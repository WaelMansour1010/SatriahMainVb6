VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEmployee_SpecialReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   20205
   Icon            =   "FrmEmployee_SpecialReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   20205
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   9135
      Left            =   90
      TabIndex        =   2
      Top             =   750
      Width           =   19395
      Begin VB.Frame Frame10 
         Height          =   6705
         Left            =   660
         TabIndex        =   123
         Top             =   600
         Visible         =   0   'False
         Width           =   18525
         Begin VB.TextBox txtPath 
            Height          =   465
            Left            =   4380
            TabIndex        =   128
            Top             =   180
            Width           =   4785
         End
         Begin VB.CommandButton Command12 
            Caption         =   " ’œÌ— «·Ï «ﬂ”Ì·"
            Height          =   615
            Left            =   1530
            TabIndex        =   127
            Top             =   150
            Width           =   1215
         End
         Begin ImpulseButton.ISButton CmdClose 
            Height          =   270
            Left            =   17820
            TabIndex        =   124
            Top             =   180
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«€·«ﬁ"
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
            ButtonImage     =   "FrmEmployee_SpecialReport.frx":038A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grd 
            Height          =   5655
            Left            =   90
            TabIndex        =   125
            Top             =   870
            Width           =   18315
            _cx             =   32306
            _cy             =   9975
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   87
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmEmployee_SpecialReport.frx":0924
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
         Begin VSFlex8Ctl.VSFlexGrid tmpGrd 
            Height          =   5655
            Left            =   0
            TabIndex        =   130
            Top             =   0
            Visible         =   0   'False
            Width           =   18315
            _cx             =   32306
            _cy             =   9975
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   87
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmEmployee_SpecialReport.frx":190C
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”«— „·› «·«ﬂ”Ì·"
            Height          =   285
            Index           =   5
            Left            =   9210
            TabIndex        =   129
            Top             =   210
            Width           =   1755
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   4695
         Left            =   21720
         TabIndex        =   37
         Top             =   11520
         Width           =   4335
         Begin VB.Image Image1 
            Height          =   3675
            Index           =   1
            Left            =   120
            Picture         =   "FrmEmployee_SpecialReport.frx":28F4
            Stretch         =   -1  'True
            Top             =   120
            Width           =   4395
         End
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
            Height          =   1095
            Left            =   480
            TabIndex        =   38
            Top             =   3840
            Width           =   2895
         End
      End
      Begin VB.CheckBox ChkStatus 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈ŸÂ«— ﬂ· «·„ÊŸ›Ì‰ „⁄ «·„‰ ÂÌ… Œœ„« Â„"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   15720
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   7455
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   19215
         Begin VB.Frame Frame8 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Ã“«¡"
            Height          =   2295
            Left            =   13620
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   4920
            Width           =   5505
            Begin VB.ListBox PenaltyTypeList 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":4E4C
               Left            =   2820
               List            =   "FrmEmployee_SpecialReport.frx":4E53
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   360
               Width           =   2610
            End
            Begin VB.ListBox SelectedPenaltyTypeList 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":4E60
               Left            =   240
               List            =   "FrmEmployee_SpecialReport.frx":4E67
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   360
               Width           =   1965
            End
            Begin VB.Label PSin 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   2265
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   1170
               Width           =   570
            End
            Begin VB.Label PMin 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   2265
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   1455
               Width           =   480
            End
            Begin VB.Label PMout 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   2265
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   825
               Width           =   375
            End
            Begin VB.Label PSout 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   2265
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   510
               Width           =   480
            End
         End
         Begin MSDataListLib.DataCombo DCBranch 
            Height          =   285
            Left            =   19320
            TabIndex        =   18
            Top             =   2160
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmpDepartments 
            Height          =   285
            Left            =   19320
            TabIndex        =   20
            Top             =   2520
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCRegionID 
            Height          =   285
            Left            =   19320
            TabIndex        =   22
            Top             =   2880
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCNationality 
            Height          =   285
            Left            =   19320
            TabIndex        =   24
            Top             =   3240
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboJobsType 
            Height          =   315
            Left            =   19320
            TabIndex        =   26
            Top             =   3600
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   19320
            TabIndex        =   32
            Top             =   3960
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSanction 
            Height          =   315
            Left            =   19320
            TabIndex        =   41
            Top             =   4320
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcstatus 
            Height          =   315
            Left            =   19320
            TabIndex        =   42
            Top             =   4680
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic10 
            Height          =   2295
            Left            =   6480
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   120
            Width           =   6255
            _cx             =   11033
            _cy             =   4048
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
            Caption         =   "«·«œ«—…"
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
            Begin VB.ListBox DepList 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":4E7A
               Left            =   3390
               List            =   "FrmEmployee_SpecialReport.frx":4E81
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   360
               Width           =   2610
            End
            Begin VB.ListBox SelectedDepList 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":4E8E
               Left            =   180
               List            =   "FrmEmployee_SpecialReport.frx":4E95
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   360
               Width           =   2565
            End
            Begin VB.Label DSin 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   2835
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   1170
               Width           =   570
            End
            Begin VB.Label DMin 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   2835
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   1455
               Width           =   480
            End
            Begin VB.Label DMout 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   2835
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   825
               Width           =   375
            End
            Begin VB.Label DSout 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   2835
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   510
               Width           =   480
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   2295
            Left            =   12840
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   120
            Width           =   6255
            _cx             =   11033
            _cy             =   4048
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
            Caption         =   "«·›—⁄"
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
            Begin VB.Frame Frame6 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   2295
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   0
               Width           =   6255
               Begin VB.ListBox SelectedBranchList 
                  Height          =   1620
                  ItemData        =   "FrmEmployee_SpecialReport.frx":4EA8
                  Left            =   240
                  List            =   "FrmEmployee_SpecialReport.frx":4EAF
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   360
                  Width           =   2565
               End
               Begin VB.ListBox BranchList 
                  Height          =   1620
                  ItemData        =   "FrmEmployee_SpecialReport.frx":4EC2
                  Left            =   3450
                  List            =   "FrmEmployee_SpecialReport.frx":4EC9
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   360
                  Width           =   2610
               End
               Begin VB.Label BSout 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   510
                  Width           =   480
               End
               Begin VB.Label BMout 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   360
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   825
                  Width           =   375
               End
               Begin VB.Label BMin 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "<<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   330
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   1455
                  Width           =   480
               End
               Begin VB.Label BSin 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   300
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   1170
                  Width           =   570
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   2295
            Left            =   12840
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   2520
            Width           =   6255
            _cx             =   11033
            _cy             =   4048
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
            Caption         =   "«·Ã‰”Ì…"
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
            Begin VB.Frame Frame7 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ã‰”Ì…"
               Height          =   2295
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   0
               Width           =   6255
               Begin VB.ListBox SelectedNationalityList 
                  Height          =   1620
                  ItemData        =   "FrmEmployee_SpecialReport.frx":4ED6
                  Left            =   240
                  List            =   "FrmEmployee_SpecialReport.frx":4EDD
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   360
                  Width           =   2565
               End
               Begin VB.ListBox NationalityList 
                  Height          =   1620
                  ItemData        =   "FrmEmployee_SpecialReport.frx":4EF0
                  Left            =   3450
                  List            =   "FrmEmployee_SpecialReport.frx":4EF7
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   360
                  Width           =   2610
               End
               Begin VB.Label NSout 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   510
                  Width           =   480
               End
               Begin VB.Label NMout 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   360
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   825
                  Width           =   375
               End
               Begin VB.Label NMin 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "<<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   330
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   1455
                  Width           =   480
               End
               Begin VB.Label NSin 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   300
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   1170
                  Width           =   570
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   2295
            Left            =   6480
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   2520
            Width           =   6255
            _cx             =   11033
            _cy             =   4048
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
            Caption         =   "«·ÊŸÌ›…"
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
            Begin VB.ListBox SelectedJobList 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":4F04
               Left            =   180
               List            =   "FrmEmployee_SpecialReport.frx":4F0B
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   360
               Width           =   2565
            End
            Begin VB.ListBox JobList 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":4F1E
               Left            =   3390
               List            =   "FrmEmployee_SpecialReport.frx":4F25
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   360
               Width           =   2610
            End
            Begin VB.Label JSout 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   2835
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   510
               Width           =   480
            End
            Begin VB.Label JMout 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   2835
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   825
               Width           =   375
            End
            Begin VB.Label JMin 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   2835
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   1455
               Width           =   480
            End
            Begin VB.Label JSin 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   2835
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   1170
               Width           =   570
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   2295
            Left            =   8940
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   4920
            Width           =   4605
            _cx             =   8123
            _cy             =   4048
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
            Caption         =   "Õ«·… «·⁄„·"
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
            Begin VB.ListBox SelectedWorkCaseList 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":4F32
               Left            =   180
               List            =   "FrmEmployee_SpecialReport.frx":4F39
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   360
               Width           =   1740
            End
            Begin VB.ListBox WorkCaseList 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":4F4C
               Left            =   2565
               List            =   "FrmEmployee_SpecialReport.frx":4F53
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   360
               Width           =   1785
            End
            Begin VB.Label CSout 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   2010
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   510
               Width           =   480
            End
            Begin VB.Label CMout 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   2010
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   825
               Width           =   375
            End
            Begin VB.Label CMin 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   2010
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   1455
               Width           =   480
            End
            Begin VB.Label CSin 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   2010
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   1170
               Width           =   570
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   2295
            Left            =   120
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   2520
            Width           =   6255
            _cx             =   11033
            _cy             =   4048
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
            Caption         =   "„ÊŸ› „Õœœ"
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
            Begin VB.Frame Frame5 
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ÊŸ› „Õœœ"
               Height          =   2295
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   0
               Width           =   6255
               Begin VB.ListBox SelectedEmpList 
                  Height          =   1620
                  ItemData        =   "FrmEmployee_SpecialReport.frx":4F60
                  Left            =   240
                  List            =   "FrmEmployee_SpecialReport.frx":4F67
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   360
                  Width           =   2565
               End
               Begin VB.ListBox EmpList 
                  Height          =   1620
                  ItemData        =   "FrmEmployee_SpecialReport.frx":4F7A
                  Left            =   3450
                  List            =   "FrmEmployee_SpecialReport.frx":4F81
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   360
                  Width           =   2610
               End
               Begin VB.Label ESout 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   480
                  Width           =   480
               End
               Begin VB.Label EMout 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   360
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   825
                  Width           =   375
               End
               Begin VB.Label EMin 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "<<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   330
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   1455
                  Width           =   480
               End
               Begin VB.Label ESin 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   300
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   1170
                  Width           =   570
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   2295
            Left            =   120
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   120
            Width           =   6255
            _cx             =   11033
            _cy             =   4048
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
            Caption         =   "«·«œ«—…/«·ﬁÿ«⁄"
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
            Begin VB.Frame Frame9 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÿ«⁄"
               Height          =   2295
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   0
               Width           =   6255
               Begin VB.ListBox SelectedSecList 
                  Height          =   1620
                  ItemData        =   "FrmEmployee_SpecialReport.frx":4F8E
                  Left            =   240
                  List            =   "FrmEmployee_SpecialReport.frx":4F95
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   360
                  Width           =   2565
               End
               Begin VB.ListBox SecList 
                  Height          =   1620
                  ItemData        =   "FrmEmployee_SpecialReport.frx":4FA8
                  Left            =   3450
                  List            =   "FrmEmployee_SpecialReport.frx":4FAF
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   360
                  Width           =   2610
               End
               Begin VB.Label SSout 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   510
                  Width           =   480
               End
               Begin VB.Label SMout 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   360
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   825
                  Width           =   375
               End
               Begin VB.Label SMin 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "<<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   330
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   1455
                  Width           =   480
               End
               Begin VB.Label SSin 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   300
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   1170
                  Width           =   570
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   2295
            Left            =   4260
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   4920
            Width           =   4665
            _cx             =   8229
            _cy             =   4048
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
            Caption         =   "ÿ—ﬁ «·œ›⁄"
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
            Begin VB.ListBox PaymentTypesList 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":4FBC
               Left            =   2595
               List            =   "FrmEmployee_SpecialReport.frx":4FC3
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   360
               Width           =   1815
            End
            Begin VB.ListBox selectedPaymentTypesList 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":4FD0
               Left            =   180
               List            =   "FrmEmployee_SpecialReport.frx":4FD7
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   360
               Width           =   1770
            End
            Begin VB.Label PySin 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   1170
               Width           =   570
            End
            Begin VB.Label PyMin 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   1455
               Width           =   480
            End
            Begin VB.Label PyMout 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   825
               Width           =   375
            End
            Begin VB.Label PySout 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   510
               Width           =   480
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   2295
            Left            =   180
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   4920
            Width           =   3915
            _cx             =   6906
            _cy             =   4048
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
            Caption         =   "„Êﬁ⁄ «·⁄„·"
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
            Begin VB.ListBox SelectEmpLocations 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":4FEA
               Left            =   180
               List            =   "FrmEmployee_SpecialReport.frx":4FF1
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   360
               Width           =   1395
            End
            Begin VB.ListBox EmpLocationsList 
               Height          =   1620
               ItemData        =   "FrmEmployee_SpecialReport.frx":5004
               Left            =   2220
               List            =   "FrmEmployee_SpecialReport.frx":500B
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   360
               Width           =   1440
            End
            Begin VB.Label PySout2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1665
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   510
               Width           =   480
            End
            Begin VB.Label PyMout2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   360
               Left            =   1665
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   825
               Width           =   375
            End
            Begin VB.Label PyMin2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   1665
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   1455
               Width           =   480
            End
            Begin VB.Label PySin2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   1665
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   1170
               Width           =   570
            End
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Ã“«¡"
            Height          =   225
            Index           =   2
            Left            =   24000
            TabIndex        =   69
            Top             =   4320
            Width           =   1125
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·⁄„·"
            Height          =   285
            Index           =   6
            Left            =   24090
            TabIndex        =   43
            Top             =   4710
            Width           =   915
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ÊŸ› „Õœœ"
            Height          =   225
            Index           =   0
            Left            =   24000
            TabIndex        =   33
            Top             =   3960
            Width           =   1125
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÊŸÌ›…"
            Height          =   225
            Index           =   8
            Left            =   24240
            TabIndex        =   27
            Top             =   3600
            Width           =   885
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ã‰”Ì…"
            Height          =   225
            Index           =   27
            Left            =   24210
            TabIndex        =   25
            Top             =   3240
            Width           =   915
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—…/«·ﬁÿ«⁄"
            Height          =   225
            Index           =   59
            Left            =   24120
            TabIndex        =   23
            Top             =   2880
            Width           =   1005
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·ﬁ”„"
            Height          =   225
            Index           =   7
            Left            =   24240
            TabIndex        =   21
            Top             =   2520
            Width           =   885
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   225
            Index           =   52
            Left            =   24210
            TabIndex        =   19
            Top             =   2280
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»‹‹‹‹‹‹‹‹«⁄…"
         Height          =   1095
         Left            =   330
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   7950
         Width           =   14535
         Begin VB.CommandButton cmdEmployeesAll 
            Caption         =   "«·„ÊŸ›Ì‰ „Ã„⁄"
            Height          =   615
            Left            =   12840
            TabIndex        =   126
            Top             =   210
            Width           =   1215
         End
         Begin VB.CommandButton Command11 
            Caption         =   "ÿ»ﬁ« ··„Â‰…"
            Height          =   315
            Left            =   6810
            TabIndex        =   122
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton Command10 
            Caption         =   "ÿ»ﬁ« ··⁄„—"
            Height          =   315
            Left            =   7920
            TabIndex        =   121
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton Command9 
            Caption         =   "«·„ÊŸ›Ì‰ «·„ √Œ—Ì‰ ›Ï «·⁄Êœ… „‰ «Ã«“…"
            Height          =   615
            Left            =   11610
            TabIndex        =   120
            Top             =   180
            Width           =   1215
         End
         Begin VB.CommandButton Command8 
            Caption         =   "ÿ—ﬁ «·œ›⁄ ··„ÊŸ›Ì‰"
            Height          =   615
            Left            =   1410
            TabIndex        =   98
            Top             =   180
            Width           =   1455
         End
         Begin VB.CommandButton Command7 
            Caption         =   " ﬁ«—Ì— «·Ã“«¡« "
            Height          =   615
            Left            =   2880
            TabIndex        =   40
            Top             =   180
            Width           =   1455
         End
         Begin VB.CommandButton BtnClear 
            Caption         =   "„”Õ"
            Height          =   615
            Left            =   -60
            TabIndex        =   39
            Top             =   180
            Width           =   1455
         End
         Begin VB.CommandButton Command6 
            Caption         =   " ﬁ—Ì— „Ã„⁄ »«·«Ã«“« "
            Height          =   615
            Left            =   4380
            TabIndex        =   36
            Top             =   180
            Width           =   1035
         End
         Begin VB.CommandButton Command5 
            Caption         =   "ÿ»«⁄Â Õ«·… «·„ÊŸ›"
            Height          =   615
            Left            =   5400
            TabIndex        =   34
            Top             =   180
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "«·„›—œ« "
            Height          =   315
            Left            =   6840
            TabIndex        =   31
            Top             =   180
            Width           =   1065
         End
         Begin VB.CommandButton Command2 
            Caption         =   "«·„ÊŸ›Ì‰ «·„‰ ÂÌ… ⁄ﬁÊœÂ„"
            Height          =   615
            Left            =   9180
            TabIndex        =   30
            Top             =   180
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "«·„ÊŸ›Ì‰ «·„⁄Ì‰Ì‰"
            Height          =   615
            Left            =   10650
            TabIndex        =   29
            Top             =   180
            Width           =   945
         End
         Begin VB.CommandButton Command3 
            Caption         =   "ÿ»ﬁ« ··Ã‰”Ì…"
            Height          =   315
            Left            =   7950
            TabIndex        =   16
            Top             =   150
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   1092
         Left            =   -5760
         TabIndex        =   11
         Top             =   11280
         Width           =   6132
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   348
            Left            =   2040
            TabIndex        =   12
            Top             =   360
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   237633537
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   348
            Left            =   4080
            TabIndex        =   28
            Top             =   360
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   237633537
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   288
            Index           =   4
            Left            =   3480
            TabIndex        =   14
            Top             =   360
            Width           =   468
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   288
            Index           =   3
            Left            =   5520
            TabIndex        =   13
            Top             =   360
            Width           =   468
         End
      End
      Begin VB.Frame XPPnlTime 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·„œ… "
         Height          =   1065
         Left            =   15000
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   7920
         Width           =   4335
         Begin MSComCtl2.DTPicker XPDtbFrom 
            Height          =   345
            Left            =   2160
            TabIndex        =   5
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   237633537
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker XPDtpTo 
            Height          =   345
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   237633537
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   285
            Index           =   2
            Left            =   3240
            TabIndex        =   8
            Top             =   360
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   7
            Top             =   360
            Width           =   465
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   492
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   1128
      _ExtentX        =   1984
      _ExtentY        =   873
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   10
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " «· ﬁ«—Ì— «·„ Œ’’… ·‘ƒ‰ «·„ÊŸ›Ì‰    "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   19215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmEmployee_SpecialReport"
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
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)

Private Const GMEM_MOVEABLE = &H2
Private Const CF_HDROP = 15

Private Type DROPFILES
    pFiles As Long
    ptX As Long
    ptY As Long
    fNC As Long
    fWide As Long
End Type



Public Sub GetData()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim Ids As String
StrSQL = "SELECT     TOP 100 PERCENT dbo.tblVacationData.ID, dbo.tblVacationData.EmpID, dbo.tblVacationData.ExpectedacationDate, dbo.tblVacationData.[Value], "
StrSQL = StrSQL & "                       dbo.tblVacationData.Status1, dbo.tblVacationData.Status2, dbo.tblVacationData.ExpectedacationDateH, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
StrSQL = StrSQL & "                      dbo.TblEmployee.jopstatusid"
StrSQL = StrSQL & " FROM         dbo.tblVacationData LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.tblVacationData.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & " where 1 = 1"

   'If ChkStatus.value = vbUnchecked Then
  '      StrSQL = StrSQL & " and dbo.TblEmployee.workstate = 1"
   ' End If
    
    Ids = "0"
    'Dim i As Integer
    
    '1 *********************************************************************************
    'If Me.SelectedBranchList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedBranchList.ListCount - 1
    '        Ids = Ids & "," & SelectedBranchList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.branchid in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '2 *********************************************************************************
    'If Me.SelectedDepList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedDepList.ListCount - 1
    '        Ids = Ids & "," & SelectedDepList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.departmentid in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '3 *********************************************************************************
    'If Me.SelectedSecList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedSecList.ListCount - 1
    '        Ids = Ids & "," & SelectedSecList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '4 *********************************************************************************
    'If Me.SelectedNationalityList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedNationalityList.ListCount - 1
    '        Ids = Ids & "," & SelectedNationalityList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '5 *********************************************************************************
    'If Me.SelectedJobList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedJobList.ListCount - 1
    '        Ids = Ids & "," & SelectedJobList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.JobTypeID in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To Me.SelectedEmpList.ListCount - 1
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            StrSQL = StrSQL & " and TblEmployee.Emp_ID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '7 **********************************************************************************
    'If Me.SelectedPenaltyTypeList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
    '        Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '8 **********************************************************************************
    'If Me.SelectedWorkCaseList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
    '        Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '*************************************************************************************
    
    
    
    
    
    StrWhere = ""
   
    'If (Me.DataCombo1.Text <> "") And (val(DataCombo1.BoundText) <> 0) Then
    '    StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID =" & Me.DataCombo1.BoundText & ""
    'End If

    If Not IsNull(Me.XPDtbFrom.value) Then
        StrWhere = StrWhere & " AND dbo.tblVacationData.ExpectedacationDate >=" & SQLDate(Me.XPDtbFrom.value, True) & ""
    End If
    
    If Not IsNull(Me.XPDtpTo.value) Then
        StrWhere = StrWhere & " AND dbo.tblVacationData.ExpectedacationDate <=" & SQLDate(Me.XPDtpTo.value, True) & ""
    End If

    StrSQL = StrSQL & StrWhere

    StrSQL = StrSQL & " ORDER BY dbo.tblVacationData.ID "
 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
    Else
    Msg = "No Data"
    End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else

 rs.MoveFirst
 print_reportVoCation StrSQL
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           
 

    End If

End Sub



Private Sub btnClear_Click()
            clear_all Me
XPDtbFrom.value = ""
XPDtpTo.value = ""
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       

 GetData
            
        Case 1


            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
            Case 3
'print_report
    End Select

End Sub

Private Sub CmdClose_Click()
Frame10.Visible = False
End Sub

Private Sub cmdEmployeesAll_Click()
Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim Ids As String
    
If IsNull(XPDtbFrom.value) Or IsNull(XPDtpTo.value) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— › —  «· «—ÌŒ"
Else
MsgBox "Please select period of date"
End If
Exit Sub
End If
MySQL = ""

' MySQL =  "  SELECT    dbo.GetEmployeeSalary(dbo.TblEmployee.Emp_ID, " & SQLDate(XPDtpTo, True) & ") AS totalsalary, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, dbo.jopstatus.color, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID,"
 MySQL = " SELECT     dbo.GetEmployeeSalary (dbo.TblEmployee.Emp_ID, " & SQLDate(XPDtpTo, True) & ") AS totalsalary, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, "
 MySQL = MySQL & "                     dbo.jopstatus.color, dbo.jopstatus.name as jopstatusName, TblEmployee.*, "
  MySQL = MySQL & "                      dbo.jopstatus.namee,"
 MySQL = MySQL & "                     dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpJobsTypes.VisaCode, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeptColor,"
 MySQL = MySQL & "                     dbo.TblEmpDepartments.DeptBr, dbo.TblEmpDepartments.Dpeterial, dbo.TblEmpDepartments.short, dbo.EmpGroupDep.GroupName AS GroupName,"
 MySQL = MySQL & "                     dbo.EmpGroupDep.Fullcode AS FullGroupCode, dbo.EmpGroupDep.Ename AS LocationNameE, dbo.TblBranchesData.branch_name,"
 MySQL = MySQL & "                     dbo.TblBranchesData.branch_namee, dbo.TblEmployee.RegionID, dbo.TblSection.name AS Section, dbo.Nationality.name AS NationlName,"
 MySQL = MySQL & "                     dbo.Nationality.namee AS Nationalitynamee, dbo.TblSection.namee AS SectionE, dbo.EmpGroupDep.GroupNameE, dbo.TblEmployee.DeanID,"
 MySQL = MySQL & "                     dbo.dean.name AS DeanName, dbo.dean.namee AS DaenNameE, dbo.TblEmployee.DeptID2, dbo.TblEmpDepartmentsDet.Name AS DeptName,"
 MySQL = MySQL & "                     BasicSalary2 = (SELECT sum(VALUE)  FROM EmpSalaryComponent TCCC WHERE TCCC.AccountCode = 1 AND TCCC.emp_ID =TblEmployee.emp_ID ),"
 MySQL = MySQL & "                     FeeFood2 = (SELECT sum(VALUE)  FROM EmpSalaryComponent TCCC WHERE TCCC.AccountCode = 2 AND TCCC.emp_ID =TblEmployee.emp_ID ),"
 MySQL = MySQL & "                     FeeMove2 = (SELECT sum(VALUE) FROM EmpSalaryComponent TCCC WHERE TCCC.AccountCode = 3 AND TCCC.emp_ID =TblEmployee.emp_ID ),"
 MySQL = MySQL & "                     FeeHome2 = (SELECT sum(VALUE)  FROM EmpSalaryComponent TCCC WHERE TCCC.AccountCode = 4 AND TCCC.emp_ID =TblEmployee.emp_ID ),"
 MySQL = MySQL & "                     FeeOther2 = (SELECT sum(VALUE)  FROM EmpSalaryComponent TCCC WHERE TCCC.AccountCode = 5 AND TCCC.emp_ID =TblEmployee.emp_ID ),"
 MySQL = MySQL & "                     FeeFixed2 = (SELECT sum(VALUE)  FROM EmpSalaryComponent TCCC WHERE TCCC.AccountCode = 6 AND TCCC.emp_ID =TblEmployee.emp_ID ),"
 MySQL = MySQL & "                     FeeLoca2 = (SELECT sum(VALUE) FROM EmpSalaryComponent TCCC WHERE TCCC.AccountCode = 7 AND TCCC.emp_ID =TblEmployee.emp_ID ),"
 MySQL = MySQL & "                     FeeTel2 = (SELECT sum(VALUE) FROM EmpSalaryComponent TCCC WHERE TCCC.AccountCode = 8 AND TCCC.emp_ID =TblEmployee.emp_ID ),"
 
 MySQL = MySQL & "                     dbo.TblEmpDepartmentsDet.NameE AS DeptNameE"
 MySQL = MySQL & "  FROM         dbo.TblEmployee LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmpDepartmentsDet ON dbo.TblEmployee.DeptID2 = dbo.TblEmpDepartmentsDet.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.dean ON dbo.TblEmployee.DeanID = dbo.dean.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.Nationality ON dbo.TblEmployee.NationlID = dbo.Nationality.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblSection ON dbo.TblEmployee.RegionID = dbo.TblSection.Id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID "


 MySQL = MySQL & " where  1 = 1 "
 
 
 
 '   If ChkStatus.value = vbUnchecked Then
 '       MySQL = MySQL & " and dbo.TblEmployee.workstate = 1"
 '   End If
    
    Ids = "0"
    Dim i As Integer
    
    '1 *********************************************************************************
    If Me.SelectedBranchList.ListCount > 0 Then
        For i = 0 To Me.SelectedBranchList.ListCount - 1
            Ids = Ids & "," & SelectedBranchList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.branchid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '2 *********************************************************************************
    If Me.SelectedDepList.ListCount > 0 Then
        For i = 0 To Me.SelectedDepList.ListCount - 1
            Ids = Ids & "," & SelectedDepList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.departmentid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '3 *********************************************************************************
    If Me.SelectedSecList.ListCount > 0 Then
        For i = 0 To Me.SelectedSecList.ListCount - 1
            Ids = Ids & "," & SelectedSecList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '4 *********************************************************************************
    If Me.SelectedNationalityList.ListCount > 0 Then
        For i = 0 To Me.SelectedNationalityList.ListCount - 1
            Ids = Ids & "," & SelectedNationalityList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '5 *********************************************************************************
    If Me.SelectedJobList.ListCount > 0 Then
        For i = 0 To Me.SelectedJobList.ListCount - 1
            Ids = Ids & "," & SelectedJobList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.JobTypeID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To Me.SelectedEmpList.ListCount - 1
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.Emp_ID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '7 **********************************************************************************
    'If Me.SelectedPenaltyTypeList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
    '        Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '8 **********************************************************************************
    If Me.SelectedWorkCaseList.ListCount > 0 Then
        For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
            Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '*************************************************************************************
   '10 **********************************************************************************
    If Me.SelectEmpLocations.ListCount > 0 Then
        For i = 0 To Me.SelectEmpLocations.ListCount - 1
            Ids = Ids & "," & SelectEmpLocations.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.GroupID in  (" & Ids & ") "
        End If
    End If
    Ids = "-1"
    
    
    If Not IsNull(XPDtbFrom.value) Then
        MySQL = MySQL & " and  TblEmployee.BignDateWork >=  " & SQLDate(Me.XPDtbFrom.value, True) & ""
    End If
    
    If Not IsNull(XPDtpTo.value) Then
        MySQL = MySQL & "  and   TblEmployee.BignDateWork <=  " & SQLDate(Me.XPDtpTo.value, True) & ""
    End If
    
    'If dcstatus.BoundText <> "" Then
    '    MySQL = MySQL + " and TblEmployee.jopstatusid =" & dcstatus.BoundText
    'End If
    
    'If DCBranch.BoundText <> "" Then
    '     MySQL = MySQL & " and TblEmployee. branchid = " & val(DCBranch.BoundText)
    'End If

    'If DcboEmpDepartments.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.departmentid  = " & val(DcboEmpDepartments.BoundText)
    'End If

    'If DCRegionID.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.RegionID  = " & val(DCRegionID.BoundText)
    'End If
    
    'If DCNationality.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.Nationality = '" & DCNationality.Text & "'"
    'End If

    'If DcboJobsType.BoundText <> "" Then
    '    MySQL = MySQL & " AND   TblEmployee.JobTypeID =  " & val(DcboJobsType.BoundText)
    'End If
    
    'If DataCombo1.BoundText <> "" Then
    '    MySQL = MySQL & " AND   dbo.TblEmployee.Emp_ID =  " & val(DataCombo1.BoundText)
    'End If
       
      Frame10.Visible = True
  loadgrid MySQL, Grd, True, False
  loadgrid MySQL, tmpGrd, True, False
   ' Set RsData = New ADODB.Recordset
   ' RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
End Sub

Private Sub Command1_Click()
Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim Ids As String
    
If IsNull(XPDtbFrom.value) Or IsNull(XPDtpTo.value) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— › —  «· «—ÌŒ"
Else
MsgBox "Please select period of date"
End If
Exit Sub
End If
MySQL = ""

' MySQL =  "  SELECT    dbo.GetEmployeeSalary(dbo.TblEmployee.Emp_ID, " & SQLDate(XPDtpTo, True) & ") AS totalsalary, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, dbo.jopstatus.color, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID,"
 MySQL = " SELECT     dbo.GetEmployeeSalary (dbo.TblEmployee.Emp_ID, " & SQLDate(XPDtpTo, True) & ") AS totalsalary, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, "
 MySQL = MySQL & "                     dbo.jopstatus.color, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom,"
 MySQL = MySQL & "                     dbo.TblEmployee.workstate, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region,"
 MySQL = MySQL & "                     dbo.TblEmployee.InsuranceState, dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama,"
 MySQL = MySQL & "                     dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.DateEndekama, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.DateEndekamah,"
 MySQL = MySQL & "                     dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLincH,"
 MySQL = MySQL & "                     dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
 MySQL = MySQL & "                     dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.EmpNum, dbo.TblEmployee.CustNum, dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala,"
 MySQL = MySQL & "                     dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork, dbo.TblEmployee.Notsstkala, dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB,"
 MySQL = MySQL & "                     dbo.TblEmployee.KafelID, dbo.TblEmployee.KafelName, dbo.TblEmployee.pasplace, dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.hdodno,"
 MySQL = MySQL & "                     dbo.TblEmployee.hdoddate, dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.kafeltel, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.kafeladd,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_sakn1, dbo.TblEmployee.Emp_Salary_bus1,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1, dbo.TblEmployee.Emp_Salary_mob1, dbo.TblEmployee.Emp_Salary_mang1,"
 MySQL = MySQL & "                     dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1, dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc, dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1, dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1, dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.placeWORK,"
 MySQL = MySQL & "                     dbo.TblEmployee.project_id, dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode,"
 MySQL = MySQL & "                     dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode, dbo.TblEmployee.BlnceVocat, dbo.TblEmployee.InstanceDateH,"
 MySQL = MySQL & "                      dbo.TblEmployee.InstanceDateM, dbo.TblEmployee.PerceTage, dbo.TblEmployee.WorkShop_Job, dbo.TblEmployee.BYHour, dbo.TblEmployee.Percentage,"
 MySQL = MySQL & "                      dbo.TblEmployee.SalaryType, dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseend,"
 MySQL = MySQL & "                     dbo.TblEmployee.DriverLicense, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.OpenBalance4,"
 MySQL = MySQL & "                     dbo.TblEmployee.OpenBalanceType4, dbo.TblEmployee.swapedempid, dbo.TblEmployee.mangerid, dbo.TblEmployee.GroupID, dbo.TblEmployee.VisaNo,"
 MySQL = MySQL & "                     dbo.TblEmployee.JobTypeID3, dbo.TblEmployee.JobTypeID2, dbo.TblEmployee.JobTypeID1, dbo.TblEmployee.LastDateH, dbo.TblEmployee.LastDate,"
 MySQL = MySQL & "                     dbo.TblEmployee.IssueDateH, dbo.TblEmployee.DOBH, dbo.TblEmployee.gradeID, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.BankCard,"
 MySQL = MySQL & "                     dbo.TblEmployee.DriverId, dbo.TblEmployee.Account_Code5, dbo.TblEmployee.Account_Code4, dbo.TblEmployee.Account_Code3,"
 MySQL = MySQL & "                     dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance2, dbo.TblEmployee.OpenBalance1, dbo.TblEmployee.OpenBalanceType1,"
 MySQL = MySQL & "                     dbo.TblEmployee.OpenBalance, dbo.TblEmployee.OpenBalanceType, dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id,"
 MySQL = MySQL & "                     dbo.TblEmployee.Fullcode, dbo.TblEmployee.prifix, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId, dbo.TblEmployee.cost_center_id, dbo.jopstatus.namee,"
 MySQL = MySQL & "                     dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpJobsTypes.VisaCode, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeptColor,"
 MySQL = MySQL & "                     dbo.TblEmpDepartments.DeptBr, dbo.TblEmpDepartments.Dpeterial, dbo.TblEmpDepartments.short, dbo.EmpGroupDep.GroupName AS LocationName,"
 MySQL = MySQL & "                     dbo.EmpGroupDep.Fullcode AS FullGroupCode, dbo.EmpGroupDep.Ename AS LocationNameE, dbo.TblBranchesData.branch_name,"
 MySQL = MySQL & "                     dbo.TblBranchesData.branch_namee, dbo.TblEmployee.RegionID, dbo.TblSection.name AS Section, dbo.Nationality.name AS Nationalityname,"
 MySQL = MySQL & "                     dbo.Nationality.namee AS Nationalitynamee, dbo.TblSection.namee AS SectionE, dbo.EmpGroupDep.GroupNameE, dbo.TblEmployee.DeanID,"
 MySQL = MySQL & "                     dbo.dean.name AS DaenName, dbo.dean.namee AS DaenNameE, dbo.TblEmployee.DeptID2, dbo.TblEmpDepartmentsDet.Name AS DeptName,"
 MySQL = MySQL & "                     dbo.TblEmpDepartmentsDet.NameE AS DeptNameE"
 MySQL = MySQL & "  FROM         dbo.TblEmployee LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmpDepartmentsDet ON dbo.TblEmployee.DeptID2 = dbo.TblEmpDepartmentsDet.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.dean ON dbo.TblEmployee.DeanID = dbo.dean.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.Nationality ON dbo.TblEmployee.NationlID = dbo.Nationality.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblSection ON dbo.TblEmployee.RegionID = dbo.TblSection.Id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
 MySQL = MySQL & " where  1 = 1 "
 
 '   If ChkStatus.value = vbUnchecked Then
 '       MySQL = MySQL & " and dbo.TblEmployee.workstate = 1"
 '   End If
    
    Ids = "0"
    Dim i As Integer
    
    '1 *********************************************************************************
    If Me.SelectedBranchList.ListCount > 0 Then
        For i = 0 To Me.SelectedBranchList.ListCount - 1
            Ids = Ids & "," & SelectedBranchList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.branchid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '2 *********************************************************************************
    If Me.SelectedDepList.ListCount > 0 Then
        For i = 0 To Me.SelectedDepList.ListCount - 1
            Ids = Ids & "," & SelectedDepList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.departmentid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '3 *********************************************************************************
    If Me.SelectedSecList.ListCount > 0 Then
        For i = 0 To Me.SelectedSecList.ListCount - 1
            Ids = Ids & "," & SelectedSecList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '4 *********************************************************************************
    If Me.SelectedNationalityList.ListCount > 0 Then
        For i = 0 To Me.SelectedNationalityList.ListCount - 1
            Ids = Ids & "," & SelectedNationalityList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '5 *********************************************************************************
    If Me.SelectedJobList.ListCount > 0 Then
        For i = 0 To Me.SelectedJobList.ListCount - 1
            Ids = Ids & "," & SelectedJobList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.JobTypeID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To Me.SelectedEmpList.ListCount - 1
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.Emp_ID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '7 **********************************************************************************
    'If Me.SelectedPenaltyTypeList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
    '        Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '8 **********************************************************************************
    If Me.SelectedWorkCaseList.ListCount > 0 Then
        For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
            Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '*************************************************************************************
   '10 **********************************************************************************
    If Me.SelectEmpLocations.ListCount > 0 Then
        For i = 0 To Me.SelectEmpLocations.ListCount - 1
            Ids = Ids & "," & SelectEmpLocations.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.GroupID in  (" & Ids & ") "
        End If
    End If
    Ids = "-1"
    
    
    If Not IsNull(XPDtbFrom.value) Then
        MySQL = MySQL & " and  TblEmployee.BignDateWork >=  " & SQLDate(Me.XPDtbFrom.value, True) & ""
    End If
    
    If Not IsNull(XPDtpTo.value) Then
        MySQL = MySQL & "  and   TblEmployee.BignDateWork <=  " & SQLDate(Me.XPDtpTo.value, True) & ""
    End If
    
    'If dcstatus.BoundText <> "" Then
    '    MySQL = MySQL + " and TblEmployee.jopstatusid =" & dcstatus.BoundText
    'End If
    
    'If DCBranch.BoundText <> "" Then
    '     MySQL = MySQL & " and TblEmployee. branchid = " & val(DCBranch.BoundText)
    'End If

    'If DcboEmpDepartments.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.departmentid  = " & val(DcboEmpDepartments.BoundText)
    'End If

    'If DCRegionID.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.RegionID  = " & val(DCRegionID.BoundText)
    'End If
    
    'If DCNationality.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.Nationality = '" & DCNationality.Text & "'"
    'End If

    'If DcboJobsType.BoundText <> "" Then
    '    MySQL = MySQL & " AND   TblEmployee.JobTypeID =  " & val(DcboJobsType.BoundText)
    'End If
    
    'If DataCombo1.BoundText <> "" Then
    '    MySQL = MySQL & " AND   dbo.TblEmployee.Emp_ID =  " & val(DataCombo1.BoundText)
    'End If
       
       
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_1.rpt"
    Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_1E.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
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
       
        
       If Not IsNull(XPDtbFrom.value) Then
    xReport.ParameterFields(2).AddCurrentValue (SQLDate(Me.XPDtbFrom.value, True))
    End If
    
    If Not IsNull(XPDtpTo.value) Then
        xReport.ParameterFields(4).AddCurrentValue (SQLDate(Me.XPDtpTo.value, True))
    End If
    
    Dim ss As Integer
    RsData.MoveLast
    ss = RsData.RecordCount
    Dim dd As String
    dd = "" & ss & ""
    xReport.ParameterFields(5).AddCurrentValue dd
  
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
           
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command10_Click()

 Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
Dim MySQL As String
Dim Ids As String

If IsNull(XPDtbFrom.value) Or IsNull(XPDtpTo.value) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— › —  «· «—ÌŒ"
Else
MsgBox "Please select period of date"
End If
Exit Sub
End If
   
' MySQL = "  SELECT  dbo.GetEmployeeSalary(dbo.TblEmployee.Emp_ID, " & SQLDate(XPDtpTo, True) & ") AS totalsalary, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, dbo.jopstatus.color, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID,"
MySQL = " SELECT     dbo.GetEmployeeSalary(dbo.TblEmployee.Emp_ID, " & SQLDate(XPDtpTo, True) & ") AS totalsalary, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, "
    MySQL = MySQL & "                   dbo.jopstatus.color, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom,"
    MySQL = MySQL & "                   dbo.TblEmployee.workstate, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region,"
    MySQL = MySQL & "                   dbo.TblEmployee.InsuranceState, dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama,"
    MySQL = MySQL & "                    dbo.TblEmployee.DateEndekama, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.DateEndekamah,"
    MySQL = MySQL & "                   dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLincH,"
    MySQL = MySQL & "                   dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
    MySQL = MySQL & "                   dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.EmpNum, dbo.TblEmployee.CustNum, dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala,"
    MySQL = MySQL & "                   dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork, dbo.TblEmployee.Notsstkala, dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB,"
    MySQL = MySQL & "                   dbo.TblEmployee.KafelID, dbo.TblEmployee.KafelName, dbo.TblEmployee.pasplace, dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.hdodno,"
    MySQL = MySQL & "                   dbo.TblEmployee.hdoddate, dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.kafeltel, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.kafeladd,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_sakn1, dbo.TblEmployee.Emp_Salary_bus1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1, dbo.TblEmployee.Emp_Salary_mob1, dbo.TblEmployee.Emp_Salary_mang1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1, dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc, dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1, dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1, dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.placeWORK,"
    MySQL = MySQL & "                   dbo.TblEmployee.project_id, dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode,"
    MySQL = MySQL & "                   dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode, dbo.TblEmployee.BlnceVocat, dbo.TblEmployee.InstanceDateH,"
    MySQL = MySQL & "                   dbo.TblEmployee.InstanceDateM, dbo.TblEmployee.PerceTage, dbo.TblEmployee.WorkShop_Job, dbo.TblEmployee.BYHour, dbo.TblEmployee.Percentage,"
    MySQL = MySQL & "                   dbo.TblEmployee.SalaryType, dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseend,"
    MySQL = MySQL & "                   dbo.TblEmployee.DriverLicense, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.OpenBalance4,"
    MySQL = MySQL & "                   dbo.TblEmployee.OpenBalanceType4, dbo.TblEmployee.swapedempid, dbo.TblEmployee.mangerid, dbo.TblEmployee.GroupID, dbo.TblEmployee.VisaNo,"
    MySQL = MySQL & "                   dbo.TblEmployee.JobTypeID3, dbo.TblEmployee.JobTypeID2, dbo.TblEmployee.JobTypeID1, dbo.TblEmployee.LastDateH, dbo.TblEmployee.LastDate,"
    MySQL = MySQL & "                   dbo.TblEmployee.IssueDateH, dbo.TblEmployee.DOBH, dbo.TblEmployee.gradeID, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.BankCard,"
    MySQL = MySQL & "                   dbo.TblEmployee.DriverId, dbo.TblEmployee.Account_Code5, dbo.TblEmployee.Account_Code4, dbo.TblEmployee.Account_Code3,"
    MySQL = MySQL & "                   dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance2, dbo.TblEmployee.OpenBalance1, dbo.TblEmployee.OpenBalanceType1,"
    MySQL = MySQL & "                   dbo.TblEmployee.OpenBalance, dbo.TblEmployee.OpenBalanceType, dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id,"
    MySQL = MySQL & "                   dbo.TblEmployee.Fullcode, dbo.TblEmployee.prifix, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId, dbo.TblEmployee.cost_center_id, dbo.jopstatus.namee,dbo.jopstatus.name StatusName,"
    MySQL = MySQL & "                   dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpJobsTypes.VisaCode, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeptColor,"
    MySQL = MySQL & "                   dbo.TblEmpDepartments.DeptBr, dbo.TblEmpDepartments.Dpeterial, dbo.TblEmpDepartments.short, dbo.EmpGroupDep.GroupName AS LocationName,"
    MySQL = MySQL & "                   dbo.EmpGroupDep.Fullcode AS FullGroupCode, dbo.EmpGroupDep.Ename AS LocationNameE, dbo.TblBranchesData.branch_name,"
    MySQL = MySQL & "                   dbo.TblBranchesData.branch_namee, dbo.TblEmployee.RegionID, dbo.TblSection.name AS Section, dbo.TblSection.namee AS SectionE,"
    MySQL = MySQL & "                   dbo.EmpGroupDep.GroupNameE, dbo.TblEmployee.NationlID, dbo.Nationality.name AS NationName, dbo.Nationality.namee AS NationNameE,"
    MySQL = MySQL & "                   dbo.TblEmployee.DeanID, dbo.dean.name AS DeanName, dbo.dean.namee AS DeanNameE, dbo.TblEmployee.DeptID2,"
    MySQL = MySQL & "                   dbo.TblEmpDepartmentsDet.Name AS DeptName, dbo.TblEmpDepartmentsDet.NameE AS DeptNameE"
    MySQL = MySQL & "     FROM         dbo.TblEmployee LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmpDepartmentsDet ON dbo.TblEmployee.DeptID2 = dbo.TblEmpDepartmentsDet.ID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.dean ON dbo.TblEmployee.DeanID = dbo.dean.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.Nationality ON dbo.TblEmployee.NationlID = dbo.Nationality.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblSection ON dbo.TblEmployee.RegionID = dbo.TblSection.Id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    
    MySQL = MySQL & "         where  1 =1   "
   

   ' If ChkStatus.value = vbUnchecked Then
   '     MySQL = MySQL & " and dbo.TblEmployee.workstate = 1"
   ' End If
    
    Ids = "0"
    Dim i As Integer
    
    '1 *********************************************************************************
    If Me.SelectedBranchList.ListCount > 0 Then
        For i = 0 To Me.SelectedBranchList.ListCount - 1
            Ids = Ids & "," & SelectedBranchList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.branchid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '2 *********************************************************************************
    If Me.SelectedDepList.ListCount > 0 Then
        For i = 0 To Me.SelectedDepList.ListCount - 1
            Ids = Ids & "," & SelectedDepList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.departmentid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '3 *********************************************************************************
    If Me.SelectedSecList.ListCount > 0 Then
        For i = 0 To Me.SelectedSecList.ListCount - 1
            Ids = Ids & "," & SelectedSecList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '4 *********************************************************************************
    If Me.SelectedNationalityList.ListCount > 0 Then
        For i = 0 To Me.SelectedNationalityList.ListCount - 1
            Ids = Ids & "," & SelectedNationalityList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '5 *********************************************************************************
    If Me.SelectedJobList.ListCount > 0 Then
        For i = 0 To Me.SelectedJobList.ListCount - 1
            Ids = Ids & "," & SelectedJobList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.JobTypeID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To Me.SelectedEmpList.ListCount - 1
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.Emp_ID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '7 **********************************************************************************
    'If Me.SelectedPenaltyTypeList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
    '        Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '8 **********************************************************************************
    If Me.SelectedWorkCaseList.ListCount > 0 Then
        For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
            Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    
    '*************************************************************************************
    
       '10 **********************************************************************************
    If Me.SelectEmpLocations.ListCount > 0 Then
        For i = 0 To Me.SelectEmpLocations.ListCount - 1
            Ids = Ids & "," & SelectEmpLocations.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.GroupID in  (" & Ids & ") "
        End If
    End If
    Ids = "-1"
    
    'If dcstatus.BoundText <> "" Then
    '    MySQL = MySQL + " and TblEmployee.jopstatusid =" & dcstatus.BoundText
    'End If
    
    'If DCBranch.BoundText <> "" Then
    '     MySQL = MySQL & " and  TblEmployee.branchid = " & val(DCBranch.BoundText)
    'End If

    'If DcboEmpDepartments.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.departmentid  = " & val(DcboEmpDepartments.BoundText)
    'End If

    'If DCRegionID.BoundText <> "" Then
    '  MySQL = MySQL & " and  TblEmployee.RegionID  = " & val(DCRegionID.BoundText)
    'End If
    
    'If DCNationality.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.Nationality = '" & DCNationality.Text & "'"
    'End If

    'If DcboJobsType.BoundText <> "" Then
    '    MySQL = MySQL & " AND   TblEmployee.JobTypeID =  " & val(DcboJobsType.BoundText)
    'End If


    'If DataCombo1.BoundText <> "" Then
    '    MySQL = MySQL & " AND   dbo.TblEmployee.Emp_ID =  " & val(DataCombo1.BoundText)
    'End If
       
    
    If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_3.rpt"
    Else
                 StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_3.rpt"
            
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
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
       
        
       If Not IsNull(XPDtbFrom.value) Then
    xReport.ParameterFields(2).AddCurrentValue (SQLDate(Me.XPDtbFrom.value, True))
    End If
    
    If Not IsNull(XPDtpTo.value) Then
  xReport.ParameterFields(4).AddCurrentValue (SQLDate(Me.XPDtpTo.value, True))
    End If
    
    
     Dim ss As Integer
RsData.MoveLast
ss = RsData.RecordCount
    Dim dd As String
    dd = "" & ss & ""
  xReport.ParameterFields(5).AddCurrentValue dd
     
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
           
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault



End Sub

Private Sub Command11_Click()



 Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
Dim MySQL As String
Dim Ids As String

If IsNull(XPDtbFrom.value) Or IsNull(XPDtpTo.value) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— › —  «· «—ÌŒ"
Else
MsgBox "Please select period of date"
End If
Exit Sub
End If
   
' MySQL = "  SELECT  dbo.GetEmployeeSalary(dbo.TblEmployee.Emp_ID, " & SQLDate(XPDtpTo, True) & ") AS totalsalary, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, dbo.jopstatus.color, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID,"
MySQL = " SELECT     dbo.GetEmployeeSalary(dbo.TblEmployee.Emp_ID, " & SQLDate(XPDtpTo, True) & ") AS totalsalary, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, "
    MySQL = MySQL & "                   dbo.jopstatus.color, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom,"
    MySQL = MySQL & "                   dbo.TblEmployee.workstate, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region,"
    MySQL = MySQL & "                   dbo.TblEmployee.InsuranceState, dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama,"
    MySQL = MySQL & "                    dbo.TblEmployee.DateEndekama, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.DateEndekamah,"
    MySQL = MySQL & "                   dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLincH,"
    MySQL = MySQL & "                   dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
    MySQL = MySQL & "                   dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.EmpNum, dbo.TblEmployee.CustNum, dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala,"
    MySQL = MySQL & "                   dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork, dbo.TblEmployee.Notsstkala, dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB,"
    MySQL = MySQL & "                   dbo.TblEmployee.KafelID, dbo.TblEmployee.KafelName, dbo.TblEmployee.pasplace, dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.hdodno,"
    MySQL = MySQL & "                   dbo.TblEmployee.hdoddate, dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.kafeltel, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.kafeladd,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_sakn1, dbo.TblEmployee.Emp_Salary_bus1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1, dbo.TblEmployee.Emp_Salary_mob1, dbo.TblEmployee.Emp_Salary_mang1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1, dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc, dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1, dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1, dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.placeWORK,"
    MySQL = MySQL & "                   dbo.TblEmployee.project_id, dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode,"
    MySQL = MySQL & "                   dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode, dbo.TblEmployee.BlnceVocat, dbo.TblEmployee.InstanceDateH,"
    MySQL = MySQL & "                   dbo.TblEmployee.InstanceDateM, dbo.TblEmployee.PerceTage, dbo.TblEmployee.WorkShop_Job, dbo.TblEmployee.BYHour, dbo.TblEmployee.Percentage,"
    MySQL = MySQL & "                   dbo.TblEmployee.SalaryType, dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseend,"
    MySQL = MySQL & "                   dbo.TblEmployee.DriverLicense, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.OpenBalance4,"
    MySQL = MySQL & "                   dbo.TblEmployee.OpenBalanceType4, dbo.TblEmployee.swapedempid, dbo.TblEmployee.mangerid, dbo.TblEmployee.GroupID, dbo.TblEmployee.VisaNo,"
    MySQL = MySQL & "                   dbo.TblEmployee.JobTypeID3, dbo.TblEmployee.JobTypeID2, dbo.TblEmployee.JobTypeID1, dbo.TblEmployee.LastDateH, dbo.TblEmployee.LastDate,"
    MySQL = MySQL & "                   dbo.TblEmployee.IssueDateH, dbo.TblEmployee.DOBH, dbo.TblEmployee.gradeID, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.BankCard,"
    MySQL = MySQL & "                   dbo.TblEmployee.DriverId, dbo.TblEmployee.Account_Code5, dbo.TblEmployee.Account_Code4, dbo.TblEmployee.Account_Code3,"
    MySQL = MySQL & "                   dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance2, dbo.TblEmployee.OpenBalance1, dbo.TblEmployee.OpenBalanceType1,"
    MySQL = MySQL & "                   dbo.TblEmployee.OpenBalance, dbo.TblEmployee.OpenBalanceType, dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id,"
    MySQL = MySQL & "                   dbo.TblEmployee.Fullcode, dbo.TblEmployee.prifix, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId, dbo.TblEmployee.cost_center_id, dbo.jopstatus.namee,dbo.jopstatus.name StatusName,"
    MySQL = MySQL & "                   dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpJobsTypes.VisaCode, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeptColor,"
    MySQL = MySQL & "                   dbo.TblEmpDepartments.DeptBr, dbo.TblEmpDepartments.Dpeterial, dbo.TblEmpDepartments.short, dbo.EmpGroupDep.GroupName AS LocationName,"
    MySQL = MySQL & "                   dbo.EmpGroupDep.Fullcode AS FullGroupCode, dbo.EmpGroupDep.Ename AS LocationNameE, dbo.TblBranchesData.branch_name,"
    MySQL = MySQL & "                   dbo.TblBranchesData.branch_namee, dbo.TblEmployee.RegionID, dbo.TblSection.name AS Section, dbo.TblSection.namee AS SectionE,"
    MySQL = MySQL & "                   dbo.EmpGroupDep.GroupNameE, dbo.TblEmployee.NationlID, dbo.Nationality.name AS NationName, dbo.Nationality.namee AS NationNameE,"
    MySQL = MySQL & "                   dbo.TblEmployee.DeanID, dbo.dean.name AS DeanName, dbo.dean.namee AS DeanNameE, dbo.TblEmployee.DeptID2,"
    MySQL = MySQL & "                   dbo.TblEmpDepartmentsDet.Name AS DeptName, dbo.TblEmpDepartmentsDet.NameE AS DeptNameE"
    MySQL = MySQL & "     FROM         dbo.TblEmployee LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmpDepartmentsDet ON dbo.TblEmployee.DeptID2 = dbo.TblEmpDepartmentsDet.ID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.dean ON dbo.TblEmployee.DeanID = dbo.dean.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.Nationality ON dbo.TblEmployee.NationlID = dbo.Nationality.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblSection ON dbo.TblEmployee.RegionID = dbo.TblSection.Id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    
    MySQL = MySQL & "         where  1 =1   "
   

   ' If ChkStatus.value = vbUnchecked Then
   '     MySQL = MySQL & " and dbo.TblEmployee.workstate = 1"
   ' End If
    
    Ids = "0"
    Dim i As Integer
    
    '1 *********************************************************************************
    If Me.SelectedBranchList.ListCount > 0 Then
        For i = 0 To Me.SelectedBranchList.ListCount - 1
            Ids = Ids & "," & SelectedBranchList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.branchid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '2 *********************************************************************************
    If Me.SelectedDepList.ListCount > 0 Then
        For i = 0 To Me.SelectedDepList.ListCount - 1
            Ids = Ids & "," & SelectedDepList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.departmentid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '3 *********************************************************************************
    If Me.SelectedSecList.ListCount > 0 Then
        For i = 0 To Me.SelectedSecList.ListCount - 1
            Ids = Ids & "," & SelectedSecList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '4 *********************************************************************************
    If Me.SelectedNationalityList.ListCount > 0 Then
        For i = 0 To Me.SelectedNationalityList.ListCount - 1
            Ids = Ids & "," & SelectedNationalityList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '5 *********************************************************************************
    If Me.SelectedJobList.ListCount > 0 Then
        For i = 0 To Me.SelectedJobList.ListCount - 1
            Ids = Ids & "," & SelectedJobList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.JobTypeID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To Me.SelectedEmpList.ListCount - 1
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.Emp_ID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '7 **********************************************************************************
    'If Me.SelectedPenaltyTypeList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
    '        Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '8 **********************************************************************************
    If Me.SelectedWorkCaseList.ListCount > 0 Then
        For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
            Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    
    '*************************************************************************************
    
       '10 **********************************************************************************
    If Me.SelectEmpLocations.ListCount > 0 Then
        For i = 0 To Me.SelectEmpLocations.ListCount - 1
            Ids = Ids & "," & SelectEmpLocations.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.GroupID in  (" & Ids & ") "
        End If
    End If
    Ids = "-1"
    
    'If dcstatus.BoundText <> "" Then
    '    MySQL = MySQL + " and TblEmployee.jopstatusid =" & dcstatus.BoundText
    'End If
    
    'If DCBranch.BoundText <> "" Then
    '     MySQL = MySQL & " and  TblEmployee.branchid = " & val(DCBranch.BoundText)
    'End If

    'If DcboEmpDepartments.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.departmentid  = " & val(DcboEmpDepartments.BoundText)
    'End If

    'If DCRegionID.BoundText <> "" Then
    '  MySQL = MySQL & " and  TblEmployee.RegionID  = " & val(DCRegionID.BoundText)
    'End If
    
    'If DCNationality.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.Nationality = '" & DCNationality.Text & "'"
    'End If

    'If DcboJobsType.BoundText <> "" Then
    '    MySQL = MySQL & " AND   TblEmployee.JobTypeID =  " & val(DcboJobsType.BoundText)
    'End If


    'If DataCombo1.BoundText <> "" Then
    '    MySQL = MySQL & " AND   dbo.TblEmployee.Emp_ID =  " & val(DataCombo1.BoundText)
    'End If
       
    
    If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_4.rpt"
    Else
                 StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_4.rpt"
            
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
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
       
        
       If Not IsNull(XPDtbFrom.value) Then
    xReport.ParameterFields(2).AddCurrentValue (SQLDate(Me.XPDtbFrom.value, True))
    End If
    
    If Not IsNull(XPDtpTo.value) Then
  xReport.ParameterFields(4).AddCurrentValue (SQLDate(Me.XPDtpTo.value, True))
    End If
    
    
     Dim ss As Integer
RsData.MoveLast
ss = RsData.RecordCount
    Dim dd As String
    dd = "" & ss & ""
  xReport.ParameterFields(5).AddCurrentValue dd
     
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
           
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault





End Sub
Private Sub ExportGridToExcelAndCopyToClipboard()
    On Error GoTo ErrorHandler

    Dim fff As String
    Dim Row As Long, Col As Long
    Dim visibleCol As Long
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object

    ' ≈‰‘«¡ «”„ «·„·› „⁄  ‰”Ìﬁ «· «—ÌŒ Ê«·Êﬁ 
    fff = App.path & "\EmployeeData_" & Format(Now, "yyyyMMdd_HHmmss") & ".xls"

    ' ≈‰‘«¡ ﬂ«∆‰ Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)

    '  ’œÌ— «·»Ì«‰«  „‰ «·Ã—Ìœ ≈·Ï „·› Excel „⁄  Ã«Â· «·√⁄„œ… «·„Œ›Ì…
    visibleCol = 1 ' ⁄œ«œ «·√⁄„œ… «·„—∆Ì… ›Ì Excel
    For Col = 0 To tmpGrd.Cols - 1
        If Not tmpGrd.ColHidden(Col) Then
            '  ’œÌ— ⁄‰Ê«‰ «·⁄„Êœ («·’› «·√Ê·)
            xlSheet.cells(1, visibleCol).value = tmpGrd.TextMatrix(0, Col)
            
            '  ’œÌ— «·»Ì«‰«  ›Ì «·⁄„Êœ
            For Row = 1 To tmpGrd.rows - 1
                xlSheet.cells(Row + 1, visibleCol).value = tmpGrd.TextMatrix(Row, Col)
            Next Row
            
            visibleCol = visibleCol + 1 ' «·«‰ ﬁ«· ≈·Ï «·⁄„Êœ «· «·Ì ›Ì Excel
        End If
    Next Col

    ' Õ›Ÿ «·„·›
    xlBook.SaveAs fff
    xlBook.Close
    xlApp.Quit

    '  Õ—Ì— «·„Ê«—œ
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    ' ‰”Œ «·„·› ≈·Ï «·Õ«›Ÿ…
    Call CopyFileToClipboard(fff)

    MsgBox " „  ’œÌ— «·»Ì«‰«  ≈·Ï Excel Ê‰”Œ «·„·› ≈·Ï «·Õ«›Ÿ… »‰Ã«Õ.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "ÕœÀ Œÿ√ √À‰«¡  ’œÌ— «·»Ì«‰« : " & Err.Description, vbCritical
    On Error Resume Next
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub

Public Sub CopyFileToClipboard(FilePath As String)
    Dim df As DROPFILES
    Dim hGlobal As Long, lpGlobal As Long
    Dim fileData() As Byte
    Dim totalSize As Long

    ' ≈⁄œ«œ DROPFILES structure
    df.pFiles = LenB(df)
    df.fWide = 0 ' ANSI° ·Ê UTF-16 «Ã⁄·Â« 1

    '  ÃÂÌ“ «·»Ì«‰« 
    fileData = StrConv(FilePath & vbNullChar & vbNullChar, vbFromUnicode)
    totalSize = LenB(df) + UBound(fileData) + 1

    '  Œ’Ì’ –«ﬂ—…
    hGlobal = GlobalAlloc(GMEM_MOVEABLE, totalSize)
    If hGlobal = 0 Then Exit Sub

    lpGlobal = GlobalLock(hGlobal)
    If lpGlobal = 0 Then Exit Sub

    ' ‰”Œ «·»Ì«‰« 
    CopyMemory ByVal lpGlobal, df, LenB(df)
    CopyMemory ByVal (lpGlobal + LenB(df)), fileData(0), UBound(fileData) + 1

    Call GlobalUnlock(hGlobal)

    ' Ê÷⁄ «·»Ì«‰«  ›Ì «·Õ«›Ÿ…
    If OpenClipboard(0&) Then
        Call EmptyClipboard
        Call SetClipboardData(CF_HDROP, hGlobal)
        Call CloseClipboard
        MsgBox " „ ‰”Œ «·„·› ≈·Ï «·Õ«›Ÿ… »‰Ã«Õ.", vbInformation
    Else
        Call GlobalFree(hGlobal)
        MsgBox "›‘· ›Ì ‰”Œ «·„·› ≈·Ï «·Õ«›Ÿ….", vbCritical
    End If
End Sub


Public Sub CopyFileToClipboard2(FilePath As String)
    Dim df As DROPFILES
    Dim hGlobal As Long, lpGlobal As Long
    Dim fileData() As Byte
    Dim totalSize As Long

    ' ≈⁄œ«œ DROPFILES structure
    df.pFiles = LenB(df)
    df.fWide = 0 ' ANSI° ·Ê UTF-16 «Ã⁄·Â« 1

    '  ÃÂÌ“ «·»Ì«‰« 
    fileData = StrConv(FilePath & vbNullChar & vbNullChar, vbFromUnicode)
    totalSize = LenB(df) + UBound(fileData) + 1

    '  Œ’Ì’ –«ﬂ—…
    hGlobal = GlobalAlloc(GMEM_MOVEABLE, totalSize)
    If hGlobal = 0 Then Exit Sub

    lpGlobal = GlobalLock(hGlobal)
    If lpGlobal = 0 Then Exit Sub

    ' ‰”Œ «·»Ì«‰« 
    CopyMemory ByVal lpGlobal, df, LenB(df)
    CopyMemory ByVal (lpGlobal + LenB(df)), fileData(0), UBound(fileData) + 1

    Call GlobalUnlock(hGlobal)

    ' Ê÷⁄ «·»Ì«‰«  ›Ì «·Õ«›Ÿ…
    If OpenClipboard(0&) Then
        Call EmptyClipboard
        Call SetClipboardData(CF_HDROP, hGlobal)
        Call CloseClipboard
        MsgBox " „ ‰”Œ «·„·› ≈·Ï «·Õ«›Ÿ… »‰Ã«Õ.", vbInformation
    Else
        Call GlobalFree(hGlobal)
        MsgBox "›‘· ›Ì ‰”Œ «·„·› ≈·Ï «·Õ«›Ÿ….", vbCritical
    End If
End Sub

Private Sub Command12_Click()

ExportGridToExcelAndCopyToClipboard
Exit Sub

Dim fff As String
fff = App.path & "\EmployeeData" & CInt(Time) & ".xls"
'tmpGrd.saveGrid fff, flexFileExcel, flexXLSaveFixedCells Or flexXLSaveRaw







Dim i As Long, j As Long
Dim cellValue As String

For i = 1 To tmpGrd.rows - 1 '  Ã«Â· «·’› «·√Ê· ≈–« ﬂ«‰ —ƒÊ” √⁄„œ…
    For j = 0 To tmpGrd.Cols - 1
        If Not tmpGrd.ColHidden(j) Then
            cellValue = Trim(tmpGrd.TextMatrix(i, j))
            
            '  Õﬁﬁ ≈‰ «·ﬁÌ„… —ﬁ„Ì… ÊÿÊ·Â« √ﬂ»— „‰ 10
            If IsNumeric(cellValue) And Len(cellValue) > 10 Then
                tmpGrd.TextMatrix(i, j) = "'" & cellValue
            End If
        End If
    Next j
Next i

'tmpGrd.saveGrid fff, flexFileExcel, flexXLSaveFixedCells

Dim xlApp As Object
Dim xlBook As Object
Dim xlSheet As Object
Dim Row As Long, Col As Long

Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Sheets(1)

For Row = 0 To tmpGrd.rows - 1
    For Col = 0 To tmpGrd.Cols - 1
        xlSheet.cells(Row + 1, Col + 1).value = tmpGrd.TextMatrix(Row, Col)
    Next Col
Next Row

xlBook.SaveAs App.path & "\EmployeeData_" & Format(Now, "yyyyMMdd_HHmmss") & ".xls"
xlBook.Close
xlApp.Quit

Set xlSheet = Nothing
Set xlBook = Nothing
Set xlApp = Nothing



Call CopyFileToClipboard(fff)


'
'Dim i As Long
'For i = 0 To grd.Cols - 1
'    If Not grd.ColHidden(i) Then
'        tmpGrd.Cols = tmpGrd.Cols + 1
'        tmpGrd.ColKey(tmpGrd.Cols - 1) = grd.ColKey(i)
'        tmpGrd.TextMatrix(0, tmpGrd.Cols - 1) = grd.TextMatrix(0, i)
'    End If
'Next
''tmpGrd = Grd
'Dim fff As String
'
'    fff = GetGridFileName(tmpGrd, "»Ì«‰«  «·„ÊŸ›Ì‰")
'    'FFF = "D:\ddd54dd.xls"
'   ' tmpGrd.saveGrid fff, flexFileExcel, flexXLSaveFixedRows Or flexXLSaveFixedCols
'    fff = App.path & "\" & "EmployeeData" & CInt(Time) & ".xls"
'    tmpGrd.saveGrid fff, flexFileExcel, _
'       flexXLSaveFixedCells Or flexXLSaveRaw
'txtPath = fff
''Grd.saveGrid "C:\book1.xls", flexFileExcel, _
''       flexXLSaveFixedCells Or flexXLSaveRaw


End Sub
Public Function GetGridFileName(ByVal G As Object, Optional MainFormName As String = "") As String
    Dim GlobalGridName As String
    Dim indexs As String
    Dim MainContainerName As String

    On Error Resume Next
    indexs = G.Index

    MainContainerName = GetMainForm(G.Container)
    GlobalGridName = MainContainerName & "\" & G.Name & indexs & MainFormName
    GlobalGridName = "Import"
    GetGridFileName = App.path & GlobalGridName & ".xls"

End Function
Public Function GetMainForm(ByVal obj) As String
    Dim n As String
    On Error Resume Next
    n = obj.Container.Name

    If n = "" Then
        GetMainForm = obj.Name
    Else
        GetMainForm = GetMainForm(obj.Container)
    End If
End Function


Private Sub Command2_Click()

Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim Ids As String
    
If IsNull(XPDtbFrom.value) Or IsNull(XPDtpTo.value) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— › —  «· «—ÌŒ"
Else
MsgBox "Please select period of date"
End If
Exit Sub
End If

 MySQL = "  SELECT     dbo.GetEmployeeSalary(dbo.TblEmployee.Emp_ID, " & SQLDate(XPDtpTo, True) & ") AS totalsalary, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, "
  MySQL = MySQL & "                    dbo.jopstatus.color, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom,"
  MySQL = MySQL & "                    dbo.TblEmployee.workstate, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region,"
  MySQL = MySQL & "                    dbo.TblEmployee.InsuranceState, dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama,"
  MySQL = MySQL & "                    dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.DateEndekama, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.DateEndekamah,"
  MySQL = MySQL & "                    dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLincH,"
  MySQL = MySQL & "                    dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
  MySQL = MySQL & "                    dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.EmpNum, dbo.TblEmployee.CustNum, dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala,"
  MySQL = MySQL & "                    dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork, dbo.TblEmployee.Notsstkala, dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB,"
  MySQL = MySQL & "                    dbo.TblEmployee.KafelID, dbo.TblEmployee.KafelName, dbo.TblEmployee.pasplace, dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.hdodno,"
  MySQL = MySQL & "                    dbo.TblEmployee.hdoddate, dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.kafeltel, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.kafeladd,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_sakn1, dbo.TblEmployee.Emp_Salary_bus1,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1, dbo.TblEmployee.Emp_Salary_mob1, dbo.TblEmployee.Emp_Salary_mang1,"
  MySQL = MySQL & "                    dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1, dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc, dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1, dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1, dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.placeWORK,"
  MySQL = MySQL & "                    dbo.TblEmployee.project_id, dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode,"
  MySQL = MySQL & "                    dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode, dbo.TblEmployee.BlnceVocat, dbo.TblEmployee.InstanceDateH,"
  MySQL = MySQL & "                    dbo.TblEmployee.InstanceDateM, dbo.TblEmployee.PerceTage, dbo.TblEmployee.WorkShop_Job, dbo.TblEmployee.BYHour, dbo.TblEmployee.Percentage,"
  MySQL = MySQL & "                    dbo.TblEmployee.SalaryType, dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseend,"
  MySQL = MySQL & "                    dbo.TblEmployee.DriverLicense, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.OpenBalance4,"
  MySQL = MySQL & "                    dbo.TblEmployee.OpenBalanceType4, dbo.TblEmployee.swapedempid, dbo.TblEmployee.mangerid, dbo.TblEmployee.GroupID, dbo.TblEmployee.VisaNo,"
  MySQL = MySQL & "                    dbo.TblEmployee.JobTypeID3, dbo.TblEmployee.JobTypeID2, dbo.TblEmployee.JobTypeID1, dbo.TblEmployee.LastDateH, dbo.TblEmployee.LastDate,"
  MySQL = MySQL & "                    dbo.TblEmployee.IssueDateH, dbo.TblEmployee.DOBH, dbo.TblEmployee.gradeID, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.BankCard,"
  MySQL = MySQL & "                    dbo.TblEmployee.DriverId, dbo.TblEmployee.Account_Code5, dbo.TblEmployee.Account_Code4, dbo.TblEmployee.Account_Code3,"
  MySQL = MySQL & "                    dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance2, dbo.TblEmployee.OpenBalance1, dbo.TblEmployee.OpenBalanceType1,"
  MySQL = MySQL & "                    dbo.TblEmployee.OpenBalance, dbo.TblEmployee.OpenBalanceType, dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id,"
  MySQL = MySQL & "                    dbo.TblEmployee.Fullcode, dbo.TblEmployee.prifix, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId, dbo.TblEmployee.cost_center_id, dbo.jopstatus.namee,"
  MySQL = MySQL & "                    dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpJobsTypes.VisaCode, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeptColor,"
  MySQL = MySQL & "                    dbo.TblEmpDepartments.DeptBr, dbo.TblEmpDepartments.Dpeterial, dbo.TblEmpDepartments.short, dbo.EmpGroupDep.GroupName AS LocationName,"
  MySQL = MySQL & "                    dbo.EmpGroupDep.Fullcode AS FullGroupCode, dbo.EmpGroupDep.Ename AS LocationNameE, dbo.TblBranchesData.branch_name,"
  MySQL = MySQL & "                    dbo.TblBranchesData.branch_namee, dbo.TblEmployee.RegionID, dbo.TblSection.name AS Section, dbo.Nationality.name AS Nationalityname,"
  MySQL = MySQL & "                    dbo.Nationality.namee AS Nationalitynamee, dbo.TblSection.namee AS SectionE, dbo.EmpGroupDep.GroupNameE, dbo.TblEmployee.DeanID,"
  MySQL = MySQL & "                    dbo.dean.name AS DeanName, dbo.dean.namee AS DeanNameE, dbo.TblEmployee.DeptID2, dbo.TblEmpDepartmentsDet.Name AS DeptName,"
  MySQL = MySQL & "                    dbo.TblEmpDepartmentsDet.NameE AS DeptNameE"
  MySQL = MySQL & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblEmpDepartmentsDet ON dbo.TblEmployee.DeptID2 = dbo.TblEmpDepartmentsDet.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.dean ON dbo.TblEmployee.DeanID = dbo.dean.id LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.Nationality ON dbo.TblEmployee.NationlID = dbo.Nationality.id LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblSection ON dbo.TblEmployee.RegionID = dbo.TblSection.Id LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
  MySQL = MySQL & "         where  1 =1   "

 '   If ChkStatus.value = vbUnchecked Then
 '       MySQL = MySQL & " and dbo.TblEmployee.workstate = 0"
 '   End If
    
    Ids = "0"
    Dim i As Integer
    
    '1 *********************************************************************************
    If Me.SelectedBranchList.ListCount > 0 Then
        For i = 0 To Me.SelectedBranchList.ListCount - 1
            Ids = Ids & "," & SelectedBranchList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.branchid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '2 *********************************************************************************
    If Me.SelectedDepList.ListCount > 0 Then
        For i = 0 To Me.SelectedDepList.ListCount - 1
            Ids = Ids & "," & SelectedDepList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.departmentid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '3 *********************************************************************************
    If Me.SelectedSecList.ListCount > 0 Then
        For i = 0 To Me.SelectedSecList.ListCount - 1
            Ids = Ids & "," & SelectedSecList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '4 *********************************************************************************
    If Me.SelectedNationalityList.ListCount > 0 Then
        For i = 0 To Me.SelectedNationalityList.ListCount - 1
            Ids = Ids & "," & SelectedNationalityList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '5 *********************************************************************************
    If Me.SelectedJobList.ListCount > 0 Then
        For i = 0 To Me.SelectedJobList.ListCount - 1
            Ids = Ids & "," & SelectedJobList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.JobTypeID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To Me.SelectedEmpList.ListCount - 1
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.Emp_ID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '7 **********************************************************************************
    'If Me.SelectedPenaltyTypeList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
    '        Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '8 **********************************************************************************
    If Me.SelectedWorkCaseList.ListCount > 0 Then
        For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
            Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '*************************************************************************************

       '10 **********************************************************************************
    If Me.SelectEmpLocations.ListCount > 0 Then
        For i = 0 To Me.SelectEmpLocations.ListCount - 1
            Ids = Ids & "," & SelectEmpLocations.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.GroupID in  (" & Ids & ") "
        End If
    End If
    Ids = "-1"

    'If dcstatus.BoundText <> "" Then
    '    MySQL = MySQL + " and TblEmployee.jopstatusid =" & dcstatus.BoundText
    'End If
    
    'If DCBranch.BoundText <> "" Then
    '     MySQL = MySQL & " and  TblEmployee.branchid = " & val(DCBranch.BoundText)
    'End If

    'If DcboEmpDepartments.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.departmentid  = " & val(DcboEmpDepartments.BoundText)
    'End If

    'If DCRegionID.BoundText <> "" Then
    '   MySQL = MySQL & " and  TblEmployee.RegionID  = " & val(DCRegionID.BoundText)
    'End If
    
    'If DCNationality.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.Nationality = '" & DCNationality.Text & "'"
    'End If

    'If DcboJobsType.BoundText <> "" Then
    '    MySQL = MySQL & " AND   TblEmployee.JobTypeID =  " & val(DcboJobsType.BoundText)
    'End If

    If Not IsNull(XPDtbFrom.value) Then
            MySQL = MySQL & "  and  TblEmployee.endWork >=  " & SQLDate(Me.XPDtbFrom.value, True) & ""
    End If
    
    If Not IsNull(XPDtpTo.value) Then
            MySQL = MySQL & " and  TblEmployee.endWork <=  " & SQLDate(Me.XPDtpTo.value, True) & ""
    End If
    
    'If DataCombo1.BoundText <> "" Then
    '    MySQL = MySQL & " AND   dbo.TblEmployee.Emp_ID =  " & val(DataCombo1.BoundText)
    'End If
       
    
    
    If SystemOptions.UserInterface = ArabicInterface Then
             
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_3.rpt"
    Else
         '   StrFileName = App.path & "\REPORTS\REPORTS NEW\rpt_Employee_3.rpt"
                     StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_3E.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
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
       
        
       If Not IsNull(XPDtbFrom.value) Then
    xReport.ParameterFields(2).AddCurrentValue (SQLDate(Me.XPDtbFrom.value, True))
    End If
    
    If Not IsNull(XPDtpTo.value) Then
  xReport.ParameterFields(4).AddCurrentValue (SQLDate(Me.XPDtpTo.value, True))
    End If
     
 
     Dim ss As Integer
RsData.MoveLast
ss = RsData.RecordCount
    Dim dd As String
    dd = "" & ss & ""
  xReport.ParameterFields(5).AddCurrentValue dd
  
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
           
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub Command3_Click()
 Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
Dim MySQL As String
Dim Ids As String

If IsNull(XPDtbFrom.value) Or IsNull(XPDtpTo.value) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— › —  «· «—ÌŒ"
Else
MsgBox "Please select period of date"
End If
Exit Sub
End If
   
' MySQL = "  SELECT  dbo.GetEmployeeSalary(dbo.TblEmployee.Emp_ID, " & SQLDate(XPDtpTo, True) & ") AS totalsalary, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, dbo.jopstatus.color, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID,"
MySQL = " SELECT     dbo.GetEmployeeSalary(dbo.TblEmployee.Emp_ID, " & SQLDate(XPDtpTo, True) & ") AS totalsalary, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, "
    MySQL = MySQL & "                   dbo.jopstatus.color, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom,"
    MySQL = MySQL & "                   dbo.TblEmployee.workstate, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region,"
    MySQL = MySQL & "                   dbo.TblEmployee.InsuranceState, dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama,"
    MySQL = MySQL & "                   dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.DateEndekama, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.DateEndekamah,"
    MySQL = MySQL & "                   dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLincH,"
    MySQL = MySQL & "                   dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp,"
    MySQL = MySQL & "                   dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.EmpNum, dbo.TblEmployee.CustNum, dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala,"
    MySQL = MySQL & "                   dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork, dbo.TblEmployee.Notsstkala, dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB,"
    MySQL = MySQL & "                   dbo.TblEmployee.KafelID, dbo.TblEmployee.KafelName, dbo.TblEmployee.pasplace, dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.hdodno,"
    MySQL = MySQL & "                   dbo.TblEmployee.hdoddate, dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.kafeltel, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.kafeladd,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_sakn1, dbo.TblEmployee.Emp_Salary_bus1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1, dbo.TblEmployee.Emp_Salary_mob1, dbo.TblEmployee.Emp_Salary_mang1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1, dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc, dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1, dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1, dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.placeWORK,"
    MySQL = MySQL & "                   dbo.TblEmployee.project_id, dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode,"
    MySQL = MySQL & "                   dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode, dbo.TblEmployee.BlnceVocat, dbo.TblEmployee.InstanceDateH,"
    MySQL = MySQL & "                   dbo.TblEmployee.InstanceDateM, dbo.TblEmployee.PerceTage, dbo.TblEmployee.WorkShop_Job, dbo.TblEmployee.BYHour, dbo.TblEmployee.Percentage,"
    MySQL = MySQL & "                   dbo.TblEmployee.SalaryType, dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseend,"
    MySQL = MySQL & "                   dbo.TblEmployee.DriverLicense, dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.OpenBalance4,"
    MySQL = MySQL & "                   dbo.TblEmployee.OpenBalanceType4, dbo.TblEmployee.swapedempid, dbo.TblEmployee.mangerid, dbo.TblEmployee.GroupID, dbo.TblEmployee.VisaNo,"
    MySQL = MySQL & "                   dbo.TblEmployee.JobTypeID3, dbo.TblEmployee.JobTypeID2, dbo.TblEmployee.JobTypeID1, dbo.TblEmployee.LastDateH, dbo.TblEmployee.LastDate,"
    MySQL = MySQL & "                   dbo.TblEmployee.IssueDateH, dbo.TblEmployee.DOBH, dbo.TblEmployee.gradeID, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.BankCard,"
    MySQL = MySQL & "                   dbo.TblEmployee.DriverId, dbo.TblEmployee.Account_Code5, dbo.TblEmployee.Account_Code4, dbo.TblEmployee.Account_Code3,"
    MySQL = MySQL & "                   dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance2, dbo.TblEmployee.OpenBalance1, dbo.TblEmployee.OpenBalanceType1,"
    MySQL = MySQL & "                   dbo.TblEmployee.OpenBalance, dbo.TblEmployee.OpenBalanceType, dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id,"
    MySQL = MySQL & "                   dbo.TblEmployee.Fullcode, dbo.TblEmployee.prifix, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId, dbo.TblEmployee.cost_center_id, dbo.jopstatus.namee,dbo.jopstatus.name StatusName,"
    MySQL = MySQL & "                   dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpJobsTypes.VisaCode, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeptColor,"
    MySQL = MySQL & "                   dbo.TblEmpDepartments.DeptBr, dbo.TblEmpDepartments.Dpeterial, dbo.TblEmpDepartments.short, dbo.EmpGroupDep.GroupName AS LocationName,"
    MySQL = MySQL & "                   dbo.EmpGroupDep.Fullcode AS FullGroupCode, dbo.EmpGroupDep.Ename AS LocationNameE, dbo.TblBranchesData.branch_name,"
    MySQL = MySQL & "                   dbo.TblBranchesData.branch_namee, dbo.TblEmployee.RegionID, dbo.TblSection.name AS Section, dbo.TblSection.namee AS SectionE,"
    MySQL = MySQL & "                   dbo.EmpGroupDep.GroupNameE, dbo.TblEmployee.NationlID, dbo.Nationality.name AS NationName, dbo.Nationality.namee AS NationNameE,"
    MySQL = MySQL & "                   dbo.TblEmployee.DeanID, dbo.dean.name AS DeanName, dbo.dean.namee AS DeanNameE, dbo.TblEmployee.DeptID2,"
    MySQL = MySQL & "                   dbo.TblEmpDepartmentsDet.Name AS DeptName, dbo.TblEmpDepartmentsDet.NameE AS DeptNameE"
    MySQL = MySQL & "     FROM         dbo.TblEmployee LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmpDepartmentsDet ON dbo.TblEmployee.DeptID2 = dbo.TblEmpDepartmentsDet.ID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.dean ON dbo.TblEmployee.DeanID = dbo.dean.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.Nationality ON dbo.TblEmployee.NationlID = dbo.Nationality.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblSection ON dbo.TblEmployee.RegionID = dbo.TblSection.Id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    
    MySQL = MySQL & "         where  1 =1   "
   

   ' If ChkStatus.value = vbUnchecked Then
   '     MySQL = MySQL & " and dbo.TblEmployee.workstate = 1"
   ' End If
    
    Ids = "0"
    Dim i As Integer
    
    '1 *********************************************************************************
    If Me.SelectedBranchList.ListCount > 0 Then
        For i = 0 To Me.SelectedBranchList.ListCount - 1
            Ids = Ids & "," & SelectedBranchList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.branchid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '2 *********************************************************************************
    If Me.SelectedDepList.ListCount > 0 Then
        For i = 0 To Me.SelectedDepList.ListCount - 1
            Ids = Ids & "," & SelectedDepList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.departmentid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '3 *********************************************************************************
    If Me.SelectedSecList.ListCount > 0 Then
        For i = 0 To Me.SelectedSecList.ListCount - 1
            Ids = Ids & "," & SelectedSecList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '4 *********************************************************************************
    If Me.SelectedNationalityList.ListCount > 0 Then
        For i = 0 To Me.SelectedNationalityList.ListCount - 1
            Ids = Ids & "," & SelectedNationalityList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '5 *********************************************************************************
    If Me.SelectedJobList.ListCount > 0 Then
        For i = 0 To Me.SelectedJobList.ListCount - 1
            Ids = Ids & "," & SelectedJobList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.JobTypeID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To Me.SelectedEmpList.ListCount - 1
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.Emp_ID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '7 **********************************************************************************
    'If Me.SelectedPenaltyTypeList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
    '        Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '8 **********************************************************************************
    If Me.SelectedWorkCaseList.ListCount > 0 Then
        For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
            Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    
    '*************************************************************************************
    
       '10 **********************************************************************************
    If Me.SelectEmpLocations.ListCount > 0 Then
        For i = 0 To Me.SelectEmpLocations.ListCount - 1
            Ids = Ids & "," & SelectEmpLocations.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.GroupID in  (" & Ids & ") "
        End If
    End If
    Ids = "-1"
    
    'If dcstatus.BoundText <> "" Then
    '    MySQL = MySQL + " and TblEmployee.jopstatusid =" & dcstatus.BoundText
    'End If
    
    'If DCBranch.BoundText <> "" Then
    '     MySQL = MySQL & " and  TblEmployee.branchid = " & val(DCBranch.BoundText)
    'End If

    'If DcboEmpDepartments.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.departmentid  = " & val(DcboEmpDepartments.BoundText)
    'End If

    'If DCRegionID.BoundText <> "" Then
    '  MySQL = MySQL & " and  TblEmployee.RegionID  = " & val(DCRegionID.BoundText)
    'End If
    
    'If DCNationality.BoundText <> "" Then
    '    MySQL = MySQL & " and  TblEmployee.Nationality = '" & DCNationality.Text & "'"
    'End If

    'If DcboJobsType.BoundText <> "" Then
    '    MySQL = MySQL & " AND   TblEmployee.JobTypeID =  " & val(DcboJobsType.BoundText)
    'End If


    'If DataCombo1.BoundText <> "" Then
    '    MySQL = MySQL & " AND   dbo.TblEmployee.Emp_ID =  " & val(DataCombo1.BoundText)
    'End If
       
    
    If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_2.rpt"
    Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_2E.rpt"
            
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
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
       
        
       If Not IsNull(XPDtbFrom.value) Then
    xReport.ParameterFields(2).AddCurrentValue (SQLDate(Me.XPDtbFrom.value, True))
    End If
    
    If Not IsNull(XPDtpTo.value) Then
  xReport.ParameterFields(4).AddCurrentValue (SQLDate(Me.XPDtpTo.value, True))
    End If
    
    
     Dim ss As Integer
RsData.MoveLast
ss = RsData.RecordCount
    Dim dd As String
    dd = "" & ss & ""
  xReport.ParameterFields(5).AddCurrentValue dd
     
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
           
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault



End Sub

Private Sub Command4_Click()
 Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim MySQL As String
    Dim Ids As String
'
'If IsNull(xpdtbfrom.value) Or IsNull(XPDtpTo.value) Then
'If SystemOptions.UserInterface = ArabicInterface Then
'MsgBox "Ì—ÃÏ «Œ Ì«— › —  «· «—ÌŒ"
'Else
'MsgBox "Please select period of date"
'End If
'Exit Sub
'End If
   
  
MySQL = " SELECT     dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, dbo.jopstatus.color, dbo.jopstatus.name, dbo.TblEmployee.Emp_ID, "
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Remark,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.workstate, dbo.TblEmployee.DepartmentID,"
MySQL = MySQL & "                      dbo.TblEmployee.JobTypeID, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region, dbo.TblEmployee.InsuranceState, dbo.TblEmployee.InsuranceValue,"
MySQL = MySQL & "                      dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama, dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.DateEndekama,"
MySQL = MySQL & "                      dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc,"
MySQL = MySQL & "                      dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket,"
MySQL = MySQL & "                      dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp, dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.EmpNum, dbo.TblEmployee.CustNum,"
MySQL = MySQL & "                      dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork, dbo.TblEmployee.Notsstkala,"
MySQL = MySQL & "                      dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB, dbo.TblEmployee.KafelID, dbo.TblEmployee.KafelName, dbo.TblEmployee.pasplace,"
MySQL = MySQL & "                      dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.hdodno, dbo.TblEmployee.hdoddate, dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.kafeltel,"
MySQL = MySQL & "                      dbo.TblEmployee.jopstatusid, dbo.TblEmployee.kafeladd, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_others,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_sakn1, dbo.TblEmployee.Emp_Salary_bus1, dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_mob1, dbo.TblEmployee.Emp_Salary_mang1, dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc, dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc, dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1, dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1,"
MySQL = MySQL & "                      dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.placeWORK, dbo.TblEmployee.project_id, dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Dateexppoketh,"
MySQL = MySQL & "                      dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode, dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode,"
MySQL = MySQL & "                      dbo.TblEmployee.BlnceVocat, dbo.TblEmployee.InstanceDateH, dbo.TblEmployee.InstanceDateM, dbo.TblEmployee.PerceTage, dbo.TblEmployee.WorkShop_Job,"
MySQL = MySQL & "                      dbo.TblEmployee.BYHour, dbo.TblEmployee.Percentage, dbo.TblEmployee.SalaryType, dbo.TblEmployee.DriverLicenseendH,"
MySQL = MySQL & "                      dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseend, dbo.TblEmployee.DriverLicense, dbo.TblEmployee.lastHolidaydateH,"
MySQL = MySQL & "                      dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.OpenBalance4, dbo.TblEmployee.OpenBalanceType4, dbo.TblEmployee.swapedempid,"
MySQL = MySQL & "                      dbo.TblEmployee.mangerid, dbo.TblEmployee.GroupID, dbo.TblEmployee.VisaNo, dbo.TblEmployee.JobTypeID3, dbo.TblEmployee.JobTypeID2,"
MySQL = MySQL & "                      dbo.TblEmployee.JobTypeID1, dbo.TblEmployee.LastDateH, dbo.TblEmployee.LastDate, dbo.TblEmployee.IssueDateH, dbo.TblEmployee.DOBH,"
MySQL = MySQL & "                      dbo.TblEmployee.gradeID, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.BankCard, dbo.TblEmployee.DriverId, dbo.TblEmployee.Account_Code5,"
MySQL = MySQL & "                      dbo.TblEmployee.Account_Code4, dbo.TblEmployee.Account_Code3, dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance2,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalance1, dbo.TblEmployee.OpenBalanceType1, dbo.TblEmployee.OpenBalance, dbo.TblEmployee.OpenBalanceType,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id, dbo.TblEmployee.Fullcode, dbo.TblEmployee.prifix,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
MySQL = MySQL & "                      dbo.TblEmployee.BranchId, dbo.TblEmployee.cost_center_id, dbo.jopstatus.namee, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpJobsTypes.VisaCode,"
MySQL = MySQL & "                      dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeptColor, dbo.TblEmpDepartments.DeptBr, dbo.TblEmpDepartments.Dpeterial,"
MySQL = MySQL & "                      dbo.TblEmpDepartments.short, dbo.EmpGroupDep.GroupName AS LocationName, dbo.EmpGroupDep.Fullcode AS FullGroupCode,"
MySQL = MySQL & "                      dbo.EmpGroupDep.Ename AS LocationNameE, dbo.EmpSalaryComponent.AccountName, dbo.EmpSalaryComponent.[Value], dbo.projects.Project_name,"
MySQL = MySQL & "                      dbo.projects.Project_nameE, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Nationality.name AS Nationalityname,"
MySQL = MySQL & "                      dbo.Nationality.namee AS Nationalitynamee, dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.TblEmployee.DeanID, dbo.dean.name AS DeanName,"
MySQL = MySQL & "                      dbo.dean.namee AS DeanNameE, dbo.TblEmployee.DeptID2, dbo.TblEmpDepartmentsDet.Name AS DeptName,"
MySQL = MySQL & "                      dbo.TblEmpDepartmentsDet.NameE AS DeptNameE"
MySQL = MySQL & " FROM         dbo.Nationality RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpDepartmentsDet ON dbo.TblEmployee.DeptID2 = dbo.TblEmpDepartmentsDet.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.dean ON dbo.TblEmployee.DeanID = dbo.dean.id ON dbo.Nationality.id = dbo.TblEmployee.NationlID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.mofrdat LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode ON"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_ID = dbo.EmpSalaryComponent.emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
MySQL = MySQL & "       WHERE     (1 = 1)  "

  ' If ChkStatus.value = vbUnchecked Then
  '      MySQL = MySQL & " and dbo.TblEmployee.workstate = 1"
  '  End If
    

  '  If dcstatus.BoundText <> "" Then
  '      MySQL = MySQL + " and TblEmployee.jopstatusid =" & dcstatus.BoundText
  '  End If
    
    Ids = "0"
    Dim i As Integer
    
    '1 *********************************************************************************
    If Me.SelectedBranchList.ListCount > 0 Then
        For i = 0 To Me.SelectedBranchList.ListCount - 1
            Ids = Ids & "," & SelectedBranchList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.branchid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '2 *********************************************************************************
    If Me.SelectedDepList.ListCount > 0 Then
        For i = 0 To Me.SelectedDepList.ListCount - 1
            Ids = Ids & "," & SelectedDepList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.departmentid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '3 *********************************************************************************
    If Me.SelectedSecList.ListCount > 0 Then
        For i = 0 To Me.SelectedSecList.ListCount - 1
            Ids = Ids & "," & SelectedSecList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '4 *********************************************************************************
    If Me.SelectedNationalityList.ListCount > 0 Then
        For i = 0 To Me.SelectedNationalityList.ListCount - 1
            Ids = Ids & "," & SelectedNationalityList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '5 *********************************************************************************
    If Me.SelectedJobList.ListCount > 0 Then
        For i = 0 To Me.SelectedJobList.ListCount - 1
            Ids = Ids & "," & SelectedJobList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.JobTypeID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To Me.SelectedEmpList.ListCount - 1
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.Emp_ID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    
     MySQL = MySQL & " AND ISNULL(dbo.mofrdat.mofrad_name,'') <> '' AND ISNULL(Value,0) <> 0 "
    '7 **********************************************************************************
    'If Me.SelectedPenaltyTypeList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
    '        Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '8 **********************************************************************************
    If Me.SelectedWorkCaseList.ListCount > 0 Then
        For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
            Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '*************************************************************************************
       '10 **********************************************************************************
    If Me.SelectEmpLocations.ListCount > 0 Then
        For i = 0 To Me.SelectEmpLocations.ListCount - 1
            Ids = Ids & "," & SelectEmpLocations.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.GroupID in  (" & Ids & ") "
        End If
    End If
    Ids = "-1"
    'If val(DCBranch.BoundText) <> 0 And Me.DCBranch.Text <> "" Then
    '     MySQL = MySQL & " and  dbo.TblBranchesData.branch_id= " & val(DCBranch.BoundText)
    'End If

    'If val(DcboEmpDepartments.BoundText) <> 0 Then
    '    MySQL = MySQL & " and  TblEmployee.departmentid  = " & val(DcboEmpDepartments.BoundText)
    'End If

    'If val(DCRegionID.BoundText) <> 0 Then
    '  MySQL = MySQL & " and  TblEmployee.RegionID  = " & val(DCRegionID.BoundText)
    'End If
    
    'If val(DCNationality.BoundText) <> 0 Then
    '    MySQL = MySQL & " and  TblEmployee.Nationality = '" & DCNationality.Text & "'"
    'End If

    'If val(DcboJobsType.BoundText) <> 0 Then
    '    MySQL = MySQL & " AND   TblEmployee.JobTypeID =  " & val(DcboJobsType.BoundText)
    'End If

    'If val(DataCombo1.BoundText) <> 0 Then
    '    MySQL = MySQL & " AND   dbo.TblEmployee.Emp_ID =  " & val(DataCombo1.BoundText)
    'End If
    
    Dim X As Integer
    If SystemOptions.UserInterface = ArabicInterface Then
    X = MsgBox("Â·«  —Ìœ ⁄—÷ «›ﬁÌ", vbCritical + vbYesNoCancel)
    Else
    X = MsgBox("You want horizontal viewing", vbCritical + vbYesNoCancel)
    End If
    If X = vbNo Then
                If SystemOptions.UserInterface = ArabicInterface Then
                        StrFileName = App.path & "\REPORTS\REPORTS NEW\rpt_EmpMofrad.rpt"
                Else
                      StrFileName = App.path & "\REPORTS\REPORTS NEW\rpt_EmpMofradE.rpt"
                End If
    ElseIf X = vbYes Then
                
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrFileName = App.path & "\REPORTS\REPORTS NEW\rpt_EmpMofrad1.rpt"
                Else
                      StrFileName = App.path & "\REPORTS\REPORTS NEW\rpt_EmpMofradE1.rpt"
                End If
    Else
    Exit Sub
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
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
       
        
       If Not IsNull(XPDtbFrom.value) Then
    xReport.ParameterFields(2).AddCurrentValue (SQLDate(Me.XPDtbFrom.value, True))
    End If
    
    If Not IsNull(XPDtpTo.value) Then
  xReport.ParameterFields(4).AddCurrentValue (SQLDate(Me.XPDtpTo.value, True))
    End If
    
    
     Dim ss As Integer
RsData.MoveLast
ss = RsData.RecordCount
    Dim dd As String
    dd = "" & ss & ""
  xReport.ParameterFields(5).AddCurrentValue dd
     
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
           
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault




End Sub

Function printingReport(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim Ids As String
 
 MySQL = " SELECT     dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, "
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.BlnceVocat, dbo.TblEmployee.InstanceDateH, dbo.TblEmployee.InstanceDateM,"
MySQL = MySQL & "                      dbo.TblEmployee.PerceTage, dbo.TblEmployee.WorkShop_Job, dbo.TblEmployee.BYHour, dbo.TblEmployee.Percentage, dbo.TblEmployee.SalaryType,"
MySQL = MySQL & "                      dbo.TblEmployee.DriverLicenseendH, dbo.TblEmployee.DriverLicenseStartdH, dbo.TblEmployee.DriverLicenseend, dbo.TblEmployee.DriverLicense,"
MySQL = MySQL & "                      dbo.TblEmployee.lastHolidaydateH, dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.OpenBalance4, dbo.TblEmployee.OpenBalanceType4,"
MySQL = MySQL & "                      dbo.TblEmployee.swapedempid, dbo.TblEmployee.mangerid, dbo.TblEmployee.VisaNo, dbo.TblEmployee.LastDateH, dbo.TblEmployee.JobTypeID3,"
MySQL = MySQL & "                      TblEmpJobsTypes_4.JobTypeName AS JobTypeName3, TblEmpJobsTypes_4.JobTypeNamee AS JobTypeNamee3, dbo.TblEmployee.JobTypeID2,"
MySQL = MySQL & "                      TblEmpJobsTypes_3.JobTypeName AS JobTypeName2, TblEmpJobsTypes_3.JobTypeNamee AS JobTypeNamee2, dbo.TblEmployee.JobTypeID1,"
MySQL = MySQL & "                      TblEmpJobsTypes_2.JobTypeName AS JobTypeName1, TblEmpJobsTypes_2.JobTypeNamee AS JobTypeNamee1, dbo.TblEmployee.LastDate,"
MySQL = MySQL & "                      dbo.TblEmployee.IssueDateH, dbo.TblEmployee.DOBH, dbo.TblEmployee.gradeID, dbo.TblEmployee.InsuranceNO, dbo.TblEmployee.BankCard,"
MySQL = MySQL & "                      dbo.TblEmployee.DriverId, dbo.TblEmployee.Account_Code5, dbo.TblEmployee.Account_Code4, dbo.TblEmployee.Account_Code3, dbo.TblEmployee.OpenBalance2,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceType2, dbo.TblEmployee.OpenBalance1, dbo.TblEmployee.OpenBalanceType1, dbo.TblEmployee.OpenBalance,"
MySQL = MySQL & "                      dbo.TblEmployee.OpenBalanceType, dbo.TblEmployee.OpenBalanceDate, dbo.TblEmployee.opening_balance_voucher_id, dbo.TblEmployee.prifix,"
MySQL = MySQL & "                      dbo.TblEmployee.BranchId, dbo.TblEmployee.cost_center_id, dbo.TblEmployee.term_fullcode, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_id, dbo.TblEmployee.opr_fullcode, dbo.TblEmployee.dateendpoketh,"
MySQL = MySQL & "                      dbo.TblEmployee.Dateexppoketh, dbo.TblEmployee.Account_Code2, dbo.TblEmployee.project_id, dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.Account_code1,"
MySQL = MySQL & "                      dbo.TblEmployee.Account_code, dbo.TblEmployee.Emp_Salary_mang1, dbo.TblEmployee.Emp_Salary_mob1, dbo.TblEmployee.Emp_Salary_others1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_bus1, dbo.TblEmployee.Emp_Salary_sakn1, dbo.TblEmployee.Emp_Salary_mang,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_sakn,"
MySQL = MySQL & "                      dbo.TblEmployee.kafeladd, dbo.TblEmployee.kafeltel, dbo.TblEmployee.DOB, dbo.TblEmployee.Notsstkala, dbo.TblEmployee.EndWork,"
MySQL = MySQL & "                      dbo.TblEmployee.BignDateWork, dbo.TblEmployee.CustNum, dbo.TblEmployee.EmpNum, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.Dateexppoket,"
MySQL = MySQL & "                      dbo.TblEmployee.NumPoket, dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLinc,"
MySQL = MySQL & "                      dbo.TblEmployee.NumLicn, dbo.TblEmployee.placeEkama, dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.InsuranceState,"
MySQL = MySQL & "                      dbo.TblEmployee.Region, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.pasplace, dbo.TblEmployee.DateExpoekama, dbo.TblEmployee.DateEndekama,"
MySQL = MySQL & "                      dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.Emp_mobile, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_Mail,"
MySQL = MySQL & "                      dbo.TblEmployee.workstate, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.hdomnfaz, dbo.TblEmployee.KafelName, dbo.TblEmployee.hdoddate,"
MySQL = MySQL & "                      dbo.TblEmployee.hdodno, dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.DateEndPasp, dbo.TblEmployee.NumPasp, dbo.TblEmployee.KafelID,"
MySQL = MySQL & "                      dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.NumEkama, dbo.TblEmployee.Emp_Salary_others,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.placeWORK, dbo.TblEmployee.JobTypeID, TblEmpJobsTypes_1.JobTypeName,"
MySQL = MySQL & "                      TblEmpJobsTypes_1.JobTypeNamee, dbo.TblEmployee.dean, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_ID, dbo.EmpSalaryComponent.AccountCode,"
MySQL = MySQL & "                      dbo.EmpSalaryComponent.AccountName, dbo.EmpSalaryComponent.[Value], dbo.EmpSalaryComponent.des, dbo.EmpSalaryComponent.eq_text,"
MySQL = MySQL & "                      dbo.EmpSalaryComponent.specific_value, dbo.EmpSalaryComponent.percentage AS percentageComp, dbo.EmpSalaryComponent.min_val,"
MySQL = MySQL & "                      dbo.EmpSalaryComponent.max_val, dbo.EmpSalaryComponent.is_fixed, dbo.EmpSalaryComponent.mofrad_type, dbo.EmpSalaryComponent.ModDate,"
MySQL = MySQL & "                      dbo.EmpSalaryComponent.Flagx, dbo.EmpSalaryComponent.EntIncresDataM, dbo.EmpSalaryComponent.EntIncresDataH, dbo.mofrdat.mofrad_name,"
MySQL = MySQL & "                      dbo.mofrdat.specific_value AS specific_valueM, dbo.TblEmployee.DepartmentID, dbo.TblEmpDepartments.DepartmentName,"
MySQL = MySQL & "                      dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmployee.GroupID, dbo.EmpGroupDep.GroupName, dbo.EmpGroupDep.Ename,"
MySQL = MySQL & "                      dbo.TblEmpHolidaysDetails.fromdate, dbo.TblEmpHolidaysDetails.todate, dbo.TblEmpHolidaysDetails.fromdateH, dbo.TblEmpHolidaysDetails.todateH,"
MySQL = MySQL & "                      dbo.TblEmpHolidaysDetails.des AS DesHoliday, dbo.TblEmpHolidaysDetails.[Day], dbo.TblEmpHolidaysDetails.[Month], dbo.TblEmpHolidaysDetails.[year],"
MySQL = MySQL & "                      dbo.mofrdat.mofrad_namee, dbo.Nationality.name AS Nationalityname, dbo.Nationality.namee AS Nationalitynamee, dbo.EmpGroupDep.GroupNameE,"
MySQL = MySQL & "                      dbo.TblEmployee.DeanID, dbo.dean.name AS DeanName, dbo.dean.namee AS DeanNameE, dbo.TblEmployee.DeptID2,"
MySQL = MySQL & "                      dbo.TblEmpDepartmentsDet.Name AS DeptName, dbo.TblEmpDepartmentsDet.NameE AS DeptNameE"
MySQL = MySQL & " FROM         dbo.TblEmpJobsTypes TblEmpJobsTypes_4 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpDepartmentsDet ON dbo.TblEmployee.DeptID2 = dbo.TblEmpDepartmentsDet.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.dean ON dbo.TblEmployee.DeanID = dbo.dean.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.Nationality ON dbo.TblEmployee.NationlID = dbo.Nationality.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpHolidaysDetails ON dbo.TblEmployee.Emp_ID = dbo.TblEmpHolidaysDetails.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID ON"
MySQL = MySQL & "                      TblEmpJobsTypes_4.JobTypeID = dbo.TblEmployee.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_1 ON dbo.TblEmployee.JobTypeID = TblEmpJobsTypes_1.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_2 ON dbo.TblEmployee.JobTypeID2 = TblEmpJobsTypes_2.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_3 ON dbo.TblEmployee.JobTypeID3 = TblEmpJobsTypes_3.JobTypeID FULL OUTER JOIN"
MySQL = MySQL & "                      dbo.mofrdat RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode ON"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_id = dbo.EmpSalaryComponent.Emp_id"

MySQL = MySQL & " Where 1 = 1 "

  '  If ChkStatus.value = vbUnchecked Then
  '      MySQL = MySQL & " and dbo.TblEmployee.workstate = 1"
  '  End If
    
    Ids = "0"
    Dim i As Integer
    
    '1 *********************************************************************************
    'If Me.SelectedBranchList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedBranchList.ListCount - 1
    '        Ids = Ids & "," & SelectedBranchList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.branchid in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '2 *********************************************************************************
    'If Me.SelectedDepList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedDepList.ListCount - 1
    '        Ids = Ids & "," & SelectedDepList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.departmentid in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '3 *********************************************************************************
    'If Me.SelectedSecList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedSecList.ListCount - 1
    '        Ids = Ids & "," & SelectedSecList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '4 *********************************************************************************
    'If Me.SelectedNationalityList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedNationalityList.ListCount - 1
    '        Ids = Ids & "," & SelectedNationalityList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '5 *********************************************************************************
    'If Me.SelectedJobList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedJobList.ListCount - 1
    '        Ids = Ids & "," & SelectedJobList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.JobTypeID in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To 0
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.Emp_ID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '7 **********************************************************************************
    'If Me.SelectedPenaltyTypeList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
    '        Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '8 **********************************************************************************
    'If Me.SelectedWorkCaseList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
    '        Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '*************************************************************************************
' MySQL = MySQL & "    Where (dbo.TblQuesEmp.id = " & val(XPTxtID.text) & ")"
    If SystemOptions.UserInterface = ArabicInterface Then
      StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmp.rpt"
   Else
   StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportEmpE.rpt"
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
Dim valmofrd As String

valmofrd = GetEmployeeSalaryAccordingToComponent(val(DataCombo1.BoundText), "", 0)

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If


    xReport.ParameterFields(3).AddCurrentValue user_name
       xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(valmofrd), "0.00"), 0, True, ".")

  xReport.ParameterFields(12).AddCurrentValue valmofrd
'If C1Tab1.CurrTab = 0 Then
    xReport.ParameterFields(13).AddCurrentValue ""
'End If

'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
        
          Dim xLogo As CRAXDRT.OLEObject
   ' StrFileName = App.path & "\"& SystemOptions.ImagesPath &"\" & PICNAME & ".JPG"
  If 1 = 1 Then
  
          If Dir(App.path & "\" & SystemOptions.ImagesPath & "\" & val(DataCombo1.BoundText) & ".JPG") <> "" Then
          
    
        
             Set xLogo = xReport.Areas(1).Sections(2).AddPictureObject(App.path & "\" & SystemOptions.ImagesPath & "\" & val(FrmEmployee.XPTxtEmpID.text) & ".JPG", 500, 300)
             xLogo.Width = 1700
             xLogo.Height = 2000
             xLogo.backcolor = vbWhite
             xLogo.BorderColor = 255
             xLogo.CloseAtPageBreak = True
            
           End If
            
End If
  
    Set CViewer = New ClsReportViewer
    
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
            
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Private Sub Command5_Click()
Dim Msg As String
If Me.SelectedEmpList.ListCount <= 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·«»œ „‰ «Œ Ì«— „ÊŸ›"
    Else
        Msg = "Must Select Employee"
    End If
    MsgBox Msg, vbCritical
    Exit Sub
Else
    printingReport
End If
End Sub

Private Sub Command6_Click()
GetData
End Sub

Private Sub Command7_Click()
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim Ids As String
    
    If IsNull(XPDtbFrom.value) Or IsNull(XPDtpTo.value) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ «Œ Ì«— › —  «· «—ÌŒ"
        Else
            MsgBox "Please select period of date"
        End If
        Exit Sub
    End If

    MySQL = " SELECT     dbo.TblEmployeeWarrning.Emp_ID, dbo.TblEmployeeWarrning.recorddate, dbo.TblEmployeeWarrning.Freq, dbo.TblEmployeeWarrning.MaxSan, "
    MySQL = MySQL & "                   dbo.TblEmployeeWarrning.SanctionID, dbo.TblAdminSanction.Name, dbo.TblAdminSanction.NameE, dbo.TblEmployeeWarrning.DeptID,"
    MySQL = MySQL & "                   dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode,"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_Namee, dbo.TblEmployeeWarrning.Nationality, dbo.TblEmployeeWarrning.NumEkama, dbo.TblEmployeeWarrning.NumPasp,"
    MySQL = MySQL & "                   dbo.TblEmployeeWarrning.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployeeWarrning.Salary,"
    MySQL = MySQL & "                   dbo.TblEmployeeWarrning.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Nationality.name AS Nationalityname,"
    MySQL = MySQL & "                   dbo.Nationality.namee AS Nationalitynamee, dbo.TblEmployee.DeanID, dbo.dean.name AS DeanName, dbo.dean.namee AS DeanNameE, dbo.TblEmployee.DeptID2,"
    MySQL = MySQL & "                   dbo.TblEmpDepartmentsDet.Name AS DeptName, dbo.TblEmpDepartmentsDet.NameE AS DeptNameE"
    MySQL = MySQL & "    FROM         dbo.TblEmpDepartmentsDet RIGHT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmployee ON dbo.TblEmpDepartmentsDet.ID = dbo.TblEmployee.DeptID2 LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.dean ON dbo.TblEmployee.DeanID = dbo.dean.id RIGHT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmpDepartments RIGHT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblBranchesData RIGHT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmployeeWarrning ON dbo.TblBranchesData.branch_id = dbo.TblEmployeeWarrning.branch_no LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblEmpJobsTypes ON dbo.TblEmployeeWarrning.JobID = dbo.TblEmpJobsTypes.JobTypeID ON"
    MySQL = MySQL & "                   dbo.TblEmpDepartments.DeparmentID = dbo.TblEmployeeWarrning.DeptID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblAdminSanction ON dbo.TblEmployeeWarrning.SanctionID = dbo.TblAdminSanction.ID ON"
    MySQL = MySQL & "                   dbo.TblEmployee.Emp_ID = dbo.TblEmployeeWarrning.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & "                    dbo.Nationality ON dbo.TblEmployee.NationlID = dbo.Nationality.id"
    MySQL = MySQL & "         where  1 =1   "

    'If ChkStatus.value = vbUnchecked Then
    '    MySQL = MySQL & " and dbo.TblEmployee.workstate = 1"
    'End If
    
    Ids = "0"
    Dim i As Integer
    
    '1 *********************************************************************************
    If Me.SelectedBranchList.ListCount > 0 Then
        For i = 0 To Me.SelectedBranchList.ListCount - 1
            Ids = Ids & "," & SelectedBranchList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and dbo.TblBranchesData.branch_id in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '2 *********************************************************************************
    If Me.SelectedDepList.ListCount > 0 Then
        For i = 0 To Me.SelectedDepList.ListCount - 1
            Ids = Ids & "," & SelectedDepList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and dbo.TblEmpDepartments.DeparmentID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '3 *********************************************************************************
    'If Me.SelectedSecList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedSecList.ListCount - 1
    '        Ids = Ids & "," & SelectedSecList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '4 *********************************************************************************
    If Me.SelectedNationalityList.ListCount > 0 Then
        For i = 0 To Me.SelectedNationalityList.ListCount - 1
            Ids = Ids & "," & SelectedNationalityList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '5 *********************************************************************************
    If Me.SelectedJobList.ListCount > 0 Then
        For i = 0 To Me.SelectedJobList.ListCount - 1
            Ids = Ids & "," & SelectedJobList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and dbo.TblEmployeeWarrning.JobID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To Me.SelectedEmpList.ListCount - 1
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and dbo.TblEmployeeWarrning.Emp_ID  in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '7 **********************************************************************************
    If Me.SelectedPenaltyTypeList.ListCount > 0 Then
        For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
            Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '8 **********************************************************************************
    'If Me.SelectedWorkCaseList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
    '        Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '*************************************************************************************
    
    'If DcbSanction.Text <> "" And val(DcbSanction.BoundText) <> 0 Then
    '     MySQL = MySQL & " and  dbo.TblEmployeeWarrning.SanctionID   = " & val(DcbSanction.BoundText)
    'End If
    ' If DCBranch.Text <> "" And val(DCBranch.BoundText) <> 0 Then
    '     MySQL = MySQL & " and  dbo.TblBranchesData.branch_id  = " & val(DCBranch.BoundText)
    'End If
    

    'If DcboEmpDepartments.BoundText <> "" Then
    '    MySQL = MySQL & " and  dbo.TblEmpDepartments.DeparmentID   = " & val(DcboEmpDepartments.BoundText)
    'End If

   
    
    'If DCNationality.BoundText <> "" Then
    '    MySQL = MySQL & " and  dbo.TblEmployeeWarrning.Nationality = '" & DCNationality.Text & "'"
    'End If

    'If DcboJobsType.BoundText <> "" Then
    '    MySQL = MySQL & " AND   dbo.TblEmployeeWarrning.JobID =  " & val(DcboJobsType.BoundText)
    'End If
       '10 **********************************************************************************
    If Me.SelectEmpLocations.ListCount > 0 Then
        For i = 0 To Me.SelectEmpLocations.ListCount - 1
            Ids = Ids & "," & SelectEmpLocations.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.GroupID in  (" & Ids & ") "
        End If
    End If
    Ids = "-1"
    
    If Not IsNull(XPDtbFrom.value) Then
            MySQL = MySQL & "  and  dbo.TblEmployeeWarrning.recorddate >=  " & SQLDate(Me.XPDtbFrom.value, True) & ""
    End If
    
    If Not IsNull(XPDtpTo.value) Then
            MySQL = MySQL & " and  dbo.TblEmployeeWarrning.recorddate <=  " & SQLDate(Me.XPDtpTo.value, True) & ""
    End If
    

    'If DataCombo1.BoundText <> "" Then
    '    MySQL = MySQL & " AND   dbo.TblEmployeeWarrning.Emp_ID =  " & val(DataCombo1.BoundText)
    'End If
       
    
    
    If SystemOptions.UserInterface = ArabicInterface Then
             
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_EmployeScan.rpt"
    Else
         '   StrFileName = App.path & "\REPORTS\REPORTS NEW\rpt_Employee_3.rpt"
                     StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_EmployeScanE.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
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
       
        
       If Not IsNull(XPDtbFrom.value) Then
    xReport.ParameterFields(2).AddCurrentValue XPDtbFrom.value
    End If
    
    If Not IsNull(XPDtpTo.value) Then
  xReport.ParameterFields(4).AddCurrentValue XPDtpTo.value
    End If
     
 
     Dim ss As Integer
RsData.MoveLast
ss = RsData.RecordCount
    Dim dd As String
    dd = "" & ss & ""
  xReport.ParameterFields(5).AddCurrentValue dd
  
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
           
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Sub
Private Sub Command8_Click()

    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim Ids As String
    
    If IsNull(XPDtbFrom.value) Or IsNull(XPDtpTo.value) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ «Œ Ì«— › —  «· «—ÌŒ"
        Else
            MsgBox "Please select period of date"
        End If
        Exit Sub
    End If
 
    MySQL = " SELECT TblEmployee.Emp_ID, TblEmployee.Emp_Code, TblEmployee.Emp_Name, TblEmployee.Emp_Name1, TblEmployee.Emp_Name2, TblEmployee.Emp_Name3, TblEmployee.Emp_Name4, TblEmployee.Nationality, "
    MySQL = MySQL & " TblEmployee.dean, TblEmployee.JobTypeID, TblEmployee.placeWORK, TblEmployee.DepartmentID, TblEmployee.Emp_Salary, TblEmployee.Emp_Salary_others, TblEmployee.NumEkama, TblEmployee.DateEndekamah,"
    MySQL = MySQL & " TblEmployee.DateExpoekamaH, TblEmployee.KafelID, TblEmployee.NumPasp, TblEmployee.DateEndPasp, TblEmployee.DateExpPasp, TblEmployee.hdodno, TblEmployee.hdoddate, TblEmployee.KafelName,"
    MySQL = MySQL & " TblEmployee.hdomnfaz, TblEmployee.jopstatusid, TblEmployee.workstate, TblEmployee.Emp_Mail, TblEmployee.Emp_Phone, TblEmployee.Emp_mobile, TblEmployee.Emp_Remark, TblEmployee.Emp_Comm,"
    MySQL = MySQL & " TblEmployee.EmpProfitCom, TblEmployee.DateEndekama, TblEmployee.DateExpoekama, TblEmployee.pasplace, TblEmployee.SpecificationID, TblEmployee.Region, TblEmployee.InsuranceState, TblEmployee.InsuranceValue,"
    MySQL = MySQL & " TblEmployee.OtherDiscounts, TblEmployee.placeEkama, TblEmployee.NumLicn, TblEmployee.DateExpLinc, TblEmployee.DateEndLinc, TblEmployee.DateExpLincH, TblEmployee.DateEndLincH, TblEmployee.NumPoket,"
    MySQL = MySQL & " TblEmployee.Dateexppoket, TblEmployee.dateendpoket, TblEmployee.EmpNum, TblEmployee.CustNum, TblEmployee.ChekEndWork, TblEmployee.ChekStkala, TblEmployee.BignDateWork, TblEmployee.EndWork,"
    MySQL = MySQL & " TblEmployee.Notsstkala, TblEmployee.checkbox1, TblEmployee.DOB, TblEmployee.kafeltel, TblEmployee.kafeladd, TblEmployee.Emp_Salary_sakn, TblEmployee.Emp_Salary_bus, TblEmployee.Emp_Salary_food,"
    MySQL = MySQL & " TblEmployee.Emp_Salary_mob, TblEmployee.Emp_Salary_mang, TblEmployee.Emp_Salary_sakn1, TblEmployee.Emp_Salary_bus1, TblEmployee.Emp_Salary_food1, TblEmployee.Emp_Salary_others1,"
    MySQL = MySQL & " TblEmployee.Emp_Salary_mob1, TblEmployee.Emp_Salary_mang1, TblEmployee.Account_code, TblEmployee.Account_code1, TblEmployee.Emp_Salary_saknc, TblEmployee.Emp_Salary_busc,"
    MySQL = MySQL & " TblEmployee.Emp_Salary_foodc, TblEmployee.Emp_Salary_othersc, TblEmployee.Emp_Salary_mobc, TblEmployee.Emp_Salary_mangc, TblEmployee.Emp_Salary_saknc1, TblEmployee.Emp_Salary_busc1,"
    MySQL = MySQL & " TblEmployee.Emp_Salary_foodc1, TblEmployee.Emp_Salary_othersc1, TblEmployee.Emp_Salary_mobc1, TblEmployee.Emp_Salary_mangc1, TblEmployee.ItemPhoto, TblEmployee.project_id, TblEmployee.Account_Code2,"
    MySQL = MySQL & " TblEmployee.Dateexppoketh, TblEmployee.dateendpoketh, TblEmployee.opr_fullcode, TblEmployee.term_id, TblEmployee.opr_id, TblEmployee.term_fullcode, TblEmployee.cost_center_id, TblEmployee.BranchId,"
    MySQL = MySQL & " TblEmployee.Emp_Namee, TblEmployee.Emp_Namee1, TblEmployee.Emp_Namee2, TblEmployee.Emp_Namee3, TblEmployee.Emp_Namee4, TblEmployee.prifix, TblEmployee.Fullcode,"
    MySQL = MySQL & " TblEmployee.opening_balance_voucher_id, TblEmployee.OpenBalanceDate, TblEmployee.OpenBalanceType, TblEmployee.OpenBalance, TblEmployee.OpenBalanceType1, TblEmployee.OpenBalance1,"
    MySQL = MySQL & " TblEmployee.OpenBalanceType2, TblEmployee.OpenBalance2, TblEmployee.Account_Code3, TblEmployee.Account_Code4, TblEmployee.Account_Code5, TblEmployee.DriverId, TblEmployee.BankCard,"
    MySQL = MySQL & " TblEmployee.InsuranceNO, TblEmployee.gradeID, TblEmployee.DOBH, TblEmployee.IssueDateH, TblEmployee.LastDate, TblEmployee.LastDateH, TblEmployee.JobTypeID1, TblEmployee.JobTypeID2,"
    MySQL = MySQL & " TblEmployee.JobTypeID3, TblEmployee.VisaNo, TblEmployee.GroupID, TblEmployee.mangerid, TblEmployee.swapedempid, TblEmployee.OpenBalanceType4, TblEmployee.OpenBalance4, TblEmployee.lastHolidaydate,"
    MySQL = MySQL & " TblEmployee.lastHolidaydateH, TblEmployee.DriverLicense, TblEmployee.DriverLicenseend, TblEmployee.DriverLicenseStartdH, TblEmployee.DriverLicenseendH, TblEmployee.SalaryType, TblEmployee.Percentage,"
    MySQL = MySQL & " TblEmployee.BYHour, TblEmployee.WorkShop_Job, TblEmployee.PerceTage, TblEmployee.InstanceDateM, TblEmployee.InstanceDateH, TblEmployee.BlnceVocat, TblEmployee.DateMoveNo, TblEmployee.ChekDateIQ,"
    MySQL = MySQL & " TblEmployee.SectionID, TblEmployee.Sex, TblEmployee.RegionID, TblEmployee.EmployeeInsurance, TblEmployee.SalaryInstrunse, TblEmployee.EmpNotes, TblEmployee.NationalityE, TblEmployee.OpenBalanceType5,"
    MySQL = MySQL & " TblEmployee.OpenBalance5, TblEmployee.PayType, TblEmployee.BankCode, TblEmployee.MaritalStatus, TblEmployee.ContractID, TblEmployee.Emergencyperson, TblEmployee.EmergencyTele, TblEmployee.BankIAddress,"
    MySQL = MySQL & " TblEmployee.BanckName, TblEmployee.BankIBan, TblEmployee.SafEBox, TblEmployee.HowIqamaStH, TblEmployee.HowIqamaEndH, TblEmployee.ResourceBox, TblEmployee.TypeEmp, TblEmployee.MachinCode,"
    MySQL = MySQL & " TblEmployee.FlagDriver, TblEmployee.NoAdded, TblEmployee.DeptID2, TblEmployee.PassWord, TblEmployee.CrsID, TblEmployee.SalaryCode, TblEmployee.NationlID, TblEmployee.DeanID, TblBranchesData.branch_name,"
    MySQL = MySQL & " TblBranchesData.branch_namee, TblEmpJobsTypes.JobTypeName, TblEmpJobsTypes.JobTypeNamee, Nationality.name, Nationality.namee, TblEmpDepartments.DepartmentName, TblEmpDepartments.DepartmentNamee,"
    MySQL = MySQL & " jopstatus.name AS Expr1, jopstatus.namee AS Expr2, TblSection.name AS Expr3, TblSection.namee AS Expr4, emp_salary.total1, emp_salary.TotalAdvance, emp_salary.TotalDiscount, emp_salary.total2,"
    MySQL = MySQL & " emp_salary.EmpTotalNet , emp_salary.Sgn, emp_salary.m_year, emp_salary.m_month, emp_salary.payed"
    MySQL = MySQL & " FROM TblEmployee INNER JOIN"
    MySQL = MySQL & " emp_salary ON TblEmployee.Emp_ID = emp_salary.emp_id LEFT OUTER JOIN"
    MySQL = MySQL & " TblSection ON TblEmployee.Region = TblSection.Id LEFT OUTER JOIN"
    MySQL = MySQL & " TblEmpJobsTypes ON TblEmployee.JobTypeID = TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
    MySQL = MySQL & " TblEmpDepartments ON TblEmployee.DepartmentID = TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
    MySQL = MySQL & " TblBranchesData ON TblEmployee.BranchId = TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & " jopstatus ON TblEmployee.jopstatusid = jopstatus.id LEFT OUTER JOIN"
    MySQL = MySQL & " Nationality ON TblEmployee.NationlID = Nationality.id "
    MySQL = MySQL & " where  1 = 1 "
    
    Ids = "0"
    Dim i As Integer
    
    '1 *********************************************************************************
    If Me.SelectedBranchList.ListCount > 0 Then
        For i = 0 To Me.SelectedBranchList.ListCount - 1
            Ids = Ids & "," & SelectedBranchList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.branchid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '2 *********************************************************************************
    If Me.SelectedDepList.ListCount > 0 Then
        For i = 0 To Me.SelectedDepList.ListCount - 1
            Ids = Ids & "," & SelectedDepList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.departmentid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '3 *********************************************************************************
    If Me.SelectedSecList.ListCount > 0 Then
        For i = 0 To Me.SelectedSecList.ListCount - 1
            Ids = Ids & "," & SelectedSecList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '4 *********************************************************************************
    If Me.SelectedNationalityList.ListCount > 0 Then
        For i = 0 To Me.SelectedNationalityList.ListCount - 1
            Ids = Ids & "," & SelectedNationalityList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '5 *********************************************************************************
    If Me.SelectedJobList.ListCount > 0 Then
        For i = 0 To Me.SelectedJobList.ListCount - 1
            Ids = Ids & "," & SelectedJobList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.JobTypeID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To Me.SelectedEmpList.ListCount - 1
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.Emp_ID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '7 **********************************************************************************
    'If Me.SelectedPenaltyTypeList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
    '        Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '8 **********************************************************************************
    If Me.SelectedWorkCaseList.ListCount > 0 Then
        For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
            Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '9 **********************************************************************************
    If Me.selectedPaymentTypesList.ListCount > 0 Then
        For i = 0 To Me.selectedPaymentTypesList.ListCount - 1
            Ids = Ids & "," & selectedPaymentTypesList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.PayType in  (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '*************************************************************************************
       '10 **********************************************************************************
    If Me.SelectEmpLocations.ListCount > 0 Then
        For i = 0 To Me.SelectEmpLocations.ListCount - 1
            Ids = Ids & "," & SelectEmpLocations.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.GroupID in  (" & Ids & ") "
        End If
    End If
    Ids = "-1"
 
    
    If Not IsNull(XPDtbFrom.value) Then
        MySQL = MySQL & " and  TblEmployee.BignDateWork >=  " & SQLDate(Me.XPDtbFrom.value, True) & ""
    End If
    
    If Not IsNull(XPDtpTo.value) Then
        MySQL = MySQL & "  and   TblEmployee.BignDateWork <=  " & SQLDate(Me.XPDtpTo.value, True) & ""
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MySQL = MySQL & "  order by TblEmployee.Emp_ID"
    '---------------------------------------- Begin---------------------------------------------------
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_9.rpt"
    Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\rpt_Employee_9E.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
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
        If Not IsNull(XPDtbFrom.value) Then
            xReport.ParameterFields(2).AddCurrentValue (SQLDate(Me.XPDtbFrom.value, True))
        End If
        
        If Not IsNull(XPDtpTo.value) Then
            xReport.ParameterFields(4).AddCurrentValue (SQLDate(Me.XPDtpTo.value, True))
        End If
    
        Dim ss As Integer
        RsData.MoveLast
        ss = RsData.RecordCount
        Dim dd As String
        dd = "" & ss & ""
        xReport.ParameterFields(5).AddCurrentValue dd
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(6).AddCurrentValue XPDtbFrom.value
    xReport.ParameterFields(7).AddCurrentValue XPDtpTo.value
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command9_Click()





    Dim StrSQL As String
      Dim StrWhere As String, MySQL As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim Ids As String


StrSQL = " SELECT  TblEmployee.FullCode, TblEmployee.Emp_Name,  stratDate, stratDateH, AcuDate, AcuDateH, NoDayAct, NoDayDelay, EndDateH, EndDate, EmpID , Remark ,NoVacation "

StrSQL = StrSQL & " From dbo.TblVocationEntitlements"
StrSQL = StrSQL & " Left Outer Join TblEmployee On TblEmployee.Emp_ID = TblVocationEntitlements.EmpID"
StrSQL = StrSQL & " where 1 = 1 And IsNull(NoDayDelay,0) > 0"

   'If ChkStatus.value = vbUnchecked Then
  '      StrSQL = StrSQL & " and dbo.TblEmployee.workstate = 1"
   ' End If
    
    Ids = "0"
   '3 *********************************************************************************
    If Me.SelectedSecList.ListCount > 0 Then
        For i = 0 To Me.SelectedSecList.ListCount - 1
            Ids = Ids & "," & SelectedSecList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.RegionID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '4 *********************************************************************************
    If Me.SelectedNationalityList.ListCount > 0 Then
        For i = 0 To Me.SelectedNationalityList.ListCount - 1
            Ids = Ids & "," & SelectedNationalityList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.NationlID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '5 *********************************************************************************
    If Me.SelectedJobList.ListCount > 0 Then
        For i = 0 To Me.SelectedJobList.ListCount - 1
            Ids = Ids & "," & SelectedJobList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.JobTypeID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '6 *********************************************************************************
    If Me.SelectedEmpList.ListCount > 0 Then
        For i = 0 To Me.SelectedEmpList.ListCount - 1
            Ids = Ids & "," & SelectedEmpList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.Emp_ID in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '7 **********************************************************************************
    'If Me.SelectedPenaltyTypeList.ListCount > 0 Then
    '    For i = 0 To Me.SelectedPenaltyTypeList.ListCount - 1
    '        Ids = Ids & "," & SelectedPenaltyTypeList.ItemData(i)
    '    Next i
    '    If Ids <> "0" Then
    '        MySQL = MySQL & " and dbo.TblEmployeeWarrning.SanctionID  (" & Ids & ") "
    '    End If
    'End If
    'Ids = "0"
    '8 **********************************************************************************
    If Me.SelectedWorkCaseList.ListCount > 0 Then
        For i = 0 To Me.SelectedWorkCaseList.ListCount - 1
            Ids = Ids & "," & SelectedWorkCaseList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and TblEmployee.jopstatusid in (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '9 **********************************************************************************
    If Me.selectedPaymentTypesList.ListCount > 0 Then
        For i = 0 To Me.selectedPaymentTypesList.ListCount - 1
            Ids = Ids & "," & selectedPaymentTypesList.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.PayType in  (" & Ids & ") "
        End If
    End If
    Ids = "0"
    '*************************************************************************************
       '10 **********************************************************************************
    If Me.SelectEmpLocations.ListCount > 0 Then
        For i = 0 To Me.SelectEmpLocations.ListCount - 1
            Ids = Ids & "," & SelectEmpLocations.ItemData(i)
        Next i
        If Ids <> "0" Then
            MySQL = MySQL & " and tblemployee.GroupID in  (" & Ids & ") "
        End If
    End If
    Ids = "-1"
 
    '*************************************************************************************
    
    
    
    
    
    StrWhere = ""
   
    'If (Me.DataCombo1.Text <> "") And (val(DataCombo1.BoundText) <> 0) Then
    '    StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID =" & Me.DataCombo1.BoundText & ""
    'End If

  
    StrSQL = StrSQL & StrWhere & MySQL

    StrSQL = StrSQL & " ORDER BY dbo.TblVocationEntitlements.EmpID "
 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
    Else
    Msg = "No Data"
    End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else

 rs.MoveFirst
 print_reportVoCation StrSQL, 1
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           
 

    End If

End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub


Private Sub ChangeLang()
XPLbl(6).Caption = "Job Status"
ChkStatus.Caption = "All Employees With End Service"
Command5.Caption = "Employee.Status"
Command6.Caption = "Vacation Totals"
XPLbl(0).Caption = "Employee"
Command7.Caption = " Sanctions"
XPPnlTime.Caption = "Period"
 Command4.Caption = "Benefits"
Label5.Caption = "Special Reports For Employees Affairs"
 XPLbl(52).Caption = "Branch"
 XPLbl(59).Caption = "Management /Section"
 XPLbl(27).Caption = "Nationality"
 XPLbl(8).Caption = "Job"
 lbl(2).Caption = "From Date"
 lbl(0).Caption = "To"
 Frame2.Caption = "Print"
 Command1.Caption = "Employees Joined"
 Command2.Caption = "End Contracts"
 Command3.Caption = "Employees Details "
 lblCompanyname.Caption = "El-Sattaryh"
 XPLbl(7).Caption = "Department"
 BtnClear.Caption = "Clear"
 
'1 ******************************************************************
C1Elastic1.Caption = "Branch"
Frame6.Caption = "Branch"
C1Elastic1.CaptionPos = cpLeftTop
 '2 ******************************************************************
 C1Elastic10.Caption = "Department"
 C1Elastic10.CaptionPos = cpLeftTop
 '3 ******************************************************************
 C1Elastic7.Caption = "Section"
  Frame9.Caption = "Section"
 C1Elastic7.CaptionPos = cpLeftTop
 '4 ******************************************************************
 C1Elastic2.Caption = "Nationality"
 Frame7.Caption = "Nationality"
 C1Elastic2.CaptionPos = cpLeftTop
 '5 ******************************************************************
 C1Elastic3.Caption = "Job"
 C1Elastic3.CaptionPos = cpLeftTop
 '6 ******************************************************************
 C1Elastic6.Caption = "Employee"
 Frame5.Caption = "Employee"
 C1Elastic6.CaptionPos = cpLeftTop
 '7 ******************************************************************
 C1Elastic4.Caption = "Penalty Type"
 Frame8.Caption = "Penalty Type"
 C1Elastic4.CaptionPos = cpLeftTop
 '8 ******************************************************************
 C1Elastic5.Caption = "Work Status"
 C1Elastic5.CaptionPos = cpLeftTop
 '********************************************************************
End Sub

Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
   If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
   End If
    
   XPDtbFrom.value = ""
   XPDtpTo.value = Date
    
    Set Dcombos = New ClsDataCombos

       Dcombos.GetAdminSanction Me.DcbSanction
       Dcombos.GetBranches Me.DCBranch
       Dcombos.GetEmpDepartments Me.DcboEmpDepartments
       Dcombos.GetSection Me.DCRegionID
         If SystemOptions.UserInterface = ArabicInterface Then
                My_SQL = "  select id,name from jopstatus   "
         Else
         My_SQL = "  select id,namee from jopstatus   "
         End If
    fill_combo dcstatus, My_SQL


            If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from Nationality  "
    Else
        My_SQL = "  select  id,namee  from Nationality  "
    End If

    fill_combo DCNationality, My_SQL
    
   Dcombos.GetEmpJobsTypes Me.DcboJobsType
   Dcombos.GetEmployees DataCombo1
   
    Set cSearch = New clsDCboSearch
    My_SQL = "TblContract"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Resize_Form Me
    
   FillLists
    
End Sub


Function print_reportVoCation(Optional NoteSerial As String, Optional mType As Long = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
           
           If mType = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpVocationReports.rpt"
                 Else
                    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpVocationReportsE.rpt"
                End If
            Else
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpVocationReportsLate.rpt"
                 Else
                    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmpVocationReportsLate.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "Not Found Data to Show"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
     
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
  
    End If

   
  If Not IsNull(XPDtbFrom.value) And Not IsNull(XPDtpTo.value) Then
   xReport.ParameterFields(8).AddCurrentValue XPDtbFrom.value

    xReport.ParameterFields(10).AddCurrentValue XPDtpTo.value
  '  xReport.ParameterFields(11).AddCurrentValue DtpDateToH.value
    End If

  Dim total As String
  Dim totl As Double


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

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Function GetBranchIDFromCode(Optional brancHcode As String, _
Optional ByRef Emp_id As Integer) ' As Integer
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim ID As Integer
    

    
    sql = "select branch_Id from TblBranchesData where branch_code= '" & brancHcode & "'"
        
   
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        ID = IIf(IsNull(rs("branch_Id").value), 0, rs("branch_Id").value)
    Else
        ID = 0
    End If

    rs.Close
    Emp_id = ID
    'GetBranchIDFromCode = id

End Function

'kkhaled
Public Function FillLists()
    Dim listRS As ADODB.Recordset
    Set listRS = New ADODB.Recordset
    Dim i As Integer
    Dim listSQL As String
    'clear all lists ************************************************
    Me.BranchList.Clear
    Me.DepList.Clear
    Me.SecList.Clear
    Me.NationalityList.Clear
    Me.JobList.Clear
    Me.EmpList.Clear
    Me.PenaltyTypeList.Clear
    Me.WorkCaseList.Clear
    Me.PaymentTypesList.Clear
    '--------------------------------------------------
    Me.SelectedBranchList.Clear
    Me.SelectedDepList.Clear
    Me.SelectedSecList.Clear
    Me.SelectedNationalityList.Clear
    Me.SelectedJobList.Clear
    Me.SelectedEmpList.Clear
    Me.SelectedPenaltyTypeList.Clear
    Me.SelectedWorkCaseList.Clear
    Me.selectedPaymentTypesList.Clear
    
    EmpLocationsList.Clear
    SelectEmpLocations.Clear
    '1 *********************************************************************************************************************************
    Dim s As String
        s = " Select Commonname,branch_id, CSR,Privatekey,SerialNumber,SecretKey,PublickeycertPem,OrganizationName,Invoicetype,DefaultInvoicetype,branch_namee,"
        s = s & " Company_Comment,StreetName,AdditionalStreetName,BuildingNumber,PlotIdentification,CityName,PostalZone,branch_Code,branch_name,branch_id"
        s = s & " CountrySubentity,CitySubdivisionName,Company_Name_Eng,VATRegNo,Company_arabic_Name,industrey,SendingMode from TblBranchesData"
    listSQL = s
    listRS.Open listSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If listRS.RecordCount > 0 Then
            For i = 1 To listRS.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    BranchList.AddItem IIf(IsNull(listRS("branch_name").value), "", listRS("branch_name").value)
                Else
                    BranchList.AddItem IIf(IsNull(listRS("branch_namee").value), "", listRS("branch_namee").value)
                End If
                BranchList.ItemData(BranchList.NewIndex) = IIf(IsNull(listRS("branch_id").value), 0, listRS("branch_id").value)
                listRS.MoveNext
            Next i
        End If
    listRS.Close
    '2 *********************************************************************************************************************************
    listSQL = "Select * From TblEmpDepartments Order By DepartmentName"
     listRS.Open listSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If listRS.RecordCount > 0 Then
            For i = 1 To listRS.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    DepList.AddItem IIf(IsNull(listRS("DepartmentName").value), "", listRS("DepartmentName").value)
                Else
                    DepList.AddItem IIf(IsNull(listRS("DepartmentNamee").value), "", listRS("DepartmentNamee").value)
                End If
                DepList.ItemData(DepList.NewIndex) = IIf(IsNull(listRS("DeparmentID").value), 0, listRS("DeparmentID").value)
                listRS.MoveNext
            Next i
        End If
    listRS.Close
    '3 *********************************************************************************************************************************
    listSQL = "SELECT * From TblSection"
    listRS.Open listSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If listRS.RecordCount > 0 Then
            For i = 1 To listRS.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    SecList.AddItem IIf(IsNull(listRS("name").value), "", listRS("name").value)
                Else
                    SecList.AddItem IIf(IsNull(listRS("namee").value), "", listRS("namee").value)
                End If
                SecList.ItemData(SecList.NewIndex) = IIf(IsNull(listRS("Id").value), 0, listRS("Id").value)
                listRS.MoveNext
            Next i
        End If
    listRS.Close
    '4 *********************************************************************************************************************************
        listSQL = "select * from Nationality"
        listRS.Open listSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If listRS.RecordCount > 0 Then
            For i = 1 To listRS.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    NationalityList.AddItem IIf(IsNull(listRS("name").value), "", listRS("name").value)
                Else
                    NationalityList.AddItem IIf(IsNull(listRS("namee").value), "", listRS("namee").value)
                End If
                NationalityList.ItemData(NationalityList.NewIndex) = IIf(IsNull(listRS("id").value), 0, listRS("id").value)
            listRS.MoveNext
            Next i
        End If
        listRS.Close
        '5 *********************************************************************************************************************************
        listSQL = "Select * From TblEmpJobsTypes"
        listRS.Open listSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If listRS.RecordCount > 0 Then
            For i = 1 To listRS.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    JobList.AddItem IIf(IsNull(listRS("JobTypeName").value), "", listRS("JobTypeName").value)
                Else
                    JobList.AddItem IIf(IsNull(listRS("JobTypeNamee").value), "", listRS("JobTypeNamee").value)
                End If
                JobList.ItemData(JobList.NewIndex) = IIf(IsNull(listRS("JobTypeID").value), 0, listRS("JobTypeID").value)
            listRS.MoveNext
            Next i
        End If
        listRS.Close
        '6 *********************************************************************************************************************************
        listSQL = "SELECT * From TblEmployee"
        listRS.Open listSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If listRS.RecordCount > 0 Then
            For i = 1 To listRS.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    EmpList.AddItem IIf(IsNull(listRS("Emp_Name").value), "", listRS("Emp_Name").value)
                Else
                    EmpList.AddItem IIf(IsNull(listRS("Emp_Namee").value), "", listRS("Emp_Namee").value)
                End If
                EmpList.ItemData(EmpList.NewIndex) = IIf(IsNull(listRS("Emp_ID").value), 0, listRS("Emp_ID").value)
            listRS.MoveNext
            Next i
        End If
        listRS.Close
        '7 *********************************************************************************************************************************
        listSQL = "Select * From TblAdminSanction"
        listRS.Open listSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If listRS.RecordCount > 0 Then
            For i = 1 To listRS.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    PenaltyTypeList.AddItem IIf(IsNull(listRS("Name").value), "", listRS("Name").value)
                Else
                    PenaltyTypeList.AddItem IIf(IsNull(listRS("NameE").value), "", listRS("NameE").value)
                End If
                PenaltyTypeList.ItemData(PenaltyTypeList.NewIndex) = IIf(IsNull(listRS("ID").value), 0, listRS("ID").value)
            listRS.MoveNext
            Next i
        End If
        listRS.Close
        '8 *********************************************************************************************************************************
        listSQL = "select * from jopstatus"
        listRS.Open listSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If listRS.RecordCount > 0 Then
            For i = 1 To listRS.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    WorkCaseList.AddItem IIf(IsNull(listRS("name").value), "", listRS("name").value)
                Else
                    WorkCaseList.AddItem IIf(IsNull(listRS("namee").value), "", listRS("namee").value)
                End If
                WorkCaseList.ItemData(WorkCaseList.NewIndex) = IIf(IsNull(listRS("id").value), 0, listRS("id").value)
            listRS.MoveNext
            Next i
        End If
        listRS.Close
        '************************************************************************************************************************************
        listSQL = "Select GroupID id,GroupName,GroupNameE  From EmpGroupDep where LastGroup=1 Order By GroupName"
        listRS.Open listSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If listRS.RecordCount > 0 Then
            For i = 1 To listRS.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    EmpLocationsList.AddItem IIf(IsNull(listRS("GroupName").value), "", listRS("GroupName").value)
                Else
                    EmpLocationsList.AddItem IIf(IsNull(listRS("GroupNameE").value), "", listRS("GroupNameE").value)
                End If
                EmpLocationsList.ItemData(EmpLocationsList.NewIndex) = IIf(IsNull(listRS("id").value), 0, listRS("id").value)
            listRS.MoveNext
            Next i
        End If
        listRS.Close

        
        If SystemOptions.UserInterface = ArabicInterface Then
            PaymentTypesList.AddItem "‰ﬁœ«"
            PaymentTypesList.ItemData(PaymentTypesList.NewIndex) = 0
            PaymentTypesList.AddItem "‘Ìﬂ"
            PaymentTypesList.ItemData(PaymentTypesList.NewIndex) = 1
            PaymentTypesList.AddItem "’—«›"
            PaymentTypesList.ItemData(PaymentTypesList.NewIndex) = 2
            PaymentTypesList.AddItem " ÕÊÌ· »‰ﬂÌ"
            PaymentTypesList.ItemData(PaymentTypesList.NewIndex) = 3
            PaymentTypesList.AddItem "√Œ—Ï"
            PaymentTypesList.ItemData(PaymentTypesList.NewIndex) = 4
        Else
            PaymentTypesList.AddItem "Cash"
            PaymentTypesList.ItemData(PaymentTypesList.NewIndex) = 0
            PaymentTypesList.AddItem "Check"
            PaymentTypesList.ItemData(PaymentTypesList.NewIndex) = 1
            PaymentTypesList.AddItem "ATM"
            PaymentTypesList.ItemData(PaymentTypesList.NewIndex) = 2
            PaymentTypesList.AddItem "Bank Transfer"
            PaymentTypesList.ItemData(PaymentTypesList.NewIndex) = 3
            PaymentTypesList.AddItem "Others"
            PaymentTypesList.ItemData(PaymentTypesList.NewIndex) = 4
        End If
        '************************************************************************************************************************************
End Function
'1 ******************************************************************************************************************************************
Private Sub BSout_Click()
    If Me.BranchList.ListIndex > -1 Then
        Me.SelectedBranchList.AddItem BranchList.List(BranchList.ListIndex)
        SelectedBranchList.ItemData(SelectedBranchList.NewIndex) = BranchList.ItemData(BranchList.ListIndex)
    End If
End Sub
Private Sub BMout_Click()
    Dim i As Integer
    Me.SelectedBranchList.Clear
    If Me.BranchList.ListIndex > -1 Then
        For i = 0 To Me.BranchList.ListCount - 1
            Me.SelectedBranchList.AddItem BranchList.List(i)
            SelectedBranchList.ItemData(i) = BranchList.ItemData(i)
        Next i
    End If
End Sub
Private Sub BSin_Click()
    If SelectedBranchList.ListIndex > -1 Then
        SelectedBranchList.RemoveItem (SelectedBranchList.ListIndex)
    End If
End Sub
Private Sub BMin_Click()
    SelectedBranchList.Clear
End Sub
'2 ******************************************************************************************************************************************
Private Sub DSout_Click()
    If Me.DepList.ListIndex > -1 Then
        Me.SelectedDepList.AddItem DepList.List(DepList.ListIndex)
        SelectedDepList.ItemData(SelectedDepList.NewIndex) = DepList.ItemData(DepList.ListIndex)
    End If
End Sub
Private Sub DMout_Click()
    Dim i As Integer
    Me.SelectedDepList.Clear
    If Me.DepList.ListIndex > -1 Then
        For i = 0 To Me.DepList.ListCount - 1
            Me.SelectedDepList.AddItem DepList.List(i)
            SelectedDepList.ItemData(i) = DepList.ItemData(i)
        Next i
    End If
End Sub
Private Sub DSin_Click()
    If SelectedDepList.ListIndex > -1 Then
        SelectedDepList.RemoveItem (SelectedDepList.ListIndex)
    End If
End Sub
Private Sub DMin_Click()
    SelectedDepList.Clear
End Sub

'3 ******************************************************************************************************************************************
Private Sub SSout_Click()
    If Me.SecList.ListIndex > -1 Then
        Me.SelectedSecList.AddItem SecList.List(SecList.ListIndex)
        SelectedSecList.ItemData(SelectedSecList.NewIndex) = SecList.ItemData(SecList.ListIndex)
    End If
End Sub
Private Sub SMout_Click()
    Dim i As Integer
    Me.SelectedSecList.Clear
    If Me.SecList.ListIndex > -1 Then
        For i = 0 To Me.SecList.ListCount - 1
            Me.SelectedSecList.AddItem SecList.List(i)
            SelectedSecList.ItemData(i) = SecList.ItemData(i)
        Next i
    End If
End Sub
Private Sub SSin_Click()
    If SelectedBranchList.ListIndex > -1 Then
        SelectedSecList.RemoveItem (SelectedSecList.ListIndex)
    End If
End Sub
Private Sub SMin_Click()
    SelectedSecList.Clear
End Sub
'4 ******************************************************************************************************************************************
Private Sub NSout_Click()
    If Me.NationalityList.ListIndex > -1 Then
        Me.SelectedNationalityList.AddItem NationalityList.List(NationalityList.ListIndex)
        SelectedNationalityList.ItemData(SelectedNationalityList.NewIndex) = NationalityList.ItemData(NationalityList.ListIndex)
    End If
End Sub
Private Sub NMout_Click()
    Dim i As Integer
    Me.SelectedNationalityList.Clear
    If Me.NationalityList.ListIndex > -1 Then
        For i = 0 To Me.NationalityList.ListCount - 1
            Me.SelectedNationalityList.AddItem NationalityList.List(i)
            SelectedNationalityList.ItemData(i) = NationalityList.ItemData(i)
        Next i
    End If
End Sub
Private Sub NSin_Click()
    If SelectedNationalityList.ListIndex > -1 Then
        SelectedNationalityList.RemoveItem (SelectedNationalityList.ListIndex)
    End If
End Sub
Private Sub NMin_Click()
    SelectedNationalityList.Clear
End Sub
'5 ******************************************************************************************************************************************
Private Sub JSout_Click()
    If Me.JobList.ListIndex > -1 Then
        Me.SelectedJobList.AddItem JobList.List(JobList.ListIndex)
        SelectedJobList.ItemData(SelectedJobList.NewIndex) = JobList.ItemData(JobList.ListIndex)
    End If
End Sub
Private Sub JMout_Click()
    Dim i As Integer
    Me.SelectedJobList.Clear
    If Me.JobList.ListIndex > -1 Then
        For i = 0 To Me.JobList.ListCount - 1
            Me.SelectedJobList.AddItem JobList.List(i)
            SelectedJobList.ItemData(i) = JobList.ItemData(i)
        Next i
    End If
End Sub
Private Sub JSin_Click()
    If SelectedJobList.ListIndex > -1 Then
        SelectedJobList.RemoveItem (SelectedJobList.ListIndex)
    End If
End Sub
Private Sub JMin_Click()
    SelectedJobList.Clear
End Sub
'6 ******************************************************************************************************************************************
Private Sub ESout_Click()
    If Me.EmpList.ListIndex > -1 Then
        Me.SelectedEmpList.AddItem EmpList.List(EmpList.ListIndex)
        SelectedEmpList.ItemData(SelectedEmpList.NewIndex) = EmpList.ItemData(EmpList.ListIndex)
    End If
End Sub
Private Sub EMout_Click()
    Dim i As Integer
    Me.SelectedEmpList.Clear
    If Me.EmpList.ListIndex > -1 Then
        For i = 0 To Me.EmpList.ListCount - 1
            Me.SelectedEmpList.AddItem EmpList.List(i)
            SelectedEmpList.ItemData(i) = EmpList.ItemData(i)
        Next i
    End If
End Sub
Private Sub ESin_Click()
    If SelectedEmpList.ListIndex > -1 Then
        SelectedEmpList.RemoveItem (SelectedEmpList.ListIndex)
    End If
End Sub
Private Sub EMin_Click()
    SelectedEmpList.Clear
End Sub
'7 ******************************************************************************************************************************************
Private Sub PSout_Click()
    If Me.PenaltyTypeList.ListIndex > -1 Then
        Me.SelectedPenaltyTypeList.AddItem PenaltyTypeList.List(PenaltyTypeList.ListIndex)
        SelectedPenaltyTypeList.ItemData(SelectedPenaltyTypeList.NewIndex) = PenaltyTypeList.ItemData(PenaltyTypeList.ListIndex)
    End If
End Sub
Private Sub PMout_Click()
    Dim i As Integer
    Me.SelectedPenaltyTypeList.Clear
    If Me.PenaltyTypeList.ListIndex > -1 Then
        For i = 0 To Me.PenaltyTypeList.ListCount - 1
            Me.SelectedPenaltyTypeList.AddItem PenaltyTypeList.List(i)
            SelectedPenaltyTypeList.ItemData(i) = PenaltyTypeList.ItemData(i)
        Next i
    End If
End Sub
Private Sub PSin_Click()
    If SelectedPenaltyTypeList.ListIndex > -1 Then
        SelectedPenaltyTypeList.RemoveItem (SelectedPenaltyTypeList.ListIndex)
    End If
End Sub
Private Sub PMin_Click()
    SelectedPenaltyTypeList.Clear
End Sub
'8 ******************************************************************************************************************************************
Private Sub CSout_Click()
    If Me.WorkCaseList.ListIndex > -1 Then
        Me.SelectedWorkCaseList.AddItem WorkCaseList.List(WorkCaseList.ListIndex)
        SelectedWorkCaseList.ItemData(SelectedWorkCaseList.NewIndex) = WorkCaseList.ItemData(WorkCaseList.ListIndex)
    End If
End Sub
Private Sub CMout_Click()
    Dim i As Integer
    Me.SelectedWorkCaseList.Clear
    If Me.WorkCaseList.ListIndex > -1 Then
        For i = 0 To Me.WorkCaseList.ListCount - 1
            Me.SelectedWorkCaseList.AddItem WorkCaseList.List(i)
            SelectedWorkCaseList.ItemData(i) = WorkCaseList.ItemData(i)
        Next i
    End If
End Sub
Private Sub CSin_Click()
    If SelectedWorkCaseList.ListIndex > -1 Then
        SelectedWorkCaseList.RemoveItem (SelectedWorkCaseList.ListIndex)
    End If
End Sub
Private Sub CMin_Click()
    SelectedWorkCaseList.Clear
End Sub
'9 ******************************************************************************************************************************************
Private Sub PySout_Click()
    If Me.PaymentTypesList.ListIndex > -1 Then
        Me.selectedPaymentTypesList.AddItem PaymentTypesList.List(PaymentTypesList.ListIndex)
        selectedPaymentTypesList.ItemData(selectedPaymentTypesList.NewIndex) = PaymentTypesList.ItemData(PaymentTypesList.ListIndex)
    End If
End Sub
Private Sub PyMout_Click()
    Dim i As Integer
    Me.selectedPaymentTypesList.Clear
    If Me.PaymentTypesList.ListIndex > -1 Then
        For i = 0 To Me.PaymentTypesList.ListCount - 1
            Me.selectedPaymentTypesList.AddItem PaymentTypesList.List(i)
            selectedPaymentTypesList.ItemData(i) = PaymentTypesList.ItemData(i)
        Next i
    End If
End Sub
Private Sub PySin_Click()
    If selectedPaymentTypesList.ListIndex > -1 Then
        selectedPaymentTypesList.RemoveItem (selectedPaymentTypesList.ListIndex)
    End If
End Sub
Private Sub PyMin_Click()
    selectedPaymentTypesList.Clear
End Sub



Private Sub PySout2_Click()
    If Me.EmpLocationsList.ListIndex > -1 Then
        Me.SelectEmpLocations.AddItem EmpLocationsList.List(EmpLocationsList.ListIndex)
        SelectEmpLocations.ItemData(SelectEmpLocations.NewIndex) = EmpLocationsList.ItemData(EmpLocationsList.ListIndex)
    End If
End Sub
Private Sub PyMout2_Click()
    Dim i As Integer
    Me.SelectEmpLocations.Clear
    If Me.EmpLocationsList.ListIndex > -1 Then
        For i = 0 To Me.EmpLocationsList.ListCount - 1
            Me.SelectEmpLocations.AddItem EmpLocationsList.List(i)
            SelectEmpLocations.ItemData(i) = EmpLocationsList.ItemData(i)
        Next i
    End If
End Sub
Private Sub PySin2_Click()
    If SelectEmpLocations.ListIndex > -1 Then
        SelectEmpLocations.RemoveItem (SelectEmpLocations.ListIndex)
    End If
End Sub
Private Sub PyMin2_Click()
    SelectEmpLocations.Clear
End Sub

