VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmSaleBillInvestment 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14295
   Icon            =   "FrmSaleBillInvestment.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   14295
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   72
      Top             =   720
      Width           =   14295
      _cx             =   25215
      _cy             =   13361
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
      BackColor       =   -2147483624
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   14871017
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "«·—∆Ì”Ì…|»Ì«‰«  «·«ﬁ”«ÿ Ê«·œ›⁄« "
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
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·œ›⁄« "
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   7200
         Left            =   14940
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   45
         Width           =   14205
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√Ê· ﬁ”ÿ"
            Height          =   252
            Index           =   0
            Left            =   10440
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Œ— ﬁ”ÿ"
            Height          =   252
            Index           =   1
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœÊÌ"
            Height          =   252
            Index           =   2
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox TxtPaymentNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11640
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox TxtPeriod 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6840
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   240
            Width           =   705
         End
         Begin VB.ComboBox DcbPeriodsID 
            Height          =   315
            ItemData        =   "FrmSaleBillInvestment.frx":6852
            Left            =   5640
            List            =   "FrmSaleBillInvestment.frx":6854
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   675
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   240
            Width           =   4335
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
            Height          =   4755
            Left            =   120
            TabIndex        =   97
            Top             =   1080
            Width           =   13845
            _cx             =   24421
            _cy             =   8387
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   12
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmSaleBillInvestment.frx":6856
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
         Begin MSComCtl2.DTPicker FristDate 
            Height          =   270
            Left            =   9000
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   214695939
            CurrentDate     =   41640
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   315
            Left            =   5640
            TabIndex        =   30
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            Caption         =   "«÷«›…"
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
            ButtonImage     =   "FrmSaleBillInvestment.frx":691B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSComCtl2.DTPicker TempDate 
            Height          =   270
            Left            =   1560
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   214695939
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   37
            Left            =   11880
            TabIndex        =   122
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·œ›⁄« "
            Height          =   285
            Index           =   8
            Left            =   12840
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «Ê· œ›⁄Â"
            Height          =   285
            Index           =   9
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·› —Â »Ì‰ «·œ›⁄« "
            Height          =   285
            Index           =   11
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   21
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   285
            Index           =   10
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   6840
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Index           =   9
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   6720
            Width           =   2115
         End
      End
      Begin VB.Frame Frm2 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   7200
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   45
         Width           =   14205
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   4215
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   4080
            Width           =   14055
            Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
               Height          =   2355
               Left            =   120
               TabIndex        =   35
               Top             =   120
               Width           =   13845
               _cx             =   24421
               _cy             =   4154
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
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   12
               Cols            =   25
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSaleBillInvestment.frx":D17D
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
            Begin ImpulseButton.ISButton ISButton6 
               Height          =   330
               Left            =   12240
               TabIndex        =   123
               ToolTipText     =   "Õ–› «·ﬂ·"
               Top             =   2640
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   582
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ’›"
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
               ButtonImage     =   "FrmSaleBillInvestment.frx":D4F5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton4 
               Height          =   330
               Left            =   10320
               TabIndex        =   124
               ToolTipText     =   "Õ–› «·ﬂ·"
               Top             =   2640
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   582
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› «·ﬂ· "
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
               ButtonImage     =   "FrmSaleBillInvestment.frx":13D57
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   14
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   2760
               Visible         =   0   'False
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   13
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   2760
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   285
               Index           =   12
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   2760
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   6
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   3600
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   285
               Index           =   5
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   3600
               Width           =   1515
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Height          =   3735
            Left            =   120
            TabIndex        =   78
            Top             =   720
            Width           =   14055
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Height          =   3855
               Left            =   0
               TabIndex        =   79
               Top             =   0
               Width           =   14055
               Begin VB.CheckBox ChkComm 
                  Alignment       =   1  'Right Justify
                  Caption         =   " Õ„· ⁄·Ì «·⁄„Ì·"
                  Height          =   195
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.TextBox TxtNetComm 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   945
               End
               Begin VB.ComboBox DcbTyp 
                  Height          =   315
                  ItemData        =   "FrmSaleBillInvestment.frx":1A5B9
                  Left            =   1740
                  List            =   "FrmSaleBillInvestment.frx":1A5BB
                  RightToLeft     =   -1  'True
                  TabIndex        =   3
                  Top             =   720
                  Width           =   735
               End
               Begin VB.TextBox TxtRemarks 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   180
                  MaxLength       =   50
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   23
                  Top             =   3000
                  Width           =   5505
               End
               Begin VB.TextBox TxtRecordNo 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3540
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   2640
                  Width           =   2145
               End
               Begin VB.Frame Frame8 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÕœÊœ"
                  ForeColor       =   &H00C00000&
                  Height          =   1095
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   1440
                  Width           =   13935
                  Begin VB.TextBox TxteastWriiten 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   3480
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   15
                     Top             =   600
                     Width           =   2145
                  End
                  Begin VB.TextBox TxtwestWriiten 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   16
                     Top             =   600
                     Width           =   2145
                  End
                  Begin VB.TextBox TxtPriceSomW 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   6600
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   14
                     Top             =   600
                     Width           =   2145
                  End
                  Begin VB.TextBox TxtPriceHadW 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   9600
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   13
                     Top             =   600
                     Width           =   2145
                  End
                  Begin VB.TextBox txtWestlength 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   120
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   12
                     Top             =   240
                     Width           =   2145
                  End
                  Begin VB.TextBox txtSouthlength 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   6600
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   10
                     Top             =   240
                     Width           =   2145
                  End
                  Begin VB.TextBox txteastlength 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   3480
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   11
                     Top             =   240
                     Width           =   2145
                  End
                  Begin VB.TextBox txtnorthlength 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00FFFFFF&
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   9600
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   9
                     Top             =   240
                     Width           =   2145
                  End
                  Begin VB.Label Label23 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘„«·"
                     Height          =   255
                     Left            =   11400
                     RightToLeft     =   -1  'True
                     TabIndex        =   115
                     Top             =   720
                     Width           =   855
                  End
                  Begin VB.Label Label22 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘—ﬁ"
                     Height          =   255
                     Left            =   5520
                     RightToLeft     =   -1  'True
                     TabIndex        =   114
                     Top             =   600
                     Width           =   855
                  End
                  Begin VB.Label Label21 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Ã‰Ê»"
                     Height          =   255
                     Left            =   8520
                     RightToLeft     =   -1  'True
                     TabIndex        =   113
                     Top             =   600
                     Width           =   855
                  End
                  Begin VB.Label Label20 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "€—»"
                     Height          =   255
                     Left            =   2040
                     RightToLeft     =   -1  'True
                     TabIndex        =   112
                     Top             =   600
                     Width           =   1215
                  End
                  Begin VB.Label Label19 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÕœÊœ ﬂ «»Â"
                     Height          =   255
                     Left            =   12480
                     RightToLeft     =   -1  'True
                     TabIndex        =   111
                     Top             =   720
                     Width           =   1215
                  End
                  Begin VB.Label Label7 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÕœÊœ «—ﬁ«„"
                     Height          =   255
                     Left            =   12480
                     RightToLeft     =   -1  'True
                     TabIndex        =   110
                     Top             =   360
                     Width           =   1215
                  End
                  Begin VB.Label Label6 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "€—»"
                     Height          =   255
                     Left            =   2400
                     RightToLeft     =   -1  'True
                     TabIndex        =   109
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Ã‰Ê»"
                     Height          =   255
                     Left            =   8520
                     RightToLeft     =   -1  'True
                     TabIndex        =   108
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘—ﬁ"
                     Height          =   255
                     Left            =   5520
                     RightToLeft     =   -1  'True
                     TabIndex        =   107
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.Label Label3 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘„«·"
                     Height          =   255
                     Left            =   11400
                     RightToLeft     =   -1  'True
                     TabIndex        =   106
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.TextBox TxtPropertyDeed 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   210
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   1080
                  Width           =   3345
               End
               Begin VB.ComboBox DcbTypeSales 
                  Height          =   315
                  Left            =   11130
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   0
                  Top             =   360
                  Width           =   1545
               End
               Begin VB.ComboBox CboPayMentType 
                  Height          =   315
                  Left            =   7260
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   3000
                  Width           =   1545
               End
               Begin VB.TextBox TxtCusID 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   2640
                  Width           =   2145
               End
               Begin VB.TextBox Text9 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   10920
                  TabIndex        =   17
                  Top             =   2640
                  Width           =   1065
               End
               Begin VB.TextBox TxtDesLocation 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   5100
                  MaxLength       =   50
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   8
                  Top             =   1080
                  Width           =   7575
               End
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   9330
                  TabIndex        =   1
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox Txtcommission 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   210
                  RightToLeft     =   -1  'True
                  TabIndex        =   4
                  Top             =   720
                  Width           =   945
               End
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   11610
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   5
                  Top             =   720
                  Width           =   1065
               End
               Begin MSDataListLib.DataCombo DcbInvise 
                  Height          =   315
                  Left            =   13860
                  TabIndex        =   80
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   7515
                  _ExtentX        =   13256
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcCustomerType 
                  Height          =   315
                  Left            =   10410
                  TabIndex        =   21
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
                  Top             =   3000
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbCus 
                  Height          =   315
                  Left            =   7260
                  TabIndex        =   18
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   2640
                  Width           =   3675
                  _ExtentX        =   6482
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbLand 
                  Height          =   315
                  Left            =   5100
                  TabIndex        =   6
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   720
                  Width           =   6495
                  _ExtentX        =   11456
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbSales 
                  Height          =   315
                  Left            =   5100
                  TabIndex        =   2
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
                  Top             =   360
                  Width           =   4155
                  _ExtentX        =   7329
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdType 
                  Height          =   255
                  Index           =   0
                  Left            =   2280
                  TabIndex        =   132
                  Top             =   240
                  Width           =   1095
                  _Version        =   786432
                  _ExtentX        =   1931
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "„”«Â„…"
                  ForeColor       =   8388608
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdType 
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   133
                  Top             =   240
                  Width           =   1695
                  _Version        =   786432
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "«—÷ „„·Êﬂ…"
                  ForeColor       =   8388608
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   285
                  Index           =   19
                  Left            =   1080
                  RightToLeft     =   -1  'True
                  TabIndex        =   130
                  Top             =   720
                  Width           =   675
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   285
                  Index           =   0
                  Left            =   5790
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   3000
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÂÊÌ…"
                  Height          =   285
                  Index           =   17
                  Left            =   2190
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   2640
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·”Ã·"
                  Height          =   285
                  Index           =   16
                  Left            =   5880
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   2640
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
                  Height          =   285
                  Index           =   4
                  Left            =   8730
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   3000
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·»«∆⁄"
                  Height          =   285
                  Index           =   3
                  Left            =   9960
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   360
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄„Ì·"
                  Height          =   285
                  Index           =   1
                  Left            =   12210
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   2640
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·⁄„Ì·"
                  Height          =   285
                  Index           =   0
                  Left            =   12210
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   3000
                  Width           =   1890
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„⁄·Ê„«  ⁄‰ «·«—÷"
                  Height          =   285
                  Index           =   15
                  Left            =   12600
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   1080
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·’ﬂ"
                  Height          =   285
                  Index           =   11
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   1080
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·»«∆⁄"
                  Height          =   285
                  Index           =   1
                  Left            =   12450
                  TabIndex        =   84
                  Top             =   420
                  Width           =   1890
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„Ê·…"
                  Height          =   285
                  Index           =   3
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   510
                  Width           =   1515
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «·„”«Â„…"
                  Height          =   285
                  Index           =   7
                  Left            =   14250
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«—÷"
                  Height          =   285
                  Index           =   6
                  Left            =   12450
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   720
                  Width           =   1890
               End
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Height          =   735
            Left            =   120
            TabIndex        =   74
            Top             =   0
            Width           =   14055
            Begin VB.TextBox TxtBillNo 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox TxtSerial1 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   11400
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   240
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker XPDtbTrans 
               Height          =   315
               Left            =   9120
               TabIndex        =   32
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               Format          =   237502465
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo Dcbranch 
               Bindings        =   "FrmSaleBillInvestment.frx":1A5BD
               Height          =   315
               Left            =   5280
               TabIndex        =   33
               Top             =   240
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
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
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   255
               Index           =   0
               Left            =   4080
               TabIndex        =   127
               Top             =   240
               Width           =   975
               _Version        =   786432
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "›« Ê—…"
               ForeColor       =   8388608
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   255
               Index           =   1
               Left            =   3000
               TabIndex        =   128
               Top             =   240
               Width           =   975
               _Version        =   786432
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„—œÊœ« "
               ForeColor       =   8388608
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "»‰«¡ ⁄·Ï ›« Ê—… —ﬁ„"
               Height          =   285
               Index           =   18
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   285
               Index           =   7
               Left            =   7920
               TabIndex        =   77
               Top             =   240
               Width           =   1605
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·›« Ê—…"
               Height          =   285
               Index           =   4
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   240
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· «—ÌŒ"
               Height          =   285
               Index           =   2
               Left            =   10530
               TabIndex        =   75
               Top             =   255
               Width           =   885
            End
         End
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmSaleBillInvestment.frx":1A5D2
      Left            =   15480
      List            =   "FrmSaleBillInvestment.frx":1A5E2
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   0
      Width           =   14505
      Begin VB.TextBox TxtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox tXTRootAccount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   45
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmSaleBillInvestment.frx":1A5FB
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   46
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmSaleBillInvestment.frx":1A995
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   47
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmSaleBillInvestment.frx":1AD2F
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   48
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmSaleBillInvestment.frx":1B0C9
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "›« Ê—… «·„»Ì⁄« "
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
         Index           =   2
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmSaleBillInvestment.frx":1B463
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   52
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   15480
      TabIndex        =   53
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1425
      Left            =   0
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   8280
      Width           =   14235
      _cx             =   25109
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
      AutoSizeChildren=   0
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
      Begin VB.Frame Frame9 
         Height          =   690
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   134
         Top             =   0
         Width           =   4485
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   195
            Index           =   35
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   240
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   255
            Width           =   675
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   55
         Top             =   600
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   36
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSaleBillInvestment.frx":1C868
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8880
            TabIndex        =   38
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSaleBillInvestment.frx":230CA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   37
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSaleBillInvestment.frx":23464
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   39
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSaleBillInvestment.frx":29CC6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   40
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSaleBillInvestment.frx":2A060
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   41
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSaleBillInvestment.frx":2A5FA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   3960
            TabIndex        =   68
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   240
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
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
            ButtonImage     =   "FrmSaleBillInvestment.frx":2A994
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1920
            TabIndex        =   69
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   240
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSaleBillInvestment.frx":311F6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   10080
         TabIndex        =   61
         Top             =   120
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   570
         Left            =   120
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   -840
         Width           =   13965
         _cx             =   24633
         _cy             =   1005
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
      End
      Begin ImpulseButton.ISButton ISButton3 
         Height          =   330
         Left            =   4200
         TabIndex        =   119
         ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ButtonStyle     =   1
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
         ButtonImage     =   "FrmSaleBillInvestment.frx":31590
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   13080
         TabIndex        =   62
         Top             =   120
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   15600
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSaleBillInvestment.frx":37DF2
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSaleBillInvestment.frx":3818C
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSaleBillInvestment.frx":38526
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSaleBillInvestment.frx":388C0
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSaleBillInvestment.frx":38C5A
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSaleBillInvestment.frx":38FF4
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSaleBillInvestment.frx":3938E
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSaleBillInvestment.frx":39928
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   63
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ÕœÌÀ"
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
      ButtonImage     =   "FrmSaleBillInvestment.frx":39CC2
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… "
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
      ButtonImage     =   "FrmSaleBillInvestment.frx":40524
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   67
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
      BackColor       =   14871017
      FontSize        =   9.75
      FontName        =   "Arial"
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmSaleBillInvestment.frx":46D86
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·„” Œœ„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   15480
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmSaleBillInvestment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
                Dim RootAccount1 As String
                        Dim RootAccount2 As String
                        Dim RootAccount3 As String
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim Account_Code_dynamic As String
 Dim CommissionAccount As String
 Dim ii As Long
 Public LonRow As Double
Public LngCol As Double


Sub GetInformationCustomer(Optional Cus_ID As Double)
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Dim sql As String
If Cus_ID <> 0 Then
sql = "select TypeInvestor ,CustomerTypeID ,CustGID ,RecordNo from TblCustemers where CusID =" & Cus_ID & " "
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
DcCustomerType.BoundText = IIf(IsNull(Rs6("TypeInvestor").value), "", Rs6("TypeInvestor").value)
TxtRecordNo.text = IIf(IsNull(Rs6("CustGID").value), "", Rs6("CustGID").value)
TxtCusID.text = IIf(IsNull(Rs6("CustGID").value), "", Rs6("CustGID").value)
Else
TxtCusID = ""
TxtRecordNo = ""
DcCustomerType.BoundText = ""
End If
End If
End Sub

Private Sub CboPayMentType_Change()
If val(Me.CboPayMentType.ListIndex) = 1 Then
Frame7.Enabled = False
Else
Frame7.Enabled = True
End If
End Sub

Private Sub CboPayMentType_Click()
CboPayMentType_Change
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.text, , 200
End Sub

Private Sub DcbCus_Change()
DcbCus_Click (0)
End Sub

Private Sub DcbCus_Click(Area As Integer)

  If val(DcbCus.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCus.BoundText, EmpCode
    Me.Text9.text = EmpCode
If Me.TxtModFlg.text <> "R" Then
If val(DcbCus.BoundText) <> 0 Then
GetInformationCustomer val(DcbCus.BoundText)
End If
End If
End Sub

Private Sub Dcbranch_Change()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.text = ""
End If
End Sub
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim DebitAcc As String
Dim CreditAcc As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords



    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
   
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
       ' Msg = EleHeader.Caption & " ??? " & txtid & " EEC??I" & Date
 If TxtRemarks.text = "" Then
        Msg = "„’—Ê›«  «· ÿÊÌ— »—ﬁ„ " & TxtSerial1 & "  ··„”«Â„…  " & DcbInvise.text & "  ··«—÷ " & DcbLand.text
        Msg = "›« Ê—… —ﬁ„ " & TxtSerial1 & "  "
 Else
 Msg = Me.TxtRemarks.text
 End If
 
        
 
 notes_id = general_noteid

  
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim ProjectID As Integer
    
    BranchID = 1
    
    With GridInstallments
Dim netval As Double
Dim netvalMinusCommission As Double
Dim netvalPlusCommission As Double
Dim Comm As Double
Dim substr As String
line_no = 1
        For i = .FixedRows To .rows - 1
    BranchID = val(dcBranch.BoundText)
  
  
            If val(.TextMatrix(i, .ColIndex("Net"))) > 0 And .TextMatrix(i, .ColIndex("CodeUnit")) <> "" Then     'C?C??? C???E??E IC??
    Comm = val(.TextMatrix(i, .ColIndex("Comm")))
 
            netval = val(.TextMatrix(i, .ColIndex("Net")))
          netval = Round(netval, 2)
                  netvalMinusCommission = val(.TextMatrix(i, .ColIndex("Net"))) - val(.TextMatrix(i, .ColIndex("Comm")))
            netvalMinusCommission = Round(netvalMinusCommission, 2)
            
            netvalPlusCommission = val(.TextMatrix(i, .ColIndex("Net"))) + val(.TextMatrix(i, .ColIndex("Comm")))
            netvalPlusCommission = Round(netvalPlusCommission, 2)
             
             
    
                'substr = " ·  " & .TextMatrix(i, .ColIndex("InvesName")) & "  " & .TextMatrix(i, .ColIndex("unit")) & "  · " & .TextMatrix(i, .ColIndex("CodeUnit"))
                
                substr = " ·  " & .TextMatrix(i, .ColIndex("unit")) & "   " & .TextMatrix(i, .ColIndex("CodeUnit"))
'                substr = substr & "  " & "  · " & .TextMatrix(i, .ColIndex("PartName"))
                substr = substr & "  " & "  —ﬁ„ «·ﬁÿ⁄… " & .TextMatrix(i, .ColIndex("BlockName"))
               substr = substr & "  " & "  «·„”«Õ… " & .TextMatrix(i, .ColIndex("Area"))
               Msg = "›« Ê—… —ﬁ„ " & TxtSerial1 & "  "
               substr = Msg & substr
             If Rd(0).value = True Then 'sales
            DebitAcc = GetMyAccountCode("TblCustemers", "CusID", val(DcbCus.BoundText))
                       If RdType(0).value = True Then '„”«Â„…
                        CreditAcc = GetMyAccountCode("TblActivateInvestment", "id", val(.TextMatrix(i, .ColIndex("InvesID"))), "Account_Code5")
                        Else '«—÷ „„·Êﬂ…
                        CreditAcc = GetMyAccountCode("TblBuyLanReEst", "id", val(.TextMatrix(i, .ColIndex("InvesID"))))
            
                        End If
            
             Else 'sales return
               
                        If RdType(0).value = True Then '„”«Â„…
                        DebitAcc = GetMyAccountCode("TblActivateInvestment", "id", val(.TextMatrix(i, .ColIndex("InvesID"))), "Account_Code5")
                        Else '«—÷ „„·Êﬂ…
                        DebitAcc = GetMyAccountCode("TblBuyLanReEst", "id", val(.TextMatrix(i, .ColIndex("InvesID"))))
            
                        End If
                        
             
             
             CreditAcc = GetMyAccountCode("TblCustemers", "CusID", val(DcbCus.BoundText))
            
 
             End If
             
 Dim CustomerValue As Double
 Dim SalesValue As Double
 Dim commvalue As Double
 
' Dim netval As Double
'Dim netvalMinusCommission As Double
'Dim netvalPlusCommission As Double

 If ChkComm.value = vbChecked Then ' Õ„· ⁄·Ì «·⁄„Ì·
 
CustomerValue = netvalPlusCommission
 SalesValue = netval
  commvalue = Comm
 Else
 CustomerValue = netval
 SalesValue = netvalMinusCommission
 commvalue = Comm
 End If
 If Me.TxtRemarks.text = "" Then
Msg = substr
 End If
 
 
               If Rd(0).value = True Then 'sales
                If ModAccounts.AddNewDev(LngDevID, line_no, DebitAcc, CustomerValue, 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                  
                If ModAccounts.AddNewDev(LngDevID, line_no, CreditAcc, SalesValue, 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
          If val(TxtNetComm.text) > 0 Then '«·⁄„Ê·« 
    
             line_no = line_no + 1
                          If ModAccounts.AddNewDev(LngDevID, line_no, CommissionAccount, commvalue, 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                 line_no = line_no + 1
                 
            End If
                 
                Else 'return
                
                      If ModAccounts.AddNewDev(LngDevID, line_no, DebitAcc, SalesValue, 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                  
         If val(TxtNetComm.text) > 0 Then '«·⁄„Ê·« 
    
             line_no = line_no + 1
                          If ModAccounts.AddNewDev(LngDevID, line_no, CommissionAccount, commvalue, 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                 line_no = line_no + 1
                 
            End If
                If ModAccounts.AddNewDev(LngDevID, line_no, CreditAcc, CustomerValue, 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
          
          
                End If
            
                 
                
            End If
     
     
     Next i
     
     End With
           
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function
 Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
If Rd(0).value = True Then
des = " „»Ì⁄«  «·«—«÷Ì »—ﬁ„ " & TxtSerial1 & "  ··⁄„Ì·  " & DcbCus.text & "  ··»«∆⁄  " & DcbSales.text
If RdType(0).value = True Then '„”«Â„« 
 notytype = 9005
 Else
 notytype = 9007
 End If
ElseIf Rd(1).value = True Then
des = " „»Ì⁄«  ‘—«¡ «·«—«÷Ì »—ﬁ„ " & TxtSerial1 & "  ··⁄„Ì·" & DcbCus.text & "  ··»«∆⁄  " & DcbSales.text
If RdType(0).value = True Then '„”«Â„« 
 notytype = 9006
 Else
 notytype = 9008
 End If
 
End If

Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim sql As String
tablename = "TblSaleBilllInvestment"
Filedname = "ID"
NoteSerial1 = TxtSerial1.text
Notevalue = val(lbl(13).Caption)
 

 BranchID = val(dcBranch.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TXTNoteID.text = NoteID
                                                     TxtNoteSerial.text = NoteSerial
                                     Else
                                                 If TXTNoteID.text = "" Or TxtNoteSerial.text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TXTNoteID.text = NoteID
                                                                TxtNoteSerial.text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TXTNoteID.text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TXTNoteID.text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent
 

     End If

End Function


Private Sub DcbSales_Change()
Dim EmpCode  As String
If val(Me.DcbTypeSales.ListIndex) = 0 Then
If val(Me.DcbSales.BoundText) = 0 Then Exit Sub
           Me.TxtSearchCode.text = get_EMPLOYEE_Data(val(Me.DcbSales.BoundText), "Fullcode")
Else
  If val(DcbSales.BoundText) = 0 Then Exit Sub
    GetTblCustemersCode , , DcbSales.BoundText, EmpCode
    Me.TxtSearchCode.text = EmpCode
End If
End Sub

Private Sub DcbSales_Click(Area As Integer)
DcbSales_Change
End Sub

Private Sub DcbTyp_Change()
If Me.TxtModFlg.text <> "R" Then
If val(DcbTyp.ListIndex) = 1 Then
lbl(19).Caption = "‰”»…"
TxtNetComm.text = Round((val(Txtcommission.text) * val(lbl(13).Caption)) / 100, 2)
Else
TxtNetComm.text = Txtcommission.text
lbl(19).Caption = "ﬁÌ„…"
End If
End If
End Sub

Private Sub DcbTyp_Click()
DcbTyp_Change
End Sub

Private Sub DcbTypeSales_Change()

Dim Dcombos As ClsDataCombos
Set Dcombos = New ClsDataCombos
            Dcombos.ClearMyDataCombo Me.DcbSales
    Select Case DcbTypeSales.ListIndex

        Case 0
            Set Dcombos = New ClsDataCombos
            Dcombos.GetEmployees DcbSales
          Case 1
          Set Dcombos = New ClsDataCombos
            Dcombos.GetCustomersSuppliers 2, Me.DcbSales, True
          End Select
End Sub

Private Sub DcbTypeSales_Click()
DcbTypeSales_Change
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
   If SystemOptions.UserInterface = ArabicInterface Then
                GridInstallments.ColComboList(GridInstallments.ColIndex("TypeDis")) = "#1; ·«ÌÊÃœ Œ’„|#2; Œ’„ »ﬁÌ„…|#3; Œ’„ ‰”»…"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               GridInstallments.ColComboList(GridInstallments.ColIndex("TypeDis")) = "#1;No Discount |#2;No value Discount |#3; Percent Discount "
            End If
     If SystemOptions.UserInterface = ArabicInterface Then
     With DcbTyp
     .AddItem "ﬁÌ„…"
     .AddItem "‰”»…"
     End With
     Else
      With DcbTyp
     .AddItem "Value"
     .AddItem "Percentage"
     End With
     End If
    conection = "select * from TblSaleBilllInvestment order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me
      If SystemOptions.UserInterface = ArabicInterface Then
    With CboPayMentType
    .Clear
    .AddItem "‰ﬁœÌ"
    .AddItem "«Ã·"
    End With
       With DcbPeriodsID
    .Clear
    .AddItem "ÌÊ„"
    .AddItem "‘Â—"
    .AddItem "”‰…"
    End With
 Else
     With DcbPeriodsID
    .Clear
    .AddItem "Day"
    .AddItem "Month"
    .AddItem "Year"
    End With
    With CboPayMentType
    .Clear
    .AddItem "Cash"
    .AddItem "Credit"
    End With
End If
 If SystemOptions.UserInterface = ArabicInterface Then
 With DcbTypeSales
 .Clear
 .AddItem "„ÊŸ›"
 .AddItem "„Ê—œ"
 End With
 Else
  With DcbTypeSales
 .Clear
 .AddItem "Employee"
 .AddItem "Vendor"
 End With
 End If
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  where type=20 or type =1   order by CusName"
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  where type=20 or type =1   order by CusNamee"
    End If

    fill_combo DcbCus, My_SQL
 
    Dim Dcombos As New ClsDataCombos
   ' Dcombos.GetCustomersSuppliers 1, Me.DcbCus, True
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetUsers Me.DCboUserName
    'Dcombos.GetBuyLandRealEstate DcbLand
    Dcombos.GetInvestmentActive Me.DcbInvise
    Dcombos.GetInvStoreType Me.DcCustomerType
    'Dcombos.GetCustomerType Me.DcCustomerType
    BtnLast_Click

    ShowTip
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
   Me.Refresh
   FiLLTXT
ErrTrap:
End Sub

' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
    Dim i As Integer
             If Me.TxtModFlg.text = "E" Then

                 StrSQL = "Delete From TblSaleBilllInvestmentDet Where SBINVID =" & val(TxtSerial1.text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Update TblDivInvestInformation  set SalesPayed=Null,SalesBlocPayed=null,SalID=0 where SalID=" & val(TxtSerial1.text) & "  "
       Cn.Execute StrSQL
          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords

              
              End If
    
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
     
    If ChkComm.value = vbChecked Then
    RsSavRec.Fields("CusComm").value = 1
    Else
    RsSavRec.Fields("CusComm").value = 0
    End If
    
  
    RsSavRec.Fields("BranchID").value = val(Me.dcBranch.BoundText)
    RsSavRec.Fields("SellerType").value = val(Me.DcbTypeSales.ListIndex)
    RsSavRec.Fields("SellerID").value = val(Me.DcbSales.BoundText)
    RsSavRec.Fields("InvesID").value = val(Me.DcbInvise.BoundText)
    RsSavRec.Fields("LandID").value = val(Me.DcbLand.BoundText)
    RsSavRec.Fields("commission").value = val(Txtcommission.text)
    RsSavRec.Fields("DesLocation").value = TxtDesLocation.text
    RsSavRec.Fields("Remarks").value = TxtRemarks.text
    RsSavRec.Fields("PropertyDeed").value = TxtPropertyDeed.text
    RsSavRec.Fields("NorthlengthStr").value = TxtPriceHadW.text
    RsSavRec.Fields("SouthlengthStr").value = TxtPriceSomW.text
    RsSavRec.Fields("EastlengthStr").value = TxteastWriiten.text
    RsSavRec.Fields("WestlengthStr").value = TxtwestWriiten.text
    RsSavRec.Fields("Northlength").value = val(TxtNorthLength.text)
    RsSavRec.Fields("Southlength").value = val(TxtSouthLength.text)
    RsSavRec.Fields("Eastlength").value = val(TxtEastLength.text)
    RsSavRec.Fields("Westlength").value = val(txtWestlength.text)
    RsSavRec.Fields("Cus_Tpe").value = val(Me.DcCustomerType.BoundText)
    RsSavRec.Fields("Cus_ID").value = val(Me.DcbCus.BoundText)
    RsSavRec.Fields("Payment").value = val(Me.CboPayMentType.ListIndex)
    RsSavRec.Fields("RecordNo").value = (Me.TxtRecordNo.text)
    RsSavRec.Fields("CusID").value = (Me.TxtCusID.text)
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.Fields("FristDate").value = FristDate.value
    RsSavRec.Fields("PaymentNo").value = val(Me.TxtPaymentNo.text)
    RsSavRec.Fields("Period").value = val(Me.txtPeriod.text)
    RsSavRec.Fields("PeriodType").value = val(Me.DcbPeriodsID.ListIndex)
    RsSavRec.Fields("RemarkPay").value = (Me.Text11.text)
     If opt(0).value = True Then
    RsSavRec.Fields("Typepartial").value = 0
    ElseIf opt(1).value = True Then
    RsSavRec.Fields("Typepartial").value = 1
    ElseIf opt(2).value = True Then
    RsSavRec.Fields("Typepartial").value = 2
    Else
    RsSavRec.Fields("Typepartial").value = Null
    
    End If
     RsSavRec.Fields("TypeCom").value = val(Me.DcbTyp.ListIndex)
     
     RsSavRec.Fields("BillNo").value = val(TxtBillNo.text)
     If Rd(1).value = True Then
      RsSavRec.Fields("TypeRetSal").value = 1
     Else
      RsSavRec.Fields("TypeRetSal").value = 0
      End If
    RsSavRec.Fields("TotaPtofit").value = val(lbl(14).Caption)
  

  If RdType(1).value = True Then
  RsSavRec.Fields("TypDiv").value = 1
  Else
  RsSavRec.Fields("TypDiv").value = 0
  End If
  RsSavRec.Fields("NetComm").value = val(Me.TxtNetComm.text)
  RsSavRec.update
  
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblSaleBilllInvestmentDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim netCommissionValue As Double
    Dim netvalue As Double
    netCommissionValue = val(TxtNetComm)
    netvalue = val(lbl(13).Caption)
   
    Dim str2 As String
    With Me.GridInstallments
       For i = .FixedRows To .rows - 1
       If val(.TextMatrix(i, .ColIndex("BlockID"))) <> 0 Then
          .TextMatrix(i, .ColIndex("CommP")) = ((.TextMatrix(i, .ColIndex("Net")) / netvalue))
           .TextMatrix(i, .ColIndex("Comm")) = Round(netCommissionValue * .TextMatrix(i, .ColIndex("CommP")), 2)
       RsDevsub.AddNew
         If RdType(1).value = True Then
  RsDevsub("TypDiv").value = 1
  Else
  RsDevsub("TypDiv").value = 0
  End If
                RsDevsub("SBINVID").value = val(Me.TxtSerial1.text)
                RsDevsub("TypeTrns").value = 0
                RsDevsub("CommP").value = IIf((.TextMatrix(i, .ColIndex("CommP"))) = "", Null, val(.TextMatrix(i, .ColIndex("CommP"))))
                RsDevsub("Comm").value = IIf((.TextMatrix(i, .ColIndex("Comm"))) = "", Null, val(.TextMatrix(i, .ColIndex("Comm"))))
                
                RsDevsub("InvesID").value = IIf((.TextMatrix(i, .ColIndex("InvesID"))) = "", Null, val(.TextMatrix(i, .ColIndex("InvesID"))))
                RsDevsub("unitId").value = IIf((.TextMatrix(i, .ColIndex("unitId"))) = "", Null, val(.TextMatrix(i, .ColIndex("unitId"))))
                RsDevsub("unitunidpart").value = IIf((.TextMatrix(i, .ColIndex("unitunidpart"))) = "", Null, val(.TextMatrix(i, .ColIndex("unitunidpart"))))
                RsDevsub("DivID").value = IIf((.TextMatrix(i, .ColIndex("DivID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DivID"))))
                RsDevsub("Area").value = IIf((.TextMatrix(i, .ColIndex("Area"))) = "", Null, val(.TextMatrix(i, .ColIndex("Area"))))
                RsDevsub("Valu").value = IIf((.TextMatrix(i, .ColIndex("Valu"))) = "", Null, val(.TextMatrix(i, .ColIndex("Valu"))))
                RsDevsub("Total").value = IIf((.TextMatrix(i, .ColIndex("Total"))) = "", Null, val(.TextMatrix(i, .ColIndex("Total"))))
                RsDevsub("DisValue").value = IIf((.TextMatrix(i, .ColIndex("DisValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("DisValue"))))
                RsDevsub("TypeDis").value = IIf((.TextMatrix(i, .ColIndex("TypeDis"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypeDis"))))
                RsDevsub("Net").value = IIf((.TextMatrix(i, .ColIndex("Net"))) = "", Null, val(.TextMatrix(i, .ColIndex("Net"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remarks"))) = "", Null, (.TextMatrix(i, .ColIndex("Remarks"))))
                RsDevsub("MeterValue").value = IIf((.TextMatrix(i, .ColIndex("MeterValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("MeterValue"))))
                RsDevsub("TotalCost").value = IIf((.TextMatrix(i, .ColIndex("TotalCost"))) = "", Null, val(.TextMatrix(i, .ColIndex("TotalCost"))))
                RsDevsub("Profit").value = IIf((.TextMatrix(i, .ColIndex("Profit"))) = "", Null, val(.TextMatrix(i, .ColIndex("Profit"))))
                RsDevsub("BlockID").value = IIf((.TextMatrix(i, .ColIndex("BlockID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BlockID"))))
                RsDevsub("TotalArea").value = IIf((.TextMatrix(i, .ColIndex("TotalArea"))) = "", Null, val(.TextMatrix(i, .ColIndex("TotalArea"))))
                RsDevsub("PartName").value = IIf((.TextMatrix(i, .ColIndex("PartName"))) = "", Null, (.TextMatrix(i, .ColIndex("PartName"))))
                RsDevsub("CodeUnit").value = IIf((.TextMatrix(i, .ColIndex("CodeUnit"))) = "", Null, (.TextMatrix(i, .ColIndex("CodeUnit"))))
       RsDevsub.update
       If Rd(0).value = True Then
       If val((.TextMatrix(i, .ColIndex("BlockID")))) <> 0 And (.TextMatrix(i, .ColIndex("BlockName"))) <> "" Then
         sql = "Update TblDivInvestInformation  set SalesBlocPayed=1,SalID=" & val(TxtSerial1.text) & " where ID= " & val(.TextMatrix(i, .ColIndex("BlockID"))) & ""
       Cn.Execute sql
      End If
      ElseIf Rd(1).value = True Then
             If val((.TextMatrix(i, .ColIndex("BlockID")))) <> 0 And (.TextMatrix(i, .ColIndex("BlockName"))) <> "" Then
       StrSQL = "Update TblSaleBilllInvestmentDet  set ReturnSal=1 where SBINVID=" & val(TxtBillNo.text) & " and BlockID= " & val(.TextMatrix(i, .ColIndex("BlockID"))) & ""
       Cn.Execute StrSQL
         sql = "Update TblDivInvestInformation  set SalesBlocPayed=null,SalID=" & val(TxtBillNo.text) & " where ID= " & val(.TextMatrix(i, .ColIndex("BlockID"))) & ""
       Cn.Execute sql
      End If
      
      End If
      End If
     Next i
    End With
   ''//////////////////////Paymnts
         Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblSaleBilllInvestmentDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    With Me.VSFlexGrid1
       For i = .FixedRows To .rows - 1
       If val(.TextMatrix(i, .ColIndex("PaymentNo"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("SBINVID").value = val(Me.TxtSerial1.text)
                RsDevsub("TypeTrns").value = 1
                RsDevsub("PartID").value = IIf((.TextMatrix(i, .ColIndex("PaymentNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("PaymentNo"))))
                RsDevsub("Valu").value = IIf((.TextMatrix(i, .ColIndex("PaymentValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("PaymentValue"))))
                RsDevsub("FristDate").value = IIf((.TextMatrix(i, .ColIndex("DatePayment"))) = "", Null, (.TextMatrix(i, .ColIndex("DatePayment"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remrk"))) = "", Null, (.TextMatrix(i, .ColIndex("Remrk"))))
       RsDevsub.update
      End If
     Next i
    End With
    createVoucher
    
    
'''///////////////
  
      Select Case Me.TxtModFlg.text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
                FiLLTXT
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub


' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
       Me.TXTNoteID.text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)

Me.TxtNoteSerial.text = IIf(IsNull(RsSavRec("NoteSerial").value), "", RsSavRec("NoteSerial").value)


If IsNull(RsSavRec.Fields("CusComm").value) Then
ChkComm.value = vbUnchecked
Else
ChkComm.value = IIf((RsSavRec.Fields("CusComm").value) = 0, vbUnchecked, vbChecked)
End If

    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbTypeSales.ListIndex = IIf(IsNull(RsSavRec.Fields("SellerType").value), -1, RsSavRec.Fields("SellerType").value)
    Me.DcbSales.BoundText = IIf(IsNull(RsSavRec.Fields("SellerID").value), "", RsSavRec.Fields("SellerID").value)
    Me.DcbInvise.BoundText = IIf(IsNull(RsSavRec.Fields("InvesID").value), "", RsSavRec.Fields("InvesID").value)
    Me.DcbLand.BoundText = IIf(IsNull(RsSavRec.Fields("LandID").value), "", RsSavRec.Fields("LandID").value)
    Txtcommission.text = IIf(IsNull(RsSavRec.Fields("commission").value), 0, RsSavRec.Fields("commission").value)
    TxtDesLocation.text = IIf(IsNull(RsSavRec.Fields("DesLocation").value), "", RsSavRec.Fields("DesLocation").value)
    TxtRemarks.text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    TxtPropertyDeed.text = IIf(IsNull(RsSavRec.Fields("PropertyDeed").value), "", RsSavRec.Fields("PropertyDeed").value)
    TxtPriceHadW.text = IIf(IsNull(RsSavRec.Fields("NorthlengthStr").value), "", RsSavRec.Fields("NorthlengthStr").value)
    TxtPriceSomW.text = IIf(IsNull(RsSavRec.Fields("SouthlengthStr").value), "", RsSavRec.Fields("SouthlengthStr").value)
    TxteastWriiten.text = IIf(IsNull(RsSavRec.Fields("EastlengthStr").value), "", RsSavRec.Fields("EastlengthStr").value)
    TxtwestWriiten.text = IIf(IsNull(RsSavRec.Fields("WestlengthStr").value), "", RsSavRec.Fields("WestlengthStr").value)
    TxtNorthLength.text = IIf(IsNull(RsSavRec.Fields("Northlength").value), 0, RsSavRec.Fields("Northlength").value)
    TxtSouthLength.text = IIf(IsNull(RsSavRec.Fields("Southlength").value), 0, RsSavRec.Fields("Southlength").value)
    TxtEastLength.text = IIf(IsNull(RsSavRec.Fields("Eastlength").value), 0, RsSavRec.Fields("Eastlength").value)
    txtWestlength.text = IIf(IsNull(RsSavRec.Fields("Westlength").value), 0, RsSavRec.Fields("Westlength").value)
    DcCustomerType.BoundText = IIf(IsNull(RsSavRec.Fields("Cus_Tpe").value), "", RsSavRec.Fields("Cus_Tpe").value)
    DcbCus.BoundText = IIf(IsNull(RsSavRec.Fields("Cus_ID").value), "", RsSavRec.Fields("Cus_ID").value)
    CboPayMentType.ListIndex = IIf(IsNull(RsSavRec.Fields("Payment").value), -1, RsSavRec.Fields("Payment").value)
    TxtRecordNo.text = IIf(IsNull(RsSavRec.Fields("RecordNo").value), "", RsSavRec.Fields("RecordNo").value)
    TxtCusID.text = IIf(IsNull(RsSavRec.Fields("CusID").value), "", RsSavRec.Fields("CusID").value)
    TxtPaymentNo.text = IIf(IsNull(RsSavRec.Fields("PaymentNo").value), "", RsSavRec.Fields("PaymentNo").value)
    txtPeriod.text = IIf(IsNull(RsSavRec.Fields("Period").value), "", RsSavRec.Fields("Period").value)
    DcbPeriodsID.ListIndex = IIf(IsNull(RsSavRec.Fields("PeriodType").value), -1, RsSavRec.Fields("PeriodType").value)
    Text11.text = IIf(IsNull(RsSavRec.Fields("RemarkPay").value), "", RsSavRec.Fields("RemarkPay").value)
    FristDate.value = IIf(IsNull(RsSavRec.Fields("FristDate").value), Date, RsSavRec.Fields("FristDate").value)
    lbl(14).Caption = IIf(IsNull(RsSavRec.Fields("TotaPtofit").value), 0, RsSavRec.Fields("TotaPtofit").value)
    TxtNetComm.text = IIf(IsNull(RsSavRec.Fields("NetComm").value), 0, RsSavRec.Fields("NetComm").value)
    
        If Not (IsNull(RsSavRec.Fields("Typepartial").value)) Then
    If RsSavRec.Fields("Typepartial").value = 0 Then
    opt(0).value = True
    ElseIf RsSavRec.Fields("Typepartial").value = 1 Then
    opt(1).value = True
    ElseIf RsSavRec.Fields("Typepartial").value = 2 Then
    opt(2).value = True
    End If
    End If
    DcbTyp.ListIndex = IIf(IsNull(RsSavRec.Fields("TypeCom").value), -1, RsSavRec.Fields("TypeCom").value)
    If Me.TxtModFlg.text = "R" Then
    TxtBillNo.text = IIf(IsNull(RsSavRec.Fields("BillNo").value), 0, RsSavRec.Fields("BillNo").value)
    If Not (IsNull(RsSavRec.Fields("TypeRetSal").value)) Then
    If RsSavRec.Fields("TypeRetSal").value = 1 Then
    Rd(1).value = True
    Else
    Rd(0).value = True
    End If
    Else
    Rd(0).value = True
    End If
    End If
    
    If Not (IsNull(RsSavRec.Fields("TypDiv").value)) Then
    If RsSavRec.Fields("TypDiv").value = 1 Then
    RdType(1).value = True
    Else
    RdType(0).value = True
    End If
    Else
    RdType(0).value = True
    End If
  
    
    ''''
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData
RelinGrid
ErrTrap:
End Sub
Function CheckReturn(Optional ID As Double = 0) As Boolean
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
CheckReturn = False
sql = "SELECT    SBINVID "
sql = sql & " From dbo.TblSaleBilllInvestmentDet"
sql = sql & " Where   SBINVID=" & ID & " and ReturnSal=1 "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
CheckReturn = True
Else
CheckReturn = False
End If
End Function

Function CheckSaleorReturn(Optional ID As Double = 0) As Boolean
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
CheckSaleorReturn = False
sql = "SELECT    id "
sql = sql & " From dbo.TblSaleBilllInvestment"
sql = sql & " Where   ID=" & ID & " and TypeRetSal=1"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
CheckSaleorReturn = True
Else
CheckSaleorReturn = False
End If
End Function
Function CheckDetrputed(Optional ID As Double = 0) As Boolean
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
CheckDetrputed = False
sql = "SELECT    SBINVID "
sql = sql & " From dbo.TblSaleBilllInvestmentDet"
sql = sql & " Where   SBINVID=" & ID & " and Payed=1"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
CheckDetrputed = True
Else
CheckDetrputed = False
End If
End Function
Function GetTOtalArea(Optional ID As Double) As Double
Dim Rs5 As ADODB.Recordset
Dim sql As String
Set Rs5 = New ADODB.Recordset
sql = " SELECT     SUM(Area) AS SumTotalArea, InviseOrder"
sql = sql & " From dbo.TblActivateInvestment"
sql = sql & " Where (InviseOrder = " & ID & ") "
sql = sql & " GROUP BY InviseOrder"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
GetTOtalArea = IIf(IsNull(Rs5("SumTotalArea").value), 0, Rs5("SumTotalArea").value)
Else
GetTOtalArea = 0
End If
End Function

Function GetDividArea(Optional InvID As Double) As Double
Dim Rs5 As ADODB.Recordset
Dim sql As String
Set Rs5 = New ADODB.Recordset
sql = "SELECT     SUM(Area) AS SumArea, InvID"
sql = sql & " From dbo.TblDivInvestInformation"
sql = sql & " Where (EffectID = 0) and (InvID=" & InvID & ")"
sql = sql & " GROUP BY InvID"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
GetDividArea = IIf(IsNull(Rs5("SumArea").value), 0, Rs5("SumArea").value)
Else
GetDividArea = 0
End If
End Function
Function LandArea(Optional Land As Double) As Double
Dim Rs5 As ADODB.Recordset
Dim sql As String
Set Rs5 = New ADODB.Recordset
sql = "select Area from TblBuyLanReEst where id=" & Land & ""
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
LandArea = IIf(IsNull(Rs5("Area").value), 0, Rs5("Area").value)
Else
LandArea = 0
End If
End Function
Function LandValue(Optional Land As Double) As Double
Dim Rs5 As ADODB.Recordset
Dim sql As String
Set Rs5 = New ADODB.Recordset
sql = "select Total from TblBuyLanReEst where id=" & Land & ""
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
LandValue = IIf(IsNull(Rs5("Total").value), 0, Rs5("Total").value)
Else
LandValue = 0
End If
End Function


Function GetCostMeter(Optional invsID As Double = 0) As Double
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Dim sql As String
sql = "select CostMeterExp from Tblinvestment where id =" & invsID & ""
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
GetCostMeter = IIf(IsNull(Rs6("CostMeterExp").value), 0, Rs6("CostMeterExp").value)
Else
GetCostMeter = 0
End If
End Function

Function GetTotalLand(Optional invsID As Double = 0) As Double
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     SUM(TotalValue) AS TotalValue, InviseOrder"
sql = sql & " From dbo.TblActivateInvestment"
sql = sql & " Where (InviseOrder = " & invsID & ")"
sql = sql & " GROUP BY InviseOrder"
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
GetTotalLand = IIf(IsNull(Rs6("TotalValue").value), 0, Rs6("TotalValue").value)
Else
GetTotalLand = 0
End If
End Function
Function GetToataExpensive(Optional invsID As Double) As Double
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
sql = "SELECT     SUM(DevlopValue) AS DevlopValue"
sql = sql & " From dbo.TblExpensesInvesment"
sql = sql & " Where (InvesID = " & invsID & ")"
sql = sql & " GROUP BY InvesID"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetToataExpensive = IIf(IsNull(Rs4("DevlopValue").value), 0, Rs4("DevlopValue").value)
Else
GetToataExpensive = 0
End If
End Function

Private Sub GridInstallments_AfterEdit(ByVal row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim X As Integer
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim InvesTotal As Double
    Dim Rs7 As ADODB.Recordset
    Set Rs7 = New ADODB.Recordset
    Dim MeterValue As Double
    With GridInstallments
     Select Case .ColKey(Col)
     Case "unit"
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("unitId"), False, True)
                 .TextMatrix(row, .ColIndex("unitId")) = StrAccountCode
                 
               
              FrmAddItemItemInvest.DivIDDet = val(StrAccountCode)
               If Rd(1).value = True Then
             FrmAddItemItemInvest.SalID = val(Me.TxtBillNo.text)
             Else
             FrmAddItemItemInvest.SalID = val(TxtSerial1.text)
             End If
              FrmAddItemItemInvest.InvID = val(.TextMatrix(row, .ColIndex("InvesID")))
             FrmAddItemItemInvest.Row1 = row
             If RdType(1).value = True Then
             FrmAddItemItemInvest.TypDiv = 1
             Else
             FrmAddItemItemInvest.TypDiv = 0
             End If
             
              Load FrmAddItemItemInvest
            
              FrmAddItemItemInvest.show
              
        Case "InvesName"
                 StrAccountCode = .ComboData
                 MeterValue = 0
                 
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("InvesID"), False, True)
                 .TextMatrix(row, .ColIndex("InvesID")) = StrAccountCode
                 If RdType(0).value = True Then
                MeterValue = GetCostMeter(val(.TextMatrix(row, .ColIndex("InvesID"))))
                If MeterValue = 0 Then
                 MeterValue = GetTotalLand(val(.TextMatrix(row, .ColIndex("InvesID")))) + GetToataExpensive(val(.TextMatrix(row, .ColIndex("InvesID"))))
                 End If
                  .TextMatrix(row, .ColIndex("TotalArea")) = GetTOtalArea(val(.TextMatrix(row, .ColIndex("InvesID")))) - GetDividArea(val(.TextMatrix(row, .ColIndex("InvesID"))))
                  Else
                  .TextMatrix(row, .ColIndex("TotalArea")) = LandArea(val(.TextMatrix(row, .ColIndex("InvesID")))) - GetDividArea(val(.TextMatrix(row, .ColIndex("InvesID"))))
                   MeterValue = LandValue(val(.TextMatrix(row, .ColIndex("InvesID"))))
                   End If
                   If val(.TextMatrix(row, .ColIndex("TotalArea"))) <> 0 Then
                   MeterValue = MeterValue / val(.TextMatrix(row, .ColIndex("TotalArea")))
                   MeterValue = Round(MeterValue, 2)
                  End If
                .TextMatrix(row, .ColIndex("MeterValue")) = Round(MeterValue, 2)
                .TextMatrix(row, .ColIndex("TotalCost")) = val(.TextMatrix(row, .ColIndex("MeterValue"))) * val(.TextMatrix(row, .ColIndex("Area")))
                
             
                If val(.TextMatrix(row, .ColIndex("TypeDis"))) = 2 Then
               .TextMatrix(row, .ColIndex("Net")) = val(.TextMatrix(row, .ColIndex("Total"))) + val(.TextMatrix(row, .ColIndex("DisValue")))
                ElseIf val(.TextMatrix(row, .ColIndex("TypeDis"))) = 3 Then
               .TextMatrix(row, .ColIndex("Net")) = val(.TextMatrix(row, .ColIndex("Total"))) - (val(.TextMatrix(row, .ColIndex("Total"))) * val(.TextMatrix(row, .ColIndex("DisValue"))) / 100)
               .TextMatrix(row, .ColIndex("Net")) = Round(val(.TextMatrix(row, .ColIndex("Net"))), 2)
              Else
              .TextMatrix(row, .ColIndex("Net")) = .TextMatrix(row, .ColIndex("Total"))
              End If
              .TextMatrix(row, .ColIndex("Profit")) = val(.TextMatrix(row, .ColIndex("Net"))) - val(.TextMatrix(row, .ColIndex("TotalCost")))
       '     Case "Name"
            
       '          StrAccountCode = .ComboData
       '          LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("PartID"), False, True)
       '          .TextMatrix(Row, .ColIndex("PartID")) = StrAccountCode
       '          .TextMatrix(Row, .ColIndex("BlockID")) = 0
       '
       '
       '          .TextMatrix(Row, .ColIndex("BlockName")) = ""
       '          If Me.Rd(1).value = True Then
       '       StrSQL = "Select *  from TblDivInvestInformation "
       '       StrSQL = StrSQL & " where (SalID=" & val(TxtBillNo.text) & ")"
       '
       '      Else
       '      StrSQL = "Select *  from TblDivInvestInformation where (DivIDDet=" & val(StrAccountCode) & " and (SalesBlocPayed IS NULL))"
       '      If Me.TxtModFlg.text = "E" Then
       '      StrSQL = StrSQL & " or (SalID=" & val(TxtSerial1.text) & " and DivIDDet=" & val(StrAccountCode) & " )"
       '      End If
       '      End If
       '       Rs7.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
       '       If Rs7.RecordCount > 0 Then
       '       If Rs7.RecordCount > 1 Then
       '        FrmAddItemItemInvest.DivIDDet = val(StrAccountCode)
       '        If Rd(1).value = True Then
       '      FrmAddItemItemInvest.SalID = val(Me.TxtBillNo.text)
       '      Else
       '      FrmAddItemItemInvest.SalID = val(TxtSerial1.text)
       '      End If
       '      FrmAddItemItemInvest.Row1 = Row
       '
       '       Load FrmAddItemItemInvest
       '
       '       FrmAddItemItemInvest.Show
       '       Else
       '       .TextMatrix(Row, .ColIndex("BlockName")) = IIf(IsNull(Rs7("BlokNo").value), "", Rs7("BlokNo").value)
       '       .TextMatrix(Row, .ColIndex("BlockID")) = IIf(IsNull(Rs7("ID").value), 0, Rs7("ID").value)
       '       .TextMatrix(Row, .ColIndex("Area")) = IIf(IsNull(Rs7("Area").value), "", Rs7("Area").value)
       '       End If
       '       Else
       '       .TextMatrix(Row, .ColIndex("Area")) = 0
       '       End If
       '
       '
       '       .TextMatrix(Row, .ColIndex("TotalCost")) = val(.TextMatrix(Row, .ColIndex("MeterValue"))) * val(.TextMatrix(Row, .ColIndex("Area")))
       '         If val(.TextMatrix(Row, .ColIndex("TypeDis"))) = 2 Then
       '       .TextMatrix(Row, .ColIndex("Net")) = val(.TextMatrix(Row, .ColIndex("Total"))) + val(.TextMatrix(Row, .ColIndex("DisValue")))
       '       ElseIf val(.TextMatrix(Row, .ColIndex("TypeDis"))) = 3 Then
       '       .TextMatrix(Row, .ColIndex("Net")) = val(.TextMatrix(Row, .ColIndex("Total"))) - (val(.TextMatrix(Row, .ColIndex("Total"))) * val(.TextMatrix(Row, .ColIndex("DisValue"))) / 100)
       '       .TextMatrix(Row, .ColIndex("Net")) = Round(val(.TextMatrix(Row, .ColIndex("Net"))), 2)
       '       Else
       '       .TextMatrix(Row, .ColIndex("Net")) = .TextMatrix(Row, .ColIndex("Total"))
       '       End If
     ' Case "BlockName"
     '
     '            StrAccountCode = .ComboData
     '            LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("BlockID"), False, True)
     '            .TextMatrix(Row, .ColIndex("BlockID")) = StrAccountCode
     '            If val(.TextMatrix(Row, .ColIndex("BlockID"))) <> 0 Then
     '         strSQL = "Select Area from TblDivInvestInformation where ID=" & val(StrAccountCode) & ""
     '         Rs7.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
     '         If Rs7.RecordCount > 0 Then
     '         .TextMatrix(Row, .ColIndex("Area")) = IIf(IsNull(Rs7("Area").value), "", Rs7("Area").value)
     '         Else
     '         .TextMatrix(Row, .ColIndex("Area")) = 0
     '         End If
     '         End If
     '         .TextMatrix(Row, .ColIndex("TotalCost")) = val(.TextMatrix(Row, .ColIndex("MeterValue"))) * val(.TextMatrix(Row, .ColIndex("Area")))
     '         .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("Area")) * val(.TextMatrix(Row, .ColIndex("Valu"))))
     '           If val(.TextMatrix(Row, .ColIndex("TypeDis"))) = 2 Then
     '         .TextMatrix(Row, .ColIndex("Net")) = val(.TextMatrix(Row, .ColIndex("Total"))) + val(.TextMatrix(Row, .ColIndex("DisValue")))
     '         ElseIf val(.TextMatrix(Row, .ColIndex("TypeDis"))) = 3 Then
     '         .TextMatrix(Row, .ColIndex("Net")) = val(.TextMatrix(Row, .ColIndex("Total"))) - (val(.TextMatrix(Row, .ColIndex("Total"))) * val(.TextMatrix(Row, .ColIndex("DisValue"))) / 100)
     '         .TextMatrix(Row, .ColIndex("Net")) = Round(val(.TextMatrix(Row, .ColIndex("Net"))), 2)
     '         Else
     '         .TextMatrix(Row, .ColIndex("Net")) = .TextMatrix(Row, .ColIndex("Total"))
     '         End If
              
              .TextMatrix(row, .ColIndex("Profit")) = val(.TextMatrix(row, .ColIndex("Net"))) - val(.TextMatrix(row, .ColIndex("TotalCost")))
        Case "Valu", "Area"
        If val(.TextMatrix(row, .ColIndex("Valu"))) < val(.TextMatrix(row, .ColIndex("MeterValue"))) Then
            If SystemOptions.UserInterface = EnglishInterface Then
               X = MsgBox("Selling price less than the cost", vbCritical + vbYesNo)
             Else
                X = MsgBox(" ”⁄— «·»Ì⁄ «ﬁ· „‰ «· ﬂ·›… Â· «‰  „Ê«›ﬁ", vbCritical + vbYesNo)
           End If
          If X = vbNo Then
          .TextMatrix(row, .ColIndex("Valu")) = 0
          Exit Sub
          End If
    End If
              .TextMatrix(row, .ColIndex("Total")) = val(.TextMatrix(row, .ColIndex("Area")) * val(.TextMatrix(row, .ColIndex("Valu"))))
               If val(.TextMatrix(row, .ColIndex("TypeDis"))) = 2 Then
              .TextMatrix(row, .ColIndex("Net")) = val(.TextMatrix(row, .ColIndex("Total"))) + val(.TextMatrix(row, .ColIndex("DisValue")))
              ElseIf val(.TextMatrix(row, .ColIndex("TypeDis"))) = 3 Then
              .TextMatrix(row, .ColIndex("Net")) = val(.TextMatrix(row, .ColIndex("Total"))) - (val(.TextMatrix(row, .ColIndex("Total"))) * val(.TextMatrix(row, .ColIndex("DisValue"))) / 100)
              .TextMatrix(row, .ColIndex("Net")) = Round(val(.TextMatrix(row, .ColIndex("Net"))), 2)
              Else
              .TextMatrix(row, .ColIndex("Net")) = .TextMatrix(row, .ColIndex("Total"))
              End If
              .TextMatrix(row, .ColIndex("Profit")) = val(.TextMatrix(row, .ColIndex("Net"))) - val(.TextMatrix(row, .ColIndex("TotalCost")))
              
        Case "TypeDis"
              .TextMatrix(row, .ColIndex("DisValue")) = ""
        Case "DisValue"
              If val(.TextMatrix(row, .ColIndex("TypeDis"))) = 2 Then
              .TextMatrix(row, .ColIndex("Net")) = val(.TextMatrix(row, .ColIndex("Total"))) - val(.TextMatrix(row, .ColIndex("DisValue")))
              ElseIf val(.TextMatrix(row, .ColIndex("TypeDis"))) = 3 Then
              .TextMatrix(row, .ColIndex("Net")) = val(.TextMatrix(row, .ColIndex("Total"))) - (val(.TextMatrix(row, .ColIndex("Total"))) * val(.TextMatrix(row, .ColIndex("DisValue"))) / 100)
              .TextMatrix(row, .ColIndex("Net")) = Round(val(.TextMatrix(row, .ColIndex("Net"))), 2)
              Else
              .TextMatrix(row, .ColIndex("Net")) = .TextMatrix(row, .ColIndex("Total"))
              End If
              .TextMatrix(row, .ColIndex("Profit")) = val(.TextMatrix(row, .ColIndex("Net"))) - val(.TextMatrix(row, .ColIndex("TotalCost")))
           End Select
   
        If row = .rows - 1 Then
    
          .rows = .rows + 1
        End If
    End With
RelinGrid1
End Sub

Private Sub GridInstallments_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)

 With GridInstallments
 If Me.TxtModFlg.text = "R" Then
Cancel = True
Else
        Select Case .ColKey(Col)
           Case "CodeUnit"
                 Cancel = True
                    Case "PartName"
                 Cancel = True
        
        Case "BlockName"
                Cancel = True
                
            Case "Area"
               ' Cancel = True
               
                 Case "Valu"
              If val(.TextMatrix(row, .ColIndex("BlockID"))) = 0 Then
                              Cancel = True
                         If SystemOptions.UserInterface = ArabicInterface Then
                         MsgBox "Ì—ÃÏ «Œ Ì«— —ﬁ„ «· Õ·Ì·Ì «Ê·«"
                         Else
                         MsgBox "Please Select  No."
                         End If
                   Exit Sub
              Else
              
              .ComboList = ""
              End If
                 Case "Total"
                Cancel = True
                
                 Case "MeterValue"
                 Cancel = True
                 Case "Net"
                 Cancel = True
                 
                 Case "Remarks"
               .ComboList = ""
               Case "DisValue"
              If val(.TextMatrix(row, .ColIndex("TypeDis"))) = 1 Or val(.TextMatrix(row, .ColIndex("TypeDis"))) = 0 Then
                              Cancel = True
              Else
              
              .ComboList = ""
              End If
        End Select
       End If
    End With
End Sub

Private Sub GridInstallments_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim Rs7 As ADODB.Recordset
    With GridInstallments

        Select Case .ColKey(Col)
        Case "unit"
        If val(.TextMatrix(row, .ColIndex("InvesID"))) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„”«Â„Â «Ê «·«—÷ «Ê·«"
Else
MsgBox "Please Select Contributory or Land"
End If
Exit Sub
End If
If Me.RdType(0).value = True Then
  StrSQL = " SELECT     *"
  StrSQL = StrSQL & " From dbo.TblSpreading"
  StrSQL = StrSQL & " WHERE     (ID IN"
  StrSQL = StrSQL & "                        (SELECT     DivMainID"
  StrSQL = StrSQL & "                           From TblDivInvesment"
  StrSQL = StrSQL & "                           WHERE     InvesID = " & val(.TextMatrix(row, .ColIndex("InvesID"))) & ")) OR"
  StrSQL = StrSQL & "                    (ID IN"
  StrSQL = StrSQL & "                        (SELECT     dbo.TblDivInvesmentDet.TypeDivi"
  StrSQL = StrSQL & "                           FROM         dbo.TblDivInvesmentDet RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                                                 dbo.TblDivInvesment ON dbo.TblDivInvesmentDet.DivInvID = dbo.TblDivInvesment.ID"
  StrSQL = StrSQL & "                           WHERE     (dbo.TblDivInvesment.InvesID = " & val(.TextMatrix(row, .ColIndex("InvesID"))) & ")))"
 Else
   StrSQL = " SELECT     *"
  StrSQL = StrSQL & " From dbo.TblSpreading"
  StrSQL = StrSQL & " WHERE     (ID IN"
  StrSQL = StrSQL & "                        (SELECT     DivMainID"
  StrSQL = StrSQL & "                           From TblDivInvesment"
  StrSQL = StrSQL & "                           WHERE     LandID = " & val(.TextMatrix(row, .ColIndex("InvesID"))) & ")) OR"
  StrSQL = StrSQL & "                    (ID IN"
  StrSQL = StrSQL & "                        (SELECT     dbo.TblDivInvesmentDet.TypeDivi"
  StrSQL = StrSQL & "                           FROM         dbo.TblDivInvesmentDet RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                                                 dbo.TblDivInvesment ON dbo.TblDivInvesmentDet.DivInvID = dbo.TblDivInvesment.ID"
  StrSQL = StrSQL & "                           WHERE     (dbo.TblDivInvesment.LandID = " & val(.TextMatrix(row, .ColIndex("InvesID"))) & ")))"
 End If
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = GridInstallments.BuildComboList(rs, "Name", "ID")
                Else
                    StrComboList = GridInstallments.BuildComboList(rs, "NameE", "ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ComboList = StrComboList

                .ComboList = StrComboList
                             
Case "Name"
If val(.TextMatrix(row, .ColIndex("InvesID"))) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„”«Â„Â «Ê «·«—÷ «Ê·«"
Else
MsgBox "Please Select Contributory or Land"
End If
Exit Sub
End If
             .TextMatrix(row, .ColIndex("Name")) = ""
              StrSQL = "SELECT     DivIDDet, PartNo"
              StrSQL = StrSQL & " From dbo.TblDivInvestInformation"
              If Me.TxtModFlg.text = "N" Then
              StrSQL = StrSQL & "  Where (Not (PartNo Is Null))AND (SalesBlocPayed IS NULL)  And (InvID = " & val(.TextMatrix(row, .ColIndex("InvesID"))) & ")and  (EffectID =1)"
              End If
                 If Me.TxtModFlg.text = "E" Then
                 StrSQL = StrSQL & "  Where ((Not (PartNo Is Null))AND (SalesBlocPayed IS NULL)"
             StrSQL = StrSQL & " or SalID=" & val(TxtSerial1.text) & ") "
             StrSQL = StrSQL & " And (InvID = " & val(.TextMatrix(row, .ColIndex("InvesID"))) & ")and  (EffectID =1)"
             End If
             
              StrSQL = StrSQL & " GROUP BY DivIDDet, PartNo"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = GridInstallments.BuildComboList(rs, "PartNo", "DivIDDet")
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ComboList = StrComboList
Case "BlockName"

 If val(.TextMatrix(row, .ColIndex("InvesID"))) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„”«Â„Â «Ê·«"
Else
MsgBox "Please Select Contributory "
End If
Exit Sub
End If
 If val(.TextMatrix(row, .ColIndex("PartID"))) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— —ﬁ„ «·»·Êﬂ «Ê·«"
Else
MsgBox "Please Select Part No "
End If
Exit Sub
End If
.TextMatrix(row, .ColIndex("BlockName")) = ""
 StrSQL = "SELECT     ID ,BlokNo"
StrSQL = StrSQL & " From dbo.TblDivInvestInformation"
StrSQL = StrSQL & "  Where (Not (BlokNo Is Null))AND (SalesBlocPayed IS NULL)and (DivIDDet=" & val(.TextMatrix(row, .ColIndex("PartID"))) & ")  "
       If Me.TxtModFlg.text = "E" Then
             StrSQL = StrSQL & " or SalID=" & val(TxtSerial1.text) & " "
             End If
StrSQL = StrSQL & " GROUP BY BlokNo, ID"

                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList = GridInstallments.BuildComboList(rs, "BlokNo", "ID")

'                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
          '      End If
                .ComboList = StrComboList
                
  Case "InvesName"
          '.TextMatrix(Row, .ColIndex("DivID")) = ""
          
          If Me.RdType(0).value = True Then
           If SystemOptions.UserInterface = ArabicInterface Then
             StrSQL = "SElect id ,name from Tblinvestment where id  in (select InvesID from TblInvesOpenSales )  "
             Else
            StrSQL = "SElect id ,NameE from Tblinvestment where id  in (select InvesID from TblInvesOpenSales )   "
            End If
           Else
                  If SystemOptions.UserInterface = ArabicInterface Then
             StrSQL = "SElect id ,name from TblBuyLanReEst where id  in (select LandID from TblDivInvesment )  "
             Else
            StrSQL = "SElect id ,NameE from TblBuyLanReEst where id  in (select LandID from TblDivInvesment )   "
            End If
            
End If
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = GridInstallments.BuildComboList(rs, "Name", "ID")
                Else
                    StrComboList = GridInstallments.BuildComboList(rs, "NameE", "ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ComboList = StrComboList

                .ComboList = StrComboList
                
        End Select
    End With
End Sub


Private Sub ISButton2_Click()
 If opt(0).value = False And opt(1).value = False And opt(2).value = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
        Else
        MsgBox "Please Select Method Number of decimal"
        End If
        Exit Sub
        End If
If val(lbl(13).Caption) = 0 Then
Exit Sub
End If
If val(TxtPaymentNo.text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· ⁄œœ «·œ›⁄«  "
Else
MsgBox "Please Enter No of  Payments"
End If
TxtPaymentNo.SetFocus
Exit Sub
End If
If val(txtPeriod.text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· «·› —… »Ì‰ «·œ›⁄«  "
Else
MsgBox "Please Enter No of  Period"
End If
txtPeriod.SetFocus
Exit Sub
End If
If val(DcbPeriodsID.ListIndex) = -1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«—    ‰Ê⁄ «·› —… "
Else
MsgBox "Please Enter Type of  Period"
End If
DcbPeriodsID.SetFocus
Exit Sub
End If


filgrid1
RelinGrid
End Sub


Private Sub ISButton3_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments TxtSerial1.text, "1303201601"
ErrTrap:
End Sub

Private Sub ISButton4_Click()
If Me.TxtModFlg.text <> "R" Then
Dim X As Integer
Dim sql As String
Dim i As Integer
With GridInstallments

If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Will be deleted all linked operations", vbCritical + vbYesNo)
    Else
        X = MsgBox("”Ê› Ì „ Õ–› ﬂ· «·⁄„·Ì«  «·„— »ÿ… »Â–«   ", vbCritical + vbYesNo)
End If
If X = vbNo Then Exit Sub
For i = 1 To .rows - 1
      
       If val((.TextMatrix(i, .ColIndex("BlockID")))) = 0 Then

         sql = "Update TblDivInvestInformation  set SalesBlocPayed=Null where ID= " & val(.TextMatrix(i, .ColIndex("BlockID"))) & ""
       Cn.Execute sql
         StrSQL = "Update TblSaleBilllInvestmentDet  set ReturnSal=Null where SBINVID=" & val(TxtBillNo.text) & " and BlockID= " & val(.TextMatrix(i, .ColIndex("BlockID"))) & ""
       Cn.Execute StrSQL

      End If

Next i
GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 2
            BtnUndo.Enabled = False
End With
End If
End Sub

Private Sub ISButton5_Click()
print_report
End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double
    Dim i As Integer
    Dim j As Integer
    
    
    CommissionAccount = get_account_code_branch(131, my_branch)
If val(TxtNetComm) > 0 Then


     If CommissionAccount = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                 MsgBox "Not creation  branch", vbCritical
                
                End If
               Exit Sub
            Else

                If CommissionAccount = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» ⁄„Ê·«  „»Ì⁄«    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                    MsgBox "Is not specified sales account", vbCritical
                    End If
       Exit Sub
                End If
            End If
       
       
End If


    '---------------------- check if data Vaclete -----------------------
      If dcBranch.text = "" And val(dcBranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            dcBranch.SetFocus
            Exit Sub
     End If
           If DcbTyp.text = "" And val(DcbTyp.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·⁄„Ê·…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Commission Type ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            DcbTyp.SetFocus
            Exit Sub
     End If
     If Rd(1).value = True Then
       If val(TxtBillNo.text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· —ﬁ„ «·›« Ê—…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select enter no of bill ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            TxtBillNo.SetFocus
            Exit Sub
     End If
     End If
     If (val(DcbLand.BoundText) = 0 Or DcbLand.text = "") And RdType(1).value = True Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·«—÷  "
     Else
     MsgBox "Please Select Land"
     End If
     DcbLand.SetFocus
     Exit Sub
     End If

     
          If val(DcbCus.BoundText) = 0 Or DcbCus.text = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·⁄„Ì·  "
     Else
     MsgBox "Please Select Customer"
     End If
     DcbCus.SetFocus
     Exit Sub
     End If
         If val(DcbSales.BoundText) = 0 Or DcbSales.text = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·»«∆⁄  "
     Else
     MsgBox "Please Select Seller"
     End If
     DcbSales.SetFocus
     Exit Sub
     End If
          
'
'          If val(Txtcommission.Text) = 0 Then
'     If SystemOptions.UserInterface = ArabicInterface Then
'     MsgBox "⁄›Ê«...«·—Ã«¡ «œŒ«· «·⁄„Ê·…  "
'     Else
'     MsgBox "Please Enter Commission"
'     End If
'     Txtcommission.SetFocus
'     Exit Sub
'     End If
     With VSFlexGrid1
      For i = .FixedRows To .rows - 1
         If opt(0).value = True And i = 1 Then
            .TextMatrix(i, .ColIndex("PaymentValue")) = val(.TextMatrix(i, .ColIndex("PaymentValue"))) + ((val(lbl(13).Caption) - val(lbl(9).Caption)))
            .TextMatrix(i, .ColIndex("PaymentValue")) = Round(val(.TextMatrix(i, .ColIndex("PaymentValue"))), 2)
            End If
             If opt(1).value = True And i = (.rows - 1) Then
            
            .TextMatrix(i, .ColIndex("PaymentValue")) = val(.TextMatrix(i, .ColIndex("PaymentValue"))) + ((val(lbl(13).Caption) - val(lbl(9).Caption)))
           .TextMatrix(i, .ColIndex("PaymentValue")) = Round(val(.TextMatrix(i, .ColIndex("PaymentValue"))), 2)
            End If
            
        Next i
      End With
      RelinGrid
      
  With Me.VSFlexGrid1
If .rows > 1 Then
If Round(val(lbl(9).Caption), 2) <> Round(val(lbl(13).Caption), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ·«‰  „Ã„Ê⁄ ﬁÌ„ «·œ›⁄«  ·«Ì”«ÊÌ  «·ﬁÌ„… «·«Ã„«·Ì…"
Else
MsgBox "Can not Save  The values of Payement not equal the Total Value"
End If
Exit Sub
End If
End If
End With
           With Me.GridInstallments

           '''''''
             For i = .FixedRows To .rows - 1
       If val(.TextMatrix(i, .ColIndex("InvesID"))) <> 0 Then
       If val(.TextMatrix(i, .ColIndex("BlockID"))) <> 0 And (.TextMatrix(i, .ColIndex("BlockName"))) <> "" Then
       If val(.TextMatrix(i, .ColIndex("Valu"))) = 0 Then
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "Ì—ÃÏ «œŒ«· «·”⁄— ›Ì «·”ÿ— —ﬁ„" & i
       Else
       MsgBox "Please Enter Price In line" & i
       End If
       Exit Sub
       End If
       Else
        If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "Ì—ÃÏ  ≈Œ Ì«— «·—ﬁ„ «· Õ·Ì·Ì «·”ÿ— —ﬁ„" & i
       Else
       MsgBox "Please Select  Part No In line" & i
       End If
       Exit Sub
      End If
      End If
     Next i
     

     
     
    End With
         If Round(val(TxtNetComm.text), 2) > Round(val(lbl(14).Caption), 2) Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...  ·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·⁄„Ê·… «ﬂ»— „‰ «·—»Õ  "
     Else
     MsgBox "Can not commission greater than the profit "
     End If
     DcbLand.SetFocus
     Exit Sub
     End If
    '------------------------------ check if Empcode exist ----------------------
'   StrVacName = IsRecExist("TblEmploymentModel", "name", Trim(TxtVacName.text), "name", "Vac_ID<>'" & Trim(TxtSerial1.text) & "'")
  ' If StrVacName <> "" Then
 '    Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·«”„ „‰ ﬁ»·"
  '     MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
  '    TxtVacName.SetFocus
 '     Exit Sub
'   End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text
            '------------------------------ new record ----------------------------
        Case "N"
                  '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.Title
    End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblSaleBilllInvestment", "ID", "")
    Me.TxtSerial1.text = StrRecID
    
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Sub GetInformationLand(Optional ID As Double = 0)
Dim Rs8 As ADODB.Recordset
Dim sql As String
Set Rs8 = New ADODB.Recordset
If ID <> 0 Then
sql = "Select * from TblBuyLanReEst where id=" & ID & " "
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
    Me.TxtPriceHadW.text = IIf(IsNull(Rs8("NorthlengthStr").value), "", Rs8("NorthlengthStr").value)
    Me.TxtPriceSomW.text = IIf(IsNull(Rs8("SouthlengthStr").value), "", Rs8("EastlengthStr").value)
    Me.TxteastWriiten.text = IIf(IsNull(Rs8("EastlengthStr").value), "", Rs8("EastlengthStr").value)
    Me.TxtwestWriiten.text = IIf(IsNull(Rs8("WestlengthStr").value), "", Rs8("WestlengthStr").value)
    Me.TxtNorthLength.text = IIf(IsNull(Rs8("Northlength").value), 0, Rs8("Northlength").value)
    Me.TxtSouthLength.text = IIf(IsNull(Rs8("Southlength").value), 0, Rs8("Southlength").value)
    Me.TxtEastLength.text = IIf(IsNull(Rs8("Eastlength").value), 0, Rs8("Eastlength").value)
    Me.txtWestlength.text = IIf(IsNull(Rs8("Westlength").value), 0, Rs8("Westlength").value)
    Me.TxtDesLocation.text = IIf(IsNull(Rs8("DesLocation").value), "", Rs8("DesLocation").value)
    TxtPropertyDeed.text = IIf(IsNull(Rs8("TitledeedNo").value), "", Rs8("TitledeedNo").value)
Else
Me.TxtPriceHadW.text = ""
Me.TxtPriceSomW.text = ""
Me.TxteastWriiten.text = ""
Me.TxtwestWriiten.text = ""
Me.TxtNorthLength.text = ""
Me.TxtSouthLength.text = ""
Me.TxtEastLength.text = ""
Me.txtWestlength.text = ""
Me.TxtDesLocation.text = ""
Me.TxtPropertyDeed.text = ""
End If
End If
End Sub
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 1
      VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 1
            'Comm
 sql = "SELECT dbo.TblSaleBilllInvestmentDet.CommP ,  dbo.TblSaleBilllInvestmentDet.Comm  ,   dbo.TblSaleBilllInvestmentDet.Valu, dbo.TblSaleBilllInvestmentDet.Remarks, dbo.TblSaleBilllInvestmentDet.FristDate, dbo.TblSaleBilllInvestmentDet.PartID, "
 sql = sql & "                     dbo.TblSaleBilllInvestmentDet.Area, dbo.TblSaleBilllInvestmentDet.Total, dbo.TblSaleBilllInvestmentDet.InvesID, dbo.TblSaleBilllInvestmentDet.Net,"
 sql = sql & "                     dbo.TblSaleBilllInvestmentDet.TypeDis, dbo.TblSaleBilllInvestmentDet.DisValue, dbo.TblSaleBilllInvestmentDet.DivID, dbo.TblSaleBilllInvestmentDet.TypeTrns,"
 sql = sql & "                     dbo.TblSaleBilllInvestmentDet.SBINVID, dbo.Tblinvestment.Name, dbo.Tblinvestment.NameE, dbo.TblSaleBilllInvestmentDet.MeterValue,"
 sql = sql & "                     dbo.TblSaleBilllInvestmentDet.TotalCost, dbo.TblSaleBilllInvestmentDet.Profit, TblDivInvestInformation_1.ID, dbo.TblSaleBilllInvestmentDet.BlockID,"
 sql = sql & "                     TblDivInvestInformation_1.DivIDDet2, TblDivInvestInformation_1.BlokNo, TblDivInvestInformation_1.PartNo,"
 sql = sql & "                     dbo.GetInvestMentPartName(dbo.TblSaleBilllInvestmentDet.PartID) AS asPartName, dbo.TblSaleBilllInvestmentDet.TotalArea,"
 sql = sql & "                     dbo.TblSaleBilllInvestmentDet.ID AS IDDet, TblDivInvestInformation_1.CodeUnit, dbo.TblSaleBilllInvestmentDet.ReturnSal, dbo.TblSaleBilllInvestmentDet.ReturnID,"
 sql = sql & "                     dbo.TblSaleBilllInvestmentDet.unitId, TblSpreading_2.Name AS UntName, TblSpreading_2.NameE AS UntNameE, dbo.TblSaleBilllInvestmentDet.unitunidpart,"
 sql = sql & "                     TblSpreading_1.Name AS UntNamepart, TblSpreading_1.NameE AS UntNamepartE, dbo.TblSaleBilllInvestmentDet.CodeUnit AS CodeUnitDet,"
 sql = sql & "                     dbo.TblSaleBilllInvestmentDet.PartName, TblSpreading_2.Name AS SperName, TblSpreading_2.NameE AS SperNameE, dbo.TblBuyLanReEst.Name AS LandName,"
 sql = sql & "                     dbo.TblBuyLanReEst.NameE AS LandNameE, dbo.TblSaleBilllInvestmentDet.TypDiv"
 sql = sql & "  FROM         dbo.TblBuyLanReEst RIGHT OUTER JOIN"
 sql = sql & "                     dbo.TblSaleBilllInvestmentDet ON dbo.TblBuyLanReEst.ID = dbo.TblSaleBilllInvestmentDet.InvesID LEFT OUTER JOIN"
 sql = sql & "                     dbo.TblSpreading TblSpreading_1 ON dbo.TblSaleBilllInvestmentDet.unitunidpart = TblSpreading_1.ID LEFT OUTER JOIN"
 sql = sql & "                     dbo.TblSpreading TblSpreading_2 ON dbo.TblSaleBilllInvestmentDet.unitId = TblSpreading_2.ID LEFT OUTER JOIN"
 sql = sql & "                     dbo.TblDivInvestInformation TblDivInvestInformation_1 ON dbo.TblSaleBilllInvestmentDet.BlockID = TblDivInvestInformation_1.ID LEFT OUTER JOIN"
 sql = sql & "                     dbo.Tblinvestment ON dbo.TblSaleBilllInvestmentDet.InvesID = dbo.Tblinvestment.ID"
sql = sql & " Where (dbo.TblSaleBilllInvestmentDet.SBINVID = " & val(TxtSerial1.text) & ") And (dbo.TblSaleBilllInvestmentDet.TypeTrns = 0)"
If Me.TxtModFlg.text <> "R" Then
If Rd(1).value = True Then
sql = sql & " and (dbo.TblSaleBilllInvestmentDet.ReturnSal is null )"
End If
End If

  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs1("IDDet").value), 0, Rs1("IDDet").value)
                   .TextMatrix(i, .ColIndex("CommP")) = IIf(IsNull(Rs1("CommP").value), 0, Rs1("CommP").value)
                   .TextMatrix(i, .ColIndex("Comm")) = IIf(IsNull(Rs1("Comm").value), 0, Rs1("Comm").value)
                     '
                   '.TextMatrix(i, .ColIndex("PartID")) = IIf(IsNull(Rs1("PartID").value), 0, Rs1("PartID").value)
                   .TextMatrix(i, .ColIndex("InvesID")) = IIf(IsNull(Rs1("InvesID").value), 0, Rs1("InvesID").value)
                  ' .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("asPartName").value), "", Rs1("asPartName").value)
                   .TextMatrix(i, .ColIndex("TotalArea")) = IIf(IsNull(Rs1("TotalArea").value), "", Rs1("TotalArea").value)
                   .TextMatrix(i, .ColIndex("Area")) = IIf(IsNull(Rs1("Area").value), "", Rs1("Area").value)
                   .TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(Rs1("Valu").value), "", Rs1("Valu").value)
                   .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(Rs1("Total").value), "", Rs1("Total").value)
                   .TextMatrix(i, .ColIndex("TypeDis")) = IIf(IsNull(Rs1("TypeDis").value), 1, Rs1("TypeDis").value)
                   .TextMatrix(i, .ColIndex("DisValue")) = IIf(IsNull(Rs1("DisValue").value), "", Rs1("DisValue").value)
                   .TextMatrix(i, .ColIndex("Net")) = IIf(IsNull(Rs1("Net").value), "", Rs1("Net").value)
                   .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   .TextMatrix(i, .ColIndex("TotalCost")) = IIf(IsNull(Rs1("TotalCost").value), 0, Rs1("TotalCost").value)
                   .TextMatrix(i, .ColIndex("MeterValue")) = IIf(IsNull(Rs1("MeterValue").value), 0, Rs1("MeterValue").value)
                   .TextMatrix(i, .ColIndex("Profit")) = IIf(IsNull(Rs1("Profit").value), 0, Rs1("Profit").value)
                   .TextMatrix(i, .ColIndex("BlockID")) = IIf(IsNull(Rs1("BlockID").value), 0, Rs1("BlockID").value)
                   .TextMatrix(i, .ColIndex("BlockName")) = IIf(IsNull(Rs1("BlokNo").value), "", Rs1("BlokNo").value)
                   .TextMatrix(i, .ColIndex("unitId")) = IIf(IsNull(Rs1("unitId").value), 0, Rs1("unitId").value)
                   .TextMatrix(i, .ColIndex("unitunidpart")) = IIf(IsNull(Rs1("unitunidpart").value), 0, Rs1("unitunidpart").value)
                   .TextMatrix(i, .ColIndex("CodeUnit")) = IIf(IsNull(Rs1("CodeUnitDet").value), "", Rs1("CodeUnitDet").value)
                   .TextMatrix(i, .ColIndex("PartName")) = IIf(IsNull(Rs1("PartName").value), "", Rs1("PartName").value)
                   
                   If SystemOptions.UserInterface = ArabicInterface Then
                   If Not IsNull(Rs1("TypDiv").value) Then
                   If (Rs1("TypDiv").value) = 1 Then
                   .TextMatrix(i, .ColIndex("InvesName")) = IIf(IsNull(Rs1("LandName").value), "", Rs1("LandName").value)
                   Else
                   .TextMatrix(i, .ColIndex("InvesName")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   End If
                   Else
                   .TextMatrix(i, .ColIndex("InvesName")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   End If
                   .TextMatrix(i, .ColIndex("unit")) = IIf(IsNull(Rs1("UntName").value), "", Rs1("UntName").value)
                   Else
                       If Not IsNull(Rs1("TypDiv").value) Then
                   If (Rs1("TypDiv").value) = 1 Then
                   .TextMatrix(i, .ColIndex("InvesName")) = IIf(IsNull(Rs1("LandNameE").value), "", Rs1("LandNameE").value)
                   Else
                   .TextMatrix(i, .ColIndex("InvesName")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                   End If
                   Else
                   .TextMatrix(i, .ColIndex("InvesName")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                   End If
                   
                   .TextMatrix(i, .ColIndex("unit")) = IIf(IsNull(Rs1("UntNameE").value), "", Rs1("UntNameE").value)
                   
                   End If
                     Rs1.MoveNext
             Next i
        End With
   '///////////////////////////
    Set Rs1 = New ADODB.Recordset
   sql = "SELECT     Valu, Remarks, FristDate, PartID"
   sql = sql & "       From dbo.TblSaleBilllInvestmentDet"
   sql = sql & "       Where (SBINVID = 1) And (TypeTrns = " & val(TxtSerial1.text) & ")"
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     With Me.VSFlexGrid1
                    For i = .FixedRows To Rs1.RecordCount
                   .rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("PaymentNo")) = IIf(IsNull(Rs1("PartID").value), 0, Rs1("PartID").value)
                   .TextMatrix(i, .ColIndex("PaymentValue")) = IIf(IsNull(Rs1("Valu").value), "", Rs1("Valu").value)
                   .TextMatrix(i, .ColIndex("DatePayment")) = IIf(IsNull(Rs1("FristDate").value), Date, Rs1("FristDate").value)
                   .TextMatrix(i, .ColIndex("Remrk")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                     Rs1.MoveNext
             Next i
        End With
      RelinGrid
      RelinGrid1
ErrTrap:
    End Sub



Private Sub DcbLand_Change()
Dim Fullcode As String
If val(DcbLand.BoundText) <> 0 Then
GetTblBuyLandRealEstate val(DcbLand.BoundText), Fullcode, 0
Me.Text1.text = Fullcode
GetInformationLand val(DcbLand.BoundText)
End If
End Sub

Private Sub DcbLand_Click(Area As Integer)
DcbLand_Change
End Sub

Private Sub ISButton6_Click()
If Me.TxtModFlg.text <> "R" Then
Dim X As Integer
Dim sql As String
With GridInstallments
If .rows < 2 Then
Exit Sub
    
Else
If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Will be deleted all linked operations", vbCritical + vbYesNo)
    Else
        X = MsgBox("”Ê› Ì „ Õ–› ﬂ· «·⁄„·Ì«  «·„— »ÿ… »Â–« «·”ÿ— Ê·«Ì„ﬂ‰ «· —«Ã⁄", vbCritical + vbYesNo)
End If
If X = vbNo Then Exit Sub
If Rd(1).value = False Then
   
       If val((.TextMatrix(.row, .ColIndex("BlockID")))) = 0 Then
         sql = "Update TblDivInvestInformation  set SalesBlocPayed=Null where ID= " & val(.TextMatrix(.row, .ColIndex("BlockID"))) & ""
       Cn.Execute sql
      End If
    
Else
  StrSQL = "Update TblSaleBilllInvestmentDet  set ReturnSal=Null where SBINVID=" & val(TxtBillNo.text) & " and BlockID= " & val(.TextMatrix(.row, .ColIndex("BlockID"))) & ""
       Cn.Execute StrSQL

End If
.RemoveItem .row
BtnUndo.Enabled = False
End If
End With
End If
End Sub

Private Sub ISButton8_Click()
FrmSearchinvestment.inde = 14
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End Sub

Private Sub lbl_Click(Index As Integer)
If Index = 13 Then
If Me.TxtModFlg.text <> "R" Then
If val(DcbTyp.ListIndex) = 1 Then
TxtNetComm.text = Round((val(Txtcommission.text) * val(lbl(13).Caption)) / 100, 2)
Else
TxtNetComm.text = Txtcommission.text
End If
End If
End If
End Sub

Private Sub Rd_Click(Index As Integer)
TxtBillNo.Visible = False
lbl(18).Visible = False
If Rd(1).value = True Then
TxtBillNo.Visible = True
lbl(18).Visible = True
End If

End Sub

Private Sub RdType_Click(Index As Integer)
Dim Dcombos As New ClsDataCombos
With Me.GridInstallments
If RdType(0).value = True Then
.TextMatrix(0, .ColIndex("InvesName")) = "«·„”«Â„…"
    Dcombos.GetLandActive DcbLand
 Else
   Dcombos.GetLandNotActive DcbLand
   .TextMatrix(0, .ColIndex("InvesName")) = "«·«—÷"
 End If
 End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim ID  As Integer
GetTblBuyLandRealEstate ID, Me.Text1.text, 1
DcbLand.BoundText = ID
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text9.text, EmpID
        DcbCus.BoundText = EmpID
    End If
End Sub

Private Sub TxtBillNo_Change()
If Me.TxtModFlg.text <> "R" Then
If val(TxtBillNo.text) <> 0 Then
If CheckDetrputed(val(TxtBillNo.text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ ⁄„· „—œÊœ«  Â–Â «·Õ—ﬂ…  „  Ê“Ì⁄ «—»«ÕÂ«"
Else
MsgBox "Can Not Return This Process profit distribution"
End If
clear_all Me
   GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 2
     VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 1

Exit Sub
End If
If CheckSaleorReturn(val(TxtBillNo.text)) = True Then
Exit Sub
End If
If Round(val(Me.TxtBillNo.text), 2) = Round(val(Me.TxtSerial1.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ —ﬁ„ «·Õ—ﬂ… Ì”«ÊÌ —ﬁ„ «·„—œÊœ«  ›Ì ‰›” «·Êﬁ "
Else
MsgBox "There can be no movement equals the number returns"
End If
Exit Sub
End If

    RsSavRec.Find "ID=" & val(TxtBillNo.text), , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        Else
        clear_all Me
           GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 2
     VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 1
End If
End If
End If
End Sub

Private Sub TxtBillNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtBillNo.text, 0)
End Sub

Private Sub TxtBillNo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearchinvestment.inde = 72
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End If
End Sub

Private Sub TxtCommission_Change()
If Me.TxtModFlg.text <> "R" Then
If val(DcbTyp.ListIndex) = 1 Then
TxtNetComm.text = Round((val(Txtcommission.text) * val(lbl(13).Caption)) / 100, 2)
Else
TxtNetComm.text = Txtcommission.text
End If
End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
If val(Me.DcbTypeSales.ListIndex) = 0 Then
DcbSales.BoundText = GeTEmpIDByEmpCode(TxtSearchCode.text, True)
Else
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode TxtSearchCode.text, EmpID
        DcbSales.BoundText = EmpID
    End If
End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "ID=" & RecId, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
  End Function
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.text)
    Me.TxtModFlg.text = "R"
    FiLLTXT
     BtnLast_Click
End Sub
' delet sub
Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim sql As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double
If CheckDetrputed(val(TxtSerial1.text)) = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «·Õ–› ·«‰Â Â–Â «·›« Ê—… „— »ÿ… » Ê“Ì⁄ «·«—»«Õ"
   Else
   MsgBox "You can not Delete this bill because it is linked to dividend-screen "
   End If
   Exit Sub
   End If
         If CheckReturn(val(TxtSerial1.text)) = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «·Õ–› ·«‰Â Â–Â «·›« Ê—…  „ ⁄„· ·Â« „—œÊœ« "
   Else
   MsgBox "You can not delete this bill because it return "
   End If
   Exit Sub
   End If
   
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
       End If
               Else
   If Rd(0).value = True Then
       StrSQL = "Update TblDivInvestInformation  set SalesPayed=Null,SalesBlocPayed=Null ,SalID=0 where SalID=" & val(TxtSerial1.text) & "  "
       Cn.Execute StrSQL
     Else
     StrSQL = "Update TblSaleBilllInvestmentDet  set ReturnSal=Null where SBINVID=" & val(TxtBillNo.text) & "  "
       Cn.Execute StrSQL
            sql = "Update TblDivInvestInformation  set SalesBlocPayed=1 where SalID=" & val(TxtBillNo.text) & ""
       Cn.Execute sql
     End If
    
                RsSavRec.Find "ID=" & val(TxtSerial1.text), , adSearchForward, 1
                       StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
                  
                  
                  StrSQL = "Delete From TblSaleBilllInvestmentDet Where SBINVID =" & val(TxtSerial1.text) & ""
                 Cn.Execute StrSQL, , adExecuteNoRecords
                                          RsSavRec.delete
                                
                                    
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 1
             VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 2
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
            lbl(13).Caption = 0
             
               '''''''''''''''''''''''''''''''

                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               End If

     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete the record"
            StrMSG = StrMSG & " Is related to with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
           Cn.Errors.Clear
    End Select

End Sub

' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
               btnSave_Click
        Case vbCancel
              Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.text = "N" Then
    XPDtbTrans.Enabled = True
   ' VSFlexGrid1.Enabled = True
   ' GridInstallments.Enabled = True
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg.text = "R" Then
  '  VSFlexGrid1.Enabled = False
   '  GridInstallments.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
           XPDtbTrans.Enabled = False
    End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.text = "E" Then
   'VSFlexGrid1.Enabled = True
   XPDtbTrans.Enabled = True
  ' GridInstallments.Enabled = True
       Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    '    Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub

' move btowen recored
Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
 
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.text <> "" Then
   If CheckDetrputed(val(TxtSerial1.text)) = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰  ⁄œÌ· ·«‰Â Â–Â «·›« Ê—… „— »ÿ… » Ê“Ì⁄ «·«—»«Õ"
   Else
   MsgBox "You can not modify this bill because it is linked to dividend-screen "
   End If
   Exit Sub
   End If
      If CheckReturn(val(TxtSerial1.text)) = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰  ⁄œÌ· ·«‰Â Â–Â «·›« Ê—…  „ ⁄„· ·Â« „—œÊœ« "
   Else
   MsgBox "You can not modify this bill because it return "
   End If
   Exit Sub
   End If
   
   
        TxtModFlg = "E"
            GridInstallments.rows = GridInstallments.rows + 1
        Me.DCboUserName.BoundText = user_id
        Frm2.Enabled = True
        Me.dcBranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & CHR(13)
            Msg = Msg & " You can not edit this the record now" & CHR(13)
            Msg = Msg & "It was being edited by another user on the network"
           
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    RdType(0).value = True
    lbl(6).Caption = 0
    TxtModFlg.text = "N"
    Rd_Click (0)
    Rd(0).value = True
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 2
     VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 1
    Me.DCboUserName.BoundText = user_id
    Me.dcBranch.BoundText = Current_branch
    dcBranch.SetFocus
  
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
      
        Exit Sub
    End If
BegnieWork:
     If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext
        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
       
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  MySQL = "SELECT     dbo.TblSaleBilllInvestment.ID, dbo.TblSaleBilllInvestment.RecordDate, dbo.TblSaleBilllInvestment.UserID, dbo.TblSaleBilllInvestment.SellerType, "
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestment.commission, dbo.TblSaleBilllInvestment.DesLocation, dbo.TblSaleBilllInvestment.Remarks, dbo.TblSaleBilllInvestment.PropertyDeed,"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestment.NorthlengthStr, dbo.TblSaleBilllInvestment.SouthlengthStr, dbo.TblSaleBilllInvestment.EastlengthStr,"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestment.WestlengthStr, dbo.TblSaleBilllInvestment.Northlength, dbo.TblSaleBilllInvestment.Southlength, dbo.TblSaleBilllInvestment.Eastlength,"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestment.Westlength, dbo.TblSaleBilllInvestment.Payment, dbo.TblSaleBilllInvestment.RecordNo, dbo.TblSaleBilllInvestment.CusID,"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestment.PaymentNo, dbo.TblSaleBilllInvestment.Period, dbo.TblSaleBilllInvestment.PeriodType, dbo.TblSaleBilllInvestment.RemarkPay,"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestment.FristDate, dbo.TblSaleBilllInvestment.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestment.SellerID, dbo.TblCustemers.CusName AS SelCusName, dbo.TblCustemers.CusNamee AS SelCusNameE,"
  MySQL = MySQL & "                    dbo.TblCustemers.Fullcode AS SelFullcode, dbo.TblEmployee.Emp_Name AS SelEmp_Name, dbo.TblEmployee.Fullcode AS SelEmpFullcode,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_Namee AS SelEmp_NameE, dbo.TblSaleBilllInvestment.LandID, dbo.TblBuyLanReEst.Name, dbo.TblBuyLanReEst.NameE,"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestment.Cus_ID, TblCustemers_1.CusName, TblCustemers_1.CusNamee, TblCustemers_1.Fullcode, dbo.TblSaleBilllInvestment.Cus_Tpe,"
  MySQL = MySQL & "                    dbo.TblInvestorType.Name AS TypName, dbo.TblInvestorType.NameE AS TypNameE, dbo.TblSaleBilllInvestmentDet.FristDate AS DetFristDate,"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestmentDet.Profit, dbo.TblSaleBilllInvestmentDet.TotalCost, dbo.TblSaleBilllInvestmentDet.MeterValue, dbo.TblSaleBilllInvestmentDet.Payed,"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestmentDet.Remarks AS DetRemarks, dbo.TblSaleBilllInvestmentDet.Net, dbo.TblSaleBilllInvestmentDet.TypeDis,"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestmentDet.DisValue, dbo.TblSaleBilllInvestmentDet.Total, dbo.TblSaleBilllInvestmentDet.Valu, dbo.TblSaleBilllInvestmentDet.Area,"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestmentDet.TypeTrns, dbo.TblSaleBilllInvestmentDet.InvesID, dbo.Tblinvestment.Name AS InvestName,"
  MySQL = MySQL & "                    dbo.Tblinvestment.NameE AS InvestNameE, dbo.TblSaleBilllInvestmentDet.DivID, dbo.TblSaleBilllInvestmentDet.PartID, dbo.TblDivInvesmentDet.PartNo"
  MySQL = MySQL & "    FROM         dbo.TblDivInvesmentDet RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestmentDet ON dbo.TblDivInvesmentDet.ID = dbo.TblSaleBilllInvestmentDet.PartID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.Tblinvestment ON dbo.TblSaleBilllInvestmentDet.InvesID = dbo.Tblinvestment.ID RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblSaleBilllInvestment ON dbo.TblSaleBilllInvestmentDet.SBINVID = dbo.TblSaleBilllInvestment.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblInvestorType ON dbo.TblSaleBilllInvestment.Cus_Tpe = dbo.TblInvestorType.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCustemers TblCustemers_1 ON dbo.TblSaleBilllInvestment.Cus_ID = TblCustemers_1.CusID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBuyLanReEst ON dbo.TblSaleBilllInvestment.LandID = dbo.TblBuyLanReEst.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblEmployee ON dbo.TblSaleBilllInvestment.SellerID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCustemers ON dbo.TblSaleBilllInvestment.SellerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.TblSaleBilllInvestment.BranchID = dbo.TblBranchesData.branch_id"
  MySQL = MySQL & "  Where (dbo.TblSaleBilllInvestment.ID =" & val(TxtSerial1.text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSalesBillInvestment.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSalesBillInvestmentE.rpt"
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

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
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

End Function

'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With
ErrTrap:
End Sub
' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            btnNew_Click
        Else
            Sendkeys "{TAB}"
        End If
    End If
    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If
    'End If
    Exit Sub
ErrTrap:
End Sub


Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Sales Bill  "
    Me.lbl(4).Caption = "No"
    Me.lbl(2).Caption = "Date"
    Me.lbl(7).Caption = "Branch"
    Label1(3).Caption = "Seller"
   lbl(1).Caption = "Type"
    Me.Label1(2).Caption = Me.Caption
    lbl(3).Caption = "Commission"
    Label1(6).Caption = "Land"
    lbl(11).Caption = "Property Deed"
    lbl(15).Caption = "Information"
    Frame8.Caption = "Border"
    Label7.Caption = "Number"
    Label19.Caption = "Writing"
    Label3.Caption = "North"
    Label23.Caption = "North"
    Label5.Caption = "South"
    Label21.Caption = "South"
    Label4.Caption = "East"
    Label22.Caption = "East"
    Label6.Caption = "West"
    Label20.Caption = "West"
    Label1(1).Caption = "Customer"
    Label1(0).Caption = "Type Cus."
    lbl(0).Caption = "Remarks"
    Label1(6).Caption = "Land"
    lbl(16).Caption = "Record No"
    lbl(17).Caption = "ID"
    Label1(4).Caption = "Payment"
    lbl(12).Caption = "Total"
    lbl(10).Caption = "Total"
    Frame7.Caption = "Data of Payment"
    ISButton2.Caption = "Add"
    Label1(11).Caption = "Period"
    Label1(8).Caption = "No"
    lbl(21).Caption = "Remarks"
    Label1(9).Caption = "First Date"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    '''''''''''''' next
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    C1Tab1.Caption = "Payments Data| Data"
    ISButton3.Caption = "Attachments"
  With Me.GridInstallments
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("InvesName")) = "Investment Name"
  .TextMatrix(0, .ColIndex("DivID")) = "No.Division"
  .TextMatrix(0, .ColIndex("Name")) = "Part No."
  .TextMatrix(0, .ColIndex("Area")) = "Area"
  .TextMatrix(0, .ColIndex("Valu")) = "Metre Value"
  .TextMatrix(0, .ColIndex("Total")) = "Total"
  .TextMatrix(0, .ColIndex("TypeDis")) = "Type Dis."
  .TextMatrix(0, .ColIndex("DisValue")) = "Dis.Value"
   .TextMatrix(0, .ColIndex("Net")) = "Net"
  .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
  End With
    With Me.VSFlexGrid1
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("PaymentNo")) = "Payment No."
  .TextMatrix(0, .ColIndex("DatePayment")) = "Date"
  .TextMatrix(0, .ColIndex("PaymentValue")) = "Value"
  .TextMatrix(0, .ColIndex("Remrk")) = "Remrk"
  End With
ErrTrap:
End Sub
Sub filgrid1()
Dim i As Integer
 VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
VSFlexGrid1.rows = 1
With VSFlexGrid1

.rows = .rows + val(TxtPaymentNo.text)
For i = 1 To .rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("PaymentValue")) = Round(val(lbl(13).Caption) / val(TxtPaymentNo.text), 2)
.TextMatrix(i, .ColIndex("PaymentNo")) = i
If i = 1 Then
.TextMatrix(i, .ColIndex("DatePayment")) = FristDate.value
TempDate.value = FristDate.value
Else
If val(Me.DcbPeriodsID.ListIndex) = 0 Then
TempDate.value = DateAdd("d", val(Me.txtPeriod.text), TempDate.value)
ElseIf val(Me.DcbPeriodsID.ListIndex) = 1 Then
TempDate.value = DateAdd("M", val(Me.txtPeriod.text), TempDate.value)
ElseIf val(Me.DcbPeriodsID.ListIndex) = 1 Then
TempDate.value = DateAdd("YYYY", val(Me.txtPeriod.text), TempDate.value)
End If
.TextMatrix(i, .ColIndex("DatePayment")) = TempDate.value
End If
.TextMatrix(i, .ColIndex("Remrk")) = Text11.text

Next i
'.AutoSize 0, .Cols - 1, False
End With
End Sub
Sub RelinGrid()
Dim summation As Double
Dim Counter As Integer
Dim i As Integer
Counter = 0
summation = 0
lbl(9).Caption = 0
With Me.VSFlexGrid1
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("PaymentNo"))) <> 0 Then
Counter = Counter + 1
.TextMatrix(i, .ColIndex("Ser")) = Counter
summation = summation + val(.TextMatrix(i, .ColIndex("PaymentValue")))
End If
Next i
lbl(9).Caption = summation

End With
End Sub
Sub RelinGrid1()
Dim summation As Double
Dim SumProfit As Double
Dim Counter As Integer
Dim i As Integer
Counter = 0
summation = 0
SumProfit = 0
lbl(13).Caption = 0
With Me.GridInstallments
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("BlockID"))) <> 0 Then
Counter = Counter + 1
.TextMatrix(i, .ColIndex("Ser")) = Counter
summation = summation + val(.TextMatrix(i, .ColIndex("Net")))
SumProfit = SumProfit + val(.TextMatrix(i, .ColIndex("Profit")))
End If
Next i
lbl(13).Caption = summation
lbl(14).Caption = SumProfit
End With
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblSaleBilllInvestment"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.text = rs.RecordCount + 1
    Else
        TxtSerial1.text = 1
    End If
   rs.Close
ErrTrap:
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal row As Long, ByVal Col As Long)
RelinGrid
End Sub

'+++++++++++++++++++++++++++++++++ end

Private Sub VSFlexGrid1_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid1
If Me.TxtModFlg.text = "R" Then
Cancel = True
Else
Select Case .ColKey(Col)
Case "PaymentNo"
Cancel = True
Case "DatePayment"
Cancel = True
Case "PaymentValue"
If opt(2).value = True Then
Cancel = False
Else
Cancel = True
End If

End Select
End If
End With

End Sub

Private Sub XPDtbTrans_Change()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.text = ""
End If
End Sub

Private Sub XPDtbTrans_Click()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.text = ""
End If
End Sub
