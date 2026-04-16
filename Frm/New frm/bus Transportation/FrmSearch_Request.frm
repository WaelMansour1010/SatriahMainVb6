VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSearch_Request 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13245
   Icon            =   "FrmSearch_Request.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame frm_MinistryRequest 
      Height          =   8412
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   -120
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Height          =   1692
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   6480
         Width           =   13092
         Begin VB.ComboBox cbType3 
            Height          =   288
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   1440
            Visible         =   0   'False
            Width           =   2292
         End
         Begin VB.TextBox txtID3 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   9720
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   240
            Width           =   1908
         End
         Begin MSComCtl2.DTPicker dtpToDate3 
            Height          =   312
            Left            =   1704
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   600
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   103153667
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpFromDate3 
            Height          =   312
            Left            =   1680
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   240
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   103153667
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH3 
            Height          =   312
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal dtpToDateH3 
            Height          =   312
            Left            =   120
            TabIndex        =   74
            Top             =   600
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   556
         End
         Begin MSDataListLib.DataCombo dcMonth3 
            Height          =   288
            Left            =   5880
            TabIndex        =   75
            Top             =   720
            Width           =   2292
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcDuration3 
            Height          =   288
            Left            =   9720
            TabIndex        =   76
            Top             =   720
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch3 
            Height          =   288
            Left            =   5880
            TabIndex        =   77
            Top             =   240
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMinistryContract3 
            Height          =   288
            Left            =   9720
            TabIndex        =   88
            Top             =   1200
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄ﬁœ «·Ê“«—Ï"
            Height          =   312
            Index           =   19
            Left            =   11844
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   1200
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   18
            Left            =   8136
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   312
            Index           =   17
            Left            =   7932
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   720
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   312
            Index           =   16
            Left            =   11844
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   3000
            TabIndex        =   81
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   3
            Left            =   3144
            TabIndex        =   80
            Top             =   600
            Width           =   1176
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’—›"
            Height          =   312
            Index           =   14
            Left            =   7692
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   1440
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   288
            Index           =   13
            Left            =   11940
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   240
            Width           =   924
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   588
         Left            =   0
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   120
         Width           =   13176
         _cx             =   23230
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     »ÕÀ «” Õﬁ«ﬁ«  ⁄ﬁÊœ «·Ê“«—…   "
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
            TabIndex        =   86
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid fg_MinistryRequest 
         Height          =   5760
         Left            =   -120
         TabIndex        =   87
         Top             =   720
         Width           =   13212
         _cx             =   23304
         _cy             =   10160
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
         FormatString    =   $"FrmSearch_Request.frx":038A
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
   Begin VB.Frame Frm_VendorRequest 
      BackColor       =   &H00E2E9E9&
      Height          =   8412
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   -120
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   1692
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   6480
         Width           =   13092
         Begin VB.TextBox txtID2 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   9720
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   1908
         End
         Begin VB.ComboBox cbType2 
            Height          =   288
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   1440
            Visible         =   0   'False
            Width           =   2292
         End
         Begin MSComCtl2.DTPicker dtpToDate2 
            Height          =   312
            Left            =   1704
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   600
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   103153667
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpFromDate2 
            Height          =   312
            Left            =   1680
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   240
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   103153667
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH2 
            Height          =   312
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal dtpToDateH2 
            Height          =   312
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   556
         End
         Begin MSDataListLib.DataCombo dcMonth2 
            Height          =   288
            Left            =   5880
            TabIndex        =   54
            Top             =   720
            Width           =   2292
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcDuration2 
            Height          =   288
            Left            =   9720
            TabIndex        =   55
            Top             =   720
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch2 
            Height          =   288
            Left            =   5880
            TabIndex        =   56
            Top             =   240
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   288
            Index           =   12
            Left            =   11940
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   240
            Width           =   924
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’—›"
            Height          =   312
            Index           =   11
            Left            =   7692
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1440
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   2
            Left            =   3144
            TabIndex        =   61
            Top             =   600
            Width           =   1176
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   3000
            TabIndex        =   60
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   312
            Index           =   8
            Left            =   11844
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   312
            Index           =   7
            Left            =   7932
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   720
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   6
            Left            =   8136
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   240
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   588
         Left            =   0
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   120
         Width           =   13176
         _cx             =   23230
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     »ÕÀ «” Õﬁ«ﬁ«  «·„ ⁄ÂœÌ‰   "
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
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid fg_VendorRequest 
         Height          =   5760
         Left            =   -120
         TabIndex        =   66
         Top             =   720
         Width           =   13212
         _cx             =   23304
         _cy             =   10160
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearch_Request.frx":04D5
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
   Begin VB.Frame Frm_VendorReceipt 
      BackColor       =   &H00E2E9E9&
      Height          =   8412
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   -120
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   1692
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   6480
         Width           =   13092
         Begin VB.ComboBox CbType1 
            Height          =   288
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   2292
         End
         Begin VB.TextBox txtID1 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   9720
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   1908
         End
         Begin MSComCtl2.DTPicker dtpToDate1 
            Height          =   312
            Left            =   1704
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   600
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   103153667
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpFromDate1 
            Height          =   312
            Left            =   1704
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   240
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   103153667
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH1 
            Height          =   312
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal dtpToDateH1 
            Height          =   312
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   556
         End
         Begin MSDataListLib.DataCombo dcMonth1 
            Height          =   288
            Left            =   5880
            TabIndex        =   33
            Top             =   720
            Width           =   2292
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcDuration1 
            Height          =   288
            Left            =   9720
            TabIndex        =   34
            Top             =   720
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch1 
            Height          =   288
            Left            =   9720
            TabIndex        =   35
            Top             =   1200
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   1
            Left            =   11856
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   1200
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   312
            Index           =   2
            Left            =   7932
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   720
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   312
            Index           =   3
            Left            =   11844
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   3000
            TabIndex        =   39
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   1
            Left            =   3144
            TabIndex        =   38
            Top             =   600
            Width           =   1176
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’—›"
            Height          =   312
            Index           =   4
            Left            =   7932
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   288
            Index           =   5
            Left            =   11940
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   924
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   588
         Left            =   0
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   120
         Width           =   13176
         _cx             =   23230
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     »ÕÀ ”‰œ ’—›   "
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
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid fg_VendorReceipt 
         Height          =   5760
         Left            =   -120
         TabIndex        =   45
         Top             =   720
         Width           =   13212
         _cx             =   23304
         _cy             =   10160
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
         FormatString    =   $"FrmSearch_Request.frx":0664
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   852
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   8400
      Width           =   13212
      Begin ImpulseButton.ISButton Cmd 
         Height          =   432
         Index           =   0
         Left            =   7020
         TabIndex        =   19
         Top             =   240
         Width           =   996
         _ExtentX        =   1746
         _ExtentY        =   767
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   432
         Index           =   1
         Left            =   5928
         TabIndex        =   20
         Top             =   240
         Width           =   1032
         _ExtentX        =   1826
         _ExtentY        =   767
         ButtonPositionImage=   1
         Caption         =   "„”Õ"
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
      Begin ImpulseButton.ISButton Cmd 
         Cancel          =   -1  'True
         Height          =   432
         Index           =   2
         Left            =   4920
         TabIndex        =   21
         Top             =   240
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   767
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
   End
   Begin VB.Frame frm_ExchangeRequest 
      BackColor       =   &H00E2E9E9&
      Height          =   8412
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   -120
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Height          =   1692
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   6480
         Width           =   13092
         Begin VB.TextBox txtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   672
            Left            =   5400
            MaxLength       =   50
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   600
            Width           =   2868
         End
         Begin VB.ComboBox CbType 
            Height          =   288
            Left            =   -240
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1200
            Visible         =   0   'False
            Width           =   3732
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   9720
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   1908
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   312
            Left            =   1704
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   600
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   103153667
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   312
            Left            =   1704
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   240
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   103153667
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH 
            Height          =   312
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal dtpToDateH 
            Height          =   312
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   556
         End
         Begin MSDataListLib.DataCombo dcMonth 
            Height          =   288
            Left            =   9720
            TabIndex        =   10
            Top             =   960
            Width           =   1932
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcDuration 
            Height          =   288
            Left            =   9720
            TabIndex        =   11
            Top             =   600
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   288
            Left            =   5400
            TabIndex        =   22
            Top             =   240
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   288
            Index           =   21
            Left            =   8580
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   600
            Width           =   924
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   24
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   312
            Index           =   9
            Left            =   11652
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   960
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   312
            Index           =   10
            Left            =   11844
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   600
            Width           =   1020
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   3000
            TabIndex        =   15
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   0
            Left            =   3144
            TabIndex        =   14
            Top             =   600
            Width           =   1176
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’—›"
            Height          =   312
            Index           =   15
            Left            =   3132
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   1200
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   288
            Index           =   0
            Left            =   11940
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   924
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   588
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   13176
         _cx             =   23230
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "      »ÕÀ ÿ·»«  ’—› „ ⁄ÂœÌ‰   "
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
      Begin VSFlex8UCtl.VSFlexGrid FgExchangeRequest 
         Height          =   5760
         Left            =   -120
         TabIndex        =   3
         Top             =   720
         Width           =   13212
         _cx             =   23304
         _cy             =   10160
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
         FormatString    =   $"FrmSearch_Request.frx":07AE
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
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   312
      Index           =   20
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   90
      Top             =   0
      Width           =   1020
   End
End
Attribute VB_Name = "FrmSearch_Request"
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
    Select Case Index

        Case 0
                If frm_ExchangeRequest.Visible = True Then
                    GetData_ExchangeRequest
                ElseIf Frm_VendorReceipt.Visible = True Then
                    GetData_VendorReceipt
                ElseIf Frm_VendorRequest.Visible = True Then
                      GetData_VendorRequest
                ElseIf frm_MinistryRequest.Visible = True Then
                        GetData_MinistryRequest
                End If
        Case 1
            clear_all Me
            ResetDate
        Case 2
            Unload Me
    End Select

End Sub


 

 

 


Private Sub dcDuration_Click(Area As Integer)
Dim i As Integer, j As Integer, Str As String
    i = val(dcDuration.BoundText)
    
    If i > 0 Then
        Str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo dcMonth, Str
    Else
        Str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo dcMonth, Str
    End If
End Sub

Private Sub dcDuration1_Click(Area As Integer)
Dim i As Integer, j As Integer, Str As String
    i = val(dcDuration1.BoundText)
    
    If i > 0 Then
        Str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo dcMonth1, Str
    Else
        Str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo dcMonth1, Str
    End If
End Sub

Private Sub dcDuration2_Click(Area As Integer)
Dim i As Integer, j As Integer, Str As String
    i = val(dcDuration2.BoundText)
    
    If i > 0 Then
        Str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo dcMonth2, Str
    Else
        Str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo dcMonth2, Str
    End If

End Sub

Private Sub dtpFromDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
 dtpFromDateH.value = ToHijriDate(dtpFromDate.value)
End Sub

Private Sub dtpFromDate1_Change()
    dtpFromDateH1.value = ToHijriDate(dtpFromDate1.value)
End Sub



Private Sub dtpFromDate2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
dtpFromDateH2.value = ToHijriDate(dtpFromDate2.value)

End Sub

Private Sub dtpFromDate3_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
dtpFromDateH3.value = ToHijriDate(dtpFromDate3.value)
End Sub

Private Sub dtpFromDateH_GotFocus()
VBA.Calendar = vbCalGreg
        dtpFromDate.value = ToGregorianDate(dtpFromDateH.value)
End Sub

Private Sub dtpFromDateH1_GotFocus()
 VBA.Calendar = vbCalGreg
        dtpFromDate1.value = ToGregorianDate(dtpFromDateH1.value)
End Sub

Private Sub dtpFromDateH2_GotFocus()
 VBA.Calendar = vbCalGreg
        dtpFromDate2.value = ToGregorianDate(dtpFromDateH2.value)
End Sub

Private Sub dtpFromDateH3_GotFocus()
 VBA.Calendar = vbCalGreg
        dtpFromDate3.value = ToGregorianDate(dtpFromDateH3.value)
End Sub

Private Sub dtpToDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
dtpToDateH.value = ToHijriDate(dtpToDate.value)
End Sub

Private Sub dtpToDate1_Change()
dtpToDateH1.value = ToHijriDate(dtpToDate1.value)
End Sub

Private Sub dtpToDate2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
dtpToDateH2.value = ToHijriDate(dtpToDate2.value)
End Sub

Private Sub dtpToDate3_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
dtpToDateH3.value = ToHijriDate(dtpToDate3.value)
End Sub

Private Sub dtpToDateH_GotFocus()
 VBA.Calendar = vbCalGreg
        dtpToDate.value = ToGregorianDate(dtpToDateH.value)
End Sub

Private Sub dtpToDateH1_GotFocus()
 VBA.Calendar = vbCalGreg
        dtpToDate1.value = ToGregorianDate(dtpToDateH1.value)
End Sub

Private Sub dtpToDateH2_GotFocus()
 VBA.Calendar = vbCalGreg
        dtpToDate2.value = ToGregorianDate(dtpToDateH2.value)
End Sub

Private Sub dtpToDateH3_GotFocus()
 VBA.Calendar = vbCalGreg
        dtpToDate3.value = ToGregorianDate(dtpToDateH3.value)

End Sub

Private Sub fg_MinistryRequest_Click()
  Dim i As Integer
   i = val(fg_MinistryRequest.TextMatrix(fg_MinistryRequest.Row, fg_MinistryRequest.ColIndex("ID")))
   
   If i > 0 Then
        If SendForm = "MR_MR" Then
              FrmRequest_MinistryContract.Retrive (i)
        End If
   End If
End Sub

Private Sub fg_VendorReceipt_Click()

   Dim i As Integer
   i = val(fg_VendorReceipt.TextMatrix(fg_VendorReceipt.Row, fg_VendorReceipt.ColIndex("ID")))
   
   If i > 0 Then
        If SendForm = "VR_VR" Then
              FrmVendorReceipt1.Retrive (i)
        ElseIf SendForm = "VR_ER" Then
              FrmVendorReceipt1.Text1.text = i
              FrmVendorReceipt1.Retrive_Depend
        End If
   End If
'Unload Me
ErrTrap:

End Sub

Private Sub fg_VendorRequest_Click()
  Dim i As Integer
   i = val(fg_VendorRequest.TextMatrix(fg_VendorRequest.Row, fg_VendorRequest.ColIndex("ID")))
   
   If i > 0 Then
        
        If SendForm = "VR_VREQ" Then
              FrmRequest1.Retrive (i)
        ElseIf SendForm = "MR_VR" Then
                FrmExchangeRequest.Text1.text = i
                FrmExchangeRequest.Retrive_Data
        End If
   End If
End Sub

Private Sub FgExchangeRequest_Click()


' On Error GoTo ErrTrap
     
   Dim i As Integer
   i = val(FgExchangeRequest.TextMatrix(FgExchangeRequest.Row, FgExchangeRequest.ColIndex("ID")))
   
   If i > 0 Then
        
        If SendForm = "ER_ER" Then
              FrmExchangeRequest.Retrive (i)
        ElseIf SendForm = "VR_ER" Then
              FrmVendorReceipt1.Text1.text = i
              FrmVendorReceipt1.Retrive_Depend
        ElseIf SendForm = "Payments" Then
                    FrmPayments.TxtOrderSuppler.text = i
        End If
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
    Dim GrdBack As ClsBackGroundPic
    Set DCboSearch = New clsDCboSearch
    Set GrdBack = New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    
    
    Dim Str As String
    
    Dcombos.GetBranches Dcbranch
    Str = "select id , name  from TblDurations "
    fill_combo dcDuration, Str
    With cbType
        If SystemOptions.UserInterface = ArabicInterface Then
                .Clear
                .AddItem ("‰ﬁœÏ")
                .AddItem ("‘Ìﬂ")
        Else
                .Clear
                .AddItem ("Cash")
                .AddItem ("Cheque")
        End If
    End With
    
    
    Dcombos.GetBranches dcBranch1
    Str = "select id , name  from TblDurations "
    fill_combo dcDuration1, Str
      With CbType1
        If SystemOptions.UserInterface = ArabicInterface Then
                .Clear
                .AddItem ("‰ﬁœÏ")
                .AddItem ("‘Ìﬂ")
        Else
                .Clear
                .AddItem ("Cash")
                .AddItem ("Cheque")
        End If
    End With
    
    Dcombos.GetBranches dcBranch2
    Str = "select id , name  from TblDurations "
    fill_combo dcDuration2, Str
      With cbType2
        If SystemOptions.UserInterface = ArabicInterface Then
                .Clear
                .AddItem ("‰ﬁœÏ")
                .AddItem ("‘Ìﬂ")
        Else
                .Clear
                .AddItem ("Cash")
                .AddItem ("Cheque")
        End If
    End With
     
    Dcombos.GetBranches dcBranch3
    Str = "select id , name  from TblDurations "
    fill_combo dcDuration3, Str
    Str = " select  IDMC  , MinistryContractNo  from TblMinistryContract  "
    fill_combo dcMinistryContract3, Str
  
      
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
      
 
    
    ResetDate
    
    
    If SendForm = "ER_ER" Or SendForm = "VR_ER" Or SendForm = "Payments" Then
            frm_ExchangeRequest.Visible = True
    ElseIf SendForm = "VR_VR" Then
            Frm_VendorReceipt.Visible = True
    ElseIf SendForm = "VR_VREQ" Or SendForm = "MR_VR" Then
            Frm_VendorRequest.Visible = True
    ElseIf SendForm = "MR_MR" Then
            frm_MinistryRequest.Visible = True
    End If
    
  
End Sub

Private Sub ResetDate()

   dtpFromDate.value = Date
   dtpToDate.value = Date
   dtpFromDateH.value = ToHijriDate(Date)
   dtpToDateH.value = ToHijriDate(Date)
   dtpFromDate.value = Null
   dtpToDate.value = Null
      
   dtpFromDate1.value = Date
   dtpToDate1.value = Date
   dtpFromDateH1.value = ToHijriDate(Date)
   dtpToDateH1.value = ToHijriDate(Date)
   dtpFromDate1.value = Null
   dtpToDate1.value = Null
   
   
   dtpFromDate2.value = Date
   dtpToDate2.value = Date
   dtpFromDateH2.value = ToHijriDate(Date)
   dtpToDateH2.value = ToHijriDate(Date)
   dtpFromDate2.value = Null
   dtpToDate2.value = Null
   
   
   dtpFromDate3.value = Date
   dtpToDate3.value = Date
   dtpFromDateH3.value = ToHijriDate(Date)
   dtpToDateH3.value = ToHijriDate(Date)
   dtpFromDate3.value = Null
   dtpToDate3.value = Null
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
mdifrmmain.Enabled = True
   ' FormPostion Me, SavePostion
   ' Set DCboSearch = Nothing
End Sub

Public Sub GetData_ExchangeRequest()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

  
    StrSQL = StrSQL & "  SELECT  TblExchangeRequest.remarks , dbo.TblExchangeRequest.ID, dbo.TblExchangeRequest.ExchangeType, dbo.TblExchangeRequest.DurationID, dbo.TblExchangeRequest.Month, "
    StrSQL = StrSQL & "             dbo.TblExchangeRequest.Date, dbo.TblExchangeRequest.DateH, dbo.TblExchangeRequest.BranchID, dbo.TblDurations.Name AS DurName,"
    StrSQL = StrSQL & "                 dbo.TblDurations_Details.Name AS MonthName, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL & "  FROM     dbo.TblExchangeRequest INNER JOIN"
    StrSQL = StrSQL & "            dbo.TblDurations ON dbo.TblExchangeRequest.DurationID = dbo.TblDurations.ID INNER JOIN"
    StrSQL = StrSQL & "            dbo.TblBranchesData ON dbo.TblExchangeRequest.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
    StrSQL = StrSQL & "            dbo.TblDurations_Details ON dbo.TblExchangeRequest.Month = dbo.TblDurations_Details.ID"
    StrSQL = StrSQL & "           where  1= 1  "
   
    If Me.txtID.text <> "" Then
            StrSQL = StrSQL & "   and  TblExchangeRequest.ID = " & val(txtID.text)
    End If
    
    If Me.cbType.ListIndex <> -1 Then
            StrSQL = StrSQL & "   and  TblExchangeRequest.ExchangeType =  " & cbType.ListIndex
    End If
    
    If Me.dcDuration.BoundText <> "" Then
            StrSQL = StrSQL & "   and  dbo.TblExchangeRequest.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
    If Me.dcMonth.BoundText <> "" Then
            StrSQL = StrSQL & "   and  Month =  " & val(Me.dcMonth.BoundText)
    End If
     
    If Not IsNull(Me.dtpFromDate.value) Then
            StrSQL = StrSQL & "   and  Date >=  " & SQLDate(dtpFromDate.value, True)
    End If
    
    If Not IsNull(Me.dtpToDate.value) Then
            StrSQL = StrSQL & "   and  Date  >=  " & SQLDate(dtpToDate.value, True)
    End If
      
    If Me.Dcbranch.BoundText <> "" Then
            StrSQL = StrSQL & "   and dbo.TblExchangeRequest.BranchID =  " & val(Me.Dcbranch.BoundText)
    End If
    
    If TxtRemarks.text <> "" Then
            StrSQL = StrSQL & "   and  TblExchangeRequest.remarks  like '%" & TxtRemarks.text & "%'"
    End If
     
    StrSQL = StrSQL
    StrSQL = StrSQL & "  Order By TblExchangeRequest.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
          '  Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.Lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
      '  MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.FgExchangeRequest
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                'Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
            Dim Typ As Integer
            
            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                
                 Typ = IIf(IsNull(rs("ExchangeType").value), -1, rs("ExchangeType").value)
                 
                 If Typ = 0 Then
                      .TextMatrix(i, .ColIndex("Type")) = "‰ﬁœÏ"
                 ElseIf Typ = 1 Then
                         .TextMatrix(i, .ColIndex("Type")) = "‘Ìﬂ"
                 End If
                 
                .TextMatrix(i, .ColIndex("DurNAme")) = IIf(IsNull(rs("DurNAme").value), "", rs("DurNAme").value)
                .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("Date")) = IIf(IsNull(rs("Date").value), "", rs("Date").value)
                .TextMatrix(i, .ColIndex("DateH")) = IIf(IsNull(rs("DateH").value), "", rs("DateH").value)
                .TextMatrix(i, .ColIndex("remarks")) = IIf(IsNull(rs("remarks").value), "", rs("remarks").value)
                 rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub


Public Sub GetData_VendorReceipt()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

  
    StrSQL = StrSQL & "  SELECT dbo.TblBranchesData.branch_name, dbo.TblVendorReceipt.ID, dbo.TblVendorReceipt.BranchID, dbo.TblVendorReceipt.ExchangeType, dbo.TblVendorReceipt.DurationID, "
    StrSQL = StrSQL & "  dbo.TblDurations.Name AS DurName, dbo.TblDurations_Details.Name AS MonthName, dbo.TblVendorReceipt.Date, dbo.TblVendorReceipt.DateH,"
    StrSQL = StrSQL & "  dbo.TblVendorReceipt.Month"
    StrSQL = StrSQL & "  FROM     dbo.TblVendorReceipt INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblDurations ON dbo.TblVendorReceipt.DurationID = dbo.TblDurations.ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblDurations_Details ON dbo.TblVendorReceipt.Month = dbo.TblDurations_Details.ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.TblVendorReceipt.BranchID = dbo.TblBranchesData.branch_id"
    StrSQL = StrSQL & "  where  1= 1  "
   
    If Me.txtID1.text <> "" Then
            StrSQL = StrSQL & "   and  TblVendorReceipt.ID = " & val(txtID1.text)
    End If
    
    If Me.CbType1.ListIndex <> -1 Then
            StrSQL = StrSQL & "   and  TblVendorReceipt.ExchangeType =  " & CbType1.ListIndex
    End If
    
    If Me.dcDuration1.BoundText <> "" Then
            StrSQL = StrSQL & "   and  dbo.TblVendorReceipt.DurationID =  " & val(Me.dcDuration1.BoundText)
    End If
 
    If Me.dcMonth1.BoundText <> "" Then
            StrSQL = StrSQL & "   and  Month =  " & val(Me.dcMonth1.BoundText)
    End If
     
    If Not IsNull(Me.dtpFromDate1.value) Then
            StrSQL = StrSQL & "   and  Date >=  " & SQLDate(dtpFromDate1.value, True)
    End If
    
    If Not IsNull(Me.dtpToDate1.value) Then
            StrSQL = StrSQL & "   and  Date  >=  " & SQLDate(dtpToDate1.value, True)
    End If
      
    If Me.dcBranch1.BoundText <> "" Then
            StrSQL = StrSQL & "   and dbo.TblVendorReceipt.BranchID =  " & val(Me.dcBranch1.BoundText)
    End If
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By TblVendorReceipt.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
          '  Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.Lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
       ' MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.fg_VendorReceipt
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                'Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
            Dim Typ As Integer
            
            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                
                 Typ = IIf(IsNull(rs("ExchangeType").value), -1, rs("ExchangeType").value)
                 
                 If Typ = 0 Then
                      .TextMatrix(i, .ColIndex("Type")) = "‰ﬁœÏ"
                 ElseIf Typ = 1 Then
                         .TextMatrix(i, .ColIndex("Type")) = "‘Ìﬂ"
                 End If
                 
                .TextMatrix(i, .ColIndex("DurNAme")) = IIf(IsNull(rs("DurNAme").value), "", rs("DurNAme").value)
                .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("Date")) = IIf(IsNull(rs("Date").value), "", rs("Date").value)
                .TextMatrix(i, .ColIndex("DateH")) = IIf(IsNull(rs("DateH").value), "", rs("DateH").value)
                 rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub

Public Sub GetData_VendorRequest()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

  
    StrSQL = StrSQL & "  SELECT dbo.TblBranchesData.branch_name, dbo.TblDurations.Name AS DurName, dbo.TblDurations_Details.Name AS MonthName, dbo.TblExchangeRequest2.BranchID, "
    StrSQL = StrSQL & "                 dbo.TblExchangeRequest2.Date, dbo.TblExchangeRequest2.DateH, dbo.TblExchangeRequest2.Month, dbo.TblExchangeRequest2.DurationID,"
    StrSQL = StrSQL & "                dbo.TblExchangeRequest2.ExchangeType , dbo.TblExchangeRequest2.id"
    StrSQL = StrSQL & "  FROM     dbo.TblDurations INNER JOIN"
    StrSQL = StrSQL & "                dbo.TblExchangeRequest2 ON dbo.TblDurations.ID = dbo.TblExchangeRequest2.DurationID INNER JOIN"
    StrSQL = StrSQL & "                dbo.TblDurations_Details ON dbo.TblExchangeRequest2.Month = dbo.TblDurations_Details.ID INNER JOIN"
    StrSQL = StrSQL & "                dbo.TblBranchesData ON dbo.TblExchangeRequest2.BranchID = dbo.TblBranchesData.branch_id "
   
    If Me.txtID2.text <> "" Then
            StrSQL = StrSQL & "   and  TblExchangeRequest2.ID = " & val(txtID2.text)
    End If
    
    If Me.cbType2.ListIndex <> -1 Then
            StrSQL = StrSQL & "   and  TblExchangeRequest2.ExchangeType =  " & cbType2.ListIndex
    End If
    
    If Me.dcDuration2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  dbo.TblExchangeRequest2.DurationID =  " & val(Me.dcDuration2.BoundText)
    End If
 
    If Me.dcMonth2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  Month =  " & val(Me.dcMonth2.BoundText)
    End If
     
    If Not IsNull(Me.dtpFromDate2.value) Then
            StrSQL = StrSQL & "   and  Date >=  " & SQLDate(dtpFromDate2.value, True)
    End If
    
    If Not IsNull(Me.dtpToDate2.value) Then
            StrSQL = StrSQL & "   and  Date  >=  " & SQLDate(dtpToDate2.value, True)
    End If
      
    If Me.dcBranch2.BoundText <> "" Then
            StrSQL = StrSQL & "   and dbo.TblExchangeRequest2.BranchID =  " & val(Me.dcBranch2.BoundText)
    End If
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By TblExchangeRequest2.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
          '  Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.Lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
      '  MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.fg_VendorRequest
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                'Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
            Dim Typ As Integer
            
            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                
                 Typ = IIf(IsNull(rs("ExchangeType").value), -1, rs("ExchangeType").value)
                 
                 If Typ = 0 Then
                      .TextMatrix(i, .ColIndex("Type")) = "‰ﬁœÏ"
                 ElseIf Typ = 1 Then
                         .TextMatrix(i, .ColIndex("Type")) = "‘Ìﬂ"
                 End If
                 
                .TextMatrix(i, .ColIndex("DurNAme")) = IIf(IsNull(rs("DurNAme").value), "", rs("DurNAme").value)
                .TextMatrix(i, .ColIndex("month")) = IIf(IsNull(rs("monthname").value), "", rs("monthname").value)
                .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("Date")) = IIf(IsNull(rs("Date").value), "", rs("Date").value)
                .TextMatrix(i, .ColIndex("DateH")) = IIf(IsNull(rs("DateH").value), "", rs("DateH").value)
                 rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub

Public Sub GetData_MinistryRequest()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

  
    StrSQL = StrSQL & "  SELECT dbo.TblBranchesData.branch_name, dbo.TblDurations.Name AS DurName, dbo.TblDurations_Details.Name AS MonthName, dbo.TblRequest_MinistryContract.DurationID, "
    StrSQL = StrSQL & "               dbo.TblRequest_MinistryContract.Month, dbo.TblRequest_MinistryContract.Date, dbo.TblRequest_MinistryContract.ExchangeType, dbo.TblRequest_MinistryContract.ID,"
    StrSQL = StrSQL & "               dbo.TblRequest_MinistryContract.DateH, dbo.TblRequest_MinistryContract.BranchID, dbo.TblRequest_MinistryContract.MinstryID, dbo.TblMinistryContract.Name,"
    StrSQL = StrSQL & "               dbo.TblMinistryContract.MinistryContractNo"
    StrSQL = StrSQL & "  FROM     dbo.TblDurations INNER JOIN"
    StrSQL = StrSQL & "                 dbo.TblRequest_MinistryContract ON dbo.TblDurations.ID = dbo.TblRequest_MinistryContract.DurationID INNER JOIN"
    StrSQL = StrSQL & "            dbo.TblDurations_Details ON dbo.TblRequest_MinistryContract.Month = dbo.TblDurations_Details.ID INNER JOIN"
    StrSQL = StrSQL & "                dbo.TblBranchesData ON dbo.TblRequest_MinistryContract.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
    StrSQL = StrSQL & "               dbo.TblMinistryContract ON dbo.TblRequest_MinistryContract.MinstryID = dbo.TblMinistryContract.IDMC"
    StrSQL = StrSQL & "  where 1 =1 "
    
    If Me.txtID3.text <> "" Then
            StrSQL = StrSQL & "   and  TblRequest_MinistryContract.ID = " & val(txtID3.text)
    End If
    
'    If Me.cbType2.ListIndex <> -1 Then
'            strSQL = strSQL & "   and  TblRequest_MinistryContract.ExchangeType =  " & cbType2.ListIndex
'    End If
    
    If Me.dcDuration3.BoundText <> "" Then
            StrSQL = StrSQL & "   and  dbo.TblRequest_MinistryContract.DurationID =  " & val(Me.dcDuration3.BoundText)
    End If
 
    If Me.dcMonth3.BoundText <> "" Then
            StrSQL = StrSQL & "   and  Month =  " & val(Me.dcMonth3.BoundText)
    End If
     
    If Not IsNull(Me.dtpFromDate3.value) Then
            StrSQL = StrSQL & "   and  Date >=  " & SQLDate(dtpFromDate3.value, True)
    End If
    
    If Not IsNull(Me.dtpToDate3.value) Then
            StrSQL = StrSQL & "   and  Date  >=  " & SQLDate(dtpToDate3.value, True)
    End If
      
    If Me.dcBranch3.BoundText <> "" Then
            StrSQL = StrSQL & "   and dbo.TblRequest_MinistryContract.BranchID =  " & val(Me.dcBranch3.BoundText)
    End If
     
    If dcMinistryContract3.BoundText <> "" Then
            StrSQL = StrSQL & "   and dbo.TblRequest_MinistryContract.MinstryID =  " & val(Me.dcMinistryContract3.BoundText)
    End If
     
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By TblRequest_MinistryContract.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
          '  Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.Lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
     '   MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.fg_MinistryRequest
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                'Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
            Dim Typ As Integer
            
            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                
                 Typ = IIf(IsNull(rs("ExchangeType").value), -1, rs("ExchangeType").value)
                 
                             
                .TextMatrix(i, .ColIndex("DurNAme")) = IIf(IsNull(rs("DurNAme").value), "", rs("DurNAme").value)
                .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("Date")) = IIf(IsNull(rs("Date").value), "", rs("Date").value)
                .TextMatrix(i, .ColIndex("DateH")) = IIf(IsNull(rs("DateH").value), "", rs("DateH").value)
                 rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub



Private Sub ChangeLang()

End Sub



