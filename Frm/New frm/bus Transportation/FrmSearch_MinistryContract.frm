VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSearch_MinistryContract 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13230
   Icon            =   "FrmSearch_MinistryContract.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   13230
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
   Begin VB.Frame frm_VehicleAllocation 
      BackColor       =   &H00E2E9E9&
      Height          =   8412
      Left            =   -30
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   -90
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   1332
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   6960
         Width           =   13092
         Begin VB.TextBox txtAllocNo 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   9108
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   240
            Width           =   2520
         End
         Begin VB.TextBox txtContractMinistryNo 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   5244
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   240
            Width           =   2208
         End
         Begin MSComCtl2.DTPicker FromDate3 
            Height          =   348
            Left            =   1584
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   240
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   96141315
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal FromDateH3 
            Height          =   312
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   556
         End
         Begin MSDataListLib.DataCombo dcMinistryContract 
            Height          =   288
            Left            =   5244
            TabIndex        =   68
            Top             =   600
            Width           =   2208
            _ExtentX        =   3889
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcDuration3 
            Height          =   288
            Left            =   9108
            TabIndex        =   69
            Top             =   960
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcSchoolFile 
            Height          =   288
            Left            =   9108
            TabIndex        =   70
            Top             =   600
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lblErrors2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Height          =   315
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   900
            Width           =   5235
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   252
            Index           =   23
            Left            =   11916
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   960
            Width           =   852
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄ﬁœ «·Ê“«—Ï"
            Height          =   312
            Index           =   22
            Left            =   7452
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   600
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„œ—”…"
            Height          =   312
            Index           =   21
            Left            =   11868
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «· Œ’Ì’"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   2880
            TabIndex        =   73
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ ⁄ﬁœ «·Ê“«—…"
            Height          =   312
            Index           =   17
            Left            =   7452
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   240
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· Œ’Ì’"
            Height          =   285
            Index           =   16
            Left            =   11970
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   270
            Width           =   810
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   588
         Left            =   0
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   0
         Width           =   13296
         _cx             =   23442
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
         Caption         =   "       »ÕÀ  Œ’Ì’ «·Õ«›·«    "
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
            TabIndex        =   78
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid fg_VA2 
         Height          =   6360
         Left            =   -120
         TabIndex        =   79
         Top             =   600
         Width           =   13212
         _cx             =   23304
         _cy             =   11218
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
         FormatString    =   $"FrmSearch_MinistryContract.frx":038A
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
   Begin VB.Frame frm_attribution 
      BackColor       =   &H00E2E9E9&
      Height          =   8292
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   1812
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   6360
         Width           =   13092
         Begin VB.TextBox txtfullcode 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   8280
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   960
            Width           =   1260
         End
         Begin VB.TextBox txtRecordno 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   10260
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   960
            Width           =   1608
         End
         Begin VB.TextBox XPTxtBoxName1 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   8280
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   600
            Width           =   3588
         End
         Begin VB.TextBox txtMinistryContractNo1 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   2400
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1800
            Visible         =   0   'False
            Width           =   1644
         End
         Begin VB.TextBox txtProcessNo1 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   10260
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   1608
         End
         Begin MSComCtl2.DTPicker dtpToDate1 
            Height          =   348
            Left            =   1584
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   600
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   96141315
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpFromDate1 
            Height          =   348
            Left            =   1584
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   240
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   96141315
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH1 
            Height          =   312
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal dtpToDateH1 
            Height          =   312
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal dtpSContractDateH1 
            Height          =   252
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   450
         End
         Begin MSComCtl2.DTPicker dtpSContractDate1 
            Height          =   348
            Left            =   1584
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   960
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   96141315
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpEContractDate1 
            Height          =   348
            Left            =   1584
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   96141315
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpEContractDateH1 
            Height          =   252
            Left            =   120
            TabIndex        =   43
            Top             =   1320
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   450
         End
         Begin MSDataListLib.DataCombo dcVendor1 
            Height          =   288
            Left            =   4404
            TabIndex        =   44
            Top             =   600
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcDuration1 
            Height          =   288
            Left            =   4404
            TabIndex        =   45
            Top             =   240
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCity1 
            Height          =   288
            Left            =   7080
            TabIndex        =   46
            Top             =   240
            Width           =   2028
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCustomer 
            Height          =   288
            Left            =   4404
            TabIndex        =   81
            Top             =   960
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lblErrors 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Height          =   315
            Left            =   4740
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   1410
            Width           =   5235
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂÊœ"
            Height          =   312
            Index           =   20
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   960
            Width           =   612
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ ⁄Âœ"
            Height          =   312
            Index           =   19
            Left            =   7572
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   960
            Width           =   468
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã· "
            Height          =   312
            Index           =   18
            Left            =   12000
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   960
            Width           =   828
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄ﬁœ"
            Height          =   288
            Index           =   14
            Left            =   12084
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   240
            Width           =   804
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ ⁄ﬁœ «·Ê“«—…"
            Height          =   312
            Index           =   13
            Left            =   1332
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   1800
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”„Ï «· ⁄«ﬁœ"
            Height          =   252
            Index           =   12
            Left            =   11772
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   600
            Width           =   1116
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   372
            Index           =   1
            Left            =   3024
            TabIndex        =   53
            Top             =   960
            Width           =   1056
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   2880
            TabIndex        =   52
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «· ⁄«ﬁœ „Ì·«œÏ"
            Height          =   252
            Index           =   11
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «‰ Â«¡ «· ⁄«ﬁœ „Ì·«œÏ"
            Height          =   372
            Index           =   7
            Left            =   2832
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   1320
            Width           =   1248
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Õ«›Ÿ…"
            Height          =   312
            Index           =   6
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «· ⁄·Ì„Ì…"
            Height          =   312
            Index           =   4
            Left            =   7080
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   600
            Width           =   972
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   252
            Index           =   2
            Left            =   5916
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   240
            Width           =   972
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   588
         Left            =   0
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   0
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
         Caption         =   "       »ÕÀ ⁄ﬁÊœ «·«”‰«œ   "
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
            TabIndex        =   58
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid fg_attr 
         Height          =   5760
         Left            =   -120
         TabIndex        =   59
         Top             =   600
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
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearch_MinistryContract.frx":0526
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
      TabIndex        =   29
      Top             =   8280
      Width           =   13212
      Begin ImpulseButton.ISButton Cmd 
         Height          =   432
         Index           =   0
         Left            =   7020
         TabIndex        =   30
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
         TabIndex        =   31
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
         TabIndex        =   32
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
   Begin VB.Frame frm_ministrycontract 
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
         Height          =   1812
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   6480
         Width           =   13092
         Begin VB.TextBox txtProcessNo 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   9720
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   1908
         End
         Begin VB.TextBox XPTxtBoxName 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   5844
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   5808
         End
         Begin VB.TextBox txtMinistryContractNo 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   5844
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   2088
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   312
            Left            =   1704
            TabIndex        =   8
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
            Format          =   96141315
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   312
            Left            =   1704
            TabIndex        =   9
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
            Format          =   96141315
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH 
            Height          =   312
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal dtpToDateH 
            Height          =   312
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal dtpSContractDateH 
            Height          =   252
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   450
         End
         Begin MSComCtl2.DTPicker dtpSContractDate 
            Height          =   252
            Left            =   1704
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   960
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   450
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   96141315
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpEContractDate 
            Height          =   252
            Left            =   1704
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   450
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   96141315
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpEContractDateH 
            Height          =   252
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   450
         End
         Begin MSDataListLib.DataCombo dcVendor 
            Height          =   288
            Left            =   5844
            TabIndex        =   16
            Top             =   960
            Width           =   2088
            _ExtentX        =   3678
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcDuration 
            Height          =   288
            Left            =   7668
            TabIndex        =   17
            Top             =   2400
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCity 
            Height          =   288
            Left            =   9708
            TabIndex        =   28
            Top             =   960
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
            Left            =   9720
            TabIndex        =   86
            Top             =   1320
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lblErrors3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Height          =   315
            Left            =   4230
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1320
            Width           =   5235
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   24
            Left            =   11856
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   1320
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   252
            Index           =   1
            Left            =   9252
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   2400
            Visible         =   0   'False
            Width           =   972
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «· ⁄·Ì„Ì…"
            Height          =   312
            Index           =   9
            Left            =   7932
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   960
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Õ«›Ÿ…"
            Height          =   312
            Index           =   10
            Left            =   11844
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   960
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «‰ Â«¡ «· ⁄«ﬁœ „Ì·«œÏ"
            Height          =   372
            Index           =   8
            Left            =   3072
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1320
            Width           =   1248
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «· ⁄«ﬁœ „Ì·«œÏ"
            Height          =   252
            Index           =   5
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   960
            Width           =   1500
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   3000
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   600
            Width           =   1176
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”„Ï «· ⁄«ﬁœ"
            Height          =   252
            Index           =   3
            Left            =   11628
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   600
            Width           =   1236
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ ⁄ﬁœ «·Ê“«—…"
            Height          =   312
            Index           =   15
            Left            =   7932
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄ﬁœ"
            Height          =   288
            Index           =   0
            Left            =   11940
            RightToLeft     =   -1  'True
            TabIndex        =   18
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
         Caption         =   "       »ÕÀ ⁄ﬁÊœ «·Ê“«—…   "
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
      Begin VSFlex8UCtl.VSFlexGrid Fg 
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
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearch_MinistryContract.frx":07F7
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
End
Attribute VB_Name = "FrmSearch_MinistryContract"
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
                If frm_ministrycontract.Visible = True Then
                    GetData
                ElseIf frm_attribution.Visible = True Then
                    GetData_Attr
                    
                ElseIf frm_VehicleAllocation.Visible = True Then
                        GetData_VehicleAllocation
                End If
        Case 1
            clear_all Me
            ResetDate
        Case 2
            Unload Me
    End Select

End Sub




Private Sub dcCity_Click(Area As Integer)
Dim str As String
Set Rs_Temp = New ADODB.Recordset
Set DCVendor.RowSource = Rs_Temp
If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
Else
    str = " Select ID , NameE   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
End If
fill_combo DCVendor, str
DCVendor.Refresh
End Sub

Private Sub dcCity1_Click(Area As Integer)

Dim str As String
Set Rs_Temp = New ADODB.Recordset
Set dcVendor1.RowSource = Rs_Temp
If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea  where cityid = " & val(dcCity1.BoundText)
Else
    str = " Select ID , NameE   from TblManagerialArea  where cityid = " & val(dcCity1.BoundText)
End If
fill_combo dcVendor1, str
dcVendor1.Refresh

End Sub

Private Sub dcCustomer_Click(Area As Integer)
Dim val1, val2, recordno As String, Fullcode As String
If dcCustomer.BoundText = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and cusid = " & dcCustomer.BoundText
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
         Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     End If
     TxtRecordNo.Text = recordno
     TxtFullcode.Text = Fullcode
     
End Sub

Private Sub fg_attr_Click()
On Error Resume Next
  Dim i As Integer
   i = val(fg_attr.TextMatrix(fg_attr.Row, fg_attr.ColIndex("IDMC")))
   
   If i > 0 Then
   
        If SendForm = "attributioncontract" Then
              FrmAttributionContract.Retrive (i)
              
        ElseIf SendForm = "AC" Then
             Set Rs_Temp = New ADODB.Recordset
            Rs_Temp.Open " select ProcessNo  from TblMinistryContract where  IDMC =  " & i, Cn, adOpenStatic, adLockOptimistic, adCmdText
             If Rs_Temp.RecordCount > 0 Then
                    FrmAttributionContract.dcMinistry.BoundText = i  ' IIf(IsNull(Rs_Temp("ProcessNo").value), "", Rs_Temp("ProcessNo").value)
             End If
        ElseIf SendForm = "VA" Then
            Set Rs_Temp = New ADODB.Recordset
            Rs_Temp.Open " select ProcessNo  from TblMinistryContract where  IDMC =  " & i, Cn, adOpenStatic, adLockOptimistic, adCmdText
             If Rs_Temp.RecordCount > 0 Then
                    FrmVehicleAllocation.dcMinistryContract.BoundText = IIf(IsNull(Rs_Temp("IDMC").value), "", Rs_Temp("IDMC").value)
             End If

        ElseIf SendForm = "ConfirmViolation" Then
                     FrmConfirmViolation.dcContract.BoundText = i

        ElseIf SendForm = "StopDeal" Then
                    FrmStopDealing.dcMinistry.BoundText = i

        ElseIf SendForm = "Report_scene" Then
                            frmReport_Scenes.dcMContract.BoundText = i
        ElseIf SendForm = "ReportScene_Attribuation" Then
                        frmReport_Scenes.dcMContract.BoundText = i


        End If
   
   
   End If


'Unload Me
ErrTrap:



End Sub

Private Sub Fg_Click()

' On Error GoTo ErrTrap
     
   Dim i As Integer
   i = val(FG.TextMatrix(FG.Row, FG.ColIndex("IDMC")))
   
   If i > 0 Then
        
        
        If SendForm = "MC" Then
              FrmMinistryContract.Retrive (i)
        ElseIf SendForm = "AC" Then
              
             Set Rs_Temp = New ADODB.Recordset
             Rs_Temp.Open " select ProcessNo  from TblMinistryContract where  IDMC =  " & i, Cn, adOpenStatic, adLockOptimistic, adCmdText
             If Rs_Temp.RecordCount > 0 Then
                    FrmAttributionContract.dcMinistry.BoundText = i  ' IIf(IsNull(Rs_Temp("ProcessNo").value), "", Rs_Temp("ProcessNo").value)
             End If
        ElseIf SendForm = "VA" Then
            ' Set Rs_Temp = New Adodb.Recordset
            ' Rs_Temp.Open " select ProcessNo  from TblMinistryContract where  IDMC =  " & i, Cn, adOpenStatic, adLockOptimistic, adCmdText
            ' If Rs_Temp.RecordCount > 0 Then
                    FrmVehicleAllocation.dcMinistryContract.BoundText = i ' IIf(IsNull(Rs_Temp("IDMC").value), "", Rs_Temp("IDMC").value)
            ' End If
        ElseIf SendForm = "R" Then
                    FrmRequest_MinistryContract.dcM.BoundText = i
        ElseIf SendForm = "ReportScene_MinistryContract" Then
                    frmReport_Scenes.dcMinistry.BoundText = i
        End If
   
   
   End If


'Unload Me
ErrTrap:

End Sub

Private Sub fg_VA2_Click()

   
   Dim i As Integer
   i = val(fg_VA2.TextMatrix(fg_VA2.Row, fg_VA2.ColIndex("IDVA")))
   
   If i > 0 Then
        
        
        If SendForm = "VA2" Then
              FrmVehicleAllocation.Retrive (i)
         End If
   
   End If


'Unload Me
ErrTrap:


End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
PutFormOnTop Me.hWnd, True
mdifrmmain.Enabled = False
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.getCountriesGovernments dcCity
    
    Dcombos.getCountriesGovernments dcCity1
    Dim str As String
    
    If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea "
    Else
    str = " Select ID , NameE   from TblManagerialArea "
    End If
    fill_combo DCVendor, str
     fill_combo dcVendor1, str
     
    str = "select id , name  from TblDurations "
    fill_combo dcDuration, str
    fill_combo dcDuration1, str
    

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
    
  ResetDate
    
    
    If SendForm = "ReportScene_MinistryContract" Or SendForm = "MC" Or SendForm = "AC" Or SendForm = "VA" Or SendForm = "R" Then
            frm_ministrycontract.Visible = True
    ElseIf SendForm = "ReportScene_Attribuation" Or SendForm = "attributioncontract" Or SendForm = "ConfirmViolation" Or SendForm = "StopDeal" Or SendForm = "Report_scene" Then
            frm_attribution.Visible = True
    ElseIf SendForm = "VA2" Then
            frm_VehicleAllocation.Visible = True
            
            Dim str2 As String
            str2 = " select id , name from TblSchooleFile  "
            fill_combo dcSchoolFile, str2
            
            
            str2 = " select id , name from TblDurations   "
            fill_combo dcDuration3, str2
            
            
           
            str2 = "      select IDMC ,MinistryContractNo  from TblMinistryContract  "
            fill_combo dcMinistryContract, str2
            
    End If
    
    Dcombos.GetCustomersSuppliers 2, dcCustomer
    Dcombos.GetBranches Dcbranch
End Sub

Private Sub ResetDate()

dtpFromDate.value = Date
   dtpToDate.value = Date
   dtpFromDateH.value = ToHijriDate(Date)
   dtpToDateH.value = ToHijriDate(Date)
   dtpSContractDate.value = Date
   dtpSContractDateH.value = ToHijriDate(Date)
   dtpEContractDate.value = Date
   dtpEContractDateH.value = ToHijriDate(Date)
   dtpFromDate1.value = Date
   dtpToDate1.value = Date
   dtpFromDateH1.value = ToHijriDate(Date)
   dtpToDateH1.value = ToHijriDate(Date)
   dtpSContractDate1.value = Date
   dtpSContractDateH1.value = ToHijriDate(Date)
   dtpEContractDate1.value = Date
   dtpEContractDateH1.value = ToHijriDate(Date)
    
   dtpFromDate.value = Null
   dtpToDate.value = Null
   dtpSContractDate.value = Null
   dtpEContractDate.value = Null
   dtpFromDate1.value = Null
   dtpToDate1.value = Null
   dtpSContractDate1.value = Null
   dtpEContractDate1.value = Null
    Fromdate3.value = Null

End Sub


Private Sub Form_Unload(Cancel As Integer)
mdifrmmain.Enabled = True
   ' FormPostion Me, SavePostion
   ' Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

        lblErrors3.Visible = False
        
        lblErrors3.Caption = ""
  
    StrSQL = "  SELECT dbo.TblMinistryContract.DurationID, dbo.TblMinistryContract.CityID, dbo.TblMinistryContract.VendorID, dbo.TblMinistryContract.ProcessNo, dbo.TblMinistryContract.Name, "
    StrSQL = StrSQL & "    dbo.TblMinistryContract.FromDate, dbo.TblMinistryContract.FromDateH, dbo.TblMinistryContract.ToDate, dbo.TblMinistryContract.ToDateH,"
    StrSQL = StrSQL & "                 dbo.TblMinistryContract.StudentCount, dbo.TblMinistryContract.StudentCustom, dbo.TblMinistryContract.DisCount, dbo.TblMinistryContract.StartContractDate,"
     StrSQL = StrSQL & "                dbo.TblMinistryContract.StartContractDateh, dbo.TblMinistryContract.EndContractDate, dbo.TblMinistryContract.EndContractDateh,"
     StrSQL = StrSQL & "                dbo.TblDurations.Name AS DurationName, dbo.TblCountriesGovernments.GovernmentName, dbo.TblMinistryContract.MinistryContractNo,"
      StrSQL = StrSQL & "               dbo.TblManagerialArea.Name AS MAName, dbo.TblMinistryContract.IDMC, dbo.TblMinistryContract.BranchID, dbo.TblBranchesData.branch_name,"
      StrSQL = StrSQL & "               dbo.TblBranchesData.branch_nameE"
      StrSQL = StrSQL & "     FROM     dbo.TblMinistryContract INNER JOIN"
       StrSQL = StrSQL & "              dbo.TblBranchesData ON dbo.TblMinistryContract.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
       StrSQL = StrSQL & "              dbo.TblManagerialArea ON dbo.TblMinistryContract.VendorID = dbo.TblManagerialArea.ID LEFT OUTER JOIN"
      StrSQL = StrSQL & "               dbo.TblCountriesGovernments ON dbo.TblMinistryContract.CityID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
      StrSQL = StrSQL & "               dbo.TblDurations ON dbo.TblMinistryContract.DurationID = dbo.TblDurations.ID"
        StrSQL = StrSQL & "   WHERE  (1 = 1) "
   
   
    If Me.txtProcessNo.Text <> "" Then
            StrSQL = StrSQL & "   and  IDMC =  '" & txtProcessNo.Text & "'"
    End If
    
    If Me.txtMinistryContractNo.Text <> "" Then
            StrSQL = StrSQL & "   and  MinistryContractNo =  '" & txtMinistryContractNo.Text & "'"
    End If

    If Me.XPTxtBoxName.Text <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.Name like  N'%" & XPTxtBoxName.Text & "%'"
    End If
    
     If Me.dcDuration.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
    If Me.dcCity.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.CityID =  " & val(Me.dcCity.BoundText)
    End If
    
     If Me.DCVendor.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.VendorID =  " & val(Me.DCVendor.BoundText)
    End If
     
    If Not IsNull(Me.dtpFromDate.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.FromDate  >=  '" & dtpFromDate.value & "'"
    End If
    
   If Not IsNull(Me.dtpToDate.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.ToDate  >=  '" & dtpToDate.value & "'"
    End If
    
     If Not IsNull(dtpSContractDate.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.StartContractDate  >=  '" & dtpSContractDate.value & "'"
    End If
    
    If Not IsNull(dtpEContractDate.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.EndContractDate  >=  '" & dtpEContractDate.value & "'"
    End If
    
     If Me.Dcbranch.BoundText <> "" Then
            StrSQL = StrSQL & "   and  dbo.TblMinistryContract.BranchID =  " & val(Me.Dcbranch.BoundText)
    End If
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By IDMC "
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
        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbApplicationModal, App.title
        lblErrors3.Visible = True
        lblErrors3.Caption = Msg
        Exit Sub
    Else
        lblErrors3.Caption = ""
        lblErrors3.Visible = False
        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                'Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                 .TextMatrix(i, .ColIndex("IDMC")) = IIf(IsNull(rs("IDMC").value), "", rs("IDMC").value)
                .TextMatrix(i, .ColIndex("ProcessNo")) = IIf(IsNull(rs("ProcessNo").value), "", rs("ProcessNo").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("MinistryContractNo")) = IIf(IsNull(rs("MinistryContractNo").value), "", rs("MinistryContractNo").value)
                .TextMatrix(i, .ColIndex("DurationName")) = IIf(IsNull(rs("DurationName").value), "", rs("DurationName").value)
                
               .TextMatrix(i, .ColIndex("MAName")) = IIf(IsNull(rs("MAName").value), "", rs("MAName").value)
               .TextMatrix(i, .ColIndex("GovernmentName")) = IIf(IsNull(rs("GovernmentName").value), "", rs("GovernmentName").value)
               .TextMatrix(i, .ColIndex("DurationName")) = IIf(IsNull(rs("DurationName").value), "", rs("DurationName").value)
                .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(rs("FromDate").value), "", rs("FromDate").value)
                .TextMatrix(i, .ColIndex("FromDateH")) = IIf(IsNull(rs("FromDateH").value), "", rs("FromDateH").value)
                .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(rs("ToDate").value), "", rs("ToDate").value)
                .TextMatrix(i, .ColIndex("ToDateH")) = IIf(IsNull(rs("ToDateH").value), "", rs("ToDateH").value)
                .TextMatrix(i, .ColIndex("EndContractDate")) = IIf(IsNull(rs("EndContractDate").value), "", rs("EndContractDate").value)
                .TextMatrix(i, .ColIndex("EndContractDateh")) = IIf(IsNull(rs("EndContractDateh").value), "", rs("EndContractDateh").value)
                .TextMatrix(i, .ColIndex("StartContractDate")) = IIf(IsNull(rs("StartContractDate").value), "", rs("StartContractDate").value)
                .TextMatrix(i, .ColIndex("StartContractDateh")) = IIf(IsNull(rs("StartContractDateh").value), "", rs("StartContractDateh").value)
                 .TextMatrix(i, .ColIndex("branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                 rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub

Public Sub GetData_Attr()

    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
   lblErrors.Caption = ""
  
  StrSQL = "SELECT dbo.TblDurations.Name AS DurationName, dbo.TblCountriesGovernments.GovernmentName, dbo.TblManagerialArea.Name AS MAName, dbo.TblAttributionContract.IDAC,"
  StrSQL = StrSQL & "  dbo.TblAttributionContract.ProcessNo , dbo.TblAttributionContract.Name , dbo.TblAttributionContract.FromDate ,"
  StrSQL = StrSQL & "  dbo.TblAttributionContract.FromDateH , dbo.TblAttributionContract.ToDate , dbo.TblAttributionContract.ToDateH ,"
  StrSQL = StrSQL & "  dbo.TblAttributionContract.VendorID , dbo.TblAttributionContract.StudentCount , dbo.TblAttributionContract.StudentCustom ,"
  StrSQL = StrSQL & "  dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.EndContractDateH , dbo.TblAttributionContract.MinistryContractNo ,"
  StrSQL = StrSQL & "  dbo.TblAttributionContract.DurationID"
  StrSQL = StrSQL & " , dbo.TblAttributionContract.StartContractDate, dbo.TblAttributionContract.StartContractDateh  "
  StrSQL = StrSQL & "  FROM     dbo.TblAttributionContract Left JOIN"
  StrSQL = StrSQL & "  dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID Left JOIN"
  StrSQL = StrSQL & "  dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID Left JOIN"
  StrSQL = StrSQL & "  dbo.TblCountriesGovernments ON dbo.TblAttributionContract.CityID = dbo.TblCountriesGovernments.GovernmentID"
    StrSQL = StrSQL & "     where 1 = 1  "
   
   
    If Me.txtProcessNo1.Text <> "" Then
            StrSQL = StrSQL & "   and  IDAC =  '" & txtProcessNo1.Text & "'"
    End If
    
    If Me.txtMinistryContractNo1.Text <> "" Then
            StrSQL = StrSQL & "   and  MinistryContractNo =  '" & txtMinistryContractNo1.Text & "'"
    End If

    If Me.XPTxtBoxName1.Text <> "" Then
            StrSQL = StrSQL & "   and  TblAttributionContract.Name like  N'%" & XPTxtBoxName1.Text & "%'"
    End If
    
     If Me.dcDuration1.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblAttributionContract.DurationID =  " & val(Me.dcDuration1.BoundText)
    End If
 
    If Me.dcCity1.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblAttributionContract.CityID =  " & val(Me.dcCity1.BoundText)
    End If
    
     If Me.dcVendor1.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblAttributionContract.Mangerialareaid =  " & val(Me.dcVendor1.BoundText)
    End If
     
    If Not IsNull(Me.dtpFromDate1.value) Then
            StrSQL = StrSQL & "   and  TblAttributionContract.FromDate  >=  '" & dtpFromDate1.value & "'"
    End If
    
   If Not IsNull(Me.dtpToDate1.value) Then
            StrSQL = StrSQL & "   and  TblAttributionContract.ToDate  >=  '" & dtpToDate1.value & "'"
    End If
    
     If Not IsNull(dtpSContractDate1.value) Then
            StrSQL = StrSQL & "   and  TblAttributionContract.StartContractDate  >=  '" & dtpSContractDate1.value & "'"
    End If
    
    If Not IsNull(dtpEContractDate1.value) Then
            StrSQL = StrSQL & "   and  TblAttributionContract.EndContractDate  >=  '" & dtpEContractDate1.value & "'"
    End If
    
    If Me.dcCustomer.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblAttributionContract.VendorID =  " & val(Me.dcCustomer.BoundText)
    End If
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By IDMC "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            'Me.Lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbApplicationModal, App.title
        'MsgBox Msg, vbApplicationModal, App.title
        lblErrors.Visible = True
        lblErrors.Caption = Msg
        Exit Sub
    Else
        lblErrors.Visible = False
        lblErrors.Caption = ""
    

        With Me.fg_attr
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                'Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                 .TextMatrix(i, .ColIndex("IDMC")) = IIf(IsNull(rs("IDAC").value), "", rs("IDAC").value)
                .TextMatrix(i, .ColIndex("ProcessNo")) = IIf(IsNull(rs("ProcessNo").value), "", rs("ProcessNo").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("MinistryContractNo")) = IIf(IsNull(rs("MinistryContractNo").value), "", rs("MinistryContractNo").value)
                .TextMatrix(i, .ColIndex("DurationName")) = IIf(IsNull(rs("DurationName").value), "", rs("DurationName").value)
                
               .TextMatrix(i, .ColIndex("MAName")) = IIf(IsNull(rs("MAName").value), "", rs("MAName").value)
               .TextMatrix(i, .ColIndex("GovernmentName")) = IIf(IsNull(rs("GovernmentName").value), "", rs("GovernmentName").value)
               .TextMatrix(i, .ColIndex("DurationName")) = IIf(IsNull(rs("DurationName").value), "", rs("DurationName").value)
                .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(rs("FromDate").value), "", rs("FromDate").value)
                .TextMatrix(i, .ColIndex("FromDateH")) = IIf(IsNull(rs("FromDateH").value), "", rs("FromDateH").value)
                .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(rs("ToDate").value), "", rs("ToDate").value)
                .TextMatrix(i, .ColIndex("ToDateH")) = IIf(IsNull(rs("ToDateH").value), "", rs("ToDateH").value)
                .TextMatrix(i, .ColIndex("EndContractDate")) = IIf(IsNull(rs("EndContractDate").value), "", rs("EndContractDate").value)
                .TextMatrix(i, .ColIndex("EndContractDateh")) = IIf(IsNull(rs("EndContractDateh").value), "", rs("EndContractDateh").value)
                .TextMatrix(i, .ColIndex("StartContractDate")) = IIf(IsNull(rs("StartContractDate").value), "", rs("StartContractDate").value)
                .TextMatrix(i, .ColIndex("StartContractDateh")) = IIf(IsNull(rs("StartContractDateh").value), "", rs("StartContractDateh").value)
                
                 rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub


Public Sub GetData_VehicleAllocation()

    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    lblErrors2.Visible = False
    lblErrors2.Caption = ""
    
    StrSQL = ""
    
    StrSQL = StrSQL & "   SELECT dbo.TblVehicleAllocation.SchoolFileID, dbo.TblVehicleAllocation.IDVA, dbo.TblVehicleAllocation.ProcessNo, dbo.TblVehicleAllocation.ProcessDate,"
    StrSQL = StrSQL & "                 dbo.TblVehicleAllocation.IDMC, dbo.TblVehicleAllocation.DurationID, dbo.TblDurations.Name AS DurationName, dbo.TblSchooleFile.Name AS SchoolName,"
    StrSQL = StrSQL & "                  dbo.TblVehicleAllocation.MinistryNo , dbo.TblMinistryContract.MinistryContractNo, dbo.TblVehicleAllocation.FromDate, dbo.TblVehicleAllocation.FromdateH"
    StrSQL = StrSQL & "   FROM     dbo.TblVehicleAllocation INNER JOIN"
    StrSQL = StrSQL & "                    dbo.TblSchooleFile ON dbo.TblVehicleAllocation.SchoolFileID = dbo.TblSchooleFile.ID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblMinistryContract ON dbo.TblVehicleAllocation.IDMC = dbo.TblMinistryContract.IDMC INNER JOIN"
    StrSQL = StrSQL & "                dbo.TblDurations ON dbo.TblVehicleAllocation.DurationID = dbo.TblDurations.ID"
  

    StrSQL = StrSQL & "     where 1 = 1  "
   
   
    If Me.txtAllocNo.Text <> "" Then
            StrSQL = StrSQL & "   and  IDva =  " & val(txtAllocNo.Text)
    End If
    
    If Me.txtContractMinistryNo.Text <> "" Then
            StrSQL = StrSQL & "   and  MinistryContractNo =  '" & txtContractMinistryNo.Text & "'"
    End If

    
     If Me.dcDuration3.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblVehicleAllocation.DurationID =  " & val(Me.dcDuration3.BoundText)
    End If
 
    If Me.dcSchoolFile.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblVehicleAllocation.SchoolFileID =  " & val(Me.dcSchoolFile.BoundText)
    End If
    
     If Me.dcMinistryContract.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblVehicleAllocation.IDMC  =  " & val(Me.dcMinistryContract.BoundText)
    End If
     
    If Not IsNull(Me.Fromdate3.value) Then
            StrSQL = StrSQL & "   and  TblVehicleAllocation.FromDate  =  '" & Fromdate3.value & "'"
    End If
    
   
     
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By IDVA "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            'Me.Lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbApplicationModal, App.title
        lblErrors2.Visible = True
        lblErrors2.Caption = Msg
        Exit Sub
        
        
    Else
        lblErrors2.Visible = False
        lblErrors2.Caption = ""
        With Me.fg_VA2
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                'Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("IDVA")) = IIf(IsNull(rs("IDVA").value), "", rs("IDVA").value)
                .TextMatrix(i, .ColIndex("MinistryNo")) = IIf(IsNull(rs("MinistryNo").value), "", rs("MinistryNo").value)
                .TextMatrix(i, .ColIndex("SchoolName")) = IIf(IsNull(rs("SchoolName").value), "", rs("SchoolName").value)
                .TextMatrix(i, .ColIndex("DurationName")) = IIf(IsNull(rs("DurationName").value), "", rs("DurationName").value)
                .TextMatrix(i, .ColIndex("MinistryContractNo")) = IIf(IsNull(rs("MinistryContractNo").value), "", rs("MinistryContractNo").value)
                .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(rs("FromDate").value), "", rs("FromDate").value)
                .TextMatrix(i, .ColIndex("FromdateH")) = IIf(IsNull(rs("FromdateH").value), "", rs("FromdateH").value)
                 rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub




Private Sub ChangeLang()

End Sub



Private Sub txtfullcode_Change()


Dim val1, val2
If TxtFullcode.Text = "" Then Exit Sub
Dim str As String, recordno As String, CusID As String
recordno = ""
CusID = ""

    str = " select * From TblCustemers where Type=2  and fullcode = '" & TxtFullcode & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
           
     Else
        TxtRecordNo.Text = ""
        dcCustomer.BoundText = ""
    End If
    
    TxtRecordNo.Text = recordno
    dcCustomer.BoundText = CusID


End Sub

Private Sub txtRecordNo_Change()
Dim val1, val2, CusID As String, Fullcode As String
If TxtRecordNo.Text = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and recordno = '" & TxtRecordNo.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
         CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
          Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     Else
        dcCustomer.BoundText = ""
        TxtFullcode.Text = ""
    End If
   dcCustomer.BoundText = CusID
    TxtFullcode.Text = Fullcode
End Sub
