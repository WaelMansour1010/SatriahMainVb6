VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDriverAllocation 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄„·Ì«   Œ’Ì’ «·”«∆ﬁ ··Õ«›·…"
   ClientHeight    =   9315
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   13455
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   580
   Icon            =   "FrmDriverAllocation.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   13455
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   15120
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   2076
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9315
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13455
      _cx             =   23733
      _cy             =   16431
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
      Align           =   5
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   612
         Left            =   72
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   600
         Width           =   13308
         _cx             =   23469
         _cy             =   1085
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
         Begin VB.TextBox xtxtID 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10560
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   120
            Width           =   1572
         End
         Begin MSComCtl2.DTPicker Date 
            Height          =   312
            Left            =   2868
            TabIndex        =   45
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   98959361
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal dateH 
            Height          =   252
            Left            =   1320
            TabIndex        =   46
            Top             =   120
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   450
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   6000
            TabIndex        =   48
            Top             =   120
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Index           =   11
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   120
            Width           =   768
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· Œ’Ì’ "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Index           =   8
            Left            =   4476
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   120
            Width           =   1080
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· Œ’Ì’"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Index           =   7
            Left            =   12012
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   120
            Width           =   1128
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   1452
         Left            =   72
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1320
         Width           =   13308
         _cx             =   23469
         _cy             =   2566
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
         Begin VB.TextBox txtcarcode 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10512
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   240
            Width           =   1572
         End
         Begin VB.TextBox txtDay 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   120
            Visible         =   0   'False
            Width           =   612
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8292
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   960
            Width           =   1032
         End
         Begin VB.TextBox txtCode 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10512
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   960
            Width           =   1572
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   312
            Left            =   2844
            TabIndex        =   3
            Top             =   240
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   98959361
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   312
            Left            =   2844
            TabIndex        =   8
            Top             =   600
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   98959361
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcDriver 
            Height          =   315
            Left            =   3165
            TabIndex        =   7
            Top             =   960
            Width           =   3900
            _ExtentX        =   6879
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
            Height          =   456
            Index           =   9
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   888
            _ExtentX        =   1561
            _ExtentY        =   794
            ButtonPositionImage=   1
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
            ButtonImage     =   "FrmDriverAllocation.frx":038A
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
         Begin MSDataListLib.DataCombo dcCars 
            Height          =   315
            Left            =   5925
            TabIndex        =   2
            Top             =   240
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin Dynamic_Byte.NourHijriCal FromDateH 
            Height          =   252
            Left            =   1296
            TabIndex        =   4
            Top             =   240
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   450
         End
         Begin Dynamic_Byte.NourHijriCal ToDateH 
            Height          =   252
            Left            =   1296
            TabIndex        =   9
            Top             =   600
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   450
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”«∆ﬁ «·Õ«·Ï"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Index           =   12
            Left            =   9528
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   600
            Visible         =   0   'False
            Width           =   888
         End
         Begin VB.Label lblEmp 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   252
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   600
            Width           =   3252
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «··ÊÕ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Index           =   1
            Left            =   9648
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   768
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ  ”·Ì„ «·Õ«›·…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   4455
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «” ·«„ «·Õ«›·…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   4350
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   600
            Width           =   1425
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·”«∆ﬁ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Index           =   6
            Left            =   7356
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   960
            Width           =   768
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÂÊÌ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Index           =   3
            Left            =   9648
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   960
            Width           =   768
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·”«∆ﬁ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Index           =   2
            Left            =   12048
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   960
            Width           =   1128
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·Õ«›·…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Index           =   5
            Left            =   12468
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   708
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   4860
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Width           =   13305
         _cx             =   23469
         _cy             =   8572
         Appearance      =   2
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmDriverAllocation.frx":6BEC
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
         ExplorerBar     =   1
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   600
         Index           =   5
         Left            =   -48
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   -48
         Width           =   13548
         _cx             =   23892
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Picture         =   "FrmDriverAllocation.frx":6E16
         Caption         =   "   ⁄„·Ì«   Œ’Ì’ «·”«∆ﬁ ··Õ«›·…  "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   0
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   0
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
         Begin VB.TextBox toid 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   375
            Index           =   0
            Left            =   1695
            TabIndex        =   21
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmDriverAllocation.frx":7AF0
            ColorButton     =   -2147483634
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
            Left            =   630
            TabIndex        =   22
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmDriverAllocation.frx":7E8A
            ColorButton     =   -2147483634
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
            Height          =   372
            Index           =   1
            Left            =   2220
            TabIndex        =   23
            Top             =   120
            Width           =   492
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmDriverAllocation.frx":8224
            ColorButton     =   -2147483634
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
            Left            =   1155
            TabIndex        =   24
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmDriverAllocation.frx":85BE
            ColorButton     =   -2147483634
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   495
         Left            =   75
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   7905
         Width           =   5760
         _cx             =   10160
         _cy             =   873
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   312
            Index           =   4
            Left            =   816
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   120
            Width           =   1104
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   312
            Index           =   0
            Left            =   3804
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   120
            Width           =   1104
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   312
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   120
            Width           =   660
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   312
            Left            =   2928
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   120
            Width           =   828
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   690
         Left            =   75
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   8505
         Width           =   13245
         _cx             =   23363
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
            Height          =   468
            Index           =   0
            Left            =   11544
            TabIndex        =   32
            Top             =   120
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   820
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
            ButtonImage     =   "FrmDriverAllocation.frx":8958
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
            Height          =   468
            Index           =   1
            Left            =   9840
            TabIndex        =   33
            Top             =   120
            Width           =   1656
            _ExtentX        =   2910
            _ExtentY        =   820
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
            ButtonImage     =   "FrmDriverAllocation.frx":F1BA
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
            Height          =   468
            Index           =   2
            Left            =   8220
            TabIndex        =   34
            Top             =   120
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   820
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
            ButtonImage     =   "FrmDriverAllocation.frx":15A1C
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
            Height          =   468
            Index           =   3
            Left            =   6564
            TabIndex        =   35
            Top             =   120
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   820
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
            ButtonImage     =   "FrmDriverAllocation.frx":1C27E
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
            Height          =   468
            Index           =   4
            Left            =   5016
            TabIndex        =   36
            Top             =   120
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   820
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
            ButtonImage     =   "FrmDriverAllocation.frx":22AE0
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
            Height          =   468
            Index           =   6
            Left            =   1848
            TabIndex        =   37
            Top             =   120
            Width           =   1656
            _ExtentX        =   2910
            _ExtentY        =   820
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
            ButtonImage     =   "FrmDriverAllocation.frx":29342
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
            Height          =   468
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   820
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
            ButtonImage     =   "FrmDriverAllocation.frx":52F64
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
            Height          =   468
            Index           =   7
            Left            =   3516
            TabIndex        =   39
            Top             =   120
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   820
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
            ButtonImage     =   "FrmDriverAllocation.frx":597C6
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
End
Attribute VB_Name = "FrmDriverAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim Rs_Temp2 As ADODB.Recordset

Public LongRow As Long

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long


    If dcBranch.BoundText = "" Then
            MsgBox ("«Œ — «·›—⁄ «Ê·« ")
            Exit Sub
    End If
        
    If Grid.Rows = 1 Then
            MsgBox (" ﬁ„ »⁄„·ÌÂ  Œ’Ì’ ”«∆ﬁÌ‰ «Ê·«  ")
        Exit Sub
    End If
    

    On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
 
     End If
   
    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
        rs.AddNew
        xtxtID.Text = CStr(new_id("TblDriverAllocation", "IDDA", "", True))
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete TblDriverAllocation_Details where IDDA=" & val(Me.xtxtID.Text)
    End If
    rs("IDDA").value = xtxtID.Text
    rs("UserID") = user_id
    rs("CreationDate") = Date
    
   rs("branchid") = dcBranch.BoundText
   rs("date") = Me.Date.value
   rs("DateH") = Me.DateH.value
   
    
    rs.update
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open "TblDriverAllocation_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Integer
    With Me.Grid

        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("Emp_id")) <> "" Then
                RsDev.AddNew
                RsDev("ID").value = CStr(new_id("TblDriverAllocation_Details", "ID", "", True))
                RsDev("IDDA").value = xtxtID.Text
                RsDev("emp_id").value = .TextMatrix(i, .ColIndex("Emp_id"))
                RsDev("emp_code").value = .TextMatrix(i, .ColIndex("Emp_Code"))
                RsDev("emp_name").value = .TextMatrix(i, .ColIndex("Emp_Name"))
                RsDev("Board").value = .TextMatrix(i, .ColIndex("car"))
                RsDev("carid").value = .TextMatrix(i, .ColIndex("carid"))
                RsDev("fromdate").value = .TextMatrix(i, .ColIndex("fromdate"))
                RsDev("todate").value = .TextMatrix(i, .ColIndex("todate"))
                RsDev("fromdateh").value = .TextMatrix(i, .ColIndex("fromdateh"))
                RsDev("todateh").value = .TextMatrix(i, .ColIndex("todateh"))
                RsDev("daycount").value = .TextMatrix(i, .ColIndex("daycount"))
                RsDev("NumEkama").value = .TextMatrix(i, .ColIndex("NumEkama"))
                RsDev.update
                
                Dim str3 As String
                str3 = "select * from TblCarsData where  id =   " & val(.TextMatrix(i, .ColIndex("carid")))
                Set Rs_Temp2 = New ADODB.Recordset
                Rs_Temp2.Open str3, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Rs_Temp2.RecordCount > 0 Then
                        Rs_Temp2("Emp_ID").value = IIf(IsNull(.TextMatrix(i, .ColIndex("Emp_id"))), Null, .TextMatrix(i, .ColIndex("Emp_id")))
                        Rs_Temp2.update
                End If
            End If
            
        Next i

    End With
 
    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.Text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '  Fg_Journal.Enabled = False
    End Select

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

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
        
            TxtModFlg.Text = "N"
            clear_all Me
            Me.xtxtID.Text = CStr(new_id("TblDriverAllocation", "IDDA", "", True))
       
         
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            Grid.Enabled = True
            
            lbl(12).Visible = False
            lblEmp.Caption = ""
        
        Case 1

            TxtModFlg.Text = "E"
           ' Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True

        Case 2
   
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
               Exit Sub
           End If

             Del_Trans
        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
       Case 6
            Unload Me

        Case 7
print_report
        Case 8
           TxtModFlg.Text = "N"
           clear_all Me
           'Me.xtxtID.text = CStr(new_id("opr_Employee", "ID", "", True))
            
           Grid.Rows = Grid.Rows + 1
           Grid.Enabled = True
  
         Case 9
         
         addrow
        
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub addrow()

If dcCars.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox ("„‰ ›÷··ﬂ «Œ — «·Õ«›·… ")
        Else
                MsgBox ("Select Vechile")
        End If
        dcCars.SetFocus
        Exit Sub
End If


If DCDriver.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«Œ —  «·”«∆ﬁ ")
        Else
        MsgBox ("Select Driver ")
        End If
        DCDriver.SetFocus
        SendKeys ("{F4}")
        Exit Sub
End If


Dim i As Integer
Grid.Rows = Grid.Rows + 1
i = Grid.Rows
i = i - 1
With Grid
  .TextMatrix(i, .ColIndex("Serial")) = i - 1
  .TextMatrix(i, .ColIndex("Emp_Code")) = txtCode.Text
  .TextMatrix(i, .ColIndex("Emp_id")) = DCDriver.BoundText
  
  .TextMatrix(i, .ColIndex("car")) = dcCars.Text
  .TextMatrix(i, .ColIndex("carid")) = dcCars.BoundText
  
  .TextMatrix(i, .ColIndex("NumEkama")) = TxtId.Text
  .TextMatrix(i, .ColIndex("Emp_Name")) = DCDriver.Text
  .TextMatrix(i, .ColIndex("FromDate")) = FromDate.value
  .TextMatrix(i, .ColIndex("FromDateH")) = FromDateH.value
  .TextMatrix(i, .ColIndex("ToDate")) = ToDate.value
  .TextMatrix(i, .ColIndex("ToDateH")) = todateH.value
  .TextMatrix(i, .ColIndex("daycount")) = val(txtDay.Text)
End With
'

dcCars.BoundText = ""
DCDriver.BoundText = ""
txtCode.Text = ""
TxtId.Text = ""
lblEmp.Caption = ""
End Sub


Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If xtxtID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From TblDriverAllocation Where IDDA =" & val(Me.xtxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 rs.MoveFirst
                 Cn.Execute "delete TblDriverAllocation_Details where IDDA=" & val(Me.xtxtID)

                If rs.RecordCount < 1 Then
                Grid.Clear flexClearScrollable, flexClearEverything
                    Grid.Rows = 2
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text
        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)
            
            lbl(12).Visible = False
            lblEmp.Caption = ""
        Case "E"
            Retrive
            Me.TxtModFlg.Text = "R"
    End Select
    Exit Sub
ErrTrap:
End Sub




Private Sub Command1_Click()

End Sub

Private Sub CmdAttach_Click()
            On Error Resume Next
'ShowAttachments XPTxtBoxID, "0701201405"
 
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtId, "15062020007"


End Sub

Private Sub Date_Change()
DateH.value = ToHijriDate(Me.Date.value)
End Sub

Private Sub DateH_LostFocus()
VBA.Calendar = vbCalGreg
Me.Date.value = ToGregorianDate(Me.DateH.value)
End Sub

Private Sub dcCars_Click(Area As Integer)
Dim val1 As String, val2 As String
If dcCars.BoundText = "" Then Exit Sub
Dim i  As Integer, str As String, Emp_id   As Integer
i = dcCars.BoundText
If i > 0 Then
    str = " select  Fullcode , name ,Emp_ID   from TblCarsData   where id =  " & i
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        val1 = IIf(IsNull(Rs_Temp("Fullcode").value), "", Rs_Temp("Fullcode").value)
        
        Emp_id = IIf(IsNull(Rs_Temp("Emp_ID").value), 0, Rs_Temp("Emp_ID").value)
       
       If IsNull(Rs_Temp("Emp_ID").value) = False Then
            Dim str5 As String
            str5 = " Select * from TblEmployee where Emp_ID =  " & Emp_id
            Set Rs_Temp2 = New ADODB.Recordset
            Rs_Temp2.Open str5, Cn, adOpenStatic, adLockOptimistic, adCmdText
            
            If Rs_Temp2.RecordCount > 0 Then
                     lblEmp.Caption = IIf(IsNull(Rs_Temp2("Emp_Name").value), "", Rs_Temp2("Emp_Name").value)
            End If
            lbl(12).Visible = True
        Else
         lbl(12).Visible = False
         lblEmp.Caption = ""
         
        End If
        
        
    End If
End If
txtcarcode.Text = val1

End Sub

Private Sub dcCars_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
            Unload FrmCasrShearches
            FrmCasrShearches.SendForm = "DriverAllocation"
            FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub dcDriver_Click(Area As Integer)

Dim val1 As String, val2 As String
If DCDriver.BoundText = "" Then Exit Sub
Dim i  As Integer, str As String
i = DCDriver.BoundText
If i > 0 Then
    str = " select * from TblEmployee  where Emp_ID =  " & i
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        val1 = IIf(IsNull(Rs_Temp("Emp_Code").value), "", Rs_Temp("Emp_Code").value)
        val2 = IIf(IsNull(Rs_Temp("NumEkama").value), "", Rs_Temp("NumEkama").value)
    End If
End If
txtCode.Text = val1
TxtId.Text = val2

End Sub

Private Sub Form_Load()

    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
        Dcombos.GetBranches Me.dcBranch
    Dim str As String
    str = "select id , BoardNO  from tblcarsdata "
    fill_combo dcCars, str

    'str = "   select   EmpID, Emp_Name,  Emp_Code,  DrivValue, Emp_mobile  from tblCarDrivers ,TblEmployee where tblCarDrivers.EmpID = TblEmployee.Emp_ID "
   ' str = "   select   EmpID, Emp_Name    from tblCarDrivers ,TblEmployee where tblCarDrivers.EmpID = TblEmployee.Emp_ID "
   
  str = "  select   e.Emp_ID Emp_ID , e.Emp_Name   Emp_Name  from TblEmployee e, TblEmpJobsTypes  j"
  str = str & "   Where e.JobTypeID = j.JobTypeID"
 str = str & "     and  ( j.JobTypeName like '%”«∆ﬁ%'  or j.JobTypeNamee like '%driver%')"
   str = str & "  AND  (BranchId=0 or BranchId is null or         BranchId in(" & Current_branchSql & "))"
   
    fill_combo DCDriver, str

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
       ChangeLang
    End If

    FromDate.value = Date
    ToDate.value = Date
    FromDateH.value = ToHijriDate(Date)
    todateH.value = ToHijriDate(Date)

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblDriverAllocation order by IDDA"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
   Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
'
End Sub

Private Sub ChangeLang()
   
 
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

Private Sub FromDate_Change()
    If Me.FromDate.value <> "" Then
            If Me.ToDate.value <> "" Then
            Me.txtDay.Text = DateDiff("d", Me.FromDate.value, Me.ToDate.value)
            End If
    End If
    
    FromDateH.value = ToHijriDate(FromDate.value)
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
           If Lngid <> 0 Then
        rs.find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
    If rs.RecordCount < 1 Then
        Exit Sub
    End If
    dcBranch.BoundText = IIf(IsNull(rs("branchid").value), "", rs("branchid").value)
    Me.Date.value = IIf(IsNull(rs("date").value), Date, rs("date").value)
    Me.DateH.value = IIf(IsNull(rs("date").value), ToHijriDate(Date), rs("date").value)
      
 
    Me.xtxtID.Text = IIf(IsNull(rs("IDDA").value), "", rs("IDDA").value)
    StrSQL = "   select * from TblDriverAllocation_Details  where idda =  " & val(xtxtID.Text)
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
            .Rows = .FixedRows + RsDev.RecordCount
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(RsDev("Emp_id").value), "", RsDev("Emp_id").value)
                .TextMatrix(i, .ColIndex("Emp_code")) = IIf(IsNull(RsDev("emp_code").value), "", RsDev("emp_code").value)
                .TextMatrix(i, .ColIndex("emp_name")) = IIf(IsNull(RsDev("emp_name").value), "", RsDev("emp_name").value)
                .TextMatrix(i, .ColIndex("car")) = IIf(IsNull(RsDev("Board").value), "", RsDev("Board").value)
                .TextMatrix(i, .ColIndex("carid")) = IIf(IsNull(RsDev("carid").value), "", RsDev("carid").value)
                .TextMatrix(i, .ColIndex("NumEkama")) = IIf(IsNull(RsDev("NumEkama").value), "", RsDev("NumEkama").value)
                .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(RsDev("FromDate").value), "", RsDev("FromDate").value)
                .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(RsDev("ToDate").value), "", RsDev("ToDate").value)
                .TextMatrix(i, .ColIndex("FromDateH")) = IIf(IsNull(RsDev("FromDateH").value), "", RsDev("FromDateH").value)
                .TextMatrix(i, .ColIndex("ToDateH")) = IIf(IsNull(RsDev("ToDateH").value), "", RsDev("ToDateH").value)
                .TextMatrix(i, .ColIndex("daycount")) = IIf(IsNull(RsDev("daycount").value), "", RsDev("daycount").value)
                RsDev.MoveNext
            Next i
        End With
    End If
     
     XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
 
    ' ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub Fromdateh_LostFocus()
 VBA.Calendar = vbCalGreg
        FromDate.value = ToGregorianDate(FromDateH.value)
End Sub



Private Sub ToDate_Change()
        If Me.FromDate.value <> "" Then
                If Me.ToDate.value <> "" Then
                Me.txtDay.Text = DateDiff("d", Me.FromDate.value, Me.ToDate.value)
                End If
        End If
        
        todateH.value = ToHijriDate(ToDate.value)
End Sub



Private Sub ToDateH_LostFocus()
 VBA.Calendar = vbCalGreg
        ToDate.value = ToGregorianDate(todateH.value)
        
       
        
End Sub

Private Sub txtcarcode_Change()
Dim val1, val2
If txtcarcode.Text = "" Then Exit Sub
Dim str As String
    str = " select  id , Fullcode , name  from TblCarsData   where Fullcode =   '" & txtcarcode.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        dcCars.BoundText = IIf(IsNull(Rs_Temp("ID").value), "", Rs_Temp("ID").value)
     Else
        dcCars.BoundText = ""
    End If
End Sub

Private Sub txtcarcode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
            Unload FrmCasrShearches
            FrmCasrShearches.SendForm = "DriverAllocation"
            FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub TxtCode_Change()
Dim val1, val2
If txtCode.Text = "" Then Exit Sub
Dim str As String
    str = " select * from TblEmployee  where Emp_Code =  '" & txtCode.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        val1 = IIf(IsNull(Rs_Temp("NumEkama").value), "", Rs_Temp("NumEkama").value)
        val2 = IIf(IsNull(Rs_Temp("Emp_ID").value), "", Rs_Temp("Emp_ID").value)
    End If
    TxtId.Text = val1
    DCDriver.BoundText = val(val2)
End Sub

Private Sub txtId_Change()
    Dim val1, val2
    If TxtId.Text = "" Then Exit Sub
    Dim i  As Integer, str As String


    str = " select * from TblEmployee  where NumEkama =  '" & TxtId.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        val1 = IIf(IsNull(Rs_Temp("Emp_Code").value), "", Rs_Temp("Emp_Code").value)
        val2 = IIf(IsNull(Rs_Temp("Emp_ID").value), "", Rs_Temp("Emp_ID").value)
    End If
    txtCode.Text = val1
    DCDriver.BoundText = val2
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
      '  CmdRemove.Enabled = True
      '  Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
'        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
        
        C1Elastic3.Enabled = True
        C1Elastic2.Enabled = True
        Grid.Enabled = True
        
    ElseIf Me.TxtModFlg.Text = "E" Then
        'CmdRemove.Enabled = True
        'Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        C1Elastic3.Enabled = True
        C1Elastic2.Enabled = True
        Grid.Enabled = True

    Else
        'Ele(1).Enabled = False

        'CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        C1Elastic3.Enabled = False
        C1Elastic2.Enabled = False
        Grid.Enabled = False
    End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

    On Error GoTo ErrTrap

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

Private Sub XPDtbTrans_Change()
 
End Sub




Function print_report(Optional NoteSerial As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

   
   
  MySQL = MySQL & "  SELECT dbo.TblVendorCars.BoardNo, dbo.TblDriverAllocation.DurationID, dbo.TblDriverAllocation.BranchID, dbo.TblBranchesData.branch_name, dbo.TblDurations.Name, "
   MySQL = MySQL & "                  dbo.TblDriverAllocation.DateH, dbo.TblDriverAllocation.Date, dbo.TblDriverAllocation_Details.CarID, dbo.TblDriverAllocation_Details.emp_id, dbo.TblEmployee.Fullcode,"
 MySQL = MySQL & "                    dbo.TblEmployee.NumEkama, dbo.TblDriverAllocation_Details.ToDate, dbo.TblDriverAllocation_Details.FromDateH, dbo.TblDriverAllocation_Details.ToDateH,"
    MySQL = MySQL & "                 dbo.TblDriverAllocation_Details.FromDate , dbo.TblDriverAllocation_Details.daycount"
 MySQL = MySQL & "  FROM     dbo.TblDriverAllocation INNER JOIN"
  MySQL = MySQL & "                   dbo.TblDriverAllocation_Details ON dbo.TblDriverAllocation.IDDA = dbo.TblDriverAllocation_Details.IDDA INNER JOIN"
     MySQL = MySQL & "                dbo.TblVendorCars ON dbo.TblDriverAllocation_Details.CarID = dbo.TblVendorCars.ID INNER JOIN"
      MySQL = MySQL & "               dbo.TblDurations ON dbo.TblDriverAllocation.DurationID = dbo.TblDurations.ID INNER JOIN"
      MySQL = MySQL & "               dbo.TblBranchesData ON dbo.TblDriverAllocation.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
       MySQL = MySQL & "              dbo.TblEmployee ON dbo.TblDriverAllocation_Details.emp_id = dbo.TblEmployee.Emp_ID"
      MySQL = MySQL & "  where  1 =1  "
   

     If Me.xtxtID.Text <> "" Then
            MySQL = MySQL & "   and  IDDA =  " & val(xtxtID.Text)
    End If
   
    
    
    
  
     If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_DriverAllocation.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_DriverAllocation.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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

