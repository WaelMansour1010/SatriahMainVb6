VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmGoldDetaiks 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ”ÃÌ·  ›«’Ì· «·ﬁÿ⁄Â"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13200
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   255
      Left            =   11400
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   44
      Text            =   "frmGoldDetails.frx":0000
      Top             =   6480
      Width           =   7815
   End
   Begin VB.TextBox txtWages 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "»Ì«‰«  «·÷„«‰"
      Height          =   1455
      Left            =   14160
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   960
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox txtvlaue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Text            =   "12"
         Top             =   720
         Width           =   855
      End
      Begin MSComCtl2.DTPicker GranteeStartDate 
         Height          =   330
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Format          =   96468993
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker GranteeEndDate 
         Height          =   330
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Format          =   96468993
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Â—"
         Height          =   255
         Index           =   12
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "› —… «·÷„«‰"
         Height          =   255
         Index           =   11
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ ‰Â«Ì… «·÷„«‰"
         Height          =   255
         Index           =   9
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ »œ«Ì… «·÷„«‰"
         Height          =   255
         Index           =   6
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   360
         Width           =   1995
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   2655
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   -6240
      Width           =   4575
      Begin VB.TextBox TxtNoOFVisits 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "12"
         Top             =   720
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTRegMaintDate 
         Height          =   330
         Left            =   840
         TabIndex        =   19
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Format          =   96468993
         CurrentDate     =   38784
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   20
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   688
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "≈÷«›…"
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
         ButtonImage     =   "frmGoldDetails.frx":0006
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   975
         Index           =   3
         Left            =   1320
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1140
         Width           =   3090
         _cx             =   5450
         _cy             =   1720
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
         Begin VB.OptionButton OptInt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌÊ„"
            Height          =   210
            Index           =   0
            Left            =   2415
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   345
            Width           =   630
         End
         Begin VB.OptionButton OptInt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘Â—"
            Height          =   225
            Index           =   1
            Left            =   1650
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   345
            Value           =   -1  'True
            Width           =   720
         End
         Begin VB.TextBox Txt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   7
            Left            =   30
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Text            =   "1"
            Top             =   585
            Width           =   915
         End
         Begin VB.OptionButton OptInt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰…"
            Height          =   225
            Index           =   2
            Left            =   990
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   345
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œ… «·› —…"
            Height          =   195
            Index           =   17
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   345
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —… »Ì‰ «·“Ì«—« "
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   18
            Left            =   1050
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   0
            Width           =   1980
         End
      End
      Begin MSDataListLib.DataCombo DCVisits 
         Height          =   315
         Left            =   960
         TabIndex        =   31
         Top             =   2160
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·“Ì«—…"
         Height          =   255
         Index           =   15
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2160
         Width           =   795
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·’Ì«‰… Ì»œ√ „‰"
         Height          =   375
         Index           =   13
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·“Ì«—« "
         Height          =   375
         Index           =   14
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   840
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   13095
      Begin VB.OptionButton GranteeTypeopt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»÷„«‰"
         Height          =   195
         Index           =   1
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton GranteeTypeopt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»œÊ‰ ÷„«‰"
         Height          =   195
         Index           =   0
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”ÿ—: "
         Height          =   255
         Index           =   0
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   3
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·’‰›: "
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   11640
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   4
         Left            =   10320
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·’‰›:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   5
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   5205
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   7
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·÷„«‰"
         Height          =   255
         Index           =   10
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   1155
      End
   End
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   405
      Left            =   1020
      TabIndex        =   1
      Top             =   3810
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   714
      Caption         =   "Õ›Ÿ"
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
   Begin VB.TextBox TxtComment 
      Alignment       =   1  'Right Justify
      Height          =   975
      Left            =   30
      MaxLength       =   255
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   10110
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   2
      Top             =   3810
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   714
      Caption         =   "«·€«¡"
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
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5235
      Left            =   15120
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   9735
      _cx             =   17171
      _cy             =   9234
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
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmGoldDetails.frx":03A0
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   390
      Index           =   21
      Left            =   11640
      TabIndex        =   5
      Top             =   3480
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " Õ–› ”ÿ—"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "frmGoldDetails.frx":0663
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   390
      Index           =   0
      Left            =   9960
      TabIndex        =   32
      Top             =   3480
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " Õ–› «·ﬂ·"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "frmGoldDetails.frx":0BFD
      DrawFocusRectangle=   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid fgCameo 
      Height          =   2820
      Left            =   0
      TabIndex        =   41
      Top             =   600
      Width           =   13155
      _cx             =   23204
      _cy             =   4974
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
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmGoldDetails.frx":1197
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin VB.Label lblTotals 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«Ã„«·Ì «·ﬁÿ⁄Â"
      Height          =   255
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«Ã„«·Ì «ÃÊ— «·’Ì«€…"
      Height          =   255
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label LBLInsWages 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«ÃÊ— «· —ﬂÌ»"
      Height          =   255
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lBLnOoFsTNES 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«Ã„«·Ì ⁄œœ «·«ÕÃ«—"
      Height          =   255
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ ‰Â«Ì… «·÷„«‰"
      Height          =   255
      Index           =   16
      Left            =   15120
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   8
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "FrmGoldDetaiks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Fg As VSFlex8UCtl.vsFlexGrid

Public LngRow As Long

Public LngCol As Long

Public AllDate As String
Public AllIDS As String
Public Allline As String

Private Sub Cmd_Click(Index As Integer)

    Select Case Index
Case 0
     cleargrid

        Case 20
          '  addrow
calcrows
        Case 21
            RemoveGridRow
    End Select

End Sub
Function cleargrid()
    With Me.fgCameo
   '     .Clear flexClearScrollable, flexClearEverything
           .Clear flexClearScrollable

        .Rows = 1
     End With
End Function
Function calcrows()
   End Function
Public Sub FillGridWithData()

    On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim strFilterText1 As String
      Dim Unitname As String
    Dim ttypename As String
     Dim typename As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim inty As Integer
    Dim intervalstr As String
Dim name As String
Dim NameE As String
Dim Remarks As String
 
    
     Dim astrSplitItems1() As String
     If AllIDS = "" Then
     Exit Sub
     End If
     
    strFilterText = "&&"
         strFilterText1 = "@@"
    astrSplitItems = Split(Me.AllIDS, strFilterText)
    fgCameo.Rows = UBound(astrSplitItems) + 1
 
    For intX = 0 To UBound(astrSplitItems) - 1
    
    astrSplitItems1 = Split(astrSplitItems(intX), strFilterText1)
   
                fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("TTypeId")) = astrSplitItems1(0)
                fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("typeid")) = astrSplitItems1(1)
                fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("uniteid")) = astrSplitItems1(2)
'                    fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("uniteid")) = astrSplitItems1(2)
        '    fgCameo.Cell(flexcpData, intX + 1, fgCameo.ColIndex("uniteid")) = fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("uniteid"))
        
        fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("type")) = astrSplitItems1(3)
                fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("price")) = astrSplitItems1(4)
                fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("weight")) = astrSplitItems1(5)
                fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("Count")) = astrSplitItems1(6)
                fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("InstallPrice")) = astrSplitItems1(7)
   
  
     
  ' ttypename As String, Optional ByRef typename
     GetGoldData val(fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("TTypeId"))), val(fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("type"))), val(fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("uniteid"))), Unitname, ttypename, typename
  
  fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("unite")) = Unitname
  
  fgCameo.TextMatrix(intX + 1, fgCameo.ColIndex("TType")) = ttypename
  
 
    Next
     
     
     fgCameo.Rows = fgCameo.Rows + 1
     
      ReLineGrid
     
     
ErrTrap:
End Sub

Private Sub RemoveGridRow()

    With Me.fgCameo

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    Dim Msg As String
    Dim ExpiryDate As Date
    Dim Askinterval As String
ReLineGrid
    If Not Fg Is Nothing Then
 
  
    '    If Me.Fg.ColIndex("GranteeType") <> -1 Then
 
    '        Fg.TextMatrix(LngRow, Fg.ColIndex("GranteeType")) = IIf(GranteeTypeopt(0).value = True, 0, 1)
    '    End If

    '    If Me.Fg.ColIndex("GranteeStartDate") <> -1 Then
 '
 '           Fg.TextMatrix(LngRow, Fg.ColIndex("GranteeStartDate")) = GranteeStartDate.value
 '       End If

       ' If Me.Fg.ColIndex("GranteeEndDate") <> -1 Then
 
       '     Fg.TextMatrix(LngRow, Fg.ColIndex("GranteeEndDate")) = GranteeEndDate.value
       ' End If

       '  If Me.Fg.ColIndex("guaranteeTime") <> -1 Then
 
       '     Fg.TextMatrix(LngRow, Fg.ColIndex("guaranteeTime")) = val(txtvlaue.text)
  
       ' End If

        If Me.Fg.ColIndex("GoldDetails") <> -1 Then
 
            Fg.TextMatrix(LngRow, Fg.ColIndex("GoldDetails")) = AllIDS
        End If


        If Me.Fg.ColIndex("Wages") <> -1 Then
 
            Fg.TextMatrix(LngRow, Fg.ColIndex("Wages")) = val(txtWages.text)
  
        End If


        If Me.Fg.ColIndex("Price") <> -1 Then
             If val(Fg.TextMatrix(LngRow, Fg.ColIndex("Price"))) = 0 Then
                  Fg.TextMatrix(LngRow, Fg.ColIndex("Price")) = val(lblTotals.Caption)
            End If
  
        End If
        
        Unload Me
    End If

End Sub

Sub addrow()
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String
 
    Me.Grid.Rows = Me.Grid.Rows + 1
    LngRow = Me.Grid.Rows - 1

    With Me.Grid
 
        .TextMatrix(LngRow, .ColIndex("MaDate")) = (DTRegMaintDate.value)
        .AutoSize 0, .Cols - 1, False
    End With
  
    ReLineGrid
 
End Sub



Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
   
AllIDS = ""

 Dim total As Double

total = 0
LBLInsWages = 0
lBLnOoFsTNES = 0
    With Me.fgCameo

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("TTypeid")) <> "" Then
            Allline = ""
                IntCounter = IntCounter + 1
               .TextMatrix(i, .ColIndex("Ser")) = IntCounter
        '        AllDate = AllDate & .TextMatrix(i, .ColIndex("MaDate")) & ","
       Allline = .TextMatrix(i, .ColIndex("TTypeId")) & "@@" & .TextMatrix(i, .ColIndex("typeid")) & "@@"
        
     Allline = Allline & .TextMatrix(i, .ColIndex("uniteid")) & "@@"
        Allline = Allline & (.TextMatrix(i, .ColIndex("type"))) & "@@"
       Allline = Allline & val(.TextMatrix(i, .ColIndex("price"))) & "@@"
       Allline = Allline & val(.TextMatrix(i, .ColIndex("weight"))) & "@@"
       Allline = Allline & val(.TextMatrix(i, .ColIndex("Count"))) & "@@"
       Allline = Allline & val(.TextMatrix(i, .ColIndex("InstallPrice"))) & "@@"
       
         AllIDS = AllIDS & Allline & "&&"
         
         .TextMatrix(i, .ColIndex("Total")) = (val(.TextMatrix(i, .ColIndex("price"))) * val(.TextMatrix(i, .ColIndex("weight")))) + (val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("InstallPrice"))))
         lBLnOoFsTNES = lBLnOoFsTNES + val(.TextMatrix(i, .ColIndex("Count")))
         LBLInsWages = LBLInsWages + (val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("InstallPrice"))))
         total = total + val(.TextMatrix(i, .ColIndex("Total")))
         
            End If

        Next i
 lblTotals.Caption = total + val(txtWages)
 
    End With
    
Text1.text = AllIDS
End Sub

 
Private Sub Command1_Click()
ReLineGrid
End Sub

Private Sub Command2_Click()
FillGridWithData
End Sub

Private Sub fgCameo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim StrAccountCode1 As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim k As Integer
Dim StrComboList As String
            
    
    With fgCameo

        Select Case .ColKey(Col)
     Case "TType"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("TTypeId"), False, True)
                .TextMatrix(Row, .ColIndex("TTypeId")) = StrAccountCode


                 .TextMatrix(Row, .ColIndex("type")) = ""

                     
     Case "type"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("typeId"), False, True)
                .TextMatrix(Row, .ColIndex("typeId")) = StrAccountCode

     Case "unite"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("uniteid"), False, True)
                .TextMatrix(Row, .ColIndex("uniteid")) = StrAccountCode


   

                   End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

       
    End With

    ReLineGrid
End Sub

Private Sub fgCameo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With fgCameo
 
             If .ColKey(Col) <> "TType" And .ColKey(Col) <> "type" And .ColKey(Col) <> "unite" Then
       
            .ComboList = ""
        End If
 
        
    End With
     
End Sub

Private Sub fgCameo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmItemCameoSearch
            FrmItemCameoSearch.show

'
End If
End Sub

Private Sub fgCameo_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With fgCameo

        Select Case .ColKey(Col)
'TType
 Case "TType"
     StrSQL = " select  id,name,nameE from TblGType "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "name", "Id")
                Else
                    StrComboList = .BuildComboList(rs, "nameE", "Id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
      '          LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("TTypeId"), False, True)
       
 Case "type"
     
  If .TextMatrix(.Row, .ColIndex("TTypeId")) <> "" Then
    If val(.TextMatrix(.Row, .ColIndex("TTypeId"))) = 1 Then ' –Â»
    
       StrSQL = " select  id,name from TblAveragGrm "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "name", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("typeid"), False, True)
    
    
    ElseIf val(.TextMatrix(.Row, .ColIndex("TTypeId"))) = 2 Then   '«
        StrSQL = " select id,name,nameE,code from TblDiamonds "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "code", "id")
                Else
                    StrComboList = .BuildComboList(rs, "code", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("typeid"), False, True)
       
       
     ElseIf val(.TextMatrix(.Row, .ColIndex("TTypeId"))) = 3 Then '«ÕÃ«—     ﬂ—Ì„…
         StrSQL = " select id,name,nameE,code  from TblGemstones "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
             If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "code", "id")
                Else
                    StrComboList = .BuildComboList(rs, "code", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("typeid"), False, True)
                
 ElseIf val(.TextMatrix(.Row, .ColIndex("TTypeId"))) = 4 Then '«ÕÃ«—  ‰’›   ﬂ—Ì„…
         StrSQL = " select id,name,nameE,code from TblSemiGemstones "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
             If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "code", "id")
                Else
                    StrComboList = .BuildComboList(rs, "code", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("typeid"), False, True)
     
   ElseIf val(.TextMatrix(.Row, .ColIndex("TTypeId"))) = 5 Then '„ﬂÊ‰«  «Œ—Ì
         StrSQL = " select id,name,nameE from TblOtherComp "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
             If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "nameE", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("typeid"), False, True)
                
     End If
  

Else
            If SystemOptions.UserInterface = ArabicInterface Then
             MsgBox " Õœœ «·‰Ê⁄ «Ê·«"
            Else
            MsgBox " Select Type Firstly"
            End If
End If

 
 Case "unite"
     StrSQL = " select UnitID,UnitName,UnitNamee from TblUnites "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "UnitName", "UnitID")
                Else
                    StrComboList = .BuildComboList(rs, "UnitNamee", "UnitID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                'LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
         
 
        End Select

    End With
End Sub

Private Sub Form_Load()
    CenterForm Me

    FormPostion Me, GetPostion

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        cahngelang
    End If

    Me.CmdOk.ButtonStyle = impActive
    Set CmdOk.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    CmdOk.ButtonPositionImage = impRightOfText

    Me.CMDCancel.ButtonStyle = impActive
    Set CMDCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    CMDCancel.ButtonPositionImage = impRightOfText
    'GranteeStartDate.value = Date
    GranteeEndDate.value = Date
    DTRegMaintDate.value = Date
    
    
     
 Dim Dcombos As New ClsDataCombos
    Dcombos.GetmaintennceType Me.DCVisits




 





End Sub

Function cahngelang()
    Me.Caption = "Guarantee Data"

    lbl(1).Caption = "ItemCode"
    lbl(2).Caption = "Item Name"
    lbl(10).Caption = "G. Type"
    GranteeTypeopt(0).Caption = "WithOut Part"
    GranteeTypeopt(1).Caption = "With Part"
    lbl(6).Caption = "Guarantee  Start Date"
    lbl(9).Caption = "Guarantee  Emd Date"
    lbl(11).Caption = "Guarantee Period"
    lbl(12).Caption = "Month"
    lbl(13).Caption = "preventive maintenance Dates"
    Cmd(20).Caption = "ADD"
    Cmd(21).Caption = "Delete Row"
    CmdOk.Caption = "Save"
    CMDCancel.Caption = "Cancel"

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("MaDate")) = "preventive maintenance Dates"
    End With

End Function

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub
 Public Function getMaintentTypesData(StrAccountCode As String, Optional ByRef name As String _
 , Optional ByRef NameE As String, Optional ByRef Remarks As String, Optional ByRef intervalstr As String)
  
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim intervaltype As Integer
  Dim interval As Double
               StrSQL = " select * from TblMaintenanceType   where id=" & StrAccountCode
                Set rs = Nothing
              rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 

                If Not (rs.BOF Or rs.EOF) Then
                   name = IIf(IsNull(rs("name").value), "", rs("name").value)
                   NameE = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                   
                    Remarks = IIf(IsNull(rs("REMARKS").value), "", rs("REMARKS").value)
                    intervaltype = IIf(IsNull(rs("intervaltype").value), 0, rs("intervaltype").value)
                    interval = IIf(IsNull(rs("interval").value), 0, rs("interval").value)
                    
                    intervalstr = ""
                      If SystemOptions.UserInterface = ArabicInterface Then
                            If intervaltype = 0 Then
                            intervalstr = "œﬁÌﬁ…"
                            ElseIf intervaltype = 1 Then
                            intervalstr = "”«⁄Â"
                             ElseIf intervaltype = 2 Then
                            intervalstr = "ÌÊ„"
                             ElseIf intervaltype = 3 Then
                            intervalstr = "«”»Ê⁄"
                             ElseIf intervaltype = 4 Then
                            intervalstr = "‘Â—"
                            ElseIf intervaltype = 5 Then
                            intervalstr = "”‰…"
                            Else
                            intervalstr = ""
                            End If
                    Else
                    
                            If intervaltype = 0 Then
                            intervalstr = "Minute"
                            ElseIf intervaltype = 1 Then
                            intervalstr = "hour"
                            ElseIf intervaltype = 2 Then
                            intervalstr = "day"
                            ElseIf intervaltype = 3 Then
                            intervalstr = "week"
                            ElseIf intervaltype = 4 Then
                            intervalstr = "Month"
                            ElseIf intervaltype = 5 Then
                            intervalstr = "Year"
                            Else
                            intervalstr = ""
                            End If
                            
                    End If



                  intervalstr = interval & "   " & intervalstr
                 End If

 End Function
Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrTrap
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Grid

        Select Case .ColKey(Col)
 
            Case "MainName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MainID"), False, True)
                .TextMatrix(Row, .ColIndex("MainID")) = StrAccountCode
             
    
 
Dim intervalstr As String
Dim name As String
Dim NameE As String
Dim Remarks As String
 
 
  getMaintentTypesData StrAccountCode, name, name, Remarks, intervalstr
                    .TextMatrix(Row, .ColIndex("REMARKS")) = Remarks
 
       
                    .TextMatrix(Row, .ColIndex("Interval")) = intervalstr
 
 

           
          
        End Select
     
 

      
    End With
 
ErrTrap:

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
                                  

    With Grid

 





        Select Case .ColKey(Col)

            Case "MainID"
                .ComboList = ""

            Case "interval"
                .ComboList = ""
        
            Case "REMARKS"
                .ComboList = ""
                  Case "MaDate"
                .ComboList = ""
                 Case "Ser"
                .ComboList = ""
                 Cancel = True
            
        End Select

    End With
End Sub

 
Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    With Grid

        Select Case .ColKey(Col)

            Case "MainName"
                'Full Path Display
                 
                StrSQL = " select * from TblMaintenanceType "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "*name", "id")
                Else
                    StrComboList = Grid.BuildComboList(rs, "*name", "id")
                End If
                
          
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Public Sub txtvlaue_Change()
    Me.GranteeEndDate.value = DateAdd("M", val(Me.txtvlaue), Me.GranteeStartDate.value)
End Sub

Private Sub txtvlaue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, txtvlaue.text, 0)
End Sub

Private Sub txtWages_Change()
ReLineGrid
End Sub
