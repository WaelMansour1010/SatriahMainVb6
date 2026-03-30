VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSearshExpens40A 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»ÕÀ «·«÷«›«  ··«’Ê·"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   Icon            =   "FrmSearshExpens40A.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   5055
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   8775
      _cx             =   15478
      _cy             =   8916
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·› —Â"
         Height          =   1215
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2880
         Width           =   3255
         Begin MSComCtl2.DTPicker DtpDateFrom 
            Height          =   330
            Left            =   120
            TabIndex        =   21
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   103153667
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DtpDateTo 
            Height          =   330
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   103153667
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   195
            Index           =   2
            Left            =   2010
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰  «—ÌŒ"
            Height          =   195
            Index           =   4
            Left            =   2010
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   1005
         End
      End
      Begin VB.Frame lbprocess 
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”‰œ"
         Height          =   645
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2760
         Width           =   5355
         Begin VB.TextBox TxtIDFrom 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   915
         End
         Begin VB.TextBox TxtIDTO 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   5
            Left            =   4575
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   6
            Left            =   2460
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   525
         End
      End
      Begin VB.Frame lblLW 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»ÕÀ »Õ”»"
         Height          =   1575
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   3480
         Width           =   5505
         Begin VB.ComboBox CboType 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            ItemData        =   "FrmSearshExpens40A.frx":038A
            Left            =   120
            List            =   "FrmSearshExpens40A.frx":038C
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   3855
         End
         Begin MSDataListLib.DataCombo DcFixedAssets 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12640511
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSandAdd 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12640511
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«’· «·«”«”Ì"
            Height          =   285
            Index           =   0
            Left            =   4230
            TabIndex        =   13
            Top             =   735
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·”‰œ"
            Height          =   285
            Index           =   3
            Left            =   4230
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   270
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«’· «·„÷«›"
            Height          =   285
            Index           =   12
            Left            =   4065
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   1200
            Width           =   1290
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2745
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   8835
         _cx             =   15584
         _cy             =   4842
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
         FormatString    =   $"FrmSearshExpens40A.frx":038E
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
   Begin VB.ComboBox CboType1 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      ItemData        =   "FrmSearshExpens40A.frx":0485
      Left            =   0
      List            =   "FrmSearshExpens40A.frx":0487
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5760
      Width           =   3855
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   0
      Top             =   5160
      Width           =   765
      _ExtentX        =   1349
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
      Height          =   375
      Index           =   1
      Left            =   810
      TabIndex        =   1
      Top             =   5160
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
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
      Height          =   375
      Index           =   2
      Left            =   30
      TabIndex        =   2
      Top             =   5160
      Width           =   735
      _ExtentX        =   1296
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
   Begin C1SizerLibCtl.C1Elastic C1Elastic2 
      Height          =   5055
      Left            =   0
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   8775
      _cx             =   15478
      _cy             =   8916
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Height          =   1485
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   3480
         Width           =   8595
         Begin VB.TextBox TxtJobTypeNamee 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   600
            Width           =   3675
         End
         Begin VB.TextBox TxtJobTypeName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   3675
         End
         Begin VB.TextBox TxtVisaCode 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   2475
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ «‰Õ·Ì“Ì"
            Height          =   195
            Index           =   13
            Left            =   7095
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ ⁄—»Ì"
            Height          =   195
            Index           =   8
            Left            =   7095
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—„“ «·ÊŸÌ›…"
            Height          =   195
            Index           =   7
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·ÊŸÌ›…"
         Height          =   645
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2760
         Width           =   5355
         Begin VB.TextBox TxtToFromJobTypeID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   915
         End
         Begin VB.TextBox TxtFromJobTypeID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   11
            Left            =   2460
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   525
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   9
            Left            =   4575
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   540
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   2745
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   8835
         _cx             =   15584
         _cy             =   4842
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
         FormatString    =   $"FrmSearshExpens40A.frx":0489
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3060
      Width           =   1785
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2700
      Width           =   2295
   End
End
Attribute VB_Name = "FrmSearshExpens40A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public Indexx As Integer
Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
 If Indexx = 1 Then
 GetDataJob
 Else
 GetData
 End If
           
        Case 1
            clear_all Me
            DtpDateFrom.value = ""
DtpDateTo.value = ""
'Me.DtpDateFrom.value = ""
'Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub

















Private Sub Fg_Click()
FrmExpenses40A.Retrive Fg.TextMatrix(Fg.Row, Fg.ColIndex("id"))
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  Set Dcombos = New ClsDataCombos
  Dcombos.GetFixedAssets Me.DcFixedAssets
   Dcombos.GetFixedAssets Me.DcbSandAdd
 With Me.CboType
        .Clear
        .AddItem "œ„Ã «’·"
        .AddItem "««÷«›… ﬁÌ„Â ·√’·"
    
    End With
     With Me.CboType1
        .Clear
        .AddItem "œ„Ã «’·"
        .AddItem "««÷«›… ﬁÌ„Â ·√’·"
    
    End With
    
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

DtpDateFrom.value = ""
DtpDateTo.value = ""
C1Elastic2.Visible = False
C1Elastic1.Visible = False
If Indexx = 1 Then
C1Elastic2.Visible = True
FrmSearshExpens40A.Caption = "»ÕÀ «·ÊŸ«∆›"
Else
C1Elastic1.Visible = True
FrmSearshExpens40A.Caption = "»ÕÀ «·«÷«›«  ··«’Ê·"
End If
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = "SELECT     dbo.TblAdditionsAssest.ID, dbo.TblAdditionsAssest.RecordDate, dbo.TblAdditionsAssest.UserID, dbo.TblAdditionsAssest.BranchID, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAdditionsAssest.TypeSand, dbo.TblAdditionsAssest.SandAdd,"
StrSQL = StrSQL & "                       dbo.TblAdditionsAssest.FixedID, FixedAssets_1.Name, FixedAssets_1.namee, FixedAssets_2.Name AS FixAddName, FixedAssets_2.namee AS FixAddNameE,"
StrSQL = StrSQL & "                       dbo.TblAdditionsAssest.PurchasePrice, dbo.TblAdditionsAssest.PurchasePrice2, dbo.TblAdditionsAssest.AccDepre, dbo.TblAdditionsAssest.AccDepre2,"
StrSQL = StrSQL & "                       dbo.TblAdditionsAssest.CurrentValue, dbo.TblAdditionsAssest.CurrentValue2, dbo.TblAdditionsAssest.NuminstallmTotal, dbo.TblAdditionsAssest.NuminstallmTotal2,"
StrSQL = StrSQL & "                       dbo.TblAdditionsAssest.NuminstallmExcu, dbo.TblAdditionsAssest.NuminstallmExcu2, dbo.TblAdditionsAssest.NuminstallmRemin,"
StrSQL = StrSQL & "                       dbo.TblAdditionsAssest.NuminstallmRemin2, dbo.TblAdditionsAssest.NuminstallmCurr, dbo.TblAdditionsAssest.NuminstallmCurr2,"
StrSQL = StrSQL & "                       dbo.TblAdditionsAssest.general_des , dbo.TblAdditionsAssest.NoteSerial"
StrSQL = StrSQL & "  FROM         dbo.TblAdditionsAssest LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.FixedAssets FixedAssets_1 ON dbo.TblAdditionsAssest.SandAdd = FixedAssets_1.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.FixedAssets FixedAssets_2 ON dbo.TblAdditionsAssest.FixedID = FixedAssets_2.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblBranchesData ON dbo.TblAdditionsAssest.BranchID = dbo.TblBranchesData.branch_id"

    BolBegine = False
    StrWhere = ""

    '///////////////////
        If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblAdditionsAssest.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAdditionsAssest.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
  

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblAdditionsAssest.ID <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAdditionsAssest.ID <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
    

    
         If Me.CboType.Text <> "" And (val(CboType.ListIndex) <> -1) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblAdditionsAssest.TypeSand =" & Me.CboType.ListIndex & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAdditionsAssest.TypeSand =" & Me.CboType.ListIndex & ""
        End If
    End If
    
          If Me.DcbSandAdd.Text <> "" And (val(DcbSandAdd.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblAdditionsAssest.SandAdd =" & Me.DcbSandAdd.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAdditionsAssest.SandAdd =" & Me.DcbSandAdd.BoundText & ""
        End If
    End If
    
 
       If Me.DcFixedAssets.Text <> "" And (val(DcFixedAssets.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblAdditionsAssest.FixedID =" & Me.DcFixedAssets.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAdditionsAssest.FixedID =" & Me.DcFixedAssets.BoundText & ""
        End If
    End If
     If Not IsNull(Me.DtpDateFrom.value) Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblAdditionsAssest.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
          StrWhere = StrWhere & " where dbo.TblAdditionsAssest.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
                   
      End If
        If Not IsNull(Me.DtpDateTo.value) Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblAdditionsAssest.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
          StrWhere = StrWhere & " where dbo.TblAdditionsAssest.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
                   
      End If


    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblAdditionsAssest.id "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
               
                CboType1.ListIndex = IIf(IsNull(rs("TypeSand").value), "", rs("TypeSand").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
            .TextMatrix(i, .ColIndex("FixAddName")) = IIf(IsNull(rs("FixAddName").value), "", rs("FixAddName").value)
       
            Else
            .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
            .TextMatrix(i, .ColIndex("FixAddName")) = IIf(IsNull(rs("FixAddNameE").value), "", rs("FixAddNameE").value)
           End If
           .TextMatrix(i, .ColIndex("TypeSand")) = CboType1.Text

            
                
                
        
            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub
Public Sub GetDataJob()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = "SELECT     JobTypeID, JobTypeName, JobTypeNamee, VisaCode"
StrSQL = StrSQL & " From dbo.TblEmpJobsTypes where 1=1"
    StrWhere = ""

    '///////////////////
     If val(Me.TxtFromJobTypeID.Text) <> 0 Then
            StrWhere = StrWhere & " and JobTypeID >=" & val(Me.TxtFromJobTypeID.Text) & ""
    End If
      If val(Me.TxtToFromJobTypeID.Text) <> 0 Then
            StrWhere = StrWhere & " and JobTypeID <=" & val(Me.TxtToFromJobTypeID.Text) & ""
    End If

    If TxtJobTypeName.Text <> "" Then
            StrWhere = StrWhere & " and JobTypeName Like N'%" & TxtJobTypeName.Text & "%'"
    End If
    If TxtJobTypeNamee.Text <> "" Then
            StrWhere = StrWhere & " and JobTypeNamee Like N'%" & TxtJobTypeNamee.Text & "%'"
    End If
     If TxtVisaCode.Text <> "" Then
            StrWhere = StrWhere & " and VisaCode Like N'%" & TxtVisaCode.Text & "%'"
    End If

    '-----------------------------------
    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By JobTypeID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.VSFlexGrid1
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("JobTypeID")) = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                .TextMatrix(i, .ColIndex("JobTypeNamee")) = IIf(IsNull(rs("JobTypeNamee").value), "", rs("JobTypeNamee").value)
                .TextMatrix(i, .ColIndex("VisaCode")) = IIf(IsNull(rs("VisaCode").value), "", rs("VisaCode").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub

Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
  If Indexx = 1 Then
   Me.Caption = "Saerch Data of Jobs"
  Else
  Me.Caption = "Saerch Added  Of Fixed Assets"
End If
lbprocess.Caption = "No Transection"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lblLW.Caption = "Saerch By"
lbl(3).Caption = "Type"
lbl(0).Caption = "Basic Assest"
lbl(12).Caption = "Equipment"
Frame1.Caption = "Added Assest"
lbl(4).Caption = "From"
lbl(2).Caption = "To"
lbl(9).Caption = "From"
lbl(11).Caption = "To"
Frame3.Caption = "Code"
lbl(7).Caption = "Visa Code"
lbl(8).Caption = "Name Arabic"
lbl(13).Caption = "Name English"
'lbl(2).Caption = "Total"
     With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("JobTypeID")) = "Code"
        .TextMatrix(0, .ColIndex("JobTypeName")) = "Name Arabic"
        .TextMatrix(0, .ColIndex("JobTypeNamee")) = "Name English"
        .TextMatrix(0, .ColIndex("VisaCode")) = "VisaCode"
    End With
 With Me.CboType
        .Clear
    
        .Clear
        .AddItem "Assets Merge"
        .AddItem "Assets Additions"
    
    End With
     With Me.CboType1
        .Clear
    
        .Clear
        .AddItem "Assets Merge"
        .AddItem "Assets Additions"
    
    End With
    
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "No Transection"
        .TextMatrix(0, .ColIndex("RecordDate")) = "RecordDate"
         .TextMatrix(0, .ColIndex("TypeSand")) = "Type"
        .TextMatrix(0, .ColIndex("FixAddName")) = "Basic Assest"
       .TextMatrix(0, .ColIndex("Name")) = "Added Assest"
    End With
  '
End Sub

Private Sub VSFlexGrid1_Click()
    On Error GoTo ErrTrap
   FrmEmpJobsTypes.FindRec val(Me.VSFlexGrid1.TextMatrix(Me.VSFlexGrid1.Row, Me.VSFlexGrid1.ColIndex("JobTypeID")))
ErrTrap:
End Sub
