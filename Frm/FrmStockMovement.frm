VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmStockMovement 
   Caption         =   " ﬁ—Ì— »Õ—ﬂ… «·„Œ“Ê‰"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   Icon            =   "FrmStockMovement.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   10260
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   7425
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10260
      _cx             =   18098
      _cy             =   13097
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
      Align           =   5
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
      GridCols        =   5
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmStockMovement.frx":058A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   450
         Index           =   3
         Left            =   30
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1305
         Width           =   10200
         _cx             =   17992
         _cy             =   794
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
         Begin VB.CommandButton Cmd 
            Caption         =   "«’‰«› ·„  ŸÂ—"
            Height          =   375
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   30
            Visible         =   0   'False
            Width           =   1245
         End
         Begin MSDataListLib.DataCombo DcboItems 
            Height          =   315
            Left            =   1920
            TabIndex        =   28
            Top             =   60
            Visible         =   0   'False
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboGroups 
            Height          =   315
            Left            =   6240
            TabIndex        =   26
            Top             =   45
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ÕœÌœ ’‰›"
            Height          =   270
            Index           =   11
            Left            =   5310
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   75
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ÕœÌœ „Ã„Ê⁄…"
            Height          =   270
            Index           =   10
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   75
            Width           =   1125
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1260
         Index           =   2
         Left            =   2085
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   5610
         _cx             =   9895
         _cy             =   2223
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
         Caption         =   "⁄Ê«„· «·»ÕÀ"
         Align           =   0
         AutoSizeChildren=   0
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
            Height          =   735
            Index           =   4
            Left            =   60
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   450
            Width           =   1755
            _cx             =   3096
            _cy             =   1296
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ForeColor       =   128
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "‰Ÿ«„ «·⁄—÷"
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
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ ÃœÊ·Ì"
               Height          =   195
               Index           =   1
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   450
               Width           =   1455
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ ‘Ã—Ï"
               Height          =   195
               Index           =   0
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   210
               Value           =   -1  'True
               Width           =   1455
            End
         End
         Begin VB.ComboBox CboType 
            Height          =   315
            Left            =   1830
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   210
            Width           =   2805
         End
         Begin MSDataListLib.DataCombo DcboStores 
            Height          =   315
            Left            =   1830
            TabIndex        =   19
            Top             =   870
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.CheckBox Chk 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄—÷ Ã„Ì⁄ «·√’‰«›"
            Height          =   225
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   180
            Width           =   1725
         End
         Begin VB.ComboBox CboCostType 
            Height          =   315
            Left            =   1830
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   540
            Width           =   2805
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„·Ì…"
            Height          =   255
            Index           =   12
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   255
            Index           =   7
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   900
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «· ﬂ·›…"
            Height          =   255
            Index           =   6
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   570
            Width           =   975
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1260
         Index           =   1
         Left            =   7710
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   2520
         _cx             =   4445
         _cy             =   2223
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
         Caption         =   " ÕœÌœ  «—ÌŒ «·»ÕÀ "
         Align           =   0
         AutoSizeChildren=   0
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
         Begin MSComCtl2.DTPicker DtpFrom 
            Height          =   345
            Left            =   360
            TabIndex        =   4
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   106233857
            CurrentDate     =   39209
         End
         Begin MSComCtl2.DTPicker DtpTO 
            Height          =   345
            Left            =   360
            TabIndex        =   15
            Top             =   630
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   106233857
            CurrentDate     =   39209
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   255
            Index           =   5
            Left            =   1860
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   690
            Width           =   525
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   255
            Index           =   4
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   300
            Width           =   435
         End
      End
      Begin MSComctlLib.ProgressBar PrgBar 
         Height          =   300
         Left            =   30
         TabIndex        =   5
         Top             =   6540
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin ImpulseButton.ISButton CmdPrint 
         Height          =   645
         Left            =   30
         TabIndex        =   6
         Top             =   645
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1138
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
         ButtonImage     =   "FrmStockMovement.frx":062C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   4755
         Left            =   30
         TabIndex        =   7
         Top             =   1770
         Width           =   10200
         _cx             =   17992
         _cy             =   8387
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmStockMovement.frx":09C6
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   0
         Left            =   30
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   6855
         Width           =   10200
         _cx             =   17992
         _cy             =   953
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
         Begin ImpulseButton.ISButton CmdExit 
            Height          =   360
            Left            =   0
            TabIndex        =   9
            Top             =   90
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   635
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
            BackStyle       =   0
            ButtonImage     =   "FrmStockMovement.frx":0BBD
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   9
            Left            =   5490
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   90
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï ﬁÌ„… «·„Œ“Ê‰ »œ«Ì… «·› —…"
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   8
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   90
            Width           =   1620
         End
         Begin VB.Image Img 
            Height          =   240
            Left            =   1620
            Picture         =   "FrmStockMovement.frx":0F57
            Top             =   150
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·√’‰«›:-"
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   0
            Left            =   9105
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   90
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   3
            Left            =   8430
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   90
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï ﬁÌ„… «·„Œ“Ê‰ ‰Â«Ì… «·› —…"
            ForeColor       =   &H00000080&
            Height          =   405
            Index           =   1
            Left            =   3690
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   60
            Width           =   1620
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   2
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   90
            Width           =   1035
         End
      End
      Begin ImpulseButton.ISButton CmdDo 
         Height          =   600
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1058
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ‰›Ì–"
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
         ButtonImage     =   "FrmStockMovement.frx":12E1
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
   End
End
Attribute VB_Name = "FrmStockMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BolInProgress As Boolean
Dim IntCostItemPriceType As StockCostType
Dim cDboSearch(1) As clsDCboSearch

Private Sub CboCostType_Change()

    If Me.CboCostType.ListIndex = 0 Then
        IntCostItemPriceType = LastPurPriceType
        Me.Fg.TextMatrix(0, Fg.ColIndex("ItemCostPrice")) = "”⁄— «· ﬂ·›…(«Œ— ”⁄— ‘—«¡)"
    ElseIf Me.CboCostType.ListIndex = 1 Then
        IntCostItemPriceType = WeightAverage
        Me.Fg.TextMatrix(0, Fg.ColIndex("ItemCostPrice")) = "”⁄— «· ﬂ·›…(„ Ê”ÿ «·”⁄—)"
    ElseIf Me.CboCostType.ListIndex = 2 Then
        IntCostItemPriceType = FirstInFirstOut
        Me.Fg.TextMatrix(0, Fg.ColIndex("ItemCostPrice")) = "”⁄— «· ﬂ·›…(«·Ê«—œ √Ê·« Ì’—› √Ê·«)"
    ElseIf Me.CboCostType.ListIndex = 3 Then
        IntCostItemPriceType = ModernWeightAverage
        Me.Fg.TextMatrix(0, Fg.ColIndex("ItemCostPrice")) = "”⁄— «· ﬂ·›…(„ Ê”ÿ «· ﬂ·›… «·ÃœÌœ)"
    End If

    NewOperation
    Fg.AutoSize 0, Fg.Cols - 1, False
End Sub

Private Sub CboCostType_Click()
    CboCostType_Change
End Sub

Private Sub CboType_Change()

    If Me.CboType.ListIndex = 0 Then
        Me.lbl(4).Visible = False
        Me.DtpFrom.Visible = False
        FgStockCountSetup
    ElseIf Me.CboType.ListIndex = 1 Then
        Me.lbl(4).Visible = True
        Me.DtpFrom.Visible = True
        FgStockMovmentSetup
    End If

End Sub

Private Sub CboType_Click()
    CboType_Change
End Sub

Private Sub Cmd_Click()
    Dim StrSQL As String
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim StrTemp As String
    'StrTemp = ""
    'For I = 0 To Me.Fg.Rows - 1
    '    StrTemp = StrTemp & Trim$(Fg.TextMatrix(I, Fg.ColIndex("ItemID")))
    '    StrTemp = StrTemp & ","
    'Next I
    'If Trim$(StrTemp) = "" Then
    '    Exit Sub
    'Else
    '    StrTemp = Mid$(StrTemp, 1, Len(StrTemp) - 1)
    'End If
    'StrSQL = "Select * From TblItems Where ItemID NOT IN (" & StrTemp & ")"
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

End Sub

Private Sub CmdDo_Click()
    Dim Msg As String

    If Me.CboType.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·⁄„·Ì…...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If Me.DcboStores.BoundText = "" Then
        Msg = "ÌÃ» ≈Œ Ì«— «·„Œ“‰..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcboStores.SetFocus
        Exit Sub
    End If

    If Me.CboCostType.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… Õ”«»  ﬂ·›… «·„Œ“Ê‰....!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    ElseIf Me.CboCostType.ListIndex = 0 Then
        IntCostItemPriceType = LastPurPriceType
    ElseIf Me.CboCostType.ListIndex = 1 Then
        IntCostItemPriceType = WeightAverage
    ElseIf Me.CboCostType.ListIndex = 2 Then
        IntCostItemPriceType = FirstInFirstOut
    ElseIf Me.CboCostType.ListIndex = 3 Then
        IntCostItemPriceType = ModernWeightAverage
    End If

    If Me.CboType.ListIndex = 0 Then

        DoStockCount
    ElseIf Me.CboType.ListIndex = 1 Then

        DoStockmovement
    End If

End Sub

Private Sub CmdPrint_Click()

    If Me.CboType.ListIndex = 0 Then
        StockCountPrint
    ElseIf Me.CboType.ListIndex = 1 Then
        StockMovementPrint
    End If

End Sub

Private Sub Fg_BeforeSort(ByVal Col As Long, _
                          Order As Integer)

    If BolInProgress = True Then
        Order = 0
    End If

End Sub

Private Sub Fg_DblClick()
'    Dim LngItemID As Long
' '   Dim Frm As FrmShowItemCostPrice
'
    With Me.Fg

        If .Col = -1 Then Exit Sub
        If .Row = -1 Then Exit Sub
        If Me.CboType.ListIndex = 0 Then Exit Sub
        LngItemID = val(.TextMatrix(.Row, .ColIndex("ItemID")))

        If val(.TextMatrix(.Row, .ColIndex("ItemID"))) <> 0 Then
            If .Col = .ColIndex("ItemID") Or .Col = .ColIndex("ItemCode") Or .Col = .ColIndex("ItemName") Then
                Load FrmSelectData
                FrmSelectData.DcboItemName.BoundText = val(.TextMatrix(.Row, .ColIndex("ItemID")))
                FrmSelectData.TxtItemCode.text = val(.TextMatrix(.Row, .ColIndex("ItemCode")))
                FrmSelectData.DcboStores.BoundText = Me.DcboStores.BoundText
                FrmSelectData.DtpFrom.value = Me.DtpFrom.value
                FrmSelectData.DtpTO.value = Me.DtpTO.value
                FrmSelectData.show
            ElseIf .Col = .ColIndex("EndStock") Then
                Load FrmSearchSerial
                FrmSearchSerial.XPTxtCode.text = val(.TextMatrix(.Row, .ColIndex("ItemCode")))
                FrmSearchSerial.DCboItemsName.BoundText = val(.TextMatrix(.Row, .ColIndex("ItemID")))
                FrmSearchSerial.show
            ElseIf .Col = .ColIndex("ItemCostPrice") Or .Col = .ColIndex("StockCost") Then
                Me.MousePointer = vbArrowHourglass
                Set Frm = New FrmShowItemCostPrice
                Frm.LoadData LngItemID, val(.TextMatrix(.Row, .ColIndex("EndStock"))), IntCostItemPriceType
                Frm.show
                Me.MousePointer = vbDefault
            End If
        End If

    End With

End Sub

Private Sub Fg_MouseUp(Button As Integer, _
                       Shift As Integer, _
                       X As Single, _
                       Y As Single)
    Dim LngCurrentItemID As Long
    Dim LngMouseRow As Long

    If BolInProgress = True Then Exit Sub
    If Button = vbRightButton Then

        With Me.Fg
            LngMouseRow = .MouseRow

            If LngMouseRow = -1 Then Exit Sub
            If .Col = -1 Then Exit Sub
            mdifrmmain.MnuItemTools_ItemSerial.Tag = ""
            mdifrmmain.MnuItemTools_ItemCart.Tag = ""
            mdifrmmain.MnuItemTools_ItemData.Tag = ""
            mdifrmmain.MnuItemTools_ItemQty.Tag = ""
            mdifrmmain.MnuItemTools_ItemCostTrans.Tag = ""
        
            If val(.TextMatrix(LngMouseRow, .ColIndex("ItemID"))) <> 0 Then
                LngCurrentItemID = val(.TextMatrix(LngMouseRow, .ColIndex("ItemID")))
                mdifrmmain.MnuItemTools_ItemSerial.Enabled = False
                mdifrmmain.MnuItemTools_ItemSerial.Tag = ""
            
                mdifrmmain.MnuItemTools_ItemCart.Tag = LngCurrentItemID & "-" & Me.DcboStores.BoundText & "-" & Me.DtpFrom.value & "-" & Me.DtpTO.value
                mdifrmmain.MnuItemTools_ItemQty.Tag = LngCurrentItemID
                mdifrmmain.MnuItemTools_ItemData.Tag = LngCurrentItemID
                mdifrmmain.MnuItemTools_ItemCostTrans.Tag = LngCurrentItemID
                Me.PopupMenu mdifrmmain.MnuItemTools
            End If

        End With

    End If

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim GrdBack As ClsBackGroundPic
    Dim StrSQL As String

    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        .RowHeightMin = 320
        .AutoSizeMode = flexAutoSizeColWidth
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    Set Dcombos = New ClsDataCombos
    Dcombos.GetStores Me.DcboStores
    Dcombos.GetItemSGroups Me.DcboGroups
    Set cDboSearch(0) = New clsDCboSearch
    Set cDboSearch(0).Client = Me.DcboGroups

    Dcombos.GetItemsNames Me.DcboItems
    Set cDboSearch(1) = New clsDCboSearch
    Set cDboSearch(1).Client = Me.DcboGroups

    SetDtpickerDate Me.DtpFrom
    SetDtpickerDate Me.DtpTO

    With Me.CboType
        .Clear
        .AddItem "Ã—œ «·„Œ“Ê‰ Õ Ï  «—ÌŒ „⁄Ì‰"
        .AddItem "Õ—ﬂ… «·„Œ“Ê‰ Œ·«· › —…"
    End With

    With Me.CboCostType
        .Clear
        .AddItem "«Œ— ”⁄— ‘—«¡"
        .ItemData(0) = 1
        .AddItem "«·„ Ê”ÿ «·„—ÃÕ"
        .ItemData(1) = 2
        .AddItem "«·Ê«—œ √Ê·« Ì’—› √Ê·«"
        .ItemData(2) = 3
        .AddItem "«·„ Ê”ÿ «·„—ÃÕ «·ÃœÌœ"
        .ItemData(3) = 4
        '   .AddItem "«·Ê«—œ √ŒÌ— Ì’—› √Ê·«"
        '   .ItemData(3) = 4
        '   .AddItem "«Œ— ”⁄— »Ì⁄"
        '   .ItemData(4) = 5
    End With

    Resize_Form Me, ReportSize
    '--------------------------------
    StrSQL = " Update Transaction_details"
    StrSQL = StrSQL + " Set ItemDiscountType=0 , ItemDiscount=0"
    StrSQL = StrSQL + " Where ItemDiscountType Is Null"
    Cn.Execute StrSQL, , adExecuteNoRecords
    '--------------------------------
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If BolInProgress = True Then
        Cancel = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cDboSearch(0) = Nothing
    Set cDboSearch(1) = Nothing
End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    Select Case Index

        Case 2
            Me.lbl(2).ToolTipText = WriteNo(Me.lbl(2).Caption, 0)

        Case 9
            Me.lbl(9).ToolTipText = WriteNo(Me.lbl(9).Caption, 0)

        Case 3
            Me.lbl(3).ToolTipText = WriteNo(Me.lbl(3).Caption, 0)
    End Select

End Sub

Private Sub NewOperation()
    Me.Fg.Clear flexClearScrollable, flexClearEverything
    Me.Fg.Rows = Me.Fg.FixedRows
    Me.Fg.AutoSize 0, Fg.Cols - 1, False
    Me.lbl(2).Caption = 0
    Me.lbl(3).Caption = 0
    Me.lbl(9).Caption = 0
End Sub

Private Sub DoStockmovement()
    Dim StrSQL  As String
    Dim rs As ADODB.Recordset
    Dim RsMoveQty As ADODB.Recordset
    Dim i As Integer
    Dim Msg As String
    Dim StrStartDate As String
    Dim StrStartQtyDate As String
    Dim StrEndDate As String
    Dim SngItemCostPrice As Single
    Dim LngItemID As Long
    Dim DblItemQty As Double
    Dim EndDate As Variant
    Dim Fromdate As Variant
    Dim BolNoStartQry As Boolean
    Dim IntColName As Integer
    Dim RsData  As ADODB.Recordset
    Dim LngParentRow As Long, LngRowNum As Long

    If Me.DcboStores.BoundText = "" Then
        Msg = "ÌÃ» ≈Œ Ì«— «·„Œ“‰..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcboStores.SetFocus
        Exit Sub
    End If

    If Me.CboCostType.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… Õ”«»  ﬂ·›… «·„Œ“Ê‰....!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    ElseIf Me.CboCostType.ListIndex = 0 Then
        IntCostItemPriceType = LastPurPriceType
    ElseIf Me.CboCostType.ListIndex = 1 Then
        IntCostItemPriceType = WeightAverage
    ElseIf Me.CboCostType.ListIndex = 2 Then
        IntCostItemPriceType = FirstInFirstOut
    ElseIf Me.CboCostType.ListIndex = 3 Then
        IntCostItemPriceType = ModernWeightAverage
    End If

    If IsNull(Me.DtpFrom.value) Then
        BolNoStartQry = True
        StrStartDate = SQLDate(#1/1/1901#, True)
        StrStartQtyDate = SQLDate(#1/1/1901#, True)
    Else
        BolNoStartQry = False
        StrStartDate = SQLDate(Me.DtpFrom.value, True)
        StrStartQtyDate = SQLDate(Me.DtpFrom.value - 1, True)
    End If

    If IsNull(Me.DtpTO.value) Then
        StrEndDate = SQLDate(#1/1/2079#, True)
    Else
        StrEndDate = SQLDate(Me.DtpTO.value, True)
    End If

    Me.lbl(3).Caption = 0
    Me.lbl(9).Caption = 0
    Me.lbl(2).Caption = 0
    '--------------------
    CmdDo.Enabled = False

    '--------------------
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
    
        StrSQL = "Select ItemID, ItemCode,  ItemName,Sum(SumQuantity1)as SumQ1,Sum(SumQuantity_1)as SumQ_1,GroupID "
        StrSQL = StrSQL + " From "
        StrSQL = StrSQL + " ( "
        StrSQL = StrSQL + " SELECT ItemID, ItemCode, GroupName, ItemName, StoreID, StoreName, SumQuantity1," & "SumQuantity_1,GroupID "
        StrSQL = StrSQL + " FROM dbo.QryItemsInOutTransactions(" & StrStartDate & "," & StrEndDate & ")" & "QryItemsInOutTransactions "
        StrSQL = StrSQL + " )XTable "
        StrSQL = StrSQL + " Where StoreID =" & Me.DcboStores.BoundText & ""

        If Me.DcboGroups.BoundText <> "" Then
            StrSQL = StrSQL + " AND GroupID=" & val(Me.DcboGroups.BoundText) & ""
        End If

        StrSQL = StrSQL + " Group By ItemID,ItemCode,ItemName,GroupID"

        If Me.Opt(0).value = True Then
            StrSQL = StrSQL + " Order By GroupID,ItemID DESC "
        Else
            StrSQL = StrSQL + " Order By ItemID "
        End If

        Set RsMoveQty = New ADODB.Recordset
        RsMoveQty.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If RsMoveQty.BOF Or RsMoveQty.EOF Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.CmdDo.Enabled = True
            RsMoveQty.Close
            Set RsMoveQty = Nothing
            Exit Sub
        End If

        If BolNoStartQry = False Then
            StrSQL = "SELECT dbo.TblItems.ItemID, dbo.TblItems.ItemCode,"
            StrSQL = StrSQL + " dbo.TblItems.ItemName,Sum(QryQuantity.QTY)as Qty "
            StrSQL = StrSQL + " FROM dbo.TblItems LEFT OUTER JOIN"
            StrSQL = StrSQL + " dbo.QryQuantity(" & StrStartQtyDate & ", 0) QryQuantity"
            StrSQL = StrSQL + " ON dbo.TblItems.ItemID = QryQuantity.ItemID"
            StrSQL = StrSQL + " Where QryQuantity.StoreID=" & Me.DcboStores.BoundText & ""

            If Me.DcboGroups.BoundText <> "" Then
                StrSQL = StrSQL + " AND dbo.TblItems.GroupID=" & val(Me.DcboGroups.BoundText) & ""
            End If

            StrSQL = StrSQL + " Group BY dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName"
            StrSQL = StrSQL + " Order By dbo.TblItems.ItemID"
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        End If

    Else
        Exit Sub
    End If

    If Me.Opt(0).value = True Then 'Tree Style

        With Me.Fg
            .Redraw = flexRDNone
            .Rows = 1
            .ColPosition(.ColIndex("ItemName")) = 0

            If SystemOptions.UserInterface = ArabicInterface Then
                IntColName = 1
                .AddItem "‘Ã—… «·√’‰«›"
            Else
                .AddItem "Items Tree"
                IntColName = 1
            End If

            .Rowdata(.Rows - 1) = "1G"
            .IsSubtotal(.Rows - 1) = True
            .Cell(flexcpFontBold, .Rows - 1, 1) = True
            .GridLines = flexGridFlat
            .MergeCells = flexMergeSpill
            .OutlineBar = flexOutlineBarComplete
            .AllowUserResizing = flexResizeColumns
            .ExtendLastCol = True
            '.NodeClosedPicture = MDIFrmMain.ImgLstMenuIcons.ListImages("NodeClose").Picture
            '.NodeOpenPicture = MDIFrmMain.ImgLstMenuIcons.ListImages("NodeOpen").Picture
            .RowHeightMin = 300
            .ScrollTrack = False
            .ScrollTips = True
            .SheetBorder = vbWhite
            '---------------------------------------------------------------------
            'Load Groups Tree in the Grid
            StrSQL = " SELECT Groups.GroupID, Groups.GroupName, Groups.ParentID " & "FROM Groups Where Groups.ParentID=1"
            Set RsData = New ADODB.Recordset
        
            RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            Call LoadGridTree("1G", RsData, Fg, "Groups", "ParentID", "", , IntColName, vbBlue)
            .Redraw = True

            '---------------------------------------------------------------------
            For i = 0 To RsMoveQty.RecordCount - 1

                DoEvents
                LngParentRow = .FindRow(CStr(RsMoveQty("GroupID").value) & "G", 0, -1, False, True)

                If LngParentRow <> -1 Then
                    .AddItem "", (LngParentRow + 1)
                    Me.lbl(3).Caption = i
                
                    .Rowdata((LngParentRow + 1)) = RsMoveQty("ItemID").value & "I"
                    .RowOutlineLevel((LngParentRow + 1)) = .RowOutlineLevel(LngParentRow) + 1
                    .Cell(flexcpPicture, LngParentRow + 1, 0) = mdifrmmain.ImgLstTree.ListImages("Item").Picture
                
                    LngRowNum = LngParentRow + 1
               
                    .TextMatrix(LngRowNum, .ColIndex("Serial")) = LngRowNum
                    .TextMatrix(LngRowNum, .ColIndex("ItemID")) = IIf(IsNull(RsMoveQty("ItemID").value), "", RsMoveQty("ItemID").value)
                    LngItemID = val(.TextMatrix(LngRowNum, .ColIndex("ItemID")))
                    'If LngItemID = 709 Then
                    'Stop
                    'End If
                    .TextMatrix(LngRowNum, .ColIndex("ItemCode")) = IIf(IsNull(RsMoveQty("ItemCode").value), "", RsMoveQty("ItemCode").value)
                    .TextMatrix(LngRowNum, .ColIndex("ItemName")) = IIf(IsNull(RsMoveQty("ItemName").value), "", RsMoveQty("ItemName").value)

                    If BolNoStartQry = True Then
                        .TextMatrix(LngRowNum, .ColIndex("Qty")) = 0
                    Else
                        rs.find "ItemID=" & RsMoveQty("ItemID").value, , adSearchForward, 1

                        If Not (rs.BOF Or rs.EOF) Then
                            .TextMatrix(LngRowNum, .ColIndex("Qty")) = IIf(IsNull(rs("Qty").value), 0, rs("Qty").value)
                        Else
                            .TextMatrix(LngRowNum, .ColIndex("Qty")) = 0
                        End If
                    End If

                    '---------------------------------------
                    'Õ”«» ﬁÌ„… —’Ìœ »œ«Ì… «·› —…
                    If val(.TextMatrix(LngRowNum, .ColIndex("Qty"))) <= 0 Then
                        .TextMatrix(LngRowNum, .ColIndex("StartCost")) = 0
                    Else

                        If Not IsNull(Me.DtpFrom.value) Then
                            EndDate = Me.DtpFrom.value - 1
                        End If
                    
                        DblItemQty = val(.TextMatrix(LngRowNum, .ColIndex("Qty")))
                        SngItemCostPrice = GetCostItemPrice(LngItemID, 2, , , IntCostItemPriceType, DblItemQty, EndDate, EndDate)
                        .TextMatrix(LngRowNum, .ColIndex("StartCost")) = SngItemCostPrice * val(.TextMatrix(LngRowNum, .ColIndex("Qty")))
                        .TextMatrix(LngRowNum, .ColIndex("StartCost")) = Format(val(.TextMatrix(LngRowNum, .ColIndex("StartCost"))), SystemOptions.SysDefCurrencyForamt)
                    End If

                    '---------------------------------------
                    '«·ﬂ„Ì… «·„‰’—›…Ê«·„÷«›…„‰ «·’‰›
                    If Not (RsMoveQty.BOF Or RsMoveQty.EOF) Then
                        .TextMatrix(LngRowNum, .ColIndex("SumQuantity1")) = IIf(IsNull(RsMoveQty("SumQ1").value), 0, RsMoveQty("SumQ1").value)
                        .TextMatrix(LngRowNum, .ColIndex("SumQuantity_1")) = IIf(IsNull(RsMoveQty("SumQ_1").value), 0, RsMoveQty("SumQ_1").value)
                    Else
                        .TextMatrix(LngRowNum, .ColIndex("SumQuantity1")) = "0"
                        .TextMatrix(LngRowNum, .ColIndex("SumQuantity_1")) = "0"
                    End If

                    If val(.TextMatrix(LngRowNum, .ColIndex("SumQuantity1"))) = 0 And val(.TextMatrix(LngRowNum, .ColIndex("SumQuantity_1"))) = 0 Then
                        .Cell(flexcpForeColor, LngRowNum, 0, LngRowNum, .Cols - 1) = vbRed
                    End If

                    '-------------------------------------------
                    .TextMatrix(LngRowNum, .ColIndex("EndStock")) = (val(.TextMatrix(LngRowNum, .ColIndex("SumQuantity1"))) + val(.TextMatrix(LngRowNum, .ColIndex("Qty")))) - val(.TextMatrix(LngRowNum, .ColIndex("SumQuantity_1")))
                
                    LngItemID = .TextMatrix(LngRowNum, .ColIndex("ItemID"))
                    DblItemQty = val(.TextMatrix(LngRowNum, .ColIndex("EndStock")))

                    If Not IsNull(Me.DtpTO.value) Then
                        EndDate = Me.DtpTO.value
                    End If

                    SngItemCostPrice = GetCostItemPrice(LngItemID, 2, , , IntCostItemPriceType, DblItemQty, EndDate)

                    If IntCostItemPriceType = FirstInFirstOut Then
                        '›Ï Õ«·… «‰ ÌﬂÊ‰ ÿ—Ìﬁ… Õ”«» ﬁÌ„… «·„Œ“Ê‰
                        '«·Ê«—œ «Ê·«œ Ì’—› «Ê·« ... ›«‰ «·ﬁÌ„… «·⁄«∆œ… „‰
                        '«·œ«·…
                        'GetCostItemPrice
                        ' ”«ÊÏ ﬁÌ„… «·„Œ“Ê‰ «·ﬂ·Ï
                        'Ê·–« ·«Ì „ ÷—» «·ﬂ„Ì… ›Ï «·”⁄—
                        .TextMatrix(LngRowNum, .ColIndex("ItemCostPrice")) = SystemOptions.SysDefCurrencyForamt
                        .TextMatrix(LngRowNum, .ColIndex("StockCost")) = SngItemCostPrice
                        .TextMatrix(LngRowNum, .ColIndex("StockCost")) = Format(val(.TextMatrix(LngRowNum, .ColIndex("StockCost"))), SystemOptions.SysDefCurrencyForamt)
                    
                    Else
                        .TextMatrix(LngRowNum, .ColIndex("ItemCostPrice")) = SngItemCostPrice
                        .TextMatrix(LngRowNum, .ColIndex("StockCost")) = (SngItemCostPrice * val(.TextMatrix(LngRowNum, .ColIndex("EndStock"))))
                        .TextMatrix(LngRowNum, .ColIndex("StockCost")) = Format(val(.TextMatrix(LngRowNum, .ColIndex("StockCost"))), SystemOptions.SysDefCurrencyForamt)
                    End If
                End If

                RsMoveQty.MoveNext
            Next i

            For i = Me.Fg.FixedRows To Me.Fg.Rows - 1
                Dim XNode As VSFlex8UCtl.VSFlexNode
                Dim StrTemp As String, DblNodeChildCount As Double

                If .IsSubtotal(i) = True Then
                    Set XNode = Fg.GetNode(i)

                    If Not XNode Is Nothing Then
                        '⁄œœ «·√’‰«› «·„ÊÃÊœ… œ«Œ· Â–Â «·„Ã„Ê⁄…
                        DblNodeChildCount = ModFgLib.GetNodeChildTotal(Fg, XNode, flexSTCount)
                        StrTemp = XNode.text & " ( " & DblNodeChildCount & " ) "
                        XNode.text = StrTemp
                        '------------------------------------------------------
                        '≈Ã„«·Ï  ﬂ·›… «·√’‰«› «·„ÊÃÊœ… œ«Œ· Â–Â «·„Ã„Ê⁄…
                        DblNodeChildCount = ModFgLib.GetNodeChildTotal(Fg, XNode, flexSTSum, Fg.ColIndex("StockCost"))
                        StrTemp = " ( " & DblNodeChildCount & " ) "
                        StrTemp = XNode.text & " ( " & DblNodeChildCount & " ) "
                        XNode.text = StrTemp
                    End If
                End If

            Next i

        End With

    ElseIf Me.Opt(1).value = True Then 'Table Style

        With Me.Fg
            .Rows = .FixedRows
            Me.PrgBar.Visible = True
            Me.PrgBar.Max = RsMoveQty.RecordCount
            .Rows = .FixedRows + RsMoveQty.RecordCount

            For i = .FixedRows To .Rows - 1
                BolInProgress = True

                DoEvents
                Me.PrgBar.value = i
                Me.lbl(3).Caption = i
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsMoveQty("ItemID").value), "", RsMoveQty("ItemID").value)
                LngItemID = val(.TextMatrix(i, .ColIndex("ItemID")))
                'If LngItemID = 709 Then
                'Stop
                'End If
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsMoveQty("ItemCode").value), "", RsMoveQty("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsMoveQty("ItemName").value), "", RsMoveQty("ItemName").value)

                If BolNoStartQry = True Then
                    .TextMatrix(i, .ColIndex("Qty")) = 0
                Else
                    rs.find "ItemID=" & RsMoveQty("ItemID").value, , adSearchForward, 1

                    If Not (rs.BOF Or rs.EOF) Then
                        .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs("Qty").value), 0, rs("Qty").value)
                    Else
                        .TextMatrix(i, .ColIndex("Qty")) = 0
                    End If
                End If

                '---------------------------------------
                'Õ”«» ﬁÌ„… —’Ìœ »œ«Ì… «·› —…
                If val(.TextMatrix(i, .ColIndex("Qty"))) <= 0 Then
                    .TextMatrix(i, .ColIndex("StartCost")) = 0
                Else

                    If Not IsNull(Me.DtpFrom.value) Then
                        EndDate = Me.DtpFrom.value - 1
                    End If
                
                    DblItemQty = val(.TextMatrix(i, .ColIndex("Qty")))
                    SngItemCostPrice = GetCostItemPrice(LngItemID, 2, , , IntCostItemPriceType, DblItemQty, EndDate, EndDate)
                    .TextMatrix(i, .ColIndex("StartCost")) = SngItemCostPrice * val(.TextMatrix(i, .ColIndex("Qty")))
                    .TextMatrix(i, .ColIndex("StartCost")) = Format(val(.TextMatrix(i, .ColIndex("StartCost"))), SystemOptions.SysDefCurrencyForamt)
                End If

                '---------------------------------------
                '«·ﬂ„Ì… «·„‰’—›…Ê«·„÷«›…„‰ «·’‰›
                If Not (RsMoveQty.BOF Or RsMoveQty.EOF) Then
                    .TextMatrix(i, .ColIndex("SumQuantity1")) = IIf(IsNull(RsMoveQty("SumQ1").value), 0, RsMoveQty("SumQ1").value)
                    .TextMatrix(i, .ColIndex("SumQuantity_1")) = IIf(IsNull(RsMoveQty("SumQ_1").value), 0, RsMoveQty("SumQ_1").value)
                Else
                    .TextMatrix(i, .ColIndex("SumQuantity1")) = "0"
                    .TextMatrix(i, .ColIndex("SumQuantity_1")) = "0"
                End If

                If val(.TextMatrix(i, .ColIndex("SumQuantity1"))) = 0 And val(.TextMatrix(i, .ColIndex("SumQuantity_1"))) = 0 Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                End If

                '-------------------------------------------
                .TextMatrix(i, .ColIndex("EndStock")) = (val(.TextMatrix(i, .ColIndex("SumQuantity1"))) + val(.TextMatrix(i, .ColIndex("Qty")))) - val(.TextMatrix(i, .ColIndex("SumQuantity_1")))
            
                LngItemID = .TextMatrix(i, .ColIndex("ItemID"))
                DblItemQty = val(.TextMatrix(i, .ColIndex("EndStock")))

                If Not IsNull(Me.DtpTO.value) Then
                    EndDate = Me.DtpTO.value
                End If

                SngItemCostPrice = GetCostItemPrice(LngItemID, 2, , , IntCostItemPriceType, DblItemQty, EndDate)

                If IntCostItemPriceType = FirstInFirstOut Then
                    '›Ï Õ«·… «‰ ÌﬂÊ‰ ÿ—Ìﬁ… Õ”«» ﬁÌ„… «·„Œ“Ê‰
                    '«·Ê«—œ «Ê·«œ Ì’—› «Ê·« ... ›«‰ «·ﬁÌ„… «·⁄«∆œ… „‰
                    '«·œ«·…
                    'GetCostItemPrice
                    ' ”«ÊÏ ﬁÌ„… «·„Œ“Ê‰ «·ﬂ·Ï
                    'Ê·–« ·«Ì „ ÷—» «·ﬂ„Ì… ›Ï «·”⁄—
                    .TextMatrix(i, .ColIndex("ItemCostPrice")) = SystemOptions.SysDefCurrencyForamt
                    .TextMatrix(i, .ColIndex("StockCost")) = SngItemCostPrice
                    .TextMatrix(i, .ColIndex("StockCost")) = Format(val(.TextMatrix(i, .ColIndex("StockCost"))), SystemOptions.SysDefCurrencyForamt)
                
                Else
                    .TextMatrix(i, .ColIndex("ItemCostPrice")) = SngItemCostPrice
                    .TextMatrix(i, .ColIndex("StockCost")) = (SngItemCostPrice * val(.TextMatrix(i, .ColIndex("EndStock"))))
                    .TextMatrix(i, .ColIndex("StockCost")) = Format(val(.TextMatrix(i, .ColIndex("StockCost"))), SystemOptions.SysDefCurrencyForamt)
                End If

                RsMoveQty.MoveNext
                .ShowCell i, IIf(.Col = -1, 1, .Col)
            Next i

            BolInProgress = False
            .AutoSize 0, .Cols - 1, False
            Me.PrgBar.Visible = False
            Me.PrgBar.value = 0
            Me.lbl(2).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("StockCost"), .Rows - 1, .ColIndex("StockCost"))
            Me.lbl(2).Caption = Format(val(Me.lbl(2).Caption), SystemOptions.SysDefCurrencyForamt)
            Me.lbl(9).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("StartCost"), .Rows - 1, .ColIndex("StartCost"))
            Me.lbl(9).Caption = Format(val(Me.lbl(9).Caption), SystemOptions.SysDefCurrencyForamt)
        End With

    End If

    '-------------------
    CmdDo.Enabled = True
    '-------------------
End Sub

Private Sub DoStockCount()
    Dim StrSQL  As String
    Dim rs As ADODB.Recordset
    Dim StrEndDate As String
    Dim Msg As String
    Dim LngItemID As Long
    Dim DblItemQty As Double
    Dim SngItemCostPrice As Single
    Dim EndDate As Variant
    Dim i As Long, LngParentRow As Long, LngRowNum As Long
    Dim RsData As ADODB.Recordset
    Dim IntColName As Integer

    If IsNull(Me.DtpTO.value) Then
        StrEndDate = SQLDate(#1/1/2079#, True)
    Else
        StrEndDate = SQLDate(Me.DtpTO.value, True)
    End If

    StrSQL = "SELECT dbo.TblItems.ItemID, dbo.TblItems.ItemCode,"
    StrSQL = StrSQL + " dbo.TblItems.ItemName,dbo.TblItems.GroupID," & "Sum(QryQuantity.QTY)as Qty "
    StrSQL = StrSQL + " FROM dbo.TblItems LEFT OUTER JOIN"
    StrSQL = StrSQL + " dbo.QryQuantity(" & StrEndDate & ", 0) QryQuantity"
    StrSQL = StrSQL + " ON dbo.TblItems.ItemID = QryQuantity.ItemID"
    StrSQL = StrSQL + " Where QryQuantity.StoreID=" & Me.DcboStores.BoundText & ""

    If Me.DcboGroups.BoundText <> "" Then
        StrSQL = StrSQL + " AND dbo.TblItems.GroupID=" & val(Me.DcboGroups.BoundText) & ""
    End If

    StrSQL = StrSQL + " Group BY dbo.TblItems.ItemID, dbo.TblItems.ItemCode," & "dbo.TblItems.ItemName,dbo.TblItems.GroupID"

    StrSQL = StrSQL + " Order By dbo.TblItems.GroupID,dbo.TblItems.ItemID"

    If Me.Opt(0).value = True Then
        StrSQL = StrSQL + " DESC "
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.CmdDo.Enabled = True
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If Not IsNull(Me.DtpTO.value) Then
        EndDate = Me.DtpTO.value
    End If

    Me.PrgBar.Visible = True
    Me.PrgBar.Max = rs.RecordCount
    BolInProgress = True
        
    If Me.Opt(0).value = True Then 'Tree Style

        With Me.Fg
            .Redraw = flexRDNone
            .Rows = 1
            .ColPosition(.ColIndex("ItemName")) = 0

            If SystemOptions.UserInterface = ArabicInterface Then
                IntColName = 1
                .AddItem "‘Ã—… «·√’‰«›"
            Else
                .AddItem "Items Tree"
                IntColName = 1
            End If

            .Rowdata(.Rows - 1) = "1G"
            .IsSubtotal(.Rows - 1) = True
            .Cell(flexcpFontBold, .Rows - 1, 1) = True
            .GridLines = flexGridFlat
            .MergeCells = flexMergeSpill
            .OutlineBar = flexOutlineBarComplete
            .AllowUserResizing = flexResizeColumns
            .ExtendLastCol = True
            '.NodeClosedPicture = MDIFrmMain.ImgLstMenuIcons.ListImages("NodeClose").Picture
            '.NodeOpenPicture = MDIFrmMain.ImgLstMenuIcons.ListImages("NodeOpen").Picture
            .RowHeightMin = 300
            .ScrollTrack = False
            .ScrollTips = True
            .SheetBorder = vbWhite
            '-----------------------------------------
            '.ColHidden(.ColIndex("GroupName")) = True
            '-----------------------------------------
            StrSQL = " SELECT Groups.GroupID, Groups.GroupName, Groups.ParentID " & "FROM Groups Where Groups.ParentID=1"
            Set RsData = New ADODB.Recordset
        
            RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            Call LoadGridTree("1G", RsData, Fg, "Groups", "ParentID", "", , IntColName, vbBlue)
            .Redraw = True

            '--------------------------------------------------------------------------
            For i = 0 To rs.RecordCount - 1

                DoEvents
                LngParentRow = .FindRow(CStr(rs("GroupID").value) & "G", 0, -1, False, True)

                If LngParentRow <> -1 Then
                    .AddItem "", (LngParentRow + 1)
                    Me.lbl(3).Caption = i
                    .Rowdata((LngParentRow + 1)) = rs("ItemID").value & "I"
                    .RowOutlineLevel((LngParentRow + 1)) = .RowOutlineLevel(LngParentRow) + 1
                    .Cell(flexcpPicture, LngParentRow + 1, 0) = mdifrmmain.ImgLstTree.ListImages("Item").Picture
                
                    LngRowNum = LngParentRow + 1
               
                    .TextMatrix(LngRowNum, .ColIndex("Serial")) = i
                    .TextMatrix(LngRowNum, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                    .TextMatrix(LngRowNum, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    .TextMatrix(LngRowNum, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(LngRowNum, .ColIndex("Qty")) = IIf(IsNull(rs("Qty").value), 0, rs("Qty").value)
                    '-----------------------------------------------------------------
                    LngItemID = val(.TextMatrix(LngRowNum, .ColIndex("ItemID")))
                    DblItemQty = val(.TextMatrix(LngRowNum, .ColIndex("Qty")))
                    SngItemCostPrice = GetCostItemPrice(LngItemID, 2, , , IntCostItemPriceType, DblItemQty, EndDate)
                    .TextMatrix(LngRowNum, .ColIndex("ItemCostPrice")) = SngItemCostPrice
                    .TextMatrix(LngRowNum, .ColIndex("StockCost")) = SngItemCostPrice * val(.TextMatrix(LngRowNum, .ColIndex("Qty")))
                    .TextMatrix(LngRowNum, .ColIndex("StockCost")) = Format(val(.TextMatrix(LngRowNum, .ColIndex("StockCost"))), SystemOptions.SysDefCurrencyForamt)
                    '----------------------------------------------------------------------------------
                    .Cell(flexcpPictureAlignment, LngRowNum, 0) = flexPicAlignRightCenter
                Else
                    MsgBox "Stop My Group Is ((( NOT ))) Here"
                End If

                Me.PrgBar.value = i
                rs.MoveNext
            Next i

            For i = Me.Fg.FixedRows To Me.Fg.Rows - 1
                Dim XNode As VSFlex8UCtl.VSFlexNode
                Dim StrTemp As String, DblNodeChildCount As Double

                If .IsSubtotal(i) = True Then
                    Set XNode = Fg.GetNode(i)

                    If Not XNode Is Nothing Then
                        '⁄œœ «·√’‰«› «·„ÊÃÊœ… œ«Œ· Â–Â «·„Ã„Ê⁄…
                        DblNodeChildCount = ModFgLib.GetNodeChildTotal(Fg, XNode, flexSTCount)
                        StrTemp = XNode.text & " ( " & DblNodeChildCount & " ) "
                        XNode.text = StrTemp
                        '------------------------------------------------------
                        '≈Ã„«·Ï  ﬂ·›… «·√’‰«› «·„ÊÃÊœ… œ«Œ· Â–Â «·„Ã„Ê⁄…
                        DblNodeChildCount = ModFgLib.GetNodeChildTotal(Fg, XNode, flexSTSum, Fg.ColIndex("StockCost"))
                        StrTemp = " ( " & DblNodeChildCount & " ) "
                    
                        StrTemp = XNode.text & " ( " & DblNodeChildCount & " ) "
                        XNode.text = StrTemp
                        '                    .TextMatrix(I, .ColIndex("StockCost")) = StrTemp
                        '                    .Cell(flexcpForeColor, I, .ColIndex("StockCost")) = vbRed
                        '                    .Cell(flexcpFontBold, I, .ColIndex("StockCost")) = True
                        '                    .Cell(flexcpFontSize, I, .ColIndex("StockCost")) = 10
                    End If
                End If

            Next i

        End With

    ElseIf Opt(1).value = True Then 'Table Style
        Me.PrgBar.Visible = True
        Me.PrgBar.Max = rs.RecordCount

        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows + rs.RecordCount

            For i = .FixedRows To .Rows - 1

                DoEvents
                Me.PrgBar.value = i
                Me.lbl(3).Caption = i
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                LngItemID = val(.TextMatrix(i, .ColIndex("ItemID")))

                If LngItemID = 709 Then
                    'Stop
                End If

                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs("Qty").value), 0, rs("Qty").value)
                '----------------------------------------------------------------------------------
                DblItemQty = val(.TextMatrix(i, .ColIndex("Qty")))
                SngItemCostPrice = GetCostItemPrice(LngItemID, 2, , , IntCostItemPriceType, DblItemQty, EndDate)
                .TextMatrix(i, .ColIndex("ItemCostPrice")) = SngItemCostPrice
                .TextMatrix(i, .ColIndex("StockCost")) = SngItemCostPrice * val(.TextMatrix(i, .ColIndex("Qty")))
                .TextMatrix(i, .ColIndex("StockCost")) = Format(val(.TextMatrix(i, .ColIndex("StockCost"))), SystemOptions.SysDefCurrencyForamt)
                '----------------------------------------------------------------------------------
                rs.MoveNext
            Next i

        End With

    End If

    BolInProgress = False
    Fg.AutoSize 0, Fg.Cols - 1, False
    Me.PrgBar.Visible = False
    Me.PrgBar.value = 0
    Me.lbl(2).Caption = Fg.Aggregate(flexSTSum, Fg.FixedRows, Fg.ColIndex("StockCost"), Fg.Rows - 1, Fg.ColIndex("StockCost"))
    Me.lbl(2).Caption = Format(val(Me.lbl(2).Caption), SystemOptions.SysDefCurrencyForamt)

End Sub

Private Sub FgStockCountSetup()
    Dim i As Integer

    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
        .Cols = 0
        .Cols = 7
        .ColKey(0) = "Serial"
        .ColKey(1) = "ItemID"
        .ColKey(2) = "ItemCode"
        .ColKey(3) = "ItemName"
        .ColKey(4) = "Qty"
        .ColKey(5) = "ItemCostPrice"
        .ColKey(6) = "StockCost"
    
        .TextMatrix(0, .ColIndex("Serial")) = "„"
        .TextMatrix(0, .ColIndex("ItemID")) = "—ﬁ„ «·’‰›"
        .TextMatrix(0, .ColIndex("ItemCode")) = "ﬂÊœ «·’‰›"
        .TextMatrix(0, .ColIndex("ItemName")) = "«”„ «·’‰›"
        .TextMatrix(0, .ColIndex("Qty")) = "«·ﬂ„Ì…"
        .TextMatrix(0, .ColIndex("ItemCostPrice")) = "”⁄— «· ﬂ·›…"
        .TextMatrix(0, .ColIndex("StockCost")) = "ﬁÌ„…  ﬂ·›… «·„Œ“Ê‰"

        If SystemOptions.UserInterface = ArabicInterface Then
            .RightToLeft = True

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignRightCenter
                .FixedAlignment(i) = flexAlignRightCenter
            Next i

        Else
            .RightToLeft = False

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignLeftCenter
                .FixedAlignment(i) = flexAlignLeftCenter
            Next i

        End If

        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub FgStockMovmentSetup()
    Dim i As Integer

    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
        .Cols = 0
        .Cols = 11
        .ColKey(0) = "Serial"
        .ColKey(1) = "ItemID"
        .ColKey(2) = "ItemCode"
        .ColKey(3) = "ItemName"
        .ColKey(4) = "Qty"
        .ColKey(5) = "StartCost"
        .ColKey(6) = "SumQuantity1"
        .ColKey(7) = "SumQuantity_1"
        .ColKey(8) = "EndStock"
        .ColKey(9) = "ItemCostPrice"
        .ColKey(10) = "StockCost"
    
        .TextMatrix(0, .ColIndex("Serial")) = "„"
        .TextMatrix(0, .ColIndex("ItemID")) = "—ﬁ„ «·’‰›"
        .TextMatrix(0, .ColIndex("ItemCode")) = "ﬂÊœ «·’‰›"
        .TextMatrix(0, .ColIndex("ItemName")) = "«”„ «·’‰›"
        .TextMatrix(0, .ColIndex("Qty")) = "—’Ìœ »œ«Ì… «·› —…"
        .TextMatrix(0, .ColIndex("StartCost")) = "ﬁÌ„… —’Ìœ »œ«Ì… «·› —…"
        .TextMatrix(0, .ColIndex("SumQuantity1")) = "«·ﬂ„Ì… «·„÷«›…"
        .TextMatrix(0, .ColIndex("SumQuantity_1")) = "«·ﬂ„Ì… «·„‰’—›…"
        .TextMatrix(0, .ColIndex("EndStock")) = "«·—’Ìœ ›Ï ‰Â«Ì… «·› —…"
        .TextMatrix(0, .ColIndex("ItemCostPrice")) = "„ Ê”ÿ ”⁄— «· ﬂ·›…"
        .TextMatrix(0, .ColIndex("StockCost")) = "ﬁÌ„… «·„Œ“Ê‰ «·„ »ﬁÏ"
    
        If SystemOptions.UserInterface = ArabicInterface Then
            .RightToLeft = True

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignRightCenter
                .FixedAlignment(i) = flexAlignRightCenter
            Next i

        Else
            .RightToLeft = False

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignLeftCenter
                .FixedAlignment(i) = flexAlignLeftCenter
            Next i

        End If

        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub StockMovementPrint()
    Dim Msg As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer, j As Integer
    Dim cItemsReport As ClsItemsReport
    Dim StrCaption As String

    If ItemsInGrid(Me.Fg, Fg.ColIndex("ItemID")) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ «·√’‰«› √Ê·« ...!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    StrSQL = "Delete  From TempPrintStockMovement"
    Cn.Execute StrSQL, , adExecuteNoRecords

    Set rs = New ADODB.Recordset
    rs.Open "TempPrintStockMovement", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    With Me.Fg

        For i = .FixedRows To .Rows - 1

            DoEvents
            rs.AddNew
            rs("ItemID").value = IIf(.TextMatrix(i, .ColIndex("ItemID")) = "", Null, .TextMatrix(i, .ColIndex("ItemID")))
            rs("ItemCode").value = IIf(.TextMatrix(i, .ColIndex("ItemCode")) = "", Null, .TextMatrix(i, .ColIndex("ItemCode")))
            rs("ItemName").value = IIf(.TextMatrix(i, .ColIndex("ItemName")) = "", Null, .TextMatrix(i, .ColIndex("ItemName")))
            rs("BegainQty").value = IIf(.TextMatrix(i, .ColIndex("Qty")) = "", Null, .TextMatrix(i, .ColIndex("Qty")))
            rs("OutQty").value = IIf(.TextMatrix(i, .ColIndex("SumQuantity_1")) = "", Null, .TextMatrix(i, .ColIndex("SumQuantity_1")))
            rs("InQty").value = IIf(.TextMatrix(i, .ColIndex("SumQuantity1")) = "", Null, .TextMatrix(i, .ColIndex("SumQuantity1")))
            rs("EndQty").value = IIf(.TextMatrix(i, .ColIndex("EndStock")) = "", Null, .TextMatrix(i, .ColIndex("EndStock")))
            rs("ItemCostPrice").value = IIf(.TextMatrix(i, .ColIndex("ItemCostPrice")) = "", Null, .TextMatrix(i, .ColIndex("ItemCostPrice")))
            rs("ItemStockCost").value = IIf(.TextMatrix(i, .ColIndex("StockCost")) = "", Null, .TextMatrix(i, .ColIndex("StockCost")))
            rs("StartCost").value = IIf(.TextMatrix(i, .ColIndex("StartCost")) = "", Null, .TextMatrix(i, .ColIndex("StartCost")))
            rs.update
        Next i

    End With

    rs.Close
    Set rs = Nothing
    Set cItemsReport = New ClsItemsReport
    StrCaption = " ﬁ—Ì— »Õ—ﬂ… «·„Œ“Ê‰ "

    If Not IsNull(Me.DtpFrom.value) Then
        StrCaption = StrCaption & "›Ï «·› —… „‰ " & DisplayDate(DtpFrom.value)
    End If

    If Not IsNull(Me.DtpTO.value) Then
        StrCaption = StrCaption & " ≈·Ï " & DisplayDate(Me.DtpTO.value)
    End If

    cItemsReport.ShowStockMovement StrCaption, 0
    Set cItemsReport = Nothing
End Sub

Private Sub StockCountPrint()
    Dim Msg As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer, j As Integer
    Dim cItemsReport As ClsItemsReport
    Dim StrCaption As String

    If ItemsInGrid(Me.Fg, Fg.ColIndex("ItemID")) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ «·√’‰«› √Ê·« ...!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    StrSQL = "Delete  From TempPrintStockMovement"
    Cn.Execute StrSQL, , adExecuteNoRecords
    Set rs = New ADODB.Recordset
    rs.Open "TempPrintStockMovement", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    With Me.Fg

        For i = .FixedRows To .Rows - 1

            DoEvents

            If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
                rs.AddNew
                rs("ItemID").value = IIf(.TextMatrix(i, .ColIndex("ItemID")) = "", Null, .TextMatrix(i, .ColIndex("ItemID")))
                rs("ItemCode").value = IIf(.TextMatrix(i, .ColIndex("ItemCode")) = "", Null, .TextMatrix(i, .ColIndex("ItemCode")))
                rs("ItemName").value = IIf(.TextMatrix(i, .ColIndex("ItemName")) = "", Null, .TextMatrix(i, .ColIndex("ItemName")))
                rs("BegainQty").value = 0
                rs("OutQty").value = 0
                rs("InQty").value = 0
                rs("EndQty").value = IIf(.TextMatrix(i, .ColIndex("Qty")) = "", Null, .TextMatrix(i, .ColIndex("Qty")))
                rs("ItemCostPrice").value = IIf(.TextMatrix(i, .ColIndex("ItemCostPrice")) = "", Null, .TextMatrix(i, .ColIndex("ItemCostPrice")))
                rs("ItemStockCost").value = IIf(.TextMatrix(i, .ColIndex("StockCost")) = "", Null, .TextMatrix(i, .ColIndex("StockCost")))
                rs("StartCost").value = 0
                rs.update
            End If

        Next i

    End With

    rs.Close
    Set rs = Nothing
    Set cItemsReport = New ClsItemsReport
    StrCaption = " ﬁ—Ì— »Ã—œ «·„Œ“Ê‰ "

    If Not IsNull(Me.DtpFrom.value) Then
        StrCaption = StrCaption & "›Ï «·› —… „‰ " & DisplayDate(DtpFrom.value)
    End If

    If Not IsNull(Me.DtpTO.value) Then
        StrCaption = StrCaption & " ≈·Ï " & DisplayDate(Me.DtpTO.value)
    End If

    If Me.Opt(0).value = True Then
        cItemsReport.ShowStockMovement StrCaption, 1, 1
    ElseIf Me.Opt(1).value = True Then
        cItemsReport.ShowStockMovement StrCaption, 1, 0
    End If

    Set cItemsReport = Nothing
End Sub
