VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmVendorReceipt1 
   BackColor       =   &H00E2E9E9&
   Caption         =   "  ”‰œ ’—› „ ⁄ÂœÌ‰   "
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   408
   ClientWidth     =   16092
   Icon            =   "FrmVendorReceipt1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   16092
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic Main_CLE 
      Height          =   9360
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16092
      _cx             =   28385
      _cy             =   16510
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
         Height          =   1308
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   7836
         Width           =   16080
         _cx             =   28363
         _cy             =   2307
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Begin VB.CommandButton Command2 
            Caption         =   " ’œÌ—«·Ï «·«ﬂ”Ì·"
            Height          =   375
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   1488
         End
         Begin VB.Frame Frame9 
            Caption         =   "»Ì«‰«  „Õ«”»Ì…"
            Height          =   732
            Left            =   6936
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   0
            Width           =   8904
            Begin VB.CommandButton Command8 
               Caption         =   "ﬂ‘› Õ”«»"
               Height          =   375
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   240
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.CommandButton Command9 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   375
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox TxtNoteID 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   120
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   195
               Index           =   35
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   240
               Width           =   990
            End
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   396
            Index           =   0
            Left            =   13764
            TabIndex        =   8
            Top             =   792
            Width           =   1524
            _ExtentX        =   2688
            _ExtentY        =   699
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14871017
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Height          =   396
            Index           =   1
            Left            =   12084
            TabIndex        =   9
            Top             =   792
            Width           =   1644
            _ExtentX        =   2900
            _ExtentY        =   699
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14871017
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Height          =   396
            Index           =   2
            Left            =   10812
            TabIndex        =   10
            Top             =   792
            Width           =   1236
            _ExtentX        =   2180
            _ExtentY        =   699
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
            BackColor       =   14871017
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Height          =   396
            Index           =   3
            Left            =   9540
            TabIndex        =   11
            Top             =   792
            Width           =   1212
            _ExtentX        =   2138
            _ExtentY        =   699
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14871017
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Height          =   396
            Index           =   4
            Left            =   7308
            TabIndex        =   12
            Top             =   792
            Width           =   2172
            _ExtentX        =   3831
            _ExtentY        =   699
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   14871017
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Height          =   396
            Index           =   6
            Left            =   2712
            TabIndex        =   14
            Top             =   792
            Width           =   1632
            _ExtentX        =   2879
            _ExtentY        =   699
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Height          =   396
            Left            =   1176
            TabIndex        =   15
            Top             =   792
            Width           =   1488
            _ExtentX        =   2625
            _ExtentY        =   699
            ButtonPositionImage=   1
            Caption         =   "«·„—›ﬁ« "
            BackColor       =   14871017
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Height          =   396
            Index           =   7
            Left            =   5868
            TabIndex        =   13
            Top             =   792
            Width           =   1332
            _ExtentX        =   2350
            _ExtentY        =   699
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
            BackColor       =   14871017
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Height          =   396
            Index           =   5
            Left            =   4440
            TabIndex        =   49
            Top             =   792
            Width           =   1332
            _ExtentX        =   2350
            _ExtentY        =   699
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   336
            Left            =   2892
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   792
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   336
            Left            =   192
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   252
            Width           =   864
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   336
            Index           =   2
            Left            =   3744
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   252
            Width           =   1248
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   336
            Index           =   4
            Left            =   1116
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   252
            Width           =   1152
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   5712
         Left            =   0
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2040
         Width           =   16080
         _cx             =   28363
         _cy             =   10075
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Begin VB.CheckBox chkChooseAll 
            Alignment       =   1  'Right Justify
            Caption         =   "«Œ Ì«— «·ﬂ·"
            Height          =   372
            Left            =   14832
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   120
            Width           =   1104
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   5052
            Left            =   0
            TabIndex        =   6
            Top             =   600
            Width           =   16044
            _cx             =   28300
            _cy             =   8911
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
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
            Cols            =   49
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmVendorReceipt1.frx":038A
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
            Editable        =   1
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   372
            Left            =   3228
            TabIndex        =   42
            Top             =   120
            Visible         =   0   'False
            Width           =   10044
            _ExtentX        =   17717
            _ExtentY        =   656
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1116
         Left            =   0
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   756
         Width           =   16080
         _cx             =   28363
         _cy             =   1969
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   12900
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   624
            Width           =   1596
         End
         Begin VB.CommandButton Command1 
            Caption         =   "⁄—÷"
            Height          =   528
            Left            =   285
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   1020
         End
         Begin VB.ComboBox cbType 
            Height          =   288
            Left            =   9672
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   660
            Width           =   1548
         End
         Begin VB.TextBox txtCode 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   9672
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   264
            Width           =   1548
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   288
            Left            =   12912
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   1608
         End
         Begin MSDataListLib.DataCombo DcDur 
            Height          =   288
            Left            =   6132
            TabIndex        =   4
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   660
            Width           =   2508
            _ExtentX        =   4424
            _ExtentY        =   508
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMontth 
            Height          =   288
            Left            =   3576
            TabIndex        =   5
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   660
            Width           =   1524
            _ExtentX        =   2688
            _ExtentY        =   508
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker Date 
            Height          =   348
            Left            =   3564
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   240
            Width           =   1572
            _ExtentX        =   2773
            _ExtentY        =   614
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   109772803
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal DateH 
            Height          =   348
            Left            =   2256
            TabIndex        =   35
            Top             =   240
            Width           =   1332
            _ExtentX        =   2350
            _ExtentY        =   614
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   288
            Left            =   6144
            TabIndex        =   37
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   240
            Width           =   2508
            _ExtentX        =   4424
            _ExtentY        =   508
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï ÿ·» ’—›"
            Height          =   300
            Index           =   6
            Left            =   14616
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   624
            Width           =   1248
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   276
            Index           =   5
            Left            =   8880
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   588
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   5160
            TabIndex        =   36
            Top             =   240
            Width           =   876
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   276
            Index           =   1
            Left            =   5388
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   660
            Width           =   612
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   276
            Index           =   3
            Left            =   7608
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   660
            Width           =   1944
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·ÌœÊÏ"
            Height          =   324
            Index           =   9
            Left            =   11112
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   276
            Width           =   1452
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   324
            Index           =   8
            Left            =   14580
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   276
            Width           =   1248
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’—›"
            Height          =   240
            Index           =   0
            Left            =   11388
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   660
            Width           =   1188
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   732
         Left            =   0
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   16128
         _cx             =   28448
         _cy             =   1291
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
         Caption         =   "     ”‰œ ’—› „ ⁄ÂœÌ‰   "
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
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   18
            Top             =   120
            Width           =   495
            _ExtentX        =   868
            _ExtentY        =   614
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVendorReceipt1.frx":0A7C
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
            Height          =   345
            Index           =   2
            Left            =   90
            TabIndex        =   19
            Top             =   120
            Width           =   495
            _ExtentX        =   868
            _ExtentY        =   614
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVendorReceipt1.frx":0E16
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
            Height          =   345
            Index           =   1
            Left            =   1680
            TabIndex        =   20
            Top             =   120
            Width           =   495
            _ExtentX        =   868
            _ExtentY        =   614
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVendorReceipt1.frx":11B0
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
            Height          =   345
            Index           =   3
            Left            =   615
            TabIndex        =   21
            Top             =   120
            Width           =   495
            _ExtentX        =   868
            _ExtentY        =   614
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            FontSize        =   7.8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVendorReceipt1.frx":154A
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
   End
End
Attribute VB_Name = "FrmVendorReceipt1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim RsTemp2 As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim RsTemp3 As ADODB.Recordset
Dim RsTemp4 As ADODB.Recordset
Dim Account_Code_dynamic As String
Dim Account_Code_dynamic1 As String

Dim TTP As clstooltip

Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchId As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim strSQL As String
         strSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute strSQL, , adExecuteNoRecords



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
       ' Msg = EleHeader.Caption & " ??? " & txtID & " ??????" & Date
 'Dim msg As String
 Msg = EleHeader.Caption & " ··”‰… " & DcDur.text & " ··› —…  " & dcMontth.text
        
 
 notes_id = general_noteid

  
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    'C???? C??I?? C?C?C?CE
     
    Dim CValue As Double
    Dim branch As Integer
    Dim ProjectID As Integer
    
    BranchId = 1
    
    With Grid


line_no = 1
        For i = .FixedRows To .Rows - 1
    BranchId = val(dcBranch.BoundText)
    
            If .TextMatrix(i, .ColIndex("Value")) > 0 And Account_Code_dynamic1 <> "" And .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then     'C?C??? C???E??E IC??
                'Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") 'C?C??? C???E??E
         '       StrAccountCode = .TextMatrix(i, .ColIndex("Account_Code"))
        
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("Value")), 0, Msg & "  ﬁÌ„… «·œ›⁄Â  «·„” Õﬁ…  ··„ ⁄Âœ " & .TextMatrix(i, .ColIndex("cusname")) & " ··”Ì«—…" & .TextMatrix(i, .ColIndex("Car")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                  
            End If
            
       If .TextMatrix(i, .ColIndex("net")) > 0 And .TextMatrix(i, .ColIndex("Account_Code")) <> "" And .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then   'C?C??? C???E??E IC??
                'Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") 'C?C??? C???E??E
             StrAccountCode = .TextMatrix(i, .ColIndex("Account_Code"))
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("net")), 1, Msg & "     ’«›Ì   ··„ ⁄Âœ " & .TextMatrix(i, .ColIndex("cusname")) & " ··”Ì«—…" & .TextMatrix(i, .ColIndex("Car")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                  
            End If
            
            
            
       If .TextMatrix(i, .ColIndex("Total")) > 0 And Account_Code_dynamic <> "" And .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then   'C?C??? C???E??E IC??
                'Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") 'C?C??? C???E??E
         '       StrAccountCode = .TextMatrix(i, .ColIndex("Account_Code"))
        
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("Total")), 1, Msg & "  Õ”„Ì«      ··„ ⁄Âœ " & .TextMatrix(i, .ColIndex("cusname")) & " ··”Ì«—…" & .TextMatrix(i, .ColIndex("Car")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                  
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
des = EleHeader.Caption & " ··”‰… " & DcDur.text & " ··› —…  " & dcMontth.text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchId As Integer
 

Dim sql As String
tablename = "TblExchangeRequest"
Filedname = "ID"
NoteSerial1 = val(txtID)
Notevalue = 0

 notytype = 8069
'Notevalue = val(total)
 

 BranchId = val(dcBranch.BoundText)
NoteDate = Me.Date.value
 
'If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchId, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TxtNoteID.text = NoteID
                                                     TxtNoteSerial.text = NoteSerial
                                     Else
                                                 If TxtNoteID.text = "" Or TxtNoteSerial.text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchId, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TxtNoteID.text = NoteID
                                                                TxtNoteSerial.text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TxtNoteID.text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TxtNoteID.text), BranchId, user_id, NoteDate
rs.Resync adAffectCurrent
 

'     End If

End Function

Private Sub chkChooseAll_Click()
Dim i As Integer

For i = 1 To Grid.Rows - 1
    If Grid.TextMatrix(i, Grid.ColIndex("fullcode")) <> "" Then
            If chkChooseAll.value = 1 Then
                    Grid.TextMatrix(i, Grid.ColIndex("Status")) = 1
            Else
                    Grid.TextMatrix(i, Grid.ColIndex("Status")) = 0
            End If
    End If
Next
End Sub

Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.text = "N"
            clear_all Me
            txtID.text = CStr(new_id("TblVendorReceipt", "ID", "", True))
          '  TXTid.SetFocus
             Grid.Rows = Grid.FixedRows
        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"

        Case 2
 
                  Account_Code_dynamic = get_account_code_branch(106, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Exit Sub
        ElseIf Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ  ﬂ·›… «·‰ﬁ· ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            Exit Sub
                
        End If
     
     
                       Account_Code_dynamic1 = get_account_code_branch(107, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Exit Sub
        ElseIf Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ  œ›⁄«  „ ⁄ÂœÌ‰ „” Õﬁ…   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            Exit Sub
                
        End If
        
        
          
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_Company

        Case 5
                Unload FrmSearch_Request
                FrmSearch_Request.SendForm = "VR_VR"
                FrmSearch_Request.show
        Case 6
            Unload Me
         Case 7
   '      print_report2
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAttach_Click()
            On Error Resume Next
'ShowAttachments XPTxtBoxID, "0701201405"
 

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub Command2_Click()
   On Error Resume Next
    Dim StrFileName As String
    StrFileName = App.path & "\Report1.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
'Grid.RightToLeft = True
    Me.Grid.SaveGrid StrFileName, flexFileExcel, True
    OpenFile StrFileName
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.text, , 200
End Sub

Private Sub Date_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
TxtNoteSerial.text = ""
End If
End Sub

Private Sub Dcbranch_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
TxtNoteSerial.text = ""
End If


End Sub

Private Sub Dcbranch_Click(Area As Integer)
Dcbranch_Change
End Sub

Private Sub Dcbranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

 
End Sub

Private Sub dcEmp_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

 

End Sub

Private Sub Option1_Click()
 
End Sub

Private Sub Option2_Click()
 
End Sub

 
Private Sub Command1_Click()
 
ProgressBar1.Visible = True
 
' If check_reg = True Then
' Exit Sub
' End If
 
Fill_Grid
ProgressBar1.Visible = False
ProgressBar1.value = 0
End Sub

Private Function check_reg() As Boolean

Dim str As String
Grid.Rows = Grid.FixedRows

str = " select * from TblVendorReceipt where durationid = " & val(DcDur.BoundText) & "  and Month =   " & val(dcMontth.BoundText) & "  and BranchID = " & val(dcBranch.BoundText)
Set RsTemp = New ADODB.Recordset
RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsTemp.RecordCount > 0 Then
        MsgBox (" „  ”ÃÌ· ”‰œ ’—› ·Â–Â «·› —… „‰ ﬁ»· ")
        check_reg = True
Else

        check_reg = False
End If
End Function


Private Sub DcDur_Change()
Dim i As Integer, j As Integer, str As String
    i = val(DcDur.BoundText)
    
    If i > 0 Then
        str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo dcMontth, str
    Else
        str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo dcMontth, str
    End If

End Sub

Private Sub Fill_Grid()

Dim i As Integer, j As Integer, str As String
    i = val(DcDur.BoundText)
   
    Grid.Rows = Grid.FixedRows
  
 str = "  SELECT dbo.TblAttributionContract.IDAC, dbo.TblAttributionContract.DurationID, dbo.TblDurations.Name AS DurationName, dbo.TblCustemers.CusName,"
 str = str & "  dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.CusID, dbo.TblAttributionContract.StartContractDate,"
 str = str & "  dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.FromDate, dbo.TblDurations.FromDate AS DurFromDate, dbo.TblDurations.ToDate AS DurToDate,"
 str = str & "  dbo.TblVehicleAllocation_Details.Type, dbo.TblVehicleAllocation_Details.StudentCount, dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblVehicleAllocation_Details.rate,"
 str = str & "  dbo.TblVehicleAllocation_Details.Custom, dbo.TblVehicleAllocation_Details.DayRate, dbo.TblVehicleAllocation_Details.StudentCustom,"
 str = str & "  dbo.TblVehicleAllocation_Details.CarID, dbo.TblDurations_Details.Name, dbo.TblDurations_Details.ID AS MonthID, dbo.TblVehicleAllocation_Details.ID,"
 str = str & "  dbo.TblVehicleAllocation_Details.SchoolFileID , dbo.TblVendorCars.StopDeal, dbo.TblVendorCars.StopDate, dbo.TblVendorCars.StopDateH"
 str = str & "  , TblAttributionContract.BranchID  , dbo.TblCustemers.Account_Code, dbo.ACCOUNTS.Account_Serial , dbo.TblCustemers.IBAN, dbo.TblCustemers.BankAccount"
 str = str & "  FROM     dbo.TblVendorCars INNER JOIN "
 str = str & "  dbo.TblAttributionContract INNER JOIN "
 str = str & "  dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID INNER JOIN"
 str = str & "  dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA ON"
 str = str & "  dbo.TblVendorCars.ID = dbo.TblVehicleAllocation_Details.CarID LEFT OUTER JOIN"
 str = str & "  dbo.TblDurations_Details INNER JOIN"
 str = str & "  dbo.TblDurations ON dbo.TblDurations_Details.DID = dbo.TblDurations.ID ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID"
 str = str & "  INNER JOIN dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code "
str = str & "   Where (dbo.TblVehicleAllocation_Details.Type = 3)   and  ( TblVendorCars.StopDeal is null  or dbo.TblVendorCars.StopDeal = 0 )   "

str = str & "  and  TblVehicleAllocation_Details.id not in  (SELECT dbo.TblVendorReceipt_Details.InsID"
  str = str & "                                                FROM   dbo.TblVendorReceipt_Details INNER JOIN"
    str = str & "                                                     dbo.TblVendorReceipt ON dbo.TblVendorReceipt_Details.HID = dbo.TblVendorReceipt.ID"
    str = str & "                                              where DurationID = " & val(DcDur.BoundText) & "  and Month =  " & val(dcMontth.BoundText) & "  and BranchId =  " & val(dcBranch.BoundText) & "   ) "
   
  If DcDur.BoundText <> "" Then
          str = str & "  and  TblAttributionContract.DurationID  = " & val(DcDur.BoundText)
  End If
    
  If dcMontth.BoundText <> "" Then
          str = str & "      and  TblDurations_Details.ID = " & val(dcMontth.BoundText)
  End If
     
  If dcBranch.BoundText <> "" Then
          str = str & "      and  TblAttributionContract.BranchID  = " & val(dcBranch.BoundText)
  End If
      
    
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim V As Integer, H As Integer, WD
    Dim tot As Integer, daycount As Integer, SchoolFileID As Integer
    
    If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            With Grid
             ProgressBar1.Max = RsTemp.RecordCount
             Grid.Rows = Grid.FixedRows + RsTemp.RecordCount
             For j = Grid.FixedRows To Grid.Rows - 1
                     ProgressBar1.value = j - 2
                     .TextMatrix(j, .ColIndex("Serial")) = j - 1
                    .TextMatrix(j, .ColIndex("IDAC")) = IIf(IsNull(RsTemp("IDAC").value), "", RsTemp("IDAC").value)
                    .TextMatrix(j, .ColIndex("fullcode")) = IIf(IsNull(RsTemp("Fullcode").value), "", RsTemp("Fullcode").value)
                    .TextMatrix(j, .ColIndex("cusname")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
                    '.TextMatrix(j, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp("InstallmentNo").value), "", RsTemp("InstallmentNo").value)
                    .TextMatrix(j, .ColIndex("ID")) = IIf(IsNull(RsTemp("ID").value), "", RsTemp("ID").value)
                    .TextMatrix(j, .ColIndex("CusID")) = IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value)
                    .TextMatrix(j, .ColIndex("MonthID")) = IIf(IsNull(RsTemp("MonthID").value), "", RsTemp("MonthID").value)
                    .TextMatrix(j, .ColIndex("StartContractDate")) = IIf(IsNull(RsTemp("StartContractDate").value), "", RsTemp("StartContractDate").value)
                    .TextMatrix(j, .ColIndex("EndContractDate")) = IIf(IsNull(RsTemp("EndContractDate").value), "", RsTemp("EndContractDate").value)
                    .TextMatrix(j, .ColIndex("FromDate")) = IIf(IsNull(RsTemp("FromDate").value), "", RsTemp("FromDate").value)
                    .TextMatrix(j, .ColIndex("Car")) = IIf(IsNull(RsTemp("BoardNo").value), "", RsTemp("BoardNo").value)
                    .TextMatrix(j, .ColIndex("CarID")) = IIf(IsNull(RsTemp("CarID").value), "", RsTemp("CarID").value)
                    
                    .TextMatrix(j, .ColIndex("Account_Code")) = IIf(IsNull(RsTemp("Account_Code").value), "", RsTemp("Account_Code").value)
                    .TextMatrix(j, .ColIndex("Account_Serial")) = IIf(IsNull(RsTemp("Account_Serial").value), "", RsTemp("Account_Serial").value)
                    
                    .TextMatrix(j, .ColIndex("IBAN")) = IIf(IsNull(RsTemp("IBAN").value), "", RsTemp("IBAN").value)
                    .TextMatrix(j, .ColIndex("BankAccount")) = IIf(IsNull(RsTemp("BankAccount").value), "", RsTemp("BankAccount").value)
                    
                    If Not (IsNull(RsTemp("IDAC").value) Or IsNull(RsTemp("MonthID").value)) Then
                            'V = GetVac((RsTemp("IDAC").value), (RsTemp("MonthID").value))
                         '   GetVac (RsTemp("IDAC").value), (RsTemp("MonthID").value), RsTemp("DurationID").value, RsTemp("SchoolFileID").value, tot, daycount
                         '   H = GetHold((RsTemp("MonthID").value))
                         '   GetDeducts RsTemp("IDAC").value, DcDur.BoundText, RsTemp("MonthID").value, IIf(IsNull(RsTemp("CarID").value), 0, RsTemp("CarID").value), j
                            
                            
                    End If
                     
                     If Not IsNull(RsTemp("MonthID").value) Then
                            WD = GetMonthDays(RsTemp("MonthID").value)
                     End If
                    .TextMatrix(j, .ColIndex("VacDay")) = daycount   ' V ' + H
                    .TextMatrix(j, .ColIndex("WorkDay")) = WD - H
                     
                     Dim dayrate As Double
                     dayrate = IIf(IsNull(RsTemp("DayRate").value), 0, RsTemp("DayRate").value)
                      
                    
                     If Not (IsNull(RsTemp("IDAC").value)) Then
                           .TextMatrix(j, .ColIndex("VacValue")) = daycount * dayrate
                     End If
                                       
                     .TextMatrix(j, .ColIndex("DayRate")) = dayrate
                   
                    .TextMatrix(j, .ColIndex("Value")) = dayrate * (WD - H)
                    RsTemp.MoveNext
             Next
            End With
    End If
calculation
End Sub

Private Sub GetVac(IDMC As Integer, MonthID As Integer, DurationID As Integer, SchoolFileID As Integer, ByRef total As Integer, ByRef daycount As Integer)

        Dim str As String, cunt As Integer, CityID As Integer, DurID As Integer, DayDiff As Integer, j As Integer, i As Integer
            
        total = 0
        daycount = 0
     
     '   str = " select  d.schoolfileid , a.DurationID  from TblAttributionContract a ,  TblVehicleAllocation_Details  d where a.IDAC = d.IDVA   and d.type = 3  and  a.IDAC =  " & IDMC
     '   str = str & "   group by schoolfileid , DurationID  "
     '
    '    Set RsTemp2 = New ADODB.Recordset
    '    RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '
    '   For j = 0 To RsTemp2.RecordCount - 1
    ''
    '            DurationID = IIf(IsNull(RsTemp2("DurationID").value), 0, RsTemp2("DurationID").value)
    '            SchoolFileID = IIf(IsNull(RsTemp2("schoolfileid").value), 0, RsTemp2("schoolfileid").value)
    '
                str = " select h.DurationID , h.MonthID ,d.SchoolFileID ,sum (d.daycount) daycount , sum (d.dayvalue)  dayvalue , sum ( (d.daycount * d.dayvalue )) Total"
                str = str & " from TblconfirmVacation  h, TblConfirmVacation_Details d "
                str = str & " where  h.ID = d.HID and  DurationID = " & DurationID & "  and  MonthID = " & MonthID & " and  SchoolFileID = " & SchoolFileID
                str = str & " group by  h.DurationID , h.MonthID ,d.SchoolFileID  "
                
                Set RsTemp3 = New ADODB.Recordset
                RsTemp3.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp3.RecordCount > 0 Then
                    For i = 0 To RsTemp3.RecordCount - 1
                         total = total + IIf(IsNull(RsTemp3("Total").value), 0, RsTemp3("Total").value)
                         daycount = daycount + IIf(IsNull(RsTemp3("daycount").value), 0, RsTemp3("daycount").value)
                         RsTemp3.MoveNext
                   Next
                End If
    '            RsTemp2.MoveNext
    '   Next
          
          
     
End Sub

Private Function GetDayRate(IDMC As Integer, FromDate As String, ToDate As String)

 Dim str As String, days As Integer, net As Double, Operation As String
 str = " select * from TblAttributionContract  where idac = " & IDMC
 Set Rs_Temp = New ADODB.Recordset
 Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs_Temp.RecordCount > 0 Then
'        days = DateDiff("d", Rs_Temp("StartContractDate").value, Rs_Temp("EndContractDate").value)

        days = DateDiff("d", FromDate, ToDate)
        If days > 0 Then
                Operation = IIf(IsNull(Rs_Temp("AdditionalType").value), "", Rs_Temp("AdditionalType").value)
                If Operation = "add" Then
                       net = val(Rs_Temp("studentcount").value) * val(Rs_Temp("StudentCustom").value) + val(Rs_Temp("StudentCustom").value)
                ElseIf Operation = "sub" Then
                       net = val(Rs_Temp("studentcount").value) * val(Rs_Temp("StudentCustom").value) - val(Rs_Temp("StudentCustom").value)
                Else
                        net = val(Rs_Temp("studentcount").value) * val(Rs_Temp("StudentCustom").value)
                End If
                GetDayRate = net / days
        End If
 End If
 
End Function

Private Sub GetDeducts(IDMC As Integer, dur As Integer, MonthID As Integer, CarID As Integer, Row As Integer)

         Dim str As String, i As Integer, j As Integer, absc As Boolean
         str = "SELECT   dbo.TblConfirmViolation.ID, dbo.TblConfirmViolation.DurationID, dbo.TblConfirmViolation.ViolationID, dbo.TblConfirmViolation.MinistryContractID ,TblConfirmViolation.AbsenceCount,"
         str = str & " dbo.TblConfirmViolation.Date , dbo.TblConfirmViolation.value, dbo.TblConfirmViolation.monthid, dbo.TblViolationTypes.name ,TblViolationTypes.absence"
         str = str & " FROM     dbo.TblConfirmViolation INNER JOIN   dbo.TblViolationTypes ON dbo.TblConfirmViolation.ViolationID = dbo.TblViolationTypes.ID"
         str = str & " where DurationID = " & dur & " and MonthID = " & MonthID & "  and MinistryContractID =  " & IDMC & " and TblConfirmViolation.CarID =  " & CarID
         
         Set RsTemp4 = New ADODB.Recordset
         RsTemp4.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        
         With Grid
         If RsTemp4.RecordCount > 0 Then
         
                For i = 1 To RsTemp4.RecordCount
                
                         absc = IIf(IsNull(RsTemp4("absence").value), False, RsTemp4("absence").value)
                         
                         If absc = True Then
                                .TextMatrix(Row, .ColIndex("AbsenceCount")) = val(.TextMatrix(Row, .ColIndex("AbsenceCount"))) + IIf(IsNull(RsTemp4("AbsenceCount").value), 0, RsTemp4("AbsenceCount").value)
                                .TextMatrix(Row, .ColIndex("Avalue")) = val(.TextMatrix(Row, .ColIndex("Avalue"))) + IIf(IsNull(RsTemp4("Value").value), 0, RsTemp4("Value").value)
                         End If
                
                        For j = 1 To 20
                                If .TextMatrix(1, .ColIndex("d" & j)) = RsTemp4("Name").value Then
                                        .TextMatrix(Row, .ColIndex("d" & j)) = RsTemp4("Value").value
                                End If
                        Next
                        RsTemp4.MoveNext
                Next
         End If
         End With
         
End Sub

Private Function GetHold(MonthID As Integer)
    Dim str As String, cunt As Integer
             str = " select count (*)  cunt, DDID  from TblVacationSchedule where ISVac = 1 and  ddid =   " & MonthID & " group by DDID "
             Set RsTemp2 = New ADODB.Recordset
             RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
             
             If RsTemp2.RecordCount > 0 Then
                    cunt = IIf(IsNull(RsTemp2("cunt").value), 0, RsTemp2("cunt").value)
             End If
             
    GetHold = cunt
End Function

Private Function GetVac1(IDMC As Integer, MonthID As Integer)

        Dim str As String, cunt As Integer, CityID As Integer, DurID As Integer, DayDiff As Integer, j As Integer
        str = " select CityID , DurationID  from  TblAttributionContract where IDAC = " & IDMC
        Set RsTemp2 = New ADODB.Recordset
        RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If RsTemp2.RecordCount > 0 Then
               CityID = IIf(IsNull(RsTemp2("CityID").value), 0, RsTemp2("CityID").value)
               DurID = IIf(IsNull(RsTemp2("DurationID").value), 0, RsTemp2("DurationID").value)
        End If
        str = "select * from TblConfirmVacation  where DurationID = " & DurID & "and CityID = " & CityID & " and MonthID = " & MonthID
        Set RsTemp2 = New ADODB.Recordset
        RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If RsTemp2.RecordCount > 0 Then
            For j = 0 To RsTemp2.RecordCount - 1
                 DayDiff = DayDiff + DateDiff("d", RsTemp2("FromDate").value, RsTemp2("ToDate").value, vbSaturday)
           Next
        End If
        GetVac1 = DayDiff
End Function

Private Function GetMonthDays(MonthID As Integer)

    Dim str As String, cunt As Integer
             str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & " group by DDID "
             Set RsTemp2 = New ADODB.Recordset
             RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
             
             If RsTemp2.RecordCount > 0 Then
                    cunt = IIf(IsNull(RsTemp2("cunt").value), 0, RsTemp2("cunt").value)
             End If
             
    GetMonthDays = cunt

End Function

Private Sub dcMontth_Click(Area As Integer)
'Fill_Grid
End Sub

Private Sub Form_Activate()
'    XPTxtBoxID.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches dcBranch
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " ”‰œ  ’—› „ ⁄ÂœÌ‰  "
    LogTextE = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""

    Dim My_SQL As String
    My_SQL = " Select id , name from  TblDurations "
    fill_combo DcDur, My_SQL
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
  
   Dim strSQL As String
   strSQL = "SELECT  *  From TblVendorReceipt order by ID"
   rs.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
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
        
    Me.TxtModFlg.text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    Intialize_Deducts

    Inatial_Grid

    Exit Sub

ErrTrap:
End Sub

Private Sub Inatial_Grid()

 With Grid

        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        
       ' .MergeCol(.ColIndex("No")) = True
        .Cell(flexcpText, 0, .ColIndex("No"), 1, .ColIndex("No")) = "—ﬁ„ «·”ÿ—"

        .MergeCol(.ColIndex("cusname")) = True
        .Cell(flexcpText, 0, .ColIndex("cusname"), 1, .ColIndex("cusname")) = "«·«”„"

        .MergeCol(.ColIndex("PayNo")) = True
        .Cell(flexcpText, 0, .ColIndex("PayNo"), 1, .ColIndex("PayNo")) = "—ﬁ„ «·œ›⁄…"

        .MergeCol(.ColIndex("Value")) = True
        .Cell(flexcpText, 0, .ColIndex("Value"), 1, .ColIndex("Value")) = "«·ﬁÌ„…"

        .MergeCol(.ColIndex("Total")) = True
        .Cell(flexcpText, 0, .ColIndex("total"), 1, .ColIndex("total")) = "«Ã„«·Ï «·„” Õﬁ« "
        
        .MergeCol(.ColIndex("Net")) = True
        .Cell(flexcpText, 0, .ColIndex("Net"), 1, .ColIndex("Net")) = "«·’«›Ï «·„” Õﬁ"
        .Cell(flexcpText, 0, .ColIndex("d1"), 0, .ColIndex("d20")) = "Õ”„Ì« "
 
    End With



End Sub



Private Sub Intialize_Deducts()
Dim str As String, i As Integer
Set Rs_Temp = New ADODB.Recordset
str = " select * from TblViolationTypes  where absence = 0 or absence is null"
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
Rs_Temp.MoveFirst

If Rs_Temp.RecordCount > 0 Then
    For i = 1 To Rs_Temp.RecordCount
        Grid.TextMatrix(1, Grid.ColIndex("d" & i)) = IIf(IsNull(Rs_Temp("Name").value), "", Rs_Temp("Name").value)
        Rs_Temp.MoveNext
    Next
End If


For i = 1 To 20
    If Grid.TextMatrix(1, Grid.ColIndex("d" & i)) = "" Then
         Grid.ColWidth(Grid.ColIndex("d" & i)) = 0
    End If
Next


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

 
   lbl(0).Caption = "No."
   lbl(3).Caption = " Name Ar"
   lbl(7).Caption = " Name En"
   'Label3.Caption = "City"
   
  lbl(2).Caption = "Current Record"
  lbl(4).Caption = "Recors Count"
   
    Me.Caption = "Managerial Area"
    EleHeader.Caption = Me.Caption
   
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    CmdAttach.Caption = "Attachment"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  ”‰œ ’—› „ ⁄ÂœÌ‰   "
    LogTextE = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub



Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

     With Grid
            Select Case .ColKey(Col)
                Case "d1", "d2", "d3", "d4", "d5", "d6", "d7", "d8", "d9", "d10", "d11", "d12", "d13", "d14", "d15", "d16", "d17", "d18", "d19", "d20", "Value"
                        .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("Avalue"))) + val(.TextMatrix(Row, .ColIndex("d1"))) + val(.TextMatrix(Row, .ColIndex("d2"))) + val(.TextMatrix(Row, .ColIndex("d3"))) + val(.TextMatrix(Row, .ColIndex("d4"))) + val(.TextMatrix(Row, .ColIndex("d5"))) + val(.TextMatrix(Row, .ColIndex("d6"))) + val(.TextMatrix(Row, .ColIndex("d7"))) + val(.TextMatrix(Row, .ColIndex("d8"))) + val(.TextMatrix(Row, .ColIndex("d9"))) + val(.TextMatrix(Row, .ColIndex("d10"))) + val(.TextMatrix(Row, .ColIndex("d11"))) + val(.TextMatrix(Row, .ColIndex("d12"))) + val(.TextMatrix(Row, .ColIndex("d13"))) + val(.TextMatrix(Row, .ColIndex("d14"))) + val(.TextMatrix(Row, .ColIndex("d15"))) + val(.TextMatrix(Row, .ColIndex("d16"))) + val(.TextMatrix(Row, .ColIndex("d17"))) + val(.TextMatrix(Row, .ColIndex("d18"))) + val(.TextMatrix(Row, .ColIndex("d19"))) + val(.TextMatrix(Row, .ColIndex("d20")))
                        .TextMatrix(Row, .ColIndex("Net")) = val(.TextMatrix(Row, .ColIndex("Value"))) - val(.TextMatrix(Row, .ColIndex("Total")))
            End Select
       End With


End Sub

Private Sub calculation()
    Dim i As Integer
     With Grid
            For i = 2 To .Rows - 1
                        .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("Avalue"))) + val(.TextMatrix(i, .ColIndex("d1"))) + val(.TextMatrix(i, .ColIndex("d2"))) + val(.TextMatrix(i, .ColIndex("d3"))) + val(.TextMatrix(i, .ColIndex("d4"))) + val(.TextMatrix(i, .ColIndex("d5"))) + val(.TextMatrix(i, .ColIndex("d6"))) + val(.TextMatrix(i, .ColIndex("d7"))) + val(.TextMatrix(i, .ColIndex("d8"))) + val(.TextMatrix(i, .ColIndex("d9"))) + val(.TextMatrix(i, .ColIndex("d10"))) + val(.TextMatrix(i, .ColIndex("VacValue"))) + val(.TextMatrix(i, .ColIndex("d11"))) + val(.TextMatrix(i, .ColIndex("d12"))) + val(.TextMatrix(i, .ColIndex("d13"))) + val(.TextMatrix(i, .ColIndex("d14"))) + val(.TextMatrix(i, .ColIndex("d15"))) + val(.TextMatrix(i, .ColIndex("d16"))) + val(.TextMatrix(i, .ColIndex("d17"))) + val(.TextMatrix(i, .ColIndex("d18"))) + val(.TextMatrix(i, .ColIndex("d19"))) + val(.TextMatrix(i, .ColIndex("d20")))
                        .TextMatrix(i, .ColIndex("Net")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Total")))
            Next
       End With



End Sub

Private Sub Text1_Click()
Retrive_Depend
End Sub


Public Sub Retrive_Depend()

Grid.Rows = Grid.FixedRows
Dim str As String
str = " select * from TblExchangeRequest where id =  " & val(Text1.text)
Set Rs_Temp = New ADODB.Recordset
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs_Temp.RecordCount > 0 Then
        dcBranch.BoundText = IIf(IsNull(Rs_Temp("BranchID").value), "", Rs_Temp("BranchID").value)
        DcDur.BoundText = IIf(IsNull(Rs_Temp("DurationID").value), "", Rs_Temp("DurationID").value)
        dcMontth.BoundText = IIf(IsNull(Rs_Temp("Month").value), "", Rs_Temp("Month").value)
   


Dim query  As String
query = "  select * from TblExchangeReques_Detailst  where HID = " & val(Text1.text) & "  and  (TblExchangeReques_Detailst.paid  is null or TblExchangeReques_Detailst.paid  = 0 )  "

'query = query & "  and  InsID  not in  (SELECT dbo.TblVendorReceipt_Details.InsID"
'  query = query & "                                                FROM   dbo.TblVendorReceipt_Details INNER JOIN"
'    query = query & "                                                     dbo.TblVendorReceipt ON dbo.TblVendorReceipt_Details.HID = dbo.TblVendorReceipt.ID"
'    query = query & "                                              where DurationID = " & val(DcDur.BoundText) & "  and Month =  " & val(dcMontth.BoundText) & "  and BranchId =  " & val(dcBranch.BoundText) & "   ) "
   
     
   Dim i As Integer, ss As String
   Set RsTemp = New ADODB.Recordset
   RsTemp.Open query, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
   If RsTemp.RecordCount > 0 Then
        With Grid
        RsTemp.MoveFirst
        Grid.Rows = .FixedRows + RsTemp.RecordCount
        For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Serial")) = i - 1
         
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(RsTemp("InsID").value), "", RsTemp("InsID").value)
                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value)
                .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsTemp("fullcode").value), "", RsTemp("fullcode").value)
                .TextMatrix(i, .ColIndex("cusname")) = IIf(IsNull(RsTemp("cusname").value), "", RsTemp("cusname").value)
                .TextMatrix(i, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp("InsNo").value), "", RsTemp("InsNo").value)
                .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", RsTemp("Value").value)
                
                
               .TextMatrix(i, .ColIndex("CarID")) = IIf(IsNull(RsTemp("carid").value), "", RsTemp("carid").value)
               .TextMatrix(i, .ColIndex("Car")) = IIf(IsNull(RsTemp("boardno").value), "", RsTemp("boardno").value)
               .TextMatrix(i, .ColIndex("DayRate")) = IIf(IsNull(RsTemp("dayvalue").value), "", RsTemp("dayvalue").value)
                       
               
                .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(RsTemp("total_deduct").value), "", RsTemp("total_deduct").value)
                .TextMatrix(i, .ColIndex("Net")) = IIf(IsNull(RsTemp("net").value), "", RsTemp("net").value)
           
                RsTemp.MoveNext
        Next
        End With
   End If
    


 End If








End Sub



Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF3 Then
            Unload FrmSearch_Request
            FrmSearch_Request.SendForm = "VR_ER"
            FrmSearch_Request.show
    End If

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "”‰œ ’—› „ ⁄ÂœÌ‰ "
            Else
                Me.Caption = "Boxes Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            C1Elastic1.Enabled = False
            'C1Elastic2.Enabled = False
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "”‰œ ’—› „ ⁄ÂœÌ‰ ( ÃœÌœ )"
            Else
                Me.Caption = "Exchange Request (New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "”‰œ ’—› „ ⁄ÂœÌ‰( ÃœÌœ )"
            Else
                Me.Caption = "Exchange Request  (New)"
            End If
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            
            C1Elastic1.Enabled = True
            C1Elastic2.Enabled = True
            
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "”‰œ ’—› „ ⁄ÂœÌ‰ (  ⁄œÌ· )"
            Else
                Me.Caption = "Exchange Request (Edit)"
            End If
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            
            C1Elastic1.Enabled = True
            C1Elastic2.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub


Public Sub Retrive(Optional Lngid As Long = 0)

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    
         If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        If Lngid <> 0 Then
            rs.find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
 
    
'    Me.TxtNoteID.text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
'    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)

    
    txtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    txtCode.text = IIf(IsNull(rs("code").value), "", Trim(rs("code").value))
    cbType.ListIndex = IIf(IsNull(rs("ExchangeType").value), "", Trim(rs("ExchangeType").value))
    DcDur.BoundText = IIf(IsNull(rs("DurationID").value), "", Trim(rs("DurationID").value))
  
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Me.Date.value = IIf(IsNull(rs("Date").value), Date, rs("Date").value)
    Me.DateH.value = IIf(IsNull(rs("DateH").value), Date, rs("DateH").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    dcMontth.BoundText = IIf(IsNull(rs("Month").value), "", Trim(rs("Month").value))
    Text1.text = IIf(IsNull(rs("DependID").value), "", rs("DependID").value)
     
   Dim i As Integer, ss As String
   Set RsTemp = New ADODB.Recordset
   
   
   ss = ss & "  select * from TblVendorReceipt_Details  where HID =  " & val(txtID.text) & " order by ID"
   
   
   RsTemp.Open ss, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
   If RsTemp.RecordCount > 0 Then
        With Grid
        RsTemp.MoveFirst
        Grid.Rows = .FixedRows + RsTemp.RecordCount
        For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Serial")) = i - 1
                .TextMatrix(i, .ColIndex("Status")) = 1
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(RsTemp("InsID").value), "", RsTemp("InsID").value)
                 .TextMatrix(i, .ColIndex("IDAC")) = IIf(IsNull(RsTemp("IDAC").value), "", RsTemp("IDAC").value)
                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value)
                .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsTemp("fullcode").value), "", RsTemp("fullcode").value)
                .TextMatrix(i, .ColIndex("cusname")) = IIf(IsNull(RsTemp("cusname").value), "", RsTemp("cusname").value)
                .TextMatrix(i, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp("InsNo").value), "", RsTemp("InsNo").value)
                .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", RsTemp("Value").value)
                .TextMatrix(i, .ColIndex("CarID")) = IIf(IsNull(RsTemp("carid").value), "", RsTemp("carid").value)
                .TextMatrix(i, .ColIndex("Car")) = IIf(IsNull(RsTemp("boardno").value), "", RsTemp("boardno").value)
                .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(RsTemp("total_deduct").value), "", RsTemp("total_deduct").value)
                .TextMatrix(i, .ColIndex("Net")) = IIf(IsNull(RsTemp("net").value), "", RsTemp("net").value)
          
                    
                RsTemp.MoveNext
        Next
        End With
   End If
    
    
    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

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

Function CuurentLogdata(Optional Currentmode As String)
     
 
  
End Function
 
Private Sub SaveData()
    Dim Msg As String
    Dim strSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
    
        If DcDur.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcDur.SetFocus
            Exit Sub
        End If
    
         If cbType.ListIndex = -1 Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· ‰Ê⁄ «·’—› ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            cbType.SetFocus
            Exit Sub
        End If
    
          If dcBranch.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ «·›—⁄ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcBranch.SetFocus
            Exit Sub
        End If
        
        If checkedRow = False Then
                MsgBox ("«Œ — «·œ›⁄«  «Ê·«")
                Exit Sub
        End If
        
        
        Select Case Me.TxtModFlg.text
            Case "N"
                 rs.AddNew
                 txtID.text = CStr(new_id("TblVendorReceipt", "ID", "", True))
            Case "E"
                strSQL = "delete From TblVendorReceipt_Details where  HID =" & val(txtID.text)
                Cn.Execute strSQL, , adExecuteNoRecords
                
                          strSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.text)
        Cn.Execute strSQL, , adExecuteNoRecords



        End Select

        Cn.BeginTrans
        BeginTrans = True
          
        rs("ID").value = val(txtID.text)
        rs("Code").value = Trim(txtCode.text)
        rs("ExchangeType").value = IIf(cbType.ListIndex = -1, Null, cbType.ListIndex)
        rs("DurationID").value = val(DcDur.BoundText)
        rs("DurationName").value = DcDur.text
        rs("Month").value = dcMontth.BoundText
        rs("Date").value = Me.Date.value
        rs("DateH").value = Me.DateH.value
        rs("BranchID").value = dcBranch.BoundText
        rs("DependID").value = IIf(Text1.text = "", Null, val(Text1.text))
        
        rs.update
        
        
       Set RsTemp = New ADODB.Recordset
       RsTemp.Open "TblVendorReceipt_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       Dim i As Integer
       With Grid
   
            For i = .FixedRows To .Rows - 1
               If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked And .TextMatrix(i, .ColIndex("ID")) <> "" Then
                        RsTemp.AddNew
                        RsTemp("ID").value = CStr(new_id("TblVendorReceipt_Details", "ID", "", True))
                         RsTemp("IDAC").value = IIf(.TextMatrix(i, .ColIndex("IDAC")) = "", Null, .TextMatrix(i, .ColIndex("IDAC")))
                        RsTemp("HID").value = val(txtID.text)
                        RsTemp("CusID").value = .TextMatrix(i, .ColIndex("CusID"))
                        RsTemp("InsID").value = .TextMatrix(i, .ColIndex("ID"))
                        RsTemp("fullcode").value = .TextMatrix(i, .ColIndex("fullcode"))
                        RsTemp("cusname").value = .TextMatrix(i, .ColIndex("cusname"))
                       ' RsTemp("InsNo").value = .TextMatrix(i, .ColIndex("InstallmentNo"))
                        RsTemp("Value").value = .TextMatrix(i, .ColIndex("Value"))
                        
                        
                        RsTemp("carid").value = IIf(.TextMatrix(i, .ColIndex("CarID")) = "", Null, .TextMatrix(i, .ColIndex("CarID")))
                        RsTemp("boardno").value = .TextMatrix(i, .ColIndex("Car"))
                                               
                         RsTemp("Total_deduct").value = .TextMatrix(i, .ColIndex("Total"))
                         RsTemp("Net").value = .TextMatrix(i, .ColIndex("Net"))
                        
                       
                        
                        '  RsTemp("Account_Code").value = IIf(.TextMatrix(i, .ColIndex("Account_Code")) = "", "", .TextMatrix(i, .ColIndex("Account_Code")))
                        ' RsTemp("Account_Serial").value = IIf(.TextMatrix(i, .ColIndex("Account_Serial")) = "", "", .TextMatrix(i, .ColIndex("Account_Serial")))
                        
                         Set RsTemp4 = New ADODB.Recordset
                         RsTemp4.Open " select * from TblExchangeReques_Detailst where InsID =  " & val(.TextMatrix(i, .ColIndex("ID"))), Cn, adOpenStatic, adLockOptimistic, adCmdText
                         If RsTemp4.RecordCount > 0 Then
                                        RsTemp4("Paid") = 1
                                        RsTemp4.update
                         End If
                       
                        
                        RsTemp.update
                End If
            Next
        End With
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        createVoucher
        
       'CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ «·»Ì«‰«    " & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & Chr(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "ID='" & val(txtID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Function checkedRow() As Boolean
Dim i As Integer, check As Boolean
For i = 1 To Grid.Rows - 1
        If Grid.TextMatrix(i, Grid.ColIndex("status")) <> "" Then
                If Grid.TextMatrix(i, Grid.ColIndex("status")) <> 0 Then
                        checkedRow = True
                        Exit Function
                End If
        End If
Next
checkedRow = False
End Function


Private Sub Del_Company()
    Dim Msg As String
    Dim strSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If txtID.text <> "" Then

    
        Msg = "”Ì „ Õ–› »Ì«‰«   —ﬁ„ " & Chr(13)
        Msg = Msg + (txtID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
    
            If Not rs.RecordCount < 1 Then
                 
               strSQL = "  update  TblExchangeReques_Detailst set paid = 0 "
               strSQL = strSQL & "     where  TblExchangeReques_Detailst.InsID  in ( select InsID from TblVendorReceipt_Details where hid = " & val(txtID.text) & "  )"
               Cn.Execute strSQL, , adExecuteNoRecords
               
               
                strSQL = "delete From TblVendorReceipt_Details where  HID =" & val(txtID.text)
                Cn.Execute strSQL, , adExecuteNoRecords
            
                strSQL = "delete From TblVendorReceipt  where  ID =" & val(txtID.text)
                Cn.Execute strSQL, , adExecuteNoRecords
                
                strSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.text)
                Cn.Execute strSQL, , adExecuteNoRecords
                Grid.Rows = Grid.FixedRows

                
                   strSQL = "SELECT  *  From TblVendorReceipt "
                   rs.Close
                   rs.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                    Grid.Rows = Grid.FixedRows
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
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·Œ“‰… "
    Msg = Msg & Chr(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  Œ“‰… ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·Œ“‰… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
    '    .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtBoxName_GotFocus()

    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtBoxNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub


