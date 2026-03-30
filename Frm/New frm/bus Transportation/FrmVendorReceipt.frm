VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmVendorReceipt 
   BackColor       =   &H00E2E9E9&
   Caption         =   "   ”‰œ ’—› „ ⁄ÂœÌ‰   "
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12810
   Icon            =   "FrmVendorReceipt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   12810
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic Main_CLE 
      Height          =   9360
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12816
      _cx             =   22595
      _cy             =   16510
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   7836
         Width           =   12792
         _cx             =   22569
         _cy             =   2328
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
            Height          =   396
            Index           =   0
            Left            =   10200
            TabIndex        =   9
            Top             =   792
            Width           =   996
            _ExtentX        =   1746
            _ExtentY        =   688
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
            Left            =   8832
            TabIndex        =   10
            Top             =   792
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   688
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
            Left            =   7692
            TabIndex        =   11
            Top             =   792
            Width           =   1116
            _ExtentX        =   1958
            _ExtentY        =   688
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
            Left            =   6204
            TabIndex        =   12
            Top             =   792
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   688
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
            Left            =   5088
            TabIndex        =   13
            Top             =   792
            Width           =   1044
            _ExtentX        =   1852
            _ExtentY        =   688
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
            TabIndex        =   15
            Top             =   792
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   688
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
            Left            =   1464
            TabIndex        =   16
            Top             =   792
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   688
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
            Left            =   4056
            TabIndex        =   14
            Top             =   792
            Width           =   1008
            _ExtentX        =   1773
            _ExtentY        =   688
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
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
            Height          =   324
            Left            =   3312
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   252
            Width           =   696
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   324
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   252
            Width           =   684
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   324
            Index           =   2
            Left            =   4056
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   252
            Width           =   1092
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   324
            Index           =   4
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   252
            Width           =   1032
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   5832
         Left            =   0
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1836
         Width           =   12792
         _cx             =   22569
         _cy             =   10292
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
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   5772
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   12768
            _cx             =   22521
            _cy             =   10181
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
            Rows            =   50
            Cols            =   28
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmVendorReceipt.frx":038A
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   984
         Left            =   0
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   720
         Width           =   12792
         _cx             =   22569
         _cy             =   1746
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
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   288
            Left            =   10392
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   480
            Width           =   1008
         End
         Begin VB.TextBox txtCode 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   8064
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   480
            Width           =   1080
         End
         Begin VB.ComboBox cbType 
            Height          =   288
            Left            =   5700
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   480
            Width           =   1344
         End
         Begin VB.TextBox txtDepend 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   10392
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   132
            Width           =   1008
         End
         Begin MSDataListLib.DataCombo DcDur 
            Height          =   288
            Left            =   2940
            TabIndex        =   5
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   480
            Width           =   1488
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMontth 
            Height          =   288
            Left            =   240
            TabIndex        =   6
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   480
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’—›"
            Height          =   360
            Index           =   0
            Left            =   6804
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   480
            Width           =   972
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   300
            Index           =   8
            Left            =   11796
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   480
            Width           =   924
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·ÌœÊÏ"
            Height          =   300
            Index           =   9
            Left            =   9288
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   480
            Width           =   744
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   264
            Index           =   3
            Left            =   4596
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   480
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   264
            Index           =   1
            Left            =   2076
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   480
            Width           =   528
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï ÿ·» ’—›"
            Height          =   312
            Index           =   12
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   120
            Width           =   1308
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   612
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   12852
         _cx             =   22675
         _cy             =   1085
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
            TabIndex        =   18
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   19
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVendorReceipt.frx":0783
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
            TabIndex        =   20
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVendorReceipt.frx":0B1D
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
            TabIndex        =   21
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVendorReceipt.frx":0EB7
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
            TabIndex        =   22
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVendorReceipt.frx":1251
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
Attribute VB_Name = "FrmVendorReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim RsTemp2 As ADODB.Recordset
Dim TTP As clstooltip

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
            txtcode.SetFocus
             Grid.Rows = Grid.FixedRows
        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"

        Case 2
         
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_Company

        Case 5

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
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

 
End Sub

Private Sub DCEmP_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

 

End Sub

Private Sub Option1_Click()
 
End Sub

Private Sub Option2_Click()
 
End Sub

 
Private Sub dcDur_Click(Area As Integer)
    Dim i As Integer, j As Integer, str As String
    i = val(dcDur.BoundText)
    
    If i > 0 Then
        str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo dcMontth, str
    Else
        str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo dcMontth, str
    End If
    Grid.Rows = Grid.FixedRows
    
    str = "  SELECT dbo.TblMinistryContract_Installment.InstallmentNo, dbo.TblMinistryContract_Installment.Value, dbo.TblAttributionContract.DurationID, "
    str = str & "     dbo.TblDurations.Name AS DurationName, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblMinistryContract_Installment.Type,"
    str = str & "     dbo.TblMinistryContract_Installment.Due_DateH , dbo.TblMinistryContract_Installment.Due_Date, dbo.TblCustemers.fullcode"
    str = str & "     , dbo.TblMinistryContract_Installment.ID,   dbo.TblCustemers.CusID"
    str = str & "      FROM     dbo.TblAttributionContract INNER JOIN"
    str = str & "      dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    str = str & "      dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID LEFT OUTER JOIN"
    str = str & "      dbo.TblMinistryContract_Installment ON dbo.TblAttributionContract.IDAC = dbo.TblMinistryContract_Installment.IDMC"
    str = str & "      WHERE  (dbo.TblMinistryContract_Installment.Type = 2)  and TblAttributionContract.DurationID = " & i
    
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If RsTemp.RecordCount > 0 Then
            With Grid
             Grid.Rows = Grid.FixedRows + RsTemp.RecordCount
             RsTemp.MoveFirst
             For j = Grid.FixedRows To RsTemp.RecordCount
                    .TextMatrix(j, .ColIndex("fullcode")) = IIf(IsNull(RsTemp("Fullcode").value), "", RsTemp("Fullcode").value)
                    .TextMatrix(j, .ColIndex("cusname")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
                    .TextMatrix(j, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", RsTemp("Value").value)
                    .TextMatrix(j, .ColIndex("Net")) = IIf(IsNull(RsTemp("Value").value), "", RsTemp("Value").value)
                    .TextMatrix(j, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp("InstallmentNo").value), "", RsTemp("InstallmentNo").value)
                    .TextMatrix(j, .ColIndex("ID")) = IIf(IsNull(RsTemp("ID").value), "", RsTemp("ID").value)
                    .TextMatrix(j, .ColIndex("CusID")) = IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value)
                    RsTemp.MoveNext
             Next
            End With
    End If
    
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
    'Dcombos.GetEmployees DcboGovernmentID
   ' Dcombos.getCountriesGovernments Me.DcboGovernmentID

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " ÿ·» ’—› „ ⁄ÂœÌ‰  "
    LogTextE = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""

    Dim My_SQL As String
    My_SQL = " Select id , name from  TblDurations "
    fill_combo dcDur, My_SQL
  
   

    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
  
   Dim StrSQL As String
   StrSQL = "SELECT  *  From TblVendorReceipt order by ID"
   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
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

        .MergeCol(.ColIndex("Name")) = True
        .Cell(flexcpText, 0, .ColIndex("Name"), 1, .ColIndex("Name")) = "«·«”„"

        .MergeCol(.ColIndex("PayNo")) = True
        .Cell(flexcpText, 0, .ColIndex("PayNo"), 1, .ColIndex("PayNo")) = "—ﬁ„ «·œ›⁄…"

        .MergeCol(.ColIndex("Value")) = True
        .Cell(flexcpText, 0, .ColIndex("Value"), 1, .ColIndex("Value")) = "«·ﬁÌ„…"

          .MergeCol(.ColIndex("Total")) = True
          .Cell(flexcpText, 0, .ColIndex("total"), 1, .ColIndex("total")) = "«Ã„«·Ï «·„” Õﬁ« "
        
          .MergeCol(.ColIndex("Net")) = True
          .Cell(flexcpText, 0, .ColIndex("Net"), 1, .ColIndex("Net")) = "«·’«›Ï «·„” Õﬁ"


            .Cell(flexcpText, 0, .ColIndex("deduct"), 0, .ColIndex("other")) = "Õ”„Ì« "
 
          

'
'       .MergeCol(.ColIndex("deduct")) = True
'        .MergeCol(.ColIndex("clean")) = True
'       .MergeCol(.ColIndex("late")) = True
'       .MergeCol(.ColIndex("other")) = True
'
'        .Cell(flexcpText, 1, .ColIndex("deduct"), 1, .ColIndex("deduct")) = "€Ì«»"
'        .Cell(flexcpText, 1, .ColIndex("clean"), 1, .ColIndex("clean")) = " √ŒÌ—"
'        .Cell(flexcpText, 1, .ColIndex("late"), 1, .ColIndex("late")) = "‰Ÿ«›…"
''        .Cell(flexcpText, 1, .ColIndex("other"), 1, .ColIndex("other")) = "√Œ—Ï"
'
    End With



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

 
   Lbl(0).Caption = "No."
   Lbl(3).Caption = " Name Ar"
   Lbl(7).Caption = " Name En"
'   Label3.Caption = "City"
   
  Lbl(2).Caption = "Current Record"
  Lbl(4).Caption = "Recors Count"
   
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
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  ÿ·» ’—› „ ⁄ÂœÌ‰   "
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
        
            Case "late", "other", "deduct", "clean", "Value"
                    .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("deduct"))) + val(.TextMatrix(Row, .ColIndex("late"))) + val(.TextMatrix(Row, .ColIndex("clean"))) + val(.TextMatrix(Row, .ColIndex("other")))
                    .TextMatrix(Row, .ColIndex("Net")) = val(.TextMatrix(Row, .ColIndex("Value"))) - val(.TextMatrix(Row, .ColIndex("Total")))
        End Select
          
    End With
End Sub





Private Sub txtDepend_Change()

    If Not IsNumeric(txtDepend.text) Then Exit Sub
    
    Dim str As String
    
    str = " select * from TblExchangeRequest  where id = " & txtDepend.text
    Set RsTemp = New ADODB.Recordset
   RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If RsTemp.RecordCount > 0 Then
   
    txtcode.text = IIf(IsNull(RsTemp("code").value), "", Trim(RsTemp("code").value))
    cbType.ListIndex = IIf(IsNull(RsTemp("ExchangeType").value), "", Trim(RsTemp("ExchangeType").value))
    dcDur.BoundText = IIf(IsNull(RsTemp("DurationID").value), "", Trim(RsTemp("DurationID").value))
    dcMontth.text = IIf(IsNull(RsTemp("Month").value), "", Trim(RsTemp("Month").value))
   
   Dim i As Integer
   Set RsTemp2 = New ADODB.Recordset
   RsTemp2.Open " select * from TblExchangeReques_Detailst where HID =  " & val(RsTemp("ID").value) & " order by ID", Cn, adOpenStatic, adLockOptimistic, adCmdText
   If RsTemp2.RecordCount > 0 Then
        With Grid
        RsTemp2.MoveFirst
        Grid.Rows = .FixedRows + RsTemp2.RecordCount
        For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(RsTemp2("ID").value), "", RsTemp2("ID").value)
                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(RsTemp2("CusID").value), "", RsTemp2("CusID").value)
                .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsTemp2("fullcode").value), "", RsTemp2("fullcode").value)
                .TextMatrix(i, .ColIndex("cusname")) = IIf(IsNull(RsTemp2("cusname").value), "", RsTemp2("cusname").value)
                .TextMatrix(i, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp2("InsNo").value), "", RsTemp2("InsNo").value)
                .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsTemp2("Value").value), "", RsTemp2("Value").value)
               .TextMatrix(i, .ColIndex("d1")) = IIf(IsNull(RsTemp2("d1").value), "", RsTemp2("d1").value)
               .TextMatrix(i, .ColIndex("d2")) = IIf(IsNull(RsTemp2("d2").value), "", RsTemp2("d2").value)
               .TextMatrix(i, .ColIndex("d3")) = IIf(IsNull(RsTemp2("d3").value), "", RsTemp2("d3").value)
               .TextMatrix(i, .ColIndex("d4")) = IIf(IsNull(RsTemp2("d4").value), "", RsTemp2("d4").value)
               .TextMatrix(i, .ColIndex("d5")) = IIf(IsNull(RsTemp2("d5").value), "", RsTemp2("d5").value)
               .TextMatrix(i, .ColIndex("d6")) = IIf(IsNull(RsTemp2("d6").value), "", RsTemp2("d6").value)
               .TextMatrix(i, .ColIndex("d7")) = IIf(IsNull(RsTemp2("d7").value), "", RsTemp2("d7").value)
               .TextMatrix(i, .ColIndex("d8")) = IIf(IsNull(RsTemp2("d8").value), "", RsTemp2("d8").value)
               .TextMatrix(i, .ColIndex("d9")) = IIf(IsNull(RsTemp2("d9").value), "", RsTemp2("d9").value)
               .TextMatrix(i, .ColIndex("d10")) = IIf(IsNull(RsTemp2("d10").value), "", RsTemp2("d10").value)
                .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(RsTemp2("total_deduct").value), "", RsTemp2("total_deduct").value)
                .TextMatrix(i, .ColIndex("Net")) = IIf(IsNull(RsTemp2("net").value), "", RsTemp2("net").value)
                 
               .TextMatrix(i, .ColIndex("WorkDay")) = IIf(IsNull(RsTemp2("wokdays").value), "", RsTemp2("wokdays").value)
               .TextMatrix(i, .ColIndex("VacDay")) = IIf(IsNull(RsTemp2("stopdays").value), "", RsTemp2("stopdays").value)
               .TextMatrix(i, .ColIndex("VacValue")) = IIf(IsNull(RsTemp2("stopvalue").value), "", RsTemp2("stopvalue").value)
                         
                RsTemp2.MoveNext
        Next
        End With
   End If
            
        
   End If
    
    
    
    
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "ÿ·» ’—› „ ⁄ÂœÌ‰ "
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
        
         '   Me.XPTxtBoxID.locked = True
           ' Me.XPTxtBoxName.locked = True
          '  Me.XPMTxtRemark.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            C1Elastic1.Enabled = False
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "ÿ·» ’—› „ ⁄ÂœÌ‰ ( ÃœÌœ )"
            Else
                Me.Caption = "Exchange Request (New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "ÿ·» ’—› „ ⁄ÂœÌ‰( ÃœÌœ )"
            Else
                Me.Caption = "Exchange Request  (New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            '       Me.XPBtnMove(0).Enabled = False
            '       Me.XPBtnMove(1).Enabled = False
            '       Me.XPBtnMove(2).Enabled = False
            '       Me.XPBtnMove(3).Enabled = False
        
            'Me.XPTxtBoxID.locked = True
            'Me.XPTxtBoxName.locked = False
         '   Me.XPMTxtRemark.locked = False
             C1Elastic1.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "ÿ·» ’—› „ ⁄ÂœÌ‰ (  ⁄œÌ· )"
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
           ' Me.XPTxtBoxID.locked = True
           ' Me.XPTxtBoxName.locked = False
       '     Me.XPMTxtRemark.locked = False
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


    txtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    txtcode.text = IIf(IsNull(rs("code").value), "", Trim(rs("code").value))
    cbType.ListIndex = IIf(IsNull(rs("ExchangeType").value), "", Trim(rs("ExchangeType").value))
    dcDur.BoundText = IIf(IsNull(rs("DurationID").value), "", Trim(rs("DurationID").value))
    dcMontth.text = IIf(IsNull(rs("Month").value), "", Trim(rs("Month").value))
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    
   Dim i As Integer
   Set RsTemp = New ADODB.Recordset
   RsTemp.Open " select * from TblVendorReceipt_Details where HID =  " & val(txtID.text) & " order by ID", Cn, adOpenStatic, adLockOptimistic, adCmdText
    
   If RsTemp.RecordCount > 0 Then
        With Grid
        RsTemp.MoveFirst
        Grid.Rows = .FixedRows + RsTemp.RecordCount
        
        For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(RsTemp("ID").value), "", RsTemp("ID").value)
                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value)
                .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsTemp("fullcode").value), "", RsTemp("fullcode").value)
                .TextMatrix(i, .ColIndex("cusname")) = IIf(IsNull(RsTemp("cusname").value), "", RsTemp("cusname").value)
                .TextMatrix(i, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp("InsNo").value), "", RsTemp("InsNo").value)
                .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", RsTemp("Value").value)
               ' .TextMatrix(i, .ColIndex("deduct")) = IIf(IsNull(RsTemp("absence").value), "", RsTemp("absence").value)
               ' .TextMatrix(i, .ColIndex("clean")) = IIf(IsNull(RsTemp("clean").value), "", RsTemp("clean").value)
               ' .TextMatrix(i, .ColIndex("late")) = IIf(IsNull(RsTemp("late").value), "", RsTemp("late").value)
               ' .TextMatrix(i, .ColIndex("other")) = IIf(IsNull(RsTemp("other").value), "", RsTemp("other").value)
                
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
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
    
        If dcDur.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcDur.SetFocus
            Exit Sub
        End If
    
         If cbType.ListIndex = -1 Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· ‰Ê⁄ «·’—› ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcDur.SetFocus
            Exit Sub
        End If
    
        Select Case Me.TxtModFlg.text
            Case "N"
                 rs.AddNew
                 txtID.text = CStr(new_id("TblVendorReceipt", "ID", "", True))
            Case "E"
                

        End Select

        Cn.BeginTrans
        BeginTrans = True
          
        rs("ID").value = val(txtID.text)
        rs("Code").value = Trim(txtcode.text)
        rs("ExchangeType").value = IIf(cbType.ListIndex = -1, Null, cbType.ListIndex)
        rs("DurationID").value = val(dcDur.BoundText)
        rs("DurationName").value = dcDur.text
        rs("Month").value = dcMontth.text
        rs.update
        
        
       Set RsTemp = New ADODB.Recordset
       RsTemp.Open "TblVendorReceipt_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       Dim i As Integer
       With Grid
   
            For i = .FixedRows To .Rows - 1
               If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked And .TextMatrix(i, .ColIndex("ID")) <> "" Then
                        RsTemp.AddNew
                        RsTemp("ID").value = CStr(new_id("TblVendorReceipt_Details", "ID", "", True))
                        RsTemp("HID").value = val(txtID.text)
                        RsTemp("CusID").value = .TextMatrix(i, .ColIndex("CusID"))
                        RsTemp("InsID").value = .TextMatrix(i, .ColIndex("ID"))
                        RsTemp("fullcode").value = .TextMatrix(i, .ColIndex("fullcode"))
                        RsTemp("cusname").value = .TextMatrix(i, .ColIndex("cusname"))
                        RsTemp("InsNo").value = .TextMatrix(i, .ColIndex("InstallmentNo"))
                        RsTemp("Value").value = .TextMatrix(i, .ColIndex("Value"))
                                              
                            
                         RsTemp("d1").value = IIf(.TextMatrix(i, .ColIndex("d1")) = "", 0, .TextMatrix(i, .ColIndex("d1")))
                         RsTemp("d2").value = IIf(.TextMatrix(i, .ColIndex("d2")) = "", 0, .TextMatrix(i, .ColIndex("d2")))
                         RsTemp("d3").value = IIf(.TextMatrix(i, .ColIndex("d3")) = "", 0, .TextMatrix(i, .ColIndex("d3")))
                         RsTemp("d4").value = IIf(.TextMatrix(i, .ColIndex("d4")) = "", 0, .TextMatrix(i, .ColIndex("d4")))
                         RsTemp("d5").value = IIf(.TextMatrix(i, .ColIndex("d5")) = "", 0, .TextMatrix(i, .ColIndex("d5")))
                         RsTemp("d6").value = IIf(.TextMatrix(i, .ColIndex("d6")) = "", 0, .TextMatrix(i, .ColIndex("d6")))
                         RsTemp("d7").value = IIf(.TextMatrix(i, .ColIndex("d7")) = "", 0, .TextMatrix(i, .ColIndex("d7")))
                         RsTemp("d8").value = IIf(.TextMatrix(i, .ColIndex("d8")) = "", 0, .TextMatrix(i, .ColIndex("d8")))
                         RsTemp("d9").value = IIf(.TextMatrix(i, .ColIndex("d9")) = "", 0, .TextMatrix(i, .ColIndex("d9")))
                         RsTemp("d10").value = IIf(.TextMatrix(i, .ColIndex("d10")) = "", 0, .TextMatrix(i, .ColIndex("d10")))
                         RsTemp("Total_deduct").value = .TextMatrix(i, .ColIndex("Total"))
                         RsTemp("Net").value = .TextMatrix(i, .ColIndex("Net"))
                        
                         RsTemp("wokdays").value = IIf(.TextMatrix(i, .ColIndex("WorkDay")) = "", 0, .TextMatrix(i, .ColIndex("WorkDay")))
                         RsTemp("stopdays").value = IIf(.TextMatrix(i, .ColIndex("VacDay")) = "", 0, .TextMatrix(i, .ColIndex("VacDay")))
                         RsTemp("stopvalue").value = IIf(.TextMatrix(i, .ColIndex("VacValue")) = "", 0, .TextMatrix(i, .ColIndex("VacValue")))
                                              
                
             
                        RsTemp.update
                End If
            Next
        End With
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
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

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If txtID.text <> "" Then

    
        Msg = "”Ì „ Õ–› »Ì«‰«  «·”Ã· —ﬁ„ " & Chr(13)
        Msg = Msg + (txtID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
    
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From TblVendorReceipt where  ID =" & val(txtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                   StrSQL = "SELECT  *  From TblVendorReceipt"
                   rs.Close
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                

                If rs.RecordCount < 1 Then
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
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  Œ“‰… ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·Œ“‰… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
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




