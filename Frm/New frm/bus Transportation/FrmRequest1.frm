VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRequest1 
   BackColor       =   &H00E2E9E9&
   Caption         =   "«À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ··„ ⁄ÂœÌ‰  "
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   14685
   Icon            =   "FrmRequest1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   14685
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic Main_CLE 
      Height          =   9360
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14688
      _cx             =   25903
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   7836
         Width           =   14664
         _cx             =   25876
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
         Begin VB.CommandButton Command2 
            Caption         =   " ’œÌ—«·Ï «·«ﬂ”Ì·"
            Height          =   375
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   360
            Width           =   1488
         End
         Begin VB.Frame Frame9 
            Caption         =   "»Ì«‰«  „Õ«”»Ì…"
            Height          =   735
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   0
            Width           =   6975
            Begin VB.TextBox TxtNoteID 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   120
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   240
               Width           =   2415
            End
            Begin VB.CommandButton Command9 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   375
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton Command8 
               Caption         =   "ﬂ‘› Õ”«»"
               Height          =   375
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   240
               Visible         =   0   'False
               Width           =   1215
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
               TabIndex        =   46
               Top             =   240
               Width           =   990
            End
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   396
            Index           =   0
            Left            =   13116
            TabIndex        =   8
            Top             =   792
            Width           =   1320
            _ExtentX        =   2328
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
            Left            =   11724
            TabIndex        =   9
            Top             =   792
            Width           =   1356
            _ExtentX        =   2381
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
            Left            =   10428
            TabIndex        =   10
            Top             =   792
            Width           =   1272
            _ExtentX        =   2249
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
            Left            =   9060
            TabIndex        =   11
            Top             =   792
            Width           =   1272
            _ExtentX        =   2249
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
            Left            =   7488
            TabIndex        =   12
            Top             =   792
            Width           =   1512
            _ExtentX        =   2672
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
            Left            =   1956
            TabIndex        =   14
            Top             =   792
            Width           =   1356
            _ExtentX        =   2381
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
            Left            =   600
            TabIndex        =   15
            Top             =   792
            Width           =   1308
            _ExtentX        =   2302
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
            Left            =   4680
            TabIndex        =   13
            Top             =   792
            Width           =   1284
            _ExtentX        =   2275
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   396
            Index           =   5
            Left            =   3360
            TabIndex        =   47
            Top             =   792
            Width           =   1284
            _ExtentX        =   2275
            _ExtentY        =   688
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
            Index           =   8
            Left            =   6000
            TabIndex        =   48
            Top             =   792
            Width           =   1488
            _ExtentX        =   2619
            _ExtentY        =   688
            ButtonPositionImage=   1
            Caption         =   "≈·€«¡ «·œ›⁄« "
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
            Left            =   3852
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   840
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   324
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   252
            Width           =   864
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   324
            Index           =   2
            Left            =   4752
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   252
            Width           =   1272
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   324
            Index           =   4
            Left            =   1164
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   252
            Width           =   1164
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   5952
         Left            =   0
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1716
         Width           =   14664
         _cx             =   25876
         _cy             =   10504
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
         Begin VB.TextBox total 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   11280
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   5520
            Width           =   2220
         End
         Begin VB.TextBox TxtFATYou 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Left            =   8040
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   5535
            Width           =   1860
         End
         Begin VB.TextBox TxtFATValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Left            =   4440
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   5535
            Width           =   2460
         End
         Begin VB.TextBox TxtTotalValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   5535
            Width           =   2820
         End
         Begin VB.CheckBox chkChooseAll 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Œ Ì«— «·ﬂ·"
            Height          =   372
            Left            =   13080
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   120
            Width           =   1092
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   4815
            Left            =   0
            TabIndex        =   6
            Top             =   600
            Width           =   14730
            _cx             =   25982
            _cy             =   8493
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
            Cols            =   38
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRequest1.frx":038A
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
            Left            =   3480
            TabIndex        =   40
            Top             =   120
            Visible         =   0   'False
            Width           =   8412
            _ExtentX        =   14843
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "’«›Ì «·ﬁÌ„…"
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   13260
            TabIndex        =   57
            Top             =   5535
            Width           =   1530
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»…«·›« "
            ForeColor       =   &H00C00000&
            Height          =   300
            Index           =   66
            Left            =   10245
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   5535
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ì"
            ForeColor       =   &H00C00000&
            Height          =   300
            Index           =   68
            Left            =   3405
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   5535
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·›« "
            ForeColor       =   &H00C00000&
            Height          =   300
            Index           =   67
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   5535
            Width           =   930
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   744
         Left            =   0
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   864
         Width           =   14664
         _cx             =   25876
         _cy             =   1296
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
         Begin VB.CommandButton Command1 
            Caption         =   "⁄—÷"
            Height          =   492
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   120
            Width           =   1092
         End
         Begin VB.ComboBox cbType 
            Height          =   288
            Left            =   6768
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   864
            Visible         =   0   'False
            Width           =   1224
         End
         Begin VB.TextBox txtCode 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   9240
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   864
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   288
            Left            =   12612
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   264
            Width           =   1200
         End
         Begin MSDataListLib.DataCombo DcDur 
            Height          =   288
            Left            =   4308
            TabIndex        =   4
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   264
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMontth 
            Height          =   288
            Left            =   2040
            TabIndex        =   5
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   264
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker Date 
            Height          =   312
            Left            =   10548
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   240
            Width           =   1188
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal DateH 
            Height          =   312
            Left            =   9480
            TabIndex        =   35
            Top             =   240
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   288
            Left            =   7080
            TabIndex        =   37
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   240
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   264
            Index           =   5
            Left            =   8616
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   11760
            TabIndex        =   36
            Top             =   240
            Width           =   684
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   264
            Index           =   1
            Left            =   3456
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   264
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   264
            Index           =   3
            Left            =   6036
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   264
            Width           =   864
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·ÌœÊÏ"
            Height          =   300
            Index           =   9
            Left            =   10116
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1224
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   300
            Index           =   8
            Left            =   13284
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   264
            Width           =   1224
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’—›"
            Height          =   360
            Index           =   0
            Left            =   8232
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   864
            Visible         =   0   'False
            Width           =   804
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   732
         Left            =   0
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   14724
         _cx             =   25982
         _cy             =   1296
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
         BackColor       =   16777152
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "    «À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ··„ ⁄ÂœÌ‰   "
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
            ButtonImage     =   "FrmRequest1.frx":0923
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
            ButtonImage     =   "FrmRequest1.frx":0CBD
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
            ButtonImage     =   "FrmRequest1.frx":1057
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
            ButtonImage     =   "FrmRequest1.frx":13F1
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
Attribute VB_Name = "FrmRequest1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim RsTemp2 As ADODB.Recordset
Dim RsTemp3 As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim TTP As clstooltip
Dim Account_Code_dynamic As String
Dim Account_Code_dynamic1 As String


Private Sub chkChooseAll_Click()

If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
            Else
                  Exit Sub
  End If
            
Dim i As Integer

For i = 1 To Grid.Rows - 1
    If Grid.TextMatrix(i, Grid.ColIndex("Due_Date")) <> "" Then
            If chkChooseAll.value = 1 Then
                    Grid.TextMatrix(i, Grid.ColIndex("Status")) = 1
            Else
                    Grid.TextMatrix(i, Grid.ColIndex("Status")) = 0
            End If
    End If
Next
ClculteVAT
Relain
End Sub

Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
            txtID.Text = CStr(new_id("TblExchangeRequest2", "ID", "", True))
          '  TXTid.SetFocus
             Grid.Rows = Grid.FixedRows
        Case 1

                                             If ChekClodePeriod(Me.Date.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "E"

        Case 2
                                             If ChekClodePeriod(Me.Date.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
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
                                             If ChekClodePeriod(Me.Date.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Company

        Case 5
                    Unload FrmSearch_Request
                      FrmSearch_Request.SendForm = "VR_VREQ"
                      FrmSearch_Request.show
                    
        Case 6
            Unload Me
            
         Case 7
         print_report
         Case 8
         
               Dim Msg As String
               If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " ”Ì „ Õ–› «·„œ›Ê⁄«  ··”‰œ ° Â·  —Ìœ «” ﬂ„«· ⁄„·Ì… «·Õ–› ø "
               Else
                        Msg = " This action will delete Paid for this receipt "
               End If
               If MsgBox(Msg, vbOKCancel) = vbOK Then
                    Cancel_Paid
                    Command1_Click
                    C1Elastic1.Enabled = True
                    C1Elastic2.Enabled = True
                End If
         
   '      print_report2
    End Select

    Exit Sub
ErrTrap:
End Sub


Private Sub Cancel_Paid()

Dim str As String, AllID As String, i As Integer
If rs.RecordCount > 0 Then
        AllID = IIf(IsNull(rs("AllID").value), "", rs("AllID").value)
        rs("AllID") = Null
        rs.update
End If

If AllID = "" Then
    Exit Sub
End If


str = " select * from TblMinistryContract_Installment where id in (  " & AllID & "  )"
Set Rs_Temp = New ADODB.Recordset
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs_Temp.RecordCount > 0 Then
        For i = 0 To Rs_Temp.RecordCount - 1
                Rs_Temp("VR_Paid") = Null
                Rs_Temp("VRID") = Null
                Rs_Temp.update
                Rs_Temp.MoveNext
        Next
End If

End Sub

Private Sub CmdAttach_Click()
            On Error Resume Next
'ShowAttachments XPTxtBoxID, "0701201405"
     On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments txtID, "15062020003"


End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
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
    Msg = EleHeader.Caption & " —ﬁ„ " & txtID & " » «—ÌŒ" & Date
 
    notes_id = general_noteid

  
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    'C???? C??I?? C?C?C?CE
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim ProjectID As Integer
    
    BranchID = 1
    
    With Grid


line_no = 1
        For i = .FixedRows To .Rows - 1
    BranchID = val(dcBranch.BoundText)
    
            If .TextMatrix(i, .ColIndex("Value")) > 0 And .TextMatrix(i, .ColIndex("Account_Code")) <> "" And .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then   'C?C??? C???E??E IC??
                'Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") 'C?C??? C???E??E
                StrAccountCode = .TextMatrix(i, .ColIndex("Account_Code"))
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("Value")), 0, Msg & "  ··œ›⁄Â  " & .TextMatrix(i, .ColIndex("InstallmentNo")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
                ''/////////
                If val(.TextMatrix(i, .ColIndex("FATValue"))) > 0 And .TextMatrix(i, .ColIndex("AccountCodeVat")) <> "" Then
                If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCodeVat")), .TextMatrix(i, .ColIndex("FATValue")), 0, Msg & "  ··œ›⁄Â  " & .TextMatrix(i, .ColIndex("InstallmentNo")) & "Õ”«» «·ﬁÌ„… «·„÷«›…", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
                End If
                
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, val(.TextMatrix(i, .ColIndex("Value"))) + val(.TextMatrix(i, .ColIndex("FATValue"))), 1, Msg & "  ··⁄ﬁœ  " & "  ··œ›⁄Â  " & .TextMatrix(i, .ColIndex("InstallmentNo")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
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
des = EleHeader.Caption & " —ﬁ„ " & txtID & " » «—ÌŒ" & Date
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim sql As String
tablename = "TblExchangeRequest2"
Filedname = "ID"
NoteSerial1 = val(txtID)
Notevalue = 0

 notytype = 8068
'Notevalue = val(total)
 

 BranchID = val(dcBranch.BoundText)
NoteDate = Me.Date.value
 
'If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TXTNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial
                                     Else
                                                 If TXTNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TXTNoteID.Text = NoteID
                                                                TxtNoteSerial.Text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TXTNoteID.Text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TXTNoteID.Text), BranchID, user_id, NoteDate
rs.Resync adAffectCurrent
 

'     End If

End Function

Private Sub Command2_Click()
    On Error Resume Next
    Dim StrFileName As String
    StrFileName = App.path & "\Report1.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
'Grid.RightToLeft = True
  '  Me.Grid.saveGrid StrFileName, flexFileExcel, True
    Me.Grid.saveGrid StrFileName, flexFileCustomText, True
    OpenFile StrFileName
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub Date_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
TxtNoteSerial.Text = ""
End If

Me.dateH.value = ToHijriDate(Me.Date.value)

End Sub

Private Sub DateH_LostFocus()
VBA.Calendar = vbCalGreg
Me.Date.value = ToGregorianDate(Me.dateH.value)
End Sub

Private Sub Dcbranch_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
TxtNoteSerial.Text = ""
End If
End Sub

Private Sub Dcbranch_Click(Area As Integer)
Dcbranch_Change
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

 
Private Sub Command1_Click()
ProgressBar1.Visible = True
'If check_reg = True Then
' Exit Sub
''End If

Fill_Grid
ClculteVAT
ProgressBar1.Visible = False
ProgressBar1.value = 0
End Sub


Private Function check_reg() As Boolean
'
'Dim str As String
'str = " select * from TblExchangeRequest2  where durationid = " & val(DcDur.BoundText) & "  and Month =   " & val(dcMontth.BoundText)
'Set RsTemp = New ADODB.Recordset
'RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
'If RsTemp.RecordCount > 0 Then
'        MsgBox (" „ «À»«  «·«” Õﬁ«ﬁ ·Â–Â «·› —… ")
'        check_reg = True
'Else
'        check_reg = False
'End If







End Function

Private Sub DcDur_Change()
Dim i As Integer, j As Integer, str As String
    i = val(dcDur.BoundText)
    
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
 i = val(dcDur.BoundText)
 Grid.Rows = Grid.FixedRows
    
  str = str & "       SELECT  TblAttributionContract.IDAC  , dbo.TblMinistryContract_Installment.InstallmentNo, dbo.TblMinistryContract_Installment.Value, dbo.TblAttributionContract.DurationID,"
  str = str & "                     dbo.TblDurations.Name AS DurationName, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblMinistryContract_Installment.Type,"
  str = str & "                    dbo.TblMinistryContract_Installment.Due_DateH, dbo.TblMinistryContract_Installment.Due_Date, dbo.TblCustemers.Fullcode, dbo.TblMinistryContract_Installment.ID,"
  str = str & "                     dbo.TblCustemers.CusID, dbo.TblMinistryContract_Installment.MonthID, dbo.TblMinistryContract_Installment.IDMC, dbo.TblAttributionContract.StartContractDate,"
  str = str & "                    dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.FromDate, dbo.TblDurations.FromDate AS DurFromDate, dbo.TblDurations.ToDate AS DurToDate,"
  str = str & "                      dbo.ACCOUNTS.Account_Code , dbo.ACCOUNTS.account_serial  , dbo.TblCustemers.IBAN, dbo.TblCustemers.BankAccount"
  str = str & "     FROM     dbo.TblAttributionContract INNER JOIN"
  str = str & "                   dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID INNER JOIN"
  str = str & "                   dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
  str = str & "                dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID LEFT OUTER JOIN"
  str = str & "              dbo.TblMinistryContract_Installment ON dbo.TblAttributionContract.IDAC = dbo.TblMinistryContract_Installment.IDMC"
    
    
    If dcMontth.BoundText <> "" Then
              str = str & "      WHERE  (dbo.TblMinistryContract_Installment.Type = 2)  and TblAttributionContract.DurationID = " & i & "  and  TblMinistryContract_Installment.MonthID = " & val(dcMontth.BoundText)
    Else
              str = str & "      WHERE  (dbo.TblMinistryContract_Installment.Type = 2)  and TblAttributionContract.DurationID = " & i
    End If
    
    
   If dcBranch.BoundText <> "" Then
            str = str & " and   TblAttributionContract.BranchID  = " & val(dcBranch.BoundText)
   End If
    
   str = str & "  and  ( TblMinistryContract_Installment.VRID is null or TblMinistryContract_Installment.VR_Paid  = 0 )  "
   
    str = str & " order by TblAttributionContract.IDAC "
    
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    
   Dim V As Integer, H As Integer, WD, str2 As String
    
   Dim DID As Integer, IDMC As Integer
    Dim tot As Integer, daycount As Integer
    If RsTemp.RecordCount > 0 Then
    RsTemp.MoveFirst
            With Grid
             Grid.Rows = Grid.FixedRows + RsTemp.RecordCount
             ProgressBar1.Max = RsTemp.RecordCount
             
             For j = Grid.FixedRows To Grid.Rows - 1
             
                '   If Registered_Before(IIf(IsNull(RsTemp("ID").value), 0, RsTemp("ID").value)) = True Then
                '            .TextMatrix(j, .ColIndex("status")) = 1
                '            Else
                '             .TextMatrix(j, .ColIndex("status")) = 0
                '    End If
                    ProgressBar1.value = j - 2
                    .TextMatrix(j, .ColIndex("Serial")) = j - 1
                     .TextMatrix(j, .ColIndex("IDAC")) = IIf(IsNull(RsTemp("IDAC").value), Null, RsTemp("IDAC").value)
                    .TextMatrix(j, .ColIndex("fullcode")) = IIf(IsNull(RsTemp("Fullcode").value), "", RsTemp("Fullcode").value)
                    .TextMatrix(j, .ColIndex("cusname")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
                    .TextMatrix(j, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", RsTemp("Value").value)
                    .TextMatrix(j, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp("InstallmentNo").value), "", RsTemp("InstallmentNo").value)
                    .TextMatrix(j, .ColIndex("ID")) = IIf(IsNull(RsTemp("ID").value), "", RsTemp("ID").value)
                    .TextMatrix(j, .ColIndex("CusID")) = IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value)
                    .TextMatrix(j, .ColIndex("MonthID")) = IIf(IsNull(RsTemp("MonthID").value), "", RsTemp("MonthID").value)
                    .TextMatrix(j, .ColIndex("StartContractDate")) = IIf(IsNull(RsTemp("StartContractDate").value), "", RsTemp("StartContractDate").value)
                    .TextMatrix(j, .ColIndex("EndContractDate")) = IIf(IsNull(RsTemp("EndContractDate").value), "", RsTemp("EndContractDate").value)
                    .TextMatrix(j, .ColIndex("FromDate")) = IIf(IsNull(RsTemp("FromDate").value), "", RsTemp("FromDate").value)
                    
                    .TextMatrix(j, .ColIndex("FromDate")) = IIf(IsNull(RsTemp("FromDate").value), "", RsTemp("FromDate").value)
                    .TextMatrix(j, .ColIndex("Due_Date")) = IIf(IsNull(RsTemp("Due_Date").value), "", RsTemp("Due_Date").value)
                    .TextMatrix(j, .ColIndex("Due_DateH")) = IIf(IsNull(RsTemp("Due_DateH").value), "", RsTemp("Due_DateH").value)
                    .TextMatrix(j, .ColIndex("Account_Serial")) = IIf(IsNull(RsTemp("Account_Serial").value), "", RsTemp("Account_Serial").value)
                    .TextMatrix(j, .ColIndex("Account_Code")) = IIf(IsNull(RsTemp("Account_Code").value), "", RsTemp("Account_Code").value)
                       
                     If Not (IsNull(RsTemp("IDMC").value) Or IsNull(RsTemp("MonthID").value)) Then
                           ' V = GetVac((rstemp("IDMC").value), (rstemp("MonthID").value))
                            GetVac (RsTemp("IDMC").value), (RsTemp("MonthID").value), tot, daycount
                            H = GetHold((RsTemp("MonthID").value))
                            GetDeducts RsTemp("IDMC").value, dcDur.BoundText, RsTemp("MonthID").value, j
                     End If
                     
                    If Not IsNull(RsTemp("MonthID").value) Then
                            WD = GetMonthDays(RsTemp("MonthID").value)
                    End If
                    .TextMatrix(j, .ColIndex("VacDay")) = daycount   ' V + H
                    .TextMatrix(j, .ColIndex("WorkDay")) = WD - H
                     
                     If Not (IsNull(RsTemp("IDMC").value)) Then
                           .TextMatrix(j, .ColIndex("VacValue")) = tot  'Math.Round(GetDayRate(rstemp("IDMC").value, rstemp("DurFromDate").value, rstemp("DurToDate").value) * val(.TextMatrix(j, .ColIndex("VacDay"))), 2)
                     End If
                     
                     
                    IDMC = IIf(IsNull(RsTemp("IDMC").value), 0, RsTemp("IDMC").value)
                    DID = IIf(IsNull(RsTemp("DurationID").value), 0, RsTemp("DurationID").value)
                     
                     Dim n As String, value As Double
                     
                     str2 = "  select c.* , t.id , t.Name  from TblConfirmViolation c ,  TblViolationTypes t   where   c.Violationid = t.id  and c.MinistryContractID = " & IDMC & " and c.MonthID =  " & .TextMatrix(j, .ColIndex("MonthID")) & " and   c.DurationID =  " & DID
                     Set RsTemp2 = New ADODB.Recordset
                     RsTemp2.Open str2, Cn, adOpenStatic, adLockOptimistic, adCmdText
                     If RsTemp2.RecordCount > 0 Then
                            
                            n = IIf(IsNull(RsTemp2("Name").value), "", RsTemp2("Name").value)
                            value = IIf(IsNull(RsTemp2("value").value), "", RsTemp2("value").value)
                     End If
                           
                            
                    RsTemp.MoveNext
             Next
            End With
    End If
calculation



End Sub

Private Function Registered_Before(ID As Integer) As Boolean
 Dim str As String
 str = "  select * from TblExchangeReques_Detailst2  where insid =  " & ID
 Set Rs_Temp = New ADODB.Recordset
 Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs_Temp.RecordCount > 0 Then
        Registered_Before = True
        Exit Function
 End If
Registered_Before = False
End Function


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

Private Sub GetDeducts(IDMC As Integer, dur As Integer, MonthID As Integer, Row As Integer)

         Dim str As String, i As Integer, j As Integer
         str = "SELECT   dbo.TblConfirmViolation.ID, dbo.TblConfirmViolation.DurationID, dbo.TblConfirmViolation.ViolationID, dbo.TblConfirmViolation.MinistryContractID,"
         str = str & " dbo.TblConfirmViolation.Date , dbo.TblConfirmViolation.value, dbo.TblConfirmViolation.monthid, dbo.TblViolationTypes.name"
         str = str & " FROM     dbo.TblConfirmViolation INNER JOIN   dbo.TblViolationTypes ON dbo.TblConfirmViolation.ViolationID = dbo.TblViolationTypes.ID"
         str = str & " where DurationID = " & dur & " and MonthID = " & MonthID & "  and MinistryContractID =  " & IDMC
         
         Set RsTemp2 = New ADODB.Recordset
         RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
            
         With Grid
         If RsTemp2.RecordCount > 0 Then
                For i = 1 To RsTemp2.RecordCount
                        For j = 1 To 10
                                If .TextMatrix(1, .ColIndex("d" & j)) = RsTemp2("Name").value Then
                                        .TextMatrix(Row, .ColIndex("d" & j)) = RsTemp2("Value").value
                                End If
                        Next
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

Private Sub GetVac(IDMC As Integer, MonthID As Integer, ByRef Total As Integer, ByRef daycount As Integer)

        Dim str As String, cunt As Integer, CityID As Integer, DurID As Integer, DayDiff As Integer, j As Integer, i As Integer
        Dim DurationID As Integer, SchoolFileID As Integer
        
        'str = " select CityID , DurationID  from  TblAttributionContract where IDAC = " & IDMC
        str = " select  d.schoolfileid , a.DurationID  from TblAttributionContract a ,  TblVehicleAllocation_Details  d where a.IDAC = d.IDVA   and d.type = 3  and  a.IDAC =  " & IDMC
        str = str & "   group by schoolfileid , DurationID  "
        
        Set RsTemp2 = New ADODB.Recordset
        RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
       For j = 0 To RsTemp2.RecordCount - 1
                
                DurationID = IIf(IsNull(RsTemp2("DurationID").value), 0, RsTemp2("DurationID").value)
                SchoolFileID = IIf(IsNull(RsTemp2("schoolfileid").value), 0, RsTemp2("schoolfileid").value)
                
                str = " select h.DurationID , h.MonthID ,d.SchoolFileID ,sum (d.daycount) daycount , sum (d.dayvalue)  dayvalue , sum ( (d.daycount * d.dayvalue )) Total"
                str = str & " from TblconfirmVacation  h, TblConfirmVacation_Details d "
                str = str & " where  h.ID = d.HID and  DurationID = " & DurationID & "  and  MonthID = " & MonthID & " and  SchoolFileID = " & SchoolFileID
                str = str & " group by  h.DurationID , h.MonthID ,d.SchoolFileID  "
                Set RsTemp3 = New ADODB.Recordset
                RsTemp3.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp3.RecordCount > 0 Then
                    For i = 0 To RsTemp3.RecordCount - 1
                         Total = Total + IIf(IsNull(RsTemp3("Total").value), 0, RsTemp3("Total").value)
                         daycount = daycount + IIf(IsNull(RsTemp3("daycount").value), 0, RsTemp3("daycount").value)
                         RsTemp3.MoveNext
                   Next
                End If
                RsTemp2.MoveNext
       Next
          
          
     
End Sub

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
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
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

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " «À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ··„ ⁄ÂœÌ‰  "
    LogTexte = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
    My_SQL = " Select id , name from  TblDurations "
    fill_combo dcDur, My_SQL
  
   

    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
  
   Dim StrSQL As String
   StrSQL = "SELECT  *  From TblExchangeRequest2 order by ID"
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
        
    Me.TxtModFlg.Text = "R"
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
        .Cell(flexcpText, 0, .ColIndex("d1"), 0, .ColIndex("d10")) = "Õ”„Ì« "
 
    End With



End Sub



Private Sub Intialize_Deducts()
Dim str As String, i As Integer
Set Rs_Temp = New ADODB.Recordset
str = " select * from TblViolationTypes  "
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
Rs_Temp.MoveFirst
If Rs_Temp.RecordCount > 0 Then
    For i = 1 To Rs_Temp.RecordCount
        Grid.TextMatrix(1, Grid.ColIndex("d" & i)) = IIf(IsNull(Rs_Temp("Name").value), "", Rs_Temp("Name").value)
        Rs_Temp.MoveNext
    Next
End If


For i = 1 To 10
     
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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

 
   Lbl(0).Caption = "No."
   Lbl(3).Caption = " Name Ar"
   Lbl(7).Caption = " Name En"
   'Label3.Caption = "City"
   
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
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  «À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ··„ ⁄ÂœÌ‰   "
    LogTexte = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

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
                Case "d1", "d2", "d3", "d4", "d5", "d6", "d7", "d8", "d9", "d10", "Value"
                        .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("d1"))) + val(.TextMatrix(Row, .ColIndex("d2"))) + val(.TextMatrix(Row, .ColIndex("d3"))) + val(.TextMatrix(Row, .ColIndex("d4"))) + val(.TextMatrix(Row, .ColIndex("d5"))) + val(.TextMatrix(Row, .ColIndex("d6"))) + val(.TextMatrix(Row, .ColIndex("d7"))) + val(.TextMatrix(Row, .ColIndex("d8"))) + val(.TextMatrix(Row, .ColIndex("d9"))) + val(.TextMatrix(Row, .ColIndex("d10")))
                        .TextMatrix(Row, .ColIndex("Net")) = val(.TextMatrix(Row, .ColIndex("Value"))) - val(.TextMatrix(Row, .ColIndex("Total")))
            End Select
       End With

ClculteVAT
Relain
End Sub
Sub ClculteVAT()
If Me.TxtModFlg.Text <> "R" Then
Dim Percetage As Double
Dim i As Integer
Dim account As String
With Grid
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("Due_Date")) <> "" Then
PercentgValueAddedAccount_Transec .TextMatrix(i, .ColIndex("Due_Date")), 2, 0, account, Percetage
.TextMatrix(i, .ColIndex("AccountCodeVat")) = account
TxtFATYou.Text = Percetage
.TextMatrix(i, .ColIndex("FATYou")) = Percetage
If val(.TextMatrix(i, .ColIndex("FATYou"))) > 0 Then
.TextMatrix(i, .ColIndex("FATValue")) = (val(.TextMatrix(i, .ColIndex("Value"))) * val(.TextMatrix(i, .ColIndex("FATYou")))) / 100
Else
.TextMatrix(i, .ColIndex("FATValue")) = 0
End If
End If
.TextMatrix(i, .ColIndex("TotalValue")) = val(.TextMatrix(i, .ColIndex("FATValue"))) + val(.TextMatrix(i, .ColIndex("Value")))
Next i
End With
End If
End Sub
Private Sub calculation()
    Dim i As Integer
     With Grid
            For i = 2 To .Rows - 1
                        .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("d1"))) + val(.TextMatrix(i, .ColIndex("d2"))) + val(.TextMatrix(i, .ColIndex("d3"))) + val(.TextMatrix(i, .ColIndex("d4"))) + val(.TextMatrix(i, .ColIndex("d5"))) + val(.TextMatrix(i, .ColIndex("d6"))) + val(.TextMatrix(i, .ColIndex("d7"))) + val(.TextMatrix(i, .ColIndex("d8"))) + val(.TextMatrix(i, .ColIndex("d9"))) + val(.TextMatrix(i, .ColIndex("d10"))) + val(.TextMatrix(i, .ColIndex("VacValue")))
                        .TextMatrix(i, .ColIndex("Net")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Total")))
            Next
       End With



End Sub



Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
Select Case .ColKey(Col)

Case "Status"
            If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
            Else
                    Cancel = True
            End If

Case "IDAC"
        Cancel = True
Case "MA"
         Cancel = True
Case "fullcode"
          Cancel = True

Case "fullcode"
          Cancel = True
          
          
Case "cusname"
          Cancel = True
          
Case "Account_Serial"
          Cancel = True
          
          
Case "BankAccount"
          Cancel = True

Case "IBAN"
          Cancel = True

Case "recordno"
          Cancel = True


Case "Car"
          Cancel = True


Case "DayRate"
          Cancel = True


Case "WorkDay"
          Cancel = True


Case "Value"
          Cancel = True


Case "VacDay"
          Cancel = True


Case "VacValue"
          Cancel = True


Case "AbsenceCount"
          Cancel = True

Case "Avalue"
          Cancel = True


Case "StartContractDate"
          Cancel = True


Case "EndContractDate"
          Cancel = True


Case "FromDate"
          Cancel = True

          
Case "Total"
          Cancel = True
          
          
Case "Net"
          Cancel = True
          
        
          
End Select
End With


End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ··„ ⁄ÂœÌ‰ "
            Else
                Me.Caption = "Boxes Data"
            End If


            Me.Cmd(8).Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        Me.Cmd(5).Enabled = True
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
          '  C1Elastic2.Enabled = False
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ··„ ⁄ÂœÌ‰ ( ÃœÌœ )"
            Else
                Me.Caption = "Exchange Request (New)"
            End If
            Me.Cmd(8).Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ··„ ⁄ÂœÌ‰( ÃœÌœ )"
            Else
                Me.Caption = "Exchange Request  (New)"
            End If
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
             Me.Cmd(5).Enabled = False
             
            C1Elastic1.Enabled = True
            C1Elastic2.Enabled = True
            
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ··„ ⁄ÂœÌ‰ (  ⁄œÌ· )"
            Else
                Me.Caption = "Exchange Request (Edit)"
            End If
            Me.Cmd(8).Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
             Me.Cmd(5).Enabled = False
             
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            
            C1Elastic1.Enabled = False
         '   C1Elastic2.Enabled = False
            
            
    End Select

    Exit Sub
ErrTrap:
End Sub

Sub Relain()
Dim i As Integer
Dim SumVal As Double
Dim SumVAT As Double
SumVal = 0
SumVAT = 0
With Grid
For i = 1 To .Rows - 1
 If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked And .TextMatrix(i, .ColIndex("ID")) <> "" Then
 SumVal = SumVal + val(.TextMatrix(i, .ColIndex("Value")))
 SumVAT = SumVAT + val(.TextMatrix(i, .ColIndex("FATValue")))
 End If
Next i
End With
TxtFATValue.Text = SumVAT
Total.Text = SumVal
TxtTotalValue.Text = SumVal + SumVAT
End Sub
Public Sub Retrive(Optional Lngid As Long = 0, Optional NoteID As Long = 0)

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
      If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        If Lngid <> 0 And NoteID = 0 Then
            rs.find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
        
             If NoteID <> 0 Then
            rs.find "NoteID =" & NoteID, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
        
       '
        
    End If
    
    Grid.Rows = Grid.FixedRows
    
    Me.TXTNoteID.Text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)

    Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)

    txtID.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    txtCode.Text = IIf(IsNull(rs("code").value), "", Trim(rs("code").value))
    'cbType.ListIndex = IIf(IsNull(rs("ExchangeType").value), -1, Trim(rs("ExchangeType").value))
    dcDur.BoundText = IIf(IsNull(rs("DurationID").value), "", Trim(rs("DurationID").value))
    dcMontth.BoundText = IIf(IsNull(rs("Month").value), "", Trim(rs("Month").value))
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    Me.Date.value = IIf(IsNull(rs("Date").value), Date, rs("Date").value)
    Me.dateH.value = IIf(IsNull(rs("Dateh").value), Date, rs("Dateh").value)
    Me.Total.Text = IIf(IsNull(rs("total").value), "", rs("total").value)
    Me.TxtFATYou.Text = IIf(IsNull(rs("FATYou").value), "", rs("FATYou").value)
    Me.TxtFATValue.Text = IIf(IsNull(rs("FATValue").value), "", rs("FATValue").value)
    Me.TxtTotalValue.Text = IIf(IsNull(rs("TotalValue").value), "", rs("TotalValue").value)
    Dim AllID As String
    AllID = IIf(IsNull(rs("AllID").value), "", (rs("AllID").value))
    '***********************AhmedSalim
    AllID = "select id From dbo.TblMinistryContract_Installment WHERE     (VRID = " & txtID.Text & ")"
    '***********************AhmedSalim
    If AllID = "" Then
            Exit Sub
    End If
    
    
    
    Dim i As Integer, j As Integer, str As String
    i = val(dcDur.BoundText)
    Grid.Rows = Grid.FixedRows
    
    str = str & "       SELECT DISTINCT  TblAttributionContract.IDAC  , dbo.TblMinistryContract_Installment.InstallmentNo, dbo.TblMinistryContract_Installment.Value, dbo.TblAttributionContract.DurationID,"
    str = str & "                     dbo.TblDurations.Name AS DurationName, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblMinistryContract_Installment.Type,"
    str = str & "                    dbo.TblMinistryContract_Installment.Due_DateH, dbo.TblMinistryContract_Installment.Due_Date, dbo.TblCustemers.Fullcode, dbo.TblMinistryContract_Installment.ID,"
    str = str & "                     dbo.TblCustemers.CusID, dbo.TblMinistryContract_Installment.MonthID, dbo.TblMinistryContract_Installment.IDMC, dbo.TblAttributionContract.StartContractDate,"
    str = str & "                    dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.FromDate, dbo.TblDurations.FromDate AS DurFromDate, dbo.TblDurations.ToDate AS DurToDate,"
    str = str & "                      dbo.ACCOUNTS.Account_Code , dbo.ACCOUNTS.account_serial  , dbo.TblCustemers.IBAN, dbo.TblCustemers.BankAccount,TblMinistryContract_Installment.FATYou,TblMinistryContract_Installment.FATValue,TblMinistryContract_Installment.TotalValue,TblMinistryContract_Installment.AccountCodeVat"
    str = str & "     FROM     dbo.TblAttributionContract INNER JOIN"
    str = str & "                   dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID INNER JOIN"
    str = str & "                   dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
    str = str & "                dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID LEFT OUTER JOIN"
    str = str & "              dbo.TblMinistryContract_Installment ON dbo.TblAttributionContract.IDAC = dbo.TblMinistryContract_Installment.IDMC"
    
    
    If dcMontth.BoundText <> "" Then
              str = str & "      WHERE  (dbo.TblMinistryContract_Installment.Type = 2)  and TblAttributionContract.DurationID = " & i & "  and  TblMinistryContract_Installment.MonthID = " & val(dcMontth.BoundText)
    Else
              str = str & "      WHERE  (dbo.TblMinistryContract_Installment.Type = 2)  and TblAttributionContract.DurationID = " & i
    End If
    
    
   If dcBranch.BoundText <> "" Then
            str = str & " and   TblAttributionContract.BranchID  = " & val(dcBranch.BoundText)
   End If
    
   str = str & " and  TblMinistryContract_Installment.ID in ( " & AllID & "  ) "
   
   
    str = str & " order by TblAttributionContract.IDAC "
    
    
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    
   Dim V As Integer, H As Integer, WD, str2 As String
    
   Dim DID As Integer, IDMC As Integer
    Dim tot As Integer, daycount As Integer
    
    If RsTemp.RecordCount > 0 Then
         RsTemp.MoveFirst
            With Grid
                 Grid.Rows = Grid.FixedRows + RsTemp.RecordCount
                 ProgressBar1.Max = RsTemp.RecordCount
             
                 For j = Grid.FixedRows To Grid.Rows - 1
                        ProgressBar1.value = j - 2
                        .TextMatrix(j, .ColIndex("Serial")) = j - 1
                        .TextMatrix(j, .ColIndex("Status")) = 1
                        .TextMatrix(j, .ColIndex("IDAC")) = IIf(IsNull(RsTemp("IDAC").value), 0, RsTemp("IDAC").value)
                        .TextMatrix(j, .ColIndex("FATYou")) = IIf(IsNull(RsTemp("FATYou").value), 0, RsTemp("FATYou").value)
                        .TextMatrix(j, .ColIndex("FATValue")) = IIf(IsNull(RsTemp("FATValue").value), 0, RsTemp("FATValue").value)
                        .TextMatrix(j, .ColIndex("TotalValue")) = IIf(IsNull(RsTemp("TotalValue").value), 0, RsTemp("TotalValue").value)
                        .TextMatrix(j, .ColIndex("AccountCodeVat")) = IIf(IsNull(RsTemp("AccountCodeVat").value), "", RsTemp("AccountCodeVat").value)
                        .TextMatrix(j, .ColIndex("fullcode")) = IIf(IsNull(RsTemp("Fullcode").value), "", RsTemp("Fullcode").value)
                        .TextMatrix(j, .ColIndex("cusname")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
                        .TextMatrix(j, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", RsTemp("Value").value)
                        .TextMatrix(j, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp("InstallmentNo").value), "", RsTemp("InstallmentNo").value)
                        .TextMatrix(j, .ColIndex("ID")) = IIf(IsNull(RsTemp("ID").value), "", RsTemp("ID").value)
                        .TextMatrix(j, .ColIndex("CusID")) = IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value)
                        .TextMatrix(j, .ColIndex("MonthID")) = IIf(IsNull(RsTemp("MonthID").value), "", RsTemp("MonthID").value)
                        .TextMatrix(j, .ColIndex("StartContractDate")) = IIf(IsNull(RsTemp("StartContractDate").value), "", RsTemp("StartContractDate").value)
                        .TextMatrix(j, .ColIndex("EndContractDate")) = IIf(IsNull(RsTemp("EndContractDate").value), "", RsTemp("EndContractDate").value)
                        .TextMatrix(j, .ColIndex("FromDate")) = IIf(IsNull(RsTemp("FromDate").value), "", RsTemp("FromDate").value)
                    
                        .TextMatrix(j, .ColIndex("FromDate")) = IIf(IsNull(RsTemp("FromDate").value), "", RsTemp("FromDate").value)
                        .TextMatrix(j, .ColIndex("Due_Date")) = IIf(IsNull(RsTemp("Due_Date").value), "", RsTemp("Due_Date").value)
                        .TextMatrix(j, .ColIndex("Due_DateH")) = IIf(IsNull(RsTemp("Due_DateH").value), "", RsTemp("Due_DateH").value)
                        .TextMatrix(j, .ColIndex("Account_Serial")) = IIf(IsNull(RsTemp("Account_Serial").value), "", RsTemp("Account_Serial").value)
                        .TextMatrix(j, .ColIndex("Account_Code")) = IIf(IsNull(RsTemp("Account_Code").value), "", RsTemp("Account_Code").value)
                                                  
                    RsTemp.MoveNext
             Next
            End With
    End If
calculation
   
    
    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
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

    If Me.TxtModFlg.Text <> "R" Then
    
        If dcDur.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «·”‰… «·œ—«”Ì… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcDur.SetFocus
            Exit Sub
        End If
        
         If dcBranch.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «·›—⁄ «Ê·« ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcBranch.SetFocus
            Exit Sub
        End If
        
         If dcMontth.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ·  «·› —… «Ê·« ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcMontth.SetFocus
            Exit Sub
        End If
        
    
        If checkedRow = False Then
                MsgBox ("«Œ — «·œ›⁄«  «Ê·«")
                Exit Sub
        End If

    
        Select Case Me.TxtModFlg.Text
            Case "N"
                 rs.AddNew
                 txtID.Text = CStr(new_id("TblExchangeRequest2", "ID", "", True))
            Case "E"
             Cancel_Paid

          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
          Cn.Execute StrSQL, , adExecuteNoRecords


        End Select

        Cn.BeginTrans
        BeginTrans = True
          
        rs("ID").value = val(txtID.Text)
        rs("Code").value = Trim(txtCode.Text)
        rs("ExchangeType").value = IIf(cbType.ListIndex = -1, Null, cbType.ListIndex)
        rs("DurationID").value = val(dcDur.BoundText)
        rs("DurationName").value = dcDur.Text
        rs("Month").value = IIf(dcMontth.BoundText = "", Null, dcMontth.BoundText)
        rs("total").value = val(Total.Text)
        rs("FATYou").value = val(TxtFATYou.Text)
        rs("FATValue").value = val(TxtFATValue.Text)
        rs("TotalValue").value = val(TxtTotalValue.Text)
        rs("BranchID").value = IIf(dcBranch.BoundText = "", Null, dcBranch.BoundText)
        rs("Date").value = Me.Date.value
        rs("DateH").value = Me.dateH.value
        
       ' rs.update
        
        
        
       Dim i As Integer, AllID As String
       With Grid
   
            For i = .FixedRows To .Rows - 1
               If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked And .TextMatrix(i, .ColIndex("ID")) <> "" Then
                        ' If i = .FixedRows Then
                        '         AllID = AllID & IIf(.TextMatrix(i, .ColIndex("ID")) = "", " ", .TextMatrix(i, .ColIndex("ID")))
                        ' Else
                                AllID = AllID & IIf(.TextMatrix(i, .ColIndex("ID")) = "", " ", ",  " & .TextMatrix(i, .ColIndex("ID")))
                        ' End If
                End If
            Next
        End With
        
        
        AllID = mId$(AllID, 2)
          
        rs("AllID").value = AllID
        rs.update
        
        
        
       With Grid
            For i = .FixedRows To .Rows - 1
               If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked And .TextMatrix(i, .ColIndex("ID")) <> "" Then
                        Set RsTemp = New ADODB.Recordset
                        Dim m As String
                        m = "  select * from TblMinistryContract_Installment where id =  " & val(.TextMatrix(i, .ColIndex("ID")))
                        RsTemp.Open m, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        If RsTemp.RecordCount > 0 Then
                                RsTemp("VR_Paid").value = 1
                                RsTemp("VRID").value = val(txtID.Text)
                                RsTemp("FATYou").value = val(.TextMatrix(i, .ColIndex("FATYou")))
                                RsTemp("FATValue").value = val(.TextMatrix(i, .ColIndex("FATValue")))
                                RsTemp("TotalValue").value = val(.TextMatrix(i, .ColIndex("TotalValue")))
                                RsTemp("AccountCodeVat").value = .TextMatrix(i, .ColIndex("AccountCodeVat"))
                                RsTemp.update
                        End If
                End If
            Next
        End With
        
        
        
        
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       'CuurentLogdata
createVoucher

        Select Case Me.TxtModFlg.Text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ «·»Ì«‰«    " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
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

        TxtModFlg.Text = "R"
    End If

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

Private Function checkedRow() As Boolean

Dim i As Integer, Check As Boolean
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


Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "ID='" & val(txtID.Text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
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
            
        If txtID.Text <> "" Then

    
        Msg = "”Ì „ Õ–› »Ì«‰«   —ﬁ„ " & CHR(13)
        Msg = Msg + (txtID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
    
            If Not rs.RecordCount < 1 Then
            
            Cancel_Paid
            
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
             
                StrSQL = "delete From TblExchangeReques_Detailst2 where  HID =" & val(txtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            
                StrSQL = "delete From TblExchangeRequest2 where  ID =" & val(txtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                   StrSQL = "SELECT  *  From TblExchangeRequest2 "
                   rs.Close
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                

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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·Œ“‰… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Private Sub AddTip()
   ' Dim Wrap As String
   ' On Error GoTo ErrTrap
   ' Set TTP = New clstooltip
   ' Wrap = Chr(13) + Chr(10)
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  Œ“‰… ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·Œ“‰… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
''        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
''        .DelayTime = 600
'        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
''        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'    '    .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
'    End With
'
    Exit Sub
ErrTrap:
End Sub



Function print_report(Optional NoteSerial As Integer)
    
    On Error Resume Next
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
           
        MySQL = " SELECT     dbo.TblAttributionContract.IDAC, dbo.TblMinistryContract_Installment.InstallmentNo, dbo.TblMinistryContract_Installment.[Value], "
        MySQL = MySQL & "               dbo.TblDurations.Name AS DurationName, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblMinistryContract_Installment.Type,"
        MySQL = MySQL & "               dbo.TblMinistryContract_Installment.Due_DateH, dbo.TblMinistryContract_Installment.Due_Date, dbo.TblCustemers.Fullcode, dbo.TblMinistryContract_Installment.ID,"
        MySQL = MySQL & "               dbo.TblCustemers.CusID, dbo.TblMinistryContract_Installment.MonthID, dbo.TblMinistryContract_Installment.IDMC, dbo.TblAttributionContract.StartContractDate,"
        MySQL = MySQL & "               dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.FromDate, dbo.TblDurations.FromDate AS DurFromDate, dbo.TblDurations.ToDate AS DurToDate,"
        MySQL = MySQL & "                dbo.ACCOUNTS.Account_Code, dbo.ACCOUNTS.Account_Serial, dbo.TblCustemers.IBAN, dbo.TblCustemers.BankAccount, dbo.TblExchangeRequest2.DurationID,"
        MySQL = MySQL & "               dbo.TblExchangeRequest2.[Month] AS MonthName, dbo.TblExchangeRequest2.AllID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
        MySQL = MySQL & "               dbo.TblExchangeRequest2.[Date], dbo.TblExchangeRequest2.DateH, dbo.TblDurations_Details.Name, dbo.TblExchangeRequest2.ID AS RID,"
        MySQL = MySQL & "               dbo.TblMinistryContract_Installment.FATYou, dbo.TblMinistryContract_Installment.FATValue, dbo.TblMinistryContract_Installment.TotalValue,"
        MySQL = MySQL & "               dbo.TblExchangeRequest2.total, dbo.TblExchangeRequest2.FATYou AS HFATYou, dbo.TblExchangeRequest2.FATValue AS HFATValue,"
        MySQL = MySQL & "               dbo.TblExchangeRequest2.TotalValue AS HTotalValue"
        MySQL = MySQL & "   FROM     dbo.TblAttributionContract INNER JOIN"
        MySQL = MySQL & "   dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID INNER JOIN"
        MySQL = MySQL & "   dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
        MySQL = MySQL & "   dbo.TblExchangeRequest2 INNER JOIN"
        MySQL = MySQL & "   dbo.TblDurations ON dbo.TblExchangeRequest2.DurationID = dbo.TblDurations.ID INNER JOIN"
        MySQL = MySQL & "   dbo.TblDurations_Details ON dbo.TblExchangeRequest2.Month = dbo.TblDurations_Details.ID INNER JOIN"
        MySQL = MySQL & "   dbo.TblBranchesData ON dbo.TblExchangeRequest2.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
        MySQL = MySQL & "   dbo.TblMinistryContract_Installment ON dbo.TblExchangeRequest2.ID = dbo.TblMinistryContract_Installment.VRID ON"
        MySQL = MySQL & "   dbo.TblAttributionContract.IDAC = dbo.TblMinistryContract_Installment.IDMC"
        MySQL = MySQL & "   Where (dbo.TblMinistryContract_Installment.Type = 2)"
        

  MySQL = MySQL & "   and  TblExchangeRequest2.ID = " & val(txtID.Text)
  
     
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VendorRequestReceipt.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VendorRequestReceipt.rpt"
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

          If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
   
    End If
    
    If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(2).AddCurrentValue GetRegVATNo(val(dcBranch.BoundText))
    End If
    
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



