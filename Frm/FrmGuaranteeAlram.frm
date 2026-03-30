VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmGuaranteeAlram 
   ClientHeight    =   10905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14625
   Icon            =   "FrmGuaranteeAlram.frx":0000
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10905
   ScaleWidth      =   14625
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   10905
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14625
      _cx             =   25797
      _cy             =   19235
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
      Begin VB.ComboBox CboType 
         Height          =   315
         Left            =   6525
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   435
         Width           =   3465
      End
      Begin VB.TextBox TxtInterval 
         Alignment       =   2  'Center
         Height          =   435
         Left            =   5730
         MaxLength       =   3
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   435
         Width           =   660
      End
      Begin ImpulseButton.ISButton CmdPrint 
         Height          =   435
         Left            =   30
         TabIndex        =   8
         Top             =   435
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
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
         ButtonImage     =   "FrmGuaranteeAlram.frx":038A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   840
         Index           =   1
         Left            =   12105
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   30
         Width           =   2280
         _cx             =   4022
         _cy             =   1482
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
         ForeColor       =   128
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "ÿ—Ìﬁ… «·⁄—÷"
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
         Begin VB.OptionButton OptShowType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï"
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   90
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.OptionButton OptShowType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ›’Ì·Ï"
            Height          =   285
            Index           =   0
            Left            =   1050
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   300
            Value           =   -1  'True
            Width           =   885
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   9045
         Left            =   180
         TabIndex        =   2
         Top             =   1080
         Width           =   14505
         _cx             =   25585
         _cy             =   15954
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmGuaranteeAlram.frx":0724
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
         Height          =   660
         Index           =   0
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   10260
         Width           =   12195
         _cx             =   21511
         _cy             =   1164
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
            Height          =   390
            Left            =   0
            TabIndex        =   4
            Top             =   90
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   688
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
            ButtonImage     =   "FrmGuaranteeAlram.frx":095B
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
         Begin VB.Image Img 
            Height          =   240
            Left            =   7770
            Picture         =   "FrmGuaranteeAlram.frx":0CF5
            Top             =   90
            Width           =   240
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H000000FF&
            Height          =   405
            Index           =   0
            Left            =   2745
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   90
            Width           =   4935
         End
      End
      Begin ImpulseButton.ISButton CmdDo 
         Height          =   390
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   688
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
         ButtonImage     =   "FrmGuaranteeAlram.frx":107F
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   435
         Index           =   4
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   510
         Width           =   3315
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00C00000&
         Height          =   390
         Index           =   3
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   30
         Width           =   3315
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÌÊ„"
         Height          =   435
         Index           =   2
         Left            =   5070
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   435
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Œ — ‰Ê⁄ «·≈” ⁄·«„"
         Height          =   390
         Index           =   1
         Left            =   6405
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   30
         Width           =   3465
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   10905
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   14625
      _cx             =   25797
      _cy             =   19235
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
         Height          =   2175
         Left            =   120
         TabIndex        =   34
         Top             =   585
         Width           =   7515
         Begin VB.ComboBox CboType2 
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   420
            Left            =   1500
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   720
            Width           =   825
         End
         Begin VB.Frame fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ï «·› —…"
            Enabled         =   0   'False
            Height          =   1035
            Index           =   5
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   1080
            Width           =   2415
            Begin MSComCtl2.DTPicker XPDtbBuyFrom 
               Height          =   345
               Left            =   150
               TabIndex        =   74
               Top             =   240
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   166920193
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker XPDtpBuyTo 
               Height          =   345
               Left            =   150
               TabIndex        =   75
               Top             =   630
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   166920193
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   285
               Index           =   14
               Left            =   1710
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   660
               Width           =   465
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   315
               Index           =   15
               Left            =   1710
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   240
               Width           =   465
            End
         End
         Begin VB.ListBox ListAllStore 
            Height          =   1425
            ItemData        =   "FrmGuaranteeAlram.frx":1419
            Left            =   7245
            List            =   "FrmGuaranteeAlram.frx":1420
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   360
            Width           =   2970
         End
         Begin VB.ListBox ListStoreSelect 
            Height          =   1425
            ItemData        =   "FrmGuaranteeAlram.frx":1432
            Left            =   4080
            List            =   "FrmGuaranteeAlram.frx":1439
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   330
            Width           =   2730
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   255
            Left            =   240
            TabIndex        =   81
            Top             =   1800
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   450
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
            ButtonImage     =   "FrmGuaranteeAlram.frx":144E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   1320
            Visible         =   0   'False
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   450
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«—”«· »«·«Ì„Ì·"
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
            ButtonImage     =   "FrmGuaranteeAlram.frx":17E8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌÊ„"
            Height          =   420
            Index           =   8
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   960
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁ»·  «—ÌŒ «·«‰ Â«¡ »"
            Height          =   375
            Index           =   9
            Left            =   2415
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   750
            Width           =   1500
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "<"
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
            Height          =   345
            Left            =   6795
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "<<"
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
            Height          =   360
            Left            =   6795
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   1305
            Width           =   495
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   6795
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   630
            Width           =   495
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   6795
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·„Œ“‰"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   6090
            TabIndex        =   37
            Top             =   120
            Width           =   1470
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Height          =   2175
         Left            =   7725
         TabIndex        =   26
         Top             =   585
         Width           =   6990
         Begin VB.ListBox ListBranchSelected 
            Height          =   1425
            ItemData        =   "FrmGuaranteeAlram.frx":1B82
            Left            =   120
            List            =   "FrmGuaranteeAlram.frx":1B89
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   330
            Width           =   2655
         End
         Begin VB.ListBox ListBranchAll 
            Height          =   1425
            ItemData        =   "FrmGuaranteeAlram.frx":1BA1
            Left            =   3240
            List            =   "FrmGuaranteeAlram.frx":1BA8
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   330
            Width           =   2655
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   750
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "<<"
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
            Height          =   240
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1425
            Width           =   495
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "<"
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
            Height          =   345
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·›—⁄"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   2310
            TabIndex        =   29
            Top             =   120
            Width           =   1440
         End
      End
      Begin XtremeSuiteControls.CheckBox CheckBox1 
         Height          =   315
         Left            =   9615
         TabIndex        =   23
         Top             =   150
         Width           =   4290
         _Version        =   786432
         _ExtentX        =   7567
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "≈ŸÂ«— «·«’‰«› «· Ì ·Ì” ·Â«  «—ÌŒ «‰ Â«¡ "
         BackColor       =   14871017
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   7290
         Left            =   30
         TabIndex        =   17
         Top             =   2775
         Width           =   14655
         _cx             =   25850
         _cy             =   12859
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmGuaranteeAlram.frx":1BBB
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
         Height          =   660
         Index           =   3
         Left            =   30
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   10245
         Width           =   12360
         _cx             =   21802
         _cy             =   1164
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
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   450
            Left            =   0
            TabIndex        =   19
            Top             =   120
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   794
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
            ButtonImage     =   "FrmGuaranteeAlram.frx":1DF7
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
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   450
            Left            =   2580
            TabIndex        =   42
            Top             =   150
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   794
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
            ButtonImage     =   "FrmGuaranteeAlram.frx":2191
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H000000FF&
            Height          =   480
            Index           =   5
            Left            =   4185
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   120
            Width           =   7725
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   12030
            Picture         =   "FrmGuaranteeAlram.frx":252B
            Top             =   120
            Width           =   240
         End
      End
      Begin MSDataListLib.DataCombo DCBranch 
         Height          =   315
         Left            =   5745
         TabIndex        =   24
         Top             =   150
         Visible         =   0   'False
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   390
         Index           =   52
         Left            =   11310
         TabIndex        =   25
         Top             =   180
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00C00000&
         Height          =   390
         Index           =   7
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   30
         Width           =   3300
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   450
         Index           =   6
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   435
         Width           =   3300
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic2 
      Height          =   10905
      Left            =   0
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   0
      Width           =   14625
      _cx             =   25797
      _cy             =   19235
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
      Begin VB.Frame Frame11 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·”‰œ"
         Height          =   870
         Left            =   5955
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   2160
         Width           =   8625
         Begin MSComCtl2.DTPicker FrmDate 
            Height          =   330
            Left            =   3600
            TabIndex        =   68
            Top             =   270
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   143654915
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker TODate 
            Height          =   330
            Left            =   120
            TabIndex        =   69
            Top             =   270
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   143654915
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   13
            Left            =   2550
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   11
            Left            =   5970
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   2175
         Left            =   7590
         TabIndex        =   59
         Top             =   0
         Width           =   7005
         Begin VB.ListBox ListBranchAll2 
            Height          =   1425
            ItemData        =   "FrmGuaranteeAlram.frx":28B5
            Left            =   3240
            List            =   "FrmGuaranteeAlram.frx":28BC
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   330
            Width           =   2655
         End
         Begin VB.ListBox ListBranchSelected2 
            Height          =   1425
            ItemData        =   "FrmGuaranteeAlram.frx":28CF
            Left            =   120
            List            =   "FrmGuaranteeAlram.frx":28D6
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   330
            Width           =   2655
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·›—⁄"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   2310
            TabIndex        =   66
            Top             =   120
            Width           =   1440
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "<"
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
            Height          =   345
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Caption         =   "<<"
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
            Height          =   240
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   1425
            Width           =   495
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   750
            Width           =   495
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   420
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   2175
         Left            =   120
         TabIndex        =   51
         Top             =   0
         Width           =   7515
         Begin VB.ListBox ListStoreSelect2 
            Height          =   1425
            ItemData        =   "FrmGuaranteeAlram.frx":28EE
            Left            =   120
            List            =   "FrmGuaranteeAlram.frx":28F5
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   330
            Width           =   2730
         End
         Begin VB.ListBox ListAllStore2 
            Height          =   1425
            ItemData        =   "FrmGuaranteeAlram.frx":290A
            Left            =   3285
            List            =   "FrmGuaranteeAlram.frx":2911
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   360
            Width           =   2970
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·„Œ“‰"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   2130
            TabIndex        =   58
            Top             =   120
            Width           =   1470
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2835
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   2835
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   630
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "<<"
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
            Height          =   360
            Left            =   2835
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   1305
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "<"
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
            Height          =   345
            Left            =   2835
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   960
            Width           =   495
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
         Height          =   7200
         Left            =   30
         TabIndex        =   44
         Top             =   3015
         Width           =   14505
         _cx             =   25585
         _cy             =   12700
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmGuaranteeAlram.frx":2923
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   660
         Index           =   4
         Left            =   30
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   10260
         Width           =   12195
         _cx             =   21511
         _cy             =   1164
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
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   390
            Left            =   0
            TabIndex        =   46
            Top             =   90
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   688
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
            ButtonImage     =   "FrmGuaranteeAlram.frx":2B34
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
         Begin ImpulseButton.ISButton ISButton7 
            Height          =   360
            Left            =   1395
            TabIndex        =   72
            Top             =   135
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   635
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
            ButtonImage     =   "FrmGuaranteeAlram.frx":2ECE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H000000FF&
            Height          =   405
            Index           =   10
            Left            =   2745
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   90
            Width           =   4935
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   7770
            Picture         =   "FrmGuaranteeAlram.frx":3268
            Top             =   90
            Width           =   240
         End
      End
      Begin ImpulseButton.ISButton ISButton6 
         Height          =   375
         Left            =   1875
         TabIndex        =   48
         Top             =   2475
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
         ButtonImage     =   "FrmGuaranteeAlram.frx":35F2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   " ‰»ÌÂ«  ”‰œ«  «·ÕÃ“"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2220
         TabIndex        =   50
         Top             =   0
         Width           =   12360
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00C00000&
         Height          =   390
         Index           =   12
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   30
         Width           =   3315
      End
   End
End
Attribute VB_Name = "FrmGuaranteeAlram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ind As Integer
Public PrinterSql As String
Public CurrentMessage As String

Private Sub CboType_Change()
    SetControls
End Sub

Private Sub ChangeLang()
XPLbl(52).Caption = "Branch"
Label12.Caption = "Select Store"
    'LblCaption.Caption = Me.Caption
Label13.Caption = "Select Branch"
    Ele(1).Caption = "View Type"
    OptShowType(0).Caption = "Details"
    OptShowType(1).Caption = "Summary"
ISButton3.Caption = "Run"
    CmdDo.Caption = "Run"
    ISButton1.Caption = "Print"
    CmdPrint.Caption = "Print"
    CmdExit.Caption = "Exit"
    lbl(2).Caption = "Days"
    lbl(8).Caption = "Days"
    lbl(9).Caption = "Before the expiry date"
    lbl(1).Caption = "Select Report Type"
    ISButton2.Caption = "Exit"
    CheckBox1.RightToLeft = False
    CheckBox1.Caption = "Display items that do not have an expiry date"
    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("Serial")) = "No"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        
        .TextMatrix(0, .ColIndex("Qty")) = "Current Qty"
        .TextMatrix(0, .ColIndex("ExpiryDate")) = "Expiry Date"
    End With
    
    With Fg
        .TextMatrix(0, .ColIndex("Serial")) = "No"
        .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
 
        .TextMatrix(0, .ColIndex("Qty")) = "Current Qty"
        .TextMatrix(0, .ColIndex("ItemSerial")) = "Item Serial"
        .TextMatrix(0, .ColIndex("TransactionDate")) = "Purchase Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
        .TextMatrix(0, .ColIndex("Quantity")) = "Bill Quantity"
        .TextMatrix(0, .ColIndex("guaranteeTime")) = "guaranteeTime"
        .TextMatrix(0, .ColIndex("ExpireDate")) = "End guarantee"
        .TextMatrix(0, .ColIndex("TransactionSerial")) = "Transaction No"

    End With

End Sub

Private Sub CboType_Click()
    SetControls
End Sub

Private Sub CboType2_Change()
SetControls
End Sub

Private Sub CboType2_Click()
SetControls
End Sub

Private Sub CmdDo_Click()
    GetData
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdPrint_Click()
    Dim cItemReport As ClsItemsReport

    If Fg.Rows = Fg.FixedRows Then
        Exit Sub
    End If

    Set cItemReport = New ClsItemsReport

    If Me.CboType.ListIndex = 0 Then
        cItemReport.ShowGuarantee 0, 0, DetailDisplayType, PrinterTarget
    ElseIf Me.CboType.ListIndex = 1 Then
        cItemReport.ShowGuarantee 1, val(Me.TxtInterval.text), DetailDisplayType, WindowTarget
    ElseIf Me.CboType.ListIndex = 2 Then
        cItemReport.ShowGuarantee 2, val(Me.TxtInterval.text), DetailDisplayType, WindowTarget
    End If

    Set cItemReport = Nothing
End Sub
Function print_report(Optional NoteSerial As String)
On Error GoTo ErrTrap
  Dim SotreID As String
    Dim i As Integer
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
        sql = " SELECT     TOP 100 PERCENT dbo.Transaction_Details.ExpiryDate, SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS Qty, "
        sql = sql & "               dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.Transactions.BranchId,"
        sql = sql & "              dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_namee"
        sql = sql & "  FROM         dbo.Transaction_Details INNER JOIN"
        sql = sql & "              dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
        sql = sql & "              dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
        sql = sql & "              dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
        sql = sql & "              dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
        If CheckBox1.value = vbChecked Then
        sql = sql & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.text) & " or dbo.Transaction_Details.ExpiryDate is null) "
        Else
        sql = sql & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.text) & " ) "
        End If
  
    SotreID = "0"
    For i = 0 To Me.ListStoreSelect.ListCount - 1
    SotreID = SotreID & "," & Me.ListStoreSelect.ItemData(i)
    Next i
     If SotreID <> "0" Then
     sql = sql & " and   dbo.Transactions.StoreID in (" & SotreID & ")"
     End If
        'If val(DCBranch.BoundText) <> 0 And DCBranch.Text <> "" Then
       ' sql = sql & " and dbo.Transactions.BranchId =" & val(DCBranch.BoundText) & ""
       ' End If
        sql = sql & " GROUP BY dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, "
        sql = sql & "              dbo.TblItems.ItemNamee , dbo.TblItems.Fullcode, dbo.Transactions.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
        sql = sql & "  Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) > 0)"
        sql = sql & " ORDER BY dbo.Transaction_Details.ExpiryDate"
    sql = PrinterSql
 
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlarmExperDateOfItems.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlarmExperDateOfItemsE.rpt"
        End If

   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
     If SystemOptions.UserInterface = ArabicInterface Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
       Else
       Msg = "No Date"
     End If
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
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:
  End Function
Function print_reportBooking(Optional NoteSerial As String)
On Error GoTo ErrTrap
  Dim SotreID As String
    Dim i As Integer
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
     sql = " SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.BookingDate, dbo.Transactions.NoteSerial1, "
sql = sql & "                       dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.StoreID, dbo.TblStore.StoreName,"
sql = sql & "                      dbo.TblStore.StoreAdress, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName,"
sql = sql & "                      dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode AS ItemFullcode, dbo.TblItems.barCodeNO"
sql = sql & " FROM         dbo.TblItems RIGHT OUTER JOIN"
sql = sql & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID RIGHT OUTER JOIN"
sql = sql & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
sql = sql & " WHERE     (dbo.Transactions.Transaction_Type = 39) and dbo.Transaction_Details.ShowQty>0 and dbo.Transactions.BookingDate<" & SQLDate(Date, True) & " "
    SotreID = "0"
    For i = 0 To Me.ListStoreSelect2.ListCount - 1
    SotreID = SotreID & "," & Me.ListStoreSelect2.ItemData(i)
    Next i
     If SotreID <> "0" Then
     sql = sql & " and   dbo.Transactions.StoreID in (" & SotreID & ")"
     End If
  If Not IsNull(FrmDate.value) Then
  sql = sql & " and  dbo.Transactions.Transaction_Date >=" & SQLDate(FrmDate.value, True) & ""
  End If
    If Not IsNull(TODate.value) Then
  sql = sql & " and  dbo.Transactions.Transaction_Date <=" & SQLDate(TODate.value, True) & ""
  End If
    
 
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlarmBooking.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlarmBookingE.rpt"
        End If

   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
     If SystemOptions.UserInterface = ArabicInterface Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
       Else
       Msg = "No Date"
     End If
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
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If Not IsNull(FrmDate.value) Then
    xReport.ParameterFields(4).AddCurrentValue FrmDate.value
    End If
     If Not IsNull(TODate.value) Then
    xReport.ParameterFields(5).AddCurrentValue TODate.value
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
ErrTrap:
  End Function

Private Sub Dcbranch_Change()
Dcbranch_Click (0)
End Sub

Private Sub Dcbranch_Click(Area As Integer)
ISButton3_Click
End Sub

Private Sub Fg_MouseUp(Button As Integer, _
                       Shift As Integer, _
                       X As Single, _
                       Y As Single)
    Dim LngCurrentItemID As Long
    Dim LngMouseRow As Long

    If Button = vbRightButton Then

        With Me.Fg
            LngMouseRow = .MouseRow

            If LngMouseRow = -1 Then Exit Sub
            If .Col = -1 Then Exit Sub
            mdifrmmain.MnuItemTools_ItemSerial.Tag = ""
            mdifrmmain.MnuItemTools_ItemCart.Tag = ""
            mdifrmmain.MnuItemTools_ItemData.Tag = ""
            mdifrmmain.MnuItemTools_ItemQty.Tag = ""
        
            If val(.TextMatrix(LngMouseRow, .ColIndex("ItemID"))) <> 0 Then
                LngCurrentItemID = val(.TextMatrix(LngMouseRow, .ColIndex("ItemID")))

                If .TextMatrix(LngMouseRow, .ColIndex("ItemSerial")) <> "" Then
                    mdifrmmain.MnuItemTools_ItemSerial.Enabled = True
                    mdifrmmain.MnuItemTools_ItemSerial.Tag = LngCurrentItemID & "-" & .TextMatrix(LngMouseRow, .ColIndex("ItemSerial"))
                Else
                    mdifrmmain.MnuItemTools_ItemSerial.Enabled = False
                    mdifrmmain.MnuItemTools_ItemSerial.Tag = ""
                End If
            
                mdifrmmain.MnuItemTools_ItemCart.Tag = LngCurrentItemID & "-" & ""
                mdifrmmain.MnuItemTools_ItemQty.Tag = LngCurrentItemID
                mdifrmmain.MnuItemTools_ItemData.Tag = LngCurrentItemID
                Me.PopupMenu mdifrmmain.MnuItemTools
            End If

        End With

    End If

End Sub

Private Sub Form_Activate()
If Ind = 1 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 Me.Caption = " ‰»ÌÂ «‰ Â«¡ «·«’‰«›"
 Else
  Me.Caption = "Items   Expiry Date Alarm"
End If
 ElseIf Ind = 2 Then
If SystemOptions.UserInterface = ArabicInterface Then
 Me.Caption = " ‰»ÌÂ«  ”‰œ«  «·ÕŒÃ“"
Else
Me.Caption = "Booking Alarm"
End If
 Else
If SystemOptions.UserInterface = ArabicInterface Then
 Me.Caption = " ‰»ÌÂ «·√’‰«› «·„ÊÃÊœ… Ê«· Ï ﬁ«—» „œ… ÷„«‰Â« ⁄·Ï «·√‰ Â«¡"
Else
Me.Caption = "Items   Guarantees Alarm"
End If
 End If
End Sub

Private Sub Form_Load()
    Dim Msg As String
    Dim GrdBack As New ClsBackGroundPic
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
    FrmDate.value = Date
    TODate.value = Date
     FrmDate.value = ""
    TODate.value = ""
            With Me.CboType2
            .Clear
            .AddItem " ‰ ÂÌ ﬁ»·   "
            .AddItem "„‰  «—ÌŒ  «·Ì  «—ÌŒ"
             
        End With
        
        
If Ind = 1 Then
FillMylist
C1Elastic2.Visible = False
C1Elastic1.Visible = True
EleMain.Visible = False
ISButton3_Click
ElseIf Ind = 2 Then
FillMylist
FillGridBooking
C1Elastic2.Visible = True
C1Elastic1.Visible = False
EleMain.Visible = False
Else
C1Elastic2.Visible = False
EleMain.Visible = True
C1Elastic1.Visible = False
    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
        .Rows = .FixedRows
    End With


    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "«· «—ÌŒ «·Õ«·Ï : " & Format(Date, "yyyy/M/d")
        lbl(3).Caption = Msg

        With Me.CboType
            .Clear
            .AddItem "«·√’‰«›  «· Ï «‰ Â  „œ… ÷„«‰Â«"
            .AddItem "«·√’‰«›  ”Ê›  ‰ ÂÏ „œ… ÷„«‰Â« ›Ï Œ·«·"
            .AddItem "«·ﬂ·(«· Ï «‰ Â  „œ… ÷„«‰Â« «Ê ›Ï Œ·«·)"
        End With
        
        
        

        Msg = " ‰»ÌÂ :- Â–« «· ‰»ÌÂ ·«Ì⁄„· «Ê Ì”—Ï ”ÊÏ ⁄·Ï «·√’‰«› «· Ï   ⁄«„· »‰Ÿ«„ «·”Ì—Ì«·"
 
    Else
        Msg = "Current Date Is: " & Format(Date, "yyyy/M/d")
        lbl(3).Caption = Msg

        With Me.CboType
            .Clear
            .AddItem "Items that ended guaranteed"
            .AddItem "Items that will end guaranteed During"
            .AddItem "Items ended and will ended guarantee"
        End With

        Msg = "This alarm is working with Items  that have only serial"

    End If

    lbl(0).Caption = Msg
End If
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    Me.Height = 9240
    Me.Width = 11100
    Resize_Form Me

End Sub

Private Sub GetData()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim Msg As String

    If Me.CboType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "»—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·≈” ⁄·«„...!!!"
        Else
            Msg = "Please Select Query Type .!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    ElseIf Me.CboType.ListIndex = 1 Or Me.CboType.ListIndex = 2 Then

        If val(Me.TxtInterval.text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡ ﬂ «»… ﬁÌ„… «·› —…..!!!"
            Else
                Msg = " Please Enter Period .!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtInterval.SetFocus
            Exit Sub
        End If
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "Select * From QryGuaranteeAlram Where"

        If Me.CboType.ListIndex = 0 Then
            StrSQL = StrSQL + " ExpireDate <=#" & SQLDate(Date) & "#"
            Me.lbl(4).Caption = ""
        ElseIf Me.CboType.ListIndex = 1 Then
            StrSQL = StrSQL + " ExpireDate <=#" & SQLDate(DateAdd("d", val(Me.TxtInterval.text), Date)) & "# And ExpireDate >=#" & SQLDate(Date) & "#"

            If SystemOptions.UserInterface = ArabicInterface Then
                    
                Me.lbl(4).Caption = " «—ÌŒ ‰Â«Ì… «·› —…:" & Format(DateAdd("d", val(Me.TxtInterval.text), Date), "yyyy/M/d")
            Else
                Me.lbl(4).Caption = " Period End At" & Format(DateAdd("d", val(Me.TxtInterval.text), Date), "yyyy/M/d")
            End If

        ElseIf Me.CboType.ListIndex = 2 Then
            StrSQL = StrSQL + " ExpireDate <=#" & SQLDate(DateAdd("d", val(Me.TxtInterval.text), Date)) & "#"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(4).Caption = " «—ÌŒ ‰Â«Ì… «·› —…:" & Format(DateAdd("d", val(Me.TxtInterval.text), Date), "yyyy/M/d")
            Else
                Me.lbl(4).Caption = "Period End At" & Format(DateAdd("d", val(Me.TxtInterval.text), Date), "yyyy/M/d")
            End If
        End If

    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "Select * From QryGuaranteeAlram Where"

        If Me.CboType.ListIndex = 0 Then
            StrSQL = StrSQL + " ExpireDate <='" & SQLDate(Date) & "'"
            Me.lbl(4).Caption = ""
        ElseIf Me.CboType.ListIndex = 1 Then
            StrSQL = StrSQL + " ExpireDate <='" & SQLDate(DateAdd("d", val(Me.TxtInterval.text), Date)) & "' And ExpireDate >='" & SQLDate(Date) & "'"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(4).Caption = " «—ÌŒ ‰Â«Ì… «·› —…:" & Format(DateAdd("d", val(Me.TxtInterval.text), Date), "yyyy/M/d")
            Else
                Me.lbl(4).Caption = "Period End At" & Format(DateAdd("d", val(Me.TxtInterval.text), Date), "yyyy/M/d")
            End If

        ElseIf Me.CboType.ListIndex = 2 Then
            StrSQL = StrSQL + " ExpireDate <='" & SQLDate(DateAdd("d", val(Me.TxtInterval.text), Date)) & "'"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(4).Caption = " «—ÌŒ ‰Â«Ì… «·› —…:" & Format(DateAdd("d", val(Me.TxtInterval.text), Date), "yyyy/M/d")
            Else
                Me.lbl(4).Caption = "Period End At" & Format(DateAdd("d", val(Me.TxtInterval.text), Date), "yyyy/M/d")
            End If
        End If
    End If

    Set rs = New ADODB.Recordset

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Fg
        Fg.Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows

        If Not (rs.BOF Or rs.EOF) Then
            Fg.Rows = .FixedRows + rs.RecordCount

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
                .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs("Qty").value), "", rs("Qty").value)
                .TextMatrix(i, .ColIndex("ItemSerial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)

                If Not IsNull(rs("Transaction_Date").value) Then
                    .TextMatrix(i, .ColIndex("TransactionDate")) = Format(rs("Transaction_Date").value, "yyyy/M/d")
                End If

                .TextMatrix(i, .ColIndex("TransactionSerial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
                .TextMatrix(i, .ColIndex("guaranteeTime")) = IIf(IsNull(rs("guaranteeTime").value), "", rs("guaranteeTime").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)

                If Not IsNull(rs("ExpireDate").value) Then
                    .TextMatrix(i, .ColIndex("ExpireDate")) = Format(rs("ExpireDate").value, "yyyy/M/d")
                End If

                .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(rs("Quantity").value), "", rs("Quantity").value)
                rs.MoveNext
            Next i

        Else
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Set rs = Nothing
End Sub
Public Sub FillItemExperDate()
Dim My_SQL As String
Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
      '  My_SQL = " SELECT     TOP 100 PERCENT dbo.Transaction_Details.ExpiryDate, SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS Qty, "
      '  My_SQL = My_SQL & "               dbo.Transaction_Details.Item_ID , dbo.TblItems.itemcode, dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode"
      '  My_SQL = My_SQL & "    FROM         dbo.Transaction_Details INNER JOIN"
      '  My_SQL = My_SQL & "              dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
      '  My_SQL = My_SQL & "              dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
      '  My_SQL = My_SQL & "              dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
      '  If CheckBox1.value = vbChecked Then
      '  My_SQL = My_SQL & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.Text) & " or dbo.Transaction_Details.ExpiryDate is null) "
      '  Else
      '  My_SQL = My_SQL & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.Text) & " ) "
      '  End If
      '  My_SQL = My_SQL & "    GROUP BY dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
      '  My_SQL = My_SQL & "               dbo.TblItems.ItemNamee , dbo.TblItems.Fullcode"
      '  My_SQL = My_SQL & "  Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) > 0)"
      '  My_SQL = My_SQL & " ORDER BY dbo.Transaction_Details.ExpiryDate"
         sql = " SELECT     TOP 100 PERCENT dbo.Transaction_Details.ExpiryDate, SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS Qty, "
        sql = sql & "               dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.Transactions.BranchId,"
        sql = sql & "              dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_namee"
        sql = sql & "  FROM         dbo.Transaction_Details INNER JOIN"
        sql = sql & "              dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
        sql = sql & "              dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
        sql = sql & "              dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
        sql = sql & "              dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
        If CheckBox1.value = vbChecked Then
        
                        If CboType2.ListIndex = 0 Or CboType2.ListIndex = -1 Then '⁄œœ «Ì«„
                            sql = sql & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.text) & " or dbo.Transaction_Details.ExpiryDate is null) "
                           
                        Else
                                        sql = sql & "       Where (dbo.TransactionTypes.StockEffect <> 0)  and "
                                         If Not IsNull(XPDtbBuyFrom.value) Then
                                              sql = sql + "  (  dbo.Transaction_Details.ExpiryDate >=" & SQLDate(XPDtbBuyFrom.value, True) & ""
                                          End If
                                           If Not IsNull(Me.XPDtpBuyTo.value) Then
                                             sql = sql + " and  dbo.Transaction_Details.ExpiryDate  <=" & SQLDate(XPDtpBuyTo.value, True) & ""
                                           End If
            
                          sql = sql & "     or dbo.Transaction_Details.ExpiryDate is null) "
                          
                        
                        End If
    
        
        
        Else
                     If CboType2.ListIndex = 0 Or CboType2.ListIndex = -1 Then '⁄œœ «Ì«„
                          sql = sql & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.text) & " ) "
                    Else
                    'sql = sql & "       Where (dbo.TransactionTypes.StockEffect <> 0) AND (DATEDIFF(d, dbo.Transaction_Details.ExpiryDate, " & SQLDate(Date, True) & ") >=" & val(Me.Text1.Text) & " ) "
                    
                                      sql = sql & "       Where (dbo.TransactionTypes.StockEffect <> 0)  and "
                                         If Not IsNull(XPDtbBuyFrom.value) Then
                                              sql = sql + "  (  dbo.Transaction_Details.ExpiryDate >=" & SQLDate(XPDtbBuyFrom.value, True) & ""
                                          End If
                                           If Not IsNull(Me.XPDtpBuyTo.value) Then
                                             sql = sql + " and  dbo.Transaction_Details.ExpiryDate  <=" & SQLDate(XPDtpBuyTo.value, True) & ")"
                                           End If
            
          End If
        End If
           Dim SotreID As String
   Dim i As Integer
    SotreID = "0"
    For i = 0 To Me.ListStoreSelect.ListCount - 1
    SotreID = SotreID & "," & Me.ListStoreSelect.ItemData(i)
    Next i
    If SotreID <> "0" Then
     sql = sql & " and   dbo.Transactions.StoreID in (" & SotreID & ")"
     End If
      '  If val(DCBranch.BoundText) <> 0 And DCBranch.Text <> "" Then
      '  sql = sql & " and dbo.Transactions.BranchId =" & val(DCBranch.BoundText) & ""
      '  End If
        sql = sql & " GROUP BY dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, "
        sql = sql & "              dbo.TblItems.ItemNamee , dbo.TblItems.Fullcode, dbo.Transactions.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
        sql = sql & "  Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) > 0)"
        sql = sql & " ORDER BY dbo.Transaction_Details.ExpiryDate"
        rs.Open sql, Cn, adOpenKeyset, adLockReadOnly, adCmdText
       PrinterSql = sql
        
        CurrentMessage = "From : " & XPDtbBuyFrom.value & CHR(13)
         CurrentMessage = CurrentMessage & "TO : " & XPDtpBuyTo.value
        
    If rs.RecordCount > 0 Then
   With VSFlexGrid1
   .Rows = .Rows + rs.RecordCount - 1
                    For i = 1 To .Rows - 1
                        .TextMatrix(i, .ColIndex("Serial")) = i
                        .TextMatrix(i, .ColIndex("ExpiryDate")) = IIf(IsNull(rs.Fields("ExpiryDate").value), "", rs.Fields("ExpiryDate").value)
                        .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs.Fields("Qty").value), 0, rs.Fields("Qty").value)
                        .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs.Fields("Fullcode").value), "", rs.Fields("Fullcode").value)
                        If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs.Fields("ItemName").value), "", rs.Fields("ItemName").value)
                        Else
                        .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs.Fields("ItemNamee").value), "", rs.Fields("ItemNamee").value)
                        End If
                        
                        CurrentMessage = CurrentMessage & " Item# :  " & .TextMatrix(i, .ColIndex("Fullcode")) & vbTab
                        CurrentMessage = CurrentMessage & " Item :  " & .TextMatrix(i, .ColIndex("ItemName")) & vbTab
                        CurrentMessage = CurrentMessage & " ExpiryDate  " & .TextMatrix(i, .ColIndex("ExpiryDate")) & vbTab
                        CurrentMessage = CurrentMessage & " Qty " & .TextMatrix(i, .ColIndex("Qty")) & CHR(13)
                        CurrentMessage = "---------------------------------------------------------------------------------------------------"
                         rs.MoveNext
                    Next i
   End With
                    rs.Close
   End If
     
End Sub

Private Sub ISButton1_Click()
print_report
End Sub

Private Sub ISButton2_Click()
Unload Me
End Sub

Private Sub ISButton3_Click()
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
FillItemExperDate
End Sub

Private Sub ISButton4_Click()
ISButton3_Click

        If Email = "salimman2003@gmail.com" Then Exit Sub
        Dim RetVal As String
        
           RetVal = SendMail(Trim$(Email), _
        "", _
        "", _
        Trim$(CurrentMessage), _
        "", _
        0, _
        "", _
        "", _
        Trim$(TxtAttach.text), _
       False, True)
           MsgBox IIf(RetVal = "ok", "Message sent!", RetVal)
           
      
End Sub

Private Sub ISButton5_Click()
Unload Me
End Sub

Private Sub ISButton6_Click()
FillGridBooking
End Sub

Private Sub ISButton7_Click()
print_reportBooking
End Sub

Private Sub Label14_Click()
    Dim i As Integer
    Me.ListStoreSelect2.Clear
    For i = 0 To Me.ListAllStore2.ListCount - 1
        Me.ListStoreSelect2.AddItem ListAllStore2.List(i)
        ListStoreSelect2.ItemData(i) = ListAllStore2.ItemData(i)
    Next i
End Sub

Private Sub Label15_Click()
 If Me.ListAllStore2.ListIndex > -1 Then
    Me.ListStoreSelect2.AddItem ListAllStore2.List(ListAllStore2.ListIndex)
    ListStoreSelect2.ItemData(ListStoreSelect2.NewIndex) = ListAllStore2.ItemData(ListAllStore2.ListIndex)
End If
End Sub

Private Sub Label17_Click()
 If Me.ListBranchAll2.ListIndex > -1 Then
    Me.ListBranchSelected2.AddItem ListBranchAll2.List(ListBranchAll2.ListIndex)
    ListBranchSelected2.ItemData(ListBranchSelected2.NewIndex) = ListBranchAll2.ItemData(ListBranchAll2.ListIndex)
End If
FillMylist3
End Sub

Private Sub Label18_Click()
    Dim i As Integer
    Me.ListBranchSelected2.Clear
    For i = 0 To Me.ListBranchAll2.ListCount - 1
        Me.ListBranchSelected2.AddItem ListBranchAll2.List(i)
        ListBranchSelected2.ItemData(i) = ListBranchAll2.ItemData(i)
    Next i
   FillMylist3
End Sub

Private Sub Label19_Click()
ListBranchSelected2.Clear
Me.ListAllStore2.Clear
Me.ListStoreSelect2.Clear
End Sub

Private Sub Label2_Click()
If ListStoreSelect2.ListIndex > -1 Then
ListStoreSelect2.RemoveItem (ListStoreSelect2.ListIndex)
End If
End Sub

Private Sub Label20_Click()
If ListBranchSelected2.ListIndex > -1 Then
ListBranchSelected2.RemoveItem (ListBranchSelected2.ListIndex)
End If
FillMylist3
End Sub

Private Sub Label3_Click()
Me.ListStoreSelect2.Clear
End Sub

Private Sub Label8_Click()

 If Me.ListAllStore.ListIndex > -1 Then
    Me.ListStoreSelect.AddItem ListAllStore.List(ListAllStore.ListIndex)
    ListStoreSelect.ItemData(ListStoreSelect.NewIndex) = ListAllStore.ItemData(ListAllStore.ListIndex)
End If
End Sub

Function FillMylist()
    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
    Set rs2 = New ADODB.Recordset
    sql = " SELECT * from  TblBranchesData "
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Me.ListBranchAll.Clear
    Me.ListBranchSelected.Clear
    Me.ListBranchAll2.Clear
    Me.ListBranchSelected2.Clear
    ListAllStore.Clear
    ListStoreSelect.Clear
    ListAllStore2.Clear
    ListStoreSelect2.Clear
    If rs2.RecordCount > 0 Then
        For i = 1 To rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListBranchAll.AddItem IIf(IsNull(rs2("branch_name").value), "", rs2("branch_name").value)
                ListBranchAll2.AddItem IIf(IsNull(rs2("branch_name").value), "", rs2("branch_name").value)
            Else
                ListBranchAll.AddItem IIf(IsNull(rs2("branch_namee").value), "", rs2("branch_namee").value)
                ListBranchAll2.AddItem IIf(IsNull(rs2("branch_namee").value), "", rs2("branch_namee").value)
            End If
            ListBranchAll.ItemData(ListBranchAll.NewIndex) = IIf(IsNull(rs2("branch_id").value), 0, rs2("branch_id").value)
            ListBranchAll2.ItemData(ListBranchAll2.NewIndex) = IIf(IsNull(rs2("branch_id").value), 0, rs2("branch_id").value)
            rs2.MoveNext
        Next i
    End If
    rs2.Close
End Function
Function FillMylist2()
    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
    Dim ActivID As String
    ActivID = "0"
    For i = 0 To Me.ListBranchSelected.ListCount - 1
    ActivID = ActivID & "," & Me.ListBranchSelected.ItemData(i)
    Next i
    Me.ListAllStore.Clear
    Me.ListStoreSelect.Clear
    If ActivID = "0" Then Exit Function
    Set rs2 = New ADODB.Recordset
    sql = " SELECT * from  TblStore where BranchId in(" & ActivID & ") "
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs2.RecordCount > 0 Then
        For i = 1 To rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListAllStore.AddItem IIf(IsNull(rs2("StoreName").value), "", rs2("StoreName").value)
            Else
                ListAllStore.AddItem IIf(IsNull(rs2("StoreNamee").value), "", rs2("StoreNamee").value)
            End If
            ListAllStore.ItemData(ListAllStore.NewIndex) = IIf(IsNull(rs2("StoreID").value), 0, rs2("StoreID").value)
            rs2.MoveNext
        Next i

    End If
    rs2.Close
End Function
Function FillMylist3()
    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
    Dim ActivID As String
    ActivID = "0"
    For i = 0 To Me.ListBranchSelected2.ListCount - 1
    ActivID = ActivID & "," & Me.ListBranchSelected2.ItemData(i)
    Next i
    Me.ListAllStore2.Clear
    Me.ListStoreSelect2.Clear
    If ActivID = "0" Then Exit Function
    Set rs2 = New ADODB.Recordset
    sql = " SELECT * from  TblStore where BranchId in(" & ActivID & ") "
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs2.RecordCount > 0 Then
        For i = 1 To rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListAllStore2.AddItem IIf(IsNull(rs2("StoreName").value), "", rs2("StoreName").value)
            Else
                ListAllStore2.AddItem IIf(IsNull(rs2("StoreNamee").value), "", rs2("StoreNamee").value)
            End If
            ListAllStore2.ItemData(ListAllStore2.NewIndex) = IIf(IsNull(rs2("StoreID").value), 0, rs2("StoreID").value)
            rs2.MoveNext
        Next i

    End If
    rs2.Close
End Function
Private Sub Text1_Change()
ISButton3_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'KeyAscii = KeyAscii_Num(KeyAscii, Me.Text1.Text, 1)
End Sub
Private Sub Label10_Click()
ListBranchSelected.Clear
Me.ListAllStore.Clear
Me.ListStoreSelect.Clear
End Sub
Private Sub Label4_Click()
 If Me.ListBranchAll.ListIndex > -1 Then
    Me.ListBranchSelected.AddItem ListBranchAll.List(ListBranchAll.ListIndex)
    ListBranchSelected.ItemData(ListBranchSelected.NewIndex) = ListBranchAll.ItemData(ListBranchAll.ListIndex)
End If
FillMylist2
End Sub
Private Sub Label5_Click()
If ListStoreSelect.ListIndex > -1 Then
ListStoreSelect.RemoveItem (ListStoreSelect.ListIndex)
End If
End Sub
Private Sub Label6_Click()
Me.ListStoreSelect.Clear
End Sub

Private Sub Label7_Click()
    Dim i As Integer
    Me.ListStoreSelect.Clear
    For i = 0 To Me.ListAllStore.ListCount - 1
        Me.ListStoreSelect.AddItem ListAllStore.List(i)
        ListStoreSelect.ItemData(i) = ListAllStore.ItemData(i)
    Next i
End Sub
Private Sub Label11_Click()
If ListBranchSelected.ListIndex > -1 Then
ListBranchSelected.RemoveItem (ListBranchSelected.ListIndex)
End If
FillMylist2
End Sub
Private Sub Label9_Click()

    Dim i As Integer
    Me.ListBranchSelected.Clear
    For i = 0 To Me.ListBranchAll.ListCount - 1
        Me.ListBranchSelected.AddItem ListBranchAll.List(i)
        ListBranchSelected.ItemData(i) = ListBranchAll.ItemData(i)
    Next i
  
   FillMylist2
End Sub
Sub FillGridBooking()
Dim i As Integer
Dim sql As String
Dim SotreID As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.BookingDate, dbo.Transactions.NoteSerial1, "
sql = sql & "                       dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.StoreID, dbo.TblStore.StoreName,"
sql = sql & "                      dbo.TblStore.StoreAdress, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName,"
sql = sql & "                      dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode AS ItemFullcode, dbo.TblItems.barCodeNO"
sql = sql & " FROM         dbo.TblItems RIGHT OUTER JOIN"
sql = sql & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID RIGHT OUTER JOIN"
sql = sql & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
sql = sql & " WHERE     (dbo.Transactions.Transaction_Type = 39) and dbo.Transaction_Details.ShowQty>0 and dbo.Transactions.BookingDate<" & SQLDate(Date, True) & " "
    SotreID = "0"
    For i = 0 To Me.ListStoreSelect2.ListCount - 1
    SotreID = SotreID & "," & Me.ListStoreSelect2.ItemData(i)
    Next i
     If SotreID <> "0" Then
     sql = sql & " and   dbo.Transactions.StoreID in (" & SotreID & ")"
     End If
  If Not IsNull(FrmDate.value) Then
  sql = sql & " and  dbo.Transactions.Transaction_Date >=" & SQLDate(FrmDate.value, True) & ""
  End If
    If Not IsNull(TODate.value) Then
  sql = sql & " and  dbo.Transactions.Transaction_Date <=" & SQLDate(TODate.value, True) & ""
  End If
      With Me.VSFlexGrid2
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
      End With
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
    With Me.VSFlexGrid2
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
        .Rows = .Rows + Rs3.RecordCount
    For i = 1 To .Rows - 1
    .TextMatrix(i, .ColIndex("Serial")) = i
    .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(Rs3("Transaction_ID").value), 0, Rs3("Transaction_ID").value)
    .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(Rs3("Item_ID").value), 0, Rs3("Item_ID").value)
    .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs3("NoteSerial1").value), "", Rs3("NoteSerial1").value)
     .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(Rs3("Transaction_Date").value), "", Rs3("Transaction_Date").value)
    .TextMatrix(i, .ColIndex("BookingDate")) = IIf(IsNull(Rs3("BookingDate").value), "", Rs3("BookingDate").value)
    .TextMatrix(i, .ColIndex("ShowQty")) = IIf(IsNull(Rs3("ShowQty").value), 0, Rs3("ShowQty").value)
    .TextMatrix(i, .ColIndex("ItemFullcode")) = IIf(IsNull(Rs3("ItemFullcode").value), "", Rs3("ItemFullcode").value)
    If SystemOptions.UserInterface = ArabicInterface Then
     .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs3("CusName").value), "", Rs3("CusName").value)
     .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs3("StoreName").value), "", Rs3("StoreName").value)
     .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs3("ItemName").value), "", Rs3("ItemName").value)
    Else
     .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs3("CusNamee").value), "", Rs3("CusNamee").value)
     .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(Rs3("StoreNamee").value), "", Rs3("StoreNamee").value)
      .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs3("ItemNamee").value), "", Rs3("ItemNamee").value)
    End If
    Rs3.MoveNext
    Next i
    End With
End If
End Sub
Private Sub TxtInterval_KeyPress(KeyAscii As Integer)

    If val(TxtInterval) > (366 * 3) Then
        KeyAscii = 0
    Else
        KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtInterval.text, 1)
    End If

End Sub

Private Sub SetControls()

    If Me.CboType.ListIndex = -1 Or Me.CboType.ListIndex = 0 Then
        Me.lbl(2).Enabled = False
        Me.TxtInterval.Enabled = False
        Me.lbl(4).Enabled = False
    Else
        Me.lbl(2).Enabled = True
        Me.TxtInterval.Enabled = True
        Me.lbl(4).Enabled = True
    End If




    If Me.CboType2.ListIndex = -1 Or Me.CboType2.ListIndex = 0 Then
        Me.lbl(9).Enabled = True
        Me.Text1.Enabled = True
        Me.lbl(8).Enabled = True
        fra(5).Enabled = False
    Else
        Me.lbl(9).Enabled = False
        Me.Text1.Enabled = False
        Me.lbl(8).Enabled = False
        fra(5).Enabled = True
    End If
    
    
    
End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.VSFlexGrid2
If .ColKey(Col) <> "show" And .ColKey(Col) <> "Cancell" Then
Cancel = True
End If
End With
End Sub

Private Sub VSFlexGrid2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With VSFlexGrid2
Select Case .ColKey(Col)
Case "show"
If val(.TextMatrix(Row, .ColIndex("Transaction_ID"))) <> 0 Then
FrmPO7.show
FrmPO7.Retrive val(.TextMatrix(Row, .ColIndex("Transaction_ID")))
End If
Case "Cancell"
If val(.TextMatrix(Row, .ColIndex("Transaction_ID"))) <> 0 And val(.TextMatrix(Row, .ColIndex("Item_ID"))) <> 0 Then
Cn.Execute "Update Transaction_Details set ShowQty=0 where Item_ID =" & val(.TextMatrix(Row, .ColIndex("Item_ID"))) & " and Transaction_ID=" & val(.TextMatrix(Row, .ColIndex("Transaction_ID"))) & ""
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·€«¡ ”‰œ «·ÕÃ“"
Else
MsgBox "Booking Canceled"
End If
FillGridBooking
End If
End Select
End With
End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.VSFlexGrid2
Select Case .ColKey(Col)
 Case "show"
            .ColComboList(.ColIndex("show")) = "..."
  Case "Cancell"
            .ColComboList(.ColIndex("Cancell")) = "..."
     End Select
    End With
End Sub
