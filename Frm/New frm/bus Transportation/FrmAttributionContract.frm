VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAttributionContract 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄ﬁœ «·«”‰«œ"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13605
   Icon            =   "FrmAttributionContract.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   13605
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   10185
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13605
      _cx             =   23998
      _cy             =   17965
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
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   630
         Left            =   0
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   0
         Width           =   13545
         _cx             =   23892
         _cy             =   1111
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
         Caption         =   "     ⁄ﬁœ «·«”‰«œ    "
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
            TabIndex        =   39
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   40
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAttributionContract.frx":038A
            ColorButton     =   16777215
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
            TabIndex        =   41
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAttributionContract.frx":0724
            ColorButton     =   16777215
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
            TabIndex        =   42
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAttributionContract.frx":0ABE
            ColorButton     =   16777215
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
            TabIndex        =   43
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAttributionContract.frx":0E58
            ColorButton     =   16777215
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   4605
         Left            =   120
         TabIndex        =   44
         Top             =   3705
         Width           =   13440
         _cx             =   23707
         _cy             =   8123
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   12648447
         ForeColor       =   128
         FrontTabColor   =   14871017
         BackTabColor    =   8454143
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "»Ì«‰«  «·œ›⁄« |»Ì«‰«  «·„—ﬂ»« "
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   1
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   0   'False
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   4185
            Left            =   14085
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   45
            Width           =   13350
            _cx             =   23548
            _cy             =   7382
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
            Begin VB.TextBox lblCalStudent 
               Alignment       =   1  'Right Justify
               Height          =   420
               Left            =   270
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   4140
               Width           =   1860
            End
            Begin VB.CommandButton Command3 
               Caption         =   "«÷«›… ”ÿ—"
               Height          =   510
               Left            =   18150
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   4080
               Width           =   1650
            End
            Begin VB.CommandButton Command2 
               Caption         =   " Õ–› ”ÿ—"
               Height          =   510
               Left            =   16200
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   4080
               Width           =   1770
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Õ–› «·ﬂ·"
               Height          =   510
               Left            =   14310
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   4080
               Width           =   1710
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   3690
               Left            =   120
               TabIndex        =   92
               Top             =   390
               Width           =   13110
               _cx             =   23125
               _cy             =   6509
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial (Arabic)"
                  Size            =   10.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   16776960
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
               Rows            =   1
               Cols            =   26
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmAttributionContract.frx":11F2
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   195
               Index           =   11
               Left            =   10320
               TabIndex        =   115
               Top             =   0
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   344
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ”ÿ—"
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
               ButtonImage     =   "FrmAttributionContract.frx":1611
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   195
               Index           =   12
               Left            =   8160
               TabIndex        =   116
               Top             =   0
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   344
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› «·ﬂ·"
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
               ButtonImage     =   "FrmAttributionContract.frx":1BAB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   315
               Index           =   13
               Left            =   6000
               TabIndex        =   117
               Top             =   0
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               ButtonPositionImage=   1
               Caption         =   "«÷«›… ”ÿ—"
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
               ButtonImage     =   "FrmAttributionContract.frx":2145
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
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ï «·ÿ·«» «·–Ì‰  „  ”ﬂÌ‰Â„"
               ForeColor       =   &H000000C0&
               Height          =   375
               Index           =   23
               Left            =   2310
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   4215
               Width           =   2880
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   4185
            Left            =   45
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   45
            Width           =   13350
            _cx             =   23548
            _cy             =   7382
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
            Begin VB.TextBox TxtPaymentCount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   10530
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   105
               Width           =   1620
            End
            Begin VB.TextBox TxtPeriods 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3090
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   135
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.ComboBox DcbPeriodsID 
               Height          =   315
               ItemData        =   "FrmAttributionContract.frx":89A7
               Left            =   1560
               List            =   "FrmAttributionContract.frx":89B4
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   135
               Visible         =   0   'False
               Width           =   1530
            End
            Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
               Height          =   3405
               Left            =   120
               TabIndex        =   28
               Top             =   600
               Width           =   13170
               _cx             =   23230
               _cy             =   6006
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial (Arabic)"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
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
               BackColorAlternate=   16776960
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
               Rows            =   50
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmAttributionContract.frx":89C7
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   315
               Index           =   14
               Left            =   120
               TabIndex        =   27
               Top             =   135
               Visible         =   0   'False
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
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
               ButtonImage     =   "FrmAttributionContract.frx":8B83
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
            Begin Dynamic_Byte.NourHijriCal FirstPaymentDateH 
               Height          =   285
               Left            =   6000
               TabIndex        =   24
               Top             =   135
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   503
            End
            Begin MSComCtl2.DTPicker FirstPaymentDate 
               Height          =   285
               Left            =   7440
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   135
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   100073475
               CurrentDate     =   41640
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ «·œ›⁄« "
               Height          =   270
               Index           =   8
               Left            =   12120
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   135
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «Ê· œ›⁄Â"
               Height          =   270
               Index           =   9
               Left            =   8940
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   135
               Width           =   1350
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·› —Â »Ì‰ «·œ›⁄« "
               Height          =   270
               Index           =   11
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   135
               Visible         =   0   'False
               Width           =   1350
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   630
         Left            =   120
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   8385
         Width           =   11430
         _cx             =   20161
         _cy             =   1111
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
            Caption         =   "     „‰"
            Height          =   240
            Index           =   26
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Height          =   375
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   135
            Width           =   3375
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Height          =   375
            Left            =   165
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   135
            Width           =   2760
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   495
            Index           =   2
            Left            =   12705
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   180
            Width           =   3510
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   255
            Index           =   4
            Left            =   7995
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   180
            Width           =   1875
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   735
         Left            =   0
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   18960
         Width           =   12435
         _cx             =   21934
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   2295
         Left            =   120
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   480
         Width           =   13365
         _cx             =   23574
         _cy             =   4048
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
         Begin VB.TextBox txtfullcode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10530
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   1200
            Width           =   1620
         End
         Begin C1SizerLibCtl.C1Elastic Frame1 
            Height          =   615
            Left            =   2040
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   1680
            Visible         =   0   'False
            Width           =   1770
            _cx             =   3122
            _cy             =   1085
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
            Begin VB.OptionButton opt_all 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·ÌÊ„ ··⁄ﬁœ ﬂ«„· "
               Height          =   225
               Left            =   105
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   315
               Width           =   1590
            End
            Begin VB.OptionButton opt_one 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·ÌÊ„ ·ﬂ· ”Ì«—…"
               Height          =   225
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   24
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.TextBox txtRecordno 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7485
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   1200
            Width           =   1755
         End
         Begin VB.ComboBox cbContractType 
            Height          =   315
            ItemData        =   "FrmAttributionContract.frx":F3E5
            Left            =   3810
            List            =   "FrmAttributionContract.frx":F3E7
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   1680
            Width           =   2550
         End
         Begin VB.TextBox txtRes 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3810
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   2490
         End
         Begin VB.TextBox txtIDAC 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   10530
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   240
            Width           =   1620
         End
         Begin VB.TextBox txtIDMC 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   -240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtMinistryContractNo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7845
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   -120
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.TextBox XPTxtBoxName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10530
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   720
            Width           =   1620
         End
         Begin VB.TextBox txtProcessNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10575
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   -240
            Visible         =   0   'False
            Width           =   1620
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   315
            Left            =   1230
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   2520
            Visible         =   0   'False
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   100073475
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   315
            Left            =   1260
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   100073475
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH 
            Height          =   315
            Left            =   150
            TabIndex        =   53
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal dtpToDateH 
            Height          =   315
            Left            =   1245
            TabIndex        =   14
            Top             =   2520
            Visible         =   0   'False
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal dtpSContractDateH 
            Height          =   315
            Left            =   3825
            TabIndex        =   8
            Top             =   720
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker dtpSContractDate 
            Height          =   315
            Left            =   5100
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   720
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   100073475
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpEContractDate 
            Height          =   315
            Left            =   1260
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   720
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   100073475
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpEContractDateH 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
         End
         Begin MSDataListLib.DataCombo dcMangerialAreaID 
            Height          =   315
            Left            =   7485
            TabIndex        =   3
            Top             =   1680
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCity 
            Height          =   315
            Left            =   10530
            TabIndex        =   13
            Top             =   1680
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcDuration 
            Height          =   315
            Left            =   7485
            TabIndex        =   6
            Top             =   720
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCustomer 
            Height          =   315
            Left            =   3810
            TabIndex        =   11
            Top             =   1200
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMinistry 
            Height          =   315
            Left            =   7485
            TabIndex        =   88
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   120
            TabIndex        =   106
            Top             =   1800
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpEmbark 
            Height          =   315
            Left            =   1260
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   16776960
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   100073475
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpEmbarkH 
            Height          =   315
            Left            =   120
            TabIndex        =   111
            Top             =   1200
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·„»«‘—…"
            ForeColor       =   &H000000FF&
            Height          =   555
            Index           =   25
            Left            =   2475
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   1200
            Width           =   1125
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   315
            Index           =   24
            Left            =   1350
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   1800
            Width           =   525
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂÊœ"
            Height          =   315
            Index           =   22
            Left            =   12570
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã· "
            Height          =   315
            Index           =   21
            Left            =   9390
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1200
            Width           =   885
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· ⁄«ﬁœ"
            Height          =   315
            Index           =   16
            Left            =   6210
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»«ﬁÏ «· Œ’Ì’"
            Height          =   315
            Index           =   7
            Left            =   6165
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ ⁄Âœ"
            Height          =   315
            Index           =   6
            Left            =   6870
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   1200
            Width           =   465
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄ﬁœ"
            Height          =   285
            Index           =   0
            Left            =   12285
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄ﬁœ «·Ê“«—…"
            Height          =   315
            Index           =   15
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”„Ï «· ⁄«ﬁœ"
            Height          =   315
            Index           =   3
            Left            =   11940
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   2535
            TabIndex        =   60
            Top             =   2640
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·ÌÊ„"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2805
            TabIndex        =   59
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  »œ«Ì… «· ⁄«ﬁœ "
            Height          =   435
            Index           =   5
            Left            =   6450
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «‰ Â«¡ «· ⁄«ﬁœ "
            Height          =   555
            Index           =   8
            Left            =   2475
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Õ«›Ÿ… "
            Height          =   315
            Index           =   10
            Left            =   12600
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   1680
            Width           =   600
            WordWrap        =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «· ⁄·Ì„Ì…"
            Height          =   315
            Index           =   9
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   1680
            Width           =   1110
            WordWrap        =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄«„ «·œ—«”Ï"
            Height          =   315
            Index           =   1
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   720
            Width           =   1110
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   720
         Left            =   120
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2880
         Width           =   13365
         _cx             =   23574
         _cy             =   1270
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
         Begin VB.TextBox txtActualDayValue 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   240
            Width           =   810
         End
         Begin VB.TextBox txtDaysCount 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   9480
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   240
            Width           =   870
         End
         Begin VB.TextBox txtDayValue 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7485
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   240
            Width           =   780
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   5370
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   870
         End
         Begin VB.TextBox txtStudentCustom 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7455
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   780
         End
         Begin VB.TextBox txtStudentCount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11505
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   960
         End
         Begin VB.TextBox txtNet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2280
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtDiscount 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3780
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   630
         End
         Begin VB.CommandButton cmdSub 
            Caption         =   "-"
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
            Left            =   4410
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   240
            Width           =   300
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "+"
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
            Left            =   4950
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   405
         End
         Begin MSComCtl2.DTPicker DtatAdd 
            Height          =   330
            Left            =   0
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   -135
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   100073475
            CurrentDate     =   37140
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·›⁄·Ì… ··ÌÊ„"
            Height          =   270
            Index           =   20
            Left            =   885
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·ÌÊ„"
            Height          =   390
            Index           =   19
            Left            =   8265
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «Ì«„ «·⁄„· «·›⁄·Ì…"
            Height          =   390
            Index           =   17
            Left            =   10365
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·«Ã„«·Ì…"
            Height          =   270
            Index           =   13
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·ÿ·«»"
            Height          =   270
            Index           =   12
            Left            =   12465
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Œ’’ «·ÿ«·» ··ÌÊ„"
            Height          =   390
            Index           =   11
            Left            =   8250
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ï"
            Height          =   270
            Index           =   14
            Left            =   3030
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   18
            Left            =   4710
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   180
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   810
         Left            =   0
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   9375
         Width           =   13605
         _cx             =   23998
         _cy             =   1429
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
         Align           =   2
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
            Height          =   705
            Index           =   0
            Left            =   12420
            TabIndex        =   29
            Top             =   75
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmAttributionContract.frx":F3E9
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
            Height          =   705
            Index           =   1
            Left            =   11085
            TabIndex        =   30
            Top             =   75
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmAttributionContract.frx":15C4B
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
            Height          =   705
            Index           =   2
            Left            =   9390
            TabIndex        =   31
            Top             =   75
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmAttributionContract.frx":1C4AD
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
            Height          =   705
            Index           =   3
            Left            =   7860
            TabIndex        =   32
            Top             =   75
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmAttributionContract.frx":22D0F
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
            Height          =   705
            Index           =   4
            Left            =   6990
            TabIndex        =   33
            Top             =   75
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmAttributionContract.frx":29571
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
            Height          =   705
            Index           =   6
            Left            =   885
            TabIndex        =   36
            Top             =   75
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmAttributionContract.frx":2FDD3
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
            Height          =   705
            Left            =   150
            TabIndex        =   37
            Top             =   75
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmAttributionContract.frx":599F5
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
            Height          =   705
            Index           =   7
            Left            =   5985
            TabIndex        =   34
            Top             =   75
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmAttributionContract.frx":60257
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
            Height          =   705
            Index           =   9
            Left            =   2580
            TabIndex        =   35
            Top             =   75
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   1244
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
            ButtonImage     =   "FrmAttributionContract.frx":66AB9
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
            Height          =   705
            Index           =   5
            Left            =   3735
            TabIndex        =   108
            Top             =   75
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1244
            ButtonPositionImage=   1
            Caption         =   "„·Õﬁ 2"
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
            ButtonImage     =   "FrmAttributionContract.frx":6D31B
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
            Height          =   705
            Index           =   8
            Left            =   4455
            TabIndex        =   109
            Top             =   75
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   1244
            ButtonPositionImage=   1
            Caption         =   "„·Õﬁ 1"
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
            ButtonImage     =   "FrmAttributionContract.frx":73B7D
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
            Height          =   705
            Index           =   10
            Left            =   2145
            TabIndex        =   113
            Top             =   75
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   1244
            ButtonPositionImage=   1
            Caption         =   "‰”Œ… „„«À·Â"
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
            ButtonImage     =   "FrmAttributionContract.frx":7A3DF
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
      End
      Begin VB.Label lblResC 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9510
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   8535
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblDifC 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   10530
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   8535
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblTotalC 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   11400
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   8520
         Visible         =   0   'False
         Width           =   915
      End
   End
End
Attribute VB_Name = "FrmAttributionContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim Rs_Temp5 As ADODB.Recordset
Dim RsTemp2 As ADODB.Recordset
Dim RsTemp5 As ADODB.Recordset
Dim rsVendor As ADODB.Recordset

Dim rsDivid As ADODB.Recordset

Dim rs_det As ADODB.Recordset
Dim TTP As clstooltip
Dim rsInst As ADODB.Recordset
Dim Operation As String
 Dim rsYrs As ADODB.Recordset

Private Sub Calc_Installments()
On Error Resume Next
Dim str As String, i As Integer
str = " select TblDurations.type ,TblDurations.fromdate HFromDate , TblDurations.todate HToDate , TblDurations.FromDateH  HFromDateH ,TblDurations.TODateH  HToDateH , TblDurations.Name HName , TblDurations_Details.* from  TblDurations , TblDurations_Details   where  TblDurations_Details.DID = TblDurations.id   and DID = " & val(dcDuration.BoundText)

'Str = Str & "  and (   DatePart(yyyy , TblDurations_Details.FromDate ) >=   DatePart ( yyyy ,  cast ( '" & Format$(dtpEmbark.value, "yyyy-MM-dd") & "' as date ) )  and "
'Str = Str & " DatePart(m , TblDurations_Details.FromDate ) >=   DatePart ( m ,  cast ( '" & Format$(dtpEmbark.value, "yyyy-MM-dd") & "' as date ) )  ) "


Set Rs_Temp = New ADODB.Recordset
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

Dim year   As String, RYear As String
Dim Month As Integer, Typ As Integer, RMonth As Integer, j As Integer, cnt As Integer

   If Rs_Temp.RecordCount > 0 Then
            Rs_Temp.MoveFirst
            TxtPaymentCount.Text = Rs_Temp.RecordCount
            With FgInstallments
            cnt = .FixedRows + Rs_Temp.RecordCount
             .Rows = cnt
            For i = 1 To Rs_Temp.RecordCount
                        
                        Typ = IIf(IsNull(Rs_Temp("type").value), 0, Rs_Temp("type").value)
                        
                        If Typ = 0 Then
                                   year = Format(dtpEmbark.value, "yyyy")
                                Month = Format(dtpEmbark.value, "MM")
                                RYear = Format(Rs_Temp("FromDate").value, "yyyy")
                                RMonth = Format(Rs_Temp("FromDate").value, "MM")
                        Else
              'VBA.Calendar = vbCalHijri
              year = Format(dtpEmbarkH.value, "yyyy")
'year = Mid(dtpEmbarkH.value, 1, 4)

                                Month = Format(dtpEmbarkH.value, "MM")
                           '     Month = Mid(dtpEmbarkH.value, 9, 2)
                       '     If Month = 0 Then Month = 2
                                RYear = Format(Rs_Temp("FromDateH").value, "yyyy")
                                RMonth = Format(Rs_Temp("FromDateH").value, "MM")
                        End If
                        
                        If (year = RYear And RMonth >= Month) Or (Int(year) < Int(RYear) And RMonth <= Month) Then
                                j = j + 1
                                .TextMatrix(j, .ColIndex("QestID")) = j
                                If j = 1 Then
                                .TextMatrix(j, .ColIndex("Due_DateH")) = Format(dtpEmbarkH.value, "yyyy/MM/dd")
                               '    VBA.Calendar = vbCalGreg
                                   
                                .TextMatrix(j, .ColIndex("Due_Date")) = Format(dtpEmbark.value, "yyyy/MM/dd")
                                Else
                                
                                .TextMatrix(j, .ColIndex("Due_DateH")) = IIf(IsNull(Rs_Temp("FromDateH").value), "", Format(Rs_Temp("FromDateH").value, "yyyy/MM/dd"))
                                .TextMatrix(j, .ColIndex("Due_Date")) = IIf(IsNull(Rs_Temp("FromDate").value), "", Format(Rs_Temp("FromDate").value, "yyyy/MM/dd"))
                                End If
                                .TextMatrix(j, .ColIndex("ToDateH")) = IIf(IsNull(Rs_Temp("ToDateH").value), "", Format(Rs_Temp("ToDateH").value, "yyyy/MM/dd"))
                                .TextMatrix(j, .ColIndex("ToDate")) = IIf(IsNull(Rs_Temp("FromDate").value), "", Format(Rs_Temp("FromDate").value, "yyyy/MM/dd"))
                                .TextMatrix(j, .ColIndex("MonthID")) = IIf(IsNull(Rs_Temp("ID").value), "", Rs_Temp("ID").value)
                                Dim ID As Integer
                                ID = IIf(IsNull(Rs_Temp("ID").value), 0, Rs_Temp("ID").value)
                                If ID <> 0 Then
                                .TextMatrix(j, .ColIndex("ADays")) = GetMonthDays(ID, Typ, dtpEmbarkH.value, dtpEmbark.value) - GetHold(ID, Typ, dtpEmbarkH.value, dtpEmbark.value)
                                End If
                                .TextMatrix(j, .ColIndex("Value")) = Round(val(txtdayvalue.Text) * val(.TextMatrix(j, .ColIndex("ADays"))), 2)
                        Else
                                'cnt = cnt - 1
                                .RemoveItem (.Rows - 1)
                              '  .Rows = cnt
                        End If
                        
                        Rs_Temp.MoveNext
            Next
           End With
    End If
End Sub

Private Sub cbContractType_Click()
If cbContractType.ListIndex = 0 Then
        lbl(11).Visible = True
        txtStudentCustom.Visible = True
        txtdayvalue.Visible = False
        lbl(19).Visible = False
        Frame1.Visible = False
Else: cbContractType.ListIndex = 1
         lbl(11).Visible = False
        txtStudentCustom.Visible = False
        txtdayvalue.Visible = True
        lbl(19).Visible = True
        Frame1.Visible = True
 
End If
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
          
            XPTxtBoxName.SetFocus
            'VSFlexGrid1.Rows = VSFlexGrid1.FixedRows + 1
      '       txtIDAC.Text = CStr(new_id("TblAttributionContract", "IDAC", "", True))
             dtpEmbark_Change
             dtpFromDate_Change
dtpSContractDate_Change

dtpEContractDate_Change


Case 10
 
            TxtModFlg.Text = "N"
            Me.txtIDAC.Text = ""
 
 dcDuration.Text = ""
 dtpFromDate.value = Date
 
    dtpFromDate_Change
    
        Case 1
                             If ChekClodePeriod(dtpSContractDate.value) = True Then
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
            
            If ISAllowDeleteUpdateContract(val(txtIDAC.Text)) = False Then
                 MsgBox ("·« Ì„ﬂ‰  ⁄œÌ· «·⁄ﬁœ »”»»  ÊÃÊœ ⁄„·Ì«  „— »ÿ…  ")
           Else
                    TxtModFlg.Text = "E"
                    VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
                    dcDuration_Click 0
            End If
            
            
        Case 2
                                     If ChekClodePeriod(dtpSContractDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              If val(Me.txtIDAC) = 0 Then
              MsgBox "  —ﬁ„ «·⁄ﬁœ  ·«»œ „‰ «œŒ«·Â", vbCritical
Exit Sub
              End If
              
              
If Me.TxtModFlg = "N" Then
If CheckREpettedAttributionContract(val(Me.txtIDAC)) = True Then
MsgBox "  —ﬁ„ «·⁄ﬁœ „ÊÃÊœ „”»ﬁ« Ê·« Ì„ﬂ‰ «⁄«œÂ «œŒ«·Â", vbCritical
Exit Sub
End If
End If
            SaveData
        
        Case 3
            Undo
        Case 4

                             If ChekClodePeriod(dtpSContractDate.value) = True Then
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
            If ISAllowDeleteUpdateContract(val(txtIDAC.Text)) = False Then
                        MsgBox ("·« Ì„ﬂ‰ Õ–› «·⁄ﬁœ »”»» «Ã—«¡ ⁄„·Ì«  ⁄·Ï «·⁄ﬁœ ")
            Else
                        Del_Company
            End If
        Case 5
         '   Unload FrmSearch_MinistryContract
         '    FrmSearch_MinistryContract.SendForm = "AC"
         '    FrmSearch_MinistryContract.show
         
                Appendix2
        Case 6
            Unload Me
         Case 7
         print_report2
   
   Case 8
             Appendix1
   Case 9
            Unload FrmSearch_MinistryContract
            FrmSearch_MinistryContract.SendForm = "attributioncontract"
            FrmSearch_MinistryContract.show
   Case 14
             Calculations
             
             Case 11
             RemoveGridRow1
             Case 12
                    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            
            Case 13
             VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
             
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub RemoveGridRow1()

    With Me.VSFlexGrid1

        If .Row <= 0 Then Exit Sub
    
         
                                                        
        .RemoveItem .Row

    End With
Cal_StudentCustom
'    ReLineGrid
End Sub

Private Sub AddYear()



'Dim i As Integer
'fg_Year.Rows = fg_Year.Rows + 1
'i = fg_Year.Rows
''i = i - 1
'With fg_Year
'  .TextMatrix(i, .ColIndex("Serial")) = i - 1
''  .TextMatrix(i, .ColIndex("Year")) = txtyear.text
'  .TextMatrix(i, .ColIndex("FromDate")) = DTPicker1.value
'  .TextMatrix(i, .ColIndex("ToDate")) = DTPicker2.value
'  .TextMatrix(i, .ColIndex("FromDateH")) = NourHijriCal1.value
'  .TextMatrix(i, .ColIndex("ToDateH")) = NourHijriCal2.value
'
'End With
'txtyear.text = ""

End Sub


Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments txtIDAC, "10062020001"

End Sub

Private Sub Command2_Click()

If VSFlexGrid1.Row < VSFlexGrid1.FixedRows Then Exit Sub
VSFlexGrid1.RemoveItem (VSFlexGrid1.Row)


'txtNet.text = val(txtTotal.text) + val(txtDiscount.text)
End Sub

Private Sub Command3_Click()
VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1

'txtNet.text = val(txtTotal.text) - val(txtDiscount.text)
End Sub

Private Sub cmdAdd_Click()
Operation = "add"
TxtDiscount.Enabled = True
End Sub

Private Sub cmdSub_Click()
Operation = "sub"
TxtDiscount.Enabled = True
End Sub

Private Sub Command1_Click()
VSFlexGrid1.Rows = 1

'FrmSearch_MinistryContract.SendForm = "AC"
'FrmSearch_MinistryContract.show
End Sub

Private Sub Fill_Cars()

VSFlexGrid1.Rows = 1
        Dim StrSQL As String
        Set Rs_Temp = New ADODB.Recordset
        StrSQL = "  SELECT  ID , BoardNo , ChasisNo  , [count] ,DriverName , DriverTel  ,count ,rate  from TblVendorCars  where  ( TblVendorCars.StopDeal is null  or dbo.TblVendorCars.StopDeal = 0 )  and  customerID = " & val(dcCustomer.BoundText) & "   ORDER BY ID "
        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
       Dim count As Double, Rate As Double
       
        Dim i As Integer
        If Rs_Temp.RecordCount > 0 Then
         Rs_Temp.MoveFirst
                With VSFlexGrid1
                    .Rows = Rs_Temp.RecordCount + 1
                    For i = 1 To Rs_Temp.RecordCount
                            
                           Board_Action IIf(IsNull(Rs_Temp("ID").value), "", Rs_Temp("ID").value), i
                                                        
                         '   .TextMatrix(i, .ColIndex("CarID")) = IIf(IsNull(Rs_Temp("ID").value), "", Rs_Temp("ID").value)
                         '   .TextMatrix(i, .ColIndex("BoardNo")) = IIf(IsNull(Rs_Temp("BoardNo").value), "", Rs_Temp("BoardNo").value)
                         '   .TextMatrix(i, .ColIndex("Chasis")) = IIf(IsNull(Rs_Temp("ChasisNo").value), "", Rs_Temp("ChasisNo").value)
                         '   .TextMatrix(i, .ColIndex("Driver")) = IIf(IsNull(Rs_Temp("DriverName").value), "", Rs_Temp("DriverName").value)
                         '   .TextMatrix(i, .ColIndex("Tel")) = IIf(IsNull(Rs_Temp("DriverTel").value), "", Rs_Temp("DriverTel").value)
                         '   '.TextMatrix(i, .ColIndex("capacity")) = IIf(IsNull(Rs_Temp("count").value), "", Rs_Temp("count").value)
                         '
                         '       .TextMatrix(i, .ColIndex("rate")) = IIf(IsNull(Rs_Temp("rate").value), "", Rs_Temp("rate").value)
                         '       .TextMatrix(i, .ColIndex("SitCount")) = IIf(IsNull(Rs_Temp("count").value), "", Rs_Temp("count").value)
                         '
                         '       count = IIf(IsNull(Rs_Temp("count").value), 0, Rs_Temp("count").value)
                         '       rate = IIf(IsNull(Rs_Temp("rate").value), 0, Rs_Temp("rate").value)
                         '
                         '      If rate = 0 Then
                         '       .TextMatrix(i, .ColIndex("capacity")) = count
                         '      Else
                         '       .TextMatrix(i, .ColIndex("capacity")) = rate * count
                         '      End If
                         '
                         '
                         '
                         '    If cbContractType.ListIndex = 1 Then
                         '   If opt_all.value = True Then
                         '           .TextMatrix(i, .ColIndex("dayRate")) = txtActualDayValue.text
                         ''   Else
                         '             .TextMatrix(i, .ColIndex("dayRate")) = ""
                         '   End If
                        'Else
                        '     .TextMatrix(i, .ColIndex("dayRate")) = ""
                        'End If
                        '
                        '
                        ' Dim carid As Integer, c As Integer, totcar
                        'carid = val(.TextMatrix(i, .ColIndex("CarID")))
                        'For c = 1 To VSFlexGrid1.Rows - 1
                        '         If val(.TextMatrix(i, .ColIndex("CarID"))) = carid Then
                        '                 totcar = totcar + val(.TextMatrix(c, .ColIndex("custom")))
                        '         End If
                        'Next
                        ' .TextMatrix(i, .ColIndex("VehicleAvailableSite")) = val(.TextMatrix(i, .ColIndex("capacity"))) - totcar
                        Rs_Temp.MoveNext
                    Next
                End With
        End If
        
   
                
End Sub

Private Sub dcCity_Click(Area As Integer)

Dim str As String
Set Rs_Temp = New ADODB.Recordset
Set dcMangerialAreaID.RowSource = Rs_Temp
If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
Else
    str = " Select ID , NameE   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
End If
fill_combo dcMangerialAreaID, str
dcMangerialAreaID.Refresh

End Sub

Private Sub dcCustomer_Change()
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
    
Fill_Cars
End Sub

Private Sub dcCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
 
If KeyCode = vbKeyF3 Then
        Unload FrmCompanySearch
        FrmCompanySearch.lblSearchtype = "2016"
        FrmCompanySearch.show vbModal
End If
 
End Sub

Private Sub dcDuration_Change()
If TxtModFlg.Text <> "R" Then
        Get_Resdent
        Cal_AcutualWorkDays
End If
End Sub

Public Sub Cal_AcutualWorkDays()
    Dim str As String, cunt As Integer
    str = " select count (*) cunt,DurationID    from TblVacationSchedule where ISVac = 0 and  DurationID =   " & val(dcDuration.BoundText) & "  and DateH >= '" & Format(dtpEmbarkH.value, "yyyy/MM/dd") & "'   group by DurationID "
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
           cunt = IIf(IsNull(Rs_Temp("cunt").value), 0, Rs_Temp("cunt").value)
    End If
    txtDaysCount.Text = cunt
End Sub

Private Sub dcDuration_Click(Area As Integer)
Calc_Installments
End Sub


Private Function GetHold(MonthID As Integer, Typ As Integer, Optional FromDateH As String, Optional FromDate As Date)
    Dim str As String, cunt As Integer
             
             If Typ = 0 Then
                        str = " select count (*)  cunt, DDID  from TblVacationSchedule where ISVac = 1 and  ddid =   " & MonthID & " and Date >= " & FromDate & "  group by DDID "
             ElseIf Typ = 1 Then
                        str = " select count (*)  cunt, DDID  from TblVacationSchedule where ISVac = 1 and  ddid =   " & MonthID & " and DateH >= '" & FromDateH & "'   group by DDID "
             End If
             
             Set RsTemp2 = New ADODB.Recordset
             RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
             
             If RsTemp2.RecordCount > 0 Then
                    cunt = IIf(IsNull(RsTemp2("cunt").value), 0, RsTemp2("cunt").value)
             End If
             
    GetHold = cunt
End Function
Private Function GetMonthDays(MonthID As Integer, Typ As Integer, Optional FromDateH As String, Optional FromDate As Date)

    Dim str As String, cunt As Integer
             
             If Typ = 0 Then
                    str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and Date >= " & FromDate & "  group by DDID "
             ElseIf Typ = 1 Then
                     str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and DateH >= '" & FromDateH & "'  group by DDID "
             End If
                    
             Set RsTemp2 = New ADODB.Recordset
             RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
             
             If RsTemp2.RecordCount > 0 Then
                    cunt = IIf(IsNull(RsTemp2("cunt").value), 0, RsTemp2("cunt").value)
                    
'                    If MonthID = 6 And mId(FromDateH, 1, 4) = 1441 Then
'                  cunt = cunt + 1
'                    End If
                    
             End If
             
    GetMonthDays = cunt

End Function

Private Sub dcDuration_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

  '     Unload FrmSearch_Duration
  '      FrmSearch_Duration.SendForm = "AC"
  '      FrmSearch_Duration.show
        
End If
End Sub

Private Sub dcMangerialAreaID_Click(Area As Integer)
Dim str As String

If Me.TxtModFlg = "R" Then
 
str = "update TblAttributionContract set MangerialAreaID=" & val(dcMangerialAreaID.BoundText) & "where IDAC=" & val(txtIDAC.Text)
Cn.Execute str
rs.Resync

End If

End Sub

Private Sub dcMinistry_Change()
Get_Resdent
Cal_AcutualWorkDays
End Sub

Private Sub dcMinistry_Click(Area As Integer)
Get_Resdent
End Sub

Private Sub dcMinistry_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

       Unload FrmSearch_MinistryContract
        FrmSearch_MinistryContract.SendForm = "AC"
        FrmSearch_MinistryContract.show
        
End If
End Sub

Private Sub dtpEContractDate_Change()
 If Me.TxtModFlg.Text <> "R" Then
        dtpEContractDateH.value = ToHijriDate(dtpEContractDate.value)
     End If
     txtDaysCount.Text = DateDiff("d", dtpSContractDate.value, dtpEContractDate)
End Sub

Private Sub dtpEContractDateH_GotFocus()
   VBA.Calendar = vbCalGreg
            dtpEContractDate.value = ToGregorianDate(dtpEContractDateH.value)
End Sub


Private Sub dtpEmbark_Change()


   dtpEmbarkH.value = ToHijriDate(dtpEmbark.value)
   
   If dtpEmbarkH.value < GetDurationStart() Then
             dtpEmbarkH.value = GetDurationStart
             VBA.Calendar = vbCalGreg
             dtpEmbark.value = ToGregorianDate(GetDurationStart)
       End If
   
   Cal_AcutualWorkDays
   Calc_Installments
End Sub

Private Sub dtpEmbarkH_LostFocus()
   VBA.Calendar = vbCalGreg
            dtpEmbark.value = ToGregorianDate(dtpEmbarkH.value)
            
            
         
       If dtpEmbarkH.value < GetDurationStart() Then
             dtpEmbarkH.value = GetDurationStart
                          VBA.Calendar = vbCalGreg
             dtpEmbark.value = ToGregorianDate(GetDurationStart)
       End If
            
   Cal_AcutualWorkDays
   Calc_Installments
End Sub

Private Sub dtpFromDate_Change()
     If Me.TxtModFlg.Text <> "R" Then
        dtpFromDateH.value = ToHijriDate(dtpFromDate.value)
     End If
End Sub



Private Sub dtpFromDateH_LostFocus()
 If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            dtpFromDate.value = ToGregorianDate(dtpFromDateH.value)
        End If
End Sub

Private Sub DTPicker1_Change()
    If Me.TxtModFlg.Text <> "R" Then
   '     NourHijriCal1.value = ToHijriDate(DTPicker1.value)
     End If
End Sub
Private Sub DTPicker2_Change()
     If Me.TxtModFlg.Text <> "R" Then
     '   NourHijriCal2.value = ToHijriDate(DTPicker2.value)
     End If
End Sub

Private Sub dtpSContractDate_Change()
     If Me.TxtModFlg.Text <> "R" Then
        dtpSContractDateH.value = ToHijriDate(dtpSContractDate.value)
     End If
     
     txtDaysCount.Text = DateDiff("d", dtpSContractDate.value, dtpEContractDate)
End Sub

Private Sub dtpSContractDateH_LostFocus()
On Error Resume Next
If Me.TxtModFlg.Text <> "R" Then
        VBA.Calendar = vbCalGreg
        dtpSContractDate.value = ToGregorianDate(dtpSContractDateH.value)
End If
        
End Sub

Private Sub dtpToDate_Change()
    If Me.TxtModFlg.Text <> "R" Then
        dtpToDateH.value = ToHijriDate(dtpToDate.value)
     End If
End Sub


Private Sub dtpToDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            dtpToDate.value = ToGregorianDate(dtpToDateH.value)
        End If
End Sub



Private Sub FgInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

        Cancel = True

End Sub

Private Sub FirstPaymentDate_Change()
 If Me.TxtModFlg.Text <> "R" Then
        FirstPaymentDateH.value = ToHijriDate(FirstPaymentDate.value)
     End If
End Sub



Private Sub FirstPaymentDateH_LostFocus()

 If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            FirstPaymentDate.value = ToGregorianDate(FirstPaymentDateH.value)
        End If

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
 '   On Error GoTo ErrTrap
 
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.getCountriesGovernments dcCity
    Dcombos.GetCustomersSuppliers 2, dcCustomer, , , , 1
    Dim str As String
    
    If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea "
    Else
    str = " Select ID , NameE   from TblManagerialArea "
    End If
    fill_combo dcMangerialAreaID, str
    
    str = "select id , name  from TblDurations "
    fill_combo dcDuration, str
    
    str = " select IDMC  , MinistryContractNo  from TblMinistryContract  "
    fill_combo dcMinistry, str
     
    Dcombos.GetBranches dcBranch
     
     With cbContractType
        If SystemOptions.UserInterface = ArabicInterface Then
                .Clear
                .AddItem ("„Œ’’ «·ÿ«·»")
                .AddItem ("«Ã— ÌÊ„Ï")
        Else
                .Clear
                .AddItem ("Student Custom")
                .AddItem ("Day  Salary")
        End If
    End With
  
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " ⁄ﬁœ «·«”‰«œ  "
    LogTexte = " Open Window " & "  Ministry Contract"
    
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""
    
    Dim My_SQL As String
    
    AddTip
    Set rs = New ADODB.Recordset
   Dim StrSQL As String
   StrSQL = "SELECT  *  From TblAttributionContract"
   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
        
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
 

dtpFromDate.value = Date
dtpSContractDate.value = Date
dtpToDate.value = Date
dtpEContractDate.value = Date
'DTPicker1.value = Date
'DTPicker2.value = Date
C1Tab1.CurrTab = 0
' Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture

ErrTrap:
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

 
   lbl(0).Caption = "No."
   lbl(3).Caption = " Name Ar"
   lbl(7).Caption = " Name En"
   Label3.Caption = "City"
   
  lbl(2).Caption = "Current Record"
  lbl(4).Caption = "Recors Count"
   
    Me.Caption = "Managerial Area"
    EleHeader.Caption = Me.Caption
   
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"

    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"


End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  ⁄ﬁœ «·«”‰«œ  "
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


Private Sub NourHijriCal3_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
         '   FirstPaymentDate.value = ToGregorianDate(NourHijriCal3.value)
        End If
End Sub

Private Sub Text1_Change()

End Sub



Private Sub opt_all_Click()
Dim i As Integer

With VSFlexGrid1
For i = 1 To .Rows - 1
        
              If cbContractType.ListIndex = 1 Then
                            If opt_all.value = True Then
                                    .TextMatrix(i, .ColIndex("dayRate")) = txtActualDayValue.Text
                            Else
                                      .TextMatrix(i, .ColIndex("dayRate")) = ""
                            End If
                        Else
                             .TextMatrix(i, .ColIndex("dayRate")) = ""
                        End If
        

Next
End With

End Sub

Private Sub opt_one_Click()
Dim i As Integer

With VSFlexGrid1
For i = 1 To .Rows - 1
        
              If cbContractType.ListIndex = 1 Then
                            If opt_all.value = True Then
                                    .TextMatrix(i, .ColIndex("dayRate")) = txtActualDayValue.Text
                            Else
                                      .TextMatrix(i, .ColIndex("dayRate")) = ""
                            End If
                        Else
                             .TextMatrix(i, .ColIndex("dayRate")) = ""
                        End If
        

Next
End With
End Sub

Private Sub txtActualDayValue_Change()
Dim i As Integer

With VSFlexGrid1
For i = 1 To .Rows - 1
        
              If cbContractType.ListIndex = 1 Then
                            If opt_all.value = True Then
                                    .TextMatrix(i, .ColIndex("dayRate")) = txtActualDayValue.Text
                         '   Else
                           '           .TextMatrix(i, .ColIndex("dayRate")) = ""
                            End If
                    '    Else
                         '    .TextMatrix(i, .ColIndex("dayRate")) = ""
                        End If
        

Next
End With
End Sub

Private Sub txtDaysCount_Change()

If cbContractType.ListIndex = 0 Then
    txtTotal.Text = val(txtStudentCount.Text) * val(txtStudentCustom.Text) * val(txtDaysCount.Text)
ElseIf cbContractType.ListIndex = 1 Then
    txtTotal.Text = val(txtDaysCount.Text) * val(txtdayvalue.Text)
End If

End Sub

Private Sub txtDayValue_Change()
On Error Resume Next
If cbContractType.ListIndex = 0 Then
ElseIf cbContractType.ListIndex = 1 Then
        txtTotal.Text = val(txtDaysCount.Text) * val(txtdayvalue.Text)
End If

 Dim i As Integer
    With FgInstallments
        For i = 1 To val(TxtPaymentCount.Text)
              .TextMatrix(i, .ColIndex("Value")) = Round(val(txtdayvalue.Text) * val(.TextMatrix(i, .ColIndex("ADays"))), 2)
        Next
    End With

End Sub

Private Sub txtDiscount_Change()

If Operation = "add" Then
    TxtNet.Text = val(txtTotal.Text) + val(TxtDiscount.Text)
ElseIf Operation = "sub" Then
    TxtNet.Text = val(txtTotal.Text) - val(TxtDiscount.Text)
End If

End Sub

Private Sub Get_Resdent()
    Dim str As String
    Dim alloc As Double, StudentCount As Double, attr As Double
    
   If dcDuration.BoundText = "" Or dcMinistry.BoundText = "" Then
   Exit Sub
   End If
    
    
   ' str = "select StudentCount  from TblMinistryContract where ProcessNo =  '" & txtMinistryContractNo.text & "'"
    str = "select StudentCount  from TblMinistryContract where IDMC =  " & val(dcMinistry.BoundText)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        StudentCount = IIf(IsNull(Rs_Temp("StudentCount").value), 0, Rs_Temp("StudentCount").value)
    End If
    
    str = " select * from TblVehicleAllocation where IDMC = " & val(dcMinistry.BoundText) & " and DurationID =  " & val(dcDuration.BoundText)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        alloc = IIf(IsNull(Rs_Temp("StudentAlloc").value), 0, Rs_Temp("StudentAlloc").value)
        Else
        alloc = 0
    End If
    
    'str = " select  sum (COALESCE (studentcount , 0)) sumstudent from TblAttributionContract where MinistryContractNo =   '" & txtMinistryContractNo.text & "'"
    str = " select  sum (COALESCE (studentcount , 0)) sumstudent from TblAttributionContract where IDMC = " & val(dcMinistry.BoundText) & " and DurationID =  " & val(dcDuration.BoundText)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        attr = IIf(IsNull(Rs_Temp("sumstudent").value), 0, Rs_Temp("sumstudent").value)
        Else
        attr = 0
    End If
    
    txtRes.Text = StudentCount - alloc - attr
    
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
    
Fill_Cars
End Sub

Private Sub txtMinistryContractNo_Change()

Get_Resdent

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "⁄ﬁœ «·«”‰«œ"
            Else
                Me.Caption = "Ministry Contract"
            End If
txtIDAC.Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(9).Enabled = True
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
            
            C1Elastic3.Enabled = False
            C1Elastic5.Enabled = False
            'C1Tab1.Enabled = False
            
        Case "N"
txtIDAC.Enabled = True
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ⁄ﬁœ «·«”‰«œ ( ÃœÌœ )"
            Else
                Me.Caption = "Boxes Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "⁄ﬁœ «·«”‰«œ"
            Else
                Me.Caption = "Ministry Contarct"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            txtProcessNo.locked = False
            Me.XPTxtBoxName.locked = False
            
            C1Elastic3.Enabled = True
            C1Elastic5.Enabled = True
            'C1Tab1.Enabled = True
        Case "E"
txtIDAC.Enabled = False
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ⁄ﬁœ «·«”‰«œ (  ⁄œÌ· )"
            Else
                Me.Caption = "Boxes Data(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            txtProcessNo.locked = True
            Me.XPTxtBoxName.locked = False
       '     Me.XPMTxtRemark.locked = False
            
            C1Elastic3.Enabled = True
            C1Elastic5.Enabled = True
            'C1Tab1.Enabled = True
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
            rs.find "IDAC =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    If IsNull(rs("AdditionalType").value) Or rs("AdditionalType").value = "" Then
        TxtDiscount.Enabled = False
    Else
        Operation = rs("AdditionalType").value
        TxtDiscount.Enabled = True
    End If
    
    FgInstallments.Rows = 1
    
    txtIDAC.Text = IIf(IsNull(rs("IDAC").value), "", rs("IDAC").value)
    dcMinistry.BoundText = IIf(IsNull(rs("IDMC").value), "", rs("IDMC").value)
    
    dcCustomer.BoundText = IIf(IsNull(rs("VendorID").value), "", rs("VendorID").value)
    txtProcessNo.Text = txtIDAC.Text
    XPTxtBoxName.Text = IIf(IsNull(rs("Name").value), "", Trim(rs("Name").value))
     
     'dcVendor.BoundText = IIf(IsNull(rs("VendorID").value), "", rs("VendorID").value)
     dcCity.BoundText = IIf(IsNull(rs("CityID").value), "", rs("CityID").value)
     
    dtpFromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    dtpToDate.value = IIf(IsNull(rs("ToDate").value), Date, rs("ToDate").value)
    dtpSContractDate.value = IIf(IsNull(rs("StartContractDate").value), Date, rs("StartContractDate").value)
    dtpEContractDate.value = IIf(IsNull(rs("EndContractDate").value), Date, rs("EndContractDate").value)
    dtpFromDateH.value = IIf(IsNull(rs("FromDateh").value), Date, rs("FromDateh").value)
    dtpToDateH.value = IIf(IsNull(rs("ToDateh").value), Date, rs("ToDateh").value)
    dtpSContractDateH.value = IIf(IsNull(rs("StartContractDateh").value), Date, rs("StartContractDateh").value)
    dtpEContractDateH.value = IIf(IsNull(rs("EndContractDateh").value), Date, rs("EndContractDateh").value)
    txtStudentCustom.Text = IIf(IsNull(rs("StudentCustom").value), "", Trim(rs("StudentCustom").value))
    txtStudentCount.Text = IIf(IsNull(rs("StudentCount").value), "", Trim(rs("StudentCount").value))
    lblCalStudent.Text = IIf(IsNull(rs("StudentCount").value), "", Trim(rs("StudentCount").value))
    TxtDiscount.Text = IIf(IsNull(rs("Discount").value), "", Trim(rs("Discount").value))
    dcDuration.BoundText = IIf(IsNull(rs("DurationID").value), "", Trim(rs("DurationID").value))
    txtdayvalue.Text = IIf(IsNull(rs("DayValue").value), "", rs("DayValue").value)
    dcMangerialAreaID.BoundText = IIf(IsNull(rs("MangerialAreaID").value), "", rs("MangerialAreaID").value)
    
     txtTotal.Text = val(txtStudentCustom.Text) * val(txtStudentCount.Text) * val(txtDaysCount.Text)
     TxtNet.Text = IIf(IsNull(rs("NetValue").value), "0", rs("NetValue").value)
     dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
   
     dtpEmbark.value = IIf(IsNull(rs("HEmbarkDate").value), ToGregorianDate(GetDurationStart), rs("HEmbarkDate").value)
     dtpEmbarkH.value = IIf(IsNull(rs("HEmbarkDateH").value), GetDurationStart, rs("HEmbarkDateH").value)
      
      
    Dim s As Integer
     s = IIf(IsNull(rs("contracttype").value), -1, rs("contracttype").value)
     cbContractType.ListIndex = s
     
     If Operation = "add" Then
          TxtNet.Text = val(txtTotal.Text) + val(TxtDiscount.Text)
     ElseIf Operation = "sub" Then
          TxtNet.Text = val(txtTotal.Text) - val(TxtDiscount.Text)
     End If
     
     TxtPaymentCount.Text = IIf(IsNull(rs("PaymentCount").value), "", Trim(rs("PaymentCount").value))
     txtMinistryContractNo.Text = IIf(IsNull(rs("MinistryContractNo").value), "", rs("MinistryContractNo").value)
    
    Set rsInst = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = " SELECT * from TblMinistryContract_Installment where type = 2 and  idmc =  " & val(txtIDAC.Text)
                  FgInstallments.Clear flexClearScrollable, flexClearEverything
            FgInstallments.Rows = 1
            
    rsInst.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rsInst.RecordCount > 0 Then
    rsInst.MoveFirst
     With FgInstallments
        FgInstallments.Rows = rsInst.RecordCount + 1
        Dim j As Integer
        For j = 1 To FgInstallments.Rows - 1
                .TextMatrix(j, .ColIndex("QestID")) = IIf(IsNull(rsInst("InstallmentNo").value), "", rsInst("InstallmentNo").value)
                .TextMatrix(j, .ColIndex("value")) = IIf(IsNull(rsInst("Value").value), 0, rsInst("Value").value)
                .TextMatrix(j, .ColIndex("MonthID")) = IIf(IsNull(rsInst("MonthID").value), "", rsInst("MonthID").value)
                .TextMatrix(j, .ColIndex("Due_Date")) = IIf(IsNull(rsInst("Due_Date").value), Date, rsInst("Due_Date").value)
                .TextMatrix(j, .ColIndex("Due_DateH")) = IIf(IsNull(rsInst("Due_DateH").value), "", rsInst("Due_DateH").value)
                .TextMatrix(j, .ColIndex("ToDate")) = IIf(IsNull(rsInst("ToDate").value), "", rsInst("ToDate").value)
                .TextMatrix(j, .ColIndex("ToDateH")) = IIf(IsNull(rsInst("ToDateH").value), "", rsInst("ToDateH").value)
                .TextMatrix(j, .ColIndex("ADays")) = IIf(IsNull(rsInst("ActiveDays").value), "", rsInst("ActiveDays").value)
                .TextMatrix(j, .ColIndex("VRID")) = IIf(IsNull(rsInst("VRID").value), "", rsInst("VRID").value)
                
                rsInst.MoveNext
         Next
        End With
    End If
     
     
       Set rsVendor = New ADODB.Recordset
       Dim str As String
       str = " select * from TblVehicleAllocation_Details where type = 3 and idva =  " & val(txtIDAC.Text)
       rsVendor.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
       
                              VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
       With VSFlexGrid1
            .Rows = 1
            .Rows = rsVendor.RecordCount + 1
            For j = 1 To .Rows - 1
            
                    .TextMatrix(j, .ColIndex("serial")) = j
                   .TextMatrix(j, .ColIndex("CarID")) = IIf(IsNull(rsVendor("CarID").value), "", rsVendor("CarID").value)
                   .TextMatrix(j, .ColIndex("Chasis")) = IIf(IsNull(rsVendor("Chasis").value), "", rsVendor("Chasis").value)
                   .TextMatrix(j, .ColIndex("BoardNo")) = IIf(IsNull(rsVendor("BoardNo").value), "", rsVendor("BoardNo").value)
                   '.TextMatrix(j, .ColIndex("DriverID")) = IIf(IsNull(rsVendor("DriverID").value), "", rsVendor("DriverID").value)
                   .TextMatrix(j, .ColIndex("Driver")) = IIf(IsNull(rsVendor("Driver").value), "", rsVendor("Driver").value)
                   .TextMatrix(j, .ColIndex("BoardNo")) = IIf(IsNull(rsVendor("BoardNo").value), "", rsVendor("BoardNo").value)
                   .TextMatrix(j, .ColIndex("custom")) = IIf(IsNull(rsVendor("StudentCount").value), "", rsVendor("StudentCount").value)
                   .TextMatrix(j, .ColIndex("MAID")) = IIf(IsNull(rsVendor("MangerialAreaID").value), "", rsVendor("MangerialAreaID").value)
                   .TextMatrix(j, .ColIndex("MA")) = IIf(IsNull(rsVendor("MangerialArea").value), "", rsVendor("MangerialArea").value)
                   .TextMatrix(j, .ColIndex("MAID")) = IIf(IsNull(rsVendor("MangerialAreaID").value), "", rsVendor("MangerialAreaID").value)
                   .TextMatrix(j, .ColIndex("MA")) = IIf(IsNull(rsVendor("MangerialArea").value), "", rsVendor("MangerialArea").value)
                   .TextMatrix(j, .ColIndex("schoolfileID")) = IIf(IsNull(rsVendor("schoolfileID").value), "", rsVendor("schoolfileID").value)
                   .TextMatrix(j, .ColIndex("schoolfile")) = IIf(IsNull(rsVendor("schoolfile").value), "", rsVendor("schoolfile").value)
                   
                   .TextMatrix(j, .ColIndex("Tel")) = IIf(IsNull(rsVendor("drivertel").value), "", rsVendor("drivertel").value)
                   .TextMatrix(j, .ColIndex("rate")) = IIf(IsNull(rsVendor("Rate").value), "", rsVendor("Rate").value)
                   .TextMatrix(j, .ColIndex("ministerno")) = IIf(IsNull(rsVendor("SchoolMinistryno").value), "", rsVendor("SchoolMinistryno").value)
                   
                   .TextMatrix(j, .ColIndex("count")) = IIf(IsNull(rsVendor("SchoolStudentCount").value), "", rsVendor("SchoolStudentCount").value)
                   .TextMatrix(j, .ColIndex("allow")) = IIf(IsNull(rsVendor("SchoolStudentAvailable").value), "", rsVendor("SchoolStudentAvailable").value)
                   .TextMatrix(j, .ColIndex("dayRate")) = IIf(IsNull(rsVendor("DayRate").value), "", rsVendor("DayRate").value)
                    
                   .TextMatrix(j, .ColIndex("studentcustom")) = IIf(IsNull(rsVendor("StudentCustom").value), "", rsVendor("StudentCustom").value)
                   .TextMatrix(j, .ColIndex("SitCount")) = IIf(IsNull(rsVendor("VehicleSiteCount").value), "", rsVendor("VehicleSiteCount").value)
                   .TextMatrix(j, .ColIndex("VehicleAvailableSite")) = IIf(IsNull(rsVendor("VehicleAvailableSite").value), "", rsVendor("VehicleAvailableSite").value)
                   .TextMatrix(j, .ColIndex("capacity")) = IIf(IsNull(rsVendor("capecity").value), "", rsVendor("capecity").value)
                    
                    .TextMatrix(j, .ColIndex("Embark")) = IIf(IsNull(rsVendor("DEmbarkDate").value), "", rsVendor("DEmbarkDate").value)
                   .TextMatrix(j, .ColIndex("EmbarkH")) = IIf(IsNull(rsVendor("DEmbarkDateH").value), "", rsVendor("DEmbarkDateH").value)
                   rsVendor.MoveNext
            Next
       End With
     
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub txtNet_Change()
    On Error Resume Next
   
    txtActualDayValue.Text = Round(val(TxtNet.Text) / val(txtDaysCount.Text), 2)
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
    
Fill_Cars
End Sub

Private Sub txtStudentCount_LostFocus()
 If val(txtStudentCount.Text) > val(txtRes.Text) Then txtStudentCount.Text = "": MsgBox ("·«Ì„ﬂ‰ «‰ Ì Ã«Ê“ ⁄œœ «·ÿ«·» «·»«ﬁÏ „‰  Œ’Ì’ «·⁄ﬁœ "): Exit Sub

If cbContractType.ListIndex = 0 Then
        txtTotal.Text = val(txtStudentCount.Text) * val(txtStudentCustom.Text) * val(txtDaysCount.Text)
End If
End Sub

Private Sub txtStudentCustom_Change()
txtTotal.Text = val(txtStudentCount.Text) * val(txtStudentCustom.Text) * val(txtDaysCount.Text)
Dim i As Integer

With VSFlexGrid1
    For i = 1 To .Rows - 1
        .TextMatrix(i, .ColIndex("studentcustom")) = txtStudentCustom.Text
        .TextMatrix(i, .ColIndex("dayRate")) = val(.TextMatrix(i, .ColIndex("custom"))) * val(.TextMatrix(i, .ColIndex("studentcustom")))
                            cal_DayRate
    Next
End With

End Sub

Private Sub TxtTotal_Change()
TxtNet.Text = val(txtTotal.Text)
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim StrAccountCode As String
Dim Msg As String
'  Dim rs As New ADODB.Recordset
Dim StrSQL As String
Dim ClsAcc As New ClsAccounts
Dim LngRow As Long
Dim sql As String
Dim count As Integer
Dim Rate As Double
 
    With VSFlexGrid1

Select Case .ColKey(Col)

             Case "BoardNo"
                        
                        .TextMatrix(Row, .ColIndex("CarNo")) = 0
                        .TextMatrix(Row, .ColIndex("Chasis")) = 0
                        .TextMatrix(Row, .ColIndex("Remarks")) = ""
                        .TextMatrix(Row, .ColIndex("Driver")) = ""
                        .TextMatrix(Row, .ColIndex("Count")) = ""
                        
                        StrAccountCode = .ComboData
                        Board_Action StrAccountCode, .Row
                        
          Case "schoolfile"
                         SchoolFile_INfo .ComboData, .Row
                        
          Case "custom"
                    Dim cus As Integer, Capacity As Integer, allow As Integer, all   As Double, asas As String
           '         cus = val(.TextMatrix(.Row, .ColIndex("custom")))
                                     
           '         capacity = val(.TextMatrix(.Row, .ColIndex("VehicleAvailableSite")))
           '         allow = val(.TextMatrix(.Row, .ColIndex("allow")))
           '
           '         If cus > capacity Then
           '
           '             If capacity < 0 Then
           '                 cus = 0
           '                 .TextMatrix(.Row, .ColIndex("custom")) = 0
           '             Else
           '                 cus = capacity
           '                 .TextMatrix(.Row, .ColIndex("custom")) = capacity
           '             End If
           '         End If
                           
           '         If cus > allow Then
           '                 .TextMatrix(.Row, .ColIndex("custom")) = allow
           '                 cus = allow
           '         End If
                    
                    Dim j As Integer
           '         For j = 1 To .Rows - 1
           '                     all = all + val(.TextMatrix(j, .ColIndex("custom")))
           '         Next
           '
           '         If all > val(txtRes.Text) Then
           '               MsgBox ("·«Ì„ﬂ‰  Ã«Ê“ «·Õœ «·„”„ÊÕ »Â ·· Œ’Ì’ ··⁄ﬁœ")
           '                 .TextMatrix(.Row, .ColIndex("custom")) = ""
           '         End If
                    
                
                    Cal_StudentCustom
                    
                    If cbContractType.ListIndex = 0 Then
                            .TextMatrix(.Row, .ColIndex("dayRate")) = val(.TextMatrix(.Row, .ColIndex("custom"))) * val(.TextMatrix(.Row, .ColIndex("studentcustom")))
                            cal_DayRate
                    End If
            Case "MA"
                      StrAccountCode = .ComboData
                     .TextMatrix(Row, .ColIndex("maid")) = StrAccountCode
                     
            Case "ministerno"
            
                        Dim mm As String
                        StrAccountCode = .TextMatrix(.Row, .ColIndex("ministerno"))
                        StrSQL = "  select * from TblSchooleFile  where ministerno = '" & (StrAccountCode) & "'"
                        Set Rs_Temp = New ADODB.Recordset
                        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                        If Not (Rs_Temp.BOF Or Rs_Temp.EOF) Then
                                SchoolFile_INfo IIf(IsNull(Rs_Temp("ID").value), "", Rs_Temp("ID").value), .Row
                        End If
                                    
                Case "VehicleAvailableSite"
                
                 Case "dayRate"
                        If cbContractType.ListIndex = 1 Then
                                If opt_one.value = True Then
                                       cal_DayRate
                                End If
                        End If
                        
                   Case "studentcustom"
                            If cbContractType.ListIndex = 0 Then
                                    .TextMatrix(.Row, .ColIndex("dayRate")) = val(.TextMatrix(.Row, .ColIndex("custom"))) * val(.TextMatrix(.Row, .ColIndex("studentcustom")))
                                    cal_DayRate
                            End If
                            
                  Case "EmbarkH"
                        mDate
                            
           End Select
    End With



End Sub

Private Sub Board_Action(StrAccountCode As String, Row As Integer)
Dim StrSQL As String, count As Double, Rate As Double

 With VSFlexGrid1


                        StrSQL = "   select * from TblVendorCars  where customerid =  " & val(dcCustomer.BoundText) & "and id =  " & val(StrAccountCode)
                        Set Rs_Temp5 = New ADODB.Recordset
                        Rs_Temp5.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                                    
                        If Not (Rs_Temp5.BOF Or Rs_Temp5.EOF) Then
                                 .TextMatrix(Row, .ColIndex("BoardNo")) = IIf(IsNull(Rs_Temp5("boardno").value), "", Rs_Temp5("boardno").value)
                                .TextMatrix(Row, .ColIndex("CarID")) = IIf(IsNull(Rs_Temp5("ID").value), "", Rs_Temp5("ID").value)
                                .TextMatrix(Row, .ColIndex("Chasis")) = IIf(IsNull(Rs_Temp5("ChasisNo").value), "", Rs_Temp5("ChasisNo").value)
                                .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(Rs_Temp5("rate").value), "", Rs_Temp5("rate").value)
                                .TextMatrix(Row, .ColIndex("SitCount")) = IIf(IsNull(Rs_Temp5("count").value), "", Rs_Temp5("count").value)
                                .TextMatrix(Row, .ColIndex("Driver")) = IIf(IsNull(Rs_Temp5("DriverName").value), "", Rs_Temp5("DriverName").value)
                                .TextMatrix(Row, .ColIndex("Tel")) = IIf(IsNull(Rs_Temp5("DriverTel").value), "", Rs_Temp5("DriverTel").value)
                                 count = IIf(IsNull(Rs_Temp5("count").value), 0, Rs_Temp5("count").value)
                                 Rate = IIf(IsNull(Rs_Temp5("rate").value), 0, Rs_Temp5("rate").value)
                                 .TextMatrix(Row, .ColIndex("capacity")) = Rate * count
                                 
                                 .TextMatrix(Row, .ColIndex("Embark")) = dtpEmbark.value
                                 .TextMatrix(Row, .ColIndex("EmbarkH")) = dtpEmbarkH.value
                        End If
                        
                        If cbContractType.ListIndex = 1 Then
                            If opt_all.value = True Then
                                    .TextMatrix(Row, .ColIndex("dayRate")) = txtActualDayValue.Text
                            Else
                                   ' .TextMatrix(row, .ColIndex("dayRate")) = ""
                            End If
                        Else
'                                    .TextMatrix(row, .ColIndex("dayRate")) = ""
                        End If
                         
                        Dim CarID As Integer, c As Integer, totcar
                        CarID = val(.TextMatrix(Row, .ColIndex("CarID")))
                        
                        For c = 1 To VSFlexGrid1.Rows - 1
                                 If val(.TextMatrix(c, .ColIndex("CarID"))) = CarID Then
                                         totcar = totcar + val(.TextMatrix(c, .ColIndex("custom")))
                                 End If
                        Next
                     
                       
                       ' StrSQL = "   select * from TblVendorCars  where customerid =  " & val(dcCustomer.BoundText) & "and id =  " & val(StrAccountCode)
                       ' Set Rs_Temp = New ADODB.Recordset
                       ' Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                       
                        
                        Dim sumation As Double
                        StrSQL = ""
                        StrSQL = " select  d.CarID  , sum( d.studentcount ) sumation  ,v.DurationID  from TblAttributionContract   v , TblVehicleAllocation_Details d   "
                        StrSQL = StrSQL & "  where v.IDAC  = d.IDVA  and  type = 3 and   d.carid = " & CarID & "  and v.DurationID = " & val(dcDuration.BoundText) & " and v.IDMC =  " & val(dcMinistry.BoundText)
                        If TxtModFlg.Text = "E" Then
                                StrSQL = StrSQL & "  and v.IDAC <>  " & val(txtIDAC.Text)
                        End If
                        
                        StrSQL = StrSQL & "  group by   d.CarID ,v.DurationID  "
                        Set Rs_Temp5 = New ADODB.Recordset
                        Rs_Temp5.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        If Rs_Temp5.RecordCount > 0 Then
                                sumation = IIf(IsNull(Rs_Temp5("sumation").value), 0, Rs_Temp5("sumation").value)
                        End If
                        
                        .TextMatrix(Row, .ColIndex("VehicleAvailableSite")) = Int(val(.TextMatrix(Row, .ColIndex("capacity"))) - totcar - sumation)

        End With

End Sub


Public Sub SchoolFile_INfo(SchoolFileID As String, Row As Integer)
Dim StrAccountCode As String, StrSQL As String
            With VSFlexGrid1
            StrAccountCode = SchoolFileID
 
            StrSQL = "   select * from TblSchooleFile where id =  " & val(StrAccountCode)
            Set Rs_Temp = New ADODB.Recordset
            Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (Rs_Temp.BOF Or Rs_Temp.EOF) Then
                    .TextMatrix(Row, .ColIndex("schoolfileid")) = IIf(IsNull(Rs_Temp("ID").value), "", Rs_Temp("ID").value)
                    .TextMatrix(Row, .ColIndex("count")) = IIf(IsNull(Rs_Temp("studentcount").value), "", Rs_Temp("studentcount").value)
                    .TextMatrix(Row, .ColIndex("ministerno")) = IIf(IsNull(Rs_Temp("ministerNo").value), "", Rs_Temp("ministerNo").value)
                    .TextMatrix(Row, .ColIndex("schoolfile")) = IIf(IsNull(Rs_Temp("name").value), "", Rs_Temp("name").value)
            End If
            
            
            Dim tot As Integer, totSchl As Integer, j As Integer
            For j = 1 To .Rows - 1
                    If .TextMatrix(j, .ColIndex("schoolfileid")) = StrAccountCode Then
                            totSchl = totSchl + val(.TextMatrix(j, .ColIndex("custom")))
                    End If
            Next
                                    
                        
            ' StrSQL = "  select sum(studentalloc) tot from TblVehicleAllocation  where type= 3 and   schoolfileid   = " & val(StrAccountCode)
             StrSQL = ""
             StrSQL = StrSQL & "  select h.DurationID , h.IDMC , d.SchoolFileID   , sum ( d.studentcount) tot"
             StrSQL = StrSQL & "  from  TblAttributionContract h ,TblVehicleAllocation_Details d"
             StrSQL = StrSQL & "  where  h.idac = d.IDVA and Type = 3"
             StrSQL = StrSQL & "  and d.SchoolFileID = " & val(StrAccountCode) & " and h.IDMC = " & val(dcMinistry.BoundText) & " and h.DurationID = " & val(dcDuration.BoundText)
             If TxtModFlg.Text = "E" Then
                    StrSQL = StrSQL & " and h.IDAC <>  " & val(txtIDAC.Text)
             End If
             StrSQL = StrSQL & "  group by h.DurationID , h.IDMC , d.SchoolFileID"
            
            
             Set Rs_Temp = New ADODB.Recordset
             Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
             If Not (Rs_Temp.BOF Or Rs_Temp.EOF) Then
                    tot = IIf(IsNull(Rs_Temp("tot").value), 0, Rs_Temp("tot").value)
             End If
              .TextMatrix(Row, .ColIndex("allow")) = val(.TextMatrix(Row, .ColIndex("count"))) - tot - totSchl
             
             
            If cbContractType.ListIndex = 0 Then
                      .TextMatrix(Row, .ColIndex("studentcustom")) = txtStudentCustom.Text
            End If
                       
End With


End Sub

Private Sub cal_DayRate(Optional SCustom = False)

Dim i  As Integer, tot As Double

'If cbContractType.ListIndex = 1 Then
'        If opt_one.value = True Then
            With VSFlexGrid1
            For i = 1 To .Rows - 1
                    tot = tot + val(.TextMatrix(i, .ColIndex("dayRate")))
            Next
            End With
            txtdayvalue.Text = tot
'        End If
' End If

End Sub


Private Sub Cal_StudentCustom()

Dim i  As Integer, tot As Integer
With VSFlexGrid1
For i = 1 To .Rows - 1
        tot = tot + val(.TextMatrix(i, .ColIndex("custom")))
Next
End With
lblCalStudent.Text = tot
End Sub


Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid1
     
    If TxtModFlg.Text = "R" Then
               .ComboList = ""
          Cancel = True
          Exit Sub
    End If
     
   
        
        
    If dcMinistry.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«Œ — «·⁄ﬁœ «·Ê“«—Ï  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "   ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
              Cancel = True
                Exit Sub
               
       End If
           
        If dcDuration.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«Œ — «·”‰… «·œ—«”Ì…  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "select Duration   ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
                dcDuration.SetFocus
                SendKeys ("{F4}")
                Cancel = True
                Exit Sub
               
       End If
         
          If dcCustomer.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«Œ — «·„ ⁄Âœ  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "select Duration   ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
                dcCustomer.SetFocus
                SendKeys ("{F4}")
                Cancel = True
                Exit Sub
               
       End If
         
         
       If cbContractType.ListIndex = -1 Then
            MsgBox ("«œŒ· ‰Ê⁄ «· ⁄«ﬁœ")
            Exit Sub
            Cancel = True
        End If
    
       
     
     
     
     
     
     Select Case .ColKey(Col)
     
    Case "CarNo"
         .ComboList = ""
          Cancel = True
    Case "Chasis"
            .ComboList = ""
            Cancel = True
    Case "count"
        .ComboList = ""
         Cancel = True
    Case "Remarks"
            .ComboList = ""
      Cancel = True
    Case "Tel"
            .ComboList = ""
             Cancel = True
        Case "model"
            .ComboList = ""
             Cancel = True
                Case "Driver"
            .ComboList = ""
             Cancel = True
                Case "capacity"
            .ComboList = ""
             Cancel = True
             
                Case "count"
            .ComboList = ""
            Cancel = True
            
                Case "Remarks"
            .ComboList = ""
            
            
              Case "allow"
            .ComboList = ""
            Cancel = True
            
                Case "custom"
            .ComboList = ""
            
           SchoolFile_INfo VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, .ColIndex("Schoolfileid")), VSFlexGrid1.Row
           Board_Action VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, .ColIndex("CarID")), VSFlexGrid1.Row
            
             Case "studentcustom"
            .ComboList = ""
           'Cancel = True
                     
           
           
             Case "dayRate"
            
            .ComboList = ""
            If cbContractType.ListIndex = 1 Then
                If opt_all.value = True Then
                        Cancel = True
                End If
            ElseIf cbContractType.ListIndex = 0 Then
                    Cancel = True
             End If
            
            
            'Cancel = True
           Case "SitCount"
            .ComboList = ""
            Cancel = True
           Case "rate"
             .ComboList = ""
            Cancel = True
            
            
           Case "ministerno"
           
                   .ComboList = ""
            'Cancel = True
                Case "VehicleAvailableSite"
                Cancel = True
            
           Case "Embark"
                    Cancel = True
        
           Case "EmbarkH"
                    Cancel = True
           
            
            
            End Select
            
            End With
End Sub

Private Sub VSFlexGrid1_Click()

If TxtModFlg.Text <> "R" Then
        Select Case VSFlexGrid1.ColKey(VSFlexGrid1.Col)
        Case "Embark"
                    Unload FrmRegesterDateProject
                    FrmRegesterDateProject.SendForm = "AttributionContract"
                    FrmRegesterDateProject.show vbModal
        Case "EmbarkH"
                    Unload FrmRegesterDateProject
                    FrmRegesterDateProject.SendForm = "AttributionContract"
                    FrmRegesterDateProject.show vbModal
        End Select
End If

End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

       If VSFlexGrid1.Col = VSFlexGrid1.ColIndex("schoolfile") Then
                   If KeyCode = vbKeyF3 Then
                            Unload FrmSearch_BasicData
                             FrmSearch_BasicData.SendForm = "AC2"
                             Load FrmSearch_BasicData
                             FrmSearch_BasicData.show vbModal
                   End If
        End If

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

'Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    
     With VSFlexGrid1
     
     Select Case .ColKey(Col)
     
    
     Case "BoardNo"
          
         Set Rs_Temp = New ADODB.Recordset
         StrSQL = "  SELECT  ID ,BoardNo  from TblVendorCars  where  customerID = " & val(dcCustomer.BoundText) & "   ORDER BY ID "
          Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          StrComboList = VSFlexGrid1.BuildComboList(Rs_Temp, "BoardNo", "ID")
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList
          
       Case "schoolfile"
          Set Rs_Temp = New ADODB.Recordset
          StrSQL = "  select id , name  from TblSchooleFile  "
          Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          StrComboList = VSFlexGrid1.BuildComboList(Rs_Temp, "name", "ID")
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList
                            
         Case "MA"
            Set Rs_Temp = New ADODB.Recordset
            StrSQL = " select id , name  from TblManagerialArea    "
            Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            StrComboList = VSFlexGrid1.BuildComboList(Rs_Temp, "name", "ID")
             If StrComboList <> "" Then
                   StrComboList = "|" & StrComboList
             End If
            .ComboList = StrComboList
         
           
     End Select
   End With
   
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
    Cal_AcutualWorkDays
    Exit Sub
ErrTrap:
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    
       LogTextA = "  Õ›Ÿ ‘«‘… " & "   ⁄ﬁÊœ «·«”‰«œ  " & txtIDAC.Text _
       & CHR(13) & "   ⁄›œ «·Ê“«—Â —ﬁ„  " & dcMinistry.Text _
       & CHR(13) & "    »«ﬁÌ «· Œ’Ì’  " & txtRes.Text _
       & CHR(13) & "    «—ÌŒ «·ÌÊ„     " & dtpFromDate.value _
       & CHR(13) & " „”„” «· ⁄«ﬁœ     " & XPTxtBoxName.Text _
       & CHR(13) & " «·⁄«„ «·œ—«”Ì     " & dcDuration.Text _
       & CHR(13) & "  «—ÌŒ »œ«Ì… «·⁄›œ „     " & dtpSContractDate.value _
       & CHR(13) & "  «—ÌŒ »œ«Ì… «·⁄›œ Â‹   " & dtpSContractDateH.value _
       & CHR(13) & "  «—ÌŒ ‰Â«Ì… «·⁄›œ „   " & dtpEContractDate.value _
       & CHR(13) & "  «—ÌŒ ‰Â«Ì… «·⁄›œ Â‹   " & dtpSContractDateH.value _
       & CHR(13) & "  «·ﬂÊœ         " & TxtFullcode.Text _
       & CHR(13) & "  —ﬁ„ «·”Ã·         " & TxtRecordNo.Text _
    & CHR(13) & "    «·„ ⁄Âœ       " & dcCustomer.Text _
      & CHR(13) & "    «·„Õ«›Ÿ…       " & dcCity.Text _
     & CHR(13) & "  «·«œ«—… «· ⁄·Ì„Ì…       " & dcMangerialAreaID.Text _
     & CHR(13) & "    ‰Ê⁄ «· ⁄«ﬁœ       " & cbContractType.Text _
       & CHR(13) & "   ⁄œœ «·ÿ·«»  " & txtStudentCount.Text _
       & CHR(13) & "      ⁄œœ «Ì«„ «·⁄„· «·›⁄·Ì…  " & txtDaysCount.Text _
       & CHR(13) & "   ﬁÌ„Â «·ÌÊ„  " & txtdayvalue.Text _
       & CHR(13) & "   «·«Ã„«·Ì  " & txtTotal.Text _
       & CHR(13) & "   «·’«›Ì  " & TxtNet.Text _
       & CHR(13) & "   «·ﬁÌ„Â «·›⁄·Ì… ··ÌÊ„  " & txtActualDayValue.Text _

       
         

    If Currentmode <> "D" Then
       ' AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, "", ""
        AddToLogFile CInt(user_id), , Now, Time, LogTextA, LogTexte, Me.Name, "E", "", "", val(txtIDAC), txtIDAC.Text
    Else
     AddToLogFile CInt(user_id), , Now, Time, LogTextA, LogTexte, Me.Name, "D", "", "", val(txtIDAC), txtIDAC.Text
       ' AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D", "", ""
    End If
    
   
   
End Function
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap
    Dim str As String
    
    'function check
    
    If Me.TxtModFlg.Text <> "R" Then
    
        If XPTxtBoxName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                 MsgBox "„‰ ›÷·ﬂ √œŒ· „”„Ï «· ⁄«ﬁœ  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                 MsgBox "Please Entrer Contract Name  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            XPTxtBoxName.SetFocus
            Exit Sub
        End If
        
        If cbContractType.ListIndex = -1 Then
            MsgBox ("«œŒ· ‰Ê⁄ «· ⁄«ﬁœ")
            Exit Sub
        End If
        
            If dcBranch.BoundText = "" Then
             MsgBox ("„‰ ›÷··ﬂ «Œ — «·›—⁄ «Ê·« ")
             Exit Sub
        End If
        
        
        If dcMinistry.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«Œ — «·⁄ﬁœ «·Ê“«—Ï  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "   ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
              
                Exit Sub
               
       End If
           
        If dcDuration.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«Œ — «·› —…  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "select Duration   ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
                dcDuration.SetFocus
                SendKeys ("{F4}")
                Exit Sub
               
       End If
    
         
            If val(txtStudentCount.Text) = 0 Then
                        MsgBox "«œŒ· ⁄œœ «·ÿ·«» "
                     Exit Sub
            End If
     
            If cbContractType.ListIndex = 0 Then
                If val(txtStudentCustom.Text) = 0 Then
                     MsgBox "«œŒ· „Œ’’ ﬂ· ÿ«·»"
                     Exit Sub
                End If
            Else
                If val(txtdayvalue.Text) = 0 Then
                        MsgBox "«œŒ· ﬁÌ„… «·ÌÊ„"
                        Exit Sub
                End If
            End If
         
         
       If val(TxtPaymentCount.Text) <= 0 Then
            MsgBox "«œŒ· ⁄œœ «·œ›⁄« "
            'TxtPaymentCount.SetFocus
            Exit Sub
       End If
       
       If val(txtStudentCount.Text) <> val(lblCalStudent.Text) Then
            MsgBox "⁄œœ «·ÿ·«» «·–Ì‰  „  ”ﬂÌ‰Â„ ·« Ì”«ÊÏ ⁄œœ «·ÿ·«» "
            'TxtPaymentCount.SetFocus
            Exit Sub
       End If
       
        
        Select Case Me.TxtModFlg.Text
            Case "N"
                'StrSQL = "select * From  TblAttributionContract  where Name ='" & Trim(XPTxtBoxName.text) & "'"
                'RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                'If RsTemp.RecordCount > 0 Then
                '    Msg = "Â‰«ﬂ   ⁄«ﬁœ  „”Ã· „”»ﬁ« »Â–« «·„”„Ï" & Chr(13)
                '    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                '    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·⁄ﬁœ "
                '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                '    XPTxtBoxName.SetFocus
                '    Exit Sub
                'End If
            
             rs.AddNew
          '   txtIDAC.Text = CStr(new_id("TblAttributionContract", "IDAC", "", True))
            Case "E"
              '  StrSQL = "select * From  TblAttributionContract where Name='" & Trim(XPTxtBoxName.text) & "'"
              '  RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

              '  If RsTemp.RecordCount > 0 Then
              '      If RsTemp("IDaC").value <> val(txtIDAC.text) Then
              '          Msg = "Â‰«ﬂ  ⁄«ﬁœ  „”Ã·Â „”»ﬁ« »Â–« «·„”„Ï" & Chr(13)
              '          Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·„”„Ï «·’ÕÌÕ " & Chr(13)
              '          Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·„”„Ï "
              '          MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
              '          XPTxtBoxName.SetFocus
              '          Exit Sub
              '      End If
              '  End If
        End Select

        Cn.BeginTrans
        BeginTrans = True
          
        Select Case Me.TxtModFlg.Text
            Case "N"
                If txtIDAC.Text = "" Then
                 txtIDAC.Text = CStr(new_id("TblAttributionContract", "IDAC", "", True))
                 End If
                 
            Case "E"
                StrSQL = "delete from TblMinistryContract_Installment where type=2 and idmc =  " & val(txtIDAC.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                 StrSQL = "delete from TblVehicleAllocation_Details where type =3 and IDVA =" & val(txtIDAC.Text)
                 Cn.Execute StrSQL, , adExecuteNoRecords
                
                 StrSQL = " delete from TblAttributionInstallmentDivided where  IDAC =" & val(txtIDAC.Text)
                 Cn.Execute StrSQL, , adExecuteNoRecords
                
        End Select
               
       
        rs("IDAC").value = val(txtIDAC.Text)
        rs("IDMC").value = val(dcMinistry.BoundText)
        
        rs("VendorID").value = val(dcCustomer.BoundText)
        
        rs("ProcessNo").value = txtIDAC.Text
        rs("Name").value = IIf(Trim(XPTxtBoxName.Text) = "", Null, XPTxtBoxName.Text)
        
        rs("FromDate").value = dtpFromDate.value
        rs("ToDate").value = dtpToDate.value
        rs("FromDateH").value = dtpFromDateH.value
        rs("ToDateH").value = dtpToDateH.value
          
        rs("StartContractDate").value = dtpSContractDate.value
        rs("EndContractDate").value = dtpEContractDate.value
        rs("StartContractDateH").value = dtpSContractDateH.value
        rs("EndContractDateH").value = dtpEContractDateH.value
        rs("CityID").value = IIf(dcCity.BoundText = "", Null, dcCity.BoundText)
        'rs("VendorID").value = IIf(dcVendor.BoundText = "", Null, dcVendor.BoundText)
       
        rs("StudentCount").value = IIf(txtStudentCount.Text = "", 0, val(txtStudentCount.Text))
        rs("StudentCustom").value = IIf(txtStudentCustom.Text = "", 0, val(txtStudentCustom.Text))
        rs("DisCount").value = IIf(TxtDiscount.Text = "", 0, val(TxtDiscount.Text))
        rs("PaymentCount").value = IIf(TxtPaymentCount.Text = "", 0, val(TxtPaymentCount.Text))
        
        rs("FirstPaymentDate").value = FirstPaymentDate.value
        rs("AdditionalType").value = Operation
        rs("MinistryContractNo").value = txtMinistryContractNo.Text
        rs("DurationID").value = IIf(dcDuration.BoundText = "", Null, dcDuration.BoundText)
        
        rs("DaysCount").value = val(txtDaysCount.Text)
        rs("DayValue").value = val(txtdayvalue.Text)
        rs("ActualDayValue").value = val(txtActualDayValue.Text)
        rs("ContractType").value = cbContractType.ListIndex
        rs("NetValue").value = val(TxtNet.Text)
        rs("MangerialAreaID").value = val(dcMangerialAreaID.BoundText)
        rs("BranchID").value = val(dcBranch.BoundText)
        rs("HEmbarkDate").value = dtpEmbark.value
        rs("HEmbarkDateH").value = dtpEmbarkH.value
        
        rs.update
          
'          Cmd_Click (5)
        Dim rsIns As ADODB.Recordset
        Set rsIns = New ADODB.Recordset
        StrSQL = "select * from TblMinistryContract_Installment"
        rsIns.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        With FgInstallments
        Dim j As Integer
       ' FgInstallments.Rows = val(TxtPaymentCount.text) + 1
        
        For j = 1 To FgInstallments.Rows - 1
           If .TextMatrix(j, .ColIndex("QestID")) <> "" Then
                    rsIns.AddNew
                    rsIns("ID") = CStr(new_id("TblMinistryContract_Installment", "ID", "", True))
                    rsIns("IDMC") = val(txtIDAC.Text)
                    rsIns("InstallmentNo") = .TextMatrix(j, .ColIndex("QestID"))
                    rsIns("Value") = .TextMatrix(j, .ColIndex("value"))
                    rsIns("MonthID") = .TextMatrix(j, .ColIndex("MonthID"))
                    rsIns("Due_Date") = .TextMatrix(j, .ColIndex("Due_Date"))
                    rsIns("Due_DateH") = .TextMatrix(j, .ColIndex("Due_DateH"))
                    rsIns("ToDate") = .TextMatrix(j, .ColIndex("ToDate"))
                    rsIns("ToDateH") = .TextMatrix(j, .ColIndex("ToDateH"))
                    rsIns("ActiveDays").value = IIf(.TextMatrix(j, .ColIndex("ADays")) = "", 0, (.TextMatrix(j, .ColIndex("ADays"))))
                    rsIns("Type") = 2
                    rsIns.update
                 End If
           Next
        End With
        
     
        
        
        
        Dim rs_det As ADODB.Recordset, DetailsID As String, d As Integer
        Set rs_det = New ADODB.Recordset
        rs_det.Open "TblVehicleAllocation_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        With VSFlexGrid1
         For j = 1 To .Rows - 1
            If .TextMatrix(j, .ColIndex("BoardNo")) <> "" Then
                rs_det.AddNew
                DetailsID = new_id("TblVehicleAllocation_Details", "ID", "", True)
                rs_det("ID").value = DetailsID
                rs_det("IDVA").value = val(txtIDAC.Text)
                rs_det("CarID").value = IIf(.TextMatrix(j, .ColIndex("CarID")) = "", Null, .TextMatrix(j, .ColIndex("CarID")))
                rs_det("Chasis").value = IIf(.TextMatrix(j, .ColIndex("Chasis")) = "", Null, .TextMatrix(j, .ColIndex("Chasis")))
                rs_det("BoardNo").value = IIf(.TextMatrix(j, .ColIndex("BoardNo")) = "", Null, .TextMatrix(j, .ColIndex("BoardNo")))
               ' rs_det("DriverID").value = IIf(.TextMatrix(j, .ColIndex("DriverID")) = "", Null, .TextMatrix(j, .ColIndex("DriverID")))
                rs_det("Driver").value = IIf(.TextMatrix(j, .ColIndex("Driver")) = "", Null, .TextMatrix(j, .ColIndex("Driver")))
                rs_det("schoolfileID").value = IIf(.TextMatrix(j, .ColIndex("schoolfileID")) = "", Null, (.TextMatrix(j, .ColIndex("schoolfileID"))))
                rs_det("schoolfile").value = IIf(.TextMatrix(j, .ColIndex("schoolfile")) = "", "", (.TextMatrix(j, .ColIndex("schoolfile"))))
                rs_det("capecity").value = IIf(.TextMatrix(j, .ColIndex("capacity")) = "", 0, (.TextMatrix(j, .ColIndex("capacity"))))
                rs_det("MangerialAreaID").value = IIf(.TextMatrix(j, .ColIndex("MAID")) = "", Null, (.TextMatrix(j, .ColIndex("MAID"))))
                rs_det("MangerialArea").value = IIf(.TextMatrix(j, .ColIndex("MA")) = "", "", (.TextMatrix(j, .ColIndex("MA"))))
                rs_det("DriverTel").value = IIf(.TextMatrix(j, .ColIndex("Tel")) = "", "", (.TextMatrix(j, .ColIndex("Tel"))))
                rs_det("Rate").value = IIf(.TextMatrix(j, .ColIndex("rate")) = "", 0, (.TextMatrix(j, .ColIndex("rate"))))
                rs_det("SchoolMinistryno").value = IIf(.TextMatrix(j, .ColIndex("ministerno")) = "", "", (.TextMatrix(j, .ColIndex("ministerno"))))
                rs_det("SchoolStudentCount").value = IIf(.TextMatrix(j, .ColIndex("count")) = "", 0, (.TextMatrix(j, .ColIndex("count"))))
                rs_det("SchoolStudentAvailable").value = IIf(.TextMatrix(j, .ColIndex("allow")) = "", 0, (.TextMatrix(j, .ColIndex("allow"))))
                'rs_det("Custom").value = IIf(.TextMatrix(j, .ColIndex("MA")) = "", "", (.TextMatrix(j, .ColIndex("MA"))))
                rs_det("DayRate").value = IIf(.TextMatrix(j, .ColIndex("dayRate")) = "", 0, (.TextMatrix(j, .ColIndex("dayRate"))))
                rs_det("StudentCustom").value = IIf(.TextMatrix(j, .ColIndex("studentcustom")) = "", 0, (.TextMatrix(j, .ColIndex("studentcustom"))))
                rs_det("VehicleSiteCount").value = IIf(.TextMatrix(j, .ColIndex("SitCount")) = "", 0, (.TextMatrix(j, .ColIndex("SitCount"))))
                rs_det("VehicleAvailableSite").value = IIf(.TextMatrix(j, .ColIndex("VehicleAvailableSite")) = "", 0, (.TextMatrix(j, .ColIndex("VehicleAvailableSite"))))
                rs_det("Type").value = 3
                rs_det("StudentCount").value = IIf(.TextMatrix(j, .ColIndex("custom")) = "", 0, (.TextMatrix(j, .ColIndex("custom"))))
                rs_det("DEmbarkDate").value = IIf(.TextMatrix(j, .ColIndex("Embark")) = "", 0, (.TextMatrix(j, .ColIndex("Embark"))))
                rs_det("DEmbarkDateH").value = IIf(.TextMatrix(j, .ColIndex("EmbarkH")) = "", 0, (.TextMatrix(j, .ColIndex("EmbarkH"))))
           
                
                rs_det.update
                
                Set rsDivid = New ADODB.Recordset
                rsDivid.Open " TblAttributionInstallmentDivided  ", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                
                 For d = 1 To FgInstallments.Rows - 1
                            If FgInstallments.TextMatrix(d, FgInstallments.ColIndex("QestID")) <> "" Then
                                rsDivid.AddNew
                                rsDivid("ID") = CStr(new_id("TblAttributionInstallmentDivided", "ID", "", True))
                                rsDivid("IDAC") = val(txtIDAC.Text)
                                rsDivid("DurationID") = val(dcDuration.BoundText)
                                rsDivid("BoardNo") = .TextMatrix(j, .ColIndex("BoardNo"))
                                rsDivid("MonthID") = FgInstallments.TextMatrix(d, FgInstallments.ColIndex("MonthID"))
                                rsDivid("DetailsID") = DetailsID
                                rsDivid("DDEmbarkDate").value = IIf(.TextMatrix(j, .ColIndex("Embark")) = "", 0, (.TextMatrix(j, .ColIndex("Embark"))))
                                rsDivid("DDEmbarkDateH").value = IIf(.TextMatrix(j, .ColIndex("EmbarkH")) = "", 0, (.TextMatrix(j, .ColIndex("EmbarkH"))))
                                rsDivid.update
                             End If
                    Next
                
                
                
            End If
         Next
        End With
               
        
          
       
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·⁄ﬁœ  " & CHR(13)
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

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "IDAC ='" & val(txtIDAC.Text) & "'", , adSearchForward, adBookmarkFirst

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
            
        If txtIDAC.Text <> "" Then
    
        Msg = "”Ì „ Õ–› »Ì«‰«  ⁄ﬁœ «·«”‰«œ —ﬁ„ " & CHR(13)
        Msg = Msg + (txtIDAC.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
    
            If Not rs.RecordCount < 1 Then
                If ISAllowDeleteUpdateContract(val(txtIDAC.Text)) = False Then
                        MsgBox ("·«Ì„ﬂ‰ Õ–›  «·⁄ﬁœ „‰ «Ã·  ﬂ«„· «·»Ì«‰«  ")
                        Exit Sub
                End If
                
               '  strSQL = " delete from TblAttributionInstallmentDivided where  IDAC =" & val(txtIDAC.text)
               '  Cn.Execute strSQL, , adExecuteNoRecords
                 
                StrSQL = "delete From TblAttributionContract where  IDAC =" & val(txtIDAC.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                 StrSQL = "delete from TblMinistryContract_Installment where type=2 and idmc =  " & val(txtIDAC.Text)
                 Cn.Execute StrSQL, , adExecuteNoRecords
                
                   StrSQL = "delete from TblVehicleAllocation_Details where type = 3 and IDVA =" & val(txtIDAC.Text)
                   Cn.Execute StrSQL, , adExecuteNoRecords
CuurentLogdata ("D")

                   StrSQL = "SELECT  *  From TblAttributionContract "
                   Set rs = New ADODB.Recordset
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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… «·⁄ﬁœ "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Public Function ISAllowDeleteUpdateContract(ID As Integer) As Boolean
Dim str As String

str = " Select *  from TblConfirmViolation  where MinistryContractID =  " & ID
Set Rs_Temp5 = New ADODB.Recordset
Rs_Temp5.Open str, Cn, adOpenStatic, adLockOptimistic
If Rs_Temp5.RecordCount > 0 Then
        ISAllowDeleteUpdateContract = False
        Exit Function
End If


str = " select * from TblMinistryContract_Installment where  VR_paid = 1 and type = 2 and idmc=    " & ID
Set Rs_Temp5 = New ADODB.Recordset
Rs_Temp5.Open str, Cn, adOpenStatic, adLockOptimistic
If Rs_Temp5.RecordCount > 0 Then
        ISAllowDeleteUpdateContract = False
        Exit Function
End If


str = "   SELECT * from TblAttributionInstallmentDivided where   re_paid = 1  and   idac =  " & ID
Set Rs_Temp5 = New ADODB.Recordset
Rs_Temp5.Open str, Cn, adOpenStatic, adLockOptimistic
If Rs_Temp5.RecordCount > 0 Then
        ISAllowDeleteUpdateContract = False
        Exit Function
End If


ISAllowDeleteUpdateContract = True

End Function


Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ «·«”‰«œ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄ﬁœ «·«”‰«œ  ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ «·«”‰«œ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ⁄ﬁœ «·«”‰«œ " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ «·«”‰«œ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄ﬁœ «·«”‰«œ  «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ «·«”‰«œ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ «·«”‰«œ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« ⁄ﬁœ «·«”‰«œ  " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ «·«”‰«œ  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄ﬁœ Ê“«—… " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ «·«”‰«œ  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ «·«”‰«œ  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄⁄ﬁœ «·«”‰«œ  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄⁄ﬁœ «·«”‰«œ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ «·«”‰«œ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ «·«”‰«œ  ", 1, 15204351, -2147483630
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


Private Sub Calculations(Optional WithMsg As Boolean = True)
'    On Error GoTo ErrTrap
    Dim SngAllValue As Single
    Dim i  As Integer
    Dim DateInterval, Msg As String
    Dim NewDateH As String
    Dim NewDate As String
    Dim PreDateH As String
 
 If IsNumeric(TxtPaymentCount.Text) Then
    If Not (val(TxtPaymentCount.Text) > 0) Then
            MsgBox ("««œŒ· ⁄œœ «·œ›⁄«  «Ê·« ")
            TxtPaymentCount.SetFocus
            Exit Sub
    End If
 Else '
    Exit Sub
 End If
 
 If DcbPeriodsID.ListIndex = -1 Then
 MsgBox (" «œŒ· «·› —… »Ì‰ «·œ›⁄« ")
 Exit Sub
 End If
 
  
   If DcbPeriodsID.ListIndex = 0 Then
        DateInterval = "d"
    ElseIf DcbPeriodsID.ListIndex = 1 Then
        DateInterval = "M"
    ElseIf DcbPeriodsID.ListIndex = 2 Then
        DateInterval = "yyyy"
        Else
        DateInterval = "D"
        
    End If
    
    DtatAdd.value = FirstPaymentDate.value
    With Me.FgInstallments
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + val(TxtPaymentCount.Text)

        For i = 1 To .Rows - 1

            .TextMatrix(i, .ColIndex("QestID")) = i
         '   .TextMatrix(i, .ColIndex("Value")) = Round(val(txtNet.text) / val(TxtPaymentCount.text), 2)
            
           .TextMatrix(i, .ColIndex("Value")) = Round(val(txtdayvalue.Text) * val(.TextMatrix(i, .ColIndex("ADays"))), 2)
            
          If i = 1 Then
            .TextMatrix(i, .ColIndex("Due_DateH")) = Format(FirstPaymentDateH.value, "yyyy/MM/dd")
             .TextMatrix(i, .ColIndex("Due_Date")) = ToGregorianDate(FirstPaymentDateH.value)
             Else
             PreDateH = (Trim(.TextMatrix(i - 1, .ColIndex("Due_DateH"))))
             NewDateH = (DateAdd(DateInterval, val(TxtPeriods.Text), PreDateH))
             NewDate = ToGregorianDate(NewDateH)
             DtatAdd.value = DateAdd((DateInterval), val(TxtPeriods.Text), DtatAdd.value)
             .TextMatrix(i, .ColIndex("Due_Date")) = NewDate
             .TextMatrix(i, .ColIndex("Due_DateH")) = Format(NewDateH, "yyyy/MM/dd")
             End If
         Next i
         
      
         
         .AutoSize 1, .Cols - 1, False
         End With
    Exit Sub
ErrTrap:
End Sub

Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = ""


MySQL = MySQL & "             SELECT dbo.TblAttributionContract.IDAC, dbo.TblAttributionContract.Name, dbo.TblAttributionContract.FromDate, dbo.TblAttributionContract.FromDateH,"
MySQL = MySQL & "             dbo.TblAttributionContract.ToDate, dbo.TblAttributionContract.ToDateH, dbo.TblAttributionContract.VendorID, dbo.TblAttributionContract.CityID,"
MySQL = MySQL & "             dbo.TblAttributionContract.StudentCount, dbo.TblAttributionContract.StudentCustom, dbo.TblAttributionContract.DisCount, dbo.TblAttributionContract.MinistryContractNo,"
MySQL = MySQL & "             dbo.TblVehicleAllocation_Details.CarID, dbo.TblVehicleAllocation_Details.Driver, dbo.TblVehicleAllocation_Details.SchoolFileID, dbo.TblVehicleAllocation_Details.SchoolFile,"
 MySQL = MySQL & "            dbo.TblVehicleAllocation_Details.SchoolMinistryno, dbo.TblVehicleAllocation_Details.SchoolStudentCount, dbo.TblVehicleAllocation_Details.SchoolStudentAvailable,"
MySQL = MySQL & "             dbo.TblVehicleAllocation_Details.Custom, dbo.TblVehicleAllocation_Details.DayRate, dbo.TblVehicleAllocation_Details.MaxCap,"
MySQL = MySQL & "             dbo.TblVehicleAllocation_Details.VehicleAvailableSite, dbo.TblVehicleAllocation_Details.VehicleSiteCount, dbo.TblVehicleAllocation_Details.Capecity,"
MySQL = MySQL & "             dbo.TblVehicleAllocation_Details.DriverID, dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblAttributionContract.MangerialAreaID,"
MySQL = MySQL & "             dbo.TblManagerialArea.Name AS TblManagerialAreaName, dbo.TblManagerialArea.Namee AS TblManagerialAreaNameE,"
MySQL = MySQL & "             dbo.TblVehicleAllocation_Details.StudentCount AS StudentCountD, dbo.TblVehicleAllocation_Details.StudentCustom AS StudentCustomD,"
MySQL = MySQL & "             dbo.TblDurations.Name AS DurationName, dbo.TblAttributionContract.BranchID, dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "             dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.EndContractDateh, dbo.TblAttributionContract.StartContractDate,"
MySQL = MySQL & "             dbo.TblAttributionContract.StartContractDateh, dbo.TblAttributionContract.NetValue, dbo.TblAttributionContract.ContractType, dbo.TblAttributionContract.ActualDayValue,"
MySQL = MySQL & "             dbo.TblAttributionContract.DayValue, dbo.TblAttributionContract.DaysCount, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.CusID,"
MySQL = MySQL & "             dbo.TblCountriesGovernments.GovernmentName , dbo.TblCustemers.fullcode, dbo.TblCustemers.recordno, dbo.TblCustemers.cus_phone   , dbo.TblCustemers.cus_mobile  , dbo.TblAttributionContract.HEmbarkDate, dbo.TblAttributionContract.HEmbarkDateH  "
MySQL = MySQL & "             FROM     dbo.TblAttributionContract INNER JOIN"
MySQL = MySQL & "             dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA INNER JOIN"
MySQL = MySQL & "             dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID INNER JOIN"
MySQL = MySQL & "             dbo.TblCustemers ON dbo.TblAttributionContract.vendorid = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "             dbo.TblBranchesData ON dbo.TblAttributionContract.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "             dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID LEFT OUTER JOIN"
MySQL = MySQL & "             dbo.TblCountriesGovernments ON dbo.TblAttributionContract.CityID = dbo.TblCountriesGovernments.GovernmentID"


MySQL = MySQL & "  where    dbo.TblAttributionContract.IDAC  =" & val(txtIDAC.Text)
MySQL = MySQL & " order by IDAC "

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_AttributionContract.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_AttributionContract.rpt"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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
 

Function Appendix2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = ""

     MySQL = MySQL & "  SELECT dbo.TblAttributionContract.IDAC, dbo.TblVehicleAllocation_Details.Type, dbo.TblVehicleAllocation_Details.ID, dbo.TblVehicleAllocation_Details.StudentCount,"
     MySQL = MySQL & "    dbo.TblVehicleAllocation_Details.Chasis, dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblVehicleAllocation_Details.CarID, dbo.TblSchooleFile.Name,"
     MySQL = MySQL & "     dbo.TblVehicleAllocation_Details.SchoolFileID, dbo.TblVehicleAllocation_Details.DayRate, dbo.TblAttributionContract.FromDate, dbo.TblAttributionContract.FromDateH,"
     MySQL = MySQL & "    dbo.TblVehicleAllocation_Details.Capecity"
     MySQL = MySQL & "    FROM     dbo.TblAttributionContract INNER JOIN"
     MySQL = MySQL & "    dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA INNER JOIN"
     MySQL = MySQL & "    dbo.TblSchooleFile ON dbo.TblVehicleAllocation_Details.SchoolFileID = dbo.TblSchooleFile.ID"
     MySQL = MySQL & "    Where (dbo.TblVehicleAllocation_Details.Type = 3)"

    MySQL = MySQL & "  and    dbo.TblAttributionContract.IDAC  =" & val(txtIDAC.Text)



 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Appendix.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Appendix.rpt"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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
 



Function Appendix1(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = ""

    MySQL = MySQL & "      SELECT dbo.TblAttributionContract.IDAC, dbo.TblCountriesGovernments.GovernmentName, dbo.TblManagerialArea.Name AS Expr1,"
    MySQL = MySQL & "      SUM(dbo.TblVehicleAllocation_Details.StudentCount) AS sum_student, COUNT(dbo.TblVehicleAllocation_Details.SchoolFileID) AS count_school,"
    MySQL = MySQL & "      COUNT(dbo.TblVehicleAllocation_Details.CarID) As count_car"
    MySQL = MySQL & "      FROM     dbo.TblAttributionContract INNER JOIN"
    MySQL = MySQL & "      dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA INNER JOIN"
    MySQL = MySQL & "      dbo.TblSchooleFile ON dbo.TblVehicleAllocation_Details.SchoolFileID = dbo.TblSchooleFile.ID INNER JOIN"
    MySQL = MySQL & "      dbo.TblCountriesGovernments ON dbo.TblAttributionContract.CityID = dbo.TblCountriesGovernments.GovernmentID INNER JOIN"
    MySQL = MySQL & "      dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID"
    MySQL = MySQL & "      Where (dbo.TblVehicleAllocation_Details.Type = 3)"

    MySQL = MySQL & "  and    dbo.TblAttributionContract.IDAC  =" & val(txtIDAC.Text)

    MySQL = MySQL & "      GROUP BY dbo.TblAttributionContract.IDAC, dbo.TblCountriesGovernments.GovernmentName, dbo.TblManagerialArea.Name"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Appendix1.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Appendix1.rpt"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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

Public Sub cars(Row As Integer)
Dim i As Integer, j As Integer
i = 0
Dim CarID As Integer
CarID = val(VSFlexGrid1.TextMatrix(Row, VSFlexGrid1.ColIndex("CarID")))

With VSFlexGrid1
For j = 1 To .Rows - 1
    If .TextMatrix(j, .ColIndex("CarID")) = CarID Then
        i = i + val(.TextMatrix(j, .ColIndex("count")))
    End If
Next
End With
End Sub

Public Sub mDate()
    Dim i As Integer, curr As String, j As Integer

    With VSFlexGrid1
        curr = .TextMatrix(.Row, .ColIndex("EmbarkH"))
        
        For i = 1 To .Rows - 1
            For j = i To .Rows - 1
                If .TextMatrix(j, .ColIndex("EmbarkH")) <> "" And .TextMatrix(j, .ColIndex("EmbarkH")) < curr Then
                         curr = .TextMatrix(j, .ColIndex("EmbarkH"))
                End If
            Next
        Next
        
    End With

  '  If curr > dtpEmbarkH.value Then
            dtpEmbarkH.value = curr
            VBA.Calendar = vbCalGreg
            dtpEmbark.value = ToGregorianDate(curr)
            Calc_Installments
            Cal_AcutualWorkDays
  '  End If

End Sub

Public Function GetDurationStart() As String
    
    Dim i  As Integer, str As String
    i = val(dcDuration.BoundText)
    str = " select * from tbldurations where id =  " & i
    Set Rs_Temp5 = New ADODB.Recordset
   Rs_Temp5.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
   If Rs_Temp5.RecordCount > 0 Then
                GetDurationStart = IIf(IsNull(Rs_Temp5("fromdateh").value), "", Rs_Temp5("fromdateh").value)
   End If
       
    
End Function




