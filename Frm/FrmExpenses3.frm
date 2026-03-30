VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmExpenses3 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›« Ê—… „«·Ì…"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13215
   HelpContextID   =   280
   Icon            =   "FrmExpenses3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   13215
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic3 
      Height          =   8910
      Left            =   0
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   0
      Width           =   13215
      _cx             =   23310
      _cy             =   15716
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
      Begin VB.TextBox txtManulaVat 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   142
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txt_Currency_rate 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   139
         Text            =   "1"
         Top             =   2865
         Width           =   960
      End
      Begin VB.TextBox TxtVATCustoms 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   136
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ComboBox CBoBasedON 
         Height          =   315
         Left            =   3600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   134
         Top             =   2520
         Width           =   2055
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   1095
         Left            =   0
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   7800
         Width           =   13215
         _cx             =   23310
         _cy             =   1931
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
         Begin VB.TextBox TxtSerial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5040
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   690
            Width           =   1905
         End
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   255
            Left            =   11520
            TabIndex        =   80
            Top             =   120
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "„—«ﬂ“ «· ﬂ·›…"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   255
            BCOLO           =   192
            FCOL            =   16777215
            FCOLO           =   0
            MCOL            =   192
            MPTR            =   1
            MICON           =   "FrmExpenses3.frx":038A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   7800
            TabIndex        =   81
            Top             =   720
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   10680
            TabIndex        =   83
            Top             =   120
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   873
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   1
            Left            =   9720
            TabIndex        =   84
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   2
            Left            =   8910
            TabIndex        =   85
            Top             =   120
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   3
            Left            =   7755
            TabIndex        =   86
            Top             =   120
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   4
            Left            =   6840
            TabIndex        =   87
            Top             =   120
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   6
            Left            =   2880
            TabIndex        =   88
            Top             =   120
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton CmdHelp 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   3720
            TabIndex        =   89
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”ŒÂ „„«À·…"
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   5
            Left            =   5790
            TabIndex        =   90
            Top             =   120
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   8
            Left            =   4800
            TabIndex        =   91
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   375
            Index           =   10
            Left            =   3720
            TabIndex        =   93
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   375
            Left            =   2640
            TabIndex        =   94
            Top             =   600
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   661
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
         Begin ImpulseButton.ISButton Accredit 
            Height          =   315
            Left            =   0
            TabIndex        =   100
            Top             =   600
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            ButtonPositionImage=   1
            Caption         =   "«—”«· ··«⁄ „«œ"
            BackColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   -2147483635
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   7
            Left            =   1380
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   120
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   435
            Index           =   6
            Left            =   570
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   240
            Width           =   165
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   435
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   120
            Width           =   525
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   435
            Left            =   780
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   255
            Left            =   6960
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   270
            Index           =   8
            Left            =   12225
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   735
            Width           =   900
         End
      End
      Begin VB.TextBox txt_general_des 
         Alignment       =   1  'Right Justify
         Height          =   645
         Left            =   60
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   76
         Top             =   3240
         Width           =   5595
      End
      Begin VB.TextBox txt_ORDER_NO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   2880
         Width           =   2115
      End
      Begin VB.TextBox txtto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4110
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   72
         Top             =   2130
         Width           =   1575
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   2295
         Left            =   7200
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2040
         Width           =   6015
         _cx             =   10610
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
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3930
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   840
            Width           =   705
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3930
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3930
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   120
            Width           =   705
         End
         Begin VB.TextBox TxtChequeNumber 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   1200
            Width           =   4485
         End
         Begin MSComCtl2.DTPicker DtpChequeDueDate 
            Height          =   315
            Left            =   150
            TabIndex        =   59
            Top             =   1620
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   556
            _Version        =   393216
            Format          =   149618689
            CurrentDate     =   39614
         End
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   150
            TabIndex        =   60
            Top             =   840
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   120
            TabIndex        =   61
            Top             =   480
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCVendor 
            Height          =   315
            Left            =   120
            TabIndex        =   62
            Top             =   120
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê—œ"
            Height          =   285
            Index           =   22
            Left            =   5070
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
            Height          =   285
            Index           =   19
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1620
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   285
            Index           =   18
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·»‰ﬂ"
            Height          =   285
            Index           =   17
            Left            =   4710
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·Œ“Ì‰…"
            Height          =   285
            Index           =   16
            Left            =   4710
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   540
            Width           =   1215
         End
      End
      Begin VB.ComboBox CboPaymentType 
         Height          =   315
         Left            =   7260
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox TxtNoteserial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7260
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1230
         Width           =   1335
      End
      Begin VB.ComboBox CboPaymentType1 
         Height          =   315
         Left            =   10260
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   1230
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         DataField       =   "id"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   960
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1080
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox TxtOrderID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1080
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.TextBox TxtSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10260
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   840
         Width           =   1455
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   765
         Left            =   0
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   0
         Width           =   13215
         _cx             =   23310
         _cy             =   1349
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
         Picture         =   "FrmExpenses3.frx":03A6
         Caption         =   "›« Ê—… „«·Ì…  "
         Align           =   0
         AutoSizeChildren=   7
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
         Begin VB.TextBox oldTxtSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8040
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox ChkPurchaseFixedAssets 
            Alignment       =   1  'Right Justify
            Caption         =   "›« Ê—… ‘—«¡ «’·"
            Height          =   195
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   0
            Visible         =   0   'False
            Width           =   1935
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   375
            Index           =   0
            Left            =   1695
            TabIndex        =   31
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmExpenses3.frx":1080
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
            TabIndex        =   32
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmExpenses3.frx":141A
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
            Index           =   1
            Left            =   2220
            TabIndex        =   33
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmExpenses3.frx":17B4
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
            TabIndex        =   34
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmExpenses3.frx":1B4E
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin MSAdodcLib.Adodc numbering 
            Height          =   585
            Left            =   4680
            Top             =   0
            Visible         =   0   'False
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   1032
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   " Õ—Ìﬂ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc detect_no 
            Height          =   585
            Left            =   2640
            Top             =   0
            Visible         =   0   'False
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   1032
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   " Õ—Ìﬂ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label LblShortcutKeys 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ÃœÌœ F12 Or Enter ,  ⁄œÌ· F11 , Õ›Ÿ F10 ,  —«Ã⁄ F9 ,Õ–› F8 ,»ÕÀ F7 "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   510
            Width           =   5445
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   6600
            Picture         =   "FrmExpenses3.frx":1EE8
            Stretch         =   -1  'True
            Top             =   120
            Width           =   525
         End
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   7260
         TabIndex        =   37
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   131072001
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmExpenses3.frx":5B50
         Height          =   315
         Left            =   2160
         TabIndex        =   38
         Top             =   840
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
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
      Begin MSDataListLib.DataCombo DcCostCenter 
         Bindings        =   "FrmExpenses3.frx":5B65
         Height          =   315
         Left            =   2160
         TabIndex        =   52
         Top             =   1230
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
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
      Begin MSDataListLib.DataCombo dcproject 
         Height          =   315
         Left            =   2160
         TabIndex        =   56
         Top             =   1680
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   7
         Left            =   0
         TabIndex        =   78
         Top             =   4080
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "«·⁄—÷ «·ÃœÊ·Ï"
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   3495
         Left            =   0
         TabIndex        =   101
         Top             =   4320
         Width           =   13215
         _cx             =   23310
         _cy             =   6165
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
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "«·»Ì«‰«  «·«”«”«Ì…|Õ«·… «·«⁄ „«œ"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   3120
            Left            =   45
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   45
            Width           =   13125
            _cx             =   23151
            _cy             =   5503
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
            BackColor       =   -2147483633
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   3120
               Left            =   0
               TabIndex        =   107
               TabStop         =   0   'False
               Top             =   0
               Width           =   13125
               _cx             =   23151
               _cy             =   5503
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
               Begin VB.TextBox XPTxtVal 
                  Alignment       =   1  'Right Justify
                  Height          =   360
                  Left            =   120
                  Locked          =   -1  'True
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   2670
                  Width           =   2130
               End
               Begin VB.TextBox XPTxtValView 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   120
                  Locked          =   -1  'True
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   2670
                  Width           =   2130
               End
               Begin ALLButtonS.ALLButton CmdRemove 
                  Height          =   330
                  Left            =   12135
                  TabIndex        =   110
                  Tag             =   "Delete Row"
                  Top             =   2670
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BTYPE           =   3
                  TX              =   "Õ–› ”ÿ—"
                  ENAB            =   0   'False
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   2
                  FOCUSR          =   -1  'True
                  BCOL            =   0
                  BCOLO           =   0
                  FCOL            =   255
                  FCOLO           =   255
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "FrmExpenses3.frx":5B7A
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
                  Height          =   2475
                  Left            =   270
                  TabIndex        =   111
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   12915
                  _cx             =   22781
                  _cy             =   4366
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
                  GridLines       =   3
                  GridLinesFixed  =   2
                  GridLineWidth   =   5
                  Rows            =   2
                  Cols            =   39
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmExpenses3.frx":5B96
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
                  Begin VB.PictureBox Picture1 
                     BorderStyle     =   0  'None
                     Height          =   3915
                     Left            =   2550
                     RightToLeft     =   -1  'True
                     ScaleHeight     =   3915
                     ScaleWidth      =   9405
                     TabIndex        =   116
                     Top             =   810
                     Visible         =   0   'False
                     Width           =   9405
                     Begin VB.CommandButton Command3 
                        Caption         =   "Call des"
                        Height          =   255
                        Left            =   6240
                        RightToLeft     =   -1  'True
                        TabIndex        =   120
                        Top             =   3600
                        Width           =   1095
                     End
                     Begin VB.CommandButton Command4 
                        Caption         =   "Add des"
                        Height          =   255
                        Left            =   7440
                        RightToLeft     =   -1  'True
                        TabIndex        =   119
                        Top             =   3600
                        Width           =   1350
                     End
                     Begin VB.TextBox txtcodesub 
                        Alignment       =   1  'Right Justify
                        Height          =   285
                        Left            =   5400
                        RightToLeft     =   -1  'True
                        TabIndex        =   118
                        Top             =   3600
                        Width           =   855
                     End
                     Begin VB.TextBox TxtDese 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000018&
                        BorderStyle     =   0  'None
                        Height          =   1485
                        Left            =   120
                        MultiLine       =   -1  'True
                        RightToLeft     =   -1  'True
                        ScrollBars      =   3  'Both
                        TabIndex        =   117
                        Top             =   2040
                        Width           =   8955
                     End
                     Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                        Height          =   3900
                        Left            =   120
                        TabIndex        =   121
                        TabStop         =   0   'False
                        Top             =   0
                        Width           =   10905
                        _cx             =   19235
                        _cy             =   6879
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial (Arabic)"
                           Size            =   20.25
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   -1  'True
                        Appearance      =   6
                        MousePointer    =   0
                        Version         =   801
                        BackColor       =   16777215
                        ForeColor       =   4210688
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
                        CaptionPos      =   7
                        WordWrap        =   -1  'True
                        MaxChildSize    =   0
                        MinChildSize    =   0
                        TagWidth        =   0
                        TagPosition     =   0
                        Style           =   0
                        TagSplit        =   2
                        PicturePos      =   7
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
                        Begin VB.TextBox Text3 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H80000018&
                           BorderStyle     =   0  'None
                           Height          =   1605
                           Left            =   -2400
                           MultiLine       =   -1  'True
                           RightToLeft     =   -1  'True
                           ScrollBars      =   3  'Both
                           TabIndex        =   122
                           Top             =   720
                           Visible         =   0   'False
                           Width           =   8955
                        End
                        Begin VB.Label Label2 
                           Alignment       =   1  'Right Justify
                           BackColor       =   &H8000000C&
                           Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                           ForeColor       =   &H0000C8FF&
                           Height          =   315
                           Left            =   6840
                           RightToLeft     =   -1  'True
                           TabIndex        =   123
                           Top             =   0
                           Width           =   2445
                        End
                     End
                     Begin VB.Label Label4 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Code"
                        Height          =   255
                        Left            =   1680
                        RightToLeft     =   -1  'True
                        TabIndex        =   126
                        Top             =   1320
                        Width           =   735
                     End
                     Begin VB.Label Label5 
                        Alignment       =   1  'Right Justify
                        Height          =   495
                        Left            =   1560
                        RightToLeft     =   -1  'True
                        TabIndex        =   125
                        Top             =   1200
                        Width           =   975
                     End
                     Begin VB.Label Label6 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Code"
                        Height          =   495
                        Left            =   1920
                        RightToLeft     =   -1  'True
                        TabIndex        =   124
                        Top             =   3480
                        Width           =   735
                     End
                  End
                  Begin VB.Frame Frame3 
                     Caption         =   "Õœœ —ﬁ„ «·ﬁÌœ «·„—«œ ‰”Œ…"
                     Height          =   1215
                     Left            =   -120
                     RightToLeft     =   -1  'True
                     TabIndex        =   112
                     Top             =   3720
                     Visible         =   0   'False
                     Width           =   4215
                     Begin VB.TextBox Text4 
                        Alignment       =   1  'Right Justify
                        Height          =   285
                        Left            =   360
                        RightToLeft     =   -1  'True
                        TabIndex        =   114
                        Top             =   240
                        Width           =   2175
                     End
                     Begin VB.CommandButton Command5 
                        Caption         =   "‰”Œ"
                        Height          =   255
                        Left            =   360
                        RightToLeft     =   -1  'True
                        TabIndex        =   113
                        Top             =   720
                        Width           =   1215
                     End
                     Begin VB.Label Label7 
                        Alignment       =   1  'Right Justify
                        Caption         =   "—ﬁ„ «·ﬁÌœ"
                        Height          =   255
                        Left            =   2640
                        RightToLeft     =   -1  'True
                        TabIndex        =   115
                        Top             =   240
                        Width           =   1335
                     End
                  End
                  Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   127
                     ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   2955
                     _cx             =   1973752924
                     _cy             =   1973748268
                     Alignment       =   0
                     Appearance      =   3
                     AutoSearch      =   0   'False
                     BackColor       =   -2147483624
                     BackgroundColor =   -2147483633
                     BorderColor     =   0
                     BorderVisible   =   -1  'True
                     Caption         =   "SmartCombo1"
                     CaptionAlignment=   4
                     CaptionBackColor=   -2147483633
                     BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     CaptionForeColor=   -2147483630
                     CaptionHeight   =   15
                     CaptionOnTop    =   0   'False
                     CaptionMultiLine=   0
                     Checkbox3D      =   0   'False
                     CheckboxAlignment=   5
                     CheckboxBackColor=   16777215
                     CheckboxSize    =   13
                     CheckboxValue   =   0
                     BrowsePictureAlignment=   5
                     BrowsePictureStretchH=   0
                     BrowsePictureStretchV=   0
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
                     ForeColor       =   0
                     Gap             =   0
                     HideSelection   =   -1  'True
                     Locked          =   0   'False
                     MaxLength       =   0
                     MultiLine       =   0
                     OnFocus         =   3
                     PasswordChar    =   ""
                     Picture         =   "FrmExpenses3.frx":618A
                     PictureAlignment=   5
                     PictureBackColor=   -2147483624
                     PictureStretchH =   0
                     PictureStretchV =   0
                     Redraw          =   -1  'True
                     ScrollBar       =   0
                     Style           =   0
                     Text            =   ""
                     UnderLine       =   0   'False
                     Enabled0        =   -1  'True
                     Position0       =   0
                     Tip0            =   "Caption"
                     Visible0        =   0   'False
                     Width0          =   90
                     Enabled1        =   -1  'True
                     Position1       =   1
                     Tip1            =   ""
                     Visible1        =   -1  'True
                     Width1          =   32
                     Enabled2        =   -1  'True
                     Position2       =   2
                     Tip2            =   "Check Box (Space, Ctrl + Space)"
                     Visible2        =   0   'False
                     Width2          =   16
                     Enabled3        =   -1  'True
                     Position3       =   3
                     Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
                     Visible3        =   -1  'True
                     Width3          =   145
                     Enabled4        =   -1  'True
                     Position4       =   4
                     Tip4            =   "Left Spinner (Alt + Left)"
                     Visible4        =   0   'False
                     Width4          =   16
                     Enabled5        =   -1  'True
                     Position5       =   5
                     Tip5            =   "Right Spinner (Alt + Right)"
                     Visible5        =   0   'False
                     Width5          =   16
                     Enabled6        =   -1  'True
                     Position6       =   6
                     Tip6            =   "Up Spinner (Ctrl + Up)"
                     Visible6        =   0   'False
                     Width6          =   16
                     Enabled7        =   -1  'True
                     Position7       =   7
                     Tip7            =   "Down Spinner (Ctrl + Down)"
                     Visible7        =   0   'False
                     Width7          =   16
                     Enabled8        =   -1  'True
                     Position8       =   8
                     Tip8            =   "Browse (Alt + Enter)"
                     Visible8        =   0   'False
                     Width8          =   16
                     Enabled9        =   -1  'True
                     Position9       =   9
                     Tip9            =   " (Alt + Down, F4)"
                     Visible9        =   -1  'True
                     Width9          =   16
                     Enabled10       =   -1  'True
                     Position10      =   10
                     Tip10           =   "Right Arrow (Alt + >)"
                     Visible10       =   0   'False
                     Width10         =   16
                  End
               End
               Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
                  Height          =   2475
                  Left            =   120
                  TabIndex        =   128
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   12915
                  _cx             =   22781
                  _cy             =   4366
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
                  GridLines       =   3
                  GridLinesFixed  =   2
                  GridLineWidth   =   5
                  Rows            =   2
                  Cols            =   33
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmExpenses3.frx":6724
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
                  Begin VB.PictureBox PicDes 
                     BorderStyle     =   0  'None
                     Height          =   1635
                     Left            =   240
                     RightToLeft     =   -1  'True
                     ScaleHeight     =   1635
                     ScaleWidth      =   2925
                     TabIndex        =   129
                     Top             =   480
                     Visible         =   0   'False
                     Width           =   2925
                     Begin VB.TextBox TxtDes 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H80000018&
                        BorderStyle     =   0  'None
                        Height          =   1125
                        Left            =   30
                        MultiLine       =   -1  'True
                        RightToLeft     =   -1  'True
                        ScrollBars      =   3  'Both
                        TabIndex        =   130
                        Top             =   360
                        Visible         =   0   'False
                        Width           =   2115
                     End
                     Begin VB.Label LblDes 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H8000000C&
                        Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                        ForeColor       =   &H0000C8FF&
                        Height          =   315
                        Left            =   0
                        RightToLeft     =   -1  'True
                        TabIndex        =   131
                        Top             =   0
                        Width           =   2445
                     End
                  End
               End
               Begin VB.Label LblValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Height          =   375
                  Left            =   3390
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   2610
                  Width           =   6015
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   240
                  Index           =   2
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   2670
                  Width           =   780
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   3120
            Left            =   13860
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   45
            Width           =   13125
            _cx             =   23151
            _cy             =   5503
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   2055
               Left            =   0
               TabIndex        =   104
               Tag             =   "1"
               Top             =   0
               Width           =   13095
               _cx             =   23098
               _cy             =   3625
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmExpenses3.frx":6C17
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
            Begin VB.Label Label1100 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   210
               Left            =   13335
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   2970
               Width           =   4110
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   210
               Left            =   7515
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   2310
               Width           =   4080
            End
         End
      End
      Begin ImpulseAniLabel.ISAniLabel LblLink 
         Height          =   315
         Left            =   0
         TabIndex        =   137
         Top             =   2160
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmExpenses3.frx":6D5A
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin MSDataListLib.DataCombo DcCurrency 
         Height          =   315
         Left            =   1050
         TabIndex        =   140
         Top             =   2850
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«œŒ«· «·‰”»… «·ÌœÊÌ…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   148
         Left            =   1440
         TabIndex        =   143
         Top             =   4080
         Width           =   1800
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄„·…"
         Height          =   195
         Index           =   30
         Left            =   1965
         RightToLeft     =   -1  'True
         TabIndex        =   141
         Top             =   2925
         Width           =   690
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—’Ìœ «·Õ«·Ï:"
         Height          =   315
         Index           =   29
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   138
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «· VAT ··Ã„«—ﬂ"
         Height          =   255
         Index           =   28
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   135
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·‘—Õ «·⁄«„"
         Height          =   255
         Index           =   20
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰«¡ ⁄·Ï"
         Height          =   255
         Index           =   26
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ ›« Ê—… «·„Ê—œ"
         Height          =   255
         Index           =   0
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‘—Ê⁄"
         Height          =   255
         Index           =   14
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
         Height          =   195
         Index           =   15
         Left            =   11760
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1680
         Width           =   1245
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
         Height          =   255
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰«¡ ⁄·Ï ÿ·» —ﬁ„"
         Height          =   255
         Left            =   8460
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·›« Ê—…"
         Height          =   285
         Index           =   23
         Left            =   11760
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "«·›« Ê—… «·„«·Ì… ÊÂÌ  Œ’ ﬂ· «·„‘ —Ì«  «·‰ﬁœÌ… «Ê «·«Ã·… Ê«· Ï ·Ì” ·Â« «’‰«› „ÕœœÂ Ê·« Ì‰ Ã ⁄‰Â«  √ÀÌ— „Œ“‰Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1140
         Index           =   25
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1455
         Left            =   120
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„·«ÕŸ… Â«„…:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   24
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   855
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·›« Ê—…"
         Height          =   285
         Index           =   4
         Left            =   11760
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   870
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   3735
      Left            =   13560
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Width           =   10815
      Begin VB.Frame FraNote 
         BackColor       =   &H00E2E9E9&
         Height          =   2445
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1200
         Width           =   4635
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   -240
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1590
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1590
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TXT_A_NoteID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Text            =   "Text8"
         Top             =   3150
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„’—Ê›« "
         Height          =   285
         Index           =   3
         Left            =   10800
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1110
         Width           =   1515
      End
      Begin VB.Image ImgNote 
         Height          =   240
         Left            =   -240
         Picture         =   "FrmExpenses3.frx":6EBC
         Top             =   750
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   405
         Index           =   21
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   2190
         Width           =   1155
      End
   End
   Begin VB.OptionButton OptSort 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   195
      Index           =   1
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   240
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkLastAccount 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   195
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   195
      Index           =   0
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Txt_Numorder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1035
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   9420
      Width           =   6465
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   8
         Top             =   270
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboCreditSide 
         Height          =   315
         Left            =   90
         TabIndex        =   10
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   12
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   570
         Width           =   1485
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·› —… :"
         Height          =   315
         Index           =   13
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   315
         Index           =   11
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› œ«∆‰"
         Height          =   285
         Index           =   10
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› „œÌ‰"
         Height          =   285
         Index           =   9
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4320
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   645
      Left            =   14640
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   4755
   End
   Begin MSDataListLib.DataCombo XPCboExpensesType 
      Height          =   315
      Left            =   14040
      TabIndex        =   1
      Top             =   2760
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   9
      Left            =   5640
      TabIndex        =   15
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·‘Ìﬂ"
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   27
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   6720
      Width           =   7155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·«„—"
      Height          =   285
      Index           =   5
      Left            =   13080
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2520
      Width           =   1515
   End
End
Attribute VB_Name = "FrmExpenses3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim numbering_type As Integer
Dim departement_name  As String
Dim Balance As String
Dim balanceString As String
Dim branch_no  As String
Dim RsNotes As ADODB.Recordset
Dim BolEditOnMainAccounts As Boolean
Dim ACCStrComboList As String


Private Sub DcCurrency_Change()
    If val(Me.DcCurrency.BoundText) <> MainCurrency() Then
'        VSFlexGrid1.TextMatrix(0, VSFlexGrid1.ColIndex("value")) = "«·ﬁÌ„… »«·⁄„·… «·«Ã‰»Ì…"
'        VSFlexGrid1.TextMatrix(0, VSFlexGrid1.ColIndex("valuee")) = "«·ﬁÌ„…"
'
'        Fg_Journal.TextMatrix(0, Fg_Journal.ColIndex("value")) = "«·ﬁÌ„… »«·⁄„·… «·«Ã‰»Ì…"
'        Fg_Journal.TextMatrix(0, Fg_Journal.ColIndex("valuee")) = "«·ﬁÌ„…"
'
        
        VSFlexGrid1.ColHidden(VSFlexGrid1.ColIndex("valuee")) = False
        Fg_Journal.ColHidden(Fg_Journal.ColIndex("valuee")) = False
    Else
        VSFlexGrid1.ColHidden(VSFlexGrid1.ColIndex("valuee")) = True
        Fg_Journal.ColHidden(Fg_Journal.ColIndex("valuee")) = True
        
'        VSFlexGrid1.TextMatrix(0, VSFlexGrid1.ColIndex("valuee")) = "«·ﬁÌ„… »«·⁄„·… «·«Ã‰»Ì…"
'        VSFlexGrid1.TextMatrix(0, VSFlexGrid1.ColIndex("value")) = "«·ﬁÌ„…"
'
'        Fg_Journal.TextMatrix(0, Fg_Journal.ColIndex("valuee")) = "«·ﬁÌ„… »«·⁄„·… «·«Ã‰»Ì…"
'        Fg_Journal.TextMatrix(0, Fg_Journal.ColIndex("value")) = "«·ﬁÌ„…"
                
    End If

    If Me.TxtModFlg.Text = "" Or Me.TxtModFlg.Text = "R" Then Exit Sub
    If Me.DcCurrency.BoundText <> "" Then
        txt_Currency_rate.Text = get_currency_rate(val(Me.DcCurrency.BoundText))
    Else
        txt_Currency_rate.Text = 1
        DcCurrency.BoundText = MainCurrency()
    End If
     ReLineGrid
    ' ChAddToTotal_Click
     
     
End Sub

Private Sub DcCurrency_Click(Area As Integer)
    DcCurrency_Change
End Sub

Private Sub DcboCreditSide_Change()
    WriteCustomerBalPublic Me.DcboCreditSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If Index = 18 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(18).ToolTipText = "ﬁÌ„… „»·€ «·„ﬁ»Ê÷« :" & lbl(18).Caption
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(18).ToolTipText = "Notes Recivable Value:" & lbl(18).Caption
        End If
    End If

End Sub

Private Sub LblLink_Click()
 
    Dim FirstPeriod As Date
    getFirstPeriodDateInthisYear FirstPeriod
    ShowReport DcboCreditSide.BoundText, DcboCreditSide.Text, FirstPeriod, Date

End Sub

Private Sub LblLink_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
 
    If SystemOptions.UserInterface = ArabicInterface Then
        LblLink.ToolTipText = "—’Ìœ «·ÿ—› «·œ«∆‰:" & WriteNo(Balance, 0, True)
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        LblLink.ToolTipText = "Credit Balance:" & WriteNo(Balance, 0, True)
    End If
 
End Sub
Function checkbilllinkedtobill(NoteID As Double) As Boolean
       checkbilllinkedtobill = False: Exit Function
    Dim sql As String
    Dim rs As New ADODB.Recordset
       
    sql = " SELECT     dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1, dbo.Transactions.NoteSerial1 AS Expr1, "
    sql = sql & "                         dbo.Transactions.Transaction_Date"
    sql = sql & "   FROM         dbo.Notes INNER JOIN"
    sql = sql & "                         dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    sql = sql & "                         dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code INNER JOIN"
    sql = sql & "     dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID RIGHT OUTER JOIN"
    sql = sql & "                         dbo.Transactions ON dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 = dbo.Transactions.Transaction_ID"
    sql = sql & "   WHERE     (dbo.Notes.NoteID = " & NoteID & ")"
  
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Dim Msg As String
    If rs.RecordCount > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–… «·›« Ê—… „— »ÿ… »›« Ê—… ‘—«¡ »—ﬁ„ " & IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
            Msg = Msg & "» «—ÌŒ" & IIf(IsNull(rs("Transaction_Date").value), "", rs("Transaction_Date").value)
            Msg = Msg & "Ê·«Ì„ﬂ‰ «· ⁄œÌ· «Ê «·Õ–› «·« »⁄œ «·€«¡ «·—Ìÿ «Ê·« " & ""
        Else
            Msg = "This bill linked to the purchase invoice number" & IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
            Msg = Msg & "Date" & IIf(IsNull(rs("Transaction_Date").value), "", rs("Transaction_Date").value)
            Msg = Msg & "Can not be edited or deletion " & ""
        End If
        MsgBox Msg, vbCritical
        checkbilllinkedtobill = True
    Else
        checkbilllinkedtobill = False
    End If

    rs.Close
    
End Function
Function saveChequeBoxContents1(NoteID As Double)
    
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords

    If SystemOptions.banks_Accounts3 = False Then Exit Function
 
    'rs.Open "TblChecqueBoxContent1", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     * from dbo.TblChecqueBoxContent1 Where (1 = -1)"
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
    If CboPayMentType.ListIndex = 1 Then
        rs.AddNew
        rs("noteid").value = NoteID
     
        rs("RecordDate").value = XPDtbTrans.value
        rs("DueDate").value = DtpChequeDueDate.value
        rs("BankID").value = val(DcboBankName.BoundText)
        rs("BankName").value = DcboBankName.Text
        
        rs("ChequeNo").value = TxtChequeNumber.Text
        rs("ChequeValue").value = val(XPTxtVal.Text)
    
        rs("Remarks").value = Me.DcboDebitSide.Text
        rs("Payed").value = 0
       
        rs("DepitAccount").value = (DcboDebitSide.BoundText)
        rs("notes_all").value = NoteID
      
        rs.update
    End If

    rs.Close
End Function

Private Sub Accredit_Click()
  
    Dim BeginTrans As Boolean
    If val(XPTxtID.Text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
        Else
            MsgBox "Save Doc First", vbCritical
        End If
        Exit Sub
    End If
 
    SendTopost Me.Name, "notes_all", "NoteID", 0, val(Dcbranch.BoundText), val(XPTxtID.Text), TxtSerial1.Text
    rs.Resync
        
        If SystemOptions.UserInterface = ArabicInterface Then
            Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
        Else
            Accredit.Caption = "Sent To approval "
        End If
        fillapprovData
End Sub
Function fillapprovData()
    
    Dim Num As Integer
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
    StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
    StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
    StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
    StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
    StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
    StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
    StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsDetails.RecordCount > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
        Else
            Accredit.Caption = "Sent To approval "
        End If
        Accredit.Enabled = False
    Else
        Accredit.Enabled = True
        
        If SystemOptions.UserInterface = ArabicInterface Then
            Accredit.Caption = " «·«—”«· ··«⁄ „«œ"
        Else
            Accredit.Caption = "Sent To approval "
        End If
    End If
 
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
        
            GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
            
            If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
                GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
            Else
                GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
            End If
    
            GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
            
            If SystemOptions.UserInterface = ArabicInterface Then
                GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
            Else
                GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
            End If
            
            If SystemOptions.UserInterface = ArabicInterface Then
                GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
                GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
            GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
            RsDetails.MoveNext

            If Num = RsDetails.RecordCount Then
                If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                    Else
                        Label11.Caption = "Approved"
                    End If
                    Label11.backcolor = &H80FF80
                Else
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                    Else
                        Label11.Caption = "Currently required Approve"
                    End If
                    Label11.backcolor = &HFFFFC0
                End If
            End If
        Next Num
    Else
        GRID2.Rows = 1
    End If
    RsDetails.Close
End Function
Private Sub ALLButton1_Click()
    On Error GoTo ErrTrap

    If DcCostCenter.BoundText <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «· Ê“Ì⁄ ⁄·Ï „—«ﬂ“ «· ﬂ·›… ·«‰ﬂ «Œ —   Ê“Ì⁄ ⁄«„ ⁄·Ï „—ﬂ“  ﬂ·›… „Õœœ", vbCritical
        Else
        MsgBox "·It can not be the cost of distribution centers because you chose in distribution", vbCritical
        End If
        Exit Sub
    End If

    Dim opr_id As Double

    If Not IsNumeric(Text1.Text) Then Exit Sub
    'If Me.TxtModFlg.text = "N" Then
    opr_id = val(Me.Text1.Text)
    'Else
    'opr_id = TxtDEV_NO.text
    'End If
If CboPaymentType1.ListIndex = 0 Then
    If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) = "" Then
        If Not val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE"))) = 0 And Not val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))) = 0 Then

            marakes_taklefa_tawze3.show
            
            marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
            marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
            marakes_taklefa_tawze3.lineno = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        
        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… ", vbCritical
            Else
                MsgBox "Enter Value First ", vbCritical
            End If

            Exit Sub
        End If
            
    End If
Else
    If Not VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("AccountCode")) = "" Then
        If Not val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("VALUE"))) = 0 And Not val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("LineNo1"))) = 0 Then

            marakes_taklefa_tawze3.show
            
            marakes_taklefa_tawze3.value.Caption = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("VALUE")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("AccountCode"))
            marakes_taklefa_tawze3.account_name = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("AccountName"))
            marakes_taklefa_tawze3.lineno = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("LineNo1"))
        
        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… ", vbCritical
            Else
                MsgBox "Enter Value First ", vbCritical
            End If

            Exit Sub
        End If
            
    End If
End If
    marakes_taklefa_tawze3.opr_type = "›« Ê—… „«·Ì…"
    marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
    marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
    marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
     If CboPaymentType1.ListIndex = 0 Then
     marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
    Else
     marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & VSFlexGrid1.TextMatrix(Fg_Journal.Row, VSFlexGrid1.ColIndex("AccountCode")) & "' and  line_no=" & VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("LineNo1"))
    End If
    
    marakes_taklefa_tawze3.Adodc3.Refresh
    '    Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"

    Exit Sub
ErrTrap:
End Sub

Private Sub CBoBasedON_Click()
    CBoBasedON_Change
End Sub

Private Sub CboDes_AfterAutoCloseUp()
    PutData
    CboDes.Visible = False
End Sub

Private Sub CboPayMentType_Change()

    If Me.TxtModFlg.Text = "E" Then
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        Me.DcboBox.Text = ""
        DCVendor.Text = ""
        
    End If

    If Me.CboPayMentType.ListIndex = 0 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DCVendor.Enabled = False
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        DCVendor.Text = ""
    ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 3 Then
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.Text = ""
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Me.DCVendor.Enabled = False
     DCVendor.Text = ""
    ElseIf Me.CboPayMentType.ListIndex = 2 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = False
    
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DcboBox.Enabled = False
        Me.DCVendor.Enabled = True
            Me.DtpChequeDueDate.Enabled = True

    Else

        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = False
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DCVendor.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Function setfoxy()
    Text1.Text = CStr(new_id("foxy", "id", "", True))

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id").value = Text1.Text
 
    rs.update
    
End Function

Private Sub CboPaymentType1_Change()

    If Me.CboPaymentType1.ListIndex = 0 Then
        Fg_Journal.Visible = True
        Fg_Journal.Clear flexClearScrollable, flexClearEverything
        Fg_Journal.Rows = 3
          
        VSFlexGrid1.Visible = False

    ElseIf Me.CboPaymentType1.ListIndex = 1 Then

        Fg_Journal.Visible = False
        VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
        VSFlexGrid1.Rows = 3
        VSFlexGrid1.Visible = True
    End If

End Sub

Private Sub CboPaymentType1_Click()
    CboPaymentType1_Change
End Sub

Private Sub Cmd_Click(Index As Integer)
    'On Error GoTo ErrTrap

    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            DcCostCenter.Text = ""
            DCproject.Text = ""
            GRID2.Clear flexClearScrollable, flexClearEverything
            GRID2.Rows = 1
            Accredit.Caption = ""
            ' XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
        
            Me.DCboUserName.BoundText = user_id
            'XPDtbTrans.SetFocus
            Fg_Journal.Visible = False
            VSFlexGrid1.Visible = False

            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 3
            Fg_Journal.Enabled = True
          
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 3
            VSFlexGrid1.Enabled = True
          
            DtpChequeDueDate.value = Date
            setfoxy
            CBoBasedON.ListIndex = 0
            CboPaymentType1.ListIndex = 0
            Me.Dcbranch.BoundText = Current_branch
             
            DcCurrency.BoundText = MainCurrency()
        Case 1
            If ScreenAproved(val(XPTxtID.Text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If
            
            If ChekClodePeriod(XPDtbTrans.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
        
        If checkbilllinkedtobill(Me.TXT_A_NoteID.Text) = True Then
            Exit Sub
        End If
        
        Dim Msg As String
            If SystemOptions.banks_Accounts3 = True Then
                If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–Â «·⁄„·Ì…"
                        Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
                    Else
                        Msg = " Can Not Edit this Process"
                        Msg = Msg & CHR(13) & " There is the Process Payment checks "
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If
             
            If CheAssetPayd(val(Me.XPTxtID)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–Â «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ≈÷«›… ··«’Ê·   "
                Else
                    Msg = " Can Not Edit this Process"
                    Msg = Msg & CHR(13) & " There is the Process of adding Assest "
                End If
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If
            
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
        '    Me.DCboUserName.BoundText = user_id
            Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
            VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
            VSFlexGrid1.Enabled = True
        
            CuurentLogdata

        Case 2
            If ChekClodePeriod(XPDtbTrans.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
              
                    
    If DcCurrency.BoundText = "" Then
    
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "«Œ — «·⁄„·… «Ê·« "
        Else
            Msg = "Select Currency First"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcCurrency.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    
    End If

         
    
            If CBoBasedON.ListIndex > 0 And Trim(Txt_order_no.Text) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify NO For"
                Else
                    Msg = "Õœœ —ﬁ„ "
                End If

                Msg = Msg & "  " & CBoBasedON.Text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Txt_order_no.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Dcbranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.Dcbranch.BoundText
            DcboBox_Change
            DcboBankName_Change
            DCVendor_Click (0)
            
            Dim Account_Code_dynamic82 As String
         
            If val(TxtVATCustoms.Text) > 0 Then
                Account_Code_dynamic82 = get_account_code_branch(148, my_branch)
                If Account_Code_dynamic82 = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·Ã„«—ﬂ", vbCritical
                    Else
                        MsgBox "Please Select Customs Account", vbCritical
                    End If
                    GoTo ErrTrap
                End If
            End If
            SaveData
        Case 3
            Undo
        Case 4
            If ScreenAproved(val(XPTxtID.Text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «·Õ–›.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not delete.This process associated with approvals"
                End If
                Exit Sub
            End If
             
            If ChekClodePeriod(XPDtbTrans.value) = True Then
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
        
            If checkbilllinkedtobill(Me.TXT_A_NoteID.Text) = True Then
                Exit Sub
            End If
            Del_Trans
        Case 5
            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If


             Load FrmNotesSearch
             FrmNotesSearch.SearchType = 80
             FrmNotesSearch.show vbModal

        Case 6
            Unload Me
        Case 7
            ViewDataList
        Case 8
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            print_report TxtSerial.Text, DCVendor.Text
        Case 9
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            print_Cheque TxtChequeNumber.Text, get_Cheque_report_no(val(DcboBankName.BoundText)), TxtSerial.Text
        Case 10
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            ShowGL_cc TxtSerial.Text, , 80, , , TxtSerial1
    End Select
    Exit Sub
ErrTrap:
End Sub

Function print_Cheque(Optional ChqueNum As String = "", Optional report_no As String = "", Optional serial As String)
    hide_logo = True
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From Expanses_Order  where ChqueNum='" & ChqueNum & "' and noteserial='" & TxtSerial & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
    Else
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
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
    'MsgBox ToHijriDate(Date)

    xReport.ParameterFields(5).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 1, 2)
    xReport.ParameterFields(6).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 4, 2)
    xReport.ParameterFields(7).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 9, 2)

    xReport.ParameterFields(8).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 1, 2)
    xReport.ParameterFields(9).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 4, 2)
    xReport.ParameterFields(10).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 9, 2)
    xReport.ParameterFields(11).AddCurrentValue CStr(txtTo.Text)
    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtVal.Text)
    xReport.ParameterFields(13).AddCurrentValue CStr(Me.XPMTxtRemarks.Text)
    xReport.ParameterFields(14).AddCurrentValue CStr(LblValue.Caption)
 
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Function print_report(Optional NoteSerial As String, Optional VendorName As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    'MySQL = "Select * From Expanses_Order  where noteserial='" & NoteSerial & "'"
    If CboPaymentType1.ListIndex = 0 Then
        MySQL = "SELECT   distinct  dbo.notes_all.NoteID, DOUBLE_ENTREY_VOUCHERS.VAT ,dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial, dbo.notes_all.Note_Value, dbo.notes_all.BankID, "
        MySQL = MySQL & "   dbo.notes_all.ChqueNum, dbo.notes_all.DueDate, dbo.notes_all.UserID, dbo.notes_all.Remark, dbo.notes_all.ExpensesID, dbo.notes_all.BoxID,"
        MySQL = MySQL & "  dbo.TblUsers.UserName, dbo.TblBoxesData.BoxName, dbo.BanksData.BankName, dbo.notes_all.too, dbo.Notes.Note_Value AS [Sub-value],"
        MySQL = MySQL & "  dbo.Notes.note_value_by_characters AS sub_note_value_by_char, dbo.Notes.Remark AS sub_remark, dbo.ExpensesType.Name AS Sub_expenses_name,"
        MySQL = MySQL & "  dbo.Notes.NoteType AS DD, dbo.notes_all.note_value_by_characters, dbo.notes_all.general_des, dbo.notes_all.NoteSerial1, dbo.notes.ExpensesRemark"
        MySQL = MySQL & "  ,dbo.ExpensesType.Namee FROM         dbo.ExpensesType RIGHT OUTER JOIN"
        MySQL = MySQL & "  dbo.Notes ON dbo.ExpensesType.ID = dbo.Notes.ExpensesID LEFT OUTER JOIN"
        MySQL = MySQL & "  dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN "
        MySQL = MySQL & "  dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID LEFT OUTER JOIN "
        MySQL = MySQL & "  dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID LEFT OUTER JOIN "
        MySQL = MySQL & "  dbo.TblUsers ON dbo.Notes.UserID = dbo.TblUsers.UserID"
        MySQL = MySQL & "  RIGHT Outer join  dbo.DOUBLE_ENTREY_VOUCHERS  ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
        MySQL = MySQL & "  AND IsNull(FlgVat,0) =0"
        MySQL = MySQL & "  WHERE     (dbo.Notes.NoteType = 80) AND (NOT (dbo.ExpensesType.Name IS NULL))  and  dbo.Notes.noteserial='" & NoteSerial & "'"
        MySQL = MySQL & "    AND (dbo.Notes.NoteSerial1 = " & TxtSerial1 & ") AND IsNull(FlgVat,0) =0      "
    Else
        'MySQL = "SELECT     dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.BankID, dbo.Notes.ChqueNum, dbo.Notes.DueDate, "
        'MySQL = MySQL & "  dbo.Notes.CusID, dbo.Notes.BoxID, dbo.Notes.Note_Value, dbo.Notes.note_value_by_characters, dbo.Notes.Remark AS sub_remark, dbo.ACCOUNTS.Account_Name,"
        'MySQL = MySQL & "  dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS [sub-value],"
        'MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS expenses_remark, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
        'MySQL = MySQL & "  FROM         dbo.Notes INNER JOIN"
        'MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
        'MySQL = MySQL & "  dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
        'MySQL = MySQL & "  WHERE     (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND (dbo.Notes.NoteSerial = " & NoteSerial & ")"
        
        MySQL = " SELECT  distinct Notes.NoteDate, Notes.NoteType, Notes.NoteSerial, Notes.NoteSerial1, Notes.BankID, Notes.ChqueNum, Notes.DueDate, Notes.CusID, Notes.BoxID, Notes.Note_Value, Notes.note_value_by_characters,"
        MySQL = MySQL & " Notes.Remark AS sub_remark,DOUBLE_ENTREY_VOUCHERS.VAT, ACCOUNTS.Account_Name ,ACCOUNTS.Account_Name  Sub_expenses_name, ACCOUNTS.Account_NameEng, CAST(DOUBLE_ENTREY_VOUCHERS.Value AS NVARCHAR(10)) as Value2, "
       
        
        MySQL = MySQL & " DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description , DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, notes_all.too"
        MySQL = MySQL & " FROM Notes INNER JOIN"
        MySQL = MySQL & " DOUBLE_ENTREY_VOUCHERS ON Notes.NoteID = DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
        MySQL = MySQL & " ACCOUNTS ON DOUBLE_ENTREY_VOUCHERS.Account_Code = ACCOUNTS.Account_Code LEFT OUTER JOIN"
        MySQL = MySQL & " notes_all ON Notes.notes_all = notes_all.NoteID"

        MySQL = MySQL & "  WHERE     (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND (dbo.Notes.NoteSerial = " & NoteSerial & ") AND IsNull(FlgVat,0) = 0"
        
       
        
       MySQL = MySQL & "    AND (dbo.Notes.NoteSerial1 = " & TxtSerial1 & ")"
        
    End If

    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If

    If CboPaymentType1.ListIndex = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\New Reports\" & "FinancialInvoice.rpt"
        Else
            StrFileName = App.path & "\Reports\New Reports\" & "FinancialInvoice.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\New Reports\" & "FinancialInvoiceAccounts.rpt"
        Else
            StrFileName = App.path & "\Reports\New Reports\" & "FinancialInvoiceAccounts.rpt"
        End If

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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
        
        xReport.ParameterFields(4).AddCurrentValue "'" & LblValue & "'"
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        xReport.ParameterFields(3).AddCurrentValue "'" & LblValue & "'"
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(6).AddCurrentValue VendorName

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

Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtSerial1, "0612201403"

End Sub

Private Sub CmdHelp_Click()
'    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
'    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd


   TxtModFlg.Text = "N"
            Me.XPTxtID.Text = ""
 
            Me.DCboUserName.BoundText = user_id
              'Me.DcBranch.BoundText = Current_branch
     TxtSerial.Text = ""
     TxtSerial1.Text = ""
 
 
End Sub

Private Sub CmdRemove_Click()
    Dim X As Integer
 If CboPaymentType1.ListIndex = 0 Then
If val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("FlgVat"))) = 1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ Õ–› ”ÿ— «·›«  .Ì—ÃÏ  ’›Ì— ‰”»… «·›« "
Else
MsgBox "Can not delete VAT  "
End If
Exit Sub
End If
Else

If val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("FlgVat"))) = 1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ Õ–› ”ÿ— «·›«  .Ì—ÃÏ  ’›Ì— ‰”»… «·›« "
Else
MsgBox "Can not delete VAT  "
End If
Exit Sub
End If
End If
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    Dim sql As String

    sql = "Delete  marakes_taklefa_temp where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
    Cn.Execute sql, , adExecuteNoRecords
    
    If CboPaymentType1.ListIndex = 0 Then
        If Fg_Journal.Rows > 1 Then
            If Fg_Journal.Rows = 2 Then
                Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.Fg_Journal.Rows > 1 Then
                    If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                                        
                        With Me.Fg_Journal

                       '     If Me.TxtModFlg <> "E" Then Exit Sub
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                         
                            LogTextA = "  Õ–› «·„’—Ê›   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " »ﬁÌ„… " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                            LogTexte = "  Delete  Expensen   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " With Value " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                                                         
                            AddToLogFile CInt(user_id), 80, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", val(Me.TxtSerial), val(TxtSerial1)
                        End With
                                                        
                        Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                    End If
                End If
            End If
        End If
            
        With Fg_Journal
            Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        End With

    ElseIf CboPaymentType1.ListIndex = 1 Then

        If VSFlexGrid1.Rows > 1 Then
            If VSFlexGrid1.Rows = 2 Then
                Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.VSFlexGrid1.Rows > 1 Then
                    If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                        
                        With Me.VSFlexGrid1

                         '   If Me.TxtModFlg <> "E" Then Exit Sub
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                         
                            LogTextA = "  Õ–› «·Õ”«»   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " »ﬁÌ„… " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                            LogTexte = "  Delete  Account   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " With Value " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
                                                         
                            AddToLogFile CInt(user_id), 80, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", val(Me.TxtSerial), val(TxtSerial1)
                        End With
                        
                        Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                    End If
                End If
            End If
        End If
            
        With VSFlexGrid1
            Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        End With
             
    Else
 
        Exit Sub
    End If

End Sub

Private Sub DcboBankName_Change()
    On Error Resume Next

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then

        '    Me.DcboCreditSide.BoundText = "a2a3a2"
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
        If SystemOptions.banks_Accounts3 = True Then
            Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code2")
        Else
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
    
        'Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
  
        If CboPayMentType.ListIndex = 3 Then
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If

    End If

End Sub

Private Sub DcboBox_Change()

    If DcboBox.BoundText = "" Then Exit Sub
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub

Private Sub DcboBox_Click(Area As Integer)
    DcboBox_Change
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtSerial.Text = ""
    TxtSerial1.Text = ""
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 3
    End If

End Sub

Private Sub dcproject_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 6
             FrmProjectSearch.show vbModal
           
        End If
End Sub

Private Sub DCVendor_Click(Area As Integer)

    If DCVendor.BoundText = "" Then Exit Sub
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DCVendor.BoundText))
    End If

    Text2.Text = Me.DCVendor.BoundText
End Sub

Private Sub dcVendor_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
         FrmCompanySearch.lblSearchtype.Caption = 3005
          FrmCompanySearch.show vbModal
        
        
    End If
End Sub

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim StrComboList As String
    With Fg_Journal

        Select Case .ColKey(Col)
        Case "Vatyo"
        If val(.TextMatrix(Row, .ColIndex("Vatyo"))) = 0 Then
        .TextMatrix(Row, .ColIndex("Vat")) = 0
        If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) <> 0 Then
        .TextMatrix(Row, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("PriceTotal")))
        End If
        If .Rows > Row Then
        If val(.TextMatrix(Row + 1, .ColIndex("FlgVat"))) = 1 Then
        .RemoveItem Row + 1
        End If
        End If
        End If
        Case "PriceTotal"
          AddVAT Row
                       Case "pand"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("pandid")) = StrAccountCode
                  Case "oper"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("operid")) = StrAccountCode

        Case "Supplier"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("SupplierID"), False, True)
                .TextMatrix(Row, .ColIndex("SupplierID")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("SupplierName")) = .TextMatrix(Row, .ColIndex("Supplier"))
                StrSQL = "Select * From TblCustemers Where CusID=" & val(.TextMatrix(Row, .ColIndex("SupplierID")))
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("CusVATNO")) = IIf(IsNull(rs("VATNO").value), "", rs("VATNO").value)
                Else
                .TextMatrix(Row, .ColIndex("CusVATNO")) = ""
                End If
                
             AddVAT Row
             
            Case "DeparmentName"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("DeparmentID"), False, True)
                .TextMatrix(Row, .ColIndex("DeparmentID")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("DeparmentName")) = .TextMatrix(Row, .ColIndex("DeparmentName"))
         
             
            Case "PFuLLCode"
                .TextMatrix(Row, .ColIndex("ProjectID")) = ""
                .TextMatrix(Row, .ColIndex("ProjectName")) = ""
                StrSQL = "Select expanses_account,REVENUE_account,id,Fullcode,Project_name From projects Where Fullcode='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("ProjectName")) = rs!Project_name & ""
                    .TextMatrix(Row, .ColIndex("ProjectID")) = rs!ID & ""
                Else
                    
                    .TextMatrix(Row, .ColIndex("PFuLLCode")) = ""
                    .TextMatrix(Row, .ColIndex("ProjectID")) = ""
                    .TextMatrix(Row, .ColIndex("ProjectName")) = ""
                End If
             
    Case "ProjectName"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ProjectID"), False, True)
                .TextMatrix(Row, .ColIndex("ProjectID")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ProjectName")) = .TextMatrix(Row, .ColIndex("ProjectName"))
         '       StrSQL = "Select expanses_account,REVENUE_account From projects Where Id=" & val(.TextMatrix(Row, .ColIndex("ProjectID")))
         '       Set rs = New ADODB.Recordset
         '       rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
         '
                              
                
             
         '       If rs.RecordCount > 0 Then
         '           .TextMatrix(Row, .ColIndex("AccountCode")) = Get_Account_Parent_code(IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value))
         '           .TextMatrix(Row, .ColIndex("AccountName")) = GET_ACCOUNT_name_by_Code(IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value))
         '       Else
         '           .TextMatrix(Row, .ColIndex("AccountName")) = ""
         '       End If

           '  AddVATExp Row

            Case "CarName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
     
                StrAccountCode = .ComboData
            
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CarID"), False, True)
                .TextMatrix(Row, .ColIndex("CarID")) = StrAccountCode
            
                .TextMatrix(Row, .ColIndex("des")) = "’—›  ⁄·Ï «·„⁄œÂ/«·”Ì«—…  : " & .TextMatrix(Row, .ColIndex("CarName"))
       
            Case "ExpensesID"
              
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                .TextMatrix(Row, .ColIndex("Order_No")) = Txt_order_no.Text
                AddVAT Row
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                Else
                    .TextMatrix(Row, .ColIndex("des")) = ""
                End If


   
            Case "valuee"
                 .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                 .TextMatrix(Row, .ColIndex("value")) = Round(val(.TextMatrix(Row, .ColIndex("valuee"))) * val(txt_Currency_rate), 3)
                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
              AddVAT Row
              
              

            Case "value", "opr_fullcode"
                Dim sgl As String
                Dim project_id As Integer
                project_id = get_project_id(DCproject.BoundText, "expanses_account")
                
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                End If
                .TextMatrix(Row, .ColIndex("valuee")) = Round(val(.TextMatrix(Row, .ColIndex("value"))) * val(txt_Currency_rate), 3)
               AddVAT Row
                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
                sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid

    With Me.Fg_Journal

        If Me.TxtModFlg <> "E" Then Exit Sub

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
        If Col = .ColIndex("Account_Serial") Or Col = .ColIndex("AccountName") Then
            LogTextA = "   ⁄œÌ· «·„’—Ê› «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Account To " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("Value") Then
            LogTextA = "   ⁄œÌ· «·ﬁÌ„…  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " ··„’—Ê›   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change value" & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " To Expenses " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("Des") Then
            LogTextA = "   ⁄œÌ· «·‘—Õ  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " ··„’—Ê›   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Des " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " To Expenses " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        End If

        AddToLogFile CInt(user_id), 80, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
    End With

End Sub

Function calcnets()

    If Me.CboPaymentType1.ListIndex = 0 Then

        With Fg_Journal
            Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        End With

    Else

        With Me.VSFlexGrid1
            Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        End With

    End If

End Function

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

   If val(.TextMatrix(Row, .ColIndex("FlgVat"))) <> 0 Then
   Cancel = True
Else
 Select Case .ColKey(Col)
 Case "valuee"
                   .ComboList = ""
        Case "Vat"
                 Cancel = True
        Case "Vatyo"
              If val(.TextMatrix(Row, .ColIndex("ForcedFlg"))) = 1 Then
                 Cancel = True
              Else
              .ComboList = ""
              End If
           Case "BillNo", "PFuLLCode"
                .ComboList = ""
        
         Case "CusVATNO"
                .ComboList = ""
         Case "SupplierName"
                .ComboList = ""
        Case "ProjectName"
                .ComboList = ""
         Case "PriceTotal"
                .ComboList = ""
     Case "LineNo"
                .ComboList = ""
     
     
            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
                '  Cancel = True
            
            Case "Order_No"
                .ComboList = ""
        End Select
      End If
    End With

End Sub

Private Sub Fg_Journal_DblClick()
    Exit Sub
  
    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" Then
            CboDes.Visible = False
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
            'Exit Sub
        End If

        If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
            Exit Sub
        End If

        ' same cell or neighbour? no work
        '    If r = lNoteRow And C = lNoteCol Then Exit Sub
        '    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

        ' other cell, hide current note, if any
        If lNoteRow >= 0 And lNoteCol >= 0 Then
            Fg_Journal.SetFocus
            lNoteRow = -1
            lNoteCol = -1
        End If

        ' no note to show? then bail out
        If r <= 0 Or c <= 0 Then Exit Sub
        If typename(Fg_Journal.Cell(flexcpData, r, c)) <> "String" Then
            TxtDes.Text = ""
        Else
            '
            TxtDes.Text = Fg_Journal.Cell(flexcpData, r, c)
        End If

        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
  SendKeys "{F4}"
   SendKeys "{BACKSPACE}"
  SendKeys CHR(KeyAscii)
End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    With Fg_Journal

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
'
'                    Order_no_search.show
'                     Order_no_search.RetrunType = 4
                   
                End If

            Case "AccountName"

                If KeyCode = vbKeyF3 Then
                     FrmExpensesSearch.show
                     FrmExpensesSearch.RetrunType = 2
                End If
 
        End Select

    End With

End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset
      Dim Rs3 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String
    Dim StrComboList2 As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)
         Case "Supplier"
            StrSQL = " SELECT     CusID, CusName, CusNamee"
            StrSQL = StrSQL & "            From dbo.TblCustemers"
            StrSQL = StrSQL & "       WHERE     (Type = 2)"
         Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "CusName", "CusID")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "CusNamee", "CusID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                '«ŸÂ«— «·„⁄œ« /«·”Ì«—« 
                
                
                
                 Case "pand"
             If .TextMatrix(Row, .ColIndex("projectid")) = "" Then
             MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
             Exit Sub
             End If

                StrSQL = " SELECT     des, oprid From projects_des "
                 StrSQL = StrSQL & "    Where (project_id =" & val(.TextMatrix(Row, .ColIndex("projectid"))) & ")"
                Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = .BuildComboList(Rs3, "des", "oprid")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
            
        Case "DeparmentName"
           
            If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "Select DeparmentID,DepartmentName From TblEmpDepartments Order By DepartmentName"
            Else
                StrSQL = "Select DeparmentID,DepartmentNamee From TblEmpDepartments Order By DepartmentNamee"
            End If
         Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "DepartmentNamee", "DeparmentID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                '
         Case "ProjectName"
         If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = " SELECT     ID, LTRIM(RTRIM( Project_name )) as Project_name"
         Else
         StrSQL = " SELECT     ID, LTRIM(RTRIM( Project_namee )) as Project_name"
         End If
            StrSQL = StrSQL & "            From dbo.projects"
            
     
            
            
        If SystemOptions.UserInterface = ArabicInterface Then
                 
                StrSQL = StrSQL & " where Project_name<>N'""' and not (Project_name is null)"
   Else
     
                StrSQL = StrSQL & " where Project_nameE<>N'""' and not (Project_nameE is null)"
End If


                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Project_Name", "ID")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Project_Name", "ID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
            Case "CarName"
        
                StrSQL = "  select id,BoardNO from TblCarsData"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList2 = Fg_Journal.BuildComboList(rs, "BoardNO", "id")
       
                If StrComboList2 <> "" Then
                    StrComboList2 = "|" & StrComboList2
                End If

                .ComboList = StrComboList2

            Case "AccountName"
                 
              If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts order by Account_Name"
                 Else
                     StrSQL = "select * from Expenses_accounts_eng order by Account_Nameeng"
                 End If
            
             
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

           '     If SystemOptions.UserInterface = ArabicInterface Then
           '         StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
           '     Else
           '         StrComboList = Fg_Journal.BuildComboList(rs, "Account_NameEng", "Account_Code")
           '     End If

          If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "parent_account,account_serial,*Account_Name", "Account_Code")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "parent_account,account_serial,*Account_NameEng", "Account_Code")
                End If
            
            
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                           Case "oper"
                   
If .TextMatrix(Row, .ColIndex("projectid")) = "" Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
If .TextMatrix(Row, .ColIndex("pandid")) = "" Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·»‰œ «Ê·«"
.TextMatrix(Row, .ColIndex("oper")) = ""
Exit Sub
End If
           
                If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "SELECT     dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
               Else
               StrSQL = "SELECT     dbo.TblProcessDEF.ProcessNameE, dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "    FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "      dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEF"
                End If
               StrSQL = StrSQL & "    Where (ProjectDes_ID = " & val(.TextMatrix(Row, .ColIndex("pandid"))) & ") And (project_id = " & val(.TextMatrix(Row, .ColIndex("projectid"))) & ")"
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "ProcessName", "TblProcessDEFID")
                    Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "ProcessNameE", "TblProcessDEFID")
                    End If
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        
        
        
        
    
            Case "opr_fullcode"
                Dim project_id As Integer
                project_id = get_project_id(DCproject.BoundText, "expanses_account")

                If SystemOptions.Items_or_operation = 1 Then
                    StrSQL = "  select fullcode,name from terms_operations where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,name", "fullcode")
                ElseIf SystemOptions.Items_or_operation = 0 Then
                    StrSQL = "  select fullcode,des from projects_des where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,des", "fullcode")
         
                End If

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With

End Sub


Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String

    On Error GoTo ErrTrap
    Dim StrComboList As String
    ScreenNameArabic = "›« Ê—… „«·Ì…"
    ScreenNameEnglish = "Financial Invoice"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 80

   
  If SystemOptions.AllowEditVaTManulay = True Then
txtManulaVat.Enabled = True
txtManulaVat.Visible = True
Else
txtManulaVat.Enabled = False
txtManulaVat.Text = 0
txtManulaVat.Visible = False
End If


    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
     StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
    fill_combo Me.DcCostCenter, StrSQL
     HidFat
    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("FillData").Picture
    Resize_Form Me
    AddTip
    SetDtpickerDate XPDtbTrans
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetExpensesType XPCboExpensesType
    Dcombos.GetBranches Me.Dcbranch
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.XPCboExpensesType

    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
 
    StrSQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, StrSQL
    DcCurrency.BoundText = MainCurrency()
    With Me.CboPayMentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "«Ã·"
        .AddItem "‘Ìﬂ „”œœ"
        '.AddItem "Õ”«»  "
    End With

    With Me.CboPaymentType1
        .Clear
        .AddItem "„’«—Ì›"
        .AddItem "Õ”«»« "
    
    End With

    With Me.CBoBasedON
        .Clear
        .AddItem "»·«"
        .AddItem "√„— ‘—¡"
        .AddItem "›« Ê—… „»œ∆ÌÂ"
        .AddItem " «„— «‰ «Ã  "
    
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    'StrSQL = StrSQL & " where Project_name<>N'""' and not (Project_name is null)"
    StrSQL = " select expanses_account,Project_name from projects  where   Project_name<>N'""' and not (Project_name is null) and not(expanses_account is null) "
     StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
   StrSQL = StrSQL & "    order by Project_name"
    fill_combo DCproject, StrSQL

    'StrSQL = " select  CusID, CusName from TblCustemers  where Type=2"
    'fill_combo Me.DCVendor, StrSQL

    Dcombos.GetCustomersSuppliers 4, Me.DCVendor




     With VSFlexGrid1
                     If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT ACCOUNTS.Account_Serial,  ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '   If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                    '   End If
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                    End If
                
                Else
                
                    StrSQL = "SELECT ACCOUNTS.Account_Serial, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '     If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
                   StrSQL = StrSQL & GetAccountByBarnchUser
                   StrSQL = StrSQL & GetAccountCodeHiding
                    '     End If
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                    End If
                
                End If
         Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName,Account_Serial", "Account_Code")
                 
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        ACCStrComboList = StrComboList
           End With



 
        If mdifrmmain.MnuProjects.Visible = False Then
        '     .ColHidden(.ColIndex("project")) = True
              Fg_Journal.ColHidden(Fg_Journal.ColIndex("pand")) = True
              Fg_Journal.ColHidden(Fg_Journal.ColIndex("oper")) = True
            
        End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From notes_all where notetype=80 and bill_Type<>2"
         
    If SystemOptions.usertype <> UserAdminAll Then
        StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
    StrSQL = "select * From notes_all where notetype=80 and bill_Type<>2 AND branch_no in(" & Current_branchSql & ")"
       
            If SystemOptions.FixedCustomer = 1 Then
                              StrSQL = StrSQL & " and  UserID = " & user_id
                               End If
                               
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 80
    hide_logo = False

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

    Set TTP = Nothing
    'Set EmpReport = Nothing
    TTD.Destroy
    Exit Sub
ErrTrap:
End Sub

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, _
                               ByVal SpinningEnded As Boolean)

    If ButtonID = vdsDownArrow Then
        If CboDes.IsDropped = False Then
            If PicHeight > 0 Then
                PicDes.Height = PicHeight
                PicDes.Width = PicWidth
            Else
                PicDes.Width = CboDes.Width - 10
                PicDes.Height = CboDes.Height * 8
            End If

            Debug.Print PicHeight
            Debug.Print PicWidth
            TxtDes.Visible = True
            TxtDes.Text = Fg_Journal.Cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Des"))
            CboDes.DropDown PicDes.hwnd, vdsRightToLeft, vdsBottomToDown, vdsDownArrow, True, vdsSoftResize
            Debug.Print PicDes.Height & "Pic H " & "-----" & PicDes.Width & "Pic W"
        Else
            CboDes.CloseUp
        End If
    End If

End Sub

Private Sub CboDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys "{F4}"
    End If

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption



End Sub

Private Sub PicDes_Resize()

    With PicDes
        LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
        TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
        '    PicHeight = PicDes.Height
        '    PicWidth = PicDes.Width
    End With

End Sub

Private Sub txt_Currency_rate_Change()
    If Me.TxtModFlg.Text = "" Or Me.TxtModFlg.Text = "R" Or val(txt_Currency_rate.Text) = 0 Then Exit Sub
    
    Dim i As Long
   
        For i = 1 To Fg_Journal.Rows - 1
            Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("value")) = val(Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("valuee"))) * val(txt_Currency_rate)
        Next
        
        For i = 1 To VSFlexGrid1.Rows - 1
            VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("value")) = val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("valuee"))) * val(txt_Currency_rate)
        Next
     
     
End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If CBoBasedON.ListIndex = 3 Then
        If KeyCode = vbKeyF3 Then
           Order_no_search2.show
             Order_no_search2.RetrunType = 2
         
        End If

    Else

        If KeyCode = vbKeyF3 Then
        'wael
            Order_no_search.show
             Order_no_search.RetrunType = 1
        End If

    End If

End Sub

Private Sub TxtDes_LostFocus()
    PicHeight = PicDes.Height
    PicWidth = PicDes.Width
    CboDes.CloseUp
    CboDes.Visible = False
End Sub

Private Sub TxtDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyEscape Then
        PutData
        CboDes.CloseUp
    End If

End Sub

Private Sub txtManulaVat_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "N" Then
        
        Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 3
  
  
        VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 3
            
   End If
   
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
        
            'Me.VSFlexGrid1.Enabled = False
          '  Me.Fg_Journal.Enabled = False
            Frame1.Enabled = False
        
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            CmdRemove.Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            XPTxtVal.locked = True
            '        XPCboProfLevel.Locked = True
            '        XPTxtProfMail.Locked = True
            '        XPTxtPhone.Locked = True
            '        XPTxtMobile.Locked = True
            XPMTxtRemarks.locked = True
            XPCboExpensesType.locked = True
            Me.DcboBox.locked = True
            XPDtbTrans.Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            
            End If

        Case "N"
        
            Me.VSFlexGrid1.Enabled = True
            Me.Fg_Journal.Enabled = True
            Frame1.Enabled = True

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            CmdRemove.Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '   Me.XPBtnMove(0).Enabled = False
            '   Me.XPBtnMove(1).Enabled = False
            '   Me.XPBtnMove(2).Enabled = False
            '   Me.XPBtnMove(3).Enabled = False
        
            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
        
            Me.VSFlexGrid1.Enabled = True
            Me.Fg_Journal.Enabled = True
            Frame1.Enabled = True
       
            CmdRemove.Enabled = True
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
        
            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtNoteSerial1_Change()
If TxtNoteSerial1.Text <> "" Then
Dim Type1 As Integer
Dim txtperson As String
Dim des As String
Dim EmpID As Integer
Dim Price As Double
If Me.TxtModFlg.Text <> "R" Then
OrderExchange TxtNoteSerial1.Text, Type1, txtperson, des, Price, EmpID
CboPayMentType.ListIndex = Type1
'txtto.text = txtperson
txt_general_des.Text = des
End If
End If
End Sub

Private Sub TxtNoteserial1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
            FrmReqExchangeSearch.show
            FrmReqExchangeSearch.lbltype.Caption = 3
        End If
End Sub
Sub DeleteGridCurrRowExp(Optional CurrRow As Long)
Dim i As Integer
With VSFlexGrid1
i = .Rows
Do
i = i - 1
If val(.TextMatrix(i, .ColIndex("CurrRow"))) = CurrRow Then
.RemoveItem i
End If
Loop While i > 1
End With
End Sub
Sub AddVATExp(Optional Row As Long)
If True = True Then
Dim ForcedFlg As Integer
Dim valuee As Double
Dim AccountVATDept As String
Dim i As Integer
Dim k As Integer
Dim ClsAcc  As New ClsAccounts
With VSFlexGrid1
.TextMatrix(Row, .ColIndex("Vatyo")) = PercentgValueAddedAccount(XPDtbTrans.value, .TextMatrix(Row, .ColIndex("AccountCode")), val(Dcbranch.BoundText), ForcedFlg)

If val(txtManulaVat.Text) > 0 Then
.TextMatrix(Row, .ColIndex("Vatyo")) = val(txtManulaVat.Text)

End If


.TextMatrix(Row, .ColIndex("Rate")) = val(.TextMatrix(Row, .ColIndex("Vatyo"))) / 100 + 1
If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) > 0 And val(.TextMatrix(Row, .ColIndex("Rate"))) > 0 Then
.TextMatrix(Row, .ColIndex("value")) = Round(val(.TextMatrix(Row, .ColIndex("PriceTotal"))) / val(.TextMatrix(Row, .ColIndex("Rate"))), 2)

If .TextMatrix(Row, .ColIndex("LineNo1")) = "" Then
.TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
End If

End If
valuee = val(.TextMatrix(Row, .ColIndex("Value")))

.TextMatrix(Row, .ColIndex("ForcedFlg")) = ForcedFlg
.TextMatrix(Row, .ColIndex("Vat")) = Round((val(.TextMatrix(Row, .ColIndex("Vatyo"))) * valuee) / 100, 2)
GetValueAddedAccount XPDtbTrans.value, AccountVATDept
If AccountVATDept = "" And val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· «·Õ”«» «·„œÌ‰ ›Ì ‘«‘… «⁄œ«œ  «·›« "
Else
MsgBox "Please Enter Account In VAT Settings"
End If
.TextMatrix(Row, .ColIndex("Vat")) = 0
.TextMatrix(Row, .ColIndex("Vatyo")) = 0
Exit Sub
End If
''/////////////
If val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
   If Not .TextMatrix(.Row, .ColIndex("AccountCode")) = "" Then
    DeleteGridCurrRowExp Row
   For i = 1 To 1
         .AddItem " ", .Row + i
  k = .Row + i
.TextMatrix(k, .ColIndex("CurrRow")) = Row
 
If i = 1 Then
.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(AccountVATDept)
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_name(, AccountVATDept)
.TextMatrix(k, .ColIndex("AccountCode")) = AccountVATDept
.TextMatrix(k, .ColIndex("Value")) = .TextMatrix(Row, .ColIndex("Vat"))
Else
.TextMatrix(k, .ColIndex("AccountCode")) = DcboCreditSide.BoundText
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_name(, DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("Value")) = .TextMatrix(Row, .ColIndex("Vat"))
End If
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("Des")) = .TextMatrix(Row, .ColIndex("Des")) & " " & " ﬁÌ„… „÷«›…"
Else
.TextMatrix(k, .ColIndex("Des")) = .TextMatrix(Row, .ColIndex("Des")) & " " & " VAT"
End If
.TextMatrix(k, .ColIndex("FlgVat")) = 1
Next i
End If
End If
End With
End If
End Sub
Sub DeleteGridCurrRow(Optional CurrRow As Long)
Dim i As Integer
With Fg_Journal
i = .Rows
Do
i = i - 1
If val(.TextMatrix(i, .ColIndex("CurrRow"))) = CurrRow Then
.RemoveItem i
End If
Loop While i > 1
End With
End Sub
Sub AddVAT(Optional Row As Long)
If True = True Then
Dim ForcedFlg As Integer
Dim valuee As Double
Dim AccountVATDept As String
Dim i As Integer
Dim k As Integer
Dim ClsAcc  As New ClsAccounts
With Fg_Journal
.TextMatrix(Row, .ColIndex("Vatyo")) = PercentgValueAddedAccount(XPDtbTrans.value, .TextMatrix(Row, .ColIndex("AccountCode")), val(Dcbranch.BoundText), ForcedFlg)

If val(txtManulaVat.Text) > 0 Then
.TextMatrix(Row, .ColIndex("Vatyo")) = val(txtManulaVat.Text)

End If


.TextMatrix(Row, .ColIndex("Rate")) = val(.TextMatrix(Row, .ColIndex("Vatyo"))) / 100 + 1
If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) > 0 And val(.TextMatrix(Row, .ColIndex("Rate"))) > 0 Then
.TextMatrix(Row, .ColIndex("value")) = Round(val(.TextMatrix(Row, .ColIndex("PriceTotal"))) / val(.TextMatrix(Row, .ColIndex("Rate"))), 2)
End If
valuee = val(.TextMatrix(Row, .ColIndex("value")))

.TextMatrix(Row, .ColIndex("ForcedFlg")) = ForcedFlg
.TextMatrix(Row, .ColIndex("Vat")) = Round((val(.TextMatrix(Row, .ColIndex("Vatyo"))) * valuee) / 100, 2)
GetValueAddedAccount XPDtbTrans.value, AccountVATDept
If AccountVATDept = "" And val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· «·Õ”«» «·„œÌ‰ ›Ì ‘«‘… «⁄œ«œ  «·›« "
Else
MsgBox "Please Enter Account In VAT Settings"
End If
.TextMatrix(Row, .ColIndex("Vat")) = 0
.TextMatrix(Row, .ColIndex("Vatyo")) = 0
Exit Sub
End If
''/////////////
If val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
   If Not .TextMatrix(.Row, .ColIndex("AccountCode")) = "" Then
    DeleteGridCurrRow Row
   For i = 1 To 1
         .AddItem " ", .Row + i
  k = .Row + i
.TextMatrix(k, .ColIndex("CurrRow")) = Row
 
If i = 1 Then
'.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(AccountVATDept)
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_name(, AccountVATDept)
.TextMatrix(k, .ColIndex("AccountCode")) = AccountVATDept
.TextMatrix(k, .ColIndex("value")) = .TextMatrix(Row, .ColIndex("Vat"))
Else
.TextMatrix(k, .ColIndex("AccountCode")) = DcboCreditSide.BoundText
.TextMatrix(k, .ColIndex("AccountName")) = Get_Account_name(, DcboCreditSide.BoundText)
'.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(DcboCreditSide.BoundText)
.TextMatrix(k, .ColIndex("Value")) = .TextMatrix(Row, .ColIndex("Vat"))
End If
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("des")) = .TextMatrix(Row, .ColIndex("des")) & " " & " ﬁÌ„… „÷«›…"
Else
.TextMatrix(k, .ColIndex("des")) = .TextMatrix(Row, .ColIndex("des")) & " " & " VAT  "
End If
.TextMatrix(k, .ColIndex("FlgVat")) = 1
.TextMatrix(k, .ColIndex("ExpensesID")) = .TextMatrix(Row, .ColIndex("ExpensesID"))
.TextMatrix(k, .ColIndex("opr_fullcode")) = .TextMatrix(Row, .ColIndex("opr_fullcode"))
.TextMatrix(k, .ColIndex("CarName")) = .TextMatrix(Row, .ColIndex("CarName"))
.TextMatrix(k, .ColIndex("Order_No")) = .TextMatrix(Row, .ColIndex("Order_No"))
.TextMatrix(k, .ColIndex("CarId")) = .TextMatrix(Row, .ColIndex("CarId"))
Next i
End If
End If
End With
End If
End Sub

Private Sub TxtVATCustoms_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtVATCustoms.Text, 0)
End Sub

Public Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                 ByVal Col As Long)
    'check_cost_center
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
 
    With VSFlexGrid1

        Select Case .ColKey(Col)
           Case "Vatyo"
        If val(.TextMatrix(Row, .ColIndex("Vatyo"))) = 0 Then
        .TextMatrix(Row, .ColIndex("Vat")) = 0
        If val(.TextMatrix(Row, .ColIndex("PriceTotal"))) <> 0 Then
        .TextMatrix(Row, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("PriceTotal")))
        End If
        If .Rows > Row Then
        If val(.TextMatrix(Row + 1, .ColIndex("FlgVat"))) = 1 Then
        .RemoveItem Row + 1
        End If
        End If
        End If
   
    Case "DeparmentName"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("DeparmentID"), False, True)
                .TextMatrix(Row, .ColIndex("DeparmentID")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("DeparmentName")) = .TextMatrix(Row, .ColIndex("DeparmentName"))
        
             Case "pand"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("pandid")) = StrAccountCode
                 
    Case "PriceTotal"
          AddVATExp Row
    Case "ProjectName"
    
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ProjectID"), False, True)
                .TextMatrix(Row, .ColIndex("ProjectID")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ProjectName")) = .TextMatrix(Row, .ColIndex("ProjectName"))
       '         StrSQL = "Select expanses_account,REVENUE_account From projects Where Id=" & val(.TextMatrix(Row, .ColIndex("ProjectID")))
       '         Set rs = New ADODB.Recordset
       '         rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
       '         Dim AccountCode As Long
'
'
'                If rs.RecordCount > 0 Then
'                    .TextMatrix(Row, .ColIndex("Account_Serial")) = Get_Account_Serial(IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value))
'                    .TextMatrix(Row, .ColIndex("AccountName")) = GET_ACCOUNT_name_by_Code(IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value))
'                    .TextMatrix(Row, .ColIndex("AccountCode")) = Get_Account_Serial(IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value))
'
'                Else
'                    .TextMatrix(Row, .ColIndex("AccountName")) = ""
'                End If
'
'
           '  AddVATExp Row
          
          
    Case "Supplier"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("SupplierID"), False, True)
                .TextMatrix(Row, .ColIndex("SupplierID")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("SupplierName")) = .TextMatrix(Row, .ColIndex("Supplier"))
                StrSQL = "Select * From TblCustemers Where CusID=" & val(.TextMatrix(Row, .ColIndex("SupplierID")))
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("CusVATNO")) = IIf(IsNull(rs("VATNO").value), "", rs("VATNO").value)
                Else
                .TextMatrix(Row, .ColIndex("CusVATNO")) = ""
                End If
                
             AddVATExp Row
        Case "PFuLLCode"
                .TextMatrix(Row, .ColIndex("ProjectID")) = ""
                .TextMatrix(Row, .ColIndex("ProjectName")) = ""
                StrSQL = "Select expanses_account,REVENUE_account,id,Fullcode,Project_name From projects Where Fullcode='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = Get_Account_Serial(IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value))
                    .TextMatrix(Row, .ColIndex("AccountName")) = GET_ACCOUNT_name_by_Code(IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value))
                    .TextMatrix(Row, .ColIndex("AccountCode")) = Get_Account_code(IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value))
                    .TextMatrix(Row, .ColIndex("ProjectName")) = rs!Project_name & ""
                    .TextMatrix(Row, .ColIndex("ProjectID")) = rs!ID & ""
                Else
                    .TextMatrix(Row, .ColIndex("AccountName")) = ""
                    .TextMatrix(Row, .ColIndex("PFuLLCode")) = ""
                    .TextMatrix(Row, .ColIndex("ProjectID")) = ""
                    .TextMatrix(Row, .ColIndex("ProjectName")) = ""
                End If
        Case "ProjectName"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ProjectID"), False, True)
                .TextMatrix(Row, .ColIndex("ProjectID")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ProjectName")) = .TextMatrix(Row, .ColIndex("ProjectName"))
                StrSQL = "Select expanses_account,REVENUE_account From projects Where Id=" & val(.TextMatrix(Row, .ColIndex("ProjectID")))
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = Get_Account_Serial(IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value))
                    .TextMatrix(Row, .ColIndex("AccountName")) = GET_ACCOUNT_name_by_Code(IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value))
                    .TextMatrix(Row, .ColIndex("AccountCode")) = Get_Account_code(IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value))
                Else
                    .TextMatrix(Row, .ColIndex("AccountName")) = ""
                End If
'
           '  AddVATExp Row
            Case "Value"
                 .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                 .TextMatrix(Row, .ColIndex("valuee")) = Round(val(.TextMatrix(Row, .ColIndex("value"))) / val(txt_Currency_rate), 3)
                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
              AddVATExp Row
            Case "Valuee"
                 .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                 .TextMatrix(Row, .ColIndex("value")) = Round(val(.TextMatrix(Row, .ColIndex("valuee"))) * val(txt_Currency_rate), 3)
                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
              AddVATExp Row
              
            Case "DebitValue", "CreditValue"

                'remove destribution
     
                ' sgl = "update  marakes_taklefa_temp  set value=0 where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                ' Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))
            
                If .ColKey(Col) = "DebitValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                    ' Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                 
                    '    Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValue" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                    ' Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '     Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If

                .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
            
            Case "DebitValueE", "CreditValueE"
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))
                .TextMatrix(Row, .ColIndex("rate")) = val(txt_Currency_rate)
                If .ColKey(Col) = "DebitValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                    .TextMatrix(Row, .ColIndex("rate")) = val(txt_Currency_rate)
                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE"))
                    End If

                    '
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValueE" Then
                    .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE"))
                    End If
                 
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If
                     Case "pand"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("pandid")) = StrAccountCode
                  Case "oper"
                StrAccountCode = .ComboData
                .TextMatrix(Row, .ColIndex("operid")) = StrAccountCode
         
            Case "Account_Serial"
            
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                StrSQL = StrSQL & GetAccountByBarnchUser
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    If BolEditOnMainAccounts = False Then
                        'If LastAccount(rs("Account_Code").value) = False Then
                        '    .TextMatrix(Row, Col) = ""
                        '    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                        '    Exit Sub
                        'End If
                    End If

                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    
                    End If
                    AddVATExp Row
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)
                    
                    Dim rs2 As ADODB.Recordset
                    Dim My_SQL As String

                    If IsNull(rs("currenct_code").value) Then

                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                    
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo xx
                    End If

                    My_SQL = "  select * from currency WHERE id=" & val(rs("currenct_code").value)

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
xx:
                Else
                    'GetMsgs 130, vbExclamation
                    If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "ﬂÊœ Õ”«» Œ«ÿÏ¡", vbCritical
                    Else
                    MsgBox "Error Code Account", vbCritical
                    End If
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                
                    Exit Sub
                End If

                rs.Close
                Set rs = Nothing

            Case "AccountName"
        
                'sgl = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                'Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)

                If LngRow <> -1 Then
                    'Msg = "Â–« «·Õ”«» „ÊÃÊœ „”»ﬁ«  ›Ï «·”ÿ— " & .TextMatrix(LngRow, .ColIndex("LineNo"))
                    'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    '.TextMatrix(Row, Col) = ""
                    '.TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Exit Sub
                End If

                Set ClsAcc = New ClsAccounts

                If BolEditOnMainAccounts = False Then
                    'If LastAccount(StrAccountCode) = False Then
                    '    .TextMatrix(Row, Col) = ""
                    '    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Else

                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                    'End If
                Else
                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
 
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                End If
                AddVATExp Row
                Set ClsAcc = Nothing
            
                StrSQL = "SELECT ACCOUNTS.cost_center ,ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Name='" & Trim(.TextMatrix(Row, Col)) & "'"
                StrSQL = StrSQL & GetAccountByBarnchUser
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), vbFalse, rs("cost_center").value)
            
                    'Dim rs2 As ADODB.Recordset
                    'Dim My_SQL As String
                    If IsNull(rs("currenct_code").value) Then
                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo ll
                    End If

                    My_SQL = "  select * from currency WHERE id=" & rs("currenct_code").value

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value)
ll:
                End If

        End Select

        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ReLineGrid

    End With

    With Me.VSFlexGrid1

        If Me.TxtModFlg <> "E" Then Exit Sub

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
        If Col = .ColIndex("Account_Serial") Or Col = .ColIndex("AccountName") Then
            LogTextA = "   ⁄œÌ· «·Õ”«» «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Account To " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("Value") Then
            LogTextA = "   ⁄œÌ· «·ﬁÌ„…  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change value" & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        ElseIf Col = .ColIndex("Des") Then
            LogTextA = "   ⁄œÌ· «·‘—Õ  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
            LogTexte = "  Change Des " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        End If

        AddToLogFile CInt(user_id), 80, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
    End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

      If val(.TextMatrix(Row, .ColIndex("FlgVat"))) <> 0 Then
            Cancel = True
     Else
 Select Case .ColKey(Col)
        Case "Vat"
                 Cancel = True
        Case "Vatyo"
              If val(.TextMatrix(Row, .ColIndex("ForcedFlg"))) = 1 Then
                 Cancel = True
              Else
              .ComboList = ""
              End If
                Case "BillNo", "PFuLLCode"
                .ComboList = ""
         
         Case "CusVATNO"
                .ComboList = ""
         Case "SupplierName"
                .ComboList = ""
         Case "ProjectName"
                .ComboList = ""
                
         Case "PriceTotal"
                .ComboList = ""
     Case "LineNo"
                .ComboList = ""
     
     
            Case "Value"
                .ComboList = ""

            Case "Account_Serial"
                .ComboList = ""
                '  Cancel = True
            
        End Select
      End If
    End With

End Sub

Private Sub VSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
  SendKeys "{F4}"
  SendKeys "{BACKSPACE}"
  SendKeys CHR(KeyAscii)
End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)
 
    

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 80
 
    End If

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid1

        Select Case .ColKey(Col)
           Case "Supplier"
            StrSQL = " SELECT     CusID, CusName, CusNamee"
            StrSQL = StrSQL & "            From dbo.TblCustemers"
            StrSQL = StrSQL & "       WHERE     (Type = 2)"
         Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "CusName", "CusID")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "CusNamee", "CusID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                
        Case "DeparmentName"
           
            If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "Select DeparmentID,DepartmentName From TblEmpDepartments Order By DepartmentName"
            Else
                StrSQL = "Select DeparmentID,DepartmentNamee From TblEmpDepartments Order By DepartmentNamee"
            End If
         Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "DepartmentNamee", "DeparmentID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                '
         Case "ProjectName"
            'StrSQL = " SELECT     ID, Project_Name"
                 If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = " SELECT     ID, LTRIM(RTRIM( Project_name )) as Project_name"
         Else
         StrSQL = " SELECT     ID, LTRIM(RTRIM( Project_namee )) as Project_name"
         End If
         
            StrSQL = StrSQL & "            From dbo.projects"
            
            
            
        If SystemOptions.UserInterface = ArabicInterface Then
                 
                StrSQL = StrSQL & " where Project_name<>N'""' and not (Project_name is null)"
   Else
     
                StrSQL = StrSQL & " where Project_nameE<>N'""' and not (Project_nameE is null)"
End If


                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Project_Name", "ID")
                Else
                    StrComboList = .BuildComboList(rs, "Project_Name", "ID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                
             Case "pand"
             If .TextMatrix(Row, .ColIndex("projectid")) = "" Then
             MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
             Exit Sub
             End If

                StrSQL = " SELECT     des, oprid From projects_des "
                 StrSQL = StrSQL & "    Where (project_id =" & val(.TextMatrix(Row, .ColIndex("projectid"))) & ")"
                Set Rs3 = New ADODB.Recordset
                Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    StrComboList = VSFlexGrid1.BuildComboList(Rs3, "des", "oprid")
           

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
            Case "Value", "valuee"
                .ComboList = ""
                
                
                
                
                
            Case "AccountName"
          .ComboList = ACCStrComboList
Exit Sub
                'Full Path Display
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT  ACCOUNTS.Account_Serial,ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '   If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                    '   End If
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                    End If
                
                Else
                
                    StrSQL = "SELECT  ACCOUNTS.Account_Serial,ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '     If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
                   StrSQL = StrSQL & GetAccountByBarnchUser
                   StrSQL = StrSQL & GetAccountCodeHiding
                    '     End If
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                    End If
                
                End If
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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
Function CheAssetPayd(Optional NoteID As Long = 0) As Boolean
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
CheAssetPayd = False
sql = "select NoteID from notes_all where NoteID=" & NoteID & " and (AssestPayd =1) "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CheAssetPayd = True
Else
CheAssetPayd = False
End If
End Function
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim CarID As Integer
    Dim CarName As String

    On Error GoTo ErrTrap
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 3
          
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
          
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        'Lngid
        '  If XPTxtID.text <> 0 Then
        '      Rs.find "NoteID=" & XPTxtID.text, , adSearchForward, adBookmarkFirst
        '      If Rs.EOF Or Rs.BOF Then
        '          Exit Sub
        '      End If
        '  End If
  
        If Lngid <> 0 Then
            rs.Find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If

    End If

    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    Else
        Me.DcCostCenter.BoundText = ""
    End If
    txtManulaVat.Text = IIf(IsNull(rs("txtManulaVat").value), 0, (rs("txtManulaVat").value))
    
    TxtVATCustoms.Text = IIf(IsNull(rs("VATCustoms").value), 0, rs("VATCustoms").value)
    Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.Text1.Text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    Me.Txt_order_no.Text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
    Me.txtOrderID.Text = IIf(IsNull(rs("OrderID").value), "", rs("OrderID").value)
    Me.TxtNoteSerial1.Text = IIf(IsNull(rs("Noteseril2").value), "", rs("Noteseril2").value)
    TXT_A_NoteID.Text = IIf(IsNull(rs("A_NoteID").value), "", (rs("A_NoteID").value))
    Me.DcCurrency.BoundText = IIf(IsNull(rs("CurrncyID").value), MainCurrency(), rs("CurrncyID").value)
    If DcCurrency.Text = "" Then
        Me.DcCurrency.BoundText = MainCurrency()
    End If
    
    XPTxtID.Text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    XPTxtVal.Text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
    XPMTxtRemarks.Text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    txtTo.Text = IIf(IsNull(rs("too").value), "", rs("too").value)
    txt_general_des.Text = IIf(IsNull(rs("general_des").value), "", rs("general_des").value)
txt_Currency_rate.Text = IIf(IsNull(rs("rate").value), 1, (rs("rate").value))
    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    XPCboExpensesType.BoundText = IIf(IsNull(rs("ExpensesID").value), "", rs("ExpensesID").value)

    If (rs("bill_Type").value) = 0 Then
        Me.CboPaymentType1.ListIndex = 0
    ElseIf (rs("bill_Type").value) = 1 Then
        Me.CboPaymentType1.ListIndex = 1
    Else
        Me.CboPaymentType1.ListIndex = 0
    End If

    CboPaymentType1_Change

    If Not IsNull(rs("BasedONID").value) Then
        Me.CBoBasedON.ListIndex = rs("BasedONID").value
    Else
        Me.CBoBasedON.ListIndex = 0
 
    End If

    If IsNull(rs("NoteCashingType").value) Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        DCVendor.BoundText = ""
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        DCVendor.BoundText = ""
    ElseIf rs("NoteCashingType").value = 1 Then
        Me.CboPayMentType.ListIndex = 1
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        DCVendor.BoundText = ""

    ElseIf rs("NoteCashingType").value = 3 Then
        Me.CboPayMentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    
    ElseIf rs("NoteCashingType").value = 2 Then
        Me.CboPayMentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
    
        Me.DCVendor.BoundText = rs("CusID").value

    End If

    CboPayMentType_Change

    'ÿMe.DcboBox.BoundText = IIf(IsNull(Rs("BoxID").value), "", Rs("BoxID").value)
    'DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))

    If rs("NoteCashingType").value = 0 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    ElseIf rs("NoteCashingType").value = 1 Or rs("NoteCashingType").value = 3 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DcboBankName.BoundText))
    ElseIf rs("NoteCashingType").value = 2 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DCVendor.BoundText))
    End If

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.Txt_Numorder.Text = IIf(IsNull(rs("NumOrderInpot").value), "", rs("NumOrderInpot").value)
    Me.TxtSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)

    Me.TxtSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)

    Me.oldTxtSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(27).Caption = showLabel(TxtSerial1, oldTxtSerial1)

    Me.DCproject.BoundText = IIf(IsNull(rs("project_Expensen_account").value), "", rs("project_Expensen_account").value)

    If CboPaymentType1.ListIndex = 1 Then 'Õ”«Ì« 

        StrSQL = " SELECT       dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial,Notes.project_id ProjectID,projects.project_Name ProjectName ,"
        StrSQL = StrSQL + "              dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.UserID,"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.DOUBLE_ENTREY_VOUCHERS.FlgVat, dbo.DOUBLE_ENTREY_VOUCHERS.Vatyo,"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.Vat, dbo.DOUBLE_ENTREY_VOUCHERS.CurrRow, dbo.DOUBLE_ENTREY_VOUCHERS.Billno,"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.Rate2, dbo.DOUBLE_ENTREY_VOUCHERS.SupplierName, dbo.DOUBLE_ENTREY_VOUCHERS.CusVATNO,"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.PriceTotal, dbo.DOUBLE_ENTREY_VOUCHERS.SupplierID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
        StrSQL = StrSQL + "              dbo.TblCustemers.fullcode"
        StrSQL = StrSQL + "      FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
        StrSQL = StrSQL + "              dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.TblCustemers ON dbo.DOUBLE_ENTREY_VOUCHERS.SupplierID = dbo.TblCustemers.CusID"
        
        StrSQL = StrSQL + "             Left Outer join  dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID "
       
       
        
        StrSQL = StrSQL + "              Left Outer Join dbo.projects ON dbo.projects.ID = dbo.Notes.project_id"
        
        StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.hideline = 0 or dbo.DOUBLE_ENTREY_VOUCHERS.hideline is null)and (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = " & val(rs("A_NoteID").value) & ")"
        StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"

StrSQL = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial,DOUBLE_ENTREY_VOUCHERS.Departementid DeparmentID,TblEmpDepartments.DepartmentName, "
  StrSQL = StrSQL + "                       dbo.Notes.project_id AS ProjectID, dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.UserID,"
  StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
  StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.DOUBLE_ENTREY_VOUCHERS.FlgVat, dbo.DOUBLE_ENTREY_VOUCHERS.Vatyo,"
  StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Vat, dbo.DOUBLE_ENTREY_VOUCHERS.CurrRow, dbo.DOUBLE_ENTREY_VOUCHERS.Billno,"
  StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.Rate2, dbo.DOUBLE_ENTREY_VOUCHERS.SupplierName, dbo.DOUBLE_ENTREY_VOUCHERS.CusVATNO, dbo.DOUBLE_ENTREY_VOUCHERS.pandid, dbo.projects_des.des,"
  StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.PriceTotal, dbo.DOUBLE_ENTREY_VOUCHERS.SupplierID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
  StrSQL = StrSQL + "                       dbo.TblCustemers.Fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.project_id, dbo.projects.Fullcode AS PFuLLCode, dbo.projects.Project_name ProjectName,"
  StrSQL = StrSQL + "                       dbo.DOUBLE_ENTREY_VOUCHERS.operid, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE,"
  StrSQL = StrSQL + "                       dbo.Projects.Project_nameE"
  StrSQL = StrSQL + " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
  StrSQL = StrSQL + "                       dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"

  

StrSQL = StrSQL + "                       dbo.TblProcessDEF ON dbo.DOUBLE_ENTREY_VOUCHERS.operid = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
StrSQL = StrSQL + "                       dbo.projects_des ON dbo.DOUBLE_ENTREY_VOUCHERS.pandid = dbo.projects_des.oprid AND dbo.projects_des.oprid <> 0 LEFT OUTER JOIN"
  StrSQL = StrSQL + "                       dbo.projects ON dbo.DOUBLE_ENTREY_VOUCHERS.project_id = dbo.projects.id LEFT OUTER JOIN"
  
 
 
  
  StrSQL = StrSQL + "                       dbo.TblCustemers ON dbo.DOUBLE_ENTREY_VOUCHERS.SupplierID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
  StrSQL = StrSQL + "                       dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
  StrSQL = StrSQL + "              LEFT OUTER JOIN"
StrSQL = StrSQL + "              TblEmpDepartments On TblEmpDepartments.DeparmentID = DOUBLE_ENTREY_VOUCHERS.Departementid "
      StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.hideline = 0 or dbo.DOUBLE_ENTREY_VOUCHERS.hideline is null)and (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = " & val(rs("A_NoteID").value) & ")"
        StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
                      
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If RsDev.RecordCount > 0 Then
            RsDev.MoveFirst
        End If
    
        With Me.VSFlexGrid1
 
            .Rows = .FixedRows + RsDev.RecordCount
 
            For i = .FixedRows To .Rows - 1
                 .TextMatrix(i, .ColIndex("LineNo")) = i
                 
                 .TextMatrix(i, .ColIndex("SupplierID")) = IIf(IsNull(RsDev("SupplierID").value), 0, RsDev("SupplierID").value)
                 .TextMatrix(i, .ColIndex("CusVATNO")) = IIf(IsNull(RsDev("CusVATNO").value), "", RsDev("CusVATNO").value)
                 .TextMatrix(i, .ColIndex("SupplierName")) = IIf(IsNull(RsDev("SupplierName").value), "", RsDev("SupplierName").value)
                 .TextMatrix(i, .ColIndex("PriceTotal")) = IIf(IsNull(RsDev("PriceTotal").value), 0, RsDev("PriceTotal").value)
                 .TextMatrix(i, .ColIndex("Rate")) = IIf(IsNull(RsDev("Rate2").value), 0, RsDev("Rate2").value)
                 .TextMatrix(i, .ColIndex("BillNo")) = IIf(IsNull(RsDev("BillNo").value), "", RsDev("BillNo").value)
                 .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), "", RsDev("DEV_ID_Line_No1").value)
                 .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
                 .TextMatrix(i, .ColIndex("FlgVat")) = IIf(IsNull(RsDev("FlgVat").value), 0, RsDev("FlgVat").value)
                 .TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(RsDev("Vatyo").value), 0, RsDev("Vatyo").value)
                 .TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(RsDev("Vat").value), 0, RsDev("Vat").value)
                 .TextMatrix(i, .ColIndex("CurrRow")) = IIf(IsNull(RsDev("CurrRow").value), 0, RsDev("CurrRow").value)
                 
                 .TextMatrix(i, .ColIndex("pandid")) = IIf(IsNull(RsDev("pandid").value), "", RsDev("pandid").value)
                .TextMatrix(i, .ColIndex("pand")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
                .TextMatrix(i, .ColIndex("operid")) = IIf(IsNull(RsDev("operid").value), "", RsDev("operid").value)
                .TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(RsDev("ProcessName").value), "", RsDev("ProcessName").value)
                .TextMatrix(i, .ColIndex("DeparmentID")) = IIf(IsNull(RsDev("DeparmentID").value), "", RsDev("DeparmentID").value)
                .TextMatrix(i, .ColIndex("DeparmentName")) = IIf(IsNull(RsDev("DepartmentName").value), "", RsDev("DepartmentName").value)
              
              .TextMatrix(i, .ColIndex("PFuLLCode")) = IIf(IsNull(RsDev("PFuLLCode").value), "", RsDev("PFuLLCode").value)
                .TextMatrix(i, .ColIndex("ProjectID")) = IIf(IsNull(RsDev("ProjectID").value), "", RsDev("ProjectID").value)
                
                 .TextMatrix(i, .ColIndex("account_serial")) = IIf(IsNull(RsDev("account_serial").value), "", RsDev("account_serial").value)
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Supplier")) = IIf(IsNull(RsDev("CusName").value), "", RsDev("CusName").value)
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                    .TextMatrix(i, .ColIndex("ProjectName")) = IIf(IsNull(RsDev("ProjectName").value), "", RsDev("ProjectName").value)
                Else
                    .TextMatrix(i, .ColIndex("Supplier")) = IIf(IsNull(RsDev("CusNamee").value), "", RsDev("CusNamee").value)
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                    .TextMatrix(i, .ColIndex("ProjectName")) = IIf(IsNull(RsDev("Project_nameE").value), "", RsDev("Project_nameE").value)
                End If
        
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
                .TextMatrix(i, .ColIndex("valuee")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value) / val(txt_Currency_rate)
                .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), "", RsDev("Double_Entry_Vouchers_Description").value)
            
                RsDev.MoveNext
            Next i
    
        End With

        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        ReLineGrid
        Exit Sub
    End If

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then '«·„—Ê›« 
        '   StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(Me.XPTxtID.text)
        '   StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        ' StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.*,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name FROM    dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code WHERE     dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID =" & Val(Me.XPTxtID.text) & "Order By DEV_ID_Line_No"

        'StrSQL = "SELECT   dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode,   dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit , dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID ,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description  FROM         dbo.ACCOUNTS INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
        'StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0  and dbo.DOUBLE_ENTREY_VOUCHERS.notes_all =" & Val(Me.XPTxtID.text) & ") "
        'StrSQL = StrSQL + "ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
        StrSQL = " SELECT       Notes.project_id, Notes.DeparmentID,TblEmpDepartments.DepartmentName, dbo.DOUBLE_ENTREY_VOUCHERS.Carid, dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode,Notes.project_id ProjectID,projects.project_Name ProjectName,projects.project_Namee ProjectNamee , dbo.projects.Fullcode AS PFuLLCode , "
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,dbo.Notes.Pand, dbo.projects_des.des, dbo.Notes.Oper,"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng,"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID AS Expr1,"
        StrSQL = StrSQL + "               dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.Notes.ORDER_NO, dbo.DOUBLE_ENTREY_VOUCHERS.CurrRow,"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.Vat, dbo.DOUBLE_ENTREY_VOUCHERS.Vatyo, dbo.DOUBLE_ENTREY_VOUCHERS.FlgVat,"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.Rate2, dbo.DOUBLE_ENTREY_VOUCHERS.SupplierName, dbo.DOUBLE_ENTREY_VOUCHERS.CusVATNO,"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS.PriceTotal, dbo.DOUBLE_ENTREY_VOUCHERS.SupplierID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
        StrSQL = StrSQL + "              dbo.TblCustemers.fullcode , dbo.DOUBLE_ENTREY_VOUCHERS.BillNo,dbo.DOUBLE_ENTREY_VOUCHERS.operid,"
        
        StrSQL = StrSQL + "                        dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE "
        StrSQL = StrSQL + "    FROM         dbo.ACCOUNTS INNER JOIN"
        StrSQL = StrSQL + "              dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
        
        StrSQL = StrSQL + "              dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
        StrSQL = StrSQL + "              dbo.TblCustemers ON dbo.DOUBLE_ENTREY_VOUCHERS.SupplierID = dbo.TblCustemers.CusID"
        StrSQL = StrSQL + "              Left Outer Join dbo.projects ON dbo.projects.ID = dbo.Notes.project_id"
        
          
        StrSQL = StrSQL + "              LEFT OUTER JOIN"
        StrSQL = StrSQL + "            dbo.projects_des ON dbo.Notes.Pand = dbo.projects_des.oprid AND dbo.projects_des.oprid <> 0 LEFT OUTER JOIN"
         StrSQL = StrSQL + "                       dbo.TblProcessDEF ON dbo.DOUBLE_ENTREY_VOUCHERS.operid = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
          
        StrSQL = StrSQL + "              TblEmpDepartments On TblEmpDepartments.DeparmentID = Notes.DeparmentID "
        StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.hideline = 0 or dbo.DOUBLE_ENTREY_VOUCHERS.hideline is null)and(dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & val(Me.XPTxtID.Text) & ")"
        StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
    
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For i = 1 To RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                End If

                RsDev.MoveNext
            Next i
    
            RsDev.MoveFirst
    
            With Me.Fg_Journal

                If Me.DCproject.BoundText = "" Then
                    .Rows = .FixedRows + RsDev.RecordCount
                Else
                    .Rows = .FixedRows + RsDev.RecordCount - 1
                End If

                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("SupplierID")) = IIf(IsNull(RsDev("SupplierID").value), 0, RsDev("SupplierID").value)
                    .TextMatrix(i, .ColIndex("CusVATNO")) = IIf(IsNull(RsDev("CusVATNO").value), "", RsDev("CusVATNO").value)
                    .TextMatrix(i, .ColIndex("SupplierName")) = IIf(IsNull(RsDev("SupplierName").value), "", RsDev("SupplierName").value)
                    .TextMatrix(i, .ColIndex("PriceTotal")) = IIf(IsNull(RsDev("PriceTotal").value), 0, RsDev("PriceTotal").value)
                    .TextMatrix(i, .ColIndex("Rate")) = IIf(IsNull(RsDev("Rate2").value), 0, RsDev("Rate2").value)
                    .TextMatrix(i, .ColIndex("BillNo")) = IIf(IsNull(RsDev("BillNo").value), "", RsDev("BillNo").value)
                    .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("DEV_ID_Line_No").value), "", RsDev("DEV_ID_Line_No").value)
                    .TextMatrix(i, .ColIndex("FlgVat")) = IIf(IsNull(RsDev("FlgVat").value), 0, RsDev("FlgVat").value)
                    .TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(RsDev("Vatyo").value), 0, RsDev("Vatyo").value)
                    .TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(RsDev("Vat").value), 0, RsDev("Vat").value)
                    .TextMatrix(i, .ColIndex("CurrRow")) = IIf(IsNull(RsDev("CurrRow").value), 0, RsDev("CurrRow").value)
                    .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), "", RsDev("DEV_ID_Line_No1").value)
                    .TextMatrix(i, .ColIndex("ProjectID")) = IIf(IsNull(RsDev("project_id").value), "", RsDev("project_id").value)
                    .TextMatrix(i, .ColIndex("DeparmentID")) = IIf(IsNull(RsDev("DeparmentID").value), "", RsDev("DeparmentID").value)
                    .TextMatrix(i, .ColIndex("DeparmentName")) = IIf(IsNull(RsDev("DepartmentName").value), "", RsDev("DepartmentName").value)
                   
                    .TextMatrix(i, .ColIndex("pandid")) = IIf(IsNull(RsDev("Pand").value), "", RsDev("Pand").value)
                    .TextMatrix(i, .ColIndex("pand")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
                    
                    .TextMatrix(i, .ColIndex("operid")) = IIf(IsNull(RsDev("operid").value), "", RsDev("operid").value)
                    .TextMatrix(i, .ColIndex("oper")) = IIf(IsNull(RsDev("ProcessName").value), "", RsDev("ProcessName").value)
                            .TextMatrix(i, .ColIndex("ProjectID")) = IIf(IsNull(RsDev("project_id").value), "", RsDev("project_id").value)
                
                .TextMatrix(i, .ColIndex("PFuLLCode")) = IIf(IsNull(RsDev("PFuLLCode").value), "", RsDev("PFuLLCode").value)
                
               
                    .TextMatrix(i, .ColIndex("ExpensesID")) = get_Expenses_id(IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value))
            
                    .TextMatrix(i, .ColIndex("opr_fullcode")) = IIf(IsNull(RsDev("opr_fullcode").value), "", RsDev("opr_fullcode").value)
            
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("Supplier")) = IIf(IsNull(RsDev("CusName").value), "", RsDev("CusName").value)
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                        .TextMatrix(i, .ColIndex("ProjectName")) = IIf(IsNull(RsDev("ProjectName").value), "", RsDev("ProjectName").value)
                    Else
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                        .TextMatrix(i, .ColIndex("Supplier")) = IIf(IsNull(RsDev("CusNamee").value), "", RsDev("CusNamee").value)
                        .TextMatrix(i, .ColIndex("ProjectName")) = IIf(IsNull(RsDev("ProjectNamee").value), "", RsDev("ProjectNamee").value)
                    End If

                    'Double_Entry_Vouchers_Description
                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), "", RsDev("Double_Entry_Vouchers_Description").value)
            
                    '    .TextMatrix(I, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), _
                    '        "", RsDev("Account_Name").value)
        
                    .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
                    .TextMatrix(i, .ColIndex("valuee")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value) / val(txt_Currency_rate)
                    .TextMatrix(i, .ColIndex("Order_No")) = IIf(IsNull(RsDev("Order_No").value), "", RsDev("Order_No").value)
 
                    CarID = IIf(IsNull(RsDev("CarID").value), 0, RsDev("CarID").value)

                    If CarID <> 0 Then
                        GetCarName CarID, CarName
                        .TextMatrix(i, .ColIndex("CarId")) = IIf(IsNull(RsDev("CarID").value), "", RsDev("CarID").value)
             
                        .TextMatrix(i, .ColIndex("CarName")) = CarName
                 
                    End If

                    RsDev.MoveNext
                Next i

                Me.XPTxtVal.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
                ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
                '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), _
                '  .Rows - 1, .ColIndex("CreditValue"))
                '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), _
                '  .Rows - 1, .ColIndex("DebitValue"))
            End With

        End If

    End If

    '-----------------------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    fillapprovData
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim i As Integer
Dim Balance As String
Dim balanceString As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
    Dim OtherInformation As New ClsGLOther
    Dim StrAccount As String
    On Error GoTo ErrTrap
     Dim Posted As Integer
            If CheckAprroveScreen(Me.Name) = True Then
            Posted = 1
            Else
            Posted = 0
            End If
    If Me.TxtModFlg.Text <> "R" Then

        If Me.CboPaymentType1.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·›« Ê—… ...!!!"
            Else
                Msg = "Select Bill Type ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboPayMentType.SetFocus
            Exit Sub
        End If
    
        If Me.CboPayMentType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
            Else
                Msg = "Select Payment method ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboPayMentType.SetFocus
            Exit Sub
        End If
    
        If Me.CboPayMentType.ListIndex = 2 Then
            If Trim(Me.DCVendor.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·„Ê—œ..!!"
                Else
                    Msg = "Select vendor..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DCVendor.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
        
        End If
    
        If Me.CboPayMentType.ListIndex = 0 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
                Else
                    Msg = "Select Box..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBox.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 3 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBankName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.Text) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                Else
                    Msg = "Enter Cheque No:...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If

            '     If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
            '                 If SystemOptions.UserInterface = ArabicInterface Then
            '                     Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            '                 Else
            '                 Msg = "Cheque Due Date Not Valid...!!"
            '
            '                 End If
            '         MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '         DtpChequeDueDate.SetFocus
            '         SendKeys "{F4}"
            '         Exit Sub
            '     End If
        End If
      Dim ISVAT As Boolean
    ISVAT = False
With Fg_Journal
    For i = .FixedRows To .Rows - 1
      If val(.TextMatrix(i, .ColIndex("Vat"))) > 0 Then
      ISVAT = True
      End If
     Next i
 End With
 With VSFlexGrid1
    For i = .FixedRows To .Rows - 1
      If val(.TextMatrix(i, .ColIndex("Vat"))) > 0 Then
      ISVAT = True
      End If
     Next i
 End With
 
Dim AccountVATDept As String
If ISVAT = True And True = True Then
If GetValueAddedAccount(XPDtbTrans.value, AccountVATDept) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…"
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If
        

        If CboPaymentType1.ListIndex = 0 Then

            With Fg_Journal

                For i = .FixedRows To .Rows - 2

                    If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
                        '////////////////////////////////////////notes
               
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·«Ì ÌÊÃœ „’—Ê› ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                        Else
                            MsgBox "Select Expenses in line no" & i, vbCritical
                        End If

                        Exit Sub
              
                    End If
        
                Next i

            End With

            With Fg_Journal

                For i = .FixedRows To .Rows - 2

                    If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Or val(.TextMatrix(i, .ColIndex("value"))) <= 0 Then
                        '////////////////////////////////////////notes
               
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·«Ì ÌÊÃœ ﬁÌ„… ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                        Else
                            MsgBox "Enter Value in line no" & i, vbCritical
                        End If
               
                        Exit Sub
                    End If
        
                Next i

            End With

        End If

        'Õ”«»« 
        If Me.CboPaymentType1.ListIndex = 1 Then
      
            With Me.VSFlexGrid1

                For i = .FixedRows To .Rows - 2

                    If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
                        '////////////////////////////////////////notes
               
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·«Ì ÌÊÃœ Õ”«» ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                        Else
                            MsgBox "Select Expenses in line no" & i, vbCritical
                        End If

                        Exit Sub
              
                    End If
        
                Next i

            End With
   
            With Me.VSFlexGrid1

                For i = .FixedRows To .Rows - 2

                    If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Or val(.TextMatrix(i, .ColIndex("value"))) <= 0 Then
                        '////////////////////////////////////////notes
               
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·«Ì ÌÊÃœ ﬁÌ„… ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                        Else
                            MsgBox "Enter Value in line no" & i, vbCritical
                        End If
               
                        Exit Sub
                    End If
        
                Next i

            End With
 
        End If
    
        If Me.TxtModFlg.Text = "N" Then
            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.Text), XPDtbTrans.value) = False Then
                        Exit Sub
                    End If
                End If
            End If

        ElseIf Me.TxtModFlg.Text = "E" Then

            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.Text), XPDtbTrans.value, , , val(Me.XPTxtID.Text)) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If
    
        Dim xrow As Integer

        With Fg_Journal

            For xrow = .Rows - 1 To 2 Step -1

                If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

                    .Rows = .Rows - 1
                End If

            Next xrow

        End With
    
        With Me.VSFlexGrid1

            For xrow = .Rows - 1 To 2 Step -1

                If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

                    .Rows = .Rows - 1
                End If

            Next xrow

        End With

        calcnets

        '-------------------------------------------------------------------------------------------
 my_branch = val(Me.Dcbranch.BoundText)
        '-------------------------------------------------------------------------------------------
        If TxtSerial.Text = "" Then
            If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                Else
                    MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number ": Exit Sub
                End If

            Else
         
                If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                        MsgBox "You must Define JE Coding ": Exit Sub
                    End If

                Else
                    TxtSerial.Text = Notes_coding(val(my_branch), XPDtbTrans.value)
                End If
            End If
        End If
 
        If TxtSerial1.Text = "" Then
            If Voucher_coding(val(my_branch), XPDtbTrans.value, 8, 80) = "error" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                    MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                End If

            Else
         
                If Voucher_coding(val(my_branch), XPDtbTrans.value, 8, 80) = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                        MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                    End If

                Else
                    TxtSerial1.Text = Voucher_coding(val(my_branch), XPDtbTrans.value, 8, 80)
                End If
            End If
        End If
    
        Cn.BeginTrans
        BeginTrans = True
        On Error GoTo ErrTrap
        '///////////////NOTESALL
        Dim A_NoteID As Long

        If TxtModFlg.Text = "N" Then
            XPTxtID.Text = CStr(new_id("notes_all", "NoteID", "", True))
            Me.TxtNoteSerial.Text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=80"))
            rs.AddNew
   
            Me.oldTxtSerial1.Text = Trim$(Me.TxtSerial1.Text)
 
        ElseIf Me.TxtModFlg.Text = "E" Then
    
          '  StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where notes_all=" & val(XPTxtID.Text)
          '  Cn.Execute StrSQL, , adExecuteNoRecords
        
            StrSQL = "Delete From notes Where  notetype=80  and  notes_all=" & val(XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
       
            If DcCostCenter.BoundText <> "" Then
                StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
        
        
      StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords


        End If
    
        '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("branch_no").value = val(Me.Dcbranch.BoundText)
        rs("txtManulaVat").value = val(txtManulaVat.Text)
        rs("NoteID").value = val(XPTxtID.Text)
        rs("bill_Type").value = Me.CboPaymentType1.ListIndex
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        rs("foxy_no").value = val(Text1.Text)
        rs("order_no").value = Txt_order_no.Text
        rs("VATCustoms").value = val(TxtVATCustoms.Text)
        rs("OrderID").value = IIf(Me.txtOrderID.Text = "", Null, Trim(txtOrderID.Text))
        rs("Noteseril2").value = IIf(Me.TxtNoteSerial1.Text = "", "", Trim(TxtNoteSerial1.Text))
        rs("Note_Value").value = IIf(XPTxtVal.Text = "", Null, XPTxtVal.Text)
        rs("Remark").value = IIf(XPMTxtRemarks.Text = "", "", Trim(XPMTxtRemarks.Text))
        rs("too").value = IIf(txtTo.Text = "", "", Trim(txtTo.Text))
        rs("general_des").value = IIf(txt_general_des.Text = "", "", Trim(txt_general_des.Text))
      
      rs("rate").value = IIf(Not IsNumeric(txt_Currency_rate.Text), 1, txt_Currency_rate.Text)
        If CBoBasedON.ListIndex > -1 Then
            rs("BasedONID").value = CBoBasedON.ListIndex
        Else
            rs("BasedONID").value = 0
        End If
        rs("CurrncyID").value = IIf(DcCurrency.Text = "", MainCurrency(), val(DcCurrency.BoundText))
        rs("CusID").value = Null
        rs("NoteType").value = 80
        rs("NoteDate").value = XPDtbTrans.value
        rs("UserID").value = user_id
        rs("ExpensesID").value = IIf(XPCboExpensesType.Text = "", Null, XPCboExpensesType.BoundText)
  
        If Me.CboPayMentType.ListIndex = 0 Then
            rs("BoxID").value = val(DcboBox.BoundText)
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("NoteCashingType").value = 0
        ElseIf Me.CboPayMentType.ListIndex = 1 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 1
    
        ElseIf Me.CboPayMentType.ListIndex = 3 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 3
        
        ElseIf Me.CboPayMentType.ListIndex = 2 Then
            rs("NoteCashingType").value = 2
            rs("CusID").value = val(Me.DCVendor.BoundText)
        End If
    
        rs("project_Expensen_account").value = IIf(Me.DCproject.BoundText = "", "", Me.DCproject.BoundText)
        rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.Text) = "", Null, Trim$(Me.Txt_Numorder.Text))
        rs("Buy").value = "0"
        rs("Remark").value = IIf(txt_general_des.Text = "", "", Trim(txt_general_des.Text))
        rs("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
        rs("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
 
        rs("OldNoteSerial1").value = Trim$(Me.oldTxtSerial1.Text) '
     
        rs("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
        rs("numbering_type1").value = sand_numbering_type(8) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
     
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)

        If DCproject.BoundText <> "" Then
          '  rs("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
        Else
         '   rs("note_value_by_characters").value = WriteNo(Format(val(Me.XPTxtVal.text), "0.00"), 0, True, ".", , 0)
        End If

        If Me.TxtModFlg.Text = "N" Then
            A_NoteID = CStr(new_id("Notes", "NoteID", "", True))
            TXT_A_NoteID.Text = A_NoteID
        Else
            A_NoteID = val(TXT_A_NoteID.Text)
        End If
    
        rs("A_NoteID").value = val(A_NoteID)
     
        rs.update
    
        '/////////////////////Õ”«»«  ⁄«„Â
        Dim line_no  As Integer

        If Me.CboPaymentType1.ListIndex = 1 Then
      
            Set RsNotes = New ADODB.Recordset
           ' RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
            If TxtModFlg.Text = "N" Then
           
            ElseIf Me.TxtModFlg.Text = "E" Then
           '     StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.Text)
           '     Cn.Execute StrSQL, , adExecuteNoRecords
        
            End If
    
            '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
            ' rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
            ' rs("foxy_no").value = Val(Text1.text)

            'Õ”«»« 
            RsNotes.AddNew
            RsNotes("NoteID").value = A_NoteID
             RsNotes.update
            RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
            RsNotes("order_no").value = Txt_order_no.Text
            RsNotes("notes_all").value = Me.XPTxtID.Text
            RsNotes("Note_Value").value = IIf(Not IsNumeric(XPTxtVal.Text), 0, val(XPTxtVal.Text))
            'RsNotes("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
            RsNotes("Remark").value = IIf(txt_general_des.Text = "", "", Trim(txt_general_des.Text))
            RsNotes("too").value = IIf(txtTo.Text = "", "", Trim(txtTo.Text))
             RsNotes("ManualNo").value = IIf(txtTo.Text = "", "", Trim(txtTo.Text))
             
            
            '    RsNotes("general_des").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    
            If Me.CboPayMentType.ListIndex = 0 Then
                RsNotes("BoxID").value = val(DcboBox.BoundText)
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
                RsNotes("NoteCashingType").value = 0
            ElseIf Me.CboPayMentType.ListIndex = 1 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 1
        
            ElseIf Me.CboPayMentType.ListIndex = 3 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 3
           
            ElseIf Me.CboPayMentType.ListIndex = 2 Then
                RsNotes("CusID").value = val(DCVendor.BoundText)
            End If
    
            RsNotes("NoteType").value = 80
            RsNotes("NoteDate").value = XPDtbTrans.value
            RsNotes("UserID").value = user_id
    
            'rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
            'rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
            RsNotes("Buy").value = "0"
            ' RsNotes("Remark").value = XPMTxtRemarks.text
            RsNotes("Remark").value = IIf(txt_general_des.Text = "", "", Trim(txt_general_des.Text))
            RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
            RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
            RsNotes("numbering_type1").value = sand_numbering_type(8) '‰Ê⁄  —ﬁÌ„   ›« Ê—… „«·Ì…
     
            RsNotes("sanad_year").value = year(XPDtbTrans.value)
            RsNotes("sanad_month").value = Month(XPDtbTrans.value)
            RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
            RsNotes.update
 
            '„œÌ‰ Õ”«»« 
            With VSFlexGrid1
                line_no = 1
 
                For i = .FixedRows To .Rows - 1
    
                    Dim project_id As Integer
    
                    If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
       
                        project_id = get_project_id(DCproject.BoundText, "expanses_account")
   
                        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                    OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                    OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                    OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                    OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                    OtherInformation.PriceTotal = val(.TextMatrix(i, .ColIndex("PriceTotal")))
                    OtherInformation.SupplierID = val(.TextMatrix(i, .ColIndex("SupplierID")))
                    OtherInformation.CusVATNO = (.TextMatrix(i, .ColIndex("CusVATNO")))
                    OtherInformation.SupplierName = (.TextMatrix(i, .ColIndex("SupplierName")))
                    If val(txt_Currency_rate) <> 0 Then
                    .TextMatrix(i, .ColIndex("Rate")) = txt_Currency_rate
                    End If
                    OtherInformation.Rate = val(.TextMatrix(i, .ColIndex("Rate")))
                  project_id = val(.TextMatrix(i, .ColIndex("ProjectId")))
                        If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("Value")) * val(txt_Currency_rate.Text), 0, .TextMatrix(i, .ColIndex("Des")), A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , DcCurrency.Text, val(txt_Currency_rate.Text), , , val(.TextMatrix(i, .ColIndex("LineNo1"))), val(Me.XPTxtID.Text), project_id, , , , , , , val(Me.Dcbranch.BoundText), , , , , , , val(.TextMatrix(i, .ColIndex("DeparmentID"))), , , , , , .TextMatrix(i, .ColIndex("BillNo")), , val(.TextMatrix(i, .ColIndex("pandid"))), val(.TextMatrix(i, .ColIndex("operid"))), , , , , , , , , Posted, , OtherInformation) = False Then
                        
                        
                            GoTo ErrTrap
                    
                        End If

                        line_no = line_no + 1
            
                    End If

                Next i

            End With

            'œ«∆‰ Õ”«»« 
    
            Dim IntDEV_Type As Integer
            Dim SngDEV_Value As Single
            If val(TxtVATCustoms.Text) > 0 Then
            line_no = line_no + 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                    OtherInformation.FlgVat = 0
                    OtherInformation.Vat = 0
                    OtherInformation.Vatyo = 0
                    OtherInformation.CurrRow = 0
                    OtherInformation.PriceTotal = 0
                    OtherInformation.SupplierID = 0
                    OtherInformation.CusVATNO = ""
                    OtherInformation.SupplierName = ""
                    OtherInformation.Rate = val(txt_Currency_rate.Text)
             StrAccount = get_account_code_branch(148, my_branch)
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccount, IIf(Not IsNumeric(TxtVATCustoms.Text), 0, val(TxtVATCustoms.Text)), 0, txt_general_des.Text & "Õ”«» ﬁÌ„… «· VAT ··Ã„«—ﬂ ›Ì «·›« Ê—… «·„«·Ì…", A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , DcCurrency.Text, val(txt_Currency_rate.Text), , , , val(Me.XPTxtID.Text), , , , , , , , val(Me.Dcbranch.BoundText), , , , , , , , , , , , , , , , , , 1, , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                    
            End If
            End If
            line_no = line_no + 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                    OtherInformation.FlgVat = 0
                    OtherInformation.Vat = 0
                    OtherInformation.Vatyo = 0
                    OtherInformation.CurrRow = 0
                    OtherInformation.PriceTotal = 0
                    OtherInformation.SupplierID = 0
                    OtherInformation.CusVATNO = ""
                    OtherInformation.SupplierName = ""
                    OtherInformation.Rate = val(txt_Currency_rate.Text)
            If ModAccounts.AddNewDev(LngDevID, line_no, DcboCreditSide.BoundText, IIf(Not IsNumeric(XPTxtVal.Text), 0, val(XPTxtVal.Text)) + val(TxtVATCustoms.Text), 1, txt_general_des.Text, A_NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , DcCurrency.Text, val(txt_Currency_rate.Text), , , , val(Me.XPTxtID.Text), , , , , , , , val(Me.Dcbranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                    
            End If
        
            ' TxtModFlg.text = "R"
            GoTo ll
      
        End If
    
        '  «·„’—Ê›«  „œÌ‰
    
        '//////////////////////////////////////Notes////////////////////////////////////
        Set RsNotes = New ADODB.Recordset
     '   RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
       
            Set RsDev = New ADODB.Recordset
             StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.* FROM         dbo.DOUBLE_ENTREY_VOUCHERS WHERE     (Double_Entry_Vouchers_ID = - 1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

          '  RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            
            '«·ÿ—› «·„œÌ‰
 
            Dim ExpensesID As Double
 
            Dim NoteID As String

            With Fg_Journal

                line_no = 1
       
                project_id = get_project_id(DCproject.BoundText, "expanses_account")
                
                For i = .FixedRows To .Rows - 1
                    'project_id = get_project_id(dcproject.BoundText, "expanses_account")
   
                    If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                        '////////////////////////////////////////notes
                
                        If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
                            Else
                                MsgBox "Cant save no value in line no:  " & i - 1, vbCritical: GoTo ErrTrap
                            End If
               
                        End If

                        RsNotes.AddNew
                        NoteID = CStr(new_id("Notes", "NoteID", "", True))
                        RsNotes("NoteID").value = CStr(NoteID)
                         
                        RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
                        RsNotes("Note_Value").value = .TextMatrix(i, .ColIndex("value"))
                        RsNotes("ExpensesRemark").value = .TextMatrix(i, .ColIndex("des"))
                        RsNotes("Project_Id").value = val(.TextMatrix(i, .ColIndex("ProjectId")))
                        RsNotes("Pand").value = val(.TextMatrix(i, .ColIndex("pandid")))
                        
                        
                        '  RsNotes("Remark").value = .TextMatrix(I, .ColIndex("des"))
                
                        RsNotes("Remark").value = IIf(txt_general_des.Text = "", "", Trim(txt_general_des.Text))

                        RsNotes("foxy_no").value = val(Text1.Text)
                        '  If Me.CboPayMentType.ListIndex = 0 Then
                        '     rsnotes("BoxID").value = Val(DcboBox.BoundText)
                        '     rsnotes("BankID").value = Null
                        '     rsnotes("ChqueNum").value = Null
                        '     rsnotes("DueDate").value = Null
                        '     rsnotes("NoteCashingType").value = 0
                        ' ElseIf Me.CboPayMentType.ListIndex = 1 Then
                        '     rsnotes("BoxID").value = Null
                        '     rsnotes("BankID").value = Val(Me.DCboBankName.BoundText)
                        '     rsnotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                        '     rsnotes("DueDate").value = Me.DtpChequeDueDate.value
                        '     rsnotes("NoteCashingType").value = 1
                        ' End If
               
                        If Txt_order_no.Text <> "" Then
                            RsNotes("order_no").value = Txt_order_no.Text
                        Else
                            RsNotes("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
                        End If
            
                        RsNotes("CusID").value = Null
                        RsNotes("NoteType").value = 80
                        RsNotes("NoteDate").value = XPDtbTrans.value
                        RsNotes("UserID").value = user_id
                        RsNotes("ExpensesID").value = .TextMatrix(i, .ColIndex("ExpensesID"))
                        RsNotes("DeparmentID").value = val(.TextMatrix(i, .ColIndex("DeparmentID")))
                        RsNotes("ProjectID").value = val(.TextMatrix(i, .ColIndex("ProjectID")))
                        RsNotes("notes_all").value = Me.XPTxtID.Text
                        RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
                        RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
                        RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                        RsNotes("numbering_type1").value = sand_numbering_type(8) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
                        RsNotes("ManualNo").value = IIf(txtTo.Text = "", "", Trim(txtTo.Text))
                        RsNotes("sanad_year").value = year(XPDtbTrans.value)
                        RsNotes("sanad_month").value = Month(XPDtbTrans.value)
                        RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(.TextMatrix(I, .ColIndex("value")), "0.00"), 0, True, ".")
                
                        RsNotes.update
              
                        '////////////////////////////////////////notes
 
                        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                        
                    OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                    OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                    OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                    OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                    OtherInformation.PriceTotal = val(.TextMatrix(i, .ColIndex("PriceTotal")))
                    OtherInformation.SupplierID = val(.TextMatrix(i, .ColIndex("SupplierID")))
                    OtherInformation.CusVATNO = (.TextMatrix(i, .ColIndex("CusVATNO")))
                    OtherInformation.SupplierName = (.TextMatrix(i, .ColIndex("SupplierName")))
                    OtherInformation.Rate = val(.TextMatrix(i, .ColIndex("Rate")))
                     project_id = val(.TextMatrix(i, .ColIndex("ProjectId")))
                        If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), val(.TextMatrix(i, .ColIndex("value"))), 0, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , val(.TextMatrix(i, .ColIndex("value"))), DcCurrency.Text, val(txt_Currency_rate.Text), , , val(.TextMatrix(i, Fg_Journal.ColIndex("LineNo1"))), val(Me.XPTxtID.Text), project_id, .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , , , , val(Me.Dcbranch.BoundText), val(.TextMatrix(i, .ColIndex("CarId"))), , , , , , val(.TextMatrix(i, .ColIndex("DeparmentID"))), , , , , , .TextMatrix(i, .ColIndex("BillNo")), , val(.TextMatrix(i, .ColIndex("pandid"))), val(.TextMatrix(i, .ColIndex("operid"))), , , , , , , , , Posted, , OtherInformation) = False Then
                         
                       
                        
                            '   GoTo ErrTrap
                    
                        End If

                        line_no = line_no + 1
        
                    End If

                Next i

            End With
                If val(TxtVATCustoms.Text) > 0 Then
            line_no = line_no + 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                    OtherInformation.FlgVat = 0
                    OtherInformation.Vat = 0
                    OtherInformation.Vatyo = 0
                    OtherInformation.CurrRow = 0
                    OtherInformation.PriceTotal = 0
                    OtherInformation.SupplierID = 0
                    OtherInformation.CusVATNO = ""
                    OtherInformation.SupplierName = ""
                    OtherInformation.Rate = val(txt_Currency_rate.Text)
             StrAccount = get_account_code_branch(148, my_branch)
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccount, IIf(Not IsNumeric(TxtVATCustoms.Text), 0, val(TxtVATCustoms.Text)), 0, txt_general_des.Text & "Õ”«» ﬁÌ„… «· VAT ··Ã„«—ﬂ ›Ì «·›« Ê—… «·„«·Ì…", val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , DcCurrency.Text, val(txt_Currency_rate.Text), , , , val(Me.XPTxtID.Text), , , , , , , , val(Me.Dcbranch.BoundText), , , , , , , , , , , , , , project_id, val(Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("pandid"))), val(Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("operid"))), , 1, , , , , , , Posted, , OtherInformation) = False Then
             
                GoTo ErrTrap
                    
            End If
            End If
            
            ' «·„’—Ê›«  «·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
            RsNotes.AddNew
            NoteID = CStr(new_id("Notes", "NoteID", "", True))
            RsNotes("NoteID").value = CStr(NoteID)
             
            RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
 
            RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0)
            RsNotes("Remark").value = Me.txt_general_des
            RsNotes("foxy_no").value = val(Text1.Text)
RsNotes("ManualNo").value = IIf(txtTo.Text = "", "", Trim(txtTo.Text))
            If Me.CboPayMentType.ListIndex = 0 Then
                RsNotes("BoxID").value = val(DcboBox.BoundText)
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
                RsNotes("NoteCashingType").value = 0
            ElseIf Me.CboPayMentType.ListIndex = 1 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 1
            ElseIf Me.CboPayMentType.ListIndex = 3 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 3
                            
            ElseIf Me.CboPayMentType.ListIndex = 2 Then
                RsNotes("CusID").value = val(DCVendor.BoundText)
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = Null
    
            End If

            ' RsNotes("order_no").value = txt_ORDER_NO.text
            '              RsNotes("CusID").value = Null
            RsNotes("NoteType").value = 80
            RsNotes("NoteDate").value = XPDtbTrans.value
            RsNotes("UserID").value = user_id
            ' rsnotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
            RsNotes("notes_all").value = Me.XPTxtID.Text
            RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
            RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
            RsNotes("numbering_type1").value = sand_numbering_type(8) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
            RsNotes("sanad_year").value = year(XPDtbTrans.value)
            RsNotes("sanad_month").value = Month(XPDtbTrans.value)
            RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
                
            RsNotes.update
    
            '«·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("DEV_ID_Line_No").value = line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = DcboCreditSide.BoundText
            RsDev("Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0) + val(TxtVATCustoms.Text) '.TextMatrix(I, .ColIndex("VALUE"))
           If val(txt_Currency_rate.Text) > 0 Then
            RsDev("Valuee").value = RsDev("Value").value / val(txt_Currency_rate.Text)
            RsDev("Currency").value = DcCurrency.Text
            End If
            
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.Text  ' .TextMatrix(I, .ColIndex("des"))
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            If Posted = 1 Then
            RsDev("Posted").value = 1
            Else
            RsDev("Posted").value = Null
            End If
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.Text
            '   RsDev("project_id").value = project_id
                        
            RsDev.update
     
            'GoTo ll
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
 
            line_no = line_no + 1

            If Me.DCproject.BoundText <> "" Then
                '«·ÿ—› «·„œÌ‰   „’—Ê›«  «·„‘—Ê⁄
                RsNotes.AddNew
                NoteID = CStr(new_id("Notes", "NoteID", "", True))
                RsNotes("NoteID").value = CStr(NoteID)
                 
                RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
          
                RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0)
                RsNotes("Remark").value = txt_general_des.Text 'txtto.text

                If Me.CboPayMentType.ListIndex = 0 Then
                    RsNotes("BoxID").value = val(DcboBox.BoundText)
                    RsNotes("BankID").value = Null
                    RsNotes("ChqueNum").value = Null
                    RsNotes("DueDate").value = Null
                    RsNotes("NoteCashingType").value = 0
                ElseIf Me.CboPayMentType.ListIndex = 1 Then
                    RsNotes("BoxID").value = Null
                    RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                    RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                    RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                    RsNotes("NoteCashingType").value = 1
                     
                ElseIf Me.CboPayMentType.ListIndex = 3 Then
                    RsNotes("BoxID").value = Null
                    RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                    RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
                    RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                    RsNotes("NoteCashingType").value = 3
                            
                ElseIf Me.CboPayMentType.ListIndex = 2 Then
                    RsNotes("CusID").value = val(DCVendor.BoundText)
                    RsNotes("BoxID").value = Null
                    RsNotes("BankID").value = Null
                        
                End If
                        
                ' RsNotes("order_no").value = txt_ORDER_NO.text
                'RsNotes("CusID").value = Null
                RsNotes("NoteType").value = 80
                RsNotes("NoteDate").value = XPDtbTrans.value
                RsNotes("UserID").value = user_id
                ' rsnotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
                RsNotes("notes_all").value = Me.XPTxtID.Text
                RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.Text) '„”·”· «·ﬁÌœ
                RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.Text) '„”·”· «–‰ «·’—›
                RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                RsNotes("numbering_type1").value = sand_numbering_type(8) '‰Ê⁄  —ﬁÌ„  ›« Ê—… „«·Ì…
                RsNotes("sanad_year").value = year(XPDtbTrans.value)
                RsNotes("sanad_month").value = Month(XPDtbTrans.value)
                
                RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
                
                RsNotes.update
                
                project_id = get_project_id(DCproject.BoundText, "expanses_account")
                Set RsDev = New ADODB.Recordset
                
              '  RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                    StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
                RsDev.AddNew
                RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
                RsDev("DEV_ID_Line_No").value = line_no
                RsDev("DEV_ID_Line_No1").value = setfoxy_Line
                RsDev("Account_Code").value = DCproject.BoundText
                RsDev("Value").value = IIf(IsNumeric(XPTxtVal.Text), XPTxtVal.Text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.Text ' .TextMatrix(I, .ColIndex("des"))
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)5
                If Posted = 1 Then
                RsDev("Posted").value = 1
                Else
                RsDev("Posted").value = Null
                End If
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("notes_all").value = Me.XPTxtID.Text
               ' RsDev("project_id").value = project_id
                        
                RsDev.update
                    
                line_no = line_no + 1

                With Fg_Journal

                    For i = .FixedRows To .Rows - 1
        
                        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                            '////////////////////////////////////////notes
                
                            If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
                                Else
                                    MsgBox "Cant save enter value in line :  " & i - 1, vbCritical: GoTo ErrTrap
                                End If
               
                            End If

                          '  project_id = get_project_id(dcproject.BoundText, "expanses_account")
 project_id = 0
                            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                    OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                    OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                    OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                    OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                    OtherInformation.PriceTotal = val(.TextMatrix(i, .ColIndex("PriceTotal")))
                    OtherInformation.SupplierID = val(.TextMatrix(i, .ColIndex("SupplierID")))
                    OtherInformation.CusVATNO = (.TextMatrix(i, .ColIndex("CusVATNO")))
                    OtherInformation.SupplierName = (.TextMatrix(i, .ColIndex("SupplierName")))
                    OtherInformation.Rate = val(.TextMatrix(i, .ColIndex("Rate")))

                            If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("value")), 1, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), DcCurrency.Text, val(txt_Currency_rate.Text), , , setfoxy_Line, val(Me.XPTxtID.Text), project_id, , , , , , , val(Me.Dcbranch.BoundText), , , , , , , val(.TextMatrix(i, .ColIndex("DeparmentID"))), , , , , , .TextMatrix(i, .ColIndex("BillNo")), , , , , , , , , , , , Posted, , OtherInformation) = False Then
                                GoTo ErrTrap
                    
                            End If

                            line_no = line_no + 1
        
                        End If

                    Next i

                End With

               Dim sql As String
                sql = "Update notes    set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.Text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.Text) & " and notetype=80" & "and NoteSerial1=" & TxtSerial1
                Cn.Execute sql
                sql = "Update   notes_all  set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.Text) + val(TxtVATCustoms.Text), "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.Text) & " and notetype=80" & "and NoteSerial1=" & TxtSerial1
                Cn.Execute sql
            End If

            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            LblDevID.Caption = LngDevID
            lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
        End If

ll:
   If DCproject.BoundText = "" Then
                sql = "Update notes    set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.Text) + val(TxtVATCustoms.Text), "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.Text) & " and notetype=80" & "and NoteSerial1=" & TxtSerial1
                Cn.Execute sql
    End If
    
            '«· Ê“Ì⁄ ⁄·Ï „—ﬂ“ «· ﬂ·›… «·⁄«„
   
        '     If Me.DcCostCenter.BoundText <> "" Then
        save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.Text, "  ›« Ê—… „«·Ì…", Me.XPDtbTrans.value
        '     End If
        save_cost_center
        'Õ›Ÿ «·„’«—Ì› › ÃœÊ· «·„’«—Ì›
     
        If saveExpensesDetails(1, TxtSerial.Text, TxtSerial1.Text, Txt_order_no.Text, XPDtbTrans.value) = True Then
        End If
    
        'Õ›Ÿ »Ì«‰«  «·‘Ìﬂ« 
        saveChequeBoxContents1 (val(Me.TXT_A_NoteID.Text))

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata
    
        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved... " & CHR(13)
                    Msg = Msg + "Do you want to enter another operation?"
        
                End If

                Fg_Journal.Enabled = False

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
                End If

                lbl(27).Caption = showLabel(TxtSerial1, oldTxtSerial1)
        
                Fg_Journal.Enabled = False
        End Select

    
        TxtModFlg.Text = "R"
        fillapprovData
    End If

    WriteCustomerBalPublic Me.DcboCreditSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
   
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "cant save " & CHR(13)
            Msg = Msg + "Invalid entry value " & CHR(13)
            Msg = Msg + "Check data and try again"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorr.... Error during saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Function save_cost_center()

    'on error resume next
    If Not IsNumeric(Text1.Text) Then Exit Function
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql_str As String

    'Rs.Open "", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    sql_str = "select * from marakes_taklefa_temp where kedno=" & Text1.Text
    rs.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs.RecordCount
        rs("ok").value = 1
        rs("NoteDate").value = XPDtbTrans.value
        rs("NoteSerial").value = TxtSerial.Text
        rs("Remark").value = "   ›« Ê—… „«·Ì… —ﬁ„ " & TxtSerial1 & "    " & Me.txt_general_des
 
        rs.update
        rs.MoveNext
    Next i

End Function

Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND  kedno =" & val(Text1.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If

    'rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
StrSQL = "SELECT   *  from dbo.marakes_taklefa_temp Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
If CboPaymentType1.ListIndex = 0 Then
    With Fg_Journal
 
        .Rows = .Rows + 1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                rs.AddNew
                rs("cost_center_id").value = cost_center_id
                rs("cost_center").value = cost_center
                rs("value").value = .TextMatrix(i, .ColIndex("value"))
                rs("depit_or_credit").value = "„œÌ‰"
                rs("opr_id").value = Me.Text1.Text
                rs("kedno").value = Me.Text1.Text
                rs("opr_type").value = opr_type
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
                rs("record_date").value = record_date
                rs("general_des").value = 1
                rs.update
        
            End If

        Next i

    End With
Else

    With VSFlexGrid1
 
        .Rows = .Rows + 1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                rs.AddNew
                rs("cost_center_id").value = cost_center_id
                rs("cost_center").value = cost_center
                rs("value").value = .TextMatrix(i, .ColIndex("value"))
                rs("depit_or_credit").value = "„œÌ‰"
                rs("opr_id").value = Me.Text1.Text
                rs("kedno").value = Me.Text1.Text
                rs("opr_type").value = opr_type
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
                rs("record_date").value = record_date
                rs("general_des").value = 1
                rs.update
        
            End If

        Next i

    End With



End If

    rs.Close
End Function

Private Sub Undo()
    Dim sgl As String
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            sgl = "delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute sgl, , adExecuteNoRecords
        
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)
         
        Case "E"
            sgl = "delete  marakes_taklefa_temp  where ok is null and  kedno =" & val(Text1.Text)
            Cn.Execute sgl, , adExecuteNoRecords
        
            rs.Find "NoteID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If SystemOptions.banks_Accounts3 = True Then
        If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
        
      If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–Â «·⁄„·Ì…"
            Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
        Else
        Msg = "Cant Delete This Transaction"
            Msg = Msg & CHR(13) & " Cheque Alraedy Payed "
        End If
            
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
        
        
    End If
        If CheAssetPayd(val(Me.XPTxtID)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–Â «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ≈÷«›… ··«’Ê·   "
                    Else
                    Msg = " Can Not Delete this Process"
                    Msg = Msg & CHR(13) & " There is the Process of adding Assest "
                    
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
    If XPTxtID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
 Else
     Msg = "Confirm Delete trans#" & CHR(13)
        Msg = Msg + (TxtNoteSerial.Text) & CHR(13)
        Msg = Msg + " ø"
 End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
    Deletepost Me.Name, "notes_all", "NoteID", 0, val(Dcbranch.BoundText), val(XPTxtID.Text), TxtSerial1.Text
    
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
 
          '  StrSQL = "Delete From notes Where NoteID=" & val(TXT_A_NoteID.text)
             StrSQL = "Delete From notes Where notetype=80  and   notes_all=" & val(XPTxtID.Text)
             
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From ExpensesDetails Where NoteSerial1='" & val(TxtSerial1.Text) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
    
            StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
    
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
       
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Fg_Journal.Clear flexClearScrollable, flexClearEverything
                    Fg_Journal.Rows = 3
                    Fg_Journal.Enabled = False
                
                    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                    VSFlexGrid1.Rows = 2
                    VSFlexGrid1.Enabled = False
                
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This is Process UnAvailable"
        End If
        
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry...error douring delete"
  End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Function FillGridWithData()

End Function

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

    IntCounter = 0

    With Me.VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

End Sub

Private Sub PutData()

    'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)
    With Fg_Journal

        If Len(TxtDes.Text) > 0 Then
            .Cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.Text
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        Else
            .Cell(flexcpData, .Row, .ColIndex("Des")) = ""
            .Cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
            .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        End If

    End With

End Sub

Function sand_numbering() As String
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    Dim auto_sanad_no As String
    Dim NO As String
    auto_sanad_no = ""
    departement_name = 1
    branch_no = 1
    connection_string = Cn.ConnectionString
    numbering.ConnectionString = connection_string
    numbering.CommandType = adCmdText
    numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=1"
    numbering.Refresh

    If numbering.Recordset.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = numbering.Recordset.Fields!numbering_id
        start_at = numbering.Recordset.Fields!start_at
        end_at = numbering.Recordset.Fields!end_at

    End If

    If numbering_type = 1 Then
        detect_no.ConnectionString = connection_string
        detect_no.CommandType = adCmdText
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
        detect_no.Refresh

        If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
 
            If end_at = 0 Then end_at = detect_no.Recordset.Fields!last_sand_no + 1
 
            If detect_no.Recordset.Fields!last_sand_no >= end_at Then
                sand_numbering = "error"
                Exit Function
            End If
        End If

    Else

        If numbering_type = 2 Then
 
            detect_no.ConnectionString = connection_string
            detect_no.CommandType = adCmdText
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & mId(Format$(Now, "dd/mm/yyyy"), 4, 2)
            'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

            If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)

                If end_at = 0 Then end_at = NO + 1
                If NO >= end_at Then
                    sand_numbering = "error"
                    Exit Function
                End If
            End If

        Else

            If numbering_type = 3 Then
 
                detect_no.ConnectionString = connection_string
                detect_no.CommandType = adCmdText
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4)
                'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                detect_no.Refresh

                If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)

                    If end_at = 0 Then end_at = NO + 1
                    If NO >= end_at Then
                        sand_numbering = "error"
                        Exit Function
                    End If
                End If
 
            End If
 
        End If
    End If

    If detect_no.Recordset.RecordCount = 0 Or IsNull(detect_no.Recordset.Fields!last_sand_no) Then

        If numbering_type = 0 Then
            ' auto_sanad_no = 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = start_at
            Else
                
                If numbering_type = 2 Then
                    auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & mId(Format$(Now, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & start_at

                    End If
                End If
            End If
        End If

    Else

        If numbering_type = 0 Then
            'auto_sanad_no = x + 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = detect_no.Recordset.Fields!last_sand_no + 1
            Else
                
                If numbering_type = 2 Then
                    '  If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) Then
                    ' no = 1
                    '  auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & "1"
                    '  Else
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (NO + 1)
                    '  End If
                      
                Else

                    If numbering_type = 3 Then
                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                        'no = 1
                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                        '    Else
                        NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                        auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (NO + 1)

                        '    End If

                    End If
                End If
            End If
        End If

    End If

    sand_numbering = auto_sanad_no

    'MsgBox auto_sanad_no

End Function

Function setfoxy_Line() As Double
    
    Dim X As Double
    X = CStr(new_id("foxy", "id1", "", True))
    setfoxy_Line = X
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id1").value = X ' last_line_id
 
    rs.update
    
End Function

Private Sub CBoBasedON_Change()
TxtVATCustoms.Visible = False
lbl(28).Visible = False
    'n
    With Me.Fg_Journal

        If Me.CBoBasedON.ListIndex = 0 Then

        ElseIf Me.CBoBasedON.ListIndex = 1 Then
          TxtVATCustoms.Visible = True
          lbl(28).Visible = True
            If SystemOptions.UserInterface = ArabicInterface Then
                lbl(21).Caption = "—ﬁ„ «·«„—"
            Else
                lbl(21).Caption = "  Order No"
            End If

        ElseIf Me.CBoBasedON.ListIndex = 2 Then
          TxtVATCustoms.Visible = True
          lbl(28).Visible = True
            If SystemOptions.UserInterface = ArabicInterface Then
                lbl(21).Caption = "—ﬁ„ «·›« Ê—… «·„»œ∆ÌÂ"
            Else
                lbl(21).Caption = "Performa Invoice NO"
            End If

        End If

        .TextMatrix(0, .ColIndex("order_no")) = lbl(21).Caption

    End With

End Sub

Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "—ﬁ„ «·›« Ê—… " & TxtSerial1.Text & CHR(13) & "   «· «—ÌŒ  " & XPDtbTrans & CHR(13) & "   «·›—⁄ " & Dcbranch & CHR(13) & "   „—ﬂ“ «· ﬂ·›… «·⁄«„  " & DcCostCenter & CHR(13) & "   ÿ—Ìﬁ… «·œ›⁄  " & CboPayMentType & CHR(13) & "   «·„‘—Ê⁄  " & DCproject & CHR(13) & "   «·„Ê—œ " & DCVendor & CHR(13) & "   «·Œ“Ì‰… " & DcboBox & CHR(13) & "   «·»‰ﬂ  " & DcboBankName & CHR(13) & "   —ﬁ„ «·‘Ìﬂ " & TxtChequeNumber & CHR(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & CHR(13) & "   —ﬁ„ ›« Ê—… «·„Ê—œ " & txtTo & CHR(13) & "   »‰«¡ ⁄·Ï  " & CBoBasedON & "  »—ﬁ„  " & Txt_order_no & CHR(13) & "   «·‘—Õ «·⁄«„  " & txt_general_des & CHR(13) & "   «Ã„«·Ì «·›« Ê—…    " & XPTxtValView
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Bill No " & TxtSerial1.Text & CHR(13) & "   Date  " & XPDtbTrans & CHR(13) & "   Branch " & Dcbranch & CHR(13) & "   CC  " & DcCostCenter & CHR(13) & "  Payment Type  " & CboPayMentType & CHR(13) & "   Project  " & DCproject & CHR(13) & "   Supplier " & DCVendor & CHR(13) & "   Box " & DcboBox & CHR(13) & "   Bank  " & DcboBankName & CHR(13) & "   Cheque No:   " & TxtChequeNumber & CHR(13) & "  Due Date  " & DtpChequeDueDate & CHR(13) & "  Supplier Bill No " & txtTo & CHR(13) & "   Based On  " & CBoBasedON & "  No:  " & Txt_order_no & CHR(13) & "  Remarks  " & txt_general_des & CHR(13) & "   Bill Total   " & XPTxtValView
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 80, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TxtSerial, TxtSerial1
    Else
        AddToLogFile CInt(user_id), 80, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtSerial, TxtSerial1
    End If
    
End Function

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

Private Sub AddTip()
    Dim Wrap As String
    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    If BolRtl = True Then

        With TTP
            .Create Me.hwnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        End With

        With TTP
            .Create Me.hwnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    Else

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "Add New Record..." & Wrap & "Shortcut Key F12 OR Enter" & Wrap & "OR Alt+N", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit the Current Record..." & Wrap & "Shortcut Key F11 " & Wrap & "OR Alt+E", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save the New Record OR Save the Editing in the Current Record..." & Wrap & "Shortcut Key F10 " & Wrap & "OR Alt+S", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Cancel the New Record OR Cancel Editing in the Current Record..." & Wrap & "Shortcut Key F9 " & Wrap & "OR Alt+U", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete the Current Record..." & Wrap & "Shortcut Key F8 " & Wrap & "OR Alt+D", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Close this Screen" & Wrap & "OR Alt+X", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "Display Help for this Screen" & Wrap & "Shortcut Key F1" & Wrap, BolRtl
        End With

    End If

    Exit Sub
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

Private Sub XPCboExpensesType_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("ExpensesType", "ID", val(Me.XPCboExpensesType.BoundText))
    End If

End Sub

Private Sub XPDtbTrans_Change()

    If Me.TxtModFlg = "E" Then
        If Month(rs("NoteDate").value) = Month(XPDtbTrans.value) Then Exit Sub
    End If

    If Trim(TxtSerial1.Text) <> "" Then
        oldTxtSerial1.Text = TxtSerial1.Text
    End If

    TxtSerial.Text = ""
    TxtSerial1.Text = ""

End Sub

Private Sub XPTxtVal_Change()
    'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0)
    XPTxtValView.Text = Format(val(XPTxtVal.Text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.Text, "0.00"), 0, True, ".", , 0)

    Else

        'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0, , , , 1)
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.Text, "0.00"), 0, True, ".", , 1)

    End If
    
End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
End Sub

Private Sub XPTxtVal_Validate(Cancel As Boolean)
    'If Val(XPTxtVal.Text) = 0 Then
    '    Set TTD = New clstooltipdemand
    '    TTD.Style = TTBalloon
    '    TTD.Icon = TTIconWarning
    '    TTD.Centered = True
    '    TTD.RightToLeft = True
    '    TTD.VisibleTime = 600
    '    TTD.BackColor = 0
    '    TTD.Title = "ﬁÌ„… «·„’—Ê›« "
    '    TTD.TipText = "»—Ã«¡ ﬂ «»… ﬁÌ„… «·„’—Ê›« "
    '    TTD.PopupOnDemand = True
    '    TTD.CreateToolTip XPTxtVal.hwnd
    '    TTD.Show 0, XPTxtVal.Height / Screen.TwipsPerPixelX - 1    '//In Pixel only
    '    Cancel = True
    'Else
    '    TTD.Destroy
    'End If
End Sub

Private Sub ViewDataList()
    Dim FrmView As FrmViewList
    Dim fg As VSFlex8UCtl.vsFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    'Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set fg = FrmView.vsfGroup1.vsFlexGrid

    With fg
        .Cols = 18
        .RowHeightMin = 320
        .ExplorerBar = flexExSortShowAndMove
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & "Remark, UserName From ExpensesReport"
        StrSQL = StrSQL + " Order By NoteID"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
        'Â‰« Ìﬂ » ﬂÊœ ·⁄„· „⁄œ·  Õ„Ì· «·»Ì«‰« 
        '------------------------------------
        '
        '
        '
        '
    
        '------------------------------------
        Set .DataSource = rs
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        'Rs.Close
        'Set Rs = Nothing
        .AutoSize 0, .Cols - 1, False
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.vsFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "Note_Value"
    FrmView.vsfGroup1.sql = StrSQL
    FrmView.vsfGroup1.ShowTreeGroups = True
    FrmView.vsfGroup1.update
    FrmView.SetDblClickRetrun Me, "NoteID"
    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·»Ì«‰«  «·„’—Ê›« "
    FrmView.show
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(28).Caption = "Customs Value"
    'LblValue.Visible = False
    lbl(24).Caption = "Hint"
    CmdAttach.Caption = "Attachments"
    Label10.Caption = "Based On"
    lbl(25).Caption = "This Window Allow To Refister Financial Invoice"
    lbl(23).Caption = "Invoice Type"
    Label3.Caption = "GL No."
    lbl(14).Caption = "Project#"
    'Label1.Caption = "Manual #"
    C1Tab1.Caption = "Data|Approve"
    Me.ALLButton1.Caption = "Cost Center"
    lbl(15).Caption = "Payment Method"
    lbl(16).Caption = "Box Name"
    lbl(20).Caption = "General Des"
    lbl(21).Caption = "Order No:"

    Label8.Caption = "General C. C."
    Label1.Caption = "Branch"
    lbl(26).Caption = "Based ON"

    With Me.CBoBasedON
        .Clear
        .AddItem "Without"
        .AddItem "Purchase Invoices"
        .AddItem "Performa Invoices"
        .AddItem "Production Order"
    End With

    With Me.CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
        .AddItem "Credit"
        .AddItem "P Cheque "
    End With

    With Me.CboPaymentType1
        .Clear
        .AddItem "Expenses"
        .AddItem "Accounts"
     
    End With

    CmdRemove.Caption = "Delete Row"
    Me.Caption = "Financial Invoice"
    Me.ELe.Caption = Me.Caption

    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    Me.lbl(4).Caption = " Vchr#"
    Me.lbl(1).Caption = " Date"
    Me.lbl(3).Caption = "Expenses Type"
    Me.lbl(2).Caption = "Total"
    Me.lbl(0).Caption = "Vendor Bill#"
    Me.lbl(5).Caption = "Remarks"
    Me.lbl(8).Caption = "Issued By."
    Me.lbl(7).Caption = "Current Record."

    Fra.Caption = "GL"
    lbl(11).Caption = "GL#"
    lbl(13).Caption = "Interval"
    lbl(9).Caption = "Depit"
    lbl(10).Caption = "Credit"
    lbl(17).Caption = "Bank"
    lbl(18).Caption = "Cheque#"
    lbl(19).Caption = "Due Date"
    lbl(22).Caption = "Vendor"

    Me.Cmd(0).Caption = "&New"
    Me.Cmd(1).Caption = "&Edit"
    Me.Cmd(2).Caption = "&Save"
    Me.Cmd(3).Caption = "&Undo"
    Me.Cmd(4).Caption = "&Delete"
    Me.Cmd(5).Caption = "Sear&ch"
    Me.Cmd(6).Caption = "E&xit"
    Me.Cmd(7).Caption = "&Table View"
    Cmd(8).Caption = "Print"
    Cmd(9).Caption = "Cheque Print"
    Cmd(10).Caption = "GL Print "

    Me.CmdHelp.Caption = "&Help"

    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = " Expenses Name"
        .TextMatrix(0, .ColIndex("value")) = "Value"
        .TextMatrix(0, .ColIndex("des")) = "Description"
        .TextMatrix(0, .ColIndex("opr_fullcode")) = "Operation"
        .TextMatrix(0, .ColIndex("CarName")) = "Equip"
        .TextMatrix(0, .ColIndex("Vatyo")) = "VAT %"
        .TextMatrix(0, .ColIndex("Vat")) = "VAT"
        .TextMatrix(0, .ColIndex("PriceTotal")) = "Price Total"
        .TextMatrix(0, .ColIndex("SupplierName")) = "Cash Supplier"
        .TextMatrix(0, .ColIndex("CusVATNO")) = "VAT NO."
        .TextMatrix(0, .ColIndex("Supplier")) = "Supplier Name"
        .TextMatrix(0, .ColIndex("BillNo")) = "Bill No."
        
        .TextMatrix(0, .ColIndex("PFuLLCode")) = "Project Code."
        .TextMatrix(0, .ColIndex("ProjectName")) = "Project"
        
    End With
lbl(148).Caption = "Manual Vat %"
    With VSFlexGrid1
    
     .TextMatrix(0, .ColIndex("DeparmentName")) = "Deparment"
      .TextMatrix(0, .ColIndex("PFuLLCode")) = "Project Code."
        .TextMatrix(0, .ColIndex("ProjectName")) = "Project        "
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = " Account Name"
        .TextMatrix(0, .ColIndex("Account_Serial")) = " Account Code  "
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("Des")) = "Description"
        .TextMatrix(0, .ColIndex("Vatyo")) = "VAT %"
        .TextMatrix(0, .ColIndex("Vat")) = "VAT"
        .TextMatrix(0, .ColIndex("PriceTotal")) = "Price Total"
        .TextMatrix(0, .ColIndex("SupplierName")) = "Cash Supplier"
        .TextMatrix(0, .ColIndex("CusVATNO")) = "VAT NO."
        .TextMatrix(0, .ColIndex("Supplier")) = "Supplier Name"
        .TextMatrix(0, .ColIndex("BillNo")) = "Bill No."
    End With

End Sub
Sub HidFat()
    With Me.VSFlexGrid1
        If True = True Then
            .ColHidden(.ColIndex("Vat")) = False
            .ColHidden(.ColIndex("Vatyo")) = False
        Else
            .ColHidden(.ColIndex("Vat")) = True
            .ColHidden(.ColIndex("Vatyo")) = True
        End If
    End With

    With Fg_Journal
        If True = True Then
            .ColHidden(.ColIndex("Vat")) = False
            .ColHidden(.ColIndex("Vatyo")) = False
        Else
            .ColHidden(.ColIndex("Vat")) = True
            .ColHidden(.ColIndex("Vatyo")) = True
        End If
    End With
End Sub
