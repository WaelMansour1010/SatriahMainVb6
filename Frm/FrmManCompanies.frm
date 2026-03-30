VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManCompanies 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " »Ì«‰«  ‘—ﬂ«  «·’Ì«‰…"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "FrmManCompanies.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   7770
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„ﬁ— «·‘—ﬂ…"
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
      Height          =   1785
      Index           =   5
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   3450
      Width           =   4185
      Begin VB.TextBox TxtAddress 
         Alignment       =   1  'Right Justify
         Height          =   585
         Left            =   30
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Top             =   1140
         Width           =   2985
      End
      Begin MSDataListLib.DataCombo DcboCountryID 
         Height          =   315
         Left            =   450
         TabIndex        =   54
         Top             =   150
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboGovernmentID 
         Height          =   315
         Left            =   450
         TabIndex        =   55
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboCityID 
         Height          =   315
         Left            =   450
         TabIndex        =   56
         Top             =   810
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄‰Ê«‰ »«· ›’Ì·"
         Height          =   585
         Index           =   26
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   1140
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„œÌ‰…"
         Height          =   225
         Index           =   25
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   840
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Õ«›Ÿ…"
         Height          =   225
         Index           =   24
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   510
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·œÊ·…"
         Height          =   225
         Index           =   22
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   210
         Width           =   765
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—’Ìœ «·‘—ﬂ… «·Õ«·Ì"
      ForeColor       =   &H00000080&
      Height          =   795
      Index           =   3
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   3660
      Width           =   3435
      Begin ImpulseButton.ISButton Cmd 
         Height          =   495
         Index           =   9
         Left            =   180
         TabIndex        =   49
         Top             =   240
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "⁄—÷  ﬁ—Ì— ﬂ‘› Õ”«»"
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
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   255
         Index           =   4
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   9
         Left            =   1410
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  „Õ«”»Ì…"
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
      Height          =   2205
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1410
      Width           =   3465
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ«·… «·—’Ìœ «·√›  «ÕÏ"
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
         Height          =   1305
         Index           =   1
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   870
         Width           =   3345
         Begin VB.OptionButton OptType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œÌ‰"
            Height          =   255
            Index           =   0
            Left            =   1980
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   210
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton OptType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "œ«∆‰"
            Height          =   255
            Index           =   1
            Left            =   1110
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   210
            Width           =   765
         End
         Begin VB.OptionButton OptType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "€Ì— „Õœœ"
            Height          =   255
            Index           =   2
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   210
            Width           =   1005
         End
         Begin VB.TextBox TxtOpenBalance 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   570
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   510
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker Dtp 
            Height          =   330
            Left            =   570
            TabIndex        =   39
            Top             =   870
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CustomFormat    =   "yyyy/M/d"
            Format          =   100073475
            CurrentDate     =   38718
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ”ÃÌ·"
            Height          =   285
            Index           =   7
            Left            =   2010
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   930
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·—’Ìœ "
            Height          =   255
            Index           =   5
            Left            =   1950
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   540
            Width           =   1275
         End
      End
      Begin VB.TextBox TxtCreditlimitCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   540
         Width           =   1395
      End
      Begin VB.TextBox TxtCreditLimit 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   150
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœ «·√∆ „«‰(œ«∆‰)"
         Height          =   285
         Index           =   11
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   540
         Width           =   1665
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœ «·√∆ „«‰(„œÌ‰)"
         Height          =   285
         Index           =   8
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   180
         Width           =   1665
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·≈ ’«·"
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
      Height          =   1995
      Index           =   2
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1410
      Width           =   4185
      Begin VB.TextBox TxtResponsibleContact 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   240
         Width           =   2805
      End
      Begin VB.TextBox TxtFaxNumber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   810
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1260
         Width           =   2085
      End
      Begin VB.TextBox TxtE_mail 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1590
         Width           =   2865
      End
      Begin VB.TextBox XPTxtmobile 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   810
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   945
         Width           =   2085
      End
      Begin VB.TextBox XPTxtPhone 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   810
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   2085
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„”∆Ê· «·≈ ’«·"
         Height          =   315
         Index           =   23
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·›«ﬂ”"
         Height          =   315
         Index           =   6
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·»—Ìœ «·≈·ﬂ —Ê‰Ï"
         Height          =   285
         Index           =   12
         Left            =   2940
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1590
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Â« ›"
         Height          =   285
         Index           =   3
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÃÊ«·"
         Height          =   285
         Index           =   2
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   945
         Width           =   1215
      End
   End
   Begin VB.TextBox XPTxtCusID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5610
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   705
      Width           =   855
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   585
      Left            =   3510
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   5280
      Width           =   3045
   End
   Begin VB.TextBox XPTxtCusName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3510
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1050
      Width           =   2955
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   930
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   690
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   -30
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   7785
      _cx             =   13732
      _cy             =   1032
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
      Caption         =   " »Ì«‰«  ‘—ﬂ«  «·’Ì«‰…"
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   0
         Left            =   1185
         TabIndex        =   14
         Top             =   90
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
         ButtonImage     =   "FrmManCompanies.frx":0CCA
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
         Left            =   120
         TabIndex        =   15
         Top             =   90
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
         ButtonImage     =   "FrmManCompanies.frx":1064
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
         Left            =   1710
         TabIndex        =   16
         Top             =   90
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
         ButtonImage     =   "FrmManCompanies.frx":13FE
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
         Left            =   645
         TabIndex        =   17
         Top             =   90
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
         ButtonImage     =   "FrmManCompanies.frx":1798
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   6645
      TabIndex        =   4
      Top             =   5940
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   5925
      TabIndex        =   5
      Top             =   5940
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   5205
      TabIndex        =   6
      Top             =   5940
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   3
      Left            =   4485
      TabIndex        =   7
      Top             =   5940
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   4
      Left            =   3645
      TabIndex        =   8
      Top             =   5940
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   5
      Left            =   2820
      TabIndex        =   9
      Top             =   5940
      Width           =   795
      _ExtentX        =   1402
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
      Height          =   375
      Index           =   6
      Left            =   1050
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5940
      Width           =   825
      _ExtentX        =   1455
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
      Height          =   375
      Index           =   7
      Left            =   1980
      TabIndex        =   10
      Top             =   5940
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
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
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   5940
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "„”«⁄œ…"
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   285
      Index           =   8
      Left            =   4590
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      ButtonStyle     =   1
      ButtonPositionImage=   1
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
      ButtonImage     =   "FrmManCompanies.frx":1B32
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      Height          =   285
      Left            =   210
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   5550
      Width           =   615
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      Height          =   285
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   5550
      Width           =   465
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   0
      Left            =   1590
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   5550
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   870
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   5550
      Width           =   165
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "FrmManCompanies.frx":1ECC
      Top             =   630
      Width           =   720
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·‘—ﬂ…"
      Height          =   285
      Index           =   1
      Left            =   6510
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   285
      Index           =   1
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·‘—ﬂ…"
      Height          =   285
      Index           =   0
      Left            =   6510
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1065
      Width           =   1185
   End
End
Attribute VB_Name = "FrmManCompanies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim CusReport As ClsCustemerReport
Dim Dcombo As ClsDataCombos
Dim cSearch(2) As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim Msg As String

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            XPTxtCusID.text = CStr(new_id("TblCustemers", "CusID", "", True))
            XPTxtCusName.SetFocus

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

            Del_Member

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            FrmCustemerSearch.Show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

            PrintReport
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub Form_Activate()
    'XPTxtCusID.SetFocus
End Sub

Private Sub Form_Load()
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    SetDtpickerDate Me.Dtp
    Resize_Form Me
    AddTip
    Dim Msg As String
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    LoadDataCombos
    StrSQL = "Select * From TblCustemers Where Type=4 Order BY CusID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2

    Me.TxtModFlg.text = "R"
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

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
    Set CusReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub OptType_Click(Index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.text)
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…"
            Else
                Me.Caption = "Maintenance Companies"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            XPTxtCusID.locked = True
            XPTxtCusName.locked = True
            XPTxtPhone.locked = True
            XPTxtmobile.locked = True
            XPMTxtRemarks.locked = True
        
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            End If

            Fra(0).Enabled = False

        Case "N"
        
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…( ÃœÌœ )"
            Else
                Me.Caption = "Maintenance Companies(Enter New Company)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            XPTxtCusID.locked = True
            XPTxtCusName.locked = False
            XPTxtPhone.locked = False
            XPTxtmobile.locked = False
            XPMTxtRemarks.locked = False
            Fra(0).Enabled = True

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…(  ⁄œÌ· )"
            Else
                Me.Caption = "Maintenance Companies(Edit Current Company)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            XPTxtCusID.locked = True
            XPTxtCusName.locked = False
            XPTxtPhone.locked = False
            XPTxtmobile.locked = False
            XPMTxtRemarks.locked = False
            Fra(0).Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtOpenBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOpenBalance.text, 0)
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim SngCusBegainAccount As Single

    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If Not (rs.EOF Or rs.BOF) Then
        If Lngid <> 0 Then
            rs.find "CusID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If

        XPTxtCusID.text = IIf(IsNull(rs("CusID")), "", val(rs("CusID")))
        XPTxtCusName.text = IIf(IsNull(rs("CusName")), "", Trim(rs("CusName")))
        Me.TxtResponsibleContact.text = IIf(IsNull(rs("ResponsibleContact").value), "", rs("ResponsibleContact").value)
        XPTxtPhone.text = IIf(IsNull(rs("Cus_Phone")), "", Trim(rs("Cus_Phone")))
        XPTxtmobile.text = IIf(IsNull(rs("Cus_mobile")), "", Trim(rs("Cus_mobile")))
        XPMTxtRemarks.text = IIf(IsNull(rs("Remark")), "", Trim(rs("Remark")))
        TxtCreditLimit.text = IIf(IsNull(rs("CreditLimit").value), "0", rs("CreditLimit").value)

        If Not (IsNull(rs("OpenBalanceDate").value)) Then
            Me.Dtp.value = rs("OpenBalanceDate").value
            Me.Dtp.Enabled = True
        Else
        
            Me.Dtp.value = Date
            Me.Dtp.Enabled = False
        End If

        If Not IsNull(rs("OpenBalanceType").value) Then
            Me.TxtOpenBalance.text = IIf(IsNull(rs("OpenBalance")), "", Trim(rs("OpenBalance")))

            If rs("OpenBalanceType").value = 0 Then
                OptType(0).value = True
                OptType_Click 0
            ElseIf rs("OpenBalanceType").value = 1 Then
                OptType(1).value = True
                OptType_Click 1
            End If
        
        Else
            Me.TxtOpenBalance.text = 0
            Me.OptType(2).value = True
            OptType_Click 2
        End If

        Me.TxtCreditlimitCredit.text = IIf(IsNull(rs("CreditlimitCredit").value), "0", rs("CreditlimitCredit").value)
        Me.TxtFaxNumber.text = IIf(IsNull(rs("FaxNumber").value), "", rs("FaxNumber").value)
        Me.TxtE_mail.text = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
        SngCusBegainAccount = GetCustomerAccount(val(XPTxtCusID.text), True)

        If SngCusBegainAccount < 0 Then
            Me.lbl(4).Caption = Abs(SngCusBegainAccount)
            Me.lbl(9).Caption = "„œÌ‰"
        ElseIf SngCusBegainAccount > 0 Then
            Me.lbl(4).Caption = Abs(SngCusBegainAccount)
            Me.lbl(9).Caption = "œ«∆‰"
        Else
            Me.lbl(4).Caption = 0
            Me.lbl(9).Caption = "Œ«·’"
        End If
    End If

    Me.DcboCountryID.BoundText = IIf(IsNull(rs("CountryID")), "", rs("CountryID"))
    Me.DcboGovernmentID.BoundText = IIf(IsNull(rs("GovernmentID")), "", rs("GovernmentID"))
    Me.DcboCityID.BoundText = IIf(IsNull(rs("CityID")), "", rs("CityID"))
    Me.TxtAddress.text = IIf(IsNull(rs("Address")), "", Trim(rs("Address")))

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Member()
    Dim Msg As String
    Dim IntRes As Integer
    Dim BegainTrans As Boolean
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If XPTxtCusID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·‘—ﬂ… —ﬁ„ " & Chr(13)
        Msg = Msg + (XPTxtCusID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                Cn.BeginTrans
                BegainTrans = True
                StrSQL = "DELETE From NOTES Where NoteType=101 AND CusID=" & val(Me.XPTxtCusID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                Dim StrAccountCode As String
                StrAccountCode = rs("Account_Code").value

                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                    rs.delete
                Else
                    Exit Sub
                End If

                Cn.CommitTrans
                BegainTrans = False
                Msg = " „  ⁄„·Ì… «·Õ–›."
                MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPBtnMove_Click 2

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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·‘—ﬂ… "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate

        If BegainTrans = True Then
            Cn.RollbackTrans
            BegainTrans = False
        End If
    End If

End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "CusID='" & val(XPTxtCusID.text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub SaveData()

    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim IntRes As Integer
    Dim BeginTrans As Boolean

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If XPTxtCusName.text = "" Then
            Msg = "ÌÃ» «œŒ«· «”„ ‘—ﬂ… «·’Ì«‰…..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtCusName.SetFocus
            Exit Sub
        End If

        If Me.OptType(2).value = False Then
            If val(Me.TxtOpenBalance.text) = 0 Then
                Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·—’Ìœ...!!!"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtOpenBalance.SetFocus
                Exit Sub
            End If
        End If

        If val(Me.TxtCreditLimit.text) > 0 Then

            'Â‰«ﬂ Õœ ≈∆ „«‰ ( „œÌ‰)ÊÌÃ» «· «ﬂœ „‰ «·—’Ìœ «·√›  «ÕÏ «·„œŒ·
            If Me.OptType(0).value = True Then

                '«·—’Ìœ «·√›  «ÕÏ „œÌ‰
                If val(Me.TxtOpenBalance.text) > val(Me.TxtCreditLimit.text) Then
                    Msg = "≈‰ »Â ....!!!"
                    Msg = Msg & Chr(13) & "·ﬁœ Ê÷⁄  Õœ ≈∆ „«‰ („œÌ‰ ) ··‘—ﬂ… " & val(Me.TxtCreditLimit.text)
                    Msg = Msg & Chr(13) & "·ﬂ‰ﬂ Ê÷⁄  «·—’Ìœ «·≈›  «ÕÏ ··‘—ﬂ… „œÌ‰ »‹  " & val(Me.TxtOpenBalance.text)
                    Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·»Ì«‰«  «· Ï «œŒ· Â«...øøø"
                    IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title)

                    If IntRes = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        If val(Me.TxtCreditlimitCredit.text) > 0 Then

            'Â‰«ﬂ Õœ ≈∆ „«‰ ( „œÌ‰)ÊÌÃ» «· «ﬂœ „‰ «·—’Ìœ «·√›  «ÕÏ «·„œŒ·
            If Me.OptType(1).value = True Then

                '«·—’Ìœ «·√›  «ÕÏ œ«∆‰
                If val(Me.TxtOpenBalance.text) > val(Me.TxtCreditlimitCredit.text) Then
                    Msg = "≈‰ »Â ....!!!"
                    Msg = Msg & Chr(13) & "·ﬁœ Ê÷⁄  Õœ ≈∆ „«‰ (œ«∆‰ ) ··‘—ﬂ… " & val(Me.TxtCreditlimitCredit.text)
                    Msg = Msg & Chr(13) & "·ﬂ‰ﬂ Ê÷⁄  «·—’Ìœ «·≈›  «ÕÏ ··‘—ﬂ… œ«∆‰ »‹  " & val(Me.TxtOpenBalance.text)
                    Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·»Ì«‰«  «· Ï «œŒ· Â«...øøø"
                    IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title)

                    If IntRes = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        Select Case TxtModFlg.text

            Case "N"
                StrSQL = "Select * From TblCustemers where CusName ='" & Trim(XPTxtCusName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If RsTemp.RecordCount > 0 Then
                    Msg = "ÌÊÃœ ‘—ﬂ… „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & Chr(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTxtCusName.SetFocus
                    Exit Sub
                End If

            Case "E"
                StrSQL = "select * From TblCustemers where CusName ='" & Trim(XPTxtCusName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("CusID").value <> val(XPTxtCusID.text) Then
                        Msg = "ÌÊÃœ ‘—ﬂ… „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & Chr(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtCusName.SetFocus
                        Exit Sub
                    End If
                End If

        End Select

        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then
            rs.AddNew
            rs("CusID").value = val(XPTxtCusID.text)
        Else

            If rs("CusID").value <> val(Me.XPTxtCusID.text) Then
                rs.find "CusID=" & val(Me.XPTxtCusID.text), , adSearchForward, adBookmarkFirst
            End If

            StrSQL = "DELETE From NOTES Where NoteType=101 AND CusID=" & val(Me.XPTxtCusID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If

        rs("CusName").value = Trim(XPTxtCusName.text)
        rs("Cus_Phone").value = IIf(XPTxtPhone.text = "", "", Trim(XPTxtPhone.text))
        rs("Cus_mobile").value = IIf(XPTxtmobile.text = "", "", Trim(XPTxtmobile.text))
        rs("Type").value = 4
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))

        If Me.OptType(2).value = True Then
            rs("OpenBalance").value = 0
            rs("OpenBalanceType").value = Null
        ElseIf Me.OptType(0).value = True Then
            rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
            rs("OpenBalanceType").value = 0
        ElseIf Me.OptType(1).value = True Then
            rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
            rs("OpenBalanceType").value = 1
        End If

        rs("OpenBalanceDate").value = Me.Dtp.value
        rs("CreditLimit").value = val(Me.TxtCreditLimit.text)
        rs("CreditlimitCredit").value = val(Me.TxtCreditlimitCredit.text)
        rs("FaxNumber").value = IIf(Trim$(Me.TxtFaxNumber.text) = "", Null, Trim$(Me.TxtFaxNumber.text))
        rs("E_mail").value = IIf(Trim$(Me.TxtE_mail.text) = "", Null, Trim$(Me.TxtE_mail.text))
        rs("CountryID").value = IIf(val(Me.DcboCountryID.BoundText) = 0, Null, val(Me.DcboCountryID.BoundText))
        rs("GovernmentID").value = IIf(val(Me.DcboGovernmentID.BoundText) = 0, Null, val(Me.DcboGovernmentID.BoundText))
        rs("CityID").value = IIf(val(Me.DcboCityID.BoundText) = 0, Null, val(Me.DcboCityID.BoundText))
        rs("ResponsibleContact").value = Trim$(Me.TxtResponsibleContact.text)
        rs("Address").value = Trim$(Me.TxtAddress.text)
    
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            If Me.TxtModFlg.text = "N" Then
                rs("Account_Code").value = ModAccounts.AddNewAccount("a2a3a3", Trim$(Me.XPTxtCusName.text), True, False)
            Else

                If Not IsNull(rs("Account_Code").value) Then
                    ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtCusName.text
                End If
            End If
        End If

        rs.update

        If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
                Dim LngDevID As Long
                Dim LngOpenID As Long
                LngOpenID = ModAccounts.AddNewOpenBalance(val(Me.XPTxtCusID.text), Me.Dtp.value)
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                If Me.OptType(0).value = True Then
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code").value, val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, "a2a1a1", val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                        GoTo ErrTrap
                    End If

                ElseIf Me.OptType(1).value = True Then

                    If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code").value, val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                        GoTo ErrTrap
                    End If
                End If
            End If
        End If

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·‘—ﬂ… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
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
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = BolRtl
    Else
        BolRtl = False
    End If

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ‘—ﬂ… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·‘—ﬂ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·‘—ﬂ… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ‘—ﬂ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ‘—ﬂ…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, BolRtl
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New..." & Wrap & "Add New Customer Data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print the current customer data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit the current customer data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the current editing or Save the new customer data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo..." & Wrap & "Undo in the adding new record" & Wrap & "OR undo editing current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete...." & Wrap & "Delete the current customer data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search" & Wrap & "Search for a customer..." & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next..." & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last..." & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Maintenance Companies Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "Show Help File", BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap

    If XPTxtCusID.text <> "" Then
        Set CusReport = New ClsCustemerReport
        CusReport.CustemerData XPTxtCusID.text
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

            Case "N"
                StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)

            Case "E"
                StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        End Select

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

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

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Customers Data"
    EleHeader.Caption = Me.Caption
    XPLbl(1).Caption = "Customer Code"
    XPLbl(0).Caption = "Customer Name"
    lbl(3).Caption = "Phone"
    lbl(2).Caption = "Mobile"
    lbl(1).Caption = "Remark"
    lbl(0).Caption = "Current Record"
    lbl(4).Caption = "NO. Recordes"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"

End Sub

Private Sub LoadDataCombos(Optional BolExceptCountries As Boolean = False, _
                           Optional BolExceptGovern As Boolean = False, _
                           Optional BolExceptCities As Boolean = False)
    Set Dcombo = New ClsDataCombos

    If BolExceptCountries = False Then
        Dcombo.GetCountriesNames Me.DcboCountryID
        Set cSearch(0) = New clsDCboSearch
        Set cSearch(0).Client = Me.DcboCountryID
    End If

    If BolExceptGovern = False Then
        Dcombo.getCountriesGovernments Me.DcboGovernmentID, val(Me.DcboCountryID.BoundText)
        Set cSearch(1) = New clsDCboSearch
        Set cSearch(1).Client = Me.DcboGovernmentID
    End If

    If BolExceptCities = False Then
        Dcombo.GetCountriesGovernCities Me.DcboCityID, val(Me.DcboCountryID.BoundText), val(Me.DcboGovernmentID.BoundText)
        Set cSearch(2) = New clsDCboSearch
        Set cSearch(2).Client = Me.DcboCityID
    End If

End Sub

Private Sub DcboCityID_Change()
    LoadDataCombos False, False, True
End Sub

Private Sub DcboCityID_Click(Area As Integer)
    DcboCityID_Change
End Sub

Private Sub DcboCountryID_Change()
    LoadDataCombos True, False, False
End Sub

Private Sub DcboCountryID_Click(Area As Integer)

    If val(Me.DcboCountryID.BoundText) <> 0 Then
        DcboCountryID_Change
    End If

End Sub

Private Sub DcboGovernmentID_Change()
    LoadDataCombos False, True, False
End Sub

Private Sub DcboGovernmentID_Click(Area As Integer)
    DcboGovernmentID_Change
End Sub

