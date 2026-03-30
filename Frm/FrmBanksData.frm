VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBanksData 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·»‰Êﬂ"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   HelpContextID   =   20
   Icon            =   "FrmBanksData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   9585
   Begin VB.TextBox TxtReportName 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   3360
      Width           =   2145
   End
   Begin VB.TextBox TxtAccountName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1080
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   1440
      Width           =   2865
   End
   Begin VB.TextBox TxtBranchName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1800
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   3720
      Width           =   2145
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ì „ «· ⁄«„· ›Ï  «·«”Â„"
      Height          =   195
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox TXTBranch_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1080
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   2520
      Width           =   2865
   End
   Begin VB.CheckBox chkLoan 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ì „ «· ⁄«„· ›Ï «·ﬁ—Ê÷"
      Height          =   195
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CheckBox chkapprov 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ì „ «· ⁄«„· ›Ï «·«⁄ „«œ« "
      Height          =   195
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox TxtAddress 
      Alignment       =   1  'Right Justify
      Height          =   1035
      Left            =   5400
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   50
      Top             =   3360
      Width           =   2865
   End
   Begin VB.TextBox TxtEmail 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1080
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   2160
      Width           =   2865
   End
   Begin VB.TextBox TxtTel 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   1800
      Width           =   2865
   End
   Begin VB.TextBox TxtIBan 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1080
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   1800
      Width           =   2865
   End
   Begin VB.TextBox Txtaccount_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   1440
      Width           =   2865
   End
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
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   4560
      Width           =   3105
      Begin VB.TextBox TxtOpenBalance 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   510
         Width           =   1365
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "€Ì— „Õœœ"
         Height          =   255
         Index           =   2
         Left            =   330
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   210
         Width           =   1005
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "œ«∆‰"
         Height          =   255
         Index           =   1
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   210
         Width           =   765
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œÌ‰"
         Height          =   255
         Index           =   0
         Left            =   2190
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   210
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.TextBox txtopening_balance_voucher_id 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComCtl2.DTPicker Dtp 
         Height          =   330
         Left            =   360
         TabIndex        =   39
         Top             =   870
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   12648447
         CustomFormat    =   "yyyy/M/d"
         Format          =   98172931
         CurrentDate     =   38718
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·—’Ìœ "
         Height          =   255
         Index           =   11
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· ”ÃÌ·"
         Height          =   285
         Index           =   10
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   930
         Width           =   1215
      End
   End
   Begin VB.TextBox XPTxtBankNamee 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1080
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1095
      Width           =   2865
   End
   Begin VB.TextBox txtreport_no 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   5400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   30
      ToolTipText     =   "Ì „ Ê÷⁄ Â‰« —ﬁ„ ‰„Ê–Ã «·‘Ìﬂ ·ÿ»«⁄ Â"
      Top             =   2520
      Width           =   2865
   End
   Begin VB.TextBox txtCommision 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5760
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2160
      Width           =   2505
   End
   Begin VB.TextBox XPTxtBankID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   1065
   End
   Begin VB.TextBox XPTxtBankName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1095
      Width           =   2865
   End
   Begin VB.TextBox XPMTxtRemark 
      Alignment       =   1  'Right Justify
      Height          =   435
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4080
      Width           =   3825
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   -30
      Width           =   9555
      _cx             =   16854
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
      Caption         =   "»Ì«‰«  «·»‰Êﬂ"
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   0
         Left            =   1155
         TabIndex        =   12
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
         ButtonImage     =   "FrmBanksData.frx":164A
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
         TabIndex        =   13
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
         ButtonImage     =   "FrmBanksData.frx":19E4
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
         TabIndex        =   14
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
         ButtonImage     =   "FrmBanksData.frx":1D7E
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
         TabIndex        =   15
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
         ButtonImage     =   "FrmBanksData.frx":2118
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   4080
         Picture         =   "FrmBanksData.frx":24B2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   8160
      TabIndex        =   3
      Top             =   6450
      Width           =   795
      _ExtentX        =   1402
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
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   7320
      TabIndex        =   4
      Top             =   6450
      Width           =   795
      _ExtentX        =   1402
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
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   6465
      TabIndex        =   5
      Top             =   6450
      Width           =   795
      _ExtentX        =   1402
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
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   3
      Left            =   5655
      TabIndex        =   6
      Top             =   6450
      Width           =   795
      _ExtentX        =   1402
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
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   4
      Left            =   4845
      TabIndex        =   7
      Top             =   6450
      Width           =   795
      _ExtentX        =   1402
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
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   6
      Left            =   1350
      TabIndex        =   9
      Top             =   6450
      Width           =   795
      _ExtentX        =   1402
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
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   375
      Left            =   2190
      TabIndex        =   8
      Top             =   6450
      Width           =   795
      _ExtentX        =   1402
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
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   1080
      TabIndex        =   24
      Top             =   720
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcCurrency 
      Height          =   315
      Left            =   5400
      TabIndex        =   53
      Top             =   720
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   375
      Left            =   3000
      TabIndex        =   59
      Top             =   6450
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
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
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   360
      TabIndex        =   61
      Top             =   6450
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
   Begin MSDataListLib.DataCombo DboParentAccount 
      Height          =   315
      Left            =   150
      TabIndex        =   62
      Top             =   4560
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton ISButton2 
      Height          =   315
      Left            =   120
      TabIndex        =   71
      Top             =   3360
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      Caption         =   "Õœœ „”«— «· ﬁ—Ì—"
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
      ButtonImage     =   "FrmBanksData.frx":611A
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      LowerToggledContent=   0   'False
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ImpulseButton.ISButton CmdSearch 
      Height          =   375
      Left            =   4140
      TabIndex        =   72
      Top             =   6450
      Width           =   675
      _ExtentX        =   1191
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
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰„Ê–Ã «·ÕÊ«·Â"
      Height          =   315
      Index           =   23
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Õ”«»"
      Height          =   315
      Index           =   22
      Left            =   4110
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   315
      Index           =   21
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Õ”«» «·—∆Ì”Ì"
      Height          =   315
      Index           =   20
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Õ”«»  "
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   19
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·›—⁄"
      Height          =   315
      Index           =   18
      Left            =   4110
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FrmBanksData.frx":C97C
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
      Height          =   855
      Index           =   17
      Left            =   120
      TabIndex        =   56
      Top             =   5040
      Width           =   4965
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄„·…"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6600
      TabIndex        =   52
      Top             =   720
      Width           =   450
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄‰Ê«‰ «·»‰ﬂ"
      Height          =   315
      Index           =   16
      Left            =   8550
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   3810
      Width           =   975
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«Ì„Ì·"
      Height          =   315
      Index           =   15
      Left            =   4110
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ·Ì›Ê‰ «·»‰ﬂ"
      Height          =   315
      Index           =   14
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«Ì»«‰"
      Height          =   315
      Index           =   13
      Left            =   4110
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·Õ”«»"
      Height          =   315
      Index           =   12
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·»‰ﬂ «‰Ã·Ì“Ì"
      Height          =   315
      Index           =   9
      Left            =   4110
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1095
      Width           =   1215
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "%"
      Height          =   315
      Index           =   8
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰”»Â «·⁄„Ê·Â"
      Height          =   315
      Index           =   7
      Left            =   8550
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·»‰ﬂ"
      Height          =   315
      Index           =   6
      Left            =   2910
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4335
      TabIndex        =   25
      Top             =   720
      Width           =   690
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ‰„Ê–Ã «·‘Ìﬂ"
      Height          =   315
      Index           =   5
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   4
      Left            =   2580
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   6000
      Width           =   1155
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·»‰ﬂ ⁄—»Ì"
      Height          =   315
      Index           =   3
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1095
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   2
      Left            =   5580
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   6000
      Width           =   1155
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   315
      Index           =   1
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   5970
      Width           =   705
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   6000
      Width           =   825
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·»‰ﬂ"
      Height          =   285
      Index           =   0
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   735
      Width           =   975
   End
End
Attribute VB_Name = "FrmBanksData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim FirstPeriodDateInthisYear As Date
Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = "SELECT     dbo.BanksData.BankID, dbo.BanksData.BankName, dbo.BanksData.Remarks, dbo.BanksData.Account_Code, dbo.BanksData.Branch, "
MySQL = MySQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.BanksData.Account_Code1, dbo.BanksData.Account_Code2, dbo.BanksData.report_no,"
MySQL = MySQL & "                      dbo.BanksData.BranchId, dbo.BanksData.Account_code3, dbo.BanksData.Commision, dbo.BanksData.ParetnAccount, dbo.BanksData.BankNamee,"
MySQL = MySQL & "                      dbo.BanksData.opening_balance_voucher_id, dbo.BanksData.OpenBalanceDate, dbo.BanksData.OpenBalanceType, dbo.BanksData.OpenBalance,"
MySQL = MySQL & "                      dbo.BanksData.account_no, dbo.BanksData.IBan, dbo.BanksData.Branch_NO, dbo.BanksData.Tel, dbo.BanksData.Address, dbo.BanksData.Email,"
MySQL = MySQL & "                      dbo.BanksData.Currency_ID , dbo.BanksData.chkapprov, dbo.BanksData.chkLoan"
MySQL = MySQL & " FROM         dbo.BanksData LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.BanksData.BranchId = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.BanksData.BankID =" & val(XPTxtBankID.Text) & ")"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBanck.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBanck.rpt"
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

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            XPTxtBankID.Text = CStr(new_id("BanksData", "BankID", "", True))
            XPTxtBankName.SetFocus
 
            Me.DcBranch.BoundText = branch_id
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

            Me.Dtp.value = FirstPeriodDateInthisYear

            OptType(2).value = True
            
            
            
                        Dim Account_Code_dynamic As String
            Account_Code_dynamic = get_account_code_branch(20, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·»‰ﬂ   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
       
                End If
            End If
        
            DboParentAccount.BoundText = Account_Code_dynamic
            

        Case 1
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

            Me.Dtp.value = FirstPeriodDateInthisYear

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"

        Case 2
            LogTextA = " Õ›Ÿ ‘«‘… " & " »Ì«‰«  «·»‰Êﬂ "
            LogTexte = " Save" & "   Banks Data "

            AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Company

        Case 5

        Case 6
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAttach_Click()
            On Error Resume Next
ShowAttachments XPTxtBankID.Text, "0701201404"
 

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub
 
Private Sub CmdSearch_Click()
     
     FrmExpensesSearch.RetrunType = 20
     FrmExpensesSearch.Indx = 1
     FrmExpensesSearch.Caption = Me.Caption
     FrmExpensesSearch.show
End Sub

Private Sub DboParentAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 16112
    End If
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches DcBranch
    End If

End Sub

Private Sub Form_Activate()
    XPTxtBankID.SetFocus
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

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " »Ì«‰«  «·»‰Êﬂ "
    LogTexte = " Open Window " & "   Banks Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    Dim My_SQL As String
    'My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    'fill_combo dcBranch, My_SQL
    My_SQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, My_SQL
    
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.DcBranch

 
    Dcombos.GetAccountingCodes Me.DboParentAccount, False, True
    
    
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.DcBranch.Enabled = True
    End If

    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
   ' rs.Open "[BanksData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   Dim StrSQL As String
     If SystemOptions.usertype <> UserAdminAll Then
      
StrSQL = "SELECT  *  From BanksData    where BranchId=" & Current_branch
  Else
 StrSQL = "SELECT  *  From BanksData"
    End If
  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
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

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰  ‘«‘… " & "»Ì«‰«  «·»‰Êﬂ "
    LogTexte = " Exit Window " & "  Banks Data "
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

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()
print_report2
End Sub

Private Sub ISButton2_Click()
Dim StrFileName As String
'StrFileName = App.path & "\REPORTS\Deposits\Deposits"
'CD1.FileName = StrFileName

 CD1.filter = "RPT File|*.rpt"
 CD1.InitDir = App.path & "\ REPORTS\Deposits"
CD1.ShowOpen

TxtReportName.Text = CD1.FileTitle
End Sub

Private Sub txtCommision_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.txtCommision.Text, 0)
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            Me.Dtp.Enabled = False
ISButton2.Enabled = False
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·»‰Êﬂ"
            Else
                Me.Caption = "Banks Data"
            End If
DboParentAccount.Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.XPTxtBankID.locked = True
            Me.XPTxtBankName.locked = True
            Me.XPMTxtRemark.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

        Case "N"


            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·»‰Êﬂ(ÃœÌœ)"
            Else
                Me.Caption = "Banks Data(New)"
            End If
            ISButton2.Enabled = True
DboParentAccount.Enabled = True
            Me.Dtp.Enabled = False
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtBankID.locked = True
            Me.XPTxtBankName.locked = False
            Me.XPMTxtRemark.locked = False

        Case "E"
        ISButton2.Enabled = True
            Me.Dtp.Enabled = False
DboParentAccount.Enabled = False
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·»‰Êﬂ(  ⁄œÌ· )"
            Else
                Me.Caption = "Banks Data( Edit )"
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
        
            Me.XPTxtBankID.locked = True
            Me.XPTxtBankName.locked = False
            Me.XPMTxtRemark.locked = False
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    Dim i As Integer
    If Lngid <> 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount

            If rs("BankID").value = Lngid Then
                GoTo ll
            End If

            rs.MoveNext
        Next i

        Exit Sub
    End If
ll:
    DcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", val(rs("BranchId").value))
    XPTxtBankID.Text = IIf(IsNull(rs("BankID").value), "", val(rs("BankID").value))
    txtCommision.Text = IIf(IsNull(rs("Commision").value), 0, rs("Commision").value)

    XPTxtBankName.Text = IIf(IsNull(rs("BankName").value), "", Trim(rs("BankName").value))
    XPTxtBankNamee.Text = IIf(IsNull(rs("BankNamee").value), "", Trim(rs("BankNamee").value))
    txtBranchName.Text = IIf(IsNull(rs("BranchName").value), "", rs("BranchName").value)
    TxtReportName.Text = IIf(IsNull(rs("ReportName").value), "", (rs("ReportName").value))
    TxtAccountName.Text = IIf(IsNull(rs("AccountName").value), "", (rs("AccountName").value))
    

    txtreport_no.Text = IIf(IsNull(rs("report_no").value), "", Trim(rs("report_no").value))

    XPMTxtRemark.Text = IIf(IsNull(rs("Remarks").value), "", Trim(rs("Remarks").value))

    Txtaccount_no.Text = IIf(IsNull(rs("account_no").value), "", Trim(rs("account_no").value))
 DboParentAccount.BoundText = IIf(IsNull(rs("parent_account").value), "", (rs("parent_account").value))
 
 If DboParentAccount.BoundText = "" Then
 DboParentAccount.BoundText = Get_Account_Parent_code(IIf(IsNull(rs("ParetnAccount").value), "", (rs("ParetnAccount").value)))
 
  If DboParentAccount.BoundText = "" Then
 DboParentAccount.BoundText = Get_Account_Parent_code(IIf(IsNull(rs("Account_Code").value), "", (rs("Account_Code").value)))
End If


End If


TXTBranch_no.Text = IIf(IsNull(rs("Branch_no").value), "", Trim(rs("Branch_no").value))
TxtIBan.Text = IIf(IsNull(rs("IBan").value), "", Trim(rs("IBan").value))
TxtTel.Text = IIf(IsNull(rs("Tel").value), "", Trim(rs("Tel").value))
TxtAddress.Text = IIf(IsNull(rs("Address").value), "", Trim(rs("Address").value))
TxtEmail.Text = IIf(IsNull(rs("Email").value), "", Trim(rs("Email").value))
    DcCurrency.BoundText = IIf(IsNull(rs("Currency_ID").value), "", (rs("Currency_ID").value))

If IsNull(rs("chkapprov").value) Then
chkapprov.value = vbUnchecked
Else
             If rs("chkapprov").value = False Then
             chkapprov.value = vbUnchecked
             Else
                chkapprov.value = vbChecked
            End If
End If


If IsNull(rs("chkLoan").value) Then
chkLoan.value = vbUnchecked
Else
             If rs("chkLoan").value = False Then
             chkLoan.value = vbUnchecked
             Else
                chkLoan.value = vbChecked
            End If
End If




    Dim FirstPeriodDateInthisYear As Date
    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

    txtopening_balance_voucher_id.Text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)

    ' If Not (IsNull(rs("OpenBalanceDate").value)) Then
    '        Me.Dtp.value = rs("OpenBalanceDate").value
    '    Me.Dtp.Enabled = True
    '    Else
        
    Me.Dtp.value = FirstPeriodDateInthisYear
    Me.Dtp.Enabled = False
    '    End If
    
    If Not IsNull(rs("OpenBalanceType").value) Then
        Me.TxtOpenBalance.Text = IIf(IsNull(rs("OpenBalance")), "", Trim(rs("OpenBalance")))

        If rs("OpenBalanceType").value = 0 Then
            OptType(0).value = True
            OptType_Click 0
        ElseIf rs("OpenBalanceType").value = 1 Then
            OptType(1).value = True
            OptType_Click 1
        End If
        
    Else
        Me.TxtOpenBalance.Text = 0
        Me.OptType(2).value = True
        OptType_Click 2
    End If

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub OptType_Click(Index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.Text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.Text)
End Sub

Private Sub TxtOpenBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOpenBalance.Text, 0)
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

Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
    On Error GoTo ErrTrap
    
    If Me.TxtModFlg.Text <> "R" Then

        If Trim(DboParentAccount.BoundText) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
               Msg = "Specify Parent Acc"
           Else
                Msg = "Õœœ «·Õ”«» «·—∆Ì”Ì «Ê·« "
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DboParentAccount.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If XPTxtBankName.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·»‰ﬂ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 Else
   MsgBox "Please Enter Bank Name ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 
 End If
            
            XPTxtBankName.SetFocus
            Exit Sub
        End If

        Select Case Me.TxtModFlg.Text

            Case "N"
                StrSQL = "select * From  BanksData where BankName='" & Trim(XPTxtBankName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
        
                            Msg = "Â‰«ﬂ »‰ﬂ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·»‰ﬂ"
                       Else
                       Msg = "This Bank Already Defined" & CHR(13)
                       End If
               
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtBankName.SetFocus
                    Exit Sub
                End If

            Case "E"
                StrSQL = "select * From  BanksData where BankName='" & Trim(XPTxtBankName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("BankID").value <> val(XPTxtBankID.Text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                        
                        Msg = "Â‰«ﬂ »‰ﬂ  „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·»‰ﬂ"
                                Else
                       Msg = "This Bank Already Defined" & CHR(13)
                       End If
                       
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        XPTxtBankName.SetFocus
                        Exit Sub
                    End If
                End If

        End Select

        Cn.BeginTrans
        BeginTrans = True

        Select Case Me.TxtModFlg.Text

            Case "N"
                Dim Account_Code_dynamic As String
                Account_Code_dynamic = get_account_code_branch(20, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                    
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                  Else
                  MsgBox "error in accounts", vbCritical
                  End If
                    GoTo ErrTrap
                    
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» ··»‰Êﬂ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If
    
                If SystemOptions.bankComm = True Then
                    Dim Account_Code_dynamic1 As String
                    Account_Code_dynamic1 = get_account_code_branch(50, my_branch)
        
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ⁄„Ê·Â ··»‰Êﬂ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
         
                        End If
                    End If
    
                End If
    
                Dim rsbank As New ADODB.Recordset
                Dim X As String
                Set rsbank = New ADODB.Recordset
                rsbank.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
                rs.AddNew
   
                rs("BankID").value = val(XPTxtBankID.Text)
                Account_Code_dynamic = DboParentAccount.BoundText
                If Not (rsbank.EOF Or rsbank.BOF) Then
                    If rsbank!banks_Accounts = True Then
         
                        X = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtBankName.Text, False, False, XPTxtBankNamee.Text)
                   
                        rs("ParetnAccount").value = X
                        rs("Account_Code").value = ModAccounts.AddNewAccount(X, XPTxtBankName.Text, True, False, XPTxtBankNamee.Text)
                        rs("Account_Code1").value = ModAccounts.AddNewAccount(X, XPTxtBankName.Text & "  ‘Ìﬂ«   Õ  «· Õ’Ì· ", True, False, XPTxtBankNamee.Text & " Under Collection Cheque")
                        rs("Account_Code2").value = ModAccounts.AddNewAccount(X, XPTxtBankName.Text & " ‘Ìﬂ«  „ƒÃ·… ⁄·Ï ·‘—ﬂ…", True, False, XPTxtBankNamee.Text & " Pending Cheque")
                   
                    Else
                        rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtBankName.Text, True, False, XPTxtBankNamee.Text)
                    
                    End If

                    If SystemOptions.bankComm = True Then
                        rs("Account_Code3").value = ModAccounts.AddNewAccount(Account_Code_dynamic1, XPTxtBankName.Text & "  ⁄Ê·« ", True, False, XPTxtBankNamee.Text & " Commission ")
                    End If
                End If
       
                ' Rs("Account_Code").value = ModAccounts.AddNewAccount("a1a2a2", XPTxtBankName.text, True, False)

        End Select
        rs("BranchName").value = Trim(txtBranchName.Text)
        rs("BranchId").value = IIf(Me.DcBranch.BoundText = "", 0, val(DcBranch.BoundText))
        rs("Commision").value = val(txtCommision.Text)
        rs("BankName").value = Trim(XPTxtBankName.Text)
        rs("BankNamee").value = Trim(XPTxtBankNamee.Text)
    rs("ReportName").value = Trim(TxtReportName.Text)
    rs("AccountName").value = Trim(TxtAccountName.Text)
        rs("report_no").value = Trim(txtreport_no.Text)
     
        rs("OpenBalanceDate").value = Me.Dtp.value
        rs("Remarks").value = IIf(XPMTxtRemark.Text = "", "", Trim(XPMTxtRemark.Text))

'new
        rs("account_no").value = IIf(Txtaccount_no.Text = "", "", Trim(Txtaccount_no.Text))
rs("Branch_no").value = IIf(TXTBranch_no.Text = "", "", Trim(TXTBranch_no.Text))
rs("IBan").value = IIf(TxtIBan.Text = "", "", Trim(TxtIBan.Text))
rs("Tel").value = IIf(TxtTel.Text = "", "", Trim(TxtTel.Text))
rs("Address").value = IIf(TxtAddress.Text = "", "", Trim(TxtAddress.Text))
rs("Email").value = IIf(TxtEmail.Text = "", "", Trim(TxtEmail.Text))
'DcCurrency
rs("Currency_ID").value = IIf(DcCurrency.BoundText = "", Null, DcCurrency.BoundText)

If chkapprov.value = vbChecked Then
rs("chkapprov").value = 1
Else
rs("chkapprov").value = 0
End If

If chkLoan.value = vbChecked Then
rs("chkLoan").value = 1
Else
rs("chkLoan").value = 0
End If

        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            If Me.TxtModFlg.Text = "N" Then
                '            Rs("Account_Code").Value = ModAccounts.AddNewAccount("a1a2a2", Trim$(Me.XPTxtBankName.text), True, False)
            Else
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
        
                If Not IsNull(rs("ParetnAccount").value) Then
                    ModAccounts.EditAccount rs("ParetnAccount").value, Me.XPTxtBankName.Text, Trim(XPTxtBankNamee.Text), , , , , , , , , , , , , , , , , False
                End If
            
                If Not IsNull(rs("Account_Code").value) Then
                    ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtBankName.Text, Trim(XPTxtBankNamee.Text), , , , , , , , , , , , , , , , , True
                End If
            
                If Not IsNull(rs("Account_Code1").value) Then
                    ModAccounts.EditAccount rs("Account_Code1").value, Me.XPTxtBankName.Text & "  ‘Ìﬂ«   Õ  «· Õ’Ì· ", Trim(XPTxtBankNamee.Text) & " Under Collection Cheque", , , , , , , , , , , , , , , , , True
                End If
            
                If Not IsNull(rs("Account_Code2").value) Then
                    ModAccounts.EditAccount rs("Account_Code2").value, Me.XPTxtBankName.Text & " ‘Ìﬂ«  „ƒÃ·… ··‘—ﬂ…", Trim(XPTxtBankNamee.Text) & " Pending Cheque", , , , , , , , , , , , , , , , , True
                End If
            
                If Not IsNull(rs("Account_Code3").value) Then
                    ModAccounts.EditAccount rs("Account_Code3").value, Me.XPTxtBankName.Text & " ⁄„Ê·« ", Trim(XPTxtBankNamee.Text) & " Commisions ", , , , , , , , , , , , , , , , , True
                End If
            
            End If
        End If
    

        If Me.OptType(2).value = True Then
            rs("OpenBalance").value = 0
            rs("OpenBalanceType").value = Null
        ElseIf Me.OptType(0).value = True Then
            rs("OpenBalance").value = val(Me.TxtOpenBalance.Text)
            rs("OpenBalanceType").value = 0
        ElseIf Me.OptType(1).value = True Then
            rs("OpenBalance").value = val(Me.TxtOpenBalance.Text)
            rs("OpenBalanceType").value = 1
        End If
   
        If val(TxtOpenBalance.Text) = 0 Then
            txtopening_balance_voucher_id = 0
        End If
       
        If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
       '     If val(Me.txtopening_balance_voucher_id.text) = 0 Then
                txtopening_balance_voucher_id.Text = get_opening_balance_voucher_id
            
       '     End If '
        End If '
   

        rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.Text)
   
   '********************************************************************
 

        rs("parent_account").value = IIf(DboParentAccount.BoundText = "", Null, (DboParentAccount.BoundText))
 
  '********************************************************************
        rs.update
        Cn.CommitTrans
    
        Dim StrDes As String

        If SystemOptions.UserInterface = ArabicInterface Then
            StrDes = "«·—’Ìœ «·≈›  «ÕÏ ·‹ "
        Else
            StrDes = " Opening Balance For: "
        End If
        
        If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
                Dim LngDevID As Long
                 
                Dim LngOpenID As Long
               
                'LngOpenID = ModAccounts.AddNewOpenBalance(Val(Me.XPTxtCusID.text), Me.Dtp.value)
                ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
            
                If Me.OptType(0).value = True Then
                   
                    Account_Code_dynamic1 = get_account_code_branch(58, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
            
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code").value, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtBankName.Text) & "  " & Trim$(Me.XPTxtBankNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtBankName.Text) & "  " & Trim$(Me.XPTxtBankNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
                    '  If ModAccounts.AddNewDev(LngDevID, 2, "a2a1a1", _
                       Val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    '         GoTo ErrTrap
                    ' End If
                ElseIf Me.OptType(1).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(58, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtBankName.Text) & "  " & Trim$(Me.XPTxtBankNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
       
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code").value, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtBankName.Text) & "  " & Trim$(Me.XPTxtBankNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If

                '   update_account_opening_balance rs("Account_Code").value
                'update_account_opening_balance Account_Code_dynamic1
                 
            End If
        End If

        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        CuurentLogdata
    
        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·»‰ﬂ" & CHR(13)
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
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
  Else
  Msg = "can't Save Data " & CHR(13)
  
  End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
Else
Msg = "Error During Saving " & CHR(13)
End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«   «·»‰Êﬂ " & CHR(13) & " ﬂÊœ «·»‰ﬂ  " & XPTxtBankID.Text & CHR(13) & " «·›—⁄ " & DcBranch.Text & CHR(13) & "«·«”„ ⁄—»Ì  " & XPTxtBankName & CHR(13) & " ‰”»… «·⁄„Ê·…   " & txtCommision & CHR(13) & " ‰„Ê–Ã «·‘Ìﬂ   " & txtreport_no
                    
    LogTextA = LogTextA & CHR(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "

    If OptType(0).value = True Then
        LogTextA = LogTextA & "„œÌ‰"
    ElseIf OptType(1).value = True Then
        LogTextA = LogTextA & "œ«∆‰"
    ElseIf OptType(2).value = True Then
        LogTextA = LogTextA & "€Ì— „Õœœ"
    End If

    LogTextA = LogTextA & CHR(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ  " & TxtOpenBalance
    LogTextA = LogTextA & CHR(13) & "„·«ÕŸ«    " & XPMTxtRemark
                     
    LogTexte = "  Save Screen  " & " Banks Data " & CHR(13) & " Bank Code    " & XPTxtBankID.Text & CHR(13) & " Branch " & DcBranch.Text & CHR(13) & " Bank Name    " & XPTxtBankNamee & CHR(13) & "Comm   " & txtCommision & CHR(13) & " Cheque Template  " & txtreport_no
                    
    LogTexte = LogTexte & CHR(13) & " Opening Balance Type  "

    If OptType(0).value = True Then
        LogTexte = LogTexte & "Debit"
    ElseIf OptType(1).value = True Then
        LogTexte = LogTexte & "Credit"
    ElseIf OptType(2).value = True Then
        LogTexte = LogTexte & "NA"
    End If

    LogTexte = LogTexte & CHR(13) & " Opening Balance Value " & TxtOpenBalance
    LogTexte = LogTexte & CHR(13) & "Remarks   " & XPMTxtRemark

    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", ""
    End If

End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "BankID='" & val(XPTxtBankID.Text) & "'", , adSearchForward, adBookmarkFirst

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

Function DeleteOpeningBalance()
    Cmd_Click (1)
    OptType(2).value = True
    TxtOpenBalance.Text = 0
    Cmd_Click (2)

End Function

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim StrAccountCode2 As String
    Dim StrAccountCode3 As String
    Dim ParetnAccount As String
    On Error GoTo ErrTrap

    If XPTxtBankID.Text <> "" Then
        StrAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
        StrAccountCode1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
        StrAccountCode2 = IIf(IsNull(rs("Account_Code2").value), "", rs("Account_Code2").value)
        StrAccountCode3 = IIf(IsNull(rs("Account_Code3").value), "", rs("Account_Code3").value)
        ParetnAccount = IIf(IsNull(rs("ParetnAccount").value), "", rs("ParetnAccount").value)
    
        StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where Account_Code='" & StrAccountCode & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode1 & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode2 & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode3 & "'"
    
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·»‰ﬂ" & CHR(13)
            Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «·»‰ﬂ"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    
        Msg = "”Ì „ Õ–› »Ì«‰«  «·»‰ﬂ —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBankID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            DeleteOpeningBalance
    
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
             
                If ModAccounts.DeleteAccount(StrAccountCode, True) = True And ModAccounts.DeleteAccount(StrAccountCode1, True) = True And ModAccounts.DeleteAccount(StrAccountCode2, True) = True And ModAccounts.DeleteAccount(StrAccountCode3, True) = True And ModAccounts.DeleteAccount(ParetnAccount, True) = True Then
                    CuurentLogdata ("D")
                    rs.delete
             
                    Msg = " „  ⁄„·Ì… «·Õ–›."
                    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            
                Else
                    GoTo ErrTrap
                End If

                rs.MoveFirst

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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·»‰ﬂ "
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  »‰ﬂ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·»‰ﬂ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·»‰ﬂ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·»‰ﬂ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ »‰ﬂ" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
Check1.Caption = "Work with Stock"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
Label1.Caption = "Currency"
chkapprov.Caption = "Work With LC"
chkLoan.Caption = "Work With Loans"
lbl(20).Caption = "Parent Acc"
CmdAttach.Caption = "Attachments"

lbl(21).Caption = "Branch"
lbl(22).Caption = "Account Name"
lbl(23).Caption = "Report Name"
ISButton2.Caption = "Select Report"
    Me.Fra(1).Caption = "Open Balance State"
    OptType(0).Caption = "Debit"
    OptType(1).Caption = "Credit"
    OptType(2).Caption = "Un Sign"
    lbl(11).Caption = "Balance Value"
    lbl(10).Caption = "Rec Date"
    lbl(12).Caption = "Account#"
    ISButton1.Caption = "Print"
    lbl(13).Caption = "I-Ban"
    lbl(14).Caption = "Tel"
    
    lbl(15).Caption = "Email"
    lbl(18).Caption = "Branch No"
    chkapprov.Caption = "Work With LC"
    lbl(16).Caption = "Addreess"
    
    lbl(15).Caption = "Email"
    lbl(17).Caption = "This screen Allow to Create Banks Data"
    
     
    Me.Caption = "Banks Data"
    EleHeader.Caption = Me.Caption
    lbl(0).Caption = "Bank Code"
    Label3.Caption = "Branch"
    lbl(7).Caption = "Comm%"
    lbl(3).Caption = "Bank Name Ar"
    lbl(9).Caption = "Bank Name En"
    lbl(1).Caption = "Remarks"
    lbl(2).Caption = "Current Record"
    lbl(4).Caption = "NO. Recordes"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    'Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    lbl(5).Caption = "Report #"

End Sub

Private Sub XPTxtBankName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtBankNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
