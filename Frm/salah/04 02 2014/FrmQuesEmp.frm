VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmQUesEmp 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  «” »Ì«‰ ⁄‰ „ÊŸ›  "
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12690
   Icon            =   "FrmQuesEmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   12690
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   1695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   85
      Top             =   600
      Width           =   12735
      Begin VB.ComboBox Contract_period 
         Height          =   315
         ItemData        =   "FrmQuesEmp.frx":038A
         Left            =   13200
         List            =   "FrmQuesEmp.frx":0394
         RightToLeft     =   -1  'True
         TabIndex        =   107
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10200
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtFile 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CheckBox RegCK 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«÷—«—Ì…"
         Height          =   255
         Index           =   1
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox RegCK 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—”„Ì…"
         Height          =   255
         Index           =   0
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Ã«“…"
         Height          =   195
         Index           =   1
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«” ⁄·«„"
         Height          =   195
         Index           =   0
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   720
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   7620
         TabIndex        =   93
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   51838977
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   6960
         TabIndex        =   94
         Top             =   1080
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmQuesEmp.frx":03A2
         Height          =   315
         Left            =   2760
         TabIndex        =   95
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
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
      Begin MSDataListLib.DataCombo DcNational 
         Height          =   315
         Left            =   2760
         TabIndex        =   96
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   8550
         TabIndex        =   105
         Top             =   255
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ã‰”Ì…"
         Height          =   285
         Index           =   2
         Left            =   5430
         TabIndex        =   104
         Top             =   1065
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   285
         Index           =   3
         Left            =   11430
         TabIndex        =   103
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÿ·»"
         Height          =   285
         Index           =   4
         Left            =   11310
         TabIndex        =   102
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lblbr 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—« » «·Õ«·Ì"
         Height          =   285
         Index           =   29
         Left            =   1440
         TabIndex        =   100
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·„·›"
         Height          =   285
         Index           =   32
         Left            =   1440
         TabIndex        =   99
         Top             =   240
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   285
         Index           =   23
         Left            =   240
         TabIndex        =   98
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·«” »Ì«‰"
         Height          =   285
         Index           =   0
         Left            =   11400
         TabIndex        =   97
         Top             =   720
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   13410
      TabIndex        =   78
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   14310
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   14190
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   13710
      TabIndex        =   75
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   14190
      TabIndex        =   74
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12765
      _cx             =   22516
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
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "  «” »Ì«‰ ⁄‰ „ÊŸ›  "
      Align           =   0
      AutoSizeChildren=   0
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
         Height          =   375
         Index           =   0
         Left            =   1185
         TabIndex        =   1
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
         ButtonImage     =   "FrmQuesEmp.frx":03B7
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
         Left            =   120
         TabIndex        =   2
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
         ButtonImage     =   "FrmQuesEmp.frx":0751
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
         Left            =   1710
         TabIndex        =   3
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
         ButtonImage     =   "FrmQuesEmp.frx":0AEB
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
         Left            =   645
         TabIndex        =   4
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
         ButtonImage     =   "FrmQuesEmp.frx":0E85
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
         Left            =   5880
         Picture         =   "FrmQuesEmp.frx":121F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2400
         TabIndex        =   20
         Top             =   0
         Width           =   2205
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   2790
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7380
      Width           =   8745
      _cx             =   15425
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
      Appearance      =   0
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   7230
         TabIndex        =   6
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   6375
         TabIndex        =   7
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   5535
         TabIndex        =   8
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   4680
         TabIndex        =   9
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   3825
         TabIndex        =   10
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   0
         TabIndex        =   11
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   855
         TabIndex        =   12
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   5
         Left            =   2760
         TabIndex        =   19
         Top             =   60
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
         Index           =   9
         Left            =   1920
         TabIndex        =   22
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
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
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   8820
      TabIndex        =   13
      Top             =   6960
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   4335
      Left            =   0
      TabIndex        =   23
      Top             =   2280
      Width           =   12720
      _cx             =   22437
      _cy             =   7646
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
      ForeColor       =   0
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "«·√’‰«›|Õ«·Â «·«⁄ „«œ"
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
      DogEars         =   0   'False
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   1
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Picture(0)      =   "FrmQuesEmp.frx":4E87
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3870
         Left            =   13365
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   45
         Width           =   12630
         _cx             =   22278
         _cy             =   6826
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
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   3270
            Left            =   120
            TabIndex        =   25
            Tag             =   "1"
            Top             =   240
            Width           =   12270
            _cx             =   21643
            _cy             =   5768
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
            FormatString    =   $"FrmQuesEmp.frx":5221
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
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   3480
            Width           =   3375
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9960
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3870
         Index           =   15
         Left            =   45
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   45
         Width           =   12630
         _cx             =   22278
         _cy             =   6826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   12
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
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   1
         ChildSpacing    =   1
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
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmQuesEmp.frx":536D
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3840
            Index           =   16
            Left            =   15
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   15
            Width           =   12600
            _cx             =   22225
            _cy             =   6773
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
            Appearance      =   5
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
            Begin VB.TextBox TXTRemarks 
               Alignment       =   1  'Right Justify
               Height          =   1095
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   83
               Top             =   2280
               Width           =   4935
            End
            Begin VB.Frame lbDW 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·⁄„·"
               Height          =   2145
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   0
               Width           =   6345
               Begin VB.TextBox TxtNuWork 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   120
                  TabIndex        =   52
                  Top             =   240
                  Width           =   1665
               End
               Begin VB.TextBox Txtlong 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   960
                  Width           =   1665
               End
               Begin MSDataListLib.DataCombo DcboJobsType 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   53
                  Top             =   960
                  Width           =   1995
                  _ExtentX        =   3519
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcProject 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   54
                  Top             =   600
                  Width           =   1995
                  _ExtentX        =   3519
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker dtstar 
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   62
                  Top             =   1440
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   51838977
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker dtendtreat 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   63
                  Top             =   1440
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   51838977
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal DBEndDate 
                  Height          =   345
                  Left            =   120
                  TabIndex        =   82
                  Top             =   600
                  Width           =   1665
                  _ExtentX        =   2778
                  _ExtentY        =   450
               End
               Begin MSDataListLib.DataCombo DcbDept 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   109
                  Top             =   240
                  Width           =   1995
                  _ExtentX        =   3519
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«œ«—…"
                  Height          =   285
                  Index           =   20
                  Left            =   5160
                  TabIndex        =   110
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊŸÌ›…"
                  Height          =   195
                  Index           =   24
                  Left            =   5280
                  TabIndex        =   61
                  Top             =   960
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Êﬁ⁄ «·⁄„·"
                  Height          =   285
                  Index           =   15
                  Left            =   5160
                  TabIndex        =   60
                  Top             =   600
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ»Ì⁄… «·⁄„·"
                  Height          =   285
                  Index           =   10
                  Left            =   1800
                  TabIndex        =   59
                  Top             =   240
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«‰ Â«¡ «·«ﬁ«„…"
                  Height          =   375
                  Index           =   13
                  Left            =   1800
                  TabIndex        =   58
                  Top             =   600
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì…«·⁄„·"
                  Height          =   285
                  Index           =   9
                  Left            =   5400
                  TabIndex        =   57
                  Top             =   1440
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·⁄ﬁœ"
                  Height          =   285
                  Index           =   11
                  Left            =   1920
                  TabIndex        =   56
                  Top             =   960
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " √—ÌŒ «·⁄Êœ… „‰ «Œ— «Ã«“…"
                  Height          =   405
                  Index           =   12
                  Left            =   1440
                  TabIndex        =   55
                  Top             =   1440
                  Width           =   1845
               End
            End
            Begin VB.Frame lblLW 
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œ… «·⁄„·"
               Height          =   1350
               Left            =   6255
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   2040
               Width           =   6345
               Begin VB.TextBox TxtDay 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   4440
                  TabIndex        =   49
                  Top             =   600
                  Width           =   1305
               End
               Begin VB.TextBox TxtMonth 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   2160
                  TabIndex        =   48
                  Top             =   600
                  Width           =   1305
               End
               Begin VB.TextBox TxtYear 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   120
                  TabIndex        =   47
                  Top             =   600
                  Width           =   1305
               End
               Begin MSComCtl2.DTPicker DTfrom 
                  Height          =   315
                  Left            =   4440
                  TabIndex        =   43
                  Top             =   240
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   51838977
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker DtTo 
                  Height          =   315
                  Left            =   2160
                  TabIndex        =   46
                  Top             =   240
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   51838977
                  CurrentDate     =   41640
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÌÊ„"
                  Height          =   315
                  Index           =   35
                  Left            =   5880
                  TabIndex        =   73
                  Top             =   720
                  Width           =   285
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”‰…"
                  Height          =   315
                  Index           =   28
                  Left            =   1680
                  TabIndex        =   72
                  Top             =   720
                  Width           =   285
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   315
                  Index           =   36
                  Left            =   3600
                  TabIndex        =   71
                  Top             =   720
                  Width           =   285
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  Height          =   195
                  Index           =   33
                  Left            =   3120
                  TabIndex        =   45
                  Top             =   240
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   285
                  Index           =   34
                  Left            =   5760
                  TabIndex        =   44
                  Top             =   240
                  Width           =   405
               End
            End
            Begin VB.Frame lblds 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·—« »"
               Height          =   2040
               Left            =   6345
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   135
               Width           =   6255
               Begin VB.TextBox Txtincrease 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   825
                  Left            =   3360
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   106
                  Top             =   480
                  Width           =   2145
               End
               Begin VB.TextBox txtab 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   465
                  Left            =   120
                  TabIndex        =   66
                  Top             =   1440
                  Width           =   2145
               End
               Begin VB.TextBox txtadd 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   465
                  Left            =   3360
                  MultiLine       =   -1  'True
                  TabIndex        =   65
                  Top             =   1440
                  Width           =   2145
               End
               Begin VB.TextBox TxtOther 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   825
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   64
                  Top             =   480
                  Width           =   2145
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì«»"
                  Height          =   285
                  Index           =   18
                  Left            =   2400
                  TabIndex        =   70
                  Top             =   1560
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈÷«›Ì"
                  Height          =   285
                  Index           =   17
                  Left            =   5520
                  TabIndex        =   69
                  Top             =   1560
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ·« "
                  Height          =   285
                  Index           =   16
                  Left            =   2160
                  TabIndex        =   68
                  Top             =   720
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·“Ì«œ« "
                  Height          =   285
                  Index           =   14
                  Left            =   5400
                  TabIndex        =   67
                  Top             =   720
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   165
                  Index           =   31
                  Left            =   3480
                  TabIndex        =   41
                  Top             =   240
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—« » ⁄‰œ «· ⁄ÌÌ‰"
                  Height          =   285
                  Index           =   5
                  Left            =   4560
                  TabIndex        =   40
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì„ﬂ‰ﬂ «· ⁄œÌ· ›Ï ﬁÌ„… «·œ›⁄«  ÌœÊÌ«ı"
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
                  Height          =   255
                  Left            =   60
                  TabIndex        =   37
                  Top             =   2280
                  Visible         =   0   'False
                  Width           =   2595
               End
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   420
               Left            =   240
               TabIndex        =   38
               Top             =   3330
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   741
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "«—”«· ··«⁄ „«œ"
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
            Begin MSDataListLib.DataCombo dcjopstatus 
               Height          =   315
               Left            =   8160
               TabIndex        =   108
               Top             =   3480
               Visible         =   0   'False
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   405
               Index           =   19
               Left            =   5175
               TabIndex        =   84
               Top             =   2280
               Width           =   990
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2220
               Index           =   62
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   1020
               Width           =   570
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3840
            Index           =   9
            Left            =   15
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   15
            Width           =   12600
            _cx             =   22225
            _cy             =   6773
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
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   0
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
            Begin VB.TextBox Text8 
               Alignment       =   1  'Right Justify
               Height          =   2880
               Left            =   3285
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   825
               Width           =   690
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   1995
               Left            =   4155
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   1020
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   1995
               Index           =   67
               Left            =   2340
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   1020
               Width           =   630
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   1920
               Index           =   68
               Left            =   3975
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   1305
               Width           =   30
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2310
               Index           =   69
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   1020
               Width           =   315
            End
         End
      End
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   13470
      TabIndex        =   79
      Top             =   3570
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   13830
      TabIndex        =   80
      Top             =   1920
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   13080
      RightToLeft     =   -1  'True
      TabIndex        =   81
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Â–… «·‘«‘…  ﬁÊ„ » ”ÃÌ· ÿ·» ”›… ‰ﬁœÌ… ÊÌ „ «Õ ”«» ﬁÌ„… «·œ›⁄ «·Ì«"
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
      Height          =   660
      Index           =   25
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4770
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   120
      Top             =   4680
      Width           =   6015
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   11565
      TabIndex        =   18
      Top             =   7035
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2670
      TabIndex        =   17
      Top             =   7110
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   930
      TabIndex        =   16
      Top             =   7110
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   330
      TabIndex        =   15
      Top             =   7140
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1980
      TabIndex        =   14
      Top             =   7140
      Width           =   615
   End
End
Attribute VB_Name = "FrmQUesEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String
Public bol As Boolean
Public novalue As Boolean
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
      
      
    Cn.BeginTrans
    BeginTrans = True

'    If IsNull(rs("Posted")) Then
'        rs("Posted") = user_id
'        rs("PostedDate") = Time
'    Else
'        rs("Posted") = Null
'       rs("PostedDate") = Time
'    End If
'
'    rs.update
SendTopost Me.Name, "TblQuesEmp", "ID", val(DcbDept.BoundText), val(Dcbranch.BoundText), val(XPTxtID.Text), XPTxtID
rs.Resync



 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

    Cn.CommitTrans
    BeginTrans = False
'FillApprovedTable
    Retrive (val(Me.XPTxtID.Text))
End Sub

Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
          '  lbl(20).Caption = "0"
        '    lbl(21).Caption = "0"
           ' lbl(22).Caption = "0"
            lbl(23).Caption = "0"
            
              GRID2.Clear flexClearScrollable, flexClearEverything
    GRID2.Rows = 1
            Me.DCboUserName.BoundText = user_id
        '    TxtPaymentCounts.text = 1
Dcbranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            
            Accredit.Enabled = True
                If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
             Opt(0).value = True
             DTTo.value = Date
             
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
  If ScreenAproved(val(XPTxtID.Text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not edit.This process associated with approvals"
         End If
         Exit Sub
       End If
       
  
  
            TxtModFlg.Text = "E"
            
            Me.DCboUserName.BoundText = user_id

        Case 2
    
            Dim Msg As String

            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
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

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
             
        If ScreenAproved(val(XPTxtID.Text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «·Õ–›.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not delete.This process associated with approvals"
         End If
         Exit Sub
       End If
  
            Del_Trans

        Case 5
        bol = True
            Load FrmQuepEmpSearch
            FrmQuepEmpSearch.show

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.Text, , 200

        Case 8
           'sa CalCulateParts
            
            
                 Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                print_report val(Me.XPTxtID.Text)
        
        
            End If
        
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



MySQL = "SELECT     dbo.TblEmployee.Fullcode, dbo.TblQuesEmp.ID, dbo.TblQuesEmp.RecordDate, dbo.TblQuesEmp.EmpID, dbo.TblQuesEmp.ProjectID, dbo.TblQuesEmp.BranchID, "
   MySQL = MySQL & "                     dbo.EmpGroupDep.GroupName, dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.branch_name, dbo.TblQuesEmp.JobID,"
   MySQL = MySQL & "                     dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
   MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Namee,"
   MySQL = MySQL & "                     dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Namee4,"
   MySQL = MySQL & "                     dbo.TblQuesEmp.SalaryAppoint, dbo.TblQuesEmp.CurrSalary, dbo.TblQuesEmp.FileNo, dbo.TblQuesEmp.NationalID, dbo.TblQuesEmp.Posted,"
   MySQL = MySQL & "                     dbo.TblQuesEmp.NutWork, dbo.TblQuesEmp.EndIqama, dbo.TblQuesEmp.StartWork, dbo.TblQuesEmp.LastTreatment, dbo.TblQuesEmp.WorkFrom,"
   MySQL = MySQL & "                     dbo.TblQuesEmp.WorkTo, dbo.TblQuesEmp.[Day], dbo.TblQuesEmp.[Month], dbo.TblQuesEmp.[Year], dbo.TblQuesEmp.Increase, dbo.TblQuesEmp.Additional,"
  MySQL = MySQL & "                      dbo.TblQuesEmp.Absent, dbo.TblQuesEmp.other, dbo.TblQuesEmp.LongCont, dbo.TblQuesEmp.PostedDate, dbo.TblQuesEmp.SpecificHolidyaType1,"
  MySQL = MySQL & "                      dbo.TblQuesEmp.SpecificHolidyaType2, dbo.TblQuesEmp.EndIqamah, dbo.TblQuesEmp.Remarks, dbo.TblQuesEmp.HolidayType, dbo.TblQuesEmp.ChekDateIQ,"
 MySQL = MySQL & "                       dbo.TblQuesEmp.jopstatusid"
 MySQL = MySQL & "  FROM         dbo.TblEmployee RIGHT OUTER JOIN"
 MySQL = MySQL & "                       dbo.TblQuesEmp ON dbo.TblEmployee.Emp_ID = dbo.TblQuesEmp.EmpID LEFT OUTER JOIN"
   MySQL = MySQL & "                     dbo.TblEmpJobsTypes ON dbo.TblQuesEmp.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
 MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.TblQuesEmp.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                       dbo.EmpGroupDep ON dbo.TblQuesEmp.ProjectID = dbo.EmpGroupDep.GroupID"
 MySQL = MySQL & "    Where (dbo.TblQuesEmp.id = " & val(XPTxtID.Text) & ")"
 If Opt(0).value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "QuesEmp.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "QuesEmp.rpt"
        End If

 Else

  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "QuesEmpVo.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "QuesEmpVo.rpt"
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    'xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
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

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
Dim str As String
dateval

End Sub

Sub dateval()
 'If Me.TxtModFlg.text <> "E" Then
   Dim astrSplitItems() As String
    Dim Result As String
    
 
     Dim diff_year As Integer
    Result = ExactAge(DTFrom.value, DTTo.value)
If Result <> "" Then
 

    astrSplitItems = Split(Result, "-")
    Txtyear.Text = astrSplitItems(0)
    TxtMonth.Text = astrSplitItems(1)
    txtDay.Text = astrSplitItems(2)
 'End If
   End If
End Sub




Private Sub DTfrom_Change()
 dateval
'DcboEmpName_Change
End Sub

Private Sub DtTo_Change()
dateval
'DcboEmpName_Change
End Sub




 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 24
       ' Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If

End Sub

Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub


    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
    
        txtFile.Text = EmpCode
        
   If Me.TxtModFlg = "R" Then Exit Sub
   
   
    Dim StrSQL As String

 
        
        
        Dim IssueDate As Date
        Dim DepID As Double
        Dim specid As Double
        Dim JobTypeID As Double
        Dim gradeID As Double
        Dim Account_code2 As String
           Dim Account_code  As String
        Dim Balance As String
        Dim ProjectID As Integer
 Dim endiqama As String
        Dim national As String
        Dim endContractPerMonth As Double
       Dim BignDateWork As Date
       Dim JobTypeName As String
       Dim JobTypeIDIQ As Integer
       Dim Contract_period As Integer
     Dim Contract_periodno As Integer
   Dim dcjopstatus As Integer
Dim LastDate As Date
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth, national, , , ProjectID, , , , , endiqama, , BignDateWork, LastDate, JobTypeName, Contract_period, Contract_periodno, , dcjopstatus, JobTypeIDIQ
        
          WriteCustomerBalPublic Account_code2, Balance
          
'  lbl(22).Caption = val(Balance)
Me.Contract_period.ListIndex = Contract_period
Me.TxtLong.Text = Contract_periodno & "     " & Me.Contract_period.Text
          WriteCustomerBalPublic Account_code, Balance
        TxtNuWork.Text = JobTypeName
'  lbl(21).Caption = val(Balance)
 ' lbl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
       ' DBIssueDate.value = issuedate
        Me.DcbDept.BoundText = DepID
      dcproject.BoundText = ProjectID
      '  DcboSpecifications.BoundText = gradeID
        DcboJobsType.BoundText = JobTypeIDIQ
        lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "", 0)
        lbl(31).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "", 1)
        TxtIncrease.Text = GetEmployeeSalaryAccordingToComponentName(val(Me.DcboEmpName.BoundText), "", 0)
      TxtOther.Text = GetEmployeeSalaryAccordingToComponentName(val(Me.DcboEmpName.BoundText), "", 1)
        DcNational.Text = national
   Me.DBENDDATE.value = (endiqama)
Me.dcjopstatus.BoundText = dcjopstatus
        dtstar.value = BignDateWork
        DTFrom.value = BignDateWork
        
        DTTo.value = Date
        
    'End If
    

 dtendtreat.value = GETlASTiSSUEDATE(val(DcboEmpName.BoundText), novalue)
 If novalue = True Then
 dtendtreat.Visible = False
 Else
 dtendtreat.Visible = True
 End If

If Opt(1).value = True Then
    DTFrom.value = dtendtreat.value
End If

End Sub

'dim  novalue as true
' db.value=GETlASTiSSUEDATE(empid,novalue)
'if novalue=true then db.hide


Function GETlASTiSSUEDATE(Emp_id As Integer, Optional novalue As Boolean) As Date
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
  sql = "SELECT     MAX(todate) AS MaxDate from dbo.TblEmpHolidaysDetails WHERE     (Emp_ID = " & Emp_id & ")"
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    If Not IsNull(rs("MaxDate").value) Then
 GETlASTiSSUEDATE = IIf(IsNull(rs("MaxDate").value), Date, rs("MaxDate").value)
novalue = False
Else
 GETlASTiSSUEDATE = Date
 novalue = True
 End If
 Else
 GETlASTiSSUEDATE = Date
 novalue = True
    End If

End Function

Private Sub DtTo_Click()
dateval
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub Opt_Click(Index As Integer)
If Opt(0).value = True Then
RegCK(0).value = vbUnchecked
RegCK(1).value = vbUnchecked
End If
If Opt(1).value = True Then
RegCK(0).Enabled = True
RegCK(1).Enabled = True
Else
RegCK(0).Enabled = False
RegCK(1).Enabled = False
End If
DcboEmpName_Click (0)
dateval
End Sub

 

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    
End Sub

Private Sub XPDtbTrans_Change()

    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""

End Sub

Private Sub Dcbranch_Click(Area As Integer)
 
    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim My_SQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from jopstatus  "
    Else
        My_SQL = "  select  id,namee  from jopstatus  "
    End If

    fill_combo dcjopstatus, My_SQL
  '  With Me.Fg
     ''   .RowHeightMin = 300
      '  .WallPaper = GrdBack.Picture
      '  .AutoSize 0, .Cols - 1, False
   ' End With
    If SystemOptions.UserInterface = EnglishInterface Then
    Contract_period.AddItem "Month"
    Contract_period.AddItem "Year"
    Else
Contract_period.AddItem "‘Â—"
Contract_period.AddItem "”‰Â"
End If
    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEmpDepartments Me.DcbDept
    Dcombos.GetEmpLocations Me.dcproject
    Dcombos.GetEmpJobsTypes Me.DcboJobsType
     Dcombos.GetEmployeesNationlity Me.DcNational
 '   Dcombos.GetEmpGrades Me.DcboSpecifications
    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
  '  YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblQuesEmp     Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.Text = "R"
    Retrive


    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
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
    Label1.Visible = False
    lbl(20).Caption = "Management"
Opt(0).Caption = "Query"
Opt(1).Caption = "Vacation"
RegCK(0).Caption = "Official"
RegCK(1).Caption = "COMPELLING"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
XPTab301.Caption = "Data|Send Approved"
    Me.Caption = "Employee Questionnaire"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lblBr.Caption = "Branch"
   lbl(32).Caption = "File No"
    lbl(3).Caption = "Employee"
    lbl(2).Caption = "Nationality"
    'lbl(0).Caption = "Box"
 '   Fra(0).Caption = "payments Method"
   ' lbl(9).Caption = "Count"
   lbl(19).Caption = "Remarks"
    lbl(9).Caption = "Start"
    lbl(35).Caption = "Day"
    lbl(36).Caption = "Month"
    lbl(28).Caption = "Year"
    lbl(33).Caption = "To"
    lbl(34).Caption = "From"
    lbl(29).Caption = "Curr Salary"
    lbl(5).Caption = "Salary on Appoint"
    lblds.Caption = "Data of Salary"
    lblLW.Caption = "Long Work"
    lbl(14).Caption = "Increase"
    lbl(17).Caption = "Additional"
    lbl(16).Caption = "Allowance"
    lbl(18).Caption = "Absent"
    lbDW.Caption = "Data of Work"
    lbl(15).Caption = "Location"
    lbl(10).Caption = "Nature of Work"
    lbl(24).Caption = "Job"
    lbl(12).Caption = "Last return from vacation"
    lbl(11).Caption = "Long Contract"
    lbl(13).Caption = "Date  Expaire Ikama"
  '  Cmd(8).Caption = "Calc Dates"
        lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
   ' Caption = "Auto Discount"
    lbl(8).Caption = "By"

    Label11.Caption = "Approval Requested By"
    
    With GRID2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
    End With

End Sub

' Private Sub YearMonth()

'    Dim i As Integer
'    Dim IntDefIndex As Integer

  '  CmbMonth.Clear

 '   For i = 1 To 12
    '    CmbMonth.AddItem MonthName(i)
   ' Next

   ' CmbMonth.ListIndex = Month(Date) - 1
   ' CboYear.Clear

  '  For i = 2010 To 2050
  '      CboYear.AddItem i
'
'        If i = year(Date) Then
'            IntDefIndex = CboYear.NewIndex
'        End If

'    Next

'    CboYear.ListIndex = IntDefIndex
'End Sub

Private Sub Form_Paint()
    TTD.Destroy
End Sub

Private Sub Form_Resize()
    TTD.Destroy
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
    'Set EmpReport = Nothing
    TTD.Destroy
    Exit Sub
ErrTrap:
End Sub



Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
        Frame1.Enabled = False
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  "
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
          '  TxtAdvanceValue.Locked = True
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
        Frame1.Enabled = True
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
          '  TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
        Frame1.Enabled = True
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  (  ⁄œÌ· )"
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
          '  TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

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
            rs.find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    Me.DcbDept.BoundText = IIf(IsNull(rs("DeptID").value), "", val(rs("DeptID").value))
    XPTxtID.Text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    Me.dcjopstatus.BoundText = val(IIf(IsNull(rs("jopstatusid").value), 0, rs("jopstatusid").value))
    Me.DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    DcboJobsType.BoundText = IIf(IsNull(rs("JobID").value), "", rs("JobID").value)
    DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    DcNational.BoundText = IIf(IsNull(rs("NationalID").value), "", rs("NationalID").value)
    txtFile.Text = IIf(IsNull(rs("FileNo").value), "", rs("FileNo").value)
    lbl(23).Caption = IIf(IsNull(rs("CurrSalary").value), "", rs("CurrSalary").value)
    lbl(31).Caption = IIf(IsNull(rs("SalaryAppoint").value), "", rs("SalaryAppoint").value)
   dcproject.BoundText = IIf(IsNull(rs("ProjectID").value), "", rs("ProjectID").value)
   TxtNuWork.Text = IIf(IsNull(rs("NutWork").value), "", rs("NutWork").value)
   DBENDDATE.value = IIf(IsNull(rs("EndIqamaH").value), Date, rs("EndIqamaH").value)
   DTTo.value = IIf(IsNull(rs("StartWork").value), Date, rs("StartWork").value)
   TxtLong.Text = IIf(IsNull(rs("LongCont").value), "", rs("LongCont").value)
   DTFrom.value = IIf(IsNull(rs("WorkFrom").value), Date, rs("WorkFrom").value)
   DTTo.value = IIf(IsNull(rs("WorkTo").value), Date, rs("WorkTo").value)
   txtDay.Text = IIf(IsNull(rs("Day").value), "", rs("Day").value)
   TxtMonth.Text = IIf(IsNull(rs("month").value), "", rs("month").value)
   Txtyear.Text = IIf(IsNull(rs("Year").value), "", rs("Year").value)
   txtab.Text = IIf(IsNull(rs("Absent").value), "", rs("Absent").value)
   TxtAdd.Text = IIf(IsNull(rs("Additional").value), "", rs("Additional").value)
   TxtOther.Text = IIf(IsNull(rs("Other").value), "", rs("Other").value)
   TxtIncrease.Text = IIf(IsNull(rs("Increase").value), "", rs("Increase").value)
   TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
   
   If IsNull(rs("HolidayType").value) Then
          Opt(0).value = True
   Else
   
  '       Opt(0).value = IIf((rs("HolidayType").value) = 1, True, False)
            If rs("HolidayType").value = False Then
             Opt(0).value = True
            ElseIf rs("HolidayType").value = True Then
             Opt(1).value = True
            End If
  
  
   End If
   
'RegCK(0).value = IIf(IsNull(rs("SpecificHolidyaType1").value), vbUnchecked, IIf((rs("SpecificHolidyaType1").value) = 1, vbChecked, vbUnchecked))
' RegCK(1).value = IIf(IsNull(rs("SpecificHolidyaType2").value), vbUnchecked, IIf((rs("SpecificHolidyaType2").value) = 1, vbChecked, vbUnchecked))
   If rs("SpecificHolidyaType1").value = True Then
     RegCK(0).value = vbChecked
   Else
   RegCK(0).value = vbUnchecked
   End If
   
     If rs("SpecificHolidyaType2").value = True Then
     RegCK(1).value = vbChecked
   Else
   RegCK(1).value = vbUnchecked
   End If
  
 
 '  txtab.text = IIf(IsNull(rs("Absent").value), "", rs("Absent").value)
 '  Me.txtadd.text = IIf(IsNull(rs("Additional").value), "", rs("Additional").value)
 '  Me.TxtOther.text = IIf(IsNull(rs("Other").value), "", rs("Other").value)
  ' Me.Txtincrease.text = IIf(IsNull(rs("Increase").value), "", rs("Increase").value)
      '  DcboEmpDepartments.BoundText = IIf(IsNull(rs("DeparmentID").value), "", rs("DeparmentID").value)

   ' DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)

    
'
 '  lbl(23).Caption = IIf(IsNull(rs("basicSalary").value), "", rs("basicSalary").value)
 
 '   lbl(22).Caption = IIf(IsNull(rs("EmpDue").value), "", rs("EmpDue").value)
  ' lbl(20).Caption = IIf(IsNull(rs("Contractvalid").value), "", rs("Contractvalid").value)
 '  lbl(21).Caption = IIf(IsNull(rs("oldAdvance").value), "", rs("oldAdvance").value)
 
'TxtDiscount.text = IIf(IsNull(rs("Discount").value), "", rs("Discount").value)
'txtDiscountDES.text = IIf(IsNull(rs("DiscountDES").value), "", rs("DiscountDES").value)

 

    
   ' TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
  '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
  '  Me.TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), "", rs("PaymentCounts").value)
 
  
       If IsNull(rs("posted").value) Then
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               Accredit.Enabled = True
  Else
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " sent to Approval   "
                                               End If
                                               Accredit.Enabled = False
   End If
   
   
 '   Set RsDetails = New ADODB.Recordset
'    StrSQL = "Select * From  TblEmpAdvanceRequestDetails Where AdvanceID=" & val(XPTxtID.text)
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    Fg.Clear flexClearScrollable, flexClearEverything
'    Fg.Rows = Fg.FixedRows

'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
  '      RsDetails.MoveFirst
   '     Fg.Rows = Fg.FixedRows + RsDetails.RecordCount

     '   For i = Me.Fg.FixedRows To Fg.Rows - 1
      '      Fg.TextMatrix(i, Fg.ColIndex("PartNO")) = RsDetails("PartNO").value
     '      Fg.TextMatrix(i, Fg.ColIndex("PartValue")) = RsDetails("PartValue").value
       '     Fg.TextMatrix(i, Fg.ColIndex("PartDate")) = DisplayDate(CDate(RsDetails("PartDate").value))
       '     RsDetails.MoveNext
      '  Next i

   ' End If

   ' RsDetails.Close
   ' Set RsDetails = Nothing
    
    fillapprovData
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Public Sub retrive1(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

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
            rs.find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    Me.DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    DcboJobsType.BoundText = IIf(IsNull(rs("JobID").value), "", rs("JobID").value)
    DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    DcNational.BoundText = IIf(IsNull(rs("NationalID").value), "", rs("NationalID").value)
      txtFile.Text = IIf(IsNull(rs("FileNo").value), "", rs("FileNo").value)
       lbl(23).Caption = IIf(IsNull(rs("CurrSalary").value), "", rs("CurrSalary").value)
       lbl(31).Caption = IIf(IsNull(rs("SalaryAppoint").value), "", rs("SalaryAppoint").value)
   dcproject.BoundText = IIf(IsNull(rs("ProjectID").value), "", rs("ProjectID").value)
   TxtNuWork.Text = IIf(IsNull(rs("NutWork").value), "", rs("NutWork").value)
   DBENDDATE.value = IIf(IsNull(rs("EndIqamaH").value), "", rs("EndIqamaH").value)
   DTTo.value = IIf(IsNull(rs("StartWork").value), Date, rs("StartWork").value)
   TxtLong.Text = IIf(IsNull(rs("LongCont").value), "", rs("LongCont").value)
   DTFrom.value = IIf(IsNull(rs("WorkFrom").value), Date, rs("WorkFrom").value)
   DTTo.value = IIf(IsNull(rs("WorkTo").value), Date, rs("WorkTo").value)
   txtDay.Text = IIf(IsNull(rs("Day").value), "", rs("Day").value)
   TxtMonth.Text = IIf(IsNull(rs("month").value), "", rs("month").value)
   Txtyear.Text = IIf(IsNull(rs("Year").value), "", rs("Year").value)
   txtab.Text = IIf(IsNull(rs("Absent").value), "", rs("Absent").value)
   TxtAdd.Text = IIf(IsNull(rs("Additional").value), "", rs("Additional").value)
   TxtOther.Text = IIf(IsNull(rs("Other").value), "", rs("Other").value)
   TxtIncrease.Text = IIf(IsNull(rs("Increase").value), "", rs("Increase").value)
   TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
   
   If IsNull(rs("HolidayType").value) Then
          Opt(0).value = True
   Else
   
  '       Opt(0).value = IIf((rs("HolidayType").value) = 1, True, False)
            If rs("HolidayType").value = False Then
             Opt(0).value = True
            ElseIf rs("HolidayType").value = True Then
             Opt(1).value = True
            End If
  
  
   End If
   
'RegCK(0).value = IIf(IsNull(rs("SpecificHolidyaType1").value), vbUnchecked, IIf((rs("SpecificHolidyaType1").value) = 1, vbChecked, vbUnchecked))
' RegCK(1).value = IIf(IsNull(rs("SpecificHolidyaType2").value), vbUnchecked, IIf((rs("SpecificHolidyaType2").value) = 1, vbChecked, vbUnchecked))
  If rs("SpecificHolidyaType1").value = True Then
     RegCK(0).value = vbChecked
   Else
   RegCK(0).value = vbUnchecked
   End If
   
     If rs("SpecificHolidyaType2").value = True Then
     RegCK(1).value = vbChecked
   Else
   RegCK(1).value = vbUnchecked
   End If
   
 '  txtab.text = IIf(IsNull(rs("Absent").value), "", rs("Absent").value)
 '  Me.txtadd.text = IIf(IsNull(rs("Additional").value), "", rs("Additional").value)
 '  Me.TxtOther.text = IIf(IsNull(rs("Other").value), "", rs("Other").value)
  ' Me.Txtincrease.text = IIf(IsNull(rs("Increase").value), "", rs("Increase").value)
      '  DcboEmpDepartments.BoundText = IIf(IsNull(rs("DeparmentID").value), "", rs("DeparmentID").value)

   ' DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)

    
'
 '  lbl(23).Caption = IIf(IsNull(rs("basicSalary").value), "", rs("basicSalary").value)
 
 '   lbl(22).Caption = IIf(IsNull(rs("EmpDue").value), "", rs("EmpDue").value)
  ' lbl(20).Caption = IIf(IsNull(rs("Contractvalid").value), "", rs("Contractvalid").value)
 '  lbl(21).Caption = IIf(IsNull(rs("oldAdvance").value), "", rs("oldAdvance").value)
 
'TxtDiscount.text = IIf(IsNull(rs("Discount").value), "", rs("Discount").value)
'txtDiscountDES.text = IIf(IsNull(rs("DiscountDES").value), "", rs("DiscountDES").value)

 

    
   ' TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
  '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
  '  Me.TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), "", rs("PaymentCounts").value)
 
  
       If IsNull(rs("posted").value) Then
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               Accredit.Enabled = True
  Else
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " sent to Approval   "
                                               End If
                                               Accredit.Enabled = False
   End If
   
   
 '   Set RsDetails = New ADODB.Recordset
'    StrSQL = "Select * From  TblEmpAdvanceRequestDetails Where AdvanceID=" & val(XPTxtID.text)
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    Fg.Clear flexClearScrollable, flexClearEverything
'    Fg.Rows = Fg.FixedRows

'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
  '      RsDetails.MoveFirst
   '     Fg.Rows = Fg.FixedRows + RsDetails.RecordCount

     '   For i = Me.Fg.FixedRows To Fg.Rows - 1
      '      Fg.TextMatrix(i, Fg.ColIndex("PartNO")) = RsDetails("PartNO").value
     '      Fg.TextMatrix(i, Fg.ColIndex("PartValue")) = RsDetails("PartValue").value
       '     Fg.TextMatrix(i, Fg.ColIndex("PartDate")) = DisplayDate(CDate(RsDetails("PartDate").value))
       '     RsDetails.MoveNext
      '  Next i

   ' End If

   ' RsDetails.Close
   ' Set RsDetails = Nothing
    
    fillapprovData
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboEmpName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        Dim RsTest As New ADODB.Recordset
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then

            XPTxtID.Text = CStr(new_id("TblQuesEmp", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
           ' StrSQL = "Delete From TblEmpAdvanceRequestDetails Where ID=" & val(Me.XPTxtID.text)
         '   Cn.Execute StrSQL, , adExecuteNoRecords

        End If
           rs("ID").value = val(XPTxtID.Text)
           rs("RecordDate").value = XPDtbTrans.value
           rs("EmpID").value = IIf(DcboEmpName.BoundText = "", Null, DcboEmpName.BoundText)
           rs("jopstatusid").value = IIf(Me.dcjopstatus.BoundText = "", Null, dcjopstatus.BoundText)
           rs("JobID").value = val(Me.DcboJobsType.BoundText)
           rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
           rs("UserID").value = Me.DCboUserName.BoundText
           rs("ProjectID").value = IIf(dcproject.BoundText = "", Null, dcproject.BoundText)
           rs("Remarks").value = IIf(TxtRemarks.Text = "", Null, TxtRemarks.Text)
        If novalue = False Then
           rs("ChekDateIQ").value = 1
        Else
           rs("ChekDateIQ").value = 0
        End If
        rs("DeptID").value = IIf(Me.DcbDept.BoundText = "", Null, val(DcbDept.BoundText))
         
''''

    'rs("Additional").value = val(IIf(IsDate(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("Add"))), Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("Add")), Null))
     '  rs("Increase").value = val(IIf(IsDate(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartNO"))), Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartNO")), Null))
       ' rs("ManagerID").value = Me.DcmbManagerID.BoundTextPartValue
     '  rs("other").value = val(IIf(IsDate(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue")), Null))
   If Opt(0).value = True Then
    rs("HolidayType").value = 0
   Else
   rs("HolidayType").value = 1
   End If
   
   If RegCK(0).value = vbChecked Then
      rs("SpecificHolidyaType1").value = 1
   Else
   rs("SpecificHolidyaType1").value = 0
   End If
   
     
    If RegCK(1).value = vbChecked Then
      rs("SpecificHolidyaType2").value = 1
   Else
   rs("SpecificHolidyaType2").value = 0
   End If
   
       rs("Absent").value = val(Me.txtab.Text)
         rs("other").value = Me.TxtOther.Text
         rs("Additional").value = val(Me.TxtAdd.Text)
         rs("Increase").value = Me.TxtIncrease.Text
         
     '   rs("PaymentCounts").value = val(Me.TxtPaymentCounts.text)
     rs("SalaryAppoint").value = val(lbl(31).Caption)
      rs("CurrSalary").value = val(lbl(23).Caption)
       rs("FileNo").value = val(Me.txtFile.Text)
       rs("NationalID").value = val(Me.DcNational.BoundText)
rs("NutWork").value = Me.TxtNuWork.Text
rs("EndIqamaH").value = Me.DBENDDATE.value
rs("StartWork").value = Me.dtstar.value
rs("LastTreatment").value = Me.dtendtreat.value
rs("LongCont").value = Me.TxtLong.Text
rs("WorkFrom").value = Me.DTFrom.value
rs("WorkTo").value = Me.DTTo.value
rs("Day").value = Me.txtDay.Text
rs("month").value = Me.TxtMonth.Text
rs("Year").value = Me.Txtyear.Text
        rs.update
        Set RsDetails = New ADODB.Recordset
        RsDetails.Open "TblQuesEmp", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    '    For i = Me.Fg.FixedRows To Fg.Rows - 1
    '        RsDetails.AddNew
    '        RsDetails("ID").value = val(XPTxtID.text)
    '        RsDetails("Increase").value = Fg.TextMatrix(i, Fg.ColIndex("PartNO"))
    '        RsDetails("other").value = Fg.TextMatrix(i, Fg.ColIndex("PartValue"))
    '        RsDetails("Absent").value = Fg.TextMatrix(i, Fg.ColIndex("PartDate"))
    '        RsDetails("Additional").value = Fg.TextMatrix(i, Fg.ColIndex("Add"))
    '        RsDetails.update
        
   '     Next i
    
'        Dim NoteID As Long
'        Dim line_no As Integer
'        Dim RsNotes As New ADODB.Recordset
'        RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
'        If detect_employee_work_type = 1 Then
        
'            If Me.TxtModFlg.text = "E" Then
 
'                StrSQL = "Delete notes where NoteID=" & val(Me.TxtNoteID.text)
'                Cn.Execute StrSQL, , adExecuteNoRecords

'            End If

'            RsNotes.AddNew
'            NoteID = CStr(TxtNoteID.text)
'            RsNotes("NoteID").value = CStr(TxtNoteID.text)
'            RsNotes("NoteType").value = 8032
'            RsNotes("NoteDate").value = XPDtbTrans.value
'            RsNotes("UserID").value = user_id
'            RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '„”·”· «·ﬁÌœ
'            RsNotes("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '„”·”· «–‰ «·’—›
'            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
'            RsNotes("numbering_type1").value = sand_numbering_type(32) ' ”ÃÌ· «·”·›'‰Ê⁄  —ﬁÌ„    
'            RsNotes("sanad_year").value = year(XPDtbTrans.value)
'            RsNotes("sanad_month").value = Month(XPDtbTrans.value)
'            RsNotes("note_value_by_characters").value = WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
            '     RsNotes("remark").value = txtRemarks.text & bankDes
'            RsNotes("Branch_no").value = val(Me.Dcbranch.BoundText)
                
'            RsNotes.update
                
'            line_no = 1
        
'            Msg = "”·› „ÊŸ›Ì‰ —ﬁ„ " & val(Me.XPTxtID.text)
'            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'
'            Employee_account = get_EMPLOYEE_Account(val(Me.DcboEmpName.BoundText), "Account_Code")
'            StrAccountCode = Employee_account
'
            '        StrAccountCode = "a1a3a4" 'Õ”«» “„„ «·„ÊŸ›Ì‰
'            If ModAccounts.AddNewDev(LngDevID, 1, StrAccountCode, val(Me.TxtAdvanceValue.text), 0, Msg, NoteID, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text), , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If

'            StrAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))

'            If ModAccounts.AddNewDev(LngDevID, 2, StrAccountCode, val(Me.TxtAdvanceValue.text), 1, Msg, NoteID, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text), , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If
        
'        End If
    
        Cn.CommitTrans
        BeginTrans = False
        RsDetails.Close
        Set RsDetails = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
   '             Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
   '             Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
          Else
             Msg = " Saved  " & CHR(13)
                Msg = Msg + "Âyou need new transaction"
                
          End If
          
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                'MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
Else
MsgBox "Update success", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
            rs.find "ID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

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

    If XPTxtID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
         Deletepost Me.Name, "TblQuesEmp", "ID", val(DcbDept.BoundText), val(Dcbranch.BoundText), val(XPTxtID.Text), XPTxtID
                rs.delete
          '      StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.Text)
          '      Cn.Execute StrSQL, , adExecuteNoRecords
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub



Function FillApprovedTable()
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.Name & "')"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            currentdate = Now
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.Name
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.Text)
                   RSApproval("NoteSerial").value = val(Me.XPTxtID.Text)
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                Rs1.MoveNext
            Next i

    End If
    
    

End Function



Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData left  JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

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
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, "  «” »Ì«‰ ⁄‰ „ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  «” »Ì«‰ ⁄‰ „ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  «” »Ì«‰ ⁄‰ „ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  «” »Ì«‰ ⁄‰ „ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  «” »Ì«‰ ⁄‰ „ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  «” »Ì«‰ ⁄‰ „ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hwnd, "  «” »Ì«‰ ⁄‰ „ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  «” »Ì«‰ ⁄‰ „ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  «” »Ì«‰ ⁄‰ „ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  «” »Ì«‰ ⁄‰ „ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  «” »Ì«‰ ⁄‰ „ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

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

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

 



