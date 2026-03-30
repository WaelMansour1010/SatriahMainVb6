VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmTreament 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·Ï „‰ ÌÂ„Â «·«„—   "
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13395
   Icon            =   "FrmTreatment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   13395
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   1215
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   82
      Top             =   600
      Width           =   13335
      Begin VB.TextBox XPTxtID 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   255
         Width           =   1335
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10680
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   720
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   7320
         TabIndex        =   85
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   65273857
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   7320
         TabIndex        =   86
         Top             =   705
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmTreatment.frx":038A
         Height          =   315
         Left            =   240
         TabIndex        =   87
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
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
      Begin MSDataListLib.DataCombo DcbEmpNation 
         Bindings        =   "FrmTreatment.frx":039F
         Height          =   315
         Left            =   240
         TabIndex        =   88
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   8760
         TabIndex        =   93
         Top             =   255
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   285
         Index           =   3
         Left            =   12120
         TabIndex        =   92
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÿ·»"
         Height          =   285
         Index           =   4
         Left            =   12120
         TabIndex        =   91
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblBr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ã‰”Ì…"
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   89
         Top             =   720
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   5640
      Width           =   13335
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   8760
         TabIndex        =   76
         Top             =   240
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label XPTxtCurrent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   1890
         TabIndex        =   81
         Top             =   180
         Width           =   615
      End
      Begin VB.Label XPTxtCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   240
         TabIndex        =   80
         Top             =   180
         Width           =   495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ⁄œœ «·”Ã·« :"
         Height          =   315
         Index           =   6
         Left            =   840
         TabIndex        =   79
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·”Ã· «·Õ«·Ì:"
         Height          =   315
         Index           =   7
         Left            =   2580
         TabIndex        =   78
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ… : "
         Height          =   270
         Index           =   8
         Left            =   11520
         TabIndex        =   77
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   13920
      TabIndex        =   33
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   13440
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   13920
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   14040
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   13380
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   13395
      _cx             =   23627
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   21.75
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
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   " «·Ï „‰ ÌÂ„Â «·«„—                "
      Align           =   1
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
         TabIndex        =   20
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTreatment.frx":03B4
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
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTreatment.frx":074E
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
         Height          =   375
         Index           =   1
         Left            =   1710
         TabIndex        =   22
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTreatment.frx":0AE8
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
         Height          =   375
         Index           =   3
         Left            =   645
         TabIndex        =   23
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTreatment.frx":0E82
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
      Begin VB.Image Image1 
         Height          =   525
         Left            =   12480
         Picture         =   "FrmTreatment.frx":121C
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   5040
         Picture         =   "FrmTreatment.frx":296F
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2280
         TabIndex        =   31
         Top             =   480
         Width           =   2205
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   0
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6600
      Width           =   13305
      _cx             =   23469
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
         Left            =   12000
         TabIndex        =   10
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
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
         ButtonImage     =   "FrmTreatment.frx":65D7
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
         Left            =   10680
         TabIndex        =   11
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
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
         ButtonImage     =   "FrmTreatment.frx":CE39
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
         Left            =   9240
         TabIndex        =   12
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
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
         ButtonImage     =   "FrmTreatment.frx":1369B
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
         Left            =   7800
         TabIndex        =   13
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
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
         ButtonImage     =   "FrmTreatment.frx":19EFD
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
         Left            =   6360
         TabIndex        =   14
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
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
         ButtonImage     =   "FrmTreatment.frx":2075F
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
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   1485
         _ExtentX        =   2619
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
         ButtonImage     =   "FrmTreatment.frx":26FC1
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
         Left            =   1680
         TabIndex        =   17
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
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
         ButtonImage     =   "FrmTreatment.frx":50BE3
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
         Left            =   4920
         TabIndex        =   15
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
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
         ButtonImage     =   "FrmTreatment.frx":57445
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
         Left            =   3240
         TabIndex        =   16
         Top             =   0
         Width           =   1605
         _ExtentX        =   2831
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
         ButtonImage     =   "FrmTreatment.frx":5DCA7
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
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   13320
      TabIndex        =   25
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
      Left            =   13560
      TabIndex        =   28
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   3495
      Left            =   0
      TabIndex        =   35
      Top             =   1920
      Width           =   13305
      _cx             =   23469
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
      Picture(0)      =   "FrmTreatment.frx":64509
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3030
         Left            =   13950
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   45
         Width           =   13215
         _cx             =   23310
         _cy             =   5345
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
            Height          =   3630
            Left            =   120
            TabIndex        =   37
            Tag             =   "1"
            Top             =   240
            Width           =   13230
            _cx             =   23336
            _cy             =   6403
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
            FormatString    =   $"FrmTreatment.frx":648A3
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
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   4080
            Width           =   3375
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9960
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3030
         Index           =   15
         Left            =   45
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   45
         Width           =   13215
         _cx             =   23310
         _cy             =   5345
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
         _GridInfo       =   $"FrmTreatment.frx":649EF
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3000
            Index           =   16
            Left            =   15
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   15
            Width           =   13185
            _cx             =   23257
            _cy             =   5292
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
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Height          =   855
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   2040
               Width           =   9795
               Begin VB.TextBox TxtTelephone 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   4
                  Top             =   360
                  Width           =   3675
               End
               Begin VB.TextBox TxtCoputer 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5520
                  TabIndex        =   3
                  Top             =   360
                  Width           =   2865
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Â« ›"
                  Height          =   255
                  Index           =   12
                  Left            =   4440
                  TabIndex        =   72
                  Top             =   360
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬂ„»ÌÊ —"
                  Height          =   255
                  Index           =   11
                  Left            =   8640
                  TabIndex        =   71
                  Top             =   360
                  Width           =   960
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   1320
               Width           =   7140
               Begin VB.TextBox TxtlgTre 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3600
                  TabIndex        =   8
                  Top             =   240
                  Width           =   1905
               End
               Begin MSComCtl2.DTPicker XpDtTre 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   9
                  Top             =   240
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   65273857
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰  «—ÌŒ"
                  Height          =   285
                  Index           =   10
                  Left            =   2160
                  TabIndex        =   69
                  Top             =   255
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «· ’—ÌÕ"
                  Height          =   375
                  Index           =   9
                  Left            =   5880
                  TabIndex        =   68
                  Top             =   240
                  Width           =   840
               End
            End
            Begin VB.Frame lbty 
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·„⁄«„·…"
               Height          =   1245
               Left            =   7290
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   735
               Width           =   5910
               Begin XtremeSuiteControls.CheckBox ChPasNew 
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   62
                  Top             =   720
                  Width           =   1935
                  _Version        =   786432
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   " ÃœÌœ ÃÊ«“ «·”›—"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChIqNew 
                  Height          =   375
                  Left            =   1800
                  TabIndex        =   63
                  Top             =   240
                  Width           =   1455
                  _Version        =   786432
                  _ExtentX        =   2566
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   " ÃœÌœ ≈ﬁ«„Â"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChEnd 
                  Height          =   375
                  Left            =   4200
                  TabIndex        =   64
                  Top             =   720
                  Width           =   1455
                  _Version        =   786432
                  _ExtentX        =   2566
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "«‰Â«¡ „⁄«„·« "
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChIqEx 
                  Height          =   375
                  Left            =   3720
                  TabIndex        =   65
                  Top             =   240
                  Width           =   1935
                  _Version        =   786432
                  _ExtentX        =   3413
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "≈’œ«— ≈ﬁ«„Â"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChNeEnter 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   66
                  Top             =   240
                  Width           =   1455
                  _Version        =   786432
                  _ExtentX        =   2566
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "œŒÊ· ÃœÌœ"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin VB.Frame lbldata 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  "
               Height          =   1335
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   0
               Width           =   7260
               Begin VB.TextBox TxtIqFrom 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3600
                  TabIndex        =   6
                  Top             =   600
                  Width           =   2505
               End
               Begin VB.TextBox TxtIqama 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   120
                  TabIndex        =   7
                  Top             =   240
                  Width           =   2505
               End
               Begin VB.TextBox TxtPas 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3600
                  TabIndex        =   5
                  Top             =   240
                  Width           =   2505
               End
               Begin MSComCtl2.DTPicker XpDtExpair 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   58
                  Top             =   960
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   65273857
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker XpDtEnd 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   60
                  Top             =   960
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   65273857
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal EndDateH 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   73
                  Top             =   960
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin Dynamic_Byte.NourHijriCal ExpairDateH 
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   74
                  Top             =   960
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«‰ Â«¡"
                  Height          =   285
                  Index           =   18
                  Left            =   2520
                  TabIndex        =   61
                  Top             =   960
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«’œ«—"
                  Height          =   285
                  Index           =   17
                  Left            =   6120
                  TabIndex        =   59
                  Top             =   975
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„’œ—«·«ﬁ«„…"
                  Height          =   375
                  Index           =   16
                  Left            =   6240
                  TabIndex        =   57
                  Top             =   600
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·«ﬁ«„…"
                  Height          =   375
                  Index           =   14
                  Left            =   2640
                  TabIndex        =   56
                  Top             =   240
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ÃÊ«“"
                  Height          =   375
                  Index           =   5
                  Left            =   6240
                  TabIndex        =   55
                  Top             =   240
                  Width           =   840
               End
            End
            Begin VB.Frame lbdataEm 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·„ÊŸ›"
               Height          =   720
               Left            =   7260
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   0
               Width           =   5925
               Begin MSDataListLib.DataCombo DcboEmpDepartments 
                  Height          =   315
                  Left            =   3000
                  TabIndex        =   1
                  Top             =   240
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboJobsType 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   2
                  Top             =   240
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Êﬁ⁄"
                  Height          =   285
                  Index           =   15
                  Left            =   5040
                  TabIndex        =   50
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊŸÌ›…"
                  Height          =   285
                  Index           =   24
                  Left            =   2160
                  TabIndex        =   49
                  Top             =   240
                  Width           =   645
               End
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   660
               Left            =   120
               TabIndex        =   53
               Top             =   2160
               Visible         =   0   'False
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   1164
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
               ButtonImage     =   "FrmTreatment.frx":64A23
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
               Enabled         =   0   'False
               Height          =   2235
               Index           =   62
               Left            =   2550
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   1035
               Width           =   570
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3000
            Index           =   9
            Left            =   15
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   15
            Width           =   13185
            _cx             =   23257
            _cy             =   5292
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
               Height          =   2250
               Left            =   3435
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   615
               Width           =   735
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   1560
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   825
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   1560
               Index           =   67
               Left            =   2445
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   825
               Width           =   675
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   1500
               Index           =   68
               Left            =   4170
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1020
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
               Height          =   1815
               Index           =   69
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   825
               Width           =   315
            End
         End
      End
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
      TabIndex        =   34
      Top             =   3450
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   120
      Top             =   3720
      Width           =   6015
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   13050
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   13350
      TabIndex        =   26
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmTreament"
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

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean

    Cn.BeginTrans
    BeginTrans = True

    If IsNull(rs("Posted")) Then
        rs("Posted") = user_id
        rs("PostedDate") = Time
    Else
        rs("Posted") = Null
       rs("PostedDate") = Time
    End If
   
    rs.update
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

    Cn.CommitTrans
    BeginTrans = False
FillApprovedTable
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
 restch
   Dim cCompanyInfo As New ClsCompanyInfo
TxtCoputer.Text = cCompanyInfo.ArabComment
TxtTelephone.Text = cCompanyInfo.CompanyTel
           ' lbl(20).Caption = "0"
          ''  lbl(21).Caption = "0"
          '  lbl(22).Caption = "0"
         '   lbl(23).Caption = "0"
            
              GRID2.Clear flexClearScrollable, flexClearEverything
    GRID2.Rows = 1
            Me.DCboUserName.BoundText = user_id
          '  TxtPaymentCounts.text = 1
Dcbranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            
            Accredit.Enabled = True
                If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
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

            Del_Trans

        Case 5
            General_Search.send_form = "treat"
           Load General_Search
            General_Search.show
        
        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.Text, , 200

        Case 8
          '  CalCulateParts
            
            
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


MySQL = " SELECT dbo.TblTreatment.ID, dbo.TblTreatment.RecordDate, dbo.TblTreatment.BranchID, dbo.TblTreatment.NationalID, dbo.TblTreatment.ProjectID, dbo.TblTreatment.JobID,"
    MySQL = MySQL & "                             dbo.TblTreatment.IqamaID, dbo.TblTreatment.IqamaFrom, dbo.TblTreatment.EndDate, dbo.TblTreatment.ExpairDate, dbo.TblTreatment.LongTreatment,"
     MySQL = MySQL & "                            dbo.TblTreatment.TreatmentDate, dbo.TblTreatment.ComputerID, dbo.TblTreatment.Telephone, dbo.TblTreatment.IqamaEx, dbo.TblTreatment.IqamaNew,"
     MySQL = MySQL & "                            dbo.TblTreatment.PasNew, dbo.TblTreatment.EnterNew, dbo.TblTreatment.FinalTrea, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
     MySQL = MySQL & "                            dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Namee,"
      MySQL = MySQL & "                           dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblTreatment.EmpID,"
       MySQL = MySQL & "                          dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
          MySQL = MySQL & "                       dbo.EmpGroupDep.GroupName, dbo.TblEmployee.Nationality, dbo.TblEmployee.Fullcode, dbo.TblTreatment.Iqama, dbo.TblTreatment.ExpairDateH,"
          MySQL = MySQL & "                       dbo.TblTreatment.EndDateH , dbo.TblTreatment.pasid, dbo.TblEmployee.DateEndekamaH, dbo.TblEmployee.NumEkama"
    MySQL = MySQL & "           FROM     dbo.TblBranchesData RIGHT OUTER JOIN"
        MySQL = MySQL & "                         dbo.TblTreatment LEFT OUTER JOIN"
          MySQL = MySQL & "                       dbo.EmpGroupDep ON dbo.TblTreatment.ProjectID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
           MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TblTreatment.JobID = dbo.TblEmpJobsTypes.JobTypeID ON dbo.TblBranchesData.branch_id = dbo.TblTreatment.BranchID LEFT OUTER JOIN"
                MySQL = MySQL & "                 dbo.TblEmployee ON dbo.TblTreatment.EmpID = dbo.TblEmployee.Emp_ID"

 MySQL = MySQL & "   Where (dbo.TblTreatment.id =" & val(XPTxtID.Text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\treatment.rpt"
     Else
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\treatment.rpt"
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

Dim Sal  As Double
Sal = GetEmployeeSalaryAccordingToComponent(val(DcboEmpName.BoundText), "", 0)

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
         xReport.ParameterFields(2).AddCurrentValue ""
         xReport.ParameterFields(3).AddCurrentValue "" & Sal & ""
         xReport.ParameterFields(4).AddCurrentValue "" & Sal & ""
         StrReportTitle = "" '& StrAccountName
     
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue ""
        xReport.ParameterFields(3).AddCurrentValue "" & Sal & ""
        xReport.ParameterFields(4).AddCurrentValue "" & Sal & ""
   
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


 
  
 
End Function

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub


Private Sub txtDiscountDES_Change()

End Sub



Private Sub ENDDATEH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
     
      XpDtEnd.value = ToGregorianDate(XpDtEnd.value)
       
End If
End Sub
Private Sub ExpairDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
     
      XpDtExpair.value = ToGregorianDate(ExpairDateH.value)
       
End If
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub

 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 9
       ' Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If

End Sub

Private Sub DcboEmpName_Click(Area As Integer)
   On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
    
   If Me.TxtModFlg = "R" Then Exit Sub
   
   
    Dim StrSQL As String

 
        
        
        Dim IssueDate As Date
        Dim DepID As Double
        Dim specid As Double
        Dim JobTypeID As Double
        'Dim JobTypeIDiqama As Integer
        Dim JobTypeIDIQ As Integer
        Dim gradeID As Double
        Dim Account_code2 As String
           Dim Account_code  As String
        Dim Balance As String
        Dim endContractPerMonth As Double
        Dim national As String
        Dim Project As Integer
        Dim pasid As String
      Dim iqamaid As String
      Dim placeiqama As String
      Dim endiq As String
      Dim DateExpoekama As String
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth, national, , , Project, pasid, iqamaid, placeiqama, , endiq, , , , , , , , , JobTypeIDIQ, , DateExpoekama
        
      WriteCustomerBalPublic Account_code2, Balance
          
  'lbl(22).Caption = val(Balance)

         WriteCustomerBalPublic Account_code, Balance
          
 ' lbl(21).Caption = val(Balance)
 ' l'bl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
     '   DBIssueDate.value = issuedate
   DcboEmpDepartments.BoundText = Project
     '   DcboSpecifications.BoundText = gradeID
     Me.TxtIqFrom.Text = placeiqama
     DcbEmpNation.Text = national
        DcboJobsType.BoundText = JobTypeIDIQ
        TxtIqama.Text = iqamaid
        Me.EndDateH.value = endiq
        ExpairDateH.value = DateExpoekama
      Me.TxtPas.Text = pasid
    ''''
  
     '   lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
        
    'End If

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

Sub restch()
Me.ChIqEx.value = vbUnchecked
Me.ChIqNew.value = vbUnchecked
Me.ChEnd.value = vbUnchecked
Me.ChPasNew.value = vbUnchecked
Me.ChNeEnter.value = vbUnchecked

End Sub
Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

  '  With Me.Fg
      '  .RowHeightMin = 300
      '  .WallPaper = GrdBack.Picture
      '  .AutoSize 0, .Cols - 1, False
   ' End With

    Set TTD = New clstooltipdemand
   ' Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
  '  Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
   ' Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
   ' Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
   ' Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
  ' ' Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
   ' Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
  '  Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEmpLocations Me.DcboEmpDepartments
    Dcombos.GetEmpJobsTypes Me.DcboJobsType
    Dcombos.GetEmployeesNationlity Me.DcbEmpNation

  '  Dcombos.GetEmpGrades Me.DcboSpecifications
    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
    YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblTreatment     Order By ID"
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
'    Label1.Visible = False

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    Me.Caption = "To Whom It May Concern Letter"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(3).Caption = "Employee"
    lblBr.Caption = "Branch"
    lbl(2).Caption = "Nationality"
    lbl(15).Caption = "Location"
    lbl(24).Caption = "Job"
    lbl(5).Caption = "Passport No"
    lbl(14).Caption = "Residence No"
    lbl(16).Caption = "Issuing Authority"
    lbl(17).Caption = "Date of Issue"
    lbl(18).Caption = "Date of Expiry"
    ChEnd.Caption = "End Treatment"
    ChIqEx.Caption = "Residence Issue"
   ChIqNew.Caption = "Residence renewal"
   ChNeEnter.Caption = "New Entry"
   ChPasNew.Caption = "Passport renewal"
   ChEnd.RightToLeft = False
   ChIqEx.RightToLeft = False
   ChIqNew.RightToLeft = False
  ChNeEnter.RightToLeft = False
   ChPasNew.RightToLeft = False
   lbl(9).Caption = "Validity"
   lbl(10).Caption = "From Date"
   lbl(11).Caption = "Computer No"
   lbl(12).Caption = "Phone No"
   lbldata.Caption = "Data"
   lbdataEm.Caption = "Data of Employee"
   lbty.Caption = "Type of Treatment"
   XPTab301.Caption = "Data"
    'lbl(2).Caption = "value"
 '   lbl(0).Caption = "Box"
    'Fra(0).Caption = "payments Method"
    'lbl(9).Caption = "Count"
   ' lbl(10).Caption = "Start"
    'lbl(11).Caption = "Month"
   ' lbl(12).Caption = "Year"
    'Cmd(8).Caption = "Calc Dates"
  '  ChkSaleryDis.Caption = "Auto Discount"
   lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"

  '  With Me.FG
   '     .TextMatrix(0, .ColIndex("PartNO")) = "NO"
     '   .TextMatrix(0, .ColIndex("PartValue")) = "Value"
    '    .TextMatrix(0, .ColIndex("PartDate")) = "Date"

   ' End With

End Sub

Private Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

   ' CmbMonth.Clear

   ' For i = 1 To 12
     '   CmbMonth.AddItem MonthName(i)
  '  Next

  '  CmbMonth.ListIndex = Month(Date) - 1
  '  CboYear.Clear

   ' For i = 2010 To 2050
       ' CboYear.AddItem i

       ' If i = year(Date) Then
       '     IntDefIndex = CboYear.NewIndex
       ' End If

  ' Next

    'CboYear.ListIndex = IntDefIndex
End Sub

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

Private Sub TxtAdvanceValue_LostFocus()
 
   
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            '        Me.Caption = "«·Ï „‰ ÌÂ„Â «·«„—   "
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
            '        Me.Caption = "«·Ï „‰ ÌÂ„Â «·«„—   ( ÃœÌœ )"
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
        '    TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = "«·Ï „‰ ÌÂ„Â «·«„—   (  ⁄œÌ· )"
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
        '    TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtPaymentCounts_KeyPress(KeyAscii As Integer)

End Sub

Private Sub TxtPaymentCounts_LostFocus()
 
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

    XPTxtID.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    DcboEmpDepartments.BoundText = IIf(IsNull(rs("ProjectID").value), "", rs("ProjectID").value)
    DcboJobsType.BoundText = IIf(IsNull(rs("JobID").value), "", rs("JobID").value)
    TxtPas.Text = IIf(IsNull(rs("PasID").value), "", rs("PasID").value)
    TxtIqama.Text = IIf(IsNull(rs("Iqama").value), "", rs("Iqama").value)
   '  TxtIqama.text = val(IIf(IsNull(rs("IqamaID").value), "", rs("IqamaID").value))
    TxtIqFrom.Text = IIf(IsNull(rs("IqamaFrom").value), "", rs("IqamaFrom").value)
    XpDtExpair.value = IIf(IsNull(rs("ExpairDate").value), "", rs("ExpairDate").value)
     XpDtEnd.value = IIf(IsNull(rs("EndDate").value), "", rs("EndDate").value)
     TxtlgTre.Text = IIf(IsNull(rs("LongTreatment").value), "", rs("LongTreatment").value)
     XpDtTre.value = IIf(IsNull(rs("TreatmentDate").value), "", rs("TreatmentDate").value)
    TxtCoputer.Text = IIf(IsNull(rs("ComputerId").value), "", rs("ComputerId").value)
    TxtTelephone.Text = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
    Me.DcbEmpNation.BoundText = IIf(IsNull(rs("NationalID").value), "", rs("NationalID").value)
    ExpairDateH.value = IIf(IsNull(rs("ExpairDateH").value), "", rs("ExpairDateH").value)
     EndDateH.value = IIf(IsNull(rs("EndDateH").value), "", rs("EndDateH").value)
    If rs("IqamaEx").value = True Then
    Me.ChIqEx.value = vbChecked
    Else
    Me.ChIqEx.value = vbUnchecked
    End If
    If rs("IqamaNew").value = True Then
    Me.ChIqNew = vbChecked
    Else
    Me.ChIqNew.value = vbUnchecked
    End If
    If rs("PasNew").value = True Then
    Me.ChPasNew.value = vbChecked
    Else
    Me.ChPasNew.value = vbUnchecked
    End If
    If rs("EnterNew").value = True Then
    Me.ChNeEnter.value = vbChecked
    Else
    Me.ChNeEnter.value = vbUnchecked
    End If
    If rs("FinalTrea").value = True Then
    Me.ChEnd.value = vbChecked
    Else
    Me.ChEnd.value = vbUnchecked
    End If
    
    
    
    
  '  DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)

    

 '  lbl(23).Caption = IIf(IsNull(rs("basicSalary").value), "", rs("basicSalary").value)
 
   ' lbl(22).Caption = IIf(IsNull(rs("EmpDue").value), "", rs("EmpDue").value)
  ' lbl(20).Caption = IIf(IsNull(rs("Contractvalid").value), "", rs("Contractvalid").value)
 '  lbl(21).Caption = IIf(IsNull(rs("oldAdvance").value), "", rs("oldAdvance").value)
 
'TxtDiscount.text = IIf(IsNull(rs("Discount").value), "", rs("Discount").value)
'txtDiscountDES.text = IIf(IsNull(rs("DiscountDES").value), "", rs("DiscountDES").value)

 

  
'    TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
  '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
  '  Me.TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), "", rs("PaymentCounts").value)
 
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
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
   
   
    Set RsDetails = New ADODB.Recordset
    StrSQL = "Select * From  TblTreatment Where ID=" & val(XPTxtID.Text)
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  '  Fg.Clear flexClearScrollable, flexClearEverything
  '  Fg.Rows = Fg.FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
       ' Fg.Rows = Fg.FixedRows + RsDetails.RecordCount

       ' For i = Me.Fg.FixedRows To Fg.Rows - 1
         '   Fg.TextMatrix(i, Fg.ColIndex("PartNO")) = RsDetails("PartNO").value
         '   Fg.TextMatrix(i, Fg.ColIndex("PartValue")) = RsDetails("PartValue").value
        '    Fg.TextMatrix(i, Fg.ColIndex("PartDate")) = DisplayDate(CDate(RsDetails("PartDate").value))
         '   RsDetails.MoveNext
      '  Next i

    End If

    RsDetails.Close
    Set RsDetails = Nothing
    
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

   

      '  If CheckPartCal = False Then
        '    Exit Sub
     '   End If

      '  If CheckDate = False Then
      '      Exit Sub
      '  End If

        '”·› ”«»ﬁ…
        Dim RsTest As New ADODB.Recordset
        'Set RsTest = New ADODB.Recordset
       ' StrSQL = "SELECT dbo.TblEmpAdvanceRequest.AdvanceID, dbo.TblEmpAdvanceRequest.Emp_ID, dbo.TblEmpAdvanceRequestDetails.Payed, dbo.TblEmpAdvanceRequestDetails.PartValue FROM dbo.TblEmpAdvanceRequest INNER JOIN dbo.TblEmpAdvanceRequestDetails ON dbo.TblEmpAdvanceRequest.AdvanceID = dbo.TblEmpAdvanceRequestDetails.AdvanceID WHERE (dbo.TblEmpAdvanceRequestDetails.Payed IS NULL) AND (dbo.TblEmpAdvanceRequest.Emp_ID =" & Me.DcboEmpName.BoundText & ")"
        'RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        'If RsTest.RecordCount > 0 Then
        'MsgBox "«·„ÊŸ› " & DcboEmpName.text & "  ⁄·ÌÂ ”·› ”«»ﬁ… ·„  ”œœ »⁄œ"
        'RsTest.Close
        ' Exit Sub
        'End If

   '     If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.TxtAdvanceValue.text), Me.XPDtbTrans.value) = False Then
   '         Exit Sub
   '     End If

   '     CalCulateParts
    
 
        
 '       If TxtNoteSerial1.text = "" Then
 '           If Voucher_coding(val(my_branch), XPDtbTrans.value, 32, 8032) = "error" Then
 '               MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ”ÃÌ· ”·›  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
 '           Else
 '
 '               If Voucher_coding(val(my_branch), XPDtbTrans.value, 32, 8032) = "" Then
 '                   MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ  ”ÃÌ· ”·›   ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
 '               Else
 '                   TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 32, 8032)
 '               End If
 '           End If
'        End If
'
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then

            XPTxtID.Text = CStr(new_id("TblTreatment", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
           ' StrSQL = "Delete From TblTreatment Where ID=" & val(Me.XPTxtID.text)
          '  Cn.Execute StrSQL, , adExecuteNoRecords
'
        End If
        rs("ID").value = val(XPTxtID.Text)
        rs("RecordDate").value = XPDtbTrans.value
        rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
         rs("EmpID").value = val(Me.DcboEmpName.BoundText)
        rs("ProjectID").value = val(Me.DcboEmpDepartments.BoundText)
        rs("JobID").value = val(Me.DcboJobsType.BoundText)
        rs("PasID").value = Me.TxtPas.Text
       rs("Iqama").value = Me.TxtIqama.Text
        rs("IqamaFrom").value = Me.TxtIqFrom.Text
        rs("ExpairDate").value = Me.XpDtExpair.value
        rs("EndDate").value = Me.XpDtEnd.value
        rs("LongTreatment").value = Me.TxtlgTre.Text
      rs("TreatmentDate").value = Me.XpDtTre.value
      rs("ExpairDateh").value = Me.ExpairDateH.value
        rs("EndDateh").value = Me.EndDateH.value
        
        rs("ComputerID").value = Me.TxtCoputer.Text
        rs("Telephone").value = Me.TxtTelephone.Text
        rs("NationalID").value = IIf(Me.DcbEmpNation.BoundText = "", Null, Me.DcbEmpNation.BoundText)
        
        
        
       ' rs("FirstDate").value = IIf(IsDate(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartDate"))), Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartDate")), Null)
       ' rs("PaymentCounts").value = val(Me.TxtPaymentCounts.text)
        rs("UserID").value = Me.DCboUserName.BoundText
         If ChIqEx.value = vbChecked Then
         rs("IqamaEx").value = 1
         Else
         rs("IqamaEx").value = 0
        End If
    If ChIqNew = vbChecked Then
    rs("IqamaNew").value = 1
    Else
    rs("IqamaNew").value = 0
    End If
    If ChPasNew.value = vbChecked Then
    rs("PasNew").value = 1
    Else
     rs("PasNew").value = 0
    End If
    If ChNeEnter.value = vbChecked Then
    rs("EnterNew").value = 1
    Else
     rs("EnterNew").value = 0
     End If
    If ChEnd.value = vbChecked Then
    rs("FinalTrea").value = 1
    Else
      rs("FinalTrea").value = 0
     End If



    rs("iqamaid").value = val(TxtIqama.Text)


        rs.update
        Set RsDetails = New ADODB.Recordset
        RsDetails.Open "TblTreatment", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

       ' For i = Me.Fg.FixedRows To Fg.Rows - 1
        '    RsDetails.AddNew
        ''    RsDetails("AdvanceID").value = val(XPTxtID.text)
          '  RsDetails("PartNO").value = Fg.TextMatrix(i, Fg.ColIndex("PartNO"))
         '   RsDetails("PartValue").value = Fg.TextMatrix(i, Fg.ColIndex("PartValue"))
          '  RsDetails("PartDate").value = Fg.TextMatrix(i, Fg.ColIndex("PartDate"))
          '  RsDetails.update
      '  Next i
    
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
               End If
    
        Cn.CommitTrans
        BeginTrans = False
      '  RsDetails.Close
        Set RsDetails = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
           '     Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
           '     Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"


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
          '      MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          

If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
Else
MsgBox "Update success", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End If

        End Select

        TxtModFlg.Text = "R"
  

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
            rs.find "id='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

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
                rs.delete
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
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
StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
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
        .Create Me.hwnd, "«·Ï „‰ ÌÂ„Â «·«„—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·Ï „‰ ÌÂ„Â «·«„—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·Ï „‰ ÌÂ„Â «·«„—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·Ï „‰ ÌÂ„Â «·«„—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·Ï „‰ ÌÂ„Â «·«„—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·Ï „‰ ÌÂ„Â «·«„—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hwnd, "«·Ï „‰ ÌÂ„Â «·«„—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·Ï „‰ ÌÂ„Â «·«„—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·Ï „‰ ÌÂ„Â «·«„—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·Ï „‰ ÌÂ„Â «·«„—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·Ï „‰ ÌÂ„Â «·«„—   ", 1, 15204351, -2147483630
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

Private Sub TxtAdvanceValue_KeyPress(KeyAscii As Integer)
  
End Sub

 

 
Private Sub XpDtEnd_Change()
If Me.TxtModFlg.Text <> "R" Then
     
         EndDateH.value = ToHijriDate(XpDtEnd.value)
       
End If
End Sub

Private Sub XpDtExpair_Change()
If Me.TxtModFlg.Text <> "R" Then
     
         ExpairDateH.value = ToHijriDate(XpDtExpair.value)
       
End If
End Sub

