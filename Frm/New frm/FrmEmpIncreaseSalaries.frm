VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEmpIncreaseSalaries 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  “Ì«œ… —« » ·„ÊŸ›"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13830
   Icon            =   "FrmEmpIncreaseSalaries.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   13830
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   975
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   600
      Width           =   13815
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   116
         Top             =   120
         Width           =   2535
         Begin VB.ComboBox CBTybe 
            Height          =   315
            ItemData        =   "FrmEmpIncreaseSalaries.frx":6852
            Left            =   120
            List            =   "FrmEmpIncreaseSalaries.frx":6854
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   360
            Width           =   2235
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·“Ì«œ…"
            Height          =   285
            Index           =   35
            Left            =   1320
            TabIndex        =   117
            Top             =   0
            Width           =   1125
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   113
         Top             =   120
         Width           =   1695
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   23
            Left            =   120
            TabIndex        =   115
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—« » «·Õ«·Ì"
            Height          =   285
            Index           =   29
            Left            =   480
            TabIndex        =   114
            Top             =   0
            Width           =   1125
         End
      End
      Begin VB.ComboBox Contract_period 
         Height          =   315
         ItemData        =   "FrmEmpIncreaseSalaries.frx":6856
         Left            =   15000
         List            =   "FrmEmpIncreaseSalaries.frx":6860
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11040
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   600
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   8520
         TabIndex        =   65
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   79888385
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   4680
         TabIndex        =   66
         Top             =   585
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmEmpIncreaseSalaries.frx":686E
         Height          =   315
         Left            =   4680
         TabIndex        =   67
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
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   9960
         TabIndex        =   71
         Top             =   255
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   285
         Index           =   3
         Left            =   12510
         TabIndex        =   70
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÿ·»"
         Height          =   285
         Index           =   4
         Left            =   12510
         TabIndex        =   69
         Top             =   270
         Width           =   1245
      End
      Begin VB.Label lblbr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   14850
      TabIndex        =   58
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
      TabIndex        =   57
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   14190
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   15150
      TabIndex        =   55
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   14190
      TabIndex        =   54
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
      Width           =   13845
      _cx             =   24421
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
      Caption         =   "  “Ì«œ… —« » ·„ÊŸ›    "
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
         ButtonImage     =   "FrmEmpIncreaseSalaries.frx":6883
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
         TabIndex        =   2
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
         ButtonImage     =   "FrmEmpIncreaseSalaries.frx":6C1D
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
         TabIndex        =   3
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
         ButtonImage     =   "FrmEmpIncreaseSalaries.frx":6FB7
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
         TabIndex        =   4
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
         ButtonImage     =   "FrmEmpIncreaseSalaries.frx":7351
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
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   5880
         Picture         =   "FrmEmpIncreaseSalaries.frx":76EB
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
      Left            =   30
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8580
      Width           =   13665
      _cx             =   24104
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
         Left            =   12390
         TabIndex        =   6
         Top             =   75
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
         Left            =   11055
         TabIndex        =   7
         Top             =   75
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
         Left            =   9615
         TabIndex        =   8
         Top             =   75
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
         Left            =   8280
         TabIndex        =   9
         Top             =   75
         Width           =   1245
         _ExtentX        =   2196
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
         Left            =   6825
         TabIndex        =   10
         Top             =   75
         Width           =   1245
         _ExtentX        =   2196
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
         Left            =   120
         TabIndex        =   11
         Top             =   60
         Width           =   1365
         _ExtentX        =   2408
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
         Left            =   1935
         TabIndex        =   12
         Top             =   60
         Width           =   1395
         _ExtentX        =   2461
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
         Left            =   5280
         TabIndex        =   19
         Top             =   60
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
         Left            =   3600
         TabIndex        =   22
         Top             =   60
         Width           =   1485
         _ExtentX        =   2619
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
         ColorTextShadow =   -2147483637
      End
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   8820
      TabIndex        =   13
      Top             =   8280
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   6615
      Left            =   0
      TabIndex        =   23
      Top             =   1560
      Width           =   13800
      _cx             =   24342
      _cy             =   11668
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
      Caption         =   " ⁄œÌ· «·—« »|Õ«·Â «·«⁄ „«œ|«·Ê÷⁄ «·„ﬁ —Õ|≈Ÿ«›… „›—œ"
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
      Picture(0)      =   "FrmEmpIncreaseSalaries.frx":B353
      Flags(1)        =   2
      Picture(2)      =   "FrmEmpIncreaseSalaries.frx":11BB5
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   6150
         Left            =   14745
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   45
         Width           =   13710
         _cx             =   24183
         _cy             =   10848
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   2175
            Left            =   0
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   -120
            Width           =   13695
            _cx             =   24156
            _cy             =   3836
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
            Begin VB.ComboBox ContractUPdata 
               Height          =   315
               ItemData        =   "FrmEmpIncreaseSalaries.frx":18417
               Left            =   10080
               List            =   "FrmEmpIncreaseSalaries.frx":18419
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   124
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox TxtRemarkUPdata 
               Alignment       =   2  'Center
               Height          =   1035
               Left            =   720
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   123
               Top             =   960
               Width           =   10695
            End
            Begin MSDataListLib.DataCombo JobUPdata 
               Height          =   315
               Left            =   8220
               TabIndex        =   125
               Top             =   600
               Width           =   3195
               _ExtentX        =   5636
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·⁄ﬁœ"
               Height          =   255
               Left            =   11790
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”»«» √Œ—Ï ·· —ﬁÌ…"
               Height          =   435
               Index           =   36
               Left            =   11640
               TabIndex        =   127
               Top             =   960
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÊŸÌ›…"
               Height          =   195
               Index           =   37
               Left            =   11760
               TabIndex        =   126
               Top             =   600
               Width           =   1365
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   4095
            Left            =   0
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   2040
            Width           =   13695
            _cx             =   24156
            _cy             =   7223
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
            Begin VB.Image Image1 
               Height          =   3015
               Left            =   240
               Picture         =   "FrmEmpIncreaseSalaries.frx":1841B
               Stretch         =   -1  'True
               Top             =   360
               Width           =   13095
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6150
         Left            =   14445
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   45
         Width           =   13710
         _cx             =   24183
         _cy             =   10848
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
            TabIndex        =   25
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
            FormatString    =   $"FrmEmpIncreaseSalaries.frx":1DBED
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
            TabIndex        =   39
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
            TabIndex        =   26
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6150
         Index           =   15
         Left            =   45
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   45
         Width           =   13710
         _cx             =   24183
         _cy             =   10848
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
         _GridInfo       =   $"FrmEmpIncreaseSalaries.frx":1DD39
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6120
            Index           =   16
            Left            =   15
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   15
            Width           =   13680
            _cx             =   24130
            _cy             =   10795
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
            Begin VB.Frame Frame5 
               BackColor       =   &H00E2E9E9&
               Height          =   3555
               Left            =   135
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   2460
               Width           =   5985
               Begin VB.TextBox TxtRemark 
                  Alignment       =   2  'Center
                  Height          =   555
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   112
                  Top             =   2880
                  Width           =   4815
               End
               Begin VB.CheckBox Approved 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " „ «·«⁄ „«œ"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.TextBox TxtRemarkAccount 
                  Alignment       =   2  'Center
                  Height          =   555
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   100
                  Top             =   120
                  Width           =   4215
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg2 
                  Height          =   1365
                  Left            =   120
                  TabIndex        =   94
                  Top             =   1200
                  Width           =   5805
                  _cx             =   10239
                  _cy             =   2408
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
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmEmpIncreaseSalaries.frx":1DD6D
                  ScrollTrack     =   -1  'True
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   8
                  Left            =   4680
                  TabIndex        =   108
                  Top             =   2520
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   476
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
                  ButtonImage     =   "FrmEmpIncreaseSalaries.frx":1DE6B
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSComCtl2.DTPicker DateIncrease 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   109
                  Top             =   720
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   79888385
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   285
                  Index           =   34
                  Left            =   5040
                  TabIndex        =   111
                  Top             =   3000
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·“Ì«œ… «·„⁄ „œ…"
                  Height          =   285
                  Index           =   33
                  Left            =   1680
                  TabIndex        =   110
                  Top             =   720
                  Width           =   1845
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·“Ì«œ… «·„⁄ „œ…"
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   26
                  Left            =   2880
                  TabIndex        =   103
                  Top             =   960
                  Width           =   1845
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ«  «·«œ«—… «·„«·Ì…"
                  Height          =   405
                  Index           =   32
                  Left            =   4320
                  TabIndex        =   102
                  Top             =   120
                  Width           =   1605
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Height          =   3600
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   2460
               Width           =   7560
               Begin VB.ComboBox DcbType 
                  Height          =   315
                  ItemData        =   "FrmEmpIncreaseSalaries.frx":1E405
                  Left            =   2640
                  List            =   "FrmEmpIncreaseSalaries.frx":1E407
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   104
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.TextBox TxtRemarkHR 
                  Alignment       =   2  'Center
                  Height          =   675
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   99
                  Top             =   2880
                  Width           =   6015
               End
               Begin VB.CheckBox ChekAll 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ÕœÌœ «·ﬂ·"
                  ForeColor       =   &H00FF0000&
                  Height          =   315
                  Left            =   5880
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   2520
                  Width           =   1095
               End
               Begin VB.TextBox TxtTypeValu 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1560
                  MultiLine       =   -1  'True
                  TabIndex        =   96
                  Top             =   360
                  Width           =   1005
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg1 
                  Height          =   1725
                  Left            =   120
                  TabIndex        =   92
                  Top             =   720
                  Width           =   7365
                  _cx             =   12991
                  _cy             =   3043
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
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmEmpIncreaseSalaries.frx":1E409
                  ScrollTrack     =   -1  'True
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
               Begin ImpulseButton.ISButton Cmdd 
                  Height          =   510
                  Left            =   240
                  TabIndex        =   106
                  Top             =   240
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   900
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
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
                  ButtonImage     =   "FrmEmpIncreaseSalaries.frx":1E575
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  Height          =   285
                  Index           =   2
                  Left            =   720
                  TabIndex        =   107
                  Top             =   360
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ«  ‘ƒÊ‰ «·„ÊŸ›Ì‰"
                  Height          =   645
                  Index           =   28
                  Left            =   6240
                  TabIndex        =   101
                  Top             =   2880
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»œ·«  «·œ«Œ·Â ›Ì «·“Ì«œ…"
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   22
                  Left            =   5640
                  TabIndex        =   97
                  Top             =   360
                  Width           =   1845
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·“Ì«œ…"
                  Height          =   285
                  Index           =   21
                  Left            =   3960
                  TabIndex        =   95
                  Top             =   360
                  Width           =   1005
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Height          =   720
               Left            =   135
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   585
               Width           =   13545
               Begin VB.TextBox TxtRemarkManger 
                  Alignment       =   2  'Center
                  Height          =   555
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   89
                  Top             =   120
                  Width           =   4095
               End
               Begin VB.TextBox TxtRemarkEmp 
                  Alignment       =   2  'Center
                  Height          =   555
                  Left            =   6720
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   88
                  Top             =   120
                  Width           =   4455
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "(«”»«» «·“Ì«œ…(Œ«’ »«·„œÌ— «·„»«‘—"
                  Height          =   435
                  Index           =   19
                  Left            =   4200
                  TabIndex        =   90
                  Top             =   240
                  Width           =   2445
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "(«”»«» «·“Ì«œ…(Œ«’ »«·„ÊŸ›"
                  Height          =   435
                  Index           =   20
                  Left            =   11280
                  TabIndex        =   87
                  Top             =   240
                  Width           =   2085
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Height          =   1395
               Left            =   135
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   1185
               Width           =   13545
               Begin VB.ComboBox DataCombo5 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "FrmEmpIncreaseSalaries.frx":24DD7
                  Left            =   10200
                  List            =   "FrmEmpIncreaseSalaries.frx":24DD9
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   118
                  Top             =   960
                  Width           =   1455
               End
               Begin VB.TextBox TxtLastUpdateSalary 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   6120
                  MultiLine       =   -1  'True
                  TabIndex        =   82
                  Top             =   600
                  Width           =   1935
               End
               Begin MSComCtl2.DTPicker BignDate 
                  Height          =   315
                  Left            =   10200
                  TabIndex        =   77
                  Top             =   240
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   79888385
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker LastUpdateDate 
                  Height          =   315
                  Left            =   10200
                  TabIndex        =   78
                  Top             =   600
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   79888385
                  CurrentDate     =   38784
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   1005
                  Left            =   0
                  TabIndex        =   85
                  Top             =   120
                  Width           =   6045
                  _cx             =   10663
                  _cy             =   1773
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
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmEmpIncreaseSalaries.frx":24DDB
                  ScrollTrack     =   -1  'True
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
               Begin VB.Label Label30 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·⁄ﬁœ"
                  Height          =   255
                  Left            =   11640
                  RightToLeft     =   -1  'True
                  TabIndex        =   119
                  Top             =   960
                  Width           =   1695
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   13
                  Left            =   6360
                  TabIndex        =   84
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—« » «·Õ«·Ì"
                  Height          =   285
                  Index           =   12
                  Left            =   8520
                  TabIndex        =   83
                  Top             =   240
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„»·€ «Œ— “Ì«œ… ··—« »"
                  Height          =   525
                  Index           =   11
                  Left            =   8280
                  TabIndex        =   81
                  Top             =   600
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «Œ— “Ì«œ… ··—« »"
                  Height          =   525
                  Index           =   10
                  Left            =   11640
                  TabIndex        =   80
                  Top             =   600
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì…«·⁄„·"
                  Height          =   285
                  Index           =   9
                  Left            =   11880
                  TabIndex        =   79
                  Top             =   240
                  Width           =   1485
               End
            End
            Begin VB.Frame lbDW 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·„ÊŸ›"
               Height          =   675
               Left            =   135
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   0
               Width           =   13545
               Begin MSDataListLib.DataCombo DcboJobsType 
                  Height          =   315
                  Left            =   9360
                  TabIndex        =   43
                  Top             =   240
                  Width           =   3195
                  _ExtentX        =   5636
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcProject 
                  Height          =   315
                  Left            =   4920
                  TabIndex        =   44
                  Top             =   240
                  Width           =   3195
                  _ExtentX        =   5636
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbDepartment 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   74
                  Top             =   240
                  Width           =   3195
                  _ExtentX        =   5636
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
                  Height          =   405
                  Index           =   0
                  Left            =   3480
                  TabIndex        =   75
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊŸÌ›…"
                  Height          =   195
                  Index           =   24
                  Left            =   12480
                  TabIndex        =   46
                  Top             =   240
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„‘—Ê⁄"
                  Height          =   405
                  Index           =   15
                  Left            =   7920
                  TabIndex        =   45
                  Top             =   240
                  Width           =   1005
               End
            End
            Begin VB.Frame lblds 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·—« »"
               Height          =   3285
               Left            =   900
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   6105
               Width           =   6795
               Begin VB.TextBox Txtincrease 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   825
                  Left            =   3360
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   72
                  Top             =   480
                  Width           =   2145
               End
               Begin VB.TextBox txtab 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   465
                  Left            =   120
                  TabIndex        =   49
                  Top             =   1440
                  Width           =   2145
               End
               Begin VB.TextBox txtadd 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   465
                  Left            =   3360
                  MultiLine       =   -1  'True
                  TabIndex        =   48
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
                  TabIndex        =   47
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
                  TabIndex        =   53
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
                  TabIndex        =   52
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
                  TabIndex        =   51
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
                  TabIndex        =   50
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
               Height          =   645
               Left            =   270
               TabIndex        =   38
               Top             =   5955
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   1138
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3540
               Index           =   62
               Left            =   2625
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   1635
               Width           =   585
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6120
            Index           =   9
            Left            =   15
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   15
            Width           =   13680
            _cx             =   24130
            _cy             =   10795
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
               Height          =   4590
               Left            =   3585
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   1320
               Width           =   720
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   3165
               Left            =   4530
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   1635
               Width           =   1185
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3165
               Index           =   67
               Left            =   2535
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   1635
               Width           =   675
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   3060
               Index           =   68
               Left            =   4305
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   2085
               Width           =   45
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
               Height          =   3690
               Index           =   69
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   1635
               Width           =   375
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   6150
         Left            =   15045
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   45
         Width           =   13710
         _cx             =   24183
         _cy             =   10848
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   4860
            Left            =   3480
            TabIndex        =   131
            Top             =   240
            Width           =   10080
            _cx             =   17780
            _cy             =   8572
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
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   15
            FixedRows       =   2
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEmpIncreaseSalaries.frx":24E7F
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
         End
      End
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   14910
      TabIndex        =   59
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
      Left            =   15270
      TabIndex        =   60
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
      Left            =   14520
      RightToLeft     =   -1  'True
      TabIndex        =   61
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
      Top             =   8355
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
      Top             =   8310
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
      Top             =   8310
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   330
      TabIndex        =   15
      Top             =   8340
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1980
      TabIndex        =   14
      Top             =   8340
      Width           =   615
   End
End
Attribute VB_Name = "FrmEmpIncreaseSalaries"
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
Public LngRow As Double
Public LngCol  As Double

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
  IntCounter = 0
    With FG2

        For i = .FixedRows To .Rows - 1
IntCounter = IntCounter + 1
       
 If val(.TextMatrix(i, .ColIndex("MofradID"))) <> 0 Then
                
         .TextMatrix(i, .ColIndex("Count")) = IntCounter
            End If
        Next i
 
    End With
    


End Sub

Private Sub Approved_Click()
If Me.TxtModFlg.Text <> "R" Then
FillGraidApproved
End If
End Sub

Private Sub ChekAll_Click()
relighn
End Sub
Sub FillWithIncreaseValue()
Dim i As Integer
If DcbType.ListIndex = -1 Then
If SystemOptions.UserInterface Then
MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄ «·“Ì«œ…"
Else
MsgBox "Please select Type of Increase"
End If
DcbType.SetFocus
Exit Sub
End If
If val(TxtTypeValu.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  ÕœÌœ ﬁÌ„Â «·“Ì«œ…"
Else
MsgBox "Please Select Value"
End If
TxtTypeValu.SetFocus
Exit Sub
End If
With FG1
     For i = .FixedRows To .Rows - 1
     If .Cell(flexcpChecked, i, .ColIndex("Chek")) = flexChecked Then
     .TextMatrix(i, .ColIndex("Typeincrease")) = val(DcbType.ListIndex) + 1
    .TextMatrix(i, .ColIndex("TypeValue")) = val(TxtTypeValu.Text)
     If DcbType.ListIndex = 1 Then
      .TextMatrix(i, .ColIndex("IncreaseValue")) = (val(.TextMatrix(i, .ColIndex("CurrValue"))) * val(TxtTypeValu.Text) / 100)
 Else
    .TextMatrix(i, .ColIndex("IncreaseValue")) = val(TxtTypeValu.Text)
    .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("CurrValue"))) + val(.TextMatrix(i, .ColIndex("IncreaseValue")))
      End If
     End If
      Next i
      End With
End Sub
Sub FillGraidApproved()
Dim i As Integer
FG2.Rows = 2
With FG1
     For i = .FixedRows To .Rows - 1
     If .Cell(flexcpChecked, i, .ColIndex("Chek")) = flexChecked Then
     FG2.TextMatrix(FG2.Rows - 1, FG2.ColIndex("MofradID")) = .TextMatrix(i, .ColIndex("MofradID"))
     FG2.TextMatrix(FG2.Rows - 1, FG2.ColIndex("name")) = .TextMatrix(i, .ColIndex("name"))
     FG2.TextMatrix(FG2.Rows - 1, FG2.ColIndex("CurrValue")) = .TextMatrix(i, .ColIndex("CurrValue"))
     FG2.TextMatrix(FG2.Rows - 1, FG2.ColIndex("IncreaseValue")) = .TextMatrix(i, .ColIndex("IncreaseValue"))
     FG2.TextMatrix(FG2.Rows - 1, FG2.ColIndex("total")) = .TextMatrix(i, .ColIndex("total"))
     FG2.Rows = FG2.Rows + 1
      End If
      Next i
      End With
End Sub

Sub relighn()
Dim i As Integer
With FG1
     For i = .FixedRows To .Rows - 1
                                   If ChekAll.value = vbChecked Then
                .TextMatrix(i, .ColIndex("Chek")) = -1
                Else
                .TextMatrix(i, .ColIndex("Chek")) = 0
               
      End If
      Next i
      End With
End Sub
'Private Sub Accredit_Click()
'    Dim BeginTrans As Boolean
'
'    Cn.BeginTrans
'    BeginTrans = True
'
'    If IsNull(rs("Posted")) Then
'        rs("Posted") = user_id
'        rs("PostedDate") = Time
'    Else
'        rs("Posted") = Null
'       rs("PostedDate") = Time
'    End If
'
'    rs.update
' If SystemOptions.UserInterface = ArabicInterface Then
'    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
'Else
'Accredit.Caption = "Sent To approval "
'End If
'
'    Cn.CommitTrans
'    BeginTrans = False
'FillApprovedTable
'    Retrive (val(Me.XPTxtID.text))
'End Sub
'
Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
       
            lbl(13).Caption = "0"
            lbl(23).Caption = "0"
            GRID2.Clear flexClearScrollable, flexClearEverything
              GRID2.Rows = 1
              VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
              VSFlexGrid1.Rows = 2
            Me.DCboUserName.BoundText = user_id
        
dcBranch.BoundText = Current_branch
          
            
            Accredit.Enabled = True
                If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
            
             
        Case 1
                                   If ChekClodePeriod(XPDtbTrans.value) = True Then
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
            VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
            Me.DCboUserName.BoundText = user_id
            DateIncrease.value = Date

        Case 2
    
                                   If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                      
            Dim Msg As String

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                dcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText

            SaveData

        Case 3
            Undo

        Case 4
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

            Del_Trans

        Case 5
       'wael
          Load FrmSaerchIncreaseSalary
           FrmSaerchIncreaseSalary.show vbModal

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.Text, , 200

        Case 8
         
     RemoveGridRow1
           
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
Private Sub RemoveGridRow1()

    With Me.FG2

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
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
'If Me.CBTybe.ListIndex = 2 Then
    MySQL = " SELECT   dbo.TblEmpIncreaseSalary.ID, dbo.TblEmpIncreaseSalary.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmpIncreaseSalary.EmpID,"
    MySQL = MySQL & "   dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmpIncreaseSalary.JobID, dbo.TblEmpIncreaseSalary.DeptID,"
    MySQL = MySQL & "   dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpIncreaseSalary.RecordDate, dbo.TblEmpIncreaseSalary.BignDate,"
    MySQL = MySQL & "   dbo.TblEmpIncreaseSalary.LastUpdateDate, dbo.TblEmpIncreaseSalary.CurSalary, dbo.TblEmpIncreaseSalary.LastUpdateSalary, dbo.TblEmpIncreaseSalary.TypeValu,"
    MySQL = MySQL & "  dbo.TblEmpIncreaseSalary.RemarkEmp, dbo.TblEmpIncreaseSalary.RemarkManger, dbo.TblEmpIncreaseSalary.RemarkHR, dbo.TblEmpIncreaseSalary.RemarkAccount,"
    MySQL = MySQL & "   dbo.TblEmpIncreaseSalary.ChekAll, dbo.TblEmpIncreaseSalary.TypeIncrease, dbo.TblEmpIncreaseSalary.Approved, dbo.TblEmpIncreaseSalary.ProjectID, dbo.EmpGroupDep.GroupName,"
    MySQL = MySQL & "   dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpIncreaseSalaryDetalis.TypeValue, dbo.TblEmpIncreaseSalaryDetalis.Typeincrease AS TypeincreaseDet,"
    MySQL = MySQL & "  dbo.TblEmpIncreaseSalaryDetalis.IncreaseValue, dbo.TblEmpIncreaseSalaryDetalis.TypeID, dbo.TblEmpIncreaseSalaryDetalis.Chek, dbo.TblEmpIncreaseSalaryDetalis.IDIncrease,"
    MySQL = MySQL & "  dbo.TblEmpIncreaseSalaryDetalis.MofradID, dbo.mofrad.name, dbo.mofrad.nameE, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3,"
    MySQL = MySQL & "  dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality,"
    MySQL = MySQL & "  dbo.TblEmpIncreaseSalaryDetalis.CurrValue, dbo.TblEmpIncreaseSalary.DateIncrease, dbo.TblEmpIncreaseSalary.Remark, dbo.TblEmpIncreaseSalary.AddTybe,"
    MySQL = MySQL & "  dbo.TblEmpIncreaseSalary.ADDtype_Contract, dbo.Contract.Contract_Enddate, dbo.TblEmpIncreaseSalary.ContractUPdata, dbo.TblEmpIncreaseSalary.JobUPdata,"
    MySQL = MySQL & "   dbo.TblEmpIncreaseSalary.TxtRemarkUPdata, TblEmpJobsTypes_1.JobTypeName AS JobTypeNameUPDATA, TblEmpJobsTypes_1.JobTypeNamee AS JobTypeNameeUPDATA , dbo.mofrad.id AS IDMofrD"
    MySQL = MySQL & "        FROM         dbo.TblEmpDepartments RIGHT OUTER JOIN"
    MySQL = MySQL & "   dbo.TblEmpJobsTypes TblEmpJobsTypes_1 LEFT OUTER JOIN"
    MySQL = MySQL & "  dbo.TblEmpIncreaseSalary ON TblEmpJobsTypes_1.JobTypeID = dbo.TblEmpIncreaseSalary.JobUPdata LEFT OUTER JOIN"
    MySQL = MySQL & "  dbo.mofrad RIGHT OUTER JOIN"
    MySQL = MySQL & "   dbo.TblEmpIncreaseSalaryDetalis ON dbo.mofrad.id = dbo.TblEmpIncreaseSalaryDetalis.MofradID ON"
    MySQL = MySQL & "   dbo.TblEmpIncreaseSalary.ID = dbo.TblEmpIncreaseSalaryDetalis.IDIncrease LEFT OUTER JOIN"
    MySQL = MySQL & "  dbo.TblEmpJobsTypes ON dbo.TblEmpIncreaseSalary.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
    MySQL = MySQL & "  dbo.EmpGroupDep ON dbo.TblEmpIncreaseSalary.ProjectID = dbo.EmpGroupDep.GroupID ON dbo.TblEmpDepartments.DeparmentID = dbo.TblEmpIncreaseSalary.DeptID LEFT OUTER JOIN"
    MySQL = MySQL & "   dbo.Contract  RIGHT OUTER JOIN "
    MySQL = MySQL & "  dbo.TblEmployee ON dbo.Contract.Emp_id = dbo.TblEmployee.Emp_ID ON dbo.TblEmpIncreaseSalary.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & "   dbo.TblBranchesData ON dbo.TblEmpIncreaseSalary.BranchID = dbo.TblBranchesData.branch_id"
    MySQL = MySQL & "  Where (dbo.TblEmpIncreaseSalary.id = " & val(XPTxtID.Text) & ")"
  '  Else
    
  ' End If
  MySQL = "SELECT        dbo.TblEmpIncreaseSalary.ID, dbo.TblEmpIncreaseSalary.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmpIncreaseSalary.EmpID, dbo.TblEmployee.Emp_Name, "
MySQL = MySQL & "                           dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmpIncreaseSalary.JobID, dbo.TblEmpIncreaseSalary.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
MySQL = MySQL & "                           dbo.TblEmpIncreaseSalary.RecordDate, dbo.TblEmpIncreaseSalary.BignDate, dbo.TblEmpIncreaseSalary.LastUpdateDate, dbo.TblEmpIncreaseSalary.CurSalary, dbo.TblEmpIncreaseSalary.LastUpdateSalary,"
MySQL = MySQL & "                           dbo.TblEmpIncreaseSalary.TypeValu, dbo.TblEmpIncreaseSalary.RemarkEmp, dbo.TblEmpIncreaseSalary.RemarkManger, dbo.TblEmpIncreaseSalary.RemarkHR, dbo.TblEmpIncreaseSalary.RemarkAccount,"
MySQL = MySQL & "                           dbo.TblEmpIncreaseSalary.ChekAll, dbo.TblEmpIncreaseSalary.TypeIncrease, dbo.TblEmpIncreaseSalary.Approved, dbo.TblEmpIncreaseSalary.ProjectID, dbo.EmpGroupDep.GroupName, dbo.TblEmpJobsTypes.JobTypeName,"
MySQL = MySQL & "                           dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpIncreaseSalaryDetalis.TypeValue, dbo.TblEmpIncreaseSalaryDetalis.Typeincrease AS TypeincreaseDet, dbo.TblEmpIncreaseSalaryDetalis.IncreaseValue,"
MySQL = MySQL & "                           dbo.TblEmpIncreaseSalaryDetalis.TypeID, dbo.TblEmpIncreaseSalaryDetalis.Chek, dbo.TblEmpIncreaseSalaryDetalis.IDIncrease, dbo.TblEmpIncreaseSalaryDetalis.MofradID, dbo.mofrad.name, dbo.mofrad.nameE,"
MySQL = MySQL & "                           dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                           dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmpIncreaseSalaryDetalis.CurrValue, dbo.TblEmpIncreaseSalary.DateIncrease, dbo.TblEmpIncreaseSalary.Remark,"
MySQL = MySQL & "                           dbo.TblEmpIncreaseSalary.AddTybe, dbo.TblEmpIncreaseSalary.ADDtype_Contract, dbo.Contract.Contract_Enddate, dbo.TblEmpIncreaseSalary.ContractUPdata, dbo.TblEmpIncreaseSalary.JobUPdata,"
MySQL = MySQL & "                           dbo.TblEmpIncreaseSalary.TxtRemarkUPdata, TblEmpJobsTypes_1.JobTypeName AS JobTypeNameUPDATA, TblEmpJobsTypes_1.JobTypeNamee AS JobTypeNameeUPDATA, dbo.mofrad.id AS IDMofrD"
MySQL = MySQL & "  FROM            dbo.TblEmpDepartments RIGHT OUTER JOIN"
  MySQL = MySQL & "                         dbo.TblEmpJobsTypes AS TblEmpJobsTypes_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                           dbo.TblEmpIncreaseSalary ON TblEmpJobsTypes_1.JobTypeID = dbo.TblEmpIncreaseSalary.JobUPdata LEFT OUTER JOIN"
MySQL = MySQL & "                           dbo.mofrad RIGHT OUTER JOIN"
MySQL = MySQL & "                           dbo.TblEmpIncreaseSalaryDetalis ON dbo.mofrad.id = dbo.TblEmpIncreaseSalaryDetalis.MofradID ON dbo.TblEmpIncreaseSalary.ID = dbo.TblEmpIncreaseSalaryDetalis.IDIncrease LEFT OUTER JOIN"
MySQL = MySQL & "                           dbo.TblEmpJobsTypes ON dbo.TblEmpIncreaseSalary.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                           dbo.EmpGroupDep ON dbo.TblEmpIncreaseSalary.ProjectID = dbo.EmpGroupDep.GroupID ON dbo.TblEmpDepartments.DeparmentID = dbo.TblEmpIncreaseSalary.DeptID "

MySQL = MySQL & "                                  LEFT OUTER JOIN dbo.TblEmployee "
MySQL = MySQL & "                                       ON  dbo.TblEmpIncreaseSalary.EmpID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & "                                  LEFT OUTER JOIN dbo.Contract "
MySQL = MySQL & "                                       ON  dbo.Contract.Emp_id = dbo.TblEmployee.Emp_ID "

MySQL = MySQL & "                           LEFT OUTER JOIN "
MySQL = MySQL & "                           dbo.TblBranchesData ON dbo.TblEmpIncreaseSalary.BranchID = dbo.TblBranchesData.branch_id"
    MySQL = MySQL & "  Where (dbo.TblEmpIncreaseSalary.id = " & val(XPTxtID.Text) & ")"

     If Me.CBTybe.ListIndex = 2 Then
      If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepIncreaseSalary.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepIncreaseSalary.rpt"
        End If
    Else
      If SystemOptions.UserInterface = ArabicInterface Then
           StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepIncreaseSalaryUPdate.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepIncreaseSalaryUPdate.rpt"
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
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function


Private Sub Cmdd_Click()
FillWithIncreaseValue
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub
Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
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
        Dim MaxDate As Date
        Dim maxSala As Double
         '''''''''
        Dim ADDtype_Contract As Integer
        '''''''''''''''''''''''''''''''

        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth, national, , , ProjectID, , , , , endiqama, , BignDateWork, , JobTypeName, Contract_period, Contract_periodno, , dcjopstatus, JobTypeIDIQ, , , , , , , , , , , , , , , , , ADDtype_Contract
          

        DcbDepartment.BoundText = DepID
        DCproject.BoundText = ProjectID
        BignDate.value = BignDateWork
        DcboJobsType.BoundText = JobTypeID
       DataCombo5.ListIndex = val(ADDtype_Contract)
        lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "", 0)
         lbl(13).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "", 0)

  GetEmployeeSalaryAccordingToComponentEndservice val(DcboEmpName.BoundText)
GetMaxSalDate val(DcboEmpName.BoundText), MaxDate
GetMaxSalary val(DcboEmpName.BoundText), maxSala, MaxDate
TxtLastUpdateSalary.Text = maxSala
LastUpdateDate.value = MaxDate

End Sub

 Sub GetEmployeeSalaryAccordingToComponentEndservice(Emp_id As Integer)
                                                    
    Dim sql As String
    Dim mofrad_name As String
    Dim valuee As Double
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim Mofradd As String
    Dim i As Integer

sql = " SELECT     dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, SUM(dbo.EmpSalaryComponent.[Value]) AS SmValue, dbo.mofrdat.mofrad_type"
sql = sql & " FROM         dbo.mofrad INNER JOIN"
sql = sql & "                      dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type INNER JOIN"
sql = sql & "                      dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode"
sql = sql & " GROUP BY dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.mofrdat.mofrad_type, dbo.EmpSalaryComponent.emp_ID"
sql = sql & " Having (dbo.EmpSalaryComponent.Emp_id = " & Emp_id & ")"
sql = sql & " ORDER BY dbo.mofrdat.mofrad_type"

      rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
  With Me.FG
  .Rows = rs.RecordCount + 1
      For i = 1 To rs.RecordCount
       .TextMatrix(i, .ColIndex("Count")) = i
      .TextMatrix(i, .ColIndex("MofradID")) = IIf(IsNull(rs("mofrad_type").value), 0, rs("mofrad_type").value)
       .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
 .TextMatrix(i, .ColIndex("CurrValue")) = IIf(IsNull(rs("SmValue").value), 0, rs("SmValue").value)
  
 rs.MoveNext
      Next i
 End With
     End If
     rs.Close
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
  With Me.FG1
  .Rows = rs.RecordCount + 1
      For i = 1 To rs.RecordCount
       .TextMatrix(i, .ColIndex("Count")) = i
      .TextMatrix(i, .ColIndex("MofradID")) = IIf(IsNull(rs("mofrad_type").value), 0, rs("mofrad_type").value)
       .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
 .TextMatrix(i, .ColIndex("CurrValue")) = IIf(IsNull(rs("SmValue").value), 0, rs("SmValue").value)
  
 rs.MoveNext
      Next i
 End With
     End If
        
     
      rs.Close
  '  ReLineGrid
End Sub

Sub GetMaxSalary(Optional EmpID As Integer = 0, Optional ByRef Sala As Double, Optional MaxDate As Date)
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
  sql = "SELECT     value from dbo.EmpSalaryComponent WHERE     (emp_ID = " & EmpID & ") and (EntIncresDataM =" & SQLDate(MaxDate, True) & ")"
     rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

 Sala = IIf(IsNull(rs("value").value), 0, rs("value").value)
End If

End Sub
Sub GetMaxSalDate(Optional EmpID As Integer = 0, Optional ByRef MaxDate As Date)
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
  sql = "SELECT     MAX(EntIncresDataM) AS MaxDate  from dbo.EmpSalaryComponent WHERE     (emp_ID = " & EmpID & ")"
     rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    If Not IsNull(rs("MaxDate").value) Then
 MaxDate = IIf(IsNull(rs("MaxDate").value), Date, rs("MaxDate").value)

Else
 MaxDate = Date

 End If
 Else
MaxDate = Date
    End If

End Sub
Private Sub DcboEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 25
        Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If
End Sub

Private Sub Fg1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With FG1

        Select Case .ColKey(Col)
              Case "Typeincrease"
             If .Cell(flexcpChecked, Row, .ColIndex("Chek")) = flexUnchecked Then
    
              .TextMatrix(Row, .ColIndex("Typeincrease")) = ""
              Exit Sub
              End If
              If val(.TextMatrix(Row, .ColIndex("Typeincrease"))) = 2 Then
              MsgBox .TextMatrix(Row, .ColIndex("IncreaseValue"))
              MsgBox (val(.TextMatrix(Row, .ColIndex("CurrValue"))) * val(.TextMatrix(Row, .ColIndex("TypeValue"))) / 100)
              .TextMatrix(Row, .ColIndex("IncreaseValue")) = (val(.TextMatrix(Row, .ColIndex("CurrValue"))) * val(.TextMatrix(Row, .ColIndex("TypeValue"))) / 100)
              Else
              .TextMatrix(Row, .ColIndex("IncreaseValue")) = val(.TextMatrix(Row, .ColIndex("TypeValue")))
              End If
                Case "TypeValue"
                 If val(.TextMatrix(Row, .ColIndex("Typeincrease"))) = 2 Or val(.TextMatrix(Row, .ColIndex("Typeincrease"))) = 1 Then
              If val(.TextMatrix(Row, .ColIndex("Typeincrease"))) = 2 Then
              .TextMatrix(Row, .ColIndex("IncreaseValue")) = (val(.TextMatrix(Row, .ColIndex("CurrValue"))) * val(.TextMatrix(Row, .ColIndex("TypeValue"))) / 100)
              Else
              .TextMatrix(Row, .ColIndex("IncreaseValue")) = val(.TextMatrix(Row, .ColIndex("TypeValue")))
              
              End If
              .TextMatrix(Row, .ColIndex("total")) = val(.TextMatrix(Row, .ColIndex("CurrValue"))) + val(.TextMatrix(Row, .ColIndex("IncreaseValue")))
            Else
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ÌÃ»  ÕœÌœ ‰Ê⁄ «·“Ì«œÂ «Ê·«"
            Else
            MsgBox "Please select Type of Incncrease"
            End If
            Exit Sub
            End If
       End Select
       End With
 
End Sub



Private Sub Fg1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With FG1

        Select Case .ColKey(Col)
              Case "Typeincrease"
              If .Cell(flexcpChecked, Row, .ColIndex("Chek")) = flexUnchecked Then
              If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "Ì—ÃÏ  ÕœÌœ «·„›—œ «Ê·«"
              Else
              MsgBox "Please select Mofrd "
              End If
              .TextMatrix(Row, .ColIndex("Typeincrease")) = ""
              Exit Sub
              End If
          End Select
      End With
End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim GrdBack As ClsBackGroundPic
    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
    Dim StrSQL As String

   If SystemOptions.UserInterface = EnglishInterface Then
      DcbType.AddItem "Value"
      DcbType.AddItem "Rate"
    Else
      DcbType.AddItem "ﬁÌ„…"
      DcbType.AddItem "‰”»Â"
    End If
     If SystemOptions.UserInterface = ArabicInterface Then
                FG1.ColComboList(FG1.ColIndex("Typeincrease")) = "#1; ﬁÌ„… |#2; ‰”»Â"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               FG1.ColComboList(FG1.ColIndex("Typeincrease")) = "#1; value |#2;Rate "
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
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetEmpLocations Me.DCproject
    Dcombos.GetEmpJobsTypes Me.DcboJobsType
    Dcombos.GetEmpDepartments Me.DcbDepartment
    Dcombos.GetEmpJobsTypes Me.JobUPdata
   ' XPTab301.CurrTab = 0
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
  '  YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEmpIncreaseSalary     Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.Text = "R"
    FullComb
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
    lbl(35).Caption = "Increase Type"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    Label2.Caption = "Contract Type"
    lbl(37).Caption = "New Job"
    lbl(36).Caption = "Notice"
   XPTab301.CurrTab = 0
 
    XPTab301.Caption = "salary adjustment |Approve|Proposed Situation |Add Componenet"
'      XPTab301.TabVisible(0) = False
'   XPTab301.TabVisible(0) = True
'   XPTab301.TabVisible(0) = True
  '  XPTab301.CurrTab = 1
    'XPTab301.Caption(0) = "Data"
    'XPTab301.Caption(2) = "Proposed Situation"
    Label30.Caption = "Contract Type"
    Me.Caption = "Increase the Salary of the Employee"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lblBr.Caption = "Branch"
    lbl(29).Caption = "Curr Salary"
    lbl(3).Caption = "Employee"
    lbDW.Caption = "Data of Employee"
    lbl(24).Caption = "Job"
    lbl(15).Caption = "Project"
   lbl(0).Caption = "Department"
  
   lbl(20).Caption = "Employee Remarks"
    lbl(19).Caption = "Direct manager Remarks"
    lbl(9).Caption = "Start Work"
    lbl(12).Caption = "Curr Salary"
    
    lbl(10).Caption = "Last salary increase date"
    lbl(11).Caption = "Last increase in the amount of salary"
   ChekAll.RightToLeft = False
   ChekAll.Caption = "Select All"
    lbl(21).Caption = "Increase value "
    
    lbl(22).Caption = "Allowances in increase"
    Cmdd.Caption = "Add"
    lbl(32).Caption = "Financial Remark"
    Approved.Caption = "Approved"
    Approved.RightToLeft = False
    lbl(26).Caption = "Accredited increase"
    lbl(33).Caption = "Date accredited increase"
    lbl(28).Caption = "HR Remarks"
   Cmd(8).Caption = "Delete"
    lbl(34).Caption = "Remarks"
  
        lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
   
    lbl(8).Caption = "By"

   With VSFlexGrid1
       .TextMatrix(0, .ColIndex("LineNo")) = "Serial"
       .TextMatrix(0, .ColIndex("AccountName")) = "Mofrd Name"
      .TextMatrix(0, .ColIndex("value")) = "Value"
      .TextMatrix(0, .ColIndex("RecoedDate")) = "Date"
    End With
    

    With FG
       .TextMatrix(0, .ColIndex("Count")) = "Serial"
       .TextMatrix(0, .ColIndex("name")) = "Mofrd Name"
        .TextMatrix(0, .ColIndex("CurrValue")) = "Current value"

    End With
     With FG2
       .TextMatrix(0, .ColIndex("Count")) = "Serial"
       .TextMatrix(0, .ColIndex("name")) = "Mofrd Name"
        .TextMatrix(0, .ColIndex("CurrValue")) = "Current value"
        .TextMatrix(0, .ColIndex("increasevalue")) = "Increase value"
        .TextMatrix(0, .ColIndex("total")) = "Value after increase"

    End With
       With FG1
       .TextMatrix(0, .ColIndex("Count")) = "Serial"
       .TextMatrix(0, .ColIndex("Chek")) = "Seletc"
       .TextMatrix(0, .ColIndex("name")) = "Mofrd Name"
        .TextMatrix(0, .ColIndex("CurrValue")) = "Current value"
        .TextMatrix(0, .ColIndex("increasevalue")) = "Increase value"
         .TextMatrix(0, .ColIndex("Typeincrease")) = "Type Increase"
        .TextMatrix(0, .ColIndex("TypeValue")) = "Value/Rate"
         .TextMatrix(0, .ColIndex("total")) = "Value after increase"

    End With

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

 

 




Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer
If Me.TxtModFlg.Text <> "R" Then
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    End If
End Sub

Private Sub TxtTypeValu_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtTypeValu.Text, 1)

End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
    On Error GoTo ErrTrap
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid1

        Select Case .ColKey(Col)
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
             If CHekMofrd(val(Me.DcboEmpName.BoundText), val(.TextMatrix(Row, .ColIndex("AccountCode")))) = False Then
                StrSQL = " SELECT     *, dbo.mofrad.name, dbo.mofrad.nameE, dbo.mofrad.AddOrDiscount, dbo.mofrad.id"
                StrSQL = StrSQL & " FROM         dbo.mofrdat INNER JOIN"
                StrSQL = StrSQL & "       dbo.mofrad ON dbo.mofrdat.mofrad_type = dbo.mofrad.id "
                StrSQL = StrSQL & "         Where mofrad_code = " & StrAccountCode
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    '.TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("eq_sys").value), "", rs("eq_sys").value)
                    
                    '.TextMatrix(Row, .ColIndex("eq_text")) = IIf(IsNull(rs("eq_text").value), "", rs("eq_text").value)
                    .TextMatrix(Row, .ColIndex("mofrad_type")) = IIf(IsNull(rs("mofrad_type").value), "", rs("mofrad_type").value)
                    '.TextMatrix(Row, .ColIndex("AddOrDiscount")) = IIf(IsNull(rs("AddOrDiscount").value), "", rs("AddOrDiscount").value)
                    
                    '.TextMatrix(Row, .ColIndex("specific_value")) = IIf(IsNull(rs("specific_value").value), "", rs("specific_value").value)
                    '.TextMatrix(Row, .ColIndex("assurance")) = IIf(IsNull(rs("assurance").value), "", rs("assurance").value)
                    '.TextMatrix(Row, .ColIndex("percentage")) = IIf(IsNull(rs("percentage").value), "", rs("percentage").value)
                    '.TextMatrix(Row, .ColIndex("min_val")) = IIf(IsNull(rs("min_val").value), "", rs("min_val").value)
                    '.TextMatrix(Row, .ColIndex("max_val")) = IIf(IsNull(rs("max_val").value), "", rs("max_val").value)
                    '.TextMatrix(Row, .ColIndex("is_fixed")) = IIf(IsNull(rs("is_fixed").value), "", rs("is_fixed").value)
                    '.TextMatrix(Row, .ColIndex("Monthly")) = IIf(IsNull(rs("Monthly").value), "", rs("Monthly").value)
                   
                End If

            Else
           If SystemOptions.UserInterface = ArabicInterface Then
           MsgBox "Ì—ÃÏ «Œ Ì«— „›—œ «Œ— Â–« «·„›—œ „ÊÃÊœ „‰ ﬁ»· "
           Else
           MsgBox "This is Component is Already Exists"
           End If
           .TextMatrix(Row, .ColIndex("AccountCode")) = 0
           .TextMatrix(Row, .ColIndex("AccountName")) = ""
           .TextMatrix(Row, .ColIndex("mofrad_type")) = 0
           Exit Sub
            End If
        End Select

    
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
ErrTrap:
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid1
Select Case .ColKey(Col)
Case "AccountName"
VSFlexGrid1.ComboList = ""
Case "RecoedDate"
VSFlexGrid1.ComboList = ""
End Select
End With
End Sub
Function CHekMofrd(Optional EmpID As Integer = 0, Optional MordCode As Integer = 0) As Boolean
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = " SELECT     emp_ID, AccountCode"
sql = sql & " From dbo.EmpSalaryComponent"
sql = sql & " Where (Emp_id = " & EmpID & ") And (AccountCode = " & MordCode & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
CHekMofrd = True
Else
CHekMofrd = False
End If
End Function
Private Sub VSFlexGrid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Me.TxtModFlg.Text <> "R" Then
With VSFlexGrid1
Select Case .ColKey(Col)


        Case "RecoedDate"
        LngRow = Row
        LngCol = Col
       
        Load FrmDateOpProject
        FrmDateOpProject.Index = 2
        FrmDateOpProject.show vbModal
   End Select
 End With
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

    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "AccountName"
                'Full Path Display
                 
                StrSQL = " select * from mofrdat "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "eq_sys, *mofrad_name", "mofrad_code")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "eq_sys, *mofrad_namee", "mofrad_code")
                End If
                
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
               Case "RecoedDate"
               .ColComboList(.ColIndex("RecoedDate")) = "..."
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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
    Dim RsDev As ADODB.Recordset
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
    DateIncrease.value = IIf(IsNull(rs("DateIncrease").value), Date, rs("DateIncrease").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    Me.DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    DcboJobsType.BoundText = IIf(IsNull(rs("JobID").value), "", rs("JobID").value)
    DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    DCproject.BoundText = IIf(IsNull(rs("ProjectID").value), "", rs("ProjectID").value)
    DcbDepartment.BoundText = IIf(IsNull(rs("DeptID").value), "", rs("DeptID").value)
    BignDate.value = IIf(IsNull(rs("BignDate").value), Date, rs("BignDate").value)
    LastUpdateDate.value = IIf(IsNull(rs("LastUpdateDate").value), Date, rs("LastUpdateDate").value)
    lbl(23).Caption = IIf(IsNull(rs("CurSalary").value), 0, rs("CurSalary").value)
    lbl(13).Caption = IIf(IsNull(rs("CurSalary").value), 0, rs("CurSalary").value)
    TxtLastUpdateSalary.Text = IIf(IsNull(rs("LastUpdateSalary").value), 0, rs("LastUpdateSalary").value)
    TxtTypeValu.Text = IIf(IsNull(rs("TypeValu").value), 0, rs("TypeValu").value)
    TxtRemarkEmp.Text = IIf(IsNull(rs("RemarkEmp").value), "", rs("RemarkEmp").value)
    TxtRemarkManger.Text = IIf(IsNull(rs("RemarkManger").value), "", rs("RemarkManger").value)
    TxtRemarkHR.Text = IIf(IsNull(rs("RemarkHR").value), "", rs("RemarkHR").value)
    TxtRemarkAccount.Text = IIf(IsNull(rs("RemarkAccount").value), "", rs("RemarkAccount").value)
    Me.txtremark.Text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    
    '''''''''''''''''''''''''''''''''''''''
    CBTybe.ListIndex = IIf(IsNull(rs("AddTybe").value), 0, rs("AddTybe").value)
    DataCombo5.ListIndex = IIf(IsNull(rs("ADDtype_Contract").value), 0, rs("ADDtype_Contract").value)
    ContractUPdata.ListIndex = IIf(IsNull(rs("ContractUPdata").value), 0, rs("ContractUPdata").value)
    JobUPdata.BoundText = IIf(IsNull(rs("JobUPdata").value), "", rs("JobUPdata").value)
    Me.TxtRemarkUPdata.Text = IIf(IsNull(rs("TxtRemarkUPdata").value), "", rs("TxtRemarkUPdata").value)
     '''''''''''''''''''''''''''''''''''''''''''''''
    If rs("ChekAll").value = True Then
     ChekAll.value = vbChecked
   Else
   ChekAll.value = vbUnchecked
   End If
   DcbType.ListIndex = IIf(IsNull(rs("TypeIncrease").value), -1, rs("TypeIncrease").value)
     If rs("Approved").value = True Then
     Approved.value = vbChecked
   Else
   Approved.value = vbUnchecked
   End If
   
 '      If IsNull(rs("posted").value) Then
 '                                                  If SystemOptions.UserInterface = ArabicInterface Then
 '                                                   Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
 '                                                 Else
 '                                                   Accredit.Caption = " send to Approval   "
 '                                              End If
 '                                              Accredit.Enabled = True
 ' Else
 '                                                  If SystemOptions.UserInterface = ArabicInterface Then
 '                                                   Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
 '                                                 Else
 '                                                   Accredit.Caption = " sent to Approval   "
 '                                              End If
 '                                              Accredit.Enabled = False
 '  End If
 ''''''''''''''''''''
 Dim mofrdcode As Integer
 VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
              VSFlexGrid1.Rows = 1
    StrSQL = "SELECT     dbo.TblEmpIncreaseMofrd.mofrad_type, dbo.TblEmpIncreaseMofrd.IncreaseID, dbo.TblEmpIncreaseMofrd.Valuee, dbo.TblEmpIncreaseMofrd.RecoedDate, "
    StrSQL = StrSQL & "                  dbo.MOFRAD.name , dbo.MOFRAD.NameE, dbo.MOFRAD.id"
    StrSQL = StrSQL & " FROM         dbo.mofrad RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmpIncreaseMofrd ON dbo.mofrad.id = dbo.TblEmpIncreaseMofrd.mofrad_type"
    StrSQL = StrSQL & " Where (dbo.TblEmpIncreaseMofrd.IncreaseID = " & val(XPTxtID.Text) & ")"
  
Set RsDev = New ADODB.Recordset
       RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsDev.RecordCount > 0) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
               
                 .TextMatrix(i, .ColIndex("LineNo")) = i
                 .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Valuee").value), "", RsDev("Valuee").value)
                 .TextMatrix(i, .ColIndex("RecoedDate")) = IIf(IsNull(RsDev("RecoedDate").value), "", RsDev("RecoedDate").value)
                 .TextMatrix(i, .ColIndex("mofrad_type")) = IIf(IsNull(RsDev("id").value), "", RsDev("id").value)
                  RetrivetMofrdCode val(.TextMatrix(i, .ColIndex("mofrad_type"))), mofrdcode
                .TextMatrix(i, .ColIndex("AccountCode")) = mofrdcode
            
                
            
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                Else
                
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("nameE").value), "", RsDev("nameE").value)
               End If
                          
                RsDev.MoveNext
            Next i
 
        End With

    End If
    
 ''''''''''''
   StrSQL = " SELECT     dbo.TblEmpIncreaseSalaryDetalis.ID, dbo.TblEmpIncreaseSalaryDetalis.IDIncrease, dbo.TblEmpIncreaseSalaryDetalis.TypeID, "
   StrSQL = StrSQL & "                   dbo.TblEmpIncreaseSalaryDetalis.CurrValue, dbo.TblEmpIncreaseSalaryDetalis.IncreaseValue, dbo.TblEmpIncreaseSalaryDetalis.Chek,"
   StrSQL = StrSQL & "                      dbo.TblEmpIncreaseSalaryDetalis.MofradID , dbo.mofrad.name, dbo.mofrad.NameE"
   StrSQL = StrSQL & "   FROM         dbo.TblEmpIncreaseSalaryDetalis LEFT OUTER JOIN"
   StrSQL = StrSQL & "                      dbo.mofrad ON dbo.TblEmpIncreaseSalaryDetalis.MofradID = dbo.mofrad.id"
   StrSQL = StrSQL & "   Where (dbo.TblEmpIncreaseSalaryDetalis.typeid = 0) And (dbo.TblEmpIncreaseSalaryDetalis.IDIncrease = " & XPTxtID.Text & ")"
Set RsDev = New ADODB.Recordset
       RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsDev.RecordCount > 0) Then
        RsDev.MoveFirst
    
        With Me.FG
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                 .TextMatrix(i, .ColIndex("Count")) = i
            
                .TextMatrix(i, .ColIndex("CurrValue")) = IIf(IsNull(RsDev("CurrValue").value), "", RsDev("CurrValue").value)
            
                .TextMatrix(i, .ColIndex("MofradID")) = IIf(IsNull(RsDev("MofradID").value), "", RsDev("MofradID").value)
            
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                Else
                
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("nameE").value), "", RsDev("nameE").value)
               End If
                          
                RsDev.MoveNext
            Next i
 
        End With

    End If
    ''///////////////////////////
    
   StrSQL = " SELECT     dbo.TblEmpIncreaseSalaryDetalis.ID, dbo.TblEmpIncreaseSalaryDetalis.IDIncrease, dbo.TblEmpIncreaseSalaryDetalis.TypeID, "
   StrSQL = StrSQL & "                   dbo.TblEmpIncreaseSalaryDetalis.CurrValue, dbo.TblEmpIncreaseSalaryDetalis.IncreaseValue, dbo.TblEmpIncreaseSalaryDetalis.Chek,"
   StrSQL = StrSQL & "                      dbo.TblEmpIncreaseSalaryDetalis.MofradID , dbo.mofrad.name, dbo.mofrad.NameE ,dbo.TblEmpIncreaseSalaryDetalis.Typeincrease, dbo.TblEmpIncreaseSalaryDetalis.TypeValue"
   StrSQL = StrSQL & "   FROM         dbo.TblEmpIncreaseSalaryDetalis LEFT OUTER JOIN"
   StrSQL = StrSQL & "                      dbo.mofrad ON dbo.TblEmpIncreaseSalaryDetalis.MofradID = dbo.mofrad.id"
   StrSQL = StrSQL & "   Where (dbo.TblEmpIncreaseSalaryDetalis.typeid = 1) And (dbo.TblEmpIncreaseSalaryDetalis.IDIncrease = " & XPTxtID.Text & ")"
Set RsDev = New ADODB.Recordset
       RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsDev.RecordCount > 0) Then
        RsDev.MoveFirst
    
        With Me.FG1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                 .TextMatrix(i, .ColIndex("Count")) = i
            If (RsDev("Chek").value = True) Then
            .TextMatrix(i, .ColIndex("Chek")) = -1
            Else
            .TextMatrix(i, .ColIndex("Chek")) = ""
            End If
                .TextMatrix(i, .ColIndex("CurrValue")) = IIf(IsNull(RsDev("CurrValue").value), "", RsDev("CurrValue").value)
                .TextMatrix(i, .ColIndex("IncreaseValue")) = IIf(IsNull(RsDev("IncreaseValue").value), "", RsDev("IncreaseValue").value)
            
                .TextMatrix(i, .ColIndex("MofradID")) = IIf(IsNull(RsDev("MofradID").value), "", RsDev("MofradID").value)
                .TextMatrix(i, .ColIndex("Typeincrease")) = IIf(IsNull(RsDev("Typeincrease").value), "", RsDev("Typeincrease").value)
                .TextMatrix(i, .ColIndex("TypeValue")) = IIf(IsNull(RsDev("TypeValue").value), "", RsDev("TypeValue").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                Else
                
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("nameE").value), "", RsDev("nameE").value)
               End If
                          
                RsDev.MoveNext
            Next i
 
        End With

    End If
    ''/////////////
   StrSQL = " SELECT     dbo.TblEmpIncreaseSalaryDetalis.ID, dbo.TblEmpIncreaseSalaryDetalis.IDIncrease, dbo.TblEmpIncreaseSalaryDetalis.TypeID, "
   StrSQL = StrSQL & "                   dbo.TblEmpIncreaseSalaryDetalis.CurrValue, dbo.TblEmpIncreaseSalaryDetalis.IncreaseValue, dbo.TblEmpIncreaseSalaryDetalis.Chek,"
   StrSQL = StrSQL & "                      dbo.TblEmpIncreaseSalaryDetalis.MofradID , dbo.mofrad.name, dbo.mofrad.NameE"
   StrSQL = StrSQL & "   FROM         dbo.TblEmpIncreaseSalaryDetalis LEFT OUTER JOIN"
   StrSQL = StrSQL & "                      dbo.mofrad ON dbo.TblEmpIncreaseSalaryDetalis.MofradID = dbo.mofrad.id"
   StrSQL = StrSQL & "   Where (dbo.TblEmpIncreaseSalaryDetalis.typeid = 2) And (dbo.TblEmpIncreaseSalaryDetalis.IDIncrease = " & XPTxtID.Text & ")"
Set RsDev = New ADODB.Recordset
       RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsDev.RecordCount > 0) Then
        RsDev.MoveFirst
    
        With Me.FG2
    
            .Rows = .FixedRows + RsDev.RecordCount

             For i = .FixedRows To .Rows - 1
 
                 .TextMatrix(i, .ColIndex("Count")) = i
            
                .TextMatrix(i, .ColIndex("CurrValue")) = IIf(IsNull(RsDev("CurrValue").value), "", RsDev("CurrValue").value)
                .TextMatrix(i, .ColIndex("IncreaseValue")) = IIf(IsNull(RsDev("IncreaseValue").value), "", RsDev("IncreaseValue").value)
                 .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("IncreaseValue"))) + val(.TextMatrix(i, .ColIndex("CurrValue")))
                .TextMatrix(i, .ColIndex("MofradID")) = IIf(IsNull(RsDev("MofradID").value), "", RsDev("MofradID").value)
            
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                Else
                
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("nameE").value), "", RsDev("nameE").value)
               End If
                          
                RsDev.MoveNext
            Next i
 
        End With

    End If
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
Dim mofrdcode As Integer
    'On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
            Else
            Msg = "Please Select Employee"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboEmpName.SetFocus
           SendKeys "{F4}"
            Exit Sub
        End If
        Dim RsTest As New ADODB.Recordset
   
 
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then

            XPTxtID.Text = CStr(new_id("TblEmpIncreaseSalary", "ID", "", True))
      
        
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
           StrSQL = "Delete From TblEmpIncreaseSalaryDetalis Where IDIncrease=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL
                StrSQL = "delete    EmpSalaryComponent where IncreaseID =" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL
                 StrSQL = "delete    TblEmpIncreaseMofrd where IncreaseID =" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL

        End If
         rs("ID").value = val(XPTxtID.Text)
         rs("RecordDate").value = XPDtbTrans.value
         rs("DateIncrease").value = DateIncrease.value
         rs("EmpID").value = IIf(DcboEmpName.BoundText = "", Null, DcboEmpName.BoundText)
         rs("DeptID").value = IIf(Me.DcbDepartment.BoundText = "", Null, DcbDepartment.BoundText)
         rs("JobID").value = IIf(Me.DcboJobsType.BoundText = "", Null, DcboJobsType.BoundText)
         rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
         rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)
         rs("ProjectID").value = IIf(DCproject.BoundText = "", Null, DCproject.BoundText)
         rs("BignDate").value = BignDate.value
         rs("LastUpdateDate").value = LastUpdateDate.value
         rs("CurSalary").value = IIf(lbl(23).Caption = "", Null, val(lbl(23).Caption))
         rs("LastUpdateSalary").value = IIf(TxtLastUpdateSalary.Text = "", Null, val(TxtLastUpdateSalary.Text))
         rs("TypeValu").value = IIf(TxtTypeValu.Text = "", Null, val(TxtTypeValu.Text))
         rs("RemarkEmp").value = IIf(Me.TxtRemarkEmp.Text = "", Null, TxtRemarkEmp.Text)
         rs("RemarkManger").value = IIf(Me.TxtRemarkManger.Text = "", Null, TxtRemarkManger.Text)
         rs("RemarkHR").value = IIf(Me.TxtRemarkHR.Text = "", Null, TxtRemarkHR.Text)
         rs("RemarkAccount").value = IIf(Me.TxtRemarkAccount.Text = "", Null, TxtRemarkAccount.Text)
         rs("Remark").value = IIf(Me.txtremark.Text = "", Null, txtremark.Text)
         ' aladein ADD
        ''''''''''''''''''''''''''''''''''
        rs("AddTybe").value = val(CBTybe.ListIndex)
        rs("ADDtype_Contract").value = val(DataCombo5.ListIndex)
        rs("ContractUPdata").value = val(ContractUPdata.ListIndex)
        rs("JobUPdata").value = IIf(Me.JobUPdata.BoundText = "", Null, JobUPdata.BoundText)
        rs("TxtRemarkUPdata").value = IIf(Me.TxtRemarkUPdata.Text = "", Null, TxtRemarkUPdata.Text)
        ''''''''''''''''''''''''''''''''''''''
        If ChekAll.value = vbChecked Then
         rs("ChekAll").value = 1
       Else
         rs("ChekAll").value = 0
      End If
      rs("TypeIncrease").value = IIf(val(DcbType.ListIndex) = -1, Null, val(DcbType.ListIndex))
       If Approved.value = vbChecked Then
         rs("Approved").value = 1
       Else
         rs("Approved").value = 0
      End If

        rs.update
        ''''////////
       Dim rscomponent As ADODB.Recordset
       Set rscomponent = New ADODB.Recordset
       Dim sql As String
                    Set RsDetails = New ADODB.Recordset
          StrSQL = "SELECT     *  from dbo.TblEmpIncreaseMofrd Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     sql = "SELECT     * from dbo.EmpSalaryComponent Where (1 = -1)"
   rscomponent.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid1
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("mofrad_type"))) <> 0 Then
           RsDetails.AddNew
           RsDetails("IncreaseID").value = val(XPTxtID.Text)
           RsDetails("mofrad_type").value = val(.TextMatrix(i, .ColIndex("mofrad_type")))
           RsDetails("RecoedDate").value = IIf(.TextMatrix(i, .ColIndex("RecoedDate")) = "", Null, .TextMatrix(i, .ColIndex("RecoedDate")))
           RsDetails("Valuee").value = val(.TextMatrix(i, .ColIndex("value")))
                RsDetails.update
             rscomponent.AddNew
            rscomponent("emp_ID").value = val(DcboEmpName.BoundText)
            rscomponent("IncreaseID").value = val(XPTxtID.Text)
            rscomponent("AccountCode").value = .TextMatrix(i, .ColIndex("AccountCode"))
            rscomponent("AccountName").value = .TextMatrix(i, .ColIndex("AccountName"))
            rscomponent("EntIncresDataM").value = IIf(.TextMatrix(i, .ColIndex("RecoedDate")) = "", Null, .TextMatrix(i, .ColIndex("RecoedDate")))
            rscomponent("value").value = val(.TextMatrix(i, .ColIndex("value")))
            rscomponent("EntIncresDataH").value = IIf(.TextMatrix(i, .ColIndex("RecoedDate")) = "", Null, ToHijriDate(.TextMatrix(i, .ColIndex("RecoedDate"))))
            rscomponent("Flagx").value = 1
            rscomponent.update
           End If
        Next i
        End With
       ''/////1111111111111
              Set RsDetails = New ADODB.Recordset
          StrSQL = "SELECT     *  from dbo.TblEmpIncreaseSalaryDetalis Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With FG

   
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("MofradID"))) <> 0 Then
           RsDetails.AddNew
           RsDetails("IDIncrease").value = val(XPTxtID.Text)
           RsDetails("MofradID").value = val(.TextMatrix(i, .ColIndex("MofradID")))
           RsDetails("TypeID").value = 0
           RsDetails("CurrValue").value = val(.TextMatrix(i, .ColIndex("CurrValue")))
                RsDetails.update
           End If
        Next i
        End With
'''///222
              Set RsDetails = New ADODB.Recordset
          StrSQL = "SELECT     *  from dbo.TblEmpIncreaseSalaryDetalis Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With FG1

   
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("MofradID"))) <> 0 Then
           RsDetails.AddNew
           RsDetails("IDIncrease").value = val(XPTxtID.Text)
           RsDetails("MofradID").value = val(.TextMatrix(i, .ColIndex("MofradID")))
           RsDetails("TypeID").value = 1
           RsDetails("CurrValue").value = val(.TextMatrix(i, .ColIndex("CurrValue")))
           RsDetails("IncreaseValue").value = val(.TextMatrix(i, .ColIndex("IncreaseValue")))
           RsDetails("Typeincrease").value = val(.TextMatrix(i, .ColIndex("Typeincrease")))
           RsDetails("TypeValue").value = val(.TextMatrix(i, .ColIndex("TypeValue")))
          If .Cell(flexcpChecked, i, .ColIndex("Chek")) = flexChecked Then
          RsDetails("Chek").value = -1
         Else
         RsDetails("Chek").value = 0
          End If
         RsDetails.update
         End If
        Next i
        End With
              Set RsDetails = New ADODB.Recordset
          StrSQL = "SELECT     *  from dbo.TblEmpIncreaseSalaryDetalis Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With FG2

   
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("MofradID"))) <> 0 Then
           RsDetails.AddNew
           RsDetails("IDIncrease").value = val(XPTxtID.Text)
           RsDetails("MofradID").value = val(.TextMatrix(i, .ColIndex("MofradID")))
           
           
           RsDetails("TypeID").value = 2
           RsDetails("CurrValue").value = val(.TextMatrix(i, .ColIndex("CurrValue")))
           RsDetails("IncreaseValue").value = val(.TextMatrix(i, .ColIndex("IncreaseValue")))
         '  RsDetails("TotaIncre").value = val(.TextMatrix(i, .ColIndex("total")))
         RsDetails.update

         End If
        Next i
        End With
        Set RsDetails = New ADODB.Recordset
        StrSQL = "SELECT     * from dbo.EmpSalaryComponent Where (1 = -1)"
     RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     With FG2
     For i = .FixedRows To .Rows - 1
        If val(.TextMatrix(i, .ColIndex("MofradID"))) <> 0 Then
           RsDetails.AddNew
           
           RsDetails("Flagx").value = 1
           RsDetails("IncreaseID").value = val(XPTxtID.Text)
           RsDetails("Emp_id").value = val(Me.DcboEmpName.BoundText)
           RsDetails("mofrad_type").value = val(.TextMatrix(i, .ColIndex("MofradID")))
           RsDetails("EntIncresDataM").value = DateIncrease.value
           RsDetails("value").value = val(.TextMatrix(i, .ColIndex("IncreaseValue")))
           RetrivetMofrdCode val(.TextMatrix(i, .ColIndex("MofradID"))), mofrdcode
           RsDetails("AccountCode").value = mofrdcode
           RsDetails("AccountName").value = (.TextMatrix(i, .ColIndex("name")))
         RsDetails.update

         End If
        Next i
        End With
        Cn.CommitTrans
        BeginTrans = False
        RsDetails.Close
        Set RsDetails = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                Msg = "This is Record Already Saved" & CHR(13)
                Msg = Msg & "You Need To Enter another Record "
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                MsgBox "Saved SuccessFully"
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
 Sub RetrivetMofrdCode(Optional mofrdtype As Integer, Optional ByRef mofrad_code As Integer)
     Dim sql As String
    Dim rs As New ADODB.Recordset
 
  sql = "SELECT     mofrad_code from  mofrdat WHERE  mofrad_type=" & mofrdtype & "   "
     rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

 mofrad_code = IIf(IsNull(rs("mofrad_code").value), 0, rs("mofrad_code").value)
End If
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
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        MsgBox "ConFirm Delete"
        End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From TblEmpIncreaseSalary Where ID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblEmpIncreaseSalaryDetalis Where IDIncrease=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 StrSQL = "delete    TblEmpIncreaseMofrd where IncreaseID =" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL
                  StrSQL = "delete    EmpSalaryComponent where IncreaseID =" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL
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



'Function FillApprovedTable()
' Dim RSApproval  As New ADODB.Recordset
'   Set RSApproval = New ADODB.Recordset
'   Dim currentdate As Date
'   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'
' Dim sql As String
'  Dim Rs1 As New ADODB.Recordset
' Dim i As Integer
'    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
'  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
'  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
'  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
'  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
'sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.name & "')"
'sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "
'
'    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Rs1.RecordCount > 0 Then
'            currentdate = Now
'            For i = 1 To Rs1.RecordCount
'              RSApproval.AddNew
'                RSApproval("ScreenName").value = Me.name
'                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
'               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
'                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
'                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
'                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
'                   RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
'                RSApproval("Transaction_Date").value = Date
'
'                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.name), currentdate)
'               RSApproval("SendTime").value = currentdate
'
'                 If i = 1 Then
'                        RSApproval("Currcursor").value = 1
'                         RSApproval("FromUser").value = user_name
'                End If
'
'                RSApproval.update
'                Rs1.MoveNext
'            Next i
'
'    End If
'
'
'
'End Function



'Function fillapprovData()
'Dim Num As Integer
' Dim RsDetails As New ADODB.Recordset
' Dim StrSQL As String
'
'
' StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
'StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
'StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
'StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
'StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
'StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
'StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.name & "')"
'StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"
'
'    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
' If Not (RsDetails.EOF Or RsDetails.BOF) Then
'        GRID2.Rows = RsDetails.RecordCount + 1
'
'
'        For Num = 1 To RsDetails.RecordCount
'
'       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
'    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
'   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
'   Else
'    GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
'    End If
'
'        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
'           If SystemOptions.UserInterface = ArabicInterface Then
'            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
'          Else
'             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
'          End If
'            If SystemOptions.UserInterface = ArabicInterface Then
'            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
'            Else
'            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
'            End If
'            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
'          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
'
'
'RsDetails.MoveNext
'If Num = RsDetails.RecordCount Then
'
'        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
'                                If SystemOptions.UserInterface = ArabicInterface Then
'                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
'                                 Else
'                                       Label11.Caption = "Approved"
'                                 End If
'                            Label11.backcolor = &H80FF80
'        Else
'                             If SystemOptions.UserInterface = ArabicInterface Then
'                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
'                            Else
'                                     Label11.Caption = "Currently required Approve"
'                            End If
'                 Label11.backcolor = &HFFFFC0
'        End If
'
'End If
'
'        Next Num
'Else
' GRID2.Rows = 1
'    End If
'RsDetails.Close
'
'End Function


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
        .Create Me.hwnd, " “Ì«œ… —« » ·„ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "   “Ì«œ… —« » ·„ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  “Ì«œ… —« » ·„ÊŸ› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  “Ì«œ… —« » ·„ÊŸ› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  “Ì«œ… —« » ·„ÊŸ› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  “Ì«œ… —« » ·„ÊŸ› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hwnd, "  “Ì«œ… —« » ·„ÊŸ› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  “Ì«œ… —« » ·„ÊŸ› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  “Ì«œ… —« » ·„ÊŸ› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "   “Ì«œ… —« » ·„ÊŸ›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "   “Ì«œ… —« » ·„ÊŸ›  ", 1, 15204351, -2147483630
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
   Private Sub FullComb()
    If SystemOptions.UserInterface = EnglishInterface Then
    With Me.DataCombo5
        .Clear
        .AddItem "Single"
        .AddItem "Family"
        End With
    Else
    With Me.DataCombo5
        .Clear
        .AddItem "›—œÌ"
        .AddItem "⁄«∆·Ì"
        End With
    End If
    
       With CBTybe
       .Clear
      If SystemOptions.UserInterface = ArabicInterface Then
         .AddItem " «· —ﬁÌ… ›ﬁÿ "
         .AddItem "«· —ﬁÌ… „⁄  ⁄œÌ· «·—« » "
        .AddItem " ⁄œÌ· —« » ›ﬁÿ "
         .AddItem "„‰Õ „“«Ì« «Œ—Ï "
         Else
         .AddItem "Promotion Only "
         .AddItem "Promotion With Salary Adjustment"
         .AddItem "Salary Adjustment Only"
         .AddItem "Granting Other Benefits"
     End If
     End With
       If SystemOptions.UserInterface = EnglishInterface Then
    With Me.ContractUPdata
        .Clear
        .AddItem "Single"
        .AddItem "Family"
        End With
    Else
    With Me.ContractUPdata
        .Clear
        .AddItem "›—œÌ"
        .AddItem "⁄«∆·Ì"
        End With
    End If
 End Sub

 


