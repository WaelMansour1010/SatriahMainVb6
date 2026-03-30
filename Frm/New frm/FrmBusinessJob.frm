VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBusinessJob 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  ﬂ·Ì› »„Â„… ⁄„·  "
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12720
   Icon            =   "FrmBusinessJob.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   12720
   Begin VB.TextBox TxtSearchCode 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   13920
      TabIndex        =   70
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   13440
      TabIndex        =   69
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
      TabIndex        =   67
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
      TabIndex        =   63
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   735
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   13140
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   0
      Width           =   12765
      _cx             =   22516
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
      Caption         =   "  ﬂ·Ì› »„Â„… ⁄„·  "
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
         TabIndex        =   46
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
         ButtonImage     =   "FrmBusinessJob.frx":6852
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
         TabIndex        =   47
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
         ButtonImage     =   "FrmBusinessJob.frx":6BEC
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
         TabIndex        =   48
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
         ButtonImage     =   "FrmBusinessJob.frx":6F86
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
         TabIndex        =   49
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
         ButtonImage     =   "FrmBusinessJob.frx":7320
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2280
         TabIndex        =   68
         Top             =   0
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   5640
      TabIndex        =   50
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   245497857
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DcboEmpName 
      Height          =   315
      Left            =   5640
      TabIndex        =   2
      Top             =   1185
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   120
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   7920
      Width           =   12585
      _cx             =   22199
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
         Left            =   11310
         TabIndex        =   35
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
         Left            =   9975
         TabIndex        =   36
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
         Left            =   8655
         TabIndex        =   37
         Top             =   75
         Width           =   1245
         _ExtentX        =   2196
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
         Left            =   7320
         TabIndex        =   38
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
         Left            =   5985
         TabIndex        =   39
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
         Left            =   0
         TabIndex        =   43
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
         Left            =   1575
         TabIndex        =   42
         Top             =   60
         Width           =   1275
         _ExtentX        =   2249
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
         Left            =   4560
         TabIndex        =   40
         Top             =   60
         Width           =   1245
         _ExtentX        =   2196
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
         Left            =   3000
         TabIndex        =   41
         Top             =   60
         Width           =   1365
         _ExtentX        =   2408
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
      Height          =   312
      Left            =   8520
      TabIndex        =   52
      Top             =   7320
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   13200
      TabIndex        =   53
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
      Left            =   14520
      TabIndex        =   64
      Top             =   2400
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
   Begin MSDataListLib.DataCombo Dcbranch 
      Bindings        =   "FrmBusinessJob.frx":76BA
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   720
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   5532
      Left            =   120
      TabIndex        =   72
      Top             =   1680
      Width           =   12480
      _cx             =   22013
      _cy             =   9758
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
      Caption         =   "ÿ·» «· ﬂ·Ì›|Õ«·… «·«⁄ „«œ|≈‰Â«¡ «·„Â„…|⁄—÷ «·„Â«„"
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
      Picture(0)      =   "FrmBusinessJob.frx":76CF
      Picture(1)      =   "FrmBusinessJob.frx":DF31
      Picture(2)      =   "FrmBusinessJob.frx":14793
      Picture(3)      =   "FrmBusinessJob.frx":1AFF5
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   4830
         Left            =   13725
         TabIndex        =   167
         TabStop         =   0   'False
         Top             =   45
         Width           =   12390
         _cx             =   21855
         _cy             =   8520
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
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Height          =   5070
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   168
            Top             =   0
            Width           =   12390
            Begin VB.TextBox TXTEndTask 
               Alignment       =   1  'Right Justify
               Height          =   1920
               Left            =   720
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   172
               Top             =   1920
               Width           =   10305
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00E2E9E9&
               Height          =   1215
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   169
               Top             =   240
               Width           =   10335
               Begin VB.CheckBox ChkStates 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " „ ≈‰Â«¡ ÿ·»  ﬂ·Ì› „Â„… «·⁄„·"
                  Height          =   195
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   170
                  Top             =   720
                  Value           =   1  'Checked
                  Width           =   2655
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·ÿ·»"
                  Height          =   255
                  Index           =   57
                  Left            =   9000
                  TabIndex        =   171
                  Top             =   240
                  Width           =   1200
               End
            End
            Begin ImpulseButton.ISButton Printfinshed 
               Height          =   495
               Left            =   720
               TabIndex        =   173
               Top             =   3960
               Width           =   3285
               _ExtentX        =   5794
               _ExtentY        =   873
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄… ≈‰Â«¡ „Â„… «·⁄„·"
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
               ButtonImage     =   "FrmBusinessJob.frx":21857
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
               Caption         =   "«·„Â«„ «·„‰Ã“…"
               Height          =   255
               Index           =   58
               Left            =   7680
               TabIndex        =   174
               Top             =   1560
               Width           =   3240
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   4830
         Left            =   13425
         RightToLeft     =   -1  'True
         TabIndex        =   148
         Top             =   45
         Width           =   12390
         Begin VB.Frame Frame7 
            BackColor       =   &H00E2E9E9&
            Height          =   615
            Left            =   120
            TabIndex        =   164
            Top             =   4440
            Width           =   12135
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·≈Ã„«·Ì"
               Height          =   285
               Index           =   59
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblL 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00000080&
               Height          =   315
               Index           =   10
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   165
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            Height          =   735
            Left            =   240
            TabIndex        =   156
            Top             =   120
            Width           =   12135
            Begin VB.OptionButton Option3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ ÂÌ…"
               Height          =   315
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton Option2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”«—Ì…"
               Height          =   315
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ·"
               Height          =   315
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   157
               Top             =   240
               Width           =   1095
            End
            Begin VB.Image Imge 
               Height          =   255
               Left            =   2520
               Picture         =   "FrmBusinessJob.frx":280B9
               Stretch         =   -1  'True
               Top             =   240
               Width           =   255
            End
            Begin VB.Image Imgw 
               Height          =   255
               Left            =   5280
               Picture         =   "FrmBusinessJob.frx":28996
               Stretch         =   -1  'True
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·„Â„… "
               Height          =   195
               Index           =   64
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   240
               Width           =   1305
            End
         End
         Begin VB.Frame lbreg 
            BackColor       =   &H00E2E9E9&
            Height          =   735
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   840
            Width           =   6735
            Begin MSComCtl2.DTPicker DtpDateFrom 
               Height          =   330
               Left            =   2760
               TabIndex        =   151
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   245432323
               CurrentDate     =   38887
            End
            Begin MSComCtl2.DTPicker DtpDateTo 
               Height          =   330
               Left            =   120
               TabIndex        =   152
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   245432323
               CurrentDate     =   38887
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   315
               Index           =   63
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   240
               Width           =   840
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   315
               Index           =   61
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   154
               Top             =   240
               Width           =   660
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ  «·„Â„…"
               Height          =   195
               Index           =   60
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   240
               Width           =   1425
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
            Height          =   2745
            Left            =   120
            TabIndex        =   149
            Top             =   1680
            Width           =   12195
            _cx             =   21511
            _cy             =   4842
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   14871017
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
            BackColorAlternate=   16777088
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483633
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
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmBusinessJob.frx":2AA31
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
         Begin ImpulseButton.ISButton CmdSerch 
            Height          =   615
            Left            =   240
            TabIndex        =   163
            Top             =   960
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   1085
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
            BackStyle       =   0
            ButtonImage     =   "FrmBusinessJob.frx":2AB5C
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   4830
         Left            =   13125
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   45
         Width           =   12390
         _cx             =   21855
         _cy             =   8520
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
            TabIndex        =   74
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
            FormatString    =   $"FrmBusinessJob.frx":313BE
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
            Left            =   8880
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   3960
            Width           =   3375
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   4440
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4830
         Index           =   15
         Left            =   45
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   45
         Width           =   12390
         _cx             =   21855
         _cy             =   8520
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
         _GridInfo       =   $"FrmBusinessJob.frx":31501
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4800
            Index           =   16
            Left            =   15
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   15
            Width           =   12360
            _cx             =   21802
            _cy             =   8467
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
            Begin VB.TextBox txtTichetBack 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   6360
               MaxLength       =   10
               TabIndex        =   20
               Top             =   3090
               Width           =   1695
            End
            Begin VB.TextBox txtTichetGo 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   9570
               MaxLength       =   10
               TabIndex        =   19
               Top             =   3090
               Width           =   1275
            End
            Begin VB.TextBox txtvisa2 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   8175
               MaxLength       =   10
               TabIndex        =   17
               Top             =   2745
               Width           =   1335
            End
            Begin VB.TextBox txtVisa3 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   6360
               MaxLength       =   10
               TabIndex        =   18
               Top             =   2745
               Width           =   1695
            End
            Begin VB.TextBox txtVisa1 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   9570
               MaxLength       =   10
               TabIndex        =   16
               Top             =   2745
               Width           =   1275
            End
            Begin VB.TextBox txtTask 
               Alignment       =   1  'Right Justify
               Height          =   1140
               Left            =   6360
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   21
               Top             =   3540
               Width           =   4545
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Height          =   4830
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   -120
               Width           =   6060
               Begin VB.TextBox txtReason 
                  Alignment       =   1  'Right Justify
                  Height          =   852
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   34
                  Top             =   3720
                  Width           =   4572
               End
               Begin VB.CheckBox btnNotOk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "·« √Ê«›ﬁ "
                  Height          =   252
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   3240
                  Width           =   1572
               End
               Begin VB.CheckBox chkOk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ê«›ﬁ"
                  Height          =   252
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   3240
                  Width           =   1572
               End
               Begin VB.TextBox TxttotalExpenses 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   120
                  MaxLength       =   10
                  TabIndex        =   29
                  Top             =   1320
                  Width           =   1635
               End
               Begin VB.TextBox TxtTicketExpenses 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   120
                  MaxLength       =   10
                  TabIndex        =   27
                  Top             =   960
                  Width           =   1635
               End
               Begin VB.TextBox txtcarExpenses2 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   120
                  MaxLength       =   10
                  TabIndex        =   25
                  Top             =   600
                  Width           =   1635
               End
               Begin VB.TextBox TxtfoodExpenses 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   120
                  MaxLength       =   10
                  TabIndex        =   23
                  Top             =   240
                  Width           =   1635
               End
               Begin VB.TextBox txtoldExpenses 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   3000
                  Locked          =   -1  'True
                  MaxLength       =   10
                  TabIndex        =   28
                  Top             =   1320
                  Width           =   1635
               End
               Begin VB.TextBox TxtJobExpenses 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   3000
                  MaxLength       =   10
                  TabIndex        =   26
                  Top             =   960
                  Width           =   1635
               End
               Begin VB.TextBox TxtHousingExpenses 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   3000
                  MaxLength       =   10
                  TabIndex        =   24
                  Top             =   600
                  Width           =   1635
               End
               Begin VB.TextBox txtcarExpenses 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   3000
                  MaxLength       =   10
                  TabIndex        =   22
                  Top             =   240
                  Width           =   1635
               End
               Begin ImpulseButton.ISButton Accredit 
                  Height          =   240
                  Left            =   0
                  TabIndex        =   135
                  Top             =   4680
                  Visible         =   0   'False
                  Width           =   1848
                  _ExtentX        =   3254
                  _ExtentY        =   423
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
               Begin MSDataListLib.DataCombo dcoManger 
                  Height          =   288
                  Left            =   120
                  TabIndex        =   30
                  Top             =   2160
                  Width           =   4512
                  _ExtentX        =   7964
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcoCeo 
                  Height          =   288
                  Left            =   120
                  TabIndex        =   31
                  Top             =   2640
                  Width           =   4512
                  _ExtentX        =   7964
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”»» «·—›÷"
                  Height          =   750
                  Index           =   56
                  Left            =   4560
                  TabIndex        =   147
                  Top             =   3720
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„œÌ— «·⁄«„ /«· ‰›Ì–Ï"
                  Height          =   525
                  Index           =   55
                  Left            =   4800
                  TabIndex        =   146
                  Top             =   2640
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„œÌ— «·„»«‘—"
                  Height          =   285
                  Index           =   53
                  Left            =   4800
                  TabIndex        =   143
                  Top             =   2160
                  Width           =   1125
               End
               Begin VB.Line Line3 
                  X1              =   0
                  X2              =   6120
                  Y1              =   1800
                  Y2              =   1800
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’«›Ì «·„” Õﬁ ’—›Â"
                  Height          =   405
                  Index           =   46
                  Left            =   1800
                  TabIndex        =   134
                  Top             =   1320
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " –ﬂ—… ”›—"
                  Height          =   285
                  Index           =   45
                  Left            =   1800
                  TabIndex        =   133
                  Top             =   960
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ· «” Œœ«„ ”Ì«—…"
                  Height          =   405
                  Index           =   44
                  Left            =   1800
                  TabIndex        =   132
                  Top             =   480
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ⁄«„"
                  Height          =   285
                  Index           =   43
                  Left            =   1800
                  TabIndex        =   131
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ”„ ·„« ”»ﬁ ’—›Â"
                  Height          =   285
                  Index           =   42
                  Left            =   4680
                  TabIndex        =   130
                  Top             =   1320
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ· „Â„… ⁄„·"
                  Height          =   285
                  Index           =   41
                  Left            =   4680
                  TabIndex        =   129
                  Top             =   960
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ· ”ﬂ‰"
                  Height          =   285
                  Index           =   40
                  Left            =   4680
                  TabIndex        =   128
                  Top             =   600
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Õ—Êﬁ«  / „ ”Ì«—…"
                  Height          =   285
                  Index           =   39
                  Left            =   4680
                  TabIndex        =   127
                  Top             =   240
                  Width           =   1365
               End
            End
            Begin VB.TextBox TxtDiscount 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   600
               MaxLength       =   10
               TabIndex        =   113
               Top             =   1560
               Width           =   375
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
               Height          =   2415
               Index           =   0
               Left            =   3150
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   1560
               Width           =   1515
               Begin VB.TextBox TxtPaymentCounts 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   4110
                  MaxLength       =   2
                  TabIndex        =   106
                  Top             =   240
                  Width           =   825
               End
               Begin VB.ComboBox CmbMonth 
                  Height          =   315
                  Left            =   4110
                  Style           =   2  'Dropdown List
                  TabIndex        =   105
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.CheckBox ChkSaleryDis 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Œ’„ „‰ «·„— »  ·ﬁ«∆Ì«"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3960
                  TabIndex        =   104
                  Top             =   2160
                  Value           =   1  'Checked
                  Visible         =   0   'False
                  Width           =   1965
               End
               Begin VB.ComboBox CboYear 
                  Height          =   315
                  Left            =   4110
                  Style           =   2  'Dropdown List
                  TabIndex        =   103
                  Top             =   1320
                  Width           =   1095
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   435
                  Index           =   8
                  Left            =   4080
                  TabIndex        =   102
                  Top             =   1680
                  Width           =   1965
                  _ExtentX        =   3466
                  _ExtentY        =   767
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈Õ”»  Ê«—ÌŒ «·”œ«œ"
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
                  ButtonImage     =   "FrmBusinessJob.frx":31535
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   2325
                  Left            =   90
                  TabIndex        =   107
                  Top             =   210
                  Width           =   3855
                  _cx             =   6800
                  _cy             =   4101
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
                  Rows            =   50
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBusinessJob.frx":318CF
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·œ›⁄« "
                  Height          =   285
                  Index           =   9
                  Left            =   4830
                  TabIndex        =   112
                  Top             =   300
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «Ê· œ›⁄…"
                  Height          =   285
                  Index           =   10
                  Left            =   4380
                  TabIndex        =   111
                  Top             =   690
                  Width           =   1665
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
                  TabIndex        =   110
                  Top             =   2280
                  Visible         =   0   'False
                  Width           =   2595
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   315
                  Index           =   11
                  Left            =   5250
                  TabIndex        =   109
                  Top             =   990
                  Width           =   405
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”‰…"
                  Height          =   315
                  Index           =   12
                  Left            =   5250
                  TabIndex        =   108
                  Top             =   1320
                  Width           =   405
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  „«·Ì…"
               Height          =   1275
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   -1140
               Width           =   1515
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·„” Õﬁ«  ··„ÊŸ›"
                  Height          =   285
                  Index           =   17
                  Left            =   3960
                  TabIndex        =   100
                  Top             =   720
                  Width           =   1965
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·⁄ﬁœ «·„ »ﬁÌ…"
                  Height          =   285
                  Index           =   18
                  Left            =   1560
                  TabIndex        =   99
                  Top             =   720
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”·› ·„  ”œœ"
                  Height          =   285
                  Index           =   19
                  Left            =   1800
                  TabIndex        =   98
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   285
                  Index           =   16
                  Left            =   -240
                  TabIndex        =   97
                  Top             =   720
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   20
                  Left            =   960
                  TabIndex        =   96
                  Top             =   720
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   21
                  Left            =   960
                  TabIndex        =   95
                  Top             =   360
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   22
                  Left            =   3240
                  TabIndex        =   94
                  Top             =   720
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„— »…"
                  Height          =   285
                  Index           =   14
                  Left            =   4800
                  TabIndex        =   93
                  Top             =   360
                  Width           =   1125
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·„ÊŸ›"
               Height          =   3180
               Left            =   6180
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   -555
               Width           =   6240
               Begin VB.ComboBox CBOTransportTypeID 
                  Height          =   288
                  ItemData        =   "FrmBusinessJob.frx":3195A
                  Left            =   120
                  List            =   "FrmBusinessJob.frx":3195C
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   1440
                  Width           =   1632
               End
               Begin VB.TextBox TxtPaymentVchrNo 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   2400
                  MaxLength       =   10
                  TabIndex        =   14
                  Top             =   2880
                  Width           =   1155
               End
               Begin VB.TextBox txtPaymentVchrValue 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   120
                  MaxLength       =   10
                  TabIndex        =   15
                  Top             =   2880
                  Width           =   1395
               End
               Begin VB.TextBox TxtjobLocation 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   120
                  MaxLength       =   10
                  TabIndex        =   9
                  Top             =   1800
                  Width           =   1632
               End
               Begin VB.TextBox TxtInterval 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   3240
                  MaxLength       =   10
                  TabIndex        =   8
                  Top             =   1800
                  Width           =   1752
               End
               Begin MSDataListLib.DataCombo DcboEmpDepartments 
                  Height          =   288
                  Left            =   3240
                  TabIndex        =   5
                  Top             =   1080
                  Width           =   1752
                  _ExtentX        =   3096
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DBIssueDate 
                  Height          =   315
                  Left            =   7560
                  TabIndex        =   86
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   246808577
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcboJobsType 
                  Height          =   288
                  Left            =   120
                  TabIndex        =   4
                  Top             =   720
                  Width           =   1632
                  _ExtentX        =   2884
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker startDate 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   10
                  Top             =   2160
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   246808577
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker startTime 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   12
                  Top             =   2520
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  CustomFormat    =   "'Time: 'hh:mm tt"
                  Format          =   246808579
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin MSComCtl2.DTPicker EndDate 
                  Height          =   312
                  Left            =   120
                  TabIndex        =   11
                  Top             =   2160
                  Width           =   1692
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   246808577
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker EndTime 
                  Height          =   312
                  Left            =   120
                  TabIndex        =   13
                  Top             =   2520
                  Width           =   1692
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _Version        =   393216
                  CustomFormat    =   "'Time: 'hh:mm tt"
                  Format          =   246808579
                  UpDown          =   -1  'True
                  CurrentDate     =   40909
               End
               Begin ImpulseButton.ISButton btnQuery 
                  Height          =   330
                  Left            =   2040
                  TabIndex        =   137
                  TabStop         =   0   'False
                  ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
                  Top             =   2880
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   ""
                  BackColor       =   14737632
                  FontSize        =   9.75
                  FontName        =   "Arial"
                  FontBold        =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmBusinessJob.frx":3195E
                  ColorButton     =   14737632
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcboSpecifications 
                  Height          =   288
                  Left            =   3240
                  TabIndex        =   6
                  Top             =   1440
                  Width           =   1752
                  _ExtentX        =   3096
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   285
                  Index           =   48
                  Left            =   3120
                  TabIndex        =   138
                  Top             =   720
                  Width           =   1845
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»ﬁÌ„…"
                  Height          =   285
                  Index           =   47
                  Left            =   1560
                  TabIndex        =   136
                  Top             =   3000
                  Width           =   405
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœÌ „ﬁœ„ »‰«¡ ⁄·Ï ”‰œ ’—› —ﬁ„"
                  Height          =   405
                  Index           =   38
                  Left            =   3600
                  TabIndex        =   125
                  Top             =   2880
                  Width           =   2445
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”«⁄Â «·⁄Êœ…"
                  Height          =   285
                  Index           =   37
                  Left            =   1920
                  TabIndex        =   124
                  Top             =   2520
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”«⁄Â «·”›—"
                  Height          =   285
                  Index           =   36
                  Left            =   5040
                  TabIndex        =   123
                  Top             =   2520
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄Êœ… „‰ «·”›—"
                  Height          =   285
                  Index           =   35
                  Left            =   1920
                  TabIndex        =   122
                  Top             =   2160
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «·”›—"
                  Height          =   285
                  Index           =   34
                  Left            =   5040
                  TabIndex        =   121
                  Top             =   2160
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„ﬂ«‰ «·„Â„…"
                  Height          =   285
                  Index           =   33
                  Left            =   1920
                  TabIndex        =   120
                  Top             =   1800
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ê”Ì·Â «·‰ﬁ·"
                  Height          =   285
                  Index           =   32
                  Left            =   1800
                  TabIndex        =   119
                  Top             =   1440
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «Ì«„ «·„Â„…"
                  Height          =   288
                  Index           =   2
                  Left            =   5040
                  TabIndex        =   118
                  Top             =   1812
                  Width           =   1008
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›∆…"
                  Height          =   288
                  Index           =   31
                  Left            =   5400
                  TabIndex        =   117
                  Top             =   1440
                  Width           =   648
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ã‰”Ì…"
                  Height          =   285
                  Index           =   29
                  Left            =   5400
                  TabIndex        =   116
                  Top             =   720
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—« » «·«”«”Ì"
                  Height          =   285
                  Index           =   5
                  Left            =   1800
                  TabIndex        =   91
                  Top             =   1080
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
                  Height          =   285
                  Index           =   13
                  Left            =   6840
                  TabIndex        =   90
                  Top             =   360
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Êﬁ⁄ «·⁄„·"
                  Height          =   285
                  Index           =   15
                  Left            =   5280
                  TabIndex        =   89
                  Top             =   1080
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  ForeColor       =   &H00FF0000&
                  Height          =   288
                  Index           =   23
                  Left            =   120
                  TabIndex        =   88
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   1608
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊŸÌ›…"
                  Height          =   285
                  Index           =   24
                  Left            =   1920
                  TabIndex        =   87
                  Top             =   720
                  Width           =   1125
               End
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄Êœ…"
               Height          =   285
               Index           =   52
               Left            =   8295
               TabIndex        =   142
               Top             =   3090
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " –ﬂ—… ”›— –Â«» "
               Height          =   285
               Index           =   51
               Left            =   10905
               TabIndex        =   141
               Top             =   3090
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· √‘Ì—« "
               Height          =   255
               Index           =   50
               Left            =   11265
               TabIndex        =   140
               Top             =   2745
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Â„… ⁄„· "
               Height          =   360
               Index           =   28
               Left            =   11145
               TabIndex        =   139
               Top             =   3540
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÊÌŒ’„ „‰ «·”·› „»·€« Êﬁœ—…"
               Height          =   600
               Index           =   26
               Left            =   915
               TabIndex        =   114
               Top             =   1560
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2760
               Index           =   62
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   1335
               Width           =   120
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4800
            Index           =   9
            Left            =   15
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   15
            Width           =   12360
            _cx             =   21802
            _cy             =   8467
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
               Height          =   3615
               Left            =   795
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   1035
               Width           =   180
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   2520
               Left            =   1035
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   1275
               Width           =   240
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2520
               Index           =   67
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   1275
               Width           =   180
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   2415
               Index           =   68
               Left            =   975
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   1635
               Width           =   60
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
               Height          =   2910
               Index           =   69
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   1275
               Width           =   75
            End
         End
         Begin VB.Line Line2 
            X1              =   15
            X2              =   12390
            Y1              =   15
            Y2              =   4830
         End
      End
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   285
      Left            =   240
      TabIndex        =   144
      Top             =   0
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label LBLSTATES 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   1560
      TabIndex        =   158
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·„Â„…"
      Height          =   285
      Index           =   65
      Left            =   3120
      TabIndex        =   159
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   288
      Index           =   54
      Left            =   4800
      TabIndex        =   145
      Top             =   0
      Width           =   1008
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
      TabIndex        =   71
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
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Index           =   49
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   780
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   12810
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   285
      Index           =   4
      Left            =   11550
      TabIndex        =   62
      Top             =   750
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   285
      Index           =   3
      Left            =   11550
      TabIndex        =   61
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   7440
      TabIndex        =   60
      Top             =   735
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   276
      Index           =   8
      Left            =   11400
      TabIndex        =   59
      Top             =   7320
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   312
      Index           =   7
      Left            =   2880
      TabIndex        =   58
      Top             =   7440
      Width           =   1068
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   312
      Index           =   6
      Left            =   960
      TabIndex        =   57
      Top             =   7440
      Width           =   972
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   552
      Left            =   336
      TabIndex        =   56
      Top             =   7380
      Width           =   612
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   552
      Left            =   1980
      TabIndex        =   55
      Top             =   7380
      Width           =   732
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   13350
      TabIndex        =   54
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmBusinessJob"
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
    Retrive (val(Me.XPTxtID.text))
End Sub

Private Sub btnNotOk_Click()

If btnNotOk.value = 1 Then
chkOk.value = 0
txtReason.Visible = True
lbl(56).Visible = True
End If

End Sub
Private Sub btnQuery_Click()
            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 5
            FrmNotesSearch.m_SearchType2 = 1
            FrmNotesSearch.person = DcboEmpName.text
            FrmNotesSearch.show vbModal
End Sub
Private Sub chkOk_Click()

If chkOk.value = 1 Then
btnNotOk.value = 0
txtReason.Visible = False
lbl(56).Visible = False
End If

End Sub

 Private Sub ChkStates_Click()
 If ChkStates.value = False Then
  TXTEndTask.Enabled = False
  Printfinshed.Enabled = False
  Else
    TXTEndTask.Enabled = True
  Printfinshed.Enabled = True
 End If
 End Sub
Private Sub Cmd_Click(index As Integer)
    ' On Error GoTo ErrTrap
    Select Case index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            lbl(20).Caption = "0"
            lbl(21).Caption = "0"
            lbl(22).Caption = "0"
            lbl(23).Caption = "0"
            CBOTransportTypeID.ListIndex = 0
            
              Grid2.Clear flexClearScrollable, flexClearEverything
    Grid2.rows = 1
            Me.DCboUserName.BoundText = user_id
            TxtPaymentCounts.text = 1
dcBranch.BoundText = Current_branch
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

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
    
            Dim Msg As String

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
     'Alaa  General_Search.send_form = "BJ"
     'Alaa      Load General_Search
     'Alaa      General_Search.show
        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 8
            CalCulateParts
            
            
                 Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text)
        
        
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


   MySQL = MySQL & "  SELECT dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, TblEmployee_1.Emp_Code, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Namee,"
     MySQL = MySQL & "               dbo.TblUsers.UserName, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpDepartments.DepartmentName,"
      MySQL = MySQL & "              dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeparmentID, dbo.TblEmpJobsTypes.JobTypeID, dbo.TblEmpJobOrder.Branch_NO,"
      MySQL = MySQL & "              dbo.TblEmpGrades.name, dbo.TblEmpGrades.namee, dbo.TblEmpJobOrder.AdvanceID, dbo.TblEmpJobOrder.interval, dbo.TblEmpJobOrder.AdvanceDate,"
       MySQL = MySQL & "             dbo.TblEmpJobOrder.DeparmentID AS Expr1, dbo.TblEmpJobOrder.gradeID, dbo.TblEmpJobOrder.JobTypeID AS Expr2, dbo.TblEmpJobOrder.basicSalary,"
     MySQL = MySQL & "               dbo.TblEmpJobOrder.nationalId, dbo.TblEmpJobOrder.TransportTypeID, dbo.TblEmpJobOrder.JobLocation, dbo.TblEmpJobOrder.startDate, dbo.TblEmpJobOrder.startTime,"
      MySQL = MySQL & "              dbo.TblEmpJobOrder.EndDate, dbo.TblEmpJobOrder.EndTime, dbo.TblEmpJobOrder.PaymentVchrNo, dbo.TblEmpJobOrder.PaymentVchrValue,"
    MySQL = MySQL & "                dbo.TblEmpJobOrder.carExpenses, dbo.TblEmpJobOrder.HousingExpenses, dbo.TblEmpJobOrder.JobExpenses, dbo.TblEmpJobOrder.oldExpenses,"
        MySQL = MySQL & "            dbo.TblEmpJobOrder.foodExpenses, dbo.TblEmpJobOrder.carExpenses2, dbo.TblEmpJobOrder.TicketExpenses, dbo.TblEmpJobOrder.totalExpenses,"
       MySQL = MySQL & "             dbo.TblEmpJobOrder.Nationality, dbo.TblEmpJobOrder.Visa1, dbo.TblEmpJobOrder.ticketGo, dbo.TblEmpJobOrder.ticketBack, dbo.TblEmpJobOrder.Task,"
  MySQL = MySQL & "                  dbo.TblEmpJobOrder.Reason, dbo.TblEmpJobOrder.Visa2, dbo.TblEmpJobOrder.Visa3, dbo.TblEmpJobOrder.ok, dbo.TblEmpJobOrder.notok, dbo.TblEmpJobOrder.Manager,"
  MySQL = MySQL & "                  TblEmployee_2.Emp_Name AS Manager_name, dbo.TblEmployee.Emp_Name AS Ceo_Name"
  MySQL = MySQL & "  FROM     dbo.TblEmployee RIGHT OUTER JOIN"
  MySQL = MySQL & "              dbo.TblEmpJobOrder LEFT OUTER JOIN"
   MySQL = MySQL & "                 dbo.TblEmpGrades ON dbo.TblEmpJobOrder.gradeID = dbo.TblEmpGrades.gradeid ON dbo.TblEmployee.Emp_ID = dbo.TblEmpJobOrder.Ceo LEFT OUTER JOIN"
    MySQL = MySQL & "                dbo.TblEmployee AS TblEmployee_2 ON dbo.TblEmpJobOrder.Manager = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
   MySQL = MySQL & "                 dbo.TblEmpJobsTypes ON dbo.TblEmpJobOrder.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
  MySQL = MySQL & "                  dbo.TblEmpDepartments ON dbo.TblEmpJobOrder.DeparmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblUsers ON dbo.TblEmpJobOrder.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
  MySQL = MySQL & "                   dbo.TblEmployee AS TblEmployee_1 ON dbo.TblEmpJobOrder.Emp_id = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.TblEmpJobOrder.Branch_NO = dbo.TblBranchesData.branch_id"




   MySQL = MySQL & "  Where (dbo.TblEmpJobOrder.AdvanceID = " & val(XPTxtID.text) & ")"

 
      
        If SystemOptions.UserInterface = ArabicInterface Then
          ' StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\BusinessJob.rpt"
              StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "BusinessJob.rpt"
        Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\BusinessJob.rpt"
              'StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BusinessJob.rpt"
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
                
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtInterval.text), "0.00"), 0, True, ".")
        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
         xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(FG.TextMatrix(Me.FG.FixedRows, FG.ColIndex("PartValue"))), val(FG.TextMatrix(Me.FG.FixedRows, FG.ColIndex("PartValue"))), 0)
 xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
  'xReport.ParameterFields(10).AddCurrentValue CStr(FormatDateTime(startTime.value, vbLongTime))
   xReport.ParameterFields(11).AddCurrentValue CStr(FormatDateTime(EndTime.value, vbLongTime))
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Function print_reportt(Optional NoteSerial As String)
        
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = MySQL & "  SELECT  dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, TblEmployee_1.Emp_Code, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Namee, dbo.TblUsers.UserName,"
    MySQL = MySQL & "   dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
    MySQL = MySQL & "   dbo.TblEmpDepartments.DeparmentID, dbo.TblEmpJobsTypes.JobTypeID, dbo.TblEmpJobOrder.Branch_NO, dbo.TblEmpGrades.name, dbo.TblEmpGrades.namee, dbo.TblEmpJobOrder.AdvanceID,"
    MySQL = MySQL & "  dbo.TblEmpJobOrder.[interval], dbo.TblEmpJobOrder.AdvanceDate, dbo.TblEmpJobOrder.DeparmentID AS DeparmentIDjob, dbo.TblEmpJobOrder.gradeID,"
    MySQL = MySQL & "  dbo.TblEmpJobOrder.JobTypeID AS JobTypeIDjob, dbo.TblEmpJobOrder.basicSalary, dbo.TblEmpJobOrder.nationalId, dbo.TblEmpJobOrder.TransportTypeID, dbo.TblEmpJobOrder.JobLocation,"
    MySQL = MySQL & "  dbo.TblEmpJobOrder.startDate, dbo.TblEmpJobOrder.startTime, dbo.TblEmpJobOrder.EndDate, dbo.TblEmpJobOrder.EndTime, dbo.TblEmpJobOrder.PaymentVchrNo,"
    MySQL = MySQL & "   dbo.TblEmpJobOrder.PaymentVchrValue, dbo.TblEmpJobOrder.carExpenses, dbo.TblEmpJobOrder.HousingExpenses, dbo.TblEmpJobOrder.JobExpenses, dbo.TblEmpJobOrder.oldExpenses,"
    MySQL = MySQL & "  dbo.TblEmpJobOrder.foodExpenses, dbo.TblEmpJobOrder.carExpenses2, dbo.TblEmpJobOrder.TicketExpenses, dbo.TblEmpJobOrder.totalExpenses, dbo.TblEmpJobOrder.Nationality,"
    MySQL = MySQL & "  dbo.TblEmpJobOrder.Visa1, dbo.TblEmpJobOrder.ticketGo, dbo.TblEmpJobOrder.ticketBack, dbo.TblEmpJobOrder.Task, dbo.TblEmpJobOrder.Reason, dbo.TblEmpJobOrder.Visa2,"
    MySQL = MySQL & "  dbo.TblEmpJobOrder.Visa3, dbo.TblEmpJobOrder.ok, dbo.TblEmpJobOrder.notok, dbo.TblEmpJobOrder.Manager, TblEmployee_2.Emp_Name AS Manager_name,"
    MySQL = MySQL & "  dbo.TblEmployee.Emp_Name AS Ceo_Name, dbo.TblEmpJobOrder.AssignmentTybe, dbo.TblEmpJobOrder.EndTask"
    MySQL = MySQL & "    FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmpJobOrder LEFT OUTER JOIN"
    MySQL = MySQL & "  dbo.TblEmpGrades ON dbo.TblEmpJobOrder.gradeID = dbo.TblEmpGrades.gradeid ON dbo.TblEmployee.Emp_ID = dbo.TblEmpJobOrder.Ceo LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmployee TblEmployee_2 ON dbo.TblEmpJobOrder.Manager = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmpJobsTypes ON dbo.TblEmpJobOrder.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmpDepartments ON dbo.TblEmpJobOrder.DeparmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
    MySQL = MySQL & "  dbo.TblUsers ON dbo.TblEmpJobOrder.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmployee TblEmployee_1 ON dbo.TblEmpJobOrder.Emp_id = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblBranchesData ON dbo.TblEmpJobOrder.Branch_NO = dbo.TblBranchesData.branch_id"

    MySQL = MySQL & "  Where (dbo.TblEmpJobOrder.AdvanceID = " & val(XPTxtID.text) & ")"

      If SystemOptions.UserInterface = ArabicInterface Then
          ' StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\BusinessJobFinshed.rpt"
              StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "BusinessJobFinshed.rpt"
        Else
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\BusinessJobFinshed.rpt"
              'StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BusinessJob.rpt"
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
                
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtInterval.text), "0.00"), 0, True, ".")
        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
         xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(FG.TextMatrix(Me.FG.FixedRows, FG.ColIndex("PartValue"))), val(FG.TextMatrix(Me.FG.FixedRows, FG.ColIndex("PartValue"))), 0)
 xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
  'xReport.ParameterFields(10).AddCurrentValue CStr(FormatDateTime(startTime.value, vbLongTime))
   xReport.ParameterFields(11).AddCurrentValue CStr(FormatDateTime(EndTime.value, vbLongTime))
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub
  Private Sub CmdSerch_Click()
  GetData
  End Sub
  Public Sub GetData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
   sql = "SELECT      dbo.TblEmpJobOrder.AdvanceID, dbo.TblEmpJobOrder.Branch_NO, dbo.TblEmpJobOrder.Emp_id, dbo.TblEmpJobOrder.AdvanceDate, dbo.TblEmpJobOrder.DeparmentID,"
   sql = sql + "      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmpDepartments.DepartmentName,"
   sql = sql + "      dbo.TblEmpDepartments.DepartmentNamee , dbo.TblEmpJobOrder.AssignmentTybe"
   sql = sql + "      FROM         dbo.TblEmpJobOrder LEFT OUTER JOIN"
   sql = sql + "      dbo.TblEmpDepartments ON dbo.TblEmpJobOrder.DeparmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
   sql = sql + "     dbo.TblEmployee ON dbo.TblEmpJobOrder.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
   sql = sql + "    dbo.TblBranchesData ON dbo.TblEmpJobOrder.Branch_NO = dbo.TblBranchesData.branch_id"
    
       BolBegine = False
       StrWhere = ""
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpJobOrder.AdvanceDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpJobOrder.AdvanceDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEmpJobOrder.AdvanceDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEmpJobOrder.AdvanceDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    
      
        If (Me.Option2.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEmpJobOrder.AssignmentTybe = 0 "
         Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEmpJobOrder.AssignmentTybe = 1 "
        End If
        End If
        
        If (Me.Option3.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblEmpJobOrder.AssignmentTybe = 1 "
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblEmpJobOrder.AssignmentTybe = 0 "
        End If
        End If
  '-----------------------------------
    sql = sql & StrWhere
    sql = sql & " Order By  dbo.TblEmpJobOrder.AdvanceID"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = " ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid1
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = rs.RecordCount
            End If
                rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("AdvanceID").value), "", rs("AdvanceID").value)
                 If Not (IsNull(rs("AdvanceDate").value)) Then
                .TextMatrix(i, .ColIndex("AdvanceDate")) = Format(rs("AdvanceDate").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
               Else
              .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
              .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
              .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
               End If
              .TextMatrix(i, .ColIndex("AssignmentTybe")) = IIf(IsNull(rs("AssignmentTybe").value), "", rs("AssignmentTybe").value)
              rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub Printfinshed_Click()
   If ChkStates.value = vbChecked Then
        If TXTEndTask.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ﬂ «»… «·„Â«„ «·„‰Ã“… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Exit Sub
            Else
            MsgBox "Write Finshed Tasks ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
             Exit Sub
            End If
       End If
  End If
  print_reportt
End Sub
Private Sub txtcarExpenses_Change()
calNet
End Sub

Private Sub txtcarExpenses_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyAscii_Num(KeyAscii, txtcarExpenses.text, 0)
End Sub

Private Sub txtcarExpenses2_Change()
calNet
End Sub

Private Sub txtcarExpenses2_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, txtcarExpenses2.text, 0)
End Sub

Private Sub TxtfoodExpenses_Change()
calNet
End Sub

Private Sub TxtfoodExpenses_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtfoodExpenses.text, 0)
End Sub

Private Sub TxtHousingExpenses_Change()
calNet
End Sub

Private Sub TxtHousingExpenses_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtHousingExpenses.text, 0)
End Sub

Private Sub TxtJobExpenses_Change()
calNet
End Sub

Private Sub TxtJobExpenses_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtJobExpenses.text, 0)
End Sub

Private Sub txtoldExpenses_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, txtoldExpenses.text, 0)
End Sub

Function calNet()
txtoldExpenses.text = txtPaymentVchrValue.text
txtTotalExpenses = (val(txtcarExpenses) + val(txtcarExpenses2) + val(TxtHousingExpenses) + val(TxtJobExpenses) + val(TxtfoodExpenses) + val(TxtTicketExpenses)) - val(txtoldExpenses)

End Function

Private Sub TxtPaymentVchrNo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
      Load FrmNotesSearch
            FrmNotesSearch.SearchType = 5
            FrmNotesSearch.m_SearchType2 = 1
            FrmNotesSearch.person = DcboEmpName.text
            FrmNotesSearch.show vbModal
            
End If
End Sub

Private Sub txtPaymentVchrValue_Change()
calNet
End Sub

Private Sub txtPaymentVchrValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, txtPaymentVchrValue.text, 0)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub

 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 5
        FrmEmployeeSearch.show
  
    End If

End Sub

Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
    
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
        Dim endContractPerMonth As Double
        Dim Nationality As String
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth, Nationality
        
          WriteCustomerBalPublic Account_code2, Balance
          
  lbl(22).Caption = val(Balance)

          WriteCustomerBalPublic Account_code, Balance
          
  lbl(21).Caption = val(Balance)
  lbl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
        DBIssueDate.value = IssueDate
        DcboEmpDepartments.BoundText = DepID
        DcboSpecifications.BoundText = gradeID
        DcboJobsType.BoundText = JobTypeID
    lbl(48).Caption = Nationality
        lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
        
    'End If

End Sub

Private Sub TxtTicketExpenses_Change()
calNet
End Sub

Private Sub TxtTicketExpenses_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtTicketExpenses.text, 0)
End Sub

Private Sub TxttotalExpenses_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, txtTotalExpenses.text, 0)
End Sub

Private Sub XPDtbTrans_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""

End Sub

Private Sub Dcbranch_Click(Area As Integer)
 
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

    With Me.FG
        .RowHeightMin = 300
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
    AdditemTocCmp
 

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
        Dcombos.GetEmployees Me.dcoCeo
        Dcombos.GetEmployees Me.dcoManger
        Dcombos.GetBranches Me.dcBranch
        Dcombos.GetEmpDepartments Me.DcboEmpDepartments
        Dcombos.GetEmpJobsTypes Me.DcboJobsType
        Dcombos.GetEmpGrades Me.DcboSpecifications
    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If
 
    
'///////////////////////////
If SystemOptions.UserInterface = EnglishInterface Then
SetInterface Me
ChangeLang
End If
    
    SetDtpickerDate Me.XPDtbTrans
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
    Option1.value = True
    YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEmpJobOrder     Order By AdvanceID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.text = "R"
    Retrive
    XPTab301.CurrTab = 0
  '  XPTab301.TabVisible(3) = False
    'If SystemOptions.UserInterface = EnglishInterface Then
    '    SetInterface Me
    '    ChangeLang
    'End If

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
    Exit Sub

ErrTrap:
End Sub

Private Sub ChangeLang()
XPTab301.CurrTab = 0
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Label1.Visible = False
    Accredit.Caption = "Send To Approval"

   
    '/////////////////////
    lbl(53).Caption = "Direct manager"
    lbl(55).Caption = "CEO"
    lbl(50).Caption = "Visa"
    lbl(51).Caption = "Single Ticket"
    lbl(28).Caption = "Mession"
    lbl(56).Caption = "Refuse Reason"
    chkOk.Caption = "OK"
    btnNotOk.Caption = "Not Ok"
    lbl(52).Caption = "Return ticket"
    lbl(65).Caption = "Tybe"


    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    Me.Caption = " Task Request"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(3).Caption = "Employee"
    lbl(5).Caption = "Salary"
    lbl(49).Caption = "Branch"
    lbl(29).Caption = "Nationality"
    lbl(31).Caption = "Category"
    lbl(24).Caption = "Position"
    lbl(15).Caption = "Location"
   lbl(32).Caption = "Trans. Method"
    lbl(2).Caption = "No Day"
    lbl(33).Caption = "Place Task"
   lbl(34).Caption = "Start Trav'"
   lbl(35).Caption = "Back Trav"
  lbl(36).Caption = "Trav Hour"
  lbl(37).Caption = "Back Hour"
  lbl(38).Caption = "Criticism is submitted pursuant to Exchange support"
 lbl(47).Caption = "Value"
 lbl(43).Caption = "Food"
 lbl(40).Caption = "Housing allowance"
 lbl(41).Caption = "task allowance"
 lbl(44).Caption = "Car Allowance"
 lbl(45).Caption = "Ticket"
 lbl(46).Caption = "Net Rece Cashed"
 lbl(42).Caption = "Foregoing Cashed"
 lbl(28).Caption = "Remarks"
  ''''''''''''''''''''''''''''''''''
   ' XPTab301.CurrTab = 0
  '  XPTab301.TabCaption(1) = "End Mission"
   ' XPTab301.TabVisible(3) = False
  '  XPTab301.TabCaption(0) = "Tasks Search"
  '  XPTab301.TabCaption(2) = "Start Mission"
  '  XPTab301.TabCaption(3) = "Approval"
    
   'XPTab301.TabCaption(0) = "Tasks Search|End Mission|Start Mission|Approval"
    XPTab301.Caption = "Start Mission|Approval|Tasks Search|End Mission"
                        'ÿ·» «· ﬂ·Ì›|Õ«·… «·«⁄ „«œ|≈‰Â«¡ «·„Â„…|⁄—÷ «·„Â«„
   ' XPTab301.Caption = "Start Mission|Approval|End Mission|Tasks Search"
 ''''''''''''''''''''''''''''''''
 lbl(39).Caption = "Car fuel expenses"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
   ''''''''''''''''''''''''''''''''''
   lbl(57).Caption = "Tybe"
   lbl(58).Caption = "Finshed Tasks Remarks"
   ChkStates.Caption = "Mission Complete"
    Printfinshed.Caption = "Mission Complete Print"
    '''''''''''''''''''''''''''''''
    Label11.Caption = "Need to Approval "
    Label1100.Caption = "Need to Approval "
    ''''''''''''''''''''''''''''''''''''
   CmdSerch.Caption = "Search"
   lbl(60).Caption = "Date"
   lbl(61).Caption = "From"
   lbl(63).Caption = "To"
   lbl(64).Caption = "Type"
   Option1.Caption = "ALL"
   Option2.Caption = "ok"
   Option3.Caption = "Finshed"
    With Me.FG
        .TextMatrix(0, .ColIndex("PartNO")) = "NO"
        .TextMatrix(0, .ColIndex("PartValue")) = "Value"
        .TextMatrix(0, .ColIndex("PartDate")) = "Date"
    End With
     With Me.Grid2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "level Name"
        .TextMatrix(0, .ColIndex("EmpName")) = "Emp Name"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approv Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
    End With
    lbl(59).Caption = "Total"
    
   With Me.VSFlexGrid1
   .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "NO ."
        .TextMatrix(0, .ColIndex("AdvanceDate")) = "Date"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Emp Name"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
        .TextMatrix(0, .ColIndex("DepartmentName")) = "Department Name"
        .TextMatrix(0, .ColIndex("AssignmentTybe")) = "Assignment Tybe"
    End With
End Sub

Private Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2010 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex
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

Private Sub TxtInterval_LostFocus()
    Dim StrSQL As String
    Dim Mytot As String
    Dim MySal As String
    Exit Sub
    Dim Myrs As New ADODB.Recordset
    'StrSQL =
    Myrs.Open "SELECT * From TblEmployee  where Emp_ID=" & val(DcboEmpName.BoundText), Cn, adOpenStatic, adLockReadOnly

    If Not Myrs.EOF And Not IsNull(Myrs!Emp_Salary) Then
        MySal = Myrs!Emp_Salary
        Mytot = val(MySal) * 5

        If val(TxtInterval.text) >= Mytot Then
            MsgBox "⁄›Ê« «·”·›…  ⁄œ  «·Õœ  «·„”„ÊÕ »Â ÊÂÊ 5 «÷⁄«› ﬁÌ„Â «·—« »  " & CHR(13) & "   —« » «·„ÊŸ›    " & MySal, vbOKOnly, App.Title
            Exit Sub
   
        End If
  
    End If
   
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰"
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
            TxtInterval.locked = True
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
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰( ÃœÌœ )"
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
            TxtInterval.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰(  ⁄œÌ· )"
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
            TxtInterval.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtPaymentCounts_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtPaymentCounts.text, 1)
End Sub

Private Sub TxtPaymentCounts_LostFocus()

    If val(TxtPaymentCounts.text) > 84 Then
    
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "«·œ›«⁄  «ﬂ»— „‰ «·Õœ ", vbOKOnly, App.Title
        Else
         MsgBox "Payments is more than limit ", vbOKOnly, App.Title
    End If
         
        Exit Sub
    End If
 
End Sub

Private Sub XPBtnMove_Click(index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case index

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
            rs.Find "AdvanceID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
    

 txtVisa1.text = IIf(IsNull(rs("Visa1").value), "", (rs("Visa1").value))
 txtvisa2.text = IIf(IsNull(rs("Visa2").value), "", (rs("Visa2").value))
 txtVisa3.text = IIf(IsNull(rs("Visa3").value), "", (rs("Visa3").value))
 txtTichetGo.text = IIf(IsNull(rs("ticketGo").value), "", (rs("ticketGo").value))
 txtTichetBack.text = IIf(IsNull(rs("ticketBack").value), "", (rs("ticketBack").value))
 txtTask.text = IIf(IsNull(rs("Task").value), "", (rs("Task").value))
 If rs("Ok").value = True Then
 chkOk.value = 1
 ElseIf rs("Ok").value = False Then
 chkOk.value = 0
 End If
 If rs("notok").value = True Then
 btnNotOk.value = 1
 ElseIf rs("notok").value = False Then
 btnNotOk.value = 0
 End If
 '''''''''''''''''''''''''''''''''
  If SystemOptions.UserInterface = ArabicInterface Then
  If rs("AssignmentTybe").value = True Then
  ChkStates.value = 1
  LBLSTATES.Caption = "„‰ ÂÌ…"
  LBLSTATES.ForeColor = &HFF&
  TXTEndTask.Enabled = True
  Printfinshed.Enabled = True
  ElseIf rs("AssignmentTybe").value = False Then
  ChkStates.value = 0
  LBLSTATES.Caption = "”«—Ì…"
  LBLSTATES.ForeColor = &H8000&
  TXTEndTask.Enabled = False
  Printfinshed.Enabled = False
  End If
  Else
  If rs("AssignmentTybe").value = True Then
  ChkStates.value = 1
  LBLSTATES.Caption = "Finshed"
  LBLSTATES.ForeColor = &HFF&
  TXTEndTask.Enabled = True
  Printfinshed.Enabled = True
  ElseIf rs("AssignmentTybe").value = False Then
  ChkStates.value = 0
  LBLSTATES.Caption = "Started"
  LBLSTATES.ForeColor = &H8000&
  TXTEndTask.Enabled = False
  Printfinshed.Enabled = False
  End If
  End If
  TXTEndTask.text = IIf(IsNull(rs("EndTask").value), "", (rs("EndTask").value))
  ''''''''''''''''''''''''''''''''''''''''''''''
  If Not IsNull(rs("Manager").value) Then Me.dcoManger.BoundText = rs("Manager").value
  If Not IsNull(rs("Ceo").value) Then Me.dcoCeo.BoundText = rs("Ceo").value
     
     
     

    XPTxtID.text = IIf(IsNull(rs("AdvanceID").value), "", val(rs("AdvanceID").value))
    XPDtbTrans.value = IIf(IsNull(rs("AdvanceDate").value), Date, rs("AdvanceDate").value)
    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    
        DcboEmpDepartments.BoundText = IIf(IsNull(rs("DeparmentID").value), "", rs("DeparmentID").value)

    DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)

    DcboJobsType.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)

   
   If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(48).Caption = IIf(IsNull(rs("nationality").value), "", rs("nationality").value)
   End If

  CBOTransportTypeID.ListIndex = IIf(IsNull(rs("TransportTypeID").value), 0, rs("TransportTypeID").value)



    TxtjobLocation.text = IIf(IsNull(rs("jobLocation").value), "", rs("jobLocation").value)

  StartDate.value = rs("startDate").value

  startTime.value = rs("startTime").value
  EndDate.value = rs("EndDate").value
  EndTime.value = rs("EndTime").value
  TxtPaymentVchrNo.text = IIf(IsNull(rs("PaymentVchrNo").value), "", rs("PaymentVchrNo").value)
  txtPaymentVchrValue.text = IIf(IsNull(rs("PaymentVchrValue").value), 0, rs("PaymentVchrValue").value)
    
        txtcarExpenses.text = IIf(IsNull(rs("carExpenses").value), 0, rs("carExpenses").value)
        TxtHousingExpenses.text = IIf(IsNull(rs("HousingExpenses").value), 0, rs("HousingExpenses").value)
        TxtJobExpenses.text = IIf(IsNull(rs("JobExpenses").value), 0, rs("JobExpenses").value)
        txtoldExpenses.text = IIf(IsNull(rs("oldExpenses").value), 0, rs("oldExpenses").value)
        TxtfoodExpenses.text = IIf(IsNull(rs("foodExpenses").value), 0, rs("foodExpenses").value)
        txtcarExpenses2.text = IIf(IsNull(rs("carExpenses2").value), 0, rs("carExpenses2").value)
        TxtTicketExpenses.text = IIf(IsNull(rs("TicketExpenses").value), 0, rs("TicketExpenses").value)
  '      TxttotalExpenses.text = IIf(IsNull(rs("totalExpenses").value), 0, rs("totalExpenses").value)
 
   lbl(23).Caption = IIf(IsNull(rs("basicSalary").value), "", rs("basicSalary").value)
  Me.DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    TxtInterval.text = IIf(IsNull(rs("Interval").value), "", rs("Interval").value)
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

    If Me.TxtModFlg.text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
        
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
            Else
              Msg = "Select Employee Name !! "
        End If
            
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboEmpName.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
        If CBOTransportTypeID.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ Ê”Ì·Â «·‰ﬁ·   ..!! "
            Else
             Msg = "Seelct Transportation Method"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CBOTransportTypeID.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
       If ChkStates.value = vbChecked Then
        If TXTEndTask.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ﬂ «»… «·„Â«„ «·„‰Ã“… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Exit Sub
            Else
            MsgBox "Write Finshed Tasks ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
             Exit Sub
            End If
     End If
  End If
'
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("TblEmpJobOrder", "AdvanceID", "", True))
 
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
 

        End If
        '''''''''''''''''''''''''''''''
         rs("AssignmentTybe") = ChkStates.value
         rs("EndTask").value = IIf(TXTEndTask.text = "", Null, TXTEndTask.text)
        
        
        ''''''''''''''''''''''''''''''''''''
        rs("branch_no").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
        rs("AdvanceID").value = val(XPTxtID.text)
        rs("AdvanceDate").value = XPDtbTrans.value
        rs("Emp_ID").value = Me.DcboEmpName.BoundText
        rs("DeparmentID").value = Me.DcboEmpDepartments.BoundText
        rs("gradeID").value = val(Me.DcboSpecifications.BoundText)
        rs("JobTypeID").value = val(Me.DcboJobsType.BoundText)
        rs("basicSalary").value = val(lbl(23).Caption)
            rs("interval").value = IIf(TxtInterval.text = "", Null, val(TxtInterval.text))
            rs("UserID").value = Me.DCboUserName.BoundText
             rs("nationality").value = IIf(Me.lbl(48).Caption = "", Null, Me.lbl(48).Caption)
           rs("TransportTypeID").value = val(CBOTransportTypeID.ListIndex)
                     rs("JobLocation").value = IIf(TxtjobLocation.text = "", Null, (TxtjobLocation.text))
           rs("startDate").value = StartDate.value
           rs("startTime").value = startTime.value
           rs("EndDate").value = EndDate.value
           rs("EndTime").value = EndTime.value
              rs("PaymentVchrNo").value = IIf(TxtPaymentVchrNo.text = "", Null, (TxtPaymentVchrNo.text))
              rs("PaymentVchrValue").value = IIf(txtPaymentVchrValue.text = "", Null, val(txtPaymentVchrValue.text))
            rs("carExpenses").value = IIf(txtcarExpenses.text = "", Null, val(txtcarExpenses.text))
              rs("HousingExpenses").value = IIf(TxtHousingExpenses = "", Null, val(TxtHousingExpenses.text))
              rs("JobExpenses").value = IIf(TxtJobExpenses = "", Null, val(TxtJobExpenses.text))
              rs("oldExpenses").value = IIf(txtoldExpenses.text = "", Null, val(txtoldExpenses.text))
          rs("foodExpenses").value = IIf(TxtfoodExpenses.text = "", Null, val(TxtfoodExpenses.text))
                      rs("carExpenses2").value = IIf(txtcarExpenses2.text = "", Null, val(txtcarExpenses2.text))
                          rs("TicketExpenses").value = IIf(TxtTicketExpenses.text = "", Null, val(TxtTicketExpenses.text))
                              rs("totalExpenses").value = IIf(txtTotalExpenses.text = "", Null, val(txtTotalExpenses.text))
                              
                              
    
rs("Visa1").value = IIf(txtVisa1.text = "", Null, txtVisa1.text)
rs("Visa2").value = IIf(txtvisa2.text = "", Null, txtvisa2.text)
rs("Visa3").value = IIf(txtVisa3.text = "", Null, txtVisa3.text)
rs("ticketGo").value = IIf(txtTichetGo.text = "", Null, txtTichetGo.text)
rs("Visa3").value = IIf(txtTichetBack.text = "", Null, txtTichetBack.text)
rs("Task").value = IIf(txtTask.text = "", Null, txtTask.text)
rs("Ok") = chkOk.value
rs("notok") = btnNotOk.value
rs("Reason").value = IIf(txtReason.text = "", Null, txtReason.text)
rs("Manager").value = IIf(Me.dcoManger.BoundText = "", Null, val(Me.dcoManger.BoundText))
rs("Ceo").value = IIf(Me.dcoCeo.BoundText = "", Null, val(Me.dcoCeo.BoundText))


   rs.update
        
 
 
    
        Cn.CommitTrans
        BeginTrans = False
 
        Set RsDetails = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = " Process Data Saved " & CHR(13)
                Msg = Msg + "DO You want to add another data "
            Else
              Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
             End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                 MsgBox "ëUpdates Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
       If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
         Msg = "Data Can't  Saved " & CHR(13)
        Msg = Msg + "there is Invalid Data " & CHR(13)
        Msg = Msg + "Please Trye Again"
     
        End If
        
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
     Msg = "Sorry'... there is error while Saving Data  " & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "AdvanceID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
  Msg = "Process Data No. Will Be Delete  " & CHR(13)
        Msg = Msg + " Are You Sure you want to delete this data "

End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.text)
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
        
       If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
      Else
      Msg = "This Process Not Avilable Where ther is no data"
      End If
        
        
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
   Else
 Msg = "Sorry an error occur when deleting data  " & CHR(13)
   End If
    
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
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
                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
                   RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
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
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
   Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
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
 Grid2.rows = 1
    End If
RsDetails.Close

End Function


Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
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

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
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

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

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

Private Sub TxtInterval_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtInterval.text, 0)
End Sub

Private Function CheckDate() As Boolean
    Dim StrTemp As String
    Dim Msg  As String

    If year(Date) > val(Me.CboYear.text) Then ' ⁄«„ „÷Ï
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ ÕÌÀ «‰Â ﬁ»·  «—ÌŒ «·ÌÊ„ ÊÂ–« ·« ÌÃÊ“...!!!"
        Else
          Msg = "Invalid Date where date is previous Today ....!!!"
        End If
        
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CheckDate = False
        Exit Function
    ElseIf year(Date) = val(Me.CboYear.text) Then '‰›” «·⁄«„

        If Month(Date) > val(Me.CmbMonth.ListIndex + 1) Then
            'Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ...!!!"
            'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            'CheckDate = False
            'Exit Function
        End If
    End If

    CheckDate = True
End Function

Private Function CheckPartCal() As Boolean
    Dim Msg As String

    CheckPartCal = False

    If val(TxtInterval.text) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·”·›…...!!! "
        Else
        Msg = "Enter Advance Value !!!"
    End If
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtInterval.SetFocus
        Exit Function
    End If

    If val(TxtPaymentCounts.text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ» «œŒ«· ⁄œœ „—«   ”œÌœ «·œ›⁄…...!!! "
        Else
        Msg = "Please Enter Count of paied advance "
        End If
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtInterval.SetFocus
        Exit Function
    End If

    If CmbMonth.ListIndex = -1 Then
         If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ»  ÕœÌœ «Ê· ‘Â— · ”œÌœ «·œ›⁄…..!! "
        Else
          Msg = "Please Select first month for paid advance "
        End If
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CmbMonth.SetFocus
        Sendkeys "{F4}"
        Exit Function
    End If

    If CboYear.ListIndex = -1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ»  ÕœÌœ «Ê· ”‰… · ”œÌœ «·œ›⁄… ..!! "
        Else
           Msg = "Select First Year  "
    End If
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboYear.SetFocus
        Sendkeys "{F4}"
        Exit Function
    End If

    CheckPartCal = True
End Function

Private Sub CalCulateParts()
    Dim i As Integer
    Dim IntPartCounts As Integer
    Dim SngPartValue As Single
    Dim m_FirstDate As Date

    If CheckPartCal = False Then
        Exit Sub
    End If

    If CheckDate = False Then
        Exit Sub
    End If

    SngPartValue = val(Me.TxtInterval.text) / val(Me.TxtPaymentCounts.text)
    IntPartCounts = val(Me.TxtPaymentCounts.text)
    m_FirstDate = CDate(val(Me.CboYear.text) & "-" & Me.CmbMonth.ListIndex + 1 & "-01")

    With Me.FG
        .Clear flexClearScrollable, flexClearEverything
        .rows = .FixedRows + IntPartCounts
        .RowHeightMin = 300
        For i = 1 To IntPartCounts
            .TextMatrix(i, .ColIndex("PartNO")) = i
            .TextMatrix(i, .ColIndex("PartValue")) = SngPartValue
            .TextMatrix(i, .ColIndex("PartDate")) = DisplayDate(DateAdd("m", i - 1, m_FirstDate))
        Next i
        End With
End Sub
Private Sub AdditemTocCmp()
 On Error GoTo ErrTrap
    With CBOTransportTypeID
       .Clear
      If SystemOptions.UserInterface = ArabicInterface Then
         .AddItem " ”Ì«—… "
         .AddItem "ÿÌ—«‰ "
         Else
         .AddItem "Car "
         .AddItem "Flieght "
     End If
     End With
ErrTrap:
End Sub
'''''''''''''''''''''''''''''''' end
