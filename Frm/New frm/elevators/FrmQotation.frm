VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmQotation 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄—Ê÷ «·«”⁄— ··„’«⁄œ"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   FillColor       =   &H00C0E0FF&
   Icon            =   "FrmQotation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   13320
   Begin VB.TextBox Text15 
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
      Height          =   315
      Left            =   13440
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   3600
      Width           =   825
   End
   Begin VB.TextBox TxtDayPrice 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   13680
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   21600
      TabIndex        =   34
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   21480
      TabIndex        =   33
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   21480
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   21480
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   21480
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   13365
      _cx             =   23574
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
      Caption         =   "⁄—Ê÷ «·«”⁄«— ··„’«⁄œ"
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
         ButtonImage     =   "FrmQotation.frx":038A
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
         ButtonImage     =   "FrmQotation.frx":0724
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
         ButtonImage     =   "FrmQotation.frx":0ABE
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
         TabIndex        =   5
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
         ButtonImage     =   "FrmQotation.frx":0E58
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
         Left            =   2160
         TabIndex        =   32
         Top             =   0
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   8940
      TabIndex        =   6
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   97845249
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   270
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7980
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
         Left            =   7200
         TabIndex        =   8
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
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
         TabIndex        =   9
         Top             =   60
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
         TabIndex        =   10
         Top             =   60
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
         TabIndex        =   11
         Top             =   60
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
         TabIndex        =   12
         Top             =   60
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   25
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
         TabIndex        =   35
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
      Left            =   8340
      TabIndex        =   15
      Top             =   7560
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
      Left            =   21360
      TabIndex        =   16
      Top             =   4080
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
      Left            =   21360
      TabIndex        =   27
      Top             =   3120
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
      Bindings        =   "FrmQotation.frx":11F2
      Height          =   315
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   5895
      _ExtentX        =   10398
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
      Height          =   6375
      Left            =   0
      TabIndex        =   36
      Top             =   1080
      Width           =   13320
      _cx             =   23495
      _cy             =   11245
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
      Caption         =   "»Ì«‰« |Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "FrmQotation.frx":1207
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5910
         Left            =   13965
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   45
         Width           =   13230
         _cx             =   23336
         _cy             =   10425
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
            TabIndex        =   38
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
            FormatString    =   $"FrmQotation.frx":15A1
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
            TabIndex        =   50
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
            TabIndex        =   39
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   5910
         Index           =   15
         Left            =   45
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   45
         Width           =   13230
         _cx             =   23336
         _cy             =   10425
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
         _GridInfo       =   $"FrmQotation.frx":16ED
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5880
            Index           =   16
            Left            =   15
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   15
            Width           =   13200
            _cx             =   23283
            _cy             =   10372
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
            Frame           =   0
            FrameStyle      =   3
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
               Caption         =   "»Ì«‰«  «·œ›⁄« "
               Height          =   3180
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   3075
               Width           =   13125
               Begin VB.TextBox TxtShow 
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
                  Height          =   315
                  Left            =   3600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   2400
                  Width           =   1065
               End
               Begin VB.ComboBox DcbShowDMY 
                  Height          =   315
                  ItemData        =   "FrmQotation.frx":1721
                  Left            =   960
                  List            =   "FrmQotation.frx":172E
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   2400
                  Width           =   2535
               End
               Begin VB.TextBox TxtGranty 
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
                  Height          =   315
                  Left            =   9240
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   2400
                  Width           =   1065
               End
               Begin VB.ComboBox DcbGrantyDMY 
                  Height          =   315
                  ItemData        =   "FrmQotation.frx":1741
                  Left            =   6600
                  List            =   "FrmQotation.frx":174E
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   2400
                  Width           =   2535
               End
               Begin VB.TextBox TxtInstal 
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
                  Height          =   315
                  Left            =   3600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   2040
                  Width           =   1065
               End
               Begin VB.ComboBox DcbInstalDMY 
                  Height          =   315
                  ItemData        =   "FrmQotation.frx":1761
                  Left            =   960
                  List            =   "FrmQotation.frx":176E
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   2040
                  Width           =   2535
               End
               Begin VB.TextBox TxtSupply 
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
                  Height          =   315
                  Left            =   9240
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   2040
                  Width           =   1065
               End
               Begin VB.ComboBox DcbSupplyDMY 
                  Height          =   315
                  ItemData        =   "FrmQotation.frx":1781
                  Left            =   6600
                  List            =   "FrmQotation.frx":178E
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   2040
                  Width           =   2535
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   1725
                  Left            =   120
                  TabIndex        =   69
                  Top             =   240
                  Width           =   12885
                  _cx             =   22728
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
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmQotation.frx":17A1
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   8
                  Left            =   120
                  TabIndex        =   70
                  Top             =   2160
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   476
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–›"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmQotation.frx":1824
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "’·«ÕÌ… «·⁄—÷"
                  Height          =   285
                  Index           =   2
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   2400
                  Width           =   1410
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„œ… «·÷„«‰"
                  Height          =   285
                  Index           =   1
                  Left            =   9960
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   2400
                  Width           =   1410
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· —ﬂÌ» Œ·«·"
                  Height          =   285
                  Index           =   0
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   2040
                  Width           =   1410
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„œ… «· Ê—Ìœ"
                  Height          =   285
                  Index           =   60
                  Left            =   9960
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   2040
                  Width           =   1410
               End
            End
            Begin VB.Frame lblDataCli 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·⁄—÷"
               Height          =   3255
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   0
               Width           =   13125
               Begin VB.ComboBox DcbWith 
                  Height          =   315
                  ItemData        =   "FrmQotation.frx":1DBE
                  Left            =   120
                  List            =   "FrmQotation.frx":1DC0
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   720
                  Width           =   1995
               End
               Begin VB.TextBox TxtZibCode 
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
                  Height          =   375
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   240
                  Width           =   1995
               End
               Begin VB.TextBox TxtXY 
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
                  Height          =   375
                  Left            =   4440
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   240
                  Width           =   2145
               End
               Begin VB.TextBox TxtProject 
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
                  Height          =   375
                  Left            =   5520
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   720
                  Width           =   2145
               End
               Begin VB.TextBox TxtElevatorCount 
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
                  Height          =   375
                  Left            =   9600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   720
                  Width           =   1425
               End
               Begin VB.TextBox TxtCustomer 
                  Alignment       =   1  'Right Justify
                  Height          =   435
                  Left            =   7200
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   240
                  Width           =   4935
               End
               Begin VSFlex8UCtl.VSFlexGrid UnitsGrid 
                  Height          =   1605
                  Left            =   120
                  TabIndex        =   59
                  Top             =   1200
                  Width           =   12885
                  _cx             =   22728
                  _cy             =   2831
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
                  Rows            =   1
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmQotation.frx":1DC2
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   13
                  Left            =   12360
                  TabIndex        =   60
                  Top             =   2760
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   476
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–›"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmQotation.frx":1F70
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcboGovernmentID 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   90
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„œÌ‰…"
                  Top             =   720
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbCitM 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   91
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„œÌ‰…"
                  Top             =   240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„⁄ /»œÊ‰ €—›… „ﬂÌ‰…"
                  Height          =   255
                  Index           =   15
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·—„“ «·»—ÌœÌ"
                  Height          =   255
                  Index           =   14
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "’.»"
                  Height          =   255
                  Index           =   12
                  Left            =   6360
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   240
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   255
                  Index           =   5
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   2880
                  Width           =   2295
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   255
                  Index           =   2
                  Left            =   3360
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   2880
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "·„œÌ‰…"
                  Height          =   255
                  Index           =   11
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   840
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "·„‘—Ê⁄"
                  Height          =   255
                  Index           =   10
                  Left            =   7320
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   840
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„’⁄œ ﬂÂ—»«∆Ì"
                  Height          =   255
                  Index           =   9
                  Left            =   8400
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   840
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„ÊŸÊ⁄  Ê—Ìœ Ê —ﬂÌ» ⁄œœ"
                  Height          =   255
                  Index           =   3
                  Left            =   11040
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   840
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„ «·⁄„Ì·"
                  Height          =   255
                  Index           =   13
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3390
               Index           =   62
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   1290
               Width           =   615
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5880
            Index           =   9
            Left            =   15
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   15
            Width           =   13200
            _cx             =   23283
            _cy             =   10372
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
               Height          =   4410
               Left            =   3450
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   1275
               Width           =   720
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   2985
               Left            =   4365
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   1620
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2985
               Index           =   67
               Left            =   2445
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1620
               Width           =   690
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   2940
               Index           =   68
               Left            =   4170
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   2010
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
               Height          =   3480
               Index           =   69
               Left            =   3135
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1620
               Width           =   315
            End
         End
      End
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
      Height          =   315
      Left            =   7320
      TabIndex        =   55
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
   End
   Begin MSComCtl2.DTPicker From 
      Height          =   315
      Left            =   13560
      TabIndex        =   57
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   97845249
      CurrentDate     =   38784
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄„«—Â"
      Height          =   255
      Left            =   3660
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   0
      Width           =   855
   End
   Begin VB.Image img 
      Height          =   855
      Left            =   22680
      Picture         =   "FrmQotation.frx":250A
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   720
   End
   Begin VB.Image imgnul 
      Height          =   1095
      Left            =   22680
      Top             =   4800
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   15240
      Picture         =   "FrmQotation.frx":352E
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   720
   End
   Begin VB.Label lblBr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   780
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   20760
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄—÷"
      Height          =   285
      Index           =   4
      Left            =   12240
      TabIndex        =   24
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   9960
      TabIndex        =   23
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„” Œœ„"
      Height          =   270
      Index           =   8
      Left            =   10845
      TabIndex        =   22
      Top             =   7635
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2430
      TabIndex        =   21
      Top             =   7590
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   690
      TabIndex        =   20
      Top             =   7590
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   19
      Top             =   7620
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   18
      Top             =   7620
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   21240
      TabIndex        =   17
      Top             =   2640
      Width           =   1005
   End
End
Attribute VB_Name = "FrmQotation"
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
Public LngRow As Double

Private Sub RemoveGridRow()

    With Me.UnitsGrid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub
Private Sub RemoveGridRowFG()

    With Me.Fg

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
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
    
dcBranch.BoundText = Current_branch
  Me.DCboUserName.BoundText = user_id
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
 UnitsGrid.Rows = UnitsGrid.Rows + 1
            UnitsGrid.Enabled = True
             Fg.Rows = Fg.Rows + 1
            Fg.Enabled = True
            '
            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
    
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

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
            Load FrmSearchShowTech
            FrmSearchShowTech.show vbModal

        Case 6
            Unload Me

        Case 7
   

        Case 8
  RemoveGridRowFG
            
            
                 Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                print_report val(Me.XPTxtID.Text)
        
        
            End If
            Case 13
            RemoveGridRow
        
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


   MySQL = "SELECT     dbo.TblQutations.ID, dbo.TblQutationsDetls1.QutationsID, dbo.TblQutationsDetls1.Matrials, dbo.TblQutationsDetls1.[Count], dbo.TblQutationsDetls1.loader, "
  MySQL = MySQL & "                    dbo.TblQutationsDetls1.CountPark, dbo.TblQutationsDetls1.Spead, dbo.TblQutationsDetls1.Price, dbo.TblQutationsDetls1.DataTech, dbo.TblQutations.BranchID,"
  MySQL = MySQL & "                    dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblQutations.UserID, dbo.TblUsers.UserName, dbo.TblQutations.CityID,"
  MySQL = MySQL & "                    TblCountriesGovernments_1.GovernmentName, dbo.TblQutations.ElevatorCount, dbo.TblQutations.Supply, dbo.TblQutations.Granty, dbo.TblQutations.Instal,"
  MySQL = MySQL & "                    dbo.TblQutations.Show, dbo.TblQutations.InstalDMY, dbo.TblQutations.SupplyDMY, dbo.TblQutations.GrantyDMY, dbo.TblQutations.ShowDMY,"
  MySQL = MySQL & "                    dbo.TblQutations.RecordDate, dbo.TblQutations.RecordDateH, dbo.TblQutations.CusName, dbo.TblQutations.ProjectName, dbo.TblQutationsDetls1.Type,"
  MySQL = MySQL & "                    dbo.TblQutationsDetls1.Des, dbo.TblQutationsDetls1.Rate, dbo.TblQutations.Xy, dbo.TblQutations.ZibCode, dbo.TblQutations.CityIDH,"
  MySQL = MySQL & "                    TblCountriesGovernments_1.GovernmentName AS GovernmentNameH, dbo.TblQutations.Without ,dbo.TblQutations.ElevatorType ,   dbo.TblQutationsDetls1.ID AS IDrat"
  MySQL = MySQL & " FROM         dbo.TblQutations LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblQutations.CityIDH = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblQutations.CityID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblUsers ON dbo.TblQutations.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.TblQutations.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblQutationsDetls1 ON dbo.TblQutations.ID = dbo.TblQutationsDetls1.QutationsID"

 MySQL = MySQL & " Where (dbo.TblQutations.id = " & val(XPTxtID.Text) & ")"




        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepQutation.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepQutation.rpt"
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
       ' xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
      
    End If

    'xReport.ParameterFields(3).AddCurrentValue user_name
       xReport.ParameterFields(3).AddCurrentValue WriteNo(Format(val(lbl(5).Caption), "0.00"), 0, True, ".")
     '     xReport.ParameterFields(5).AddCurrentValue WriteNo(Format(val(TxtOFRenter.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(56).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
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
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub



Private Sub Fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With Fg
  
           If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
End With
ReLineGrid
End Sub



Private Sub NourHijriCal1_LostFocus()
      If Me.TxtModFlg.Text <> "R" Then
             VBA.Calendar = vbCalGreg
           XPDtbTrans.value = ToGregorianDate(NourHijriCal1.value)
           End If
                End Sub




Private Sub Form_Load()

    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

  '  With Me.Fg
  '      .RowHeightMin = 300
  '      .WallPaper = GrdBack.Picture
  '      .AutoSize 0, .Cols - 1, False
  '  End With


'   Set TTD = New clstooltipdemand
    
    
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
Dcombos.getCountriesGovernments Me.DcbCitM
   Dcombos.getCountriesGovernments DcboGovernmentID
    Dcombos.GetBranches Me.dcBranch
    DcbWith.AddItem "„⁄ €—›… „ﬂÌ‰« "
    DcbWith.AddItem "»œÊ‰ €—›… „ﬂÌ‰« "
  
     My_SQL = "select UserID,UserName From tblUsers "
  
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
   fill_combo DCboUserName, My_SQL
    
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblQutations     Order By ID"
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
    'Label1.Visible = False

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

 


End Sub
Private Sub Form_Paint()
   ' TTD.Destroy
End Sub

Private Sub Form_Resize()
'    TTD.Destroy
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
         '   TxtAdvanceValue.Locked = True
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
         
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
               
            UnitsGrid.Clear flexClearScrollable, flexClearEverything
            UnitsGrid.Rows = 2
            UnitsGrid.Enabled = True
               Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
         
    
            Me.DCboUserName.BoundText = user_id
        
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
          
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
          
           
            XPDtbTrans.Enabled = True
          

    End Select

    Exit Sub
ErrTrap:
End Sub










Private Sub UnitsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With UnitsGrid
  
           If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
End With
ReLineGrid
End Sub
Public Function print_report1(Optional ID As Integer = 0, Optional Row As Integer = 0)
    
     
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



   MySQL = "  SELECT     dbo.TblQutationsTech.*"
MySQL = MySQL & " From dbo.TblQutationsTech"
 MySQL = MySQL & " Where (QutationsID = " & ID & ") And (QutationsIDDet =" & Row & " )"




        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepQutationTech.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepQutationTech.rpt"
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
     Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
         StrReportTitle = ""
      
    End If

    'xReport.ParameterFields(3).AddCurrentValue user_name
     '  xReport.ParameterFields(3).AddCurrentValue WriteNo(Format(val(lbl(5).Caption), "0.00"), 0, True, ".")
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

Private Sub UnitsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With Me.UnitsGrid

        Select Case .ColKey(Col)
         Case "Print"
 print_report1 val(XPTxtID), val(.TextMatrix(Row, .ColIndex("id")))
                 Case "tech"
                  LngRow = Row

 'LngCol = Col
             ' ItemProductionDate Row, Col, , 1
                Load FrmShowTech
                FrmShowTech.show

                    
                End Select
                End With
End Sub

Private Sub UnitsGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With Me.UnitsGrid

        Select Case .ColKey(Col)

   Case "Print"
    
            .ColComboList(.ColIndex("Print")) = "..."
            
                 Case "tech"
    
            .ColComboList(.ColIndex("tech")) = "..."
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
Private Sub ReLineGrid()
    Dim I As Integer
    Dim IntCounter As Integer
    Dim sumrate As Double
    sumrate = 0
   ''''///
   IntCounter = 0
   lbl(5).Caption = 0
     With UnitsGrid

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("Matrials")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("Ser")) = IntCounter
                .TextMatrix(I, .ColIndex("total")) = val(.TextMatrix(I, .ColIndex("Price"))) * val(.TextMatrix(I, .ColIndex("Count")))
             lbl(5).Caption = val(lbl(5).Caption) + val(.TextMatrix(I, .ColIndex("total")))
                  End If
                Next I

    End With
      IntCounter = 0
     With Fg

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("Des")) <> "" Then
            sumrate = sumrate + val(.TextMatrix(I, .ColIndex("Rate")))
            If sumrate > 100 Then
            MsgBox "ÌÃ» «‰ ÌﬂÊ‰ „Ã„Ê⁄ «·‰”» 100%"
            .TextMatrix(I, .ColIndex("Rate")) = 0
            Exit Sub
            End If
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("Ser")) = IntCounter
             
                  End If
                Next I

    End With
    End Sub
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim I As Integer
    Dim StrSQL As String
  Dim ContactTime As Date
UnitsGrid.Clear flexClearScrollable, flexClearEverything
            UnitsGrid.Rows = 2
            UnitsGrid.Enabled = True
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
         
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
    NourHijriCal1.value = IIf(IsNull(rs("RecordDateH").value), "", rs("RecordDateH").value)
    dcBranch.BoundText = val(IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value))
    DcboGovernmentID.BoundText = val(IIf(IsNull(rs("CityID").value), "", rs("CityID").value))
    Me.DcbCitM.BoundText = val(IIf(IsNull(rs("CityIDH").value), "", rs("CityIDH").value))
    Me.TxtZibCode.Text = IIf(IsNull(rs("ZibCode").value), "", rs("ZibCode").value)
    Me.TxtXY.Text = IIf(IsNull(rs("Xy").value), "", rs("Xy").value)
    Me.TXtElevatorCount.Text = IIf(IsNull(rs("ElevatorCount").value), "", rs("ElevatorCount").value)
    Me.TxtCustomer.Text = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
    Me.TxtProject.Text = IIf(IsNull(rs("ProjectName").value), "", rs("ProjectName").value)
    Me.TxtSupply.Text = val(IIf(IsNull(rs("Supply").value), 0, rs("Supply").value))
    Me.TxtGranty.Text = val(IIf(IsNull(rs("Granty").value), 0, rs("Granty").value))
    Me.TxtInstal.Text = val(IIf(IsNull(rs("Instal").value), 0, rs("Instal").value))
    Me.TxtShow.Text = val(IIf(IsNull(rs("Show").value), 0, rs("Show").value))
    Me.DcbInstalDMY.ListIndex = val(IIf(IsNull(rs("InstalDMY").value), -1, rs("InstalDMY").value))
    Me.DcbSupplyDMY.ListIndex = val(IIf(IsNull(rs("SupplyDMY").value), -1, rs("SupplyDMY").value))
    Me.DcbGrantyDMY.ListIndex = val(IIf(IsNull(rs("GrantyDMY").value), -1, rs("GrantyDMY").value))
    Me.DcbShowDMY.ListIndex = val(IIf(IsNull(rs("ShowDMY").value), -1, rs("ShowDMY").value))
    Me.DcbWith.ListIndex = val(IIf(IsNull(rs("Without").value), -1, rs("Without").value))
 
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
     ' If IsNull(rs("posted").value) Then
       '                                            If SystemOptions.UserInterface = ArabicInterface Then
        '                                            Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
          '                                        Else
                                                 '   Accredit.Caption = " send to Approval   "
               ''                                End If
                                     '          Accredit.Enabled = True
'  Else
                                      '             If SystemOptions.UserInterface = ArabicInterface Then
                                        '            Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                        '          Else
                                                  '  Accredit.Caption = " sent to Approval   "
                                             '  End If
                                             '  Accredit.Enabled = False
  ' End If
   
   
    Set RsDetails = New ADODB.Recordset
StrSQL = " SELECT     QutationsID, ID, Matrials, [Count], loader, CountPark, Spead, Price, DataTech ,Type "
StrSQL = StrSQL & " From dbo.TblQutationsDetls1"
StrSQL = StrSQL & " Where (QutationsID = " & val(XPTxtID.Text) & ") and type = 0"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.UnitsGrid
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails.RecordCount

        For I = .FixedRows To .Rows - 1
   'IIf(IsNull(RsDetails("NameStatus").value), "", RsDetails("NameStatus").value) = 1
            .TextMatrix(I, .ColIndex("Ser")) = I
            .TextMatrix(I, .ColIndex("id")) = val(IIf(IsNull(RsDetails("ID").value), "", RsDetails("ID").value))
            .TextMatrix(I, .ColIndex("Matrials")) = IIf(IsNull(RsDetails("Matrials").value), "", RsDetails("Matrials").value)
            .TextMatrix(I, .ColIndex("Count")) = IIf(IsNull(RsDetails("Count").value), 0, RsDetails("Count").value)
            .TextMatrix(I, .ColIndex("loader")) = IIf(IsNull(RsDetails("loader").value), "", RsDetails("loader").value)
            .TextMatrix(I, .ColIndex("CountPark")) = IIf(IsNull(RsDetails("CountPark").value), "", RsDetails("CountPark").value)
            .TextMatrix(I, .ColIndex("Spead")) = IIf(IsNull(RsDetails("Spead").value), "", RsDetails("Spead").value)
            .TextMatrix(I, .ColIndex("Price")) = IIf(IsNull(RsDetails("Price").value), 0, RsDetails("Price").value)
            .TextMatrix(I, .ColIndex("DataTech")) = IIf(IsNull(RsDetails("DataTech").value), "", RsDetails("DataTech").value)
            .TextMatrix(I, .ColIndex("total")) = val(IIf(IsNull(RsDetails("Price").value), 0, RsDetails("Price").value)) * val(IIf(IsNull(RsDetails("Count").value), 0, RsDetails("Count").value))
                         
             '  End If
            RsDetails.MoveNext
         
        Next I
    'ReLineGridCount
    ReLineGrid
    
End With

    End If

    RsDetails.Close
    Set RsDetails = Nothing
    Set RsDetails = New ADODB.Recordset
StrSQL = "SELECT     ID, QutationsID, Rate, Des ,type "
StrSQL = StrSQL & " From dbo.TblQutationsDetls1"
StrSQL = StrSQL & " Where (QutationsID = " & val(XPTxtID.Text) & ") and type = 1"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.Fg
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails.RecordCount

        For I = .FixedRows To .Rows - 1
   'IIf(IsNull(RsDetails("NameStatus").value), "", RsDetails("NameStatus").value) = 1
            .TextMatrix(I, .ColIndex("Ser")) = I
            .TextMatrix(I, .ColIndex("id")) = val(IIf(IsNull(RsDetails("ID").value), "", RsDetails("ID").value))
            .TextMatrix(I, .ColIndex("Des")) = IIf(IsNull(RsDetails("Des").value), "", RsDetails("Des").value)
            .TextMatrix(I, .ColIndex("Rate")) = val(IIf(IsNull(RsDetails("Rate").value), 0, RsDetails("Rate").value))
                                   
             '  End If
            RsDetails.MoveNext
         
        Next I
    'ReLineGridCount
    ReLineGrid
    
End With

    End If

    RsDetails.Close
    Set RsDetails = Nothing
    
    '//////////////////////////////////////////
  '  Set RsDetails1 = New ADODB.Recordset



  '  RsDetails1.Close
  '  Set RsDetails1 = Nothing
    
  '  fillapprovData
    'ReLineGrid
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
Dim RsDetails2 As ADODB.Recordset
  Dim astrSplit2tems2() As String
Dim astrSplitItems() As String
Dim j As Integer
  Dim st As String
    Dim nElements As Integer
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim I As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
 
   If Me.DcboGovernmentID.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ   «”„ «·„œÌ‰…!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.DcboGovernmentID.SetFocus
           ' SendKeys "{F4}"
            Exit Sub
        End If
    If Me.TxtCustomer.Text = "" Then
            Msg = "Ì—ÃÏ «œŒ«· «”„ «·⁄„Ì·     !! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtCustomer.SetFocus
           ' SendKeys "{F4}"
            Exit Sub
        End If
         If Me.TxtProject.Text = "" Then
            Msg = "ÌÃ» «œŒ«· «”„ «·„‘—Ê⁄    !! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtProject.SetFocus
           ' SendKeys "{F4}"
            Exit Sub
        End If

        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then

            XPTxtID.Text = CStr(new_id("TblQutations", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
            StrSQL = "Delete From TblQutationsDetls1 Where QutationsID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
      '      StrSQL = "Delete From TblQutationsDetls2 Where QutationsID=" & val(Me.XPTxtID.text)
      '      Cn.Execute StrSQL, , adExecuteNoRecords
             StrSQL = "Delete From TblQutationsTech Where QutationsID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If
     
   

   
       rs("ID").value = val(XPTxtID.Text)
        rs("RecordDate").value = XPDtbTrans.value
       rs("RecordDateH").value = Me.NourHijriCal1.value
       
       rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
       rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)
       rs("CityID").value = IIf(Me.DcboGovernmentID.BoundText = "", Null, Me.DcboGovernmentID.BoundText)
       rs("CityIDH").value = IIf(Me.DcbCitM.BoundText = "", Null, Me.DcbCitM.BoundText)
       rs("ElevatorCount").value = val(Me.TXtElevatorCount.Text)
       rs("Xy").value = Me.TxtXY.Text
       rs("ZibCode").value = Me.TxtZibCode.Text
       rs("ProjectName").value = Me.TxtProject.Text
       rs("CusName").value = Me.TxtCustomer.Text
       rs("Supply").value = val(Me.TxtSupply.Text)
       rs("Granty").value = val(Me.TxtGranty.Text)
       rs("Instal").value = val(Me.TxtInstal.Text)
       rs("Show").value = val(Me.TxtShow.Text)
       rs("Without").value = IIf(Me.DcbWith.ListIndex = -1, Null, Me.DcbWith.ListIndex)
       
       rs("InstalDMY").value = IIf(Me.DcbInstalDMY.ListIndex = -1, Null, Me.DcbInstalDMY.ListIndex)
       rs("SupplyDMY").value = IIf(Me.DcbSupplyDMY.ListIndex = -1, Null, Me.DcbSupplyDMY.ListIndex)
       rs("GrantyDMY").value = IIf(Me.DcbGrantyDMY.ListIndex = -1, Null, Me.DcbGrantyDMY.ListIndex)
       rs("ShowDMY").value = IIf(Me.DcbShowDMY.ListIndex = -1, Null, Me.DcbShowDMY.ListIndex)
       rs.update
        '''''''''/////////////////////////////////
               Set RsDetails2 = New ADODB.Recordset
     StrSQL = "SELECT     *  from dbo.TblQutationsTech Where (1 = -1)"
   RsDetails2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
      Set RsDetails = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblQutationsDetls1 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

          
       For I = Me.UnitsGrid.FixedRows To UnitsGrid.Rows - 1
   If (UnitsGrid.TextMatrix(I, UnitsGrid.ColIndex("Matrials"))) <> "" Then
    RsDetails.AddNew
                RsDetails("QutationsID").value = val(XPTxtID.Text)
                RsDetails("type").value = 0
                RsDetails("Count").value = val(UnitsGrid.TextMatrix(I, UnitsGrid.ColIndex("Count")))
                RsDetails("loader").value = val(UnitsGrid.TextMatrix(I, UnitsGrid.ColIndex("loader")))
                RsDetails("CountPark").value = val(UnitsGrid.TextMatrix(I, UnitsGrid.ColIndex("CountPark")))
                RsDetails("Spead").value = val(UnitsGrid.TextMatrix(I, UnitsGrid.ColIndex("Spead")))
                RsDetails("Matrials").value = UnitsGrid.TextMatrix(I, UnitsGrid.ColIndex("Matrials"))
                RsDetails("Price").value = val(UnitsGrid.TextMatrix(I, UnitsGrid.ColIndex("Price")))
                RsDetails("DataTech").value = UnitsGrid.TextMatrix(I, UnitsGrid.ColIndex("DataTech"))
                RsDetails.update
                         If UnitsGrid.TextMatrix(I, UnitsGrid.ColIndex("DataTech")) <> "" Then
          st = UnitsGrid.TextMatrix(I, UnitsGrid.ColIndex("DataTech"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For j = 0 To nElements - 1
          RsDetails2.AddNew
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         RsDetails2("QutationsID").value = val(XPTxtID.Text)
         RsDetails2("QutationsIDDet").value = IIf(IsNull(RsDetails("ID").value), Null, RsDetails("ID").value)
         RsDetails2("ElevatorCount").value = val(astrSplit2tems2(0))
         RsDetails2("ParkingCount").value = val(astrSplit2tems2(1))
         RsDetails2("DoorCount").value = val(astrSplit2tems2(2))
         RsDetails2("DorDirect").value = val(astrSplit2tems2(3))
         RsDetails2("DorNo").value = val(astrSplit2tems2(4))
         RsDetails2("WellWidth").value = val(astrSplit2tems2(5))
         RsDetails2("WellDepth").value = val(astrSplit2tems2(6))
         RsDetails2("BilldWell").value = val(astrSplit2tems2(7))
         RsDetails2("space").value = val(astrSplit2tems2(8))
         RsDetails2("WellHigh").value = val(astrSplit2tems2(9))
         RsDetails2("Deptdrilled").value = val(astrSplit2tems2(10))
         RsDetails2("HighAllWell").value = val(astrSplit2tems2(11))
         RsDetails2("EngineRoom").value = val(astrSplit2tems2(12))
         RsDetails2("LoadKG").value = val(astrSplit2tems2(13))
         RsDetails2("LoadPerson").value = val((astrSplit2tems2(14)))
         RsDetails2("Spead").value = val(astrSplit2tems2(15))
         RsDetails2("Hoist").value = val(astrSplit2tems2(16))
         RsDetails2("ElectricMotor").value = val(astrSplit2tems2(17))
         RsDetails2("OperatingMethod").value = val(astrSplit2tems2(18))
         RsDetails2("OperatingMethod2").value = val(astrSplit2tems2(19))
         RsDetails2("CabinWidth").value = val(astrSplit2tems2(20))
         RsDetails2("CabinDepth").value = val(astrSplit2tems2(21))
         RsDetails2("CabinHigh").value = val(astrSplit2tems2(22))
         RsDetails2("CabinDor").value = val((astrSplit2tems2(23)))
         RsDetails2("DOrFloor").value = val(astrSplit2tems2(24))
         RsDetails2("DorWidth").value = val(astrSplit2tems2(25))
         RsDetails2("DorHight").value = val(astrSplit2tems2(26))
         RsDetails2("Walls").value = astrSplit2tems2(27)
        
         RsDetails2.update
         Next j
                  
          
          End If
    
           End If
        Next I
        
        '''''''''''''''//////////////////////////
          Set RsDetails = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblQutationsDetls1 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

          
       For I = Me.Fg.FixedRows To Fg.Rows - 1
   If (Fg.TextMatrix(I, Fg.ColIndex("Des"))) <> "" Then
    RsDetails.AddNew
                RsDetails("QutationsID").value = val(XPTxtID.Text)
                 RsDetails("type").value = 1
                RsDetails("Rate").value = val(Fg.TextMatrix(I, Fg.ColIndex("Rate")))
                RsDetails("Des").value = Fg.TextMatrix(I, Fg.ColIndex("Des"))
                RsDetails.update
    
           End If
        Next I
        
 
      

    
        Cn.CommitTrans
        BeginTrans = False
    '    RsDetails.Close
        Set RsDetails = Nothing
        Set RsDetails1 = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    
                    Exit Sub
                    Else
                    Retrive val(Me.XPTxtID.Text)
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Retrive val(Me.XPTxtID.Text)
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
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
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
  Dim StrSQL1 As String
    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From TblQutations Where ID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
' StrSQL1 = "Delete From TblQutationsDetls2 Where QutationsID=" & val(Me.XPTxtID.text)
'            Cn.Execute StrSQL1, , adExecuteNoRecords
             StrSQL = "Delete From TblQutationsDetls1 Where QutationsID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
             StrSQL = "Delete From TblQutationsTech Where QutationsID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

                If rs.RecordCount < 1 Then
             
            UnitsGrid.Clear flexClearScrollable, flexClearEverything
            UnitsGrid.Rows = 2
            UnitsGrid.Enabled = True
              Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
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

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, " ⁄—Ê÷ «·«”⁄«—··„’«⁄œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ⁄—Ê÷ «·«”⁄«—··„’«⁄œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
       .Create Me.hWnd, " ⁄—Ê÷ «·«”⁄«—··„’«⁄œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ⁄—Ê÷ «·«”⁄«—··„’«⁄œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
       .Create Me.hWnd, " ⁄—Ê÷ «·«”⁄«—··„’«⁄œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ⁄—Ê÷ «·«”⁄«—··„’«⁄œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
       .Create Me.hWnd, " ⁄—Ê÷ «·«”⁄«—··„’«⁄œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ⁄—Ê÷ «·«”⁄«—··„’«⁄œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
       .Create Me.hWnd, " ⁄—Ê÷ «·«”⁄«—··„’«⁄œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
      .Create Me.hWnd, " ⁄—Ê÷ «·«”⁄«—··„’«⁄œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
       .Create Me.hWnd, " ⁄—Ê÷ «·«”⁄«—··„’«⁄œ", 1, 15204351, -2147483630
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
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
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


Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.Text <> "R" Then
         NourHijriCal1.value = ToHijriDate(XPDtbTrans.value)
End If
End Sub

