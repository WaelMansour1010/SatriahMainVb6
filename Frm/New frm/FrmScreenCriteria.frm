VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmScreenCriteria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  ⁄—Ì› „Õœœ«  «·‘«‘«  "
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14040
   Icon            =   "FrmScreenCriteria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   14040
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame2 
      Caption         =   "ﬁÌ„ «·„Õœœ «·„ ⁄Ì—…"
      Height          =   2415
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   600
      Width           =   6735
      Begin VB.TextBox TxtSubDese 
         Alignment       =   1  'Right Justify
         Height          =   555
         Left            =   0
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   1200
         Width           =   5895
      End
      Begin VB.TextBox TxtSubDes 
         Alignment       =   1  'Right Justify
         Height          =   555
         Left            =   0
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox TxtSubValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   240
         Width           =   1215
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   20
         Left            =   1185
         TabIndex        =   46
         Top             =   1920
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   688
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
         ButtonImage     =   "FrmScreenCriteria.frx":000C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   21
         Left            =   240
         TabIndex        =   47
         Top             =   1920
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   688
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
         ButtonImage     =   "FrmScreenCriteria.frx":03A6
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ê’› «‰Ã·Ì“Ì"
         Height          =   405
         Index           =   10
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ê’› ⁄—»Ì"
         Height          =   405
         Index           =   11
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ﬁÌ„…"
         Height          =   255
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   240
         Width           =   495
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13995
      _cx             =   24686
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
      Caption         =   ""
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
         Left            =   -990
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   0
         Left            =   1185
         TabIndex        =   17
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
         ButtonImage     =   "FrmScreenCriteria.frx":0940
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
         TabIndex        =   18
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
         ButtonImage     =   "FrmScreenCriteria.frx":0CDA
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
         TabIndex        =   19
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
         ButtonImage     =   "FrmScreenCriteria.frx":1074
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
         TabIndex        =   20
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
         ButtonImage     =   "FrmScreenCriteria.frx":140E
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "  ⁄—Ì› „Õœœ«  «·‘«‘«   "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   615
         Left            =   3960
         TabIndex        =   14
         Top             =   0
         Width           =   9495
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   9210
      TabIndex        =   2
      Top             =   6600
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
      Left            =   8370
      TabIndex        =   3
      Top             =   6600
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
      Left            =   7515
      TabIndex        =   4
      Top             =   6600
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
      Left            =   6705
      TabIndex        =   5
      Top             =   6600
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
      Left            =   5895
      TabIndex        =   6
      Top             =   6600
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
      Left            =   5070
      TabIndex        =   7
      Top             =   6600
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
      Left            =   1200
      TabIndex        =   8
      Top             =   8400
      Visible         =   0   'False
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
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   4305
      Left            =   15240
      TabIndex        =   13
      Top             =   2280
      Width           =   9435
      _cx             =   16642
      _cy             =   7594
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
      BackColor       =   12648447
      ForeColor       =   -2147483630
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "„” ÊÌ«  «·«⁄ „«œ|»Ì«‰«  «·„” Œœ„Ì‰"
      Align           =   0
      CurrTab         =   1
      FirstTab        =   0
      Style           =   3
      Position        =   6
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   4215
         Left            =   45
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   45
         Width           =   8010
         _cx             =   14129
         _cy             =   7435
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   4215
         Left            =   -9990
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   45
         Width           =   8010
         _cx             =   14129
         _cy             =   7435
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
      End
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   5415
      Index           =   1
      Left            =   6720
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   600
      Width           =   7260
      _cx             =   12806
      _cy             =   9551
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
      Begin VB.Frame Frame1 
         Caption         =   "«·„Õœœ« "
         Height          =   615
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1560
         Width           =   4215
         Begin VB.ComboBox CBOCriteria 
            Height          =   315
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox Txtvalue 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬁÌ„…"
            Height          =   255
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   45
         Index           =   1
         Left            =   13215
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   3060
         Width           =   1275
      End
      Begin VB.TextBox Shifttime 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1920
         Width           =   6165
      End
      Begin VB.TextBox XPMTxtRemark 
         Alignment       =   1  'Right Justify
         Height          =   2595
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   2400
         Width           =   5895
      End
      Begin VB.TextBox XPTxtName 
         Alignment       =   1  'Right Justify
         Height          =   420
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox XPTxtSheftID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3975
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   150
         Width           =   2055
      End
      Begin VB.TextBox XPTxtNameE 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   135
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1020
         Width           =   5895
      End
      Begin MSComCtl2.DTPicker ShfitFrom 
         Height          =   450
         Left            =   11235
         TabIndex        =   28
         Top             =   780
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   794
         _Version        =   393216
         CustomFormat    =   "'Time: 'hh:mm tt"
         Format          =   77594627
         UpDown          =   -1  'True
         CurrentDate     =   40909
      End
      Begin MSComCtl2.DTPicker ShfitTo 
         Height          =   375
         Left            =   12555
         TabIndex        =   29
         Top             =   1350
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "'Time: 'hh:mm tt"
         Format          =   77594627
         UpDown          =   -1  'True
         CurrentDate     =   40909
      End
      Begin VB.Label Lb_note_value_by_characters 
         Alignment       =   1  'Right Justify
         Height          =   480
         Left            =   11985
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   4125
         Width           =   7665
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Code"
         Height          =   195
         Left            =   11400
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   3060
         Width           =   1410
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄·Ìﬁ:"
         Height          =   105
         Index           =   9
         Left            =   11595
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Tag             =   "22"
         Top             =   270
         Width           =   2895
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï"
         Height          =   360
         Index           =   7
         Left            =   10860
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1350
         Width           =   2010
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰"
         Height          =   300
         Index           =   6
         Left            =   11340
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   780
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·”«⁄« "
         Height          =   285
         Index           =   5
         Left            =   12180
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1920
         Width           =   2010
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·„Õœœ"
         Height          =   285
         Index           =   0
         Left            =   6285
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   150
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„·«ÕŸ« "
         Height          =   285
         Index           =   1
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   2370
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„ ⁄—»Ì"
         Height          =   285
         Index           =   3
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„ «‰Ã·Ì“Ì"
         Height          =   285
         Index           =   8
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1020
         Width           =   1035
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Grid 
      Height          =   3045
      Left            =   0
      TabIndex        =   44
      Top             =   3000
      Width           =   6735
      _cx             =   11880
      _cy             =   5371
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
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmScreenCriteria.frx":17A8
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2295
      Left            =   9600
      Top             =   3480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   6270
      Width           =   825
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   6270
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   2
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   6240
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   4
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   6270
      Width           =   1155
   End
End
Attribute VB_Name = "FrmScreenCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip

 

 
 
 

 
Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
         
 
 
          
        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

 
      
            TxtModFlg.text = "E"
Grid.Rows = Grid.Rows + 1
        Case 2
            SaveData

        Case 3
            Undo

        Case 4
            '  If DoPremis(Do_Delete, Me.name, True) = False Then
            '      Exit Sub
            '  End If
            Del_Company

        Case 5

        Case 6
            Unload Me
            Case 20
            addrow
            Case 21
            RemoveGridRow
            
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Activate()
    'XPTxtname.SetFocus

End Sub

Function addrow()
 Dim i As Integer
 
      If Grid.Rows = 1 Then Grid.Rows = 2
         With Grid
  i = .Rows
 
               .TextMatrix(i - 1, .ColIndex("SubValue")) = val(TxtSubValue)
                .TextMatrix(i - 1, .ColIndex("SubDes")) = TxtSubDes
                                .TextMatrix(i - 1, .ColIndex("SubDesE")) = TxtSubDese.text
                                
                  .Rows = .Rows + 1
                  TxtSubDes = ""
                  TxtSubValue = ""
                  TxtSubDese = ""
      '       .AutoSize 0, .Cols - 1, False
   
    End With
 
     
    ReLineGrid

End Function



Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If TxtSubValue.text <> "" Or TxtSubDes.text <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

     

End Sub



Private Sub RemoveGridRow()
      If Grid.Rows = 1 Then Exit Sub
    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub



Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If XPTxtSheftID.text <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  «·‘Ì›  —ﬁ„ " & Chr(13)
        Msg = Msg + (XPTxtSheftID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                rs.delete
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

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·»‰ﬂ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
   ' CmdRemove.Caption = "Delete"
   Label4.Caption = "Definition of the Determinants of Screens"
   Me.Caption = Label4.Caption
   lbl(0).Caption = "Code"
   lbl(3).Caption = "Name Arabic"
   lbl(8).Caption = "Name English"
   Frame1.Caption = "Determinants"
   Label1.Caption = "Value"
   Frame2.Caption = "Changing value Determinant"
   Label2.Caption = "Value"
   lbl(10).Caption = "Description English"
   lbl(11).Caption = "Description Arabic"
   lbl(1).Caption = "Remarks"
   Cmd(20).Caption = "Add"
   Cmd(21).Caption = "Remove"
   With Grid
   .TextMatrix(0, .ColIndex("Ser")) = "Serial"
   .TextMatrix(0, .ColIndex("SubValue")) = "Value"
   .TextMatrix(0, .ColIndex("SubDes")) = "Description Arabic"
  .TextMatrix(0, .ColIndex("SubDesE")) = "Description English"
      End With
   ' TabMain.CurrTab = 1
   ' lbl(5).Caption = "Hour"
  '  Me.Caption = "Approval Levels "
   ' Label4.Caption = Me.Caption
    'EleHeader.Caption = Me.Caption
    'lbl(0).Caption = "Code"
    'lbl(3).Caption = " Level Name"
    'lbl(1).Caption = "Remarks"
    lbl(2).Caption = "Current Record"
    lbl(4).Caption = "NO. Recordes"
    'TabMain.TabCaption(1) = "Levels"
    'TabMain.TabCaption(0) = "User"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    'Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"

With CBOCriteria
.Clear

.AddItem "Greater than"
.AddItem "Greater than or equal"
.AddItem "Less than"
.AddItem "Less than or equal"
.AddItem "Equal"
.AddItem "Not Equal"
.AddItem "JustOption"
End With

 
'    Frame10.Caption = "Employees"

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    XPTxtSheftID.text = IIf(IsNull(rs("CriteriaID").value), "", val(rs("CriteriaID").value))
    XPTxtName.text = IIf(IsNull(rs("name").value), "", Trim(rs("name").value))
    XPTxtNamee.text = IIf(IsNull(rs("namee").value), "", Trim(rs("namee").value))
TxtValue.text = IIf(IsNull(rs("value").value), "", Trim(rs("value").value))
    XPMTxtRemark.text = IIf(IsNull(rs("Remarks").value), "", Trim(rs("Remarks").value))
CBOCriteria.ListIndex = IIf(IsNull(rs("typeid").value), -1, (rs("typeid").value))
 
  
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    
    
      Dim RsDetails As New ADODB.Recordset
    StrSQL = "Select * From  tblScreenCriteriaValues Where CriteriaID=" & val(XPTxtSheftID.text)
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  With Grid
     .Clear flexClearScrollable, flexClearEverything
     .Rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
         .Rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .Rows - 1
             .TextMatrix(i, .ColIndex("Ser")) = i
             .TextMatrix(i, .ColIndex("SubValue")) = IIf(IsNull(RsDetails("SubValue").value), "", RsDetails("SubValue").value)
              .TextMatrix(i, .ColIndex("SubDes")) = IIf(IsNull(RsDetails("SubDes").value), "", RsDetails("SubDes").value)
              .TextMatrix(i, .ColIndex("SubDesE")) = IIf(IsNull(RsDetails("SubDesE").value), "", RsDetails("SubDesE").value)
              
            RsDetails.MoveNext
        Next i

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
    
    
    Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "CriteriaID='" & val(XPTxtSheftID.text) & "'", , adSearchForward, adBookmarkFirst

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
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
        If XPTxtName.text = "" Then
    
            MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·„Õœœ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtName.SetFocus
            Exit Sub
        End If

  
      If CBOCriteria.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " ÌÃ»  ÕœÌœ  «·⁄·«ﬁ…" & Chr(13)
            Else
                Msg = "Select Operand" & Chr(13)
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CBOCriteria.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
       If val(TxtValue.text) = 0 Then
    If CBOCriteria.ListIndex < 6 Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· ﬁÌ„… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtValue.SetFocus
            Exit Sub
    Else
    
    End If
    
        End If
               
        
        Select Case Me.TxtModFlg.text

            Case "N"
            
                StrSQL = "select * From  tblScreenCriteria where name='" & Trim(XPTxtName.text) & "'"
                RsTemp.Open StrSQL, Cn.ConnectionString, adOpenStatic, adLockOptimistic, adCmdText
 
                '    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    Msg = "Â‰«ﬂ „” ÊÏ „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·‘Ì› "
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtName.SetFocus
                    Exit Sub
                End If
   XPTxtSheftID.text = CStr(new_id("tblScreenCriteria", "CriteriaID", "", True))
            Case "E"
            
                StrSQL = "select * From  tblScreenCriteria where name='" & Trim(XPTxtName.text) & "'"
                RsTemp.Open StrSQL, Cn.ConnectionString, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("CriteriaID").value <> val(XPTxtSheftID.text) Then
                        Msg = "Â‰«ﬂ „” ÊÏ  „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·‘Ì› "
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        XPTxtName.SetFocus
                        Exit Sub
                    End If
                End If

        End Select

        '    Cn.BeginTrans
        '    BeginTrans = True
        Select Case Me.TxtModFlg.text

            Case "N"
                rs.AddNew
                rs("CriteriaID").value = val(XPTxtSheftID.text)
                Case "E"
                Cn.Execute "delete tblScreenCriteriaValues where CriteriaID=" & val(XPTxtSheftID.text)
        End Select

        rs("name").value = Trim(XPTxtName.text)
          rs("value").value = val(TxtValue.text)
        rs("namee").value = Trim(XPTxtNamee.text)
        rs("Remarks").value = IIf(XPMTxtRemark.text = "", "", Trim(XPMTxtRemark.text))
       rs("typeid").value = CBOCriteria.ListIndex
        rs.update
    
        '   »Ì«‰«  «·ﬁÌ„
 
    Dim RsDetails As ADODB.Recordset
 
 

        Set RsDetails = New ADODB.Recordset
        RsDetails.Open "tblScreenCriteriaValues", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
 
With Grid
        For i = .FixedRows To .Rows - 1
        If .TextMatrix(i, .ColIndex("SubDes")) <> "" Or .TextMatrix(i, .ColIndex("SubValue")) <> "" Then
            RsDetails.AddNew
           RsDetails("CriteriaID").value = val(XPTxtSheftID.text)
            RsDetails("SubValue").value = val(.TextMatrix(i, .ColIndex("SubValue")))
            RsDetails("SubDes").value = .TextMatrix(i, .ColIndex("SubDes"))
            RsDetails("SubDesE").value = .TextMatrix(i, .ColIndex("SubDesE"))
            
             
            RsDetails.update
            End If
        Next i
 End With

   

        '    Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ «·»Ì«‰«  " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ⁄„·Ì… ÃœÌœ… ‰⁄„ «Ê ·« ?"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
                MsgBox " „ Õ›Ÿ «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = Chr(13) + Chr(10)

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

    If my_language = "E" Then
        SetInterface Me
        ChangeLang
    End If

    ScreenNameArabic = " ⁄—Ì› „Õœœ«  «·‘«‘« "
    ScreenNameEnglish = " Screens Creteria Def "
    'RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"

With CBOCriteria
.AddItem "«ﬂ»— „‰"
.AddItem "«ﬂ»— „‰ «Ê Ì”«ÊÌ"
.AddItem "«ﬁ· „‰"
.AddItem "«ﬁ· „‰ «Ê Ì”«ÊÌ"
.AddItem "Ì”«ÊÌ"
.AddItem "·« Ì”«ÊÌ"
.AddItem "„Ê«’›« "
End With

Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
    'rs.Open "tblScreenCriteria", cn.ConnectionString , adOpenStatic, adLockOptimistic, adCmdTable
    rs.CursorLocation = adUseClient
    StrSQL = "select * From tblScreenCriteria"
    rs.Open StrSQL, Cn.ConnectionString, adOpenStatic, adLockOptimistic, adCmdText
    If SystemOptions.UserInterface = EnglishInterface Then
SetInterface Me
        ChangeLang
        End If
    Me.TxtModFlg.text = "R"
    XPBtnMove_Click 2

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

                'btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "  Õ›Ÿ ‘«‘… " & ScreenNameArabic & Chr(13) & " ﬂÊœ   " & XPTxtSheftID.text & Chr(13) & " «·«”„ " & XPTxtName.text & Chr(13) & "„‰  " & ShfitFrom.value & Chr(13) & " «·Ì  " & ShfitTo.value & Chr(13) & " ⁄œœ «·”«⁄«    " & Shifttime & Chr(13) & " „·«ÕŸ«    " & XPMTxtRemark
                     
    LogTextE = "  Save Screen  " & ScreenNameEnglish & Chr(13) & " Code   " & XPTxtSheftID.text & Chr(13) & " name " & XPTxtName.text & Chr(13) & "From  " & ShfitFrom.value & Chr(13) & " To  " & ShfitTo.value & Chr(13) & "  No Of Hour" & Shifttime & Chr(13) & " Remarks   " & XPMTxtRemark
                     
    If Currentmode <> "D" Then
        '      AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, "", ""
    Else
        '  AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D", "", ""
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    'RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                '            Me.Caption = "»Ì«‰«  «·‘Ì› "
            Else
                Me.Caption = "sheft Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.XPTxtSheftID.locked = True
            Me.XPTxtName.locked = True
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
                '            Me.Caption = "»Ì«‰«  «·»‰Êﬂ(ÃœÌœ)"
            Else
                Me.Caption = "Banks Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            'Me.XPBtnMove(0).Enabled = False
            'Me.XPBtnMove(1).Enabled = False
            'Me.XPBtnMove(2).Enabled = False
            'Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtSheftID.locked = True
            Me.XPTxtName.locked = False
            Me.XPMTxtRemark.locked = False

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                '            Me.Caption = "»Ì«‰«  «·»‰Êﬂ(  ⁄œÌ· )"
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
        
            Me.XPTxtSheftID.locked = True
            Me.XPTxtName.locked = False
            Me.XPMTxtRemark.locked = False
    End Select

    Exit Sub
ErrTrap:

End Sub

 

 

Private Sub TxtSubDes_GotFocus()
  SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtSubDese_GotFocus()
 SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub TxtSubValue_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 0)
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 0)
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
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

