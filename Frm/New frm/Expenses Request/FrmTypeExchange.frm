VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmTypeExchange 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÿ·» ’—›  "
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16140
   Icon            =   "FrmTypeExchange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   16140
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   18930
      TabIndex        =   27
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   19830
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   19710
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   19470
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   19710
      TabIndex        =   23
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
      Width           =   16245
      _cx             =   28654
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
      Caption         =   " ÿ·» ’—›  "
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
         ButtonImage     =   "FrmTypeExchange.frx":038A
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
         ButtonImage     =   "FrmTypeExchange.frx":0724
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
         ButtonImage     =   "FrmTypeExchange.frx":0ABE
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
         ButtonImage     =   "FrmTypeExchange.frx":0E58
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
         Left            =   3960
         Picture         =   "FrmTypeExchange.frx":11F2
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
      Left            =   3870
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8100
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
         Left            =   7920
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
         Left            =   7080
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
         Left            =   6255
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
         Left            =   5400
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
         Left            =   4545
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
         Left            =   3720
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
         Left            =   2880
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
      Begin ImpulseButton.ISButton CmdAttach 
         Height          =   375
         Left            =   1920
         TabIndex        =   32
         Top             =   60
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
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   6780
      TabIndex        =   13
      Top             =   7560
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   18990
      TabIndex        =   28
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
      Left            =   19350
      TabIndex        =   29
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
   Begin ImpulseButton.ISButton Accredit 
      Height          =   390
      Left            =   2280
      TabIndex        =   33
      Top             =   8160
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   688
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   6975
      Left            =   0
      TabIndex        =   34
      Top             =   480
      Width           =   16185
      _cx             =   28549
      _cy             =   12303
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
      Caption         =   "«·»Ì«‰« |Õ«·Â «·«⁄ „«œ|»Ì«‰«   Õ·Ì·Ì…"
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
      Picture(0)      =   "FrmTypeExchange.frx":4E5A
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   6510
         Left            =   17130
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   45
         Width           =   16095
         _cx             =   28390
         _cy             =   11483
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
         Begin VSFlex8UCtl.VSFlexGrid Fg 
            Height          =   6165
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   15855
            _cx             =   27966
            _cy             =   10874
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
            Rows            =   1
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmTypeExchange.frx":51F4
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6510
         Left            =   16830
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   45
         Width           =   16095
         _cx             =   28390
         _cy             =   11483
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
            Height          =   5910
            Left            =   120
            TabIndex        =   36
            Tag             =   "1"
            Top             =   120
            Width           =   15870
            _cx             =   27993
            _cy             =   10425
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
            FormatString    =   $"FrmTypeExchange.frx":531B
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
            Begin VB.Frame Frame6 
               Height          =   3615
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   2160
               Visible         =   0   'False
               Width           =   7695
               Begin VB.CommandButton Command7 
                  BackColor       =   &H000000FF&
                  Caption         =   "X"
                  Height          =   255
                  Left            =   7320
                  Style           =   1  'Graphical
                  TabIndex        =   104
                  Top             =   0
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
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
                  Height          =   3420
                  Index           =   22
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   240
                  Width           =   7575
               End
               Begin VB.Shape Shape5 
                  BorderWidth     =   2
                  Height          =   3375
                  Left            =   120
                  Top             =   240
                  Width           =   7575
               End
            End
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   6120
            Width           =   7095
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6510
         Index           =   15
         Left            =   45
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   45
         Width           =   16095
         _cx             =   28390
         _cy             =   11483
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
         _GridInfo       =   $"FrmTypeExchange.frx":545E
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Height          =   6480
            Left            =   15
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   15
            Width           =   16065
            Begin VB.Frame Frame5 
               Appearance      =   0  'Flat
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H80000008&
               Height          =   4455
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   7200
               Visible         =   0   'False
               Width           =   11055
               Begin ImpulseButton.ISButton ISButton4 
                  Height          =   330
                  Left            =   10080
                  TabIndex        =   89
                  ToolTipText     =   "Õ–› «·ﬂ·"
                  Top             =   120
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈€·«ﬁ"
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
                  ButtonImage     =   "FrmTypeExchange.frx":5494
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Height          =   615
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   720
               Width           =   15615
               Begin VB.CommandButton Command2 
                  Caption         =   " ›«’Ì· «„— «·‘—«¡"
                  Height          =   255
                  Left            =   2520
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "⁄—÷ "
                  Height          =   255
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   240
                  Width           =   2175
               End
               Begin VB.TextBox txtTransaction_ID 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   2175
               End
               Begin VB.TextBox TxtOrderNo 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   9120
                  Locked          =   -1  'True
                  TabIndex        =   75
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.ComboBox cbobasedOn 
                  Height          =   315
                  ItemData        =   "FrmTypeExchange.frx":BCF6
                  Left            =   10560
                  List            =   "FrmTypeExchange.frx":BCF8
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   240
                  Width           =   3855
               End
               Begin MSComCtl2.DTPicker ReqDate 
                  Height          =   315
                  Left            =   6720
                  TabIndex        =   86
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   224002049
                  CurrentDate     =   41640
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " √—ÌŒÂ"
                  Height          =   255
                  Index           =   17
                  Left            =   8280
                  TabIndex        =   87
                  Top             =   240
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "»‰«¡ ⁄·Ì"
                  Height          =   285
                  Index           =   14
                  Left            =   14400
                  TabIndex        =   73
                  Top             =   240
                  Width           =   1005
               End
            End
            Begin VB.TextBox XPTxtID 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   7290
               Locked          =   -1  'True
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.ComboBox Contract_period 
               Height          =   315
               ItemData        =   "FrmTypeExchange.frx":BCFA
               Left            =   18840
               List            =   "FrmTypeExchange.frx":BD04
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   600
               Width           =   975
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·’—›"
               Height          =   3915
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1320
               Width           =   15705
               Begin VB.Frame Frame7 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "›Ì Õ«·… «·„ÊŸ›"
                  Height          =   495
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   4935
                  Begin VB.OptionButton Option5 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”·›…"
                     Height          =   195
                     Left            =   2280
                     RightToLeft     =   -1  'True
                     TabIndex        =   113
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.OptionButton Option4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«ÃÊ— „” Õﬁ…"
                     Height          =   195
                     Left            =   3360
                     RightToLeft     =   -1  'True
                     TabIndex        =   112
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1335
                  End
                  Begin VB.OptionButton Option6 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Œ’’« "
                     Height          =   195
                     Left            =   1200
                     RightToLeft     =   -1  'True
                     TabIndex        =   111
                     Top             =   240
                     Width           =   1095
                  End
                  Begin VB.OptionButton Option7 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "»œ·«  „ﬁœ„Â"
                     Height          =   195
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   110
                     Top             =   240
                     Width           =   1095
                  End
               End
               Begin VB.TextBox TxtBankIBAN 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   660
                  Width           =   1905
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFC0&
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
                  Left            =   11520
                  TabIndex        =   102
                  Top             =   960
                  Width           =   2295
               End
               Begin VB.TextBox TxtPerson 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2760
                  MultiLine       =   -1  'True
                  TabIndex        =   101
                  Top             =   600
                  Width           =   11055
               End
               Begin VB.TextBox TxtPriceE 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H0080FFFF&
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
                  Left            =   2760
                  TabIndex        =   97
                  Top             =   960
                  Width           =   2295
               End
               Begin VB.TextBox TxtCurrencyRate 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000003&
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
                  Left            =   6600
                  TabIndex        =   95
                  Top             =   960
                  Width           =   2295
               End
               Begin VB.ComboBox DCboCashType121 
                  Height          =   315
                  ItemData        =   "FrmTypeExchange.frx":BD12
                  Left            =   11640
                  List            =   "FrmTypeExchange.frx":BD14
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   240
                  Width           =   2175
               End
               Begin VB.TextBox TxtDes 
                  Alignment       =   1  'Right Justify
                  Height          =   1755
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   50
                  Top             =   2040
                  Width           =   5415
               End
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   12960
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   2400
                  Width           =   855
               End
               Begin VB.TextBox TxtME 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   12960
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   2760
                  Width           =   855
               End
               Begin VB.TextBox TxtManger 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   12960
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   3120
                  Width           =   855
               End
               Begin MSDataListLib.DataCombo DcbExchang 
                  Bindings        =   "FrmTypeExchange.frx":BD16
                  Height          =   315
                  Left            =   6240
                  TabIndex        =   51
                  Top             =   2040
                  Width           =   7575
                  _ExtentX        =   13361
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
               Begin MSDataListLib.DataCombo DcbEmp 
                  Bindings        =   "FrmTypeExchange.frx":BD2B
                  Height          =   315
                  Left            =   6240
                  TabIndex        =   52
                  Top             =   2400
                  Width           =   6615
                  _ExtentX        =   11668
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
               Begin MSDataListLib.DataCombo DcbManEmp 
                  Bindings        =   "FrmTypeExchange.frx":BD40
                  Height          =   315
                  Left            =   6240
                  TabIndex        =   53
                  Top             =   2760
                  Width           =   6615
                  _ExtentX        =   11668
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
               Begin MSDataListLib.DataCombo DcbManger 
                  Bindings        =   "FrmTypeExchange.frx":BD55
                  Height          =   315
                  Left            =   6240
                  TabIndex        =   54
                  Top             =   3120
                  Width           =   6615
                  _ExtentX        =   11668
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
               Begin MSDataListLib.DataCombo DBCboClientName 
                  Bindings        =   "FrmTypeExchange.frx":BD6A
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   79
                  Top             =   240
                  Width           =   7695
                  _ExtentX        =   13573
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
               Begin MSDataListLib.DataCombo DCAccounts 
                  Bindings        =   "FrmTypeExchange.frx":BD7F
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   82
                  Top             =   240
                  Width           =   7695
                  _ExtentX        =   13573
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
               Begin MSDataListLib.DataCombo dcEmployee 
                  Bindings        =   "FrmTypeExchange.frx":BD94
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   83
                  Top             =   240
                  Width           =   7695
                  _ExtentX        =   13573
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
               Begin MSDataListLib.DataCombo DcbDetpartment 
                  Height          =   315
                  Left            =   6240
                  TabIndex        =   84
                  Top             =   3480
                  Width           =   7575
                  _ExtentX        =   13361
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "7"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbCurrency 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   93
                  Top             =   240
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·«Ì»«‰"
                  Height          =   285
                  Index           =   36
                  Left            =   1500
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   600
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ«·» «·’—›"
                  Height          =   285
                  Index           =   21
                  Left            =   14040
                  TabIndex        =   100
                  Top             =   2400
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H0080FFFF&
                  Caption         =   "—Ì«·"
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   20
                  Left            =   5040
                  TabIndex        =   99
                  Top             =   1680
                  Width           =   8805
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ﬁ«»·"
                  Height          =   285
                  Index           =   19
                  Left            =   5280
                  TabIndex        =   98
                  Top             =   960
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„⁄œ·"
                  Height          =   285
                  Index           =   18
                  Left            =   9840
                  TabIndex        =   96
                  Top             =   960
                  Width           =   765
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„·…"
                  Height          =   255
                  Index           =   14
                  Left            =   1830
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   240
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«œ«—… «·ÿ«·»…"
                  Height          =   375
                  Index           =   37
                  Left            =   14280
                  TabIndex        =   85
                  Top             =   3480
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„” ›Ìœ"
                  Height          =   285
                  Index           =   15
                  Left            =   14040
                  TabIndex        =   81
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„Ì·"
                  Height          =   285
                  Index           =   16
                  Left            =   10470
                  TabIndex        =   80
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì’—› «·Ï «·”«œ…"
                  Height          =   285
                  Index           =   29
                  Left            =   13920
                  TabIndex        =   62
                  Top             =   600
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·»€ Êﬁœ—Â"
                  Height          =   285
                  Index           =   0
                  Left            =   13800
                  TabIndex        =   61
                  Top             =   960
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "—Ì«·"
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   3
                  Left            =   5040
                  TabIndex        =   60
                  Top             =   1320
                  Width           =   8805
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ê–«·ﬂ ⁄»«—Â ⁄‰"
                  Height          =   285
                  Index           =   2
                  Left            =   14040
                  TabIndex        =   59
                  Top             =   2040
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‘—Õ"
                  Height          =   645
                  Index           =   5
                  Left            =   5520
                  TabIndex        =   58
                  Top             =   2400
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ«·» «·’—›"
                  Height          =   285
                  Index           =   9
                  Left            =   14040
                  TabIndex        =   57
                  Top             =   2040
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„œÌ— «·„»«‘—"
                  Height          =   285
                  Index           =   10
                  Left            =   14040
                  TabIndex        =   56
                  Top             =   2760
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ— «·«œ«—…"
                  Height          =   285
                  Index           =   12
                  Left            =   14040
                  TabIndex        =   55
                  Top             =   3120
                  Width           =   1365
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·’—›"
               Height          =   435
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   6000
               Width           =   13695
               Begin XtremeSuiteControls.RadioButton Opt 
                  Height          =   255
                  Index           =   0
                  Left            =   7920
                  TabIndex        =   43
                  Top             =   120
                  Width           =   1815
                  _Version        =   786432
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "‰ﬁœÌ"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton Opt 
                  Height          =   255
                  Index           =   1
                  Left            =   4680
                  TabIndex        =   44
                  Top             =   120
                  Width           =   1815
                  _Version        =   786432
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "‘Ìﬂ"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton Opt 
                  Height          =   255
                  Index           =   2
                  Left            =   960
                  TabIndex        =   45
                  Top             =   120
                  Width           =   1815
                  _Version        =   786432
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "ÕÊ«·Â"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin VB.TextBox TxtSerial1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   13200
               Locked          =   -1  'True
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TxtDes2 
               Alignment       =   1  'Right Justify
               Height          =   765
               Left            =   360
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   40
               Top             =   5280
               Width           =   13815
            End
            Begin MSDataListLib.DataCombo Dcbranch 
               Bindings        =   "FrmTypeExchange.frx":BDA9
               Height          =   315
               Left            =   480
               TabIndex        =   65
               Top             =   240
               Width           =   8175
               _ExtentX        =   14420
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
            Begin MSComCtl2.DTPicker XPDtbTrans 
               Height          =   315
               Left            =   10320
               TabIndex        =   66
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   113377281
               CurrentDate     =   41640
            End
            Begin VB.Label lblbr 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·›—⁄"
               Height          =   255
               Left            =   8760
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   300
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÿ·»"
               Height          =   285
               Index           =   4
               Left            =   14670
               TabIndex        =   70
               Top             =   270
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· «—ÌŒ"
               Height          =   285
               Index           =   1
               Left            =   11790
               TabIndex        =   69
               Top             =   255
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   285
               Index           =   11
               Left            =   -1320
               TabIndex        =   68
               Top             =   240
               Width           =   525
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ«  «·„—«Ã⁄Â"
               Height          =   285
               Index           =   13
               Left            =   14400
               TabIndex        =   67
               Top             =   5400
               Width           =   1365
            End
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   8
      Left            =   0
      TabIndex        =   106
      Top             =   0
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·‘Œ’ «·„”∆Ê·"
      Height          =   285
      Index           =   28
      Left            =   4080
      TabIndex        =   31
      Top             =   2160
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   18600
      RightToLeft     =   -1  'True
      TabIndex        =   30
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
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4170
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   120
      Top             =   4080
      Width           =   6015
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   9525
      TabIndex        =   18
      Top             =   7635
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   5040
      TabIndex        =   17
      Top             =   7620
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   3210
      TabIndex        =   16
      Top             =   7620
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   2490
      TabIndex        =   15
      Top             =   7620
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   4260
      TabIndex        =   14
      Top             =   7620
      Width           =   615
   End
End
Attribute VB_Name = "FrmTypeExchange"
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
'Dim cSearchDcbo  As clsDCboSearch
Dim Dcombos As ClsDataCombos
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

Private Sub Accredit_Click()
    Dim sql As String
    Dim BeginTrans As Boolean
      If val(XPTxtID.text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
      
    'sql = "update  Transactions  set Posted=" & user_id & "  where Transaction_ID=" & Val(XPTxtBillID.text)
    'Cn.Execute sql

'    Cn.BeginTrans
'    BeginTrans = True

   ' If IsNull(rs("Posted")) Then
   '     rs("Posted") = user_id
   '     rs("PostedDate") = Time
   ' Else
   '     rs("Posted") = Null
   '    rs("PostedDate") = Time
   ' End If
   '
   ' rs.update
    SendTopost Me.Name, "TblExchange", "Id", val(DcbDetpartment.BoundText), val(Dcbranch.BoundText), val(XPTxtID.text), TxtSerial1, , val(DcbDetpartment.BoundText)
    
   rs.Resync
    
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

    'Cn.CommitTrans
    'BeginTrans = False
'FillApprovedTable
  Retrive (val(XPTxtID.text))



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
Dim UserID As Integer
Dim UserID1 As Integer
Dim UserID2 As Integer
Dim EmpID As Integer
         currentdate = Now
            
                        GetApprovalDepartement val(DcbDetpartment.BoundText), UserID, EmpID, val(Me.Dcbranch.BoundText), UserID1, UserID2
            Dim currcusor As Integer
            currcusor = 1
            If UserID <> 0 Then
           '***************************************
                                 RSApproval.AddNew
                        RSApproval("ScreenName").value = Me.Name
                        RSApproval("levelo").value = 0
                       RSApproval("EmpID").value = UserID
                        RSApproval("levelorder").value = 0
                         RSApproval("currorder").value = 0
                          RSApproval("Transaction_ID").value = val(XPTxtID.text)
                          RSApproval("NoteSerial").value = TxtSerial1.text
                        RSApproval("Transaction_Date").value = Date
                        
                          RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
                       RSApproval("SendTime").value = currentdate
        
                 
                                RSApproval("Currcursor").value = currcusor
                                 RSApproval("FromUser").value = user_name
                     
                        
                        RSApproval.update
              End If
              
              
              
            If UserID1 <> 0 Then
           '***************************************
           currcusor = currcusor + 1
                                 RSApproval.AddNew
                        RSApproval("ScreenName").value = Me.Name
                        RSApproval("levelo").value = 0
                       RSApproval("EmpID").value = UserID1
                        RSApproval("levelorder").value = 0
                         RSApproval("currorder").value = 0
                          RSApproval("Transaction_ID").value = val(XPTxtID.text)
                          RSApproval("NoteSerial").value = TxtSerial1.text
                        RSApproval("Transaction_Date").value = Date
                        
                          RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
                       RSApproval("SendTime").value = currentdate
        
                 
                                RSApproval("Currcursor").value = currcusor
                                 RSApproval("FromUser").value = user_name
                     
                        
                        RSApproval.update
              End If
              
                 
           If UserID2 <> 0 Then
           '***************************************
           currcusor = currcusor + 1
                                 RSApproval.AddNew
                        RSApproval("ScreenName").value = Me.Name
                        RSApproval("levelo").value = 0
                       RSApproval("EmpID").value = UserID2
                        RSApproval("levelorder").value = 0
                         RSApproval("currorder").value = 0
                          RSApproval("Transaction_ID").value = val(XPTxtID.text)
                          RSApproval("NoteSerial").value = TxtSerial1.text
                        RSApproval("Transaction_Date").value = Date
                        
                          RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
                       RSApproval("SendTime").value = currentdate
        
                 
                                RSApproval("Currcursor").value = currcusor
                                 RSApproval("FromUser").value = user_name
                     
                        
                        RSApproval.update
              End If
   
   
    If Rs1.RecordCount > 0 Then
    
              
                 
                 
            For i = 1 To Rs1.RecordCount
            

           '****************************************
            
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.Name
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(XPTxtID.text)
                  RSApproval("NoteSerial").value = TxtSerial1.text
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 And UserID = 0 Then
                        RSApproval("Currcursor").value = 1
                       RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                Rs1.MoveNext
            Next i

    End If
    
    

End Function



Private Sub CBoBasedON_Change()
If val(Me.CBoBasedON.ListIndex) = 2 Then
Command2.Visible = True
Frame5.Visible = True
Else
Command2.Visible = False
Frame5.Visible = False
End If
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg = "E" Then
  TxtOrderNo.text = ""
   txtTransaction_ID.text = ""
   
End If
End Sub

Private Sub CBoBasedON_Click()
CBoBasedON_Change
End Sub

'    Cn.CommitTrans
'    BeginTrans = False
'FillApprovedTable
'    Retrive (val(Me.XPTxtID.text))
'End Sub

'Private Sub bClose_Click()
'Frame6.Visible = False
'If Me.ChekAccept.value = xtpChecked Then
'Frame2.Visible = True
'End If
'If Me.ChekContracted.value = xtpChecked Then
'Frame5.Visible = True
'End If
'End Sub

'Private Sub ChekAccept_Click()
'If Me.ChekAccept.value = vbChecked Then
'Me.CHekNotAccept.value = vbUnchecked
'Me.ChekContracted.value = vbUnchecked
'lbl(36).Visible = False
'Me.txtnotAccept.Visible = False
'Me.Frame2.Visible = True
'Me.Frame5.Visible = False
'Else
'Me.Frame2.Visible = False
'End If
'End Sub
'Private Sub RemoveGridRow()
'
'    With Me.Fg
'
'        If .Row <= 0 Then Exit Sub
'        .RemoveItem .Row
'    End With
'
'    ReLineGrid
'End Sub
'Private Sub RemoveGridRow2()
'
'    With Me.fg2
'
'        If .Row <= 0 Then Exit Sub
'        .RemoveItem .Row
'    End With
'
'    ReLineGrid
'End Sub

'Private Sub ChekContracted_Click()
'If Me.ChekContracted.value = xtpChecked Then
'Me.CHekNotAccept.value = xtpUnchecked
'Me.ChekAccept.value = xtpUnchecked
'lbl(36).Visible = False
'Me.txtnotAccept.Visible = False
'Me.Frame2.Visible = False
'Frame5.Visible = True
'Else
'Me.Frame5.Visible = False
'End If
'
'End Sub

'Private Sub CHekNotAccept_Click()
'If Me.CHekNotAccept.value = vbChecked Then
'Me.Frame2.Visible = False
'Me.Frame5.Visible = False
'lbl(36).Visible = True
'Me.txtnotAccept.Visible = True
'Me.ChekAccept.value = vbUnchecked
''Me.ChekContracted.value = vbUnchecked
'Else
'Me.Frame2.Visible = True
'lbl(36).Visible = False
'Me.txtnotAccept.Visible = False
'End If
'End Sub

Private Sub Cmd_Click(index As Integer)

    ' On Error GoTo ErrTrap
    Select Case index
'Case 8
'MsgBox Format(TimeFrom1.value, "hh:mm AM/PM")
'Case 8
'RemoveGridRow2
'Case 21
'RemoveGridRow
        Case 0
          

Unload FrmReqExchangeSearch
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            ' Me.DcbTO1.BoundText = 0
            '   Me.DcbTO2.BoundText = 0
            clear_all Me
            Opt(0).value = True
 
 DcbCurrency.BoundText = MainCurrency()
CBoBasedON.ListIndex = 0

            Me.DCboUserName.BoundText = user_id
        '    TxtPaymentCounts.text = 1
Dcbranch.BoundText = Current_branch
   
                 GRID2.Clear flexClearScrollable, flexClearEverything
    GRID2.rows = 1
 
             Accredit.Enabled = True
                If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
DcbExchang.BoundText = 1
Option4.value = True

          ''
        Case 1
       
        Unload FrmReqExchangeSearch
      
'Fg.Rows = Fg.Rows + 1
'Fg.Enabled = True
'fg2.Rows = fg2.Rows + 1
'fg2.Enabled = True
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            
            
                          If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "«·› —Â „€·ﬁ… "
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If


 
            
            
         If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not edit.This process associated with approvals"
         End If
         Exit Sub
       End If
       
       
            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
   DcbCurrency_Change
   

       
       
        Case 2
    
            Dim Msg As String

If SystemOptions.MonyeIssueVchrNoMust = True Then

            If val(TxtPrice.text) <= 0 Or val(TxtPriceE.text) <= 0 Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Value"
                Else
                    Msg = "Õœœ ﬁÌ„Â ’ÕÌÕÂ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  
                Exit Sub
            End If
            
End If

            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Dcbranch.SetFocus
                Sendkeys "{F4}"
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

        If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «·Õ–›.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not delete.This process associated with approvals"
         End If
         Exit Sub
       End If
       
       
            Del_Trans

        Case 5
        
            Load FrmReqExchangeSearch
            FrmReqExchangeSearch.show

        Case 6
            Unload Me

        Case 7
           ' ShowGL_cc Me.txtNoteSerial.text, , 200

        Case 8
            
            
            
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
Function print_report(Optional NoteSerial As String, Optional mType As Long = 0)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT    TblExchange.BankIBAN  ,   dbo.TblExchange.Des2, dbo.TblExchange.NoteSerial1, dbo.TblExchange.Id, dbo.TblExchange.RecordDate, dbo.TblExchange.BranchID, "
MySQL = MySQL & "                       dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblExchange.TypeExch, dbo.TblDataTypeExchange.name,"
MySQL = MySQL & "                       dbo.TblDataTypeExchange.namee, dbo.TblExchange.Price, dbo.TblExchange.ToPerson, dbo.TblExchange.EmpID, TblEmployee_4.Emp_Code,"
MySQL = MySQL & "                       TblEmployee_4.Emp_Name, TblEmployee_4.Emp_Name1, TblEmployee_4.Fullcode, TblEmployee_4.Emp_Namee4, TblEmployee_4.Emp_Namee3,"
MySQL = MySQL & "                       TblEmployee_4.Emp_Namee2, TblEmployee_4.Emp_Namee1, TblEmployee_4.Emp_Namee, TblEmployee_4.Emp_Name2, TblEmployee_4.Emp_Name3,"
MySQL = MySQL & "                       TblEmployee_4.Emp_Name4, dbo.TblExchange.ManagerID, TblEmployee_1.Emp_Code AS ManEmp_Code, TblEmployee_1.Emp_Name AS MangEmp_Name,"
MySQL = MySQL & "                       TblEmployee_1.Emp_Name1 AS mangEmp_Name1, TblEmployee_1.Emp_Name2 AS MangEmp_Name2, TblEmployee_1.Emp_Name3 AS MangEmp_Name3,"
MySQL = MySQL & "                       TblEmployee_1.Emp_Name4 AS MangEmp_Name4, TblEmployee_1.Fullcode AS MangFullcode, TblEmployee_1.Emp_Namee4 AS MangEmp_Namee4,"
MySQL = MySQL & "                       TblEmployee_1.Emp_Namee3 AS MangEmp_Namee3, TblEmployee_1.Emp_Namee2 AS MangEmp_Namee2, TblEmployee_1.Emp_Namee1 AS MangEmp_Namee1,"
MySQL = MySQL & "                       TblEmployee_1.Emp_Namee AS MangEmp_Namee, dbo.TblExchange.MempID, TblEmployee_2.Emp_Code AS ManEmpEmp_Code,"
MySQL = MySQL & "                       TblEmployee_2.Emp_Name AS MangEmpEmp_Name, TblEmployee_2.Emp_Name1 AS angEmpEmp_Name1, TblEmployee_2.Emp_Name2 AS MangEmpEmp_Name2,"
MySQL = MySQL & "                       TblEmployee_2.Emp_Name3 AS MangEmpEmp_Name3, TblEmployee_2.Emp_Name4 AS MangEmpEmp_Name4, TblEmployee_2.Fullcode AS angEmpFullcode,"
MySQL = MySQL & "                       TblEmployee_2.Emp_Namee4 AS MangEmpEmp_Namee4, TblEmployee_2.Emp_Namee3 AS angeEmpEmp_Namee3,"
MySQL = MySQL & "                       TblEmployee_2.Emp_Namee2 AS MangEmpEmp_Namee2, TblEmployee_2.Emp_Namee1 AS angEmpEmp_Namee1,"
MySQL = MySQL & "                       TblEmployee_2.Emp_Namee AS MangEMpEmp_Namee, dbo.TblExchange.Des, dbo.TblExchange.UserID, dbo.TblExchange.Type, dbo.TblExchange.Posted,"
MySQL = MySQL & "                       dbo.TblExchange.PostedDate, dbo.TblExchange.Approved, dbo.TblExchange.OrderNo, dbo.TblExchange.Transaction_ID, dbo.TblExchange.basedOn,"
MySQL = MySQL & "                       dbo.TblExchange.FromType, dbo.TblExchange.Account_Code, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng,"
MySQL = MySQL & "                       dbo.TblExchange.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode,"
MySQL = MySQL & "                       TblEmployee_3.Emp_Name AS FrmEmp_Name, TblEmployee_3.Emp_Name1 AS FrmEmp_Name1, TblEmployee_3.Emp_Name2 AS FrmEmp_Name2,"
MySQL = MySQL & "                       TblEmployee_3.Emp_Name3 AS FrmEmp_Name3, TblEmployee_3.Emp_Name4 AS FrmEmp_Name4, TblEmployee_3.Fullcode AS FrmFullcode,"
MySQL = MySQL & "                       TblEmployee_3.Emp_Namee4 AS FrmEmp_NameE4, TblEmployee_3.Emp_Namee3 AS FrmEmp_NameE3, TblEmployee_3.Emp_Namee2 AS FrmEmp_NameE2,"
MySQL = MySQL & "                       TblEmployee_3.Emp_Namee1 AS FrmEmp_NameE1, TblEmployee_3.Emp_Namee AS FrmEmp_NameE, dbo.TblExchange.EmpID1, dbo.projects.Project_name,"
MySQL = MySQL & "                       dbo.projects.Project_nameE, dbo.TblExchange.ReqDate, dbo.TblExchange.DeptID, dbo.TblEmpDepartments.DepartmentName,"
MySQL = MySQL & "                       dbo.TblEmpDepartments.DepartmentNamee, dbo.TblExchange.Rate, dbo.TblExchange.PriceE, dbo.TblExchange.CurrcyID, dbo.currency.code,"
MySQL = MySQL & "                       dbo.currency.name AS Currname, dbo.currency.nameE AS CurrnameE"
MySQL = MySQL & "  FROM         dbo.TblEmployee TblEmployee_4 RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.projects RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblExchange LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.currency ON dbo.TblExchange.CurrcyID = dbo.currency.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmpDepartments ON dbo.TblExchange.DeptID = dbo.TblEmpDepartments.DeparmentID ON dbo.projects.id = dbo.TblExchange.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee TblEmployee_3 ON dbo.TblExchange.EmpID1 = TblEmployee_3.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee TblEmployee_2 ON dbo.TblExchange.MempID = TblEmployee_2.Emp_ID ON TblEmployee_4.Emp_ID = dbo.TblExchange.EmpID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee TblEmployee_1 ON dbo.TblExchange.ManagerID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCustemers ON dbo.TblExchange.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.ACCOUNTS ON dbo.TblExchange.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblDataTypeExchange ON dbo.TblExchange.TypeExch = dbo.TblDataTypeExchange.Id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.TblExchange.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblExchange.id = " & val(Me.XPTxtID.text) & ")"





  If SystemOptions.UserInterface = ArabicInterface Then
            'StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepExchange.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepExchange.rpt"
        Else
         '   StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepExchangeE.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RepExchangeE.rpt"
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
        Msg = "Not Found Data to Show"
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
    
    
        xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(TxtPrice.text, "###.00"), 0)
       'xReport.ParameterFields(5).AddCurrentValue WriteNo(Format(TxtPriceE.Text, "###.00"), 0, , , , , , val(DcbCurrency.BoundText))

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub CmdAttach_Click()
    On Error Resume Next
ShowAttachments TxtSerial1, "0612201401"
 
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

'Private Sub Command2_Click()
'FillGridDetails
'Frame6.Visible = True
'Frame5.Visible = False
'Frame2.Visible = False
'End Sub
'Sub FillGridDetails()
'Dim StrSQL As String
'Dim i As Integer
'Dim RsDetails As ADODB.Recordset
'Set RsDetails = New ADODB.Recordset
'StrSQL = " SELECT     dbo.TblRegDateDelgate.Id, TblEmployee_1.Emp_ID, dbo.TblRegDateDelgate.RecordDate, dbo.TblRegDateDelgate.DelgID, TblEmployee_1.Emp_Code AS Emp_CodeD, "
'StrSQL = StrSQL & "                      TblEmployee_1.Emp_Name AS Emp_NameD, TblEmployee_1.Nationality AS NationalityD, TblEmployee_1.Fullcode AS FullcodeD,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.CustomerName, dbo.TblRegDateDelgate.Remark, dbo.TblRegDateDelgate.VisitID, TblTypeVisit_1.name, TblTypeVisit_1.namee,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.VisitID2, dbo.TblRegDateDelgate.SpAsID, dbo.TblSpeciaAsement.name AS nameSp, dbo.TblSpeciaAsement.namee AS nameeSp,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.Accept, dbo.TblRegDateDelgate.VisitDate, dbo.TblRegDateDelgate.Remark2, dbo.TblRegDateDelgate.PersonConc,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.Tel, dbo.TblRegDateDelgate.Mobile, dbo.TblRegDateDelgate.Email, dbo.TblRegDateDelgate.JobID, dbo.TblRegDateDelgate.LongTime,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.VisitDate1, TblTypeVisit_2.name AS name2, TblTypeVisit_2.namee AS namee2, dbo.TblRegDateDelgate.Entry, dbo.TblRegDateDelgate.Map,"
''StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.Adress, dbo.TblRegDateDelgate.NotAcept, dbo.TblRegDateDelgate.BillNo, TblEmployee_1.Emp_Namee, dbo.TblRegDateDelgate.CustomerID,"
'StrSQL = StrSQL & "                         dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblRegDateDelgate.ToTime1, dbo.TblRegTimeDelgate.name AS ToTime11,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.FromTime1, TblRegTimeDelgate_2.name AS FromTime11, dbo.TblRegDateDelgate.FromTime2, TblRegTimeDelgate_3.name AS FromTime22,"
'StrSQL = StrSQL & "                         dbo.TblRegDateDelgate.ToTime2, TblRegTimeDelgate_1.name AS ToTime22"
'StrSQL = StrSQL & "    FROM         dbo.TblRegTimeDelgate RIGHT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblRegTimeDelgate TblRegTimeDelgate_1 RIGHT OUTER JOIN"
' StrSQL = StrSQL & "                        dbo.TblRegDateDelgate ON TblRegTimeDelgate_1.Id = dbo.TblRegDateDelgate.ToTime2 LEFT OUTER JOIN"
'' StrSQL = StrSQL & "                        dbo.TblRegTimeDelgate TblRegTimeDelgate_3 ON dbo.TblRegDateDelgate.FromTime2 = TblRegTimeDelgate_3.Id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblRegTimeDelgate TblRegTimeDelgate_2 ON dbo.TblRegDateDelgate.FromTime1 = TblRegTimeDelgate_2.Id ON"
'StrSQL = StrSQL & "                         dbo.TblRegTimeDelgate.Id = dbo.TblRegDateDelgate.ToTime1 LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblCustemers ON dbo.TblRegDateDelgate.CustomerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblTypeVisit TblTypeVisit_2 ON dbo.TblRegDateDelgate.VisitID2 = TblTypeVisit_2.Id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblTypeVisit TblTypeVisit_1 ON dbo.TblRegDateDelgate.VisitID = TblTypeVisit_1.Id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblSpeciaAsement ON dbo.TblRegDateDelgate.SpAsID = dbo.TblSpeciaAsement.Id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                         dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDateDelgate.DelgID = TblEmployee_1.Emp_ID"
'StrSQL = StrSQL & "    Where (dbo.TblRegDateDelgate.customerid =" & val(Me.DcbCustomer.BoundText) & ")"
'RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
'    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
'    VSFlexGrid1.Rows = VSFlexGrid1.FixedRows
'With VSFlexGrid1
'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        .Rows = .FixedRows + RsDetails.RecordCount
'
'        For i = .FixedRows To .Rows - 1
'        .TextMatrix(i, .ColIndex("Serial")) = i
'        .TextMatrix(i, .ColIndex("PersonConc")) = IIf(IsNull(RsDetails("PersonConc").value), "", RsDetails("PersonConc").value) ' RsDetails("remark").value
'           ' .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CustomerName").value), "", RsDetails("CustomerName").value) 'RsDetails("fullcode").value
'            If SystemOptions.UserInterface = EnglishInterface Then
'           .TextMatrix(i, .ColIndex("Emp_NameD")) = IIf(IsNull(RsDetails("Emp_Namee").value), "", RsDetails("Emp_Namee").value) 'RsDetails("Emp_Namee").value
'            .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CusNamee").value), "", RsDetails("CusNamee").value)
'           Else
'           .TextMatrix(i, .ColIndex("Emp_NameD")) = IIf(IsNull(RsDetails("Emp_NameD").value), "", RsDetails("Emp_NameD").value) ' RsDetails("emp_name").value
'            .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(RsDetails("CusName").value), "", RsDetails("CusName").value)
'           End If
'            .TextMatrix(i, .ColIndex("Mobile")) = IIf(IsNull(RsDetails("Mobile").value), "", RsDetails("Mobile").value)
'             .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(RsDetails("JobID").value), "", RsDetails("JobID").value)
'              .TextMatrix(i, .ColIndex("Tel")) = IIf(IsNull(RsDetails("Tel").value), "", RsDetails("Tel").value)
'               .TextMatrix(i, .ColIndex("Email")) = IIf(IsNull(RsDetails("Email").value), "", RsDetails("Email").value)
'                .TextMatrix(i, .ColIndex("FromTim11")) = IIf(IsNull(RsDetails("FromTime11").value), "", RsDetails("FromTime11").value)
 '                .TextMatrix(i, .ColIndex("ToTime11")) = IIf(IsNull(RsDetails("ToTime11").value), "", RsDetails("ToTime11").value)
'                  .TextMatrix(i, .ColIndex("Adress")) = IIf(IsNull(RsDetails("Adress").value), "", RsDetails("Adress").value)
'                  .TextMatrix(i, .ColIndex("VisitDate1")) = IIf(IsNull(RsDetails("VisitDate1").value), "", RsDetails("VisitDate1").value)
''                  DcbTypeVisit1.BoundText = val(IIf(IsNull(RsDetails("VisitID").value), "", RsDetails("VisitID").value))
 '                 .TextMatrix(i, .ColIndex("VisitID")) = DcbTypeVisit1.text
 '               If RsDetails("Accept").value = 0 Then
 '               .TextMatrix(i, .ColIndex("Accept")) = ""
 '               End If
 '                If RsDetails("Accept").value = 1 Then
 '               .TextMatrix(i, .ColIndex("Accept")) = " „ «·“Ì«—…"
 '               End If
 '                If RsDetails("Accept").value = 2 Then
 '               .TextMatrix(i, .ColIndex("Accept")) = " „ «· ⁄«ﬁœ"
 '               End If
 ''                If RsDetails("Accept").value = 3 Then
  '              .TextMatrix(i, .ColIndex("Accept")) = "≈·€«¡ «·“Ì«—…"
  '              End If
  '               .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(RsDetails("Remark").value), "", RsDetails("Remark").value)
                
                
  '          RsDetails.MoveNext
  '      Next i

  '  End If
'End With
'    RsDetails.Close
'    Set RsDetails = Nothing
'End Sub
'Private Sub DateVisit1_KeyUp(KeyCode As Integer, Shift As Integer)
'If TxtModFlg.text <> "R" Then
'If val(Me.DcboEmpName.BoundText) = 0 Then
'MsgBox "ÌÃ»  ÕœÌœ «”„ «·„‰œÊ» «Ê·«"
'Exit Sub
'Else
'fileFgtim val(Me.DcboEmpName.BoundText), 0
'refiltimdetails val(Me.DcboEmpName.BoundText), 0
'End If
'End If
'End Sub



'Private Sub DcbCustomer_Change()
'If Me.TxtModFlg.text <> "R" Then
''Me.TxtCustomer.text = ""
'retInfoCustomer

'End If
'End Sub

'Private Sub DcbFrom1_Change()
'If TxtModFlg.text <> "R" Then
'If val(Me.DcboEmpName.BoundText) = 0 Then
'MsgBox "«Œ Ì«— «·„ÊŸ› «Ê·«"
'Exit Sub
'End If
'If Me.DcbFrom1.text <> "" And Me.DcbTO1.text <> "" Then
'If val(Me.DcbFrom1.text) >= val(Me.DcbTO1.text) Then
'MsgBox "ÌÃ» «‰ ÌﬂÊ‰ «·Êﬁ  «·«ŒÌ— «ﬂ»— „‰ Êﬁ  «·»œ«ÌÂ"
''DcbTO2.SetFocus
'Exit Sub

'Else
'chektime val(Me.DcboEmpName.BoundText), val(Me.DcbFrom1.BoundText), val(Me.DcbTO1.BoundText), 2
'fileFgtim val(Me.DcboEmpName.BoundText), 0
'refiltimdetails val(Me.DcboEmpName.BoundText), 0
'End If
'End If
'End If
'End Sub

'Private Sub DcboEmpName_Change()
'If TxtModFlg.text <> "R" Then
'fileFgtim val(Me.DcboEmpName.BoundText), 0
'refiltimdetails val(Me.DcboEmpName.BoundText), 0
'End If
'End Sub

'Private Sub DcboEmpName_Change()
'DcboEmpName_Click (0)

'End Sub






 




 

'Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
'                             Shift As Integer)
'
'    If KeyCode = vbKeyF3 Then
'        FrmEmployeeSearch.lbltype = 8
'       ' Set FrmEmployeeSearch.RetrunFrm = Me
'
'        FrmEmployeeSearch.Show
'
'    End If

Private Sub Combo1_Change()
MsgBox "1111"
End Sub

Private Sub Command1_Click()
If TxtOrderNo.text = "" Then Exit Sub
If CBoBasedON.ListIndex = 1 Then
 FrmPO8.show
  FrmPO8.Retrive val(Me.txtTransaction_ID.text)
ElseIf CBoBasedON.ListIndex = 2 Then
 FrmPO10.show
  FrmPO10.Retrive val(Me.txtTransaction_ID.text)
 ElseIf CBoBasedON.ListIndex = 3 Then
FrmVocationEntitlements.show
If val(Me.txtTransaction_ID.text) <> 0 Then
 FrmVocationEntitlements.Retrive val(Me.txtTransaction_ID.text)
 Else
 FrmVocationEntitlements.Retrive val(Me.TxtOrderNo.text)
 End If
  ElseIf CBoBasedON.ListIndex = 4 Then
 End_oF_service.show
 If val(Me.txtTransaction_ID.text) <> 0 Then
 End_oF_service.Retrive val(Me.txtTransaction_ID.text)
 Else
 End_oF_service.Retrive val(Me.TxtOrderNo.text)
 End If
 
ElseIf CBoBasedON.ListIndex = 5 Then
 FrmBankPledge4.show
 FrmBankPledge4.Retrive val(Me.TxtOrderNo.text)
  End If
End Sub

Private Sub Command2_Click()
Frame5.Visible = True
End Sub

Private Sub Command7_Click()
Frame6.Visible = False
End Sub

Private Sub DBCboClientName_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
txtperson.text = DBCboClientName.text

End If
GetIBan
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeyF3 Then
        
                   If DCboCashType121.ListIndex = 0 Then
                        
                        FrmCustemerSearch.SearchType = 9915
                        FrmCustemerSearch.show vbModal
                ElseIf DCboCashType121.ListIndex = 1 Then
                                  FrmCompanySearch.lblSearchtype.Caption = 9915
                      FrmCompanySearch.show vbModal
              
                
                       
                 End If


    End If
    
    
End Sub

Private Sub DCAccounts_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
txtperson.text = DCAccounts.text
End If
End Sub

Private Sub DCAccounts_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then
            Unload Account_search
        Account_search.show
        Account_search.case_id = 9915
        
   End If
   
End Sub

Private Sub DcbCurrency_Change()
DcbCurrency_Click (0)
End Sub
Sub CalCuteCurrency()
'Exit Sub
If val(TxtCurrencyRate.text) = 0 Then
TxtCurrencyRate.text = ""
End If
If val(val(TxtCurrencyRate.text)) <> 0 Then

'Text2.Text = Format(Text2.Text, "###.00")
TxtPriceE.text = Round(val(Format(TxtPrice.text, "###.00")) / val(TxtCurrencyRate.text), 2)
TxtPriceE.text = Format(TxtPriceE.text, "#,##0.00")
Else
TxtPriceE.text = (TxtPrice.text)
End If
If val(TxtPrice.text) = 0 Then TxtPriceE.text = "": Exit Sub
' TxtPrice.Text = Format(TxtPrice.Text, "#,##0.00")
 
End Sub
Private Sub DcbCurrency_Click(Area As Integer)
    If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    If Me.DcbCurrency.BoundText <> "" Then
        TxtCurrencyRate.text = get_currency_rate(Me.DcbCurrency.BoundText)
    Else
        TxtCurrencyRate.text = ""
    End If
CalCuteCurrency
End Sub

Private Sub DcbEmp_Change()
DcbEmp_Click (0)
End Sub

Private Sub DcbEmp_Click(Area As Integer)

     On Error Resume Next
       If val(Me.DcbEmp.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcbEmp.BoundText, EmpCode
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
        Dim MandId As Integer
        Dim endContractPerMonth As Double
        get_employee_information val(Me.DcbEmp.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth, , MandId
        
Me.DcbManEmp.BoundText = MandId
DcbDetpartment.BoundText = DepID
End Sub

Private Sub DcbManEmp_Change()
DcbManEmp_Click (0)
End Sub

Private Sub DcbManEmp_Click(Area As Integer)
    If val(Me.DcbManEmp.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcbManEmp.BoundText, EmpCode
    TxtME.text = EmpCode
    
End Sub

Private Sub DcbManger_Change()
DcbManger_Click (0)
End Sub

Private Sub DcbManger_Click(Area As Integer)
    If val(Me.DcbManger.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbManger.BoundText, EmpCode
    TxtManger.text = EmpCode
End Sub
Private Sub DCboCashType121_Change()
Set Dcombos = New ClsDataCombos
     Frame7.Visible = False
Select Case DCboCashType121.ListIndex

        Case 0
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
            Me.DBCboClientName.Visible = True
            DCEmployee.Visible = False
            DCAccounts.Visible = False
               If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(16).Caption = "«”„ «·⁄„Ì·"
            Else
                Me.lbl(16).Caption = "Customer Name"
            End If
            Case 1
        
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
            Me.DBCboClientName.Visible = True
            DCEmployee.Visible = False
            DCAccounts.Visible = False

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(16).Caption = "«”„ «·„Ê—œ"
            Else
                Me.lbl(16).Caption = "Vendor Name"
            End If
            Case 2
    
            Dcombos.GetPersons Me.DBCboClientName
            Me.DBCboClientName.Visible = True
         
            DCEmployee.Visible = False
            DCAccounts.Visible = False
             If SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(16).Caption = "name"
            Else
                Me.lbl(16).Caption = "„ﬁ«Ê· «·»«ÿ‰"
            End If
             Case 3
      
             Dim My_SQL As String
             If SystemOptions.UserInterface = ArabicInterface Then
                    My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null) order by Project_name" '
            Else
                    My_SQL = "  select id,Project_nameE from projects where not(REVENUE_account is null) order by Project_name" '
            End If
            fill_combo Me.DBCboClientName, My_SQL
         
            Me.DBCboClientName.Visible = True
          
            DCEmployee.Visible = False
            DCAccounts.Visible = False

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(16).Caption = "«”„ «·„‘—Ê⁄"
            Else
                Me.lbl(16).Caption = "project Name"
            End If
          Case 4
     

If DCboCashType121.ListIndex = 4 Then
Frame7.Visible = True
End If

            Dcombos.GetEmployees Me.DCEmployee
            Me.DCEmployee.Visible = True
            
            DBCboClientName.Visible = False
            DCAccounts.Visible = False
            
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(16).Caption = "«”„ «·„ÊŸ›"
            Else
                Me.lbl(16).Caption = "Employee  Name"
            End If
             Case 5

            Dcombos.GetAccountingCodes Me.DCAccounts, True
            'Dcombos.GetAccountingCodes Me.DcAccounts1, True
            DCAccounts.Visible = True
            Me.DCEmployee.Visible = False
            DBCboClientName.Visible = False
        
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(16).Caption = "«”„ «·Õ”«»"
            Else
                Me.lbl(16).Caption = "Accounts Nam  "
            End If
         End Select
        
End Sub
Private Sub GetIBan()
Dim s As String
s = "Select BankIBAN  FROM TblCustemers AS tc WHERE tc.CusID = " & val(DBCboClientName.BoundText)
Dim rsDummy As New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsDummy.EOF Then
    TxtBankIBAN = rsDummy!BankIBAN & ""
End If
End Sub
Private Sub DCboCashType121_Click()
DCboCashType121_Change
End Sub

Private Sub DcEmployee_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
txtperson.text = DCEmployee.text
End If
End Sub

Private Sub DCEmployee_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
         FrmEmployeeSearch.lbltype = 9915
       Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
        
End If
End Sub
Sub RetriveTender(Optional ID As Double = 0)
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " select * from TblBankPledge4 where id =" & ID & ""
Rs3.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs3.RecordCount > 0 Then
TxtPrice.text = IIf(IsNull(Rs3("CopyPrice").value), 0, Rs3("CopyPrice").value)
ReqDate.value = IIf(IsNull(Rs3("RecordDate").value), Date, Rs3("RecordDate").value)
If Not IsNull(Rs3("PaymentType").value) Then
If (Rs3("PaymentType").value) = 1 Then
Opt(1).value = True
ElseIf (Rs3("PaymentType").value) = 2 Then
Opt(2).value = True
Else
Opt(0).value = True
End If
Else
Opt(0).value = True
End If
Else
TxtPrice.text = 0
ReqDate.value = 0
End If
End Sub
Sub RetriveEndSevice(Optional ID As Double = 0)
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " select * from End_oF_service where id =" & ID & ""
Rs3.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs3.RecordCount > 0 Then
DCboCashType121.ListIndex = 4
DCEmployee.BoundText = IIf(IsNull(Rs3("EmpID").value), "", Rs3("EmpID").value)
TxtPrice.text = IIf(IsNull(Rs3("LastTotal").value), 0, Rs3("LastTotal").value)
TxtDes.text = IIf(IsNull(Rs3("Reaons").value), "", Rs3("Reaons").value)
ReqDate.value = IIf(IsNull(Rs3("record_date").value), Date, Rs3("record_date").value)
DcbEmp.BoundText = IIf(IsNull(Rs3("EmpID").value), "", Rs3("EmpID").value)
Else

DCEmployee.BoundText = 0
TxtPrice.text = 0
TxtDes.text = 0
ReqDate.value = 0
DcbEmp.BoundText = 0
End If
End Sub
Function CheckVacationPayed(Optional ID As Double) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = "Select Id from tblVocationEntitlements where ID=" & ID & " and PayedPayment =1"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckVacationPayed = True
Else
CheckVacationPayed = False
End If
End Function
Function CheckTenderPayed(Optional ID As Double) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = "Select Id from TblBankPledge4 where ID=" & ID & " and PayedPayment =1"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckTenderPayed = True
Else
CheckTenderPayed = False
End If
End Function
Sub RetriveVacation(Optional ID As Double = 0)
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     EmpID, Remark, Salary, Increase, SalaryVocation, SalEntitOther, ValueTickt,"
sql = sql & "                      IsNull(Salary ,0)+ IsNull(Increase ,0)+ IsNull(SalaryVocation ,0)+ IsNull(SalEntitOther,0) + IsNull(ValueTickt,0) - IsNull(Decrease,0)- IsNull(Advance,0) - IsNull(Other,0) AS summ, Advance, Other, RecordDate, ID"
sql = sql & "  From dbo.TblVocationEntitlements"
sql = sql & " Where ID = " & ID & ""
Rs3.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs3.RecordCount > 0 Then
DCboCashType121.ListIndex = 3
DCEmployee.BoundText = IIf(IsNull(Rs3("EmpID").value), "", Rs3("EmpID").value)
TxtPrice.text = IIf(IsNull(Rs3("summ").value), 0, Rs3("summ").value)
TxtDes.text = IIf(IsNull(Rs3("Remark").value), "", Rs3("Remark").value)
ReqDate.value = IIf(IsNull(Rs3("RecordDate").value), Date, Rs3("RecordDate").value)
DcbEmp.BoundText = IIf(IsNull(Rs3("EmpID").value), "", Rs3("EmpID").value)
Else

DCEmployee.BoundText = 0
TxtPrice.text = 0
TxtDes.text = 0
ReqDate.value = 0
DcbEmp.BoundText = 0
End If
End Sub
'End Sub



Function Calculate_TotalSelected() As Double
    Dim i As Integer
    On Error Resume Next
 
    If Fg.rows = 1 Then Exit Function
    Calculate_TotalSelected = 0

    For i = 1 To Fg.rows - 1
        
        If Fg.cell(flexcpChecked, i, Fg.ColIndex("selec")) = flexChecked Then
            
            Calculate_TotalSelected = Calculate_TotalSelected + val(Fg.TextMatrix(i, Fg.ColIndex("Total")))
'            branchs_nos = val(Grid1.TextMatrix(i, Grid1.ColIndex("EmpTotalNet"))) + "," + branchs_nos
         End If

    Next i
 
 
    
    
  ' FillGridWithData3
   
End Function
Private Sub FG_AfterEdit(ByVal row As Long, ByVal Col As Long)
With Fg
Select Case .ColKey(Col)
Case "ExpQty"
If val(.TextMatrix(row, .ColIndex("ExpQty"))) > val(.TextMatrix(row, .ColIndex("QtyAlaw"))) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ ·«‰ «·ﬂ„ÌÂ «·„’—Ê›Â «ﬂ»— „‰ «·ﬂ„Ì… «·„ Ê›—Â"
Else
MsgBox "Can not Qty Expended Larger Larger than the Qty available "
End If
.TextMatrix(row, .ColIndex("ExpQty")) = 0
Exit Sub
End If
.TextMatrix(row, .ColIndex("Total")) = val(.TextMatrix(row, .ColIndex("showPrice"))) * val(.TextMatrix(row, .ColIndex("ExpQty")))
End Select

TxtPrice.text = Calculate_TotalSelected
End With



End Sub

Private Sub FG_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg
Select Case .ColKey(Col)
Case "Name"
Cancel = True
Case "ShowQty"
Cancel = True
Case "QtyAlaw"
Cancel = True
Case "showPrice"
Cancel = True
Case "Total"
Cancel = True

End Select
End With
End Sub

Private Sub GRID2_CellButtonClick(ByVal row As Long, ByVal Col As Long)
With GRID2
Select Case .ColKey(Col)
Case "Remarks"
Frame6.Visible = False
lbl(22).Caption = ""
lbl(22).Caption = .TextMatrix(row, .ColIndex("Remarks"))
Frame6.Visible = True
End Select
End With
End Sub

'Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
'       If val(DcboEmpName.BoundText) = 0 Then Exit Sub


'    Dim EmpCode  As String
 
'    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
  '  TxtSearchCode.text = EmpCode
    
        'txtFile.text = EmpCode
        
'   If Me.TxtModFlg = "R" Then Exit Sub
'
'
'    Dim StrSQL As String
'
'
'        GetEmployeeSalaryAccordingToComponentAll val(Me.DcboEmpName.BoundText)
'
'        Dim IssueDate As Date
'        Dim depid As Double
'        Dim specid As Double
'        Dim JobTypeID As Double
'        Dim gradeID As Double
'        Dim Account_code2 As String
'           Dim Account_Code  As String
'        Dim Balance As String
'        Dim projectid As Integer
' Dim endiqama As String
'        Dim national As String
'        Dim endContractPerMonth As Double
'       Dim BignDateWork As Date
'       Dim JobTypeName As String
'       Dim JobTypeIDIQ As Integer
'       Dim iqama As String
'       Dim Contract_period As Integer
'     Dim Contract_periodno As Integer
'   Dim dcjopstatus As Integer
'Dim LastDate As Date
'        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, depid, specid, JobTypeID, gradeID, Account_code2, Account_Code, endContractPerMonth, national, , , projectid, , iqama, , , endiqama, , BignDateWork, LastDate, JobTypeName, Contract_period, Contract_periodno, , dcjopstatus, JobTypeIDIQ
        
'          WriteCustomerBalPublic Account_code2, Balance
          
'  lbl(22).Caption = val(Balance)
'Me.Contract_period.ListIndex = Contract_period
'Me.Txtlong.text = Contract_periodno & "     " & Me.Contract_period.text
'          WriteCustomerBalPublic Account_Code, Balance
      '  TxtNuWork.text = JobTypeName
'  lbl(21).Caption = val(Balance)
 ' lbl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
       ' DBIssueDate.value = issuedate
      '  DcboEmpDepartments.BoundText = depid
     ' DcProject.BoundText = projectid
      '  DcboSpecifications.BoundText = gradeID
'        DcboJobsType.BoundText = JobTypeIDIQ
'        lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "", 0)
'        lbl(31).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "", 1)
       ' Txtincrease.text = GetEmployeeSalaryAccordingToComponentName(val(Me.DcboEmpName.BoundText), "", 0)
    '  TxtOther.text = GetEmployeeSalaryAccordingToComponentName(val(Me.DcboEmpName.BoundText), "", 1)
    '    DcNational.text = national
  ' Me.DBEndDate.value = (endiqama)
'Me.dcjopstatus.BoundText = dcjopstatus
     '   Me.IssueDate.value = BignDateWork
       ' Me.TxtIqamaNo.text = iqama
 

'End Sub

' Sub GetEmployeeSalaryAccordingToComponentAll(Emp_id As Integer)
'
'  Dim sql As String
'    Dim mofrad_name As String
'    Dim valuee As Double
'    Dim rs As New ADODB.Recordset
'    Dim Balance As Double
'    Dim Mofradd As String
'    Dim i As Integer
'    Mofradd = ""
'
'    sql = "SELECT     dbo.EmpSalaryComponent.[Value],dbo.mofrdat.mofrad_name,dbo.mofrdat.mofrad_type "
''    sql = sql & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
 '   sql = sql & " dbo.mofrdat ON dbo.EmpSalaryComponent.AccountCode = dbo.mofrdat.mofrad_code"
 '   sql = sql & " WHERE   (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
 '
 '     rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'  With Me.Fg
'  .Rows = rs.RecordCount + 1
'      For i = 1 To rs.RecordCount
'       .TextMatrix(i, .ColIndex("Serial")) = i
'      .TextMatrix(i, .ColIndex("mofrdID")) = IIf(IsNull(rs("mofrad_type").value), 0, rs("mofrad_type").value)
'       .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
' .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(rs("value").value), 0, rs("value").value)
'
'
' rs.MoveNext
'      Next i
' End With
'     End If
     
     

'    rs.Close
    
'End Sub






'Private Sub DcbTO1_Change()
'If TxtModFlg.text <> "R" Then
'If val(Me.DcboEmpName.BoundText) = 0 Then
'MsgBox "«Œ Ì«— «·„ÊŸ› «Ê·«"
'Exit Sub
'End If
'If Me.DcbFrom1.text <> "" And Me.DcbTO1.text <> "" Then
'If val(Me.DcbFrom1.text) >= val(Me.DcbTO1.text) Then
'MsgBox "ÌÃ» «‰ ÌﬂÊ‰ «·Êﬁ  «·«ŒÌ— «ﬂ»— „‰ Êﬁ  «·»œ«ÌÂ"
''DcbTO2.SetFocus
'Exit Sub
'
'Else
'chektime val(Me.DcboEmpName.BoundText), val(Me.DcbFrom1.BoundText), val(Me.DcbTO1.BoundText), 2
'fileFgtim val(Me.DcboEmpName.BoundText), 0
'refiltimdetails val(Me.DcboEmpName.BoundText), 0
'End If
'End If
'End If
'End Sub





'Private Sub DcbTO2_Change()
'If TxtModFlg.text <> "R" Then
'If val(Me.DcboEmpName.BoundText) = 0 Then
'MsgBox "«Œ Ì«— «·„ÊŸ› «Ê·«"
''Exit Sub
'End If
'If Me.DcbFrom2.text <> "" And Me.DcbTO2.text <> "" Then
'If val(Me.DcbFrom2.text) >= val(Me.DcbTO2.text) Then
'MsgBox "ÌÃ» «‰ ÌﬂÊ‰ «·Êﬁ  «·«ŒÌ— «ﬂ»— „‰ Êﬁ  «·»œ«ÌÂ"
''DcbTO2.SetFocus
'Exit Sub
'
'Else
'chektime val(Me.DcboEmpName.BoundText), val(Me.DcbFrom2.BoundText), val(Me.DcbTO2.BoundText), 1
'End If
'End If
'End If
'End Sub

'Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'Dim StrAccountCode As String
'Dim StrAccountCode1 As String
'    Dim Msg As String
'    Dim rs As New ADODB.Recordset
'    Dim StrSQL As String
'    Dim ClsAcc As New ClsAccounts
'    Dim LngRow As Long
'Dim StrComboList As String
'Dim bol As Boolean
'Dim Tye As Integer
'    With Fg
'
'
'
'        Select Case .ColKey(Col)
'
'            Case "empname"
'
'                StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("code"), False, True)
'                ChekRepeat val(StrAccountCode), Row, bol
'                If bol = False Then
'               ' If StrAccountCode <> "" Then
'                .TextMatrix(Row, .ColIndex("empid")) = val(StrAccountCode)
'                StrSQL = " select Fullcode from  TblEmployee where Emp_ID=" & StrAccountCode & ""
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                If rs.RecordCount > 0 Then
'                .TextMatrix(Row, .ColIndex("code")) = rs("Fullcode").value
''                End If
 '               Else
 '               MsgBox "·«Ì„ﬂ‰ «Œ Ì«— «·„‰œÊ» Ê·«Ì„ﬂ‰ «· ﬂ—«—"
 ''               .TextMatrix(Row, .ColIndex("empname")) = ""
  '              .TextMatrix(Row, .ColIndex("code")) = ""
  ''              Exit Sub
   '             End If
   '           If TxtModFlg.text <> "R" Then
'fileFgtim val(StrAccountCode), Tye
'refiltimdetails val(StrAccountCode), Tye
'If Tye = 1 Then
'MsgBox .TextMatrix(Row, .ColIndex("empname")) & "·«Ì„ﬂ‰ «Œ Ì«—"
'.TextMatrix(Row, .ColIndex("empname")) = ""
'                .TextMatrix(Row, .ColIndex("code")) = ""
'              '  Exit Sub
'End If
'End If
'          Case "code"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
'                StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("code"), False, True)
'               If StrAccountCode <> "" Then
'                .TextMatrix(Row, .ColIndex("empid")) = StrAccountCode
'                 StrSQL = " select * from  TblEmployee where Emp_ID=" & StrAccountCode & ""
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                If SystemOptions.UserInterface = ArabicInterface Then
'                .TextMatrix(Row, .ColIndex("empname")) = rs("Emp_Name").value
'                Else
'                .TextMatrix(Row, .ColIndex("empname")) = rs("Emp_Namee").value
'                End If
'                End If
'                   End Select
'
'        If Row = .Rows - 1 Then
'
'            .Rows = .Rows + 1
'        End If
'
        ' ReLineGrid
'    End With

'    ReLineGrid
'End Sub


     
    



Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

'Private Sub menue_Click(Index As Integer)
'If Index = 2 Then
' Load FrmCustemers
'            FrmCustemers.Show
'            End If
'End Sub

 
'Private Sub XPDtbTrans_Change()
'If Me.TxtModFlg.text <> "R" Then
     
'         XPDtbTransH.value = ToHijriDate(XPDtbTrans.value)
       
'End If
'    If Trim(TxtNoteSerial1.text) <> "" Then
'        oldtxtNoteSerial1.text = TxtNoteSerial1.text
'    End If
'
'    TxtNoteSerial.text = ""
'    TxtNoteSerial1.text = ""

'End Sub

Private Sub Dcbranch_Click(Area As Integer)
 
    TxtSerial1.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub Form_Load()
    
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim My_SQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
    StrSQL = " select id,code from currency"
    fill_combo Me.DcbCurrency, StrSQL
'Frame6.Visible = False
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
     Dcombos.GetEmpDepartments Me.DcbDetpartment
    Dcombos.GetTypeExchange Me.DcbExchang
      Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.DcbEmp
        Dcombos.GetEmployees Me.DcbManEmp
    Dcombos.GetEmployees Me.DcbManger
    Dcombos.GetBranches Me.Dcbranch
   ' Dcombos.GetPersons Me.DBCboClientName
   ' Dcombos.GetPersons Me.DBCboClientName

If SystemOptions.UserInterface = ArabicInterface Then
CBoBasedON.AddItem "»·«"
CBoBasedON.AddItem "ÿ·» ‘—«¡"
CBoBasedON.AddItem "«„— ‘—«¡"
DCboCashType121.AddItem "  ⁄„Ì· "
DCboCashType121.AddItem "  „Ê—œ"
DCboCashType121.AddItem "  „ﬁ«Ê· »«ÿ‰"
DCboCashType121.AddItem "  „‘—Ê⁄"
DCboCashType121.AddItem "  „ÊŸ›"
DCboCashType121.AddItem "  Õ”«»"
CBoBasedON.AddItem "„” Õﬁ«  ≈Ã«“…"
CBoBasedON.AddItem "‰Â«Ì… Œœ„…"
CBoBasedON.AddItem "‰„Ê–Ã ‘—«¡ „‰«›”…"
Else
CBoBasedON.AddItem "Without"
CBoBasedON.AddItem "Purchase Request "
CBoBasedON.AddItem "Purchase Order"

DCboCashType121.AddItem "From Customer"
DCboCashType121.AddItem "From Vendor"
DCboCashType121.AddItem "From Contractor Batm"
DCboCashType121.AddItem "From Project"
DCboCashType121.AddItem "From Employee"
DCboCashType121.AddItem "From Account"
CBoBasedON.AddItem "From Vacation Dues"
CBoBasedON.AddItem "From End Service"
Me.CBoBasedON.AddItem "Tender Purchase Form"
End If

    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If


    SetDtpickerDate Me.XPDtbTrans
  '  YearMonth
  Dim InlineSQL As String
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblExchange     WHERE 1=1   "
StrSQL = StrSQL & "    AND  BranchID in(" & Current_branchSql & ")  "

If SystemOptions.FixedCustomer = 1 Then
StrSQL = StrSQL & "    AND  (  UserID = " & user_id & ""
InlineSQL = "SELECT     Transaction_ID From dbo.ApprovalData WHERE     (ScreenName = N'" & Me.Name & "') AND (EmpID = " & user_id & ")"
StrSQL = StrSQL & "    OR   ID IN ( " & InlineSQL & "))"
End If

StrSQL = StrSQL & " Order By ID   "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.text = "R"
            If SystemOptions.UserInterface = EnglishInterface Then
           
        SetInterface Me
        ChangeLang
    End If
    Me.Opt(0).value = False
     Me.Opt(1).value = False
     Me.Opt(2).value = False
    Retrive
   

 

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
  '  Label1.Visible = False
  lbl(18).Caption = "Rate"
  lbl(19).Caption = "Amount"
  XPLbl(14).Caption = "Currency"
  Label11.Caption = "Approved is Required Now"
lbl(16).Caption = "Customer"
lbl(13).Caption = "Remark"
CmdAttach.Caption = "Attachments"
XPTab301.Caption = "Data|Approve Status|Data Analysis"
lbl(15).Caption = "Recipient"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
Command1.Caption = "Show"
lbl(14).Caption = "therefore"
lbl(17).Caption = "Date"
lbl(37).Caption = "Management"
 
Accredit.Caption = "Send Approved "
    Me.Caption = "Issue Request  "
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lblbr.Caption = "Branch"
    Frame2.Caption = "Type Exchange"
    Frame3.Caption = "Data of Exchange "
   lbl(29).Caption = " To"
   lbl(0).Caption = "Amount"
   lbl(2).Caption = "That Vis"

   lbl(9).Caption = "Exchange Requ"
   lbl(10).Caption = "His Manger"
   lbl(12).Caption = "Manager"
   lbl(5).Caption = "Description"
 ISButton4.Caption = "Exit"
Command2.Caption = "Show Order"
   Opt(0).RightToLeft = False
      Opt(1).RightToLeft = False
         Opt(2).RightToLeft = False
        Opt(0).Caption = "Cash"
     Opt(1).Caption = "Check"
     Opt(2).Caption = "Transfer"
  With GRID2
 
  .TextMatrix((0), .ColIndex("Approved")) = "Approved"
  .TextMatrix((0), .ColIndex("levelName")) = "Level"
  .TextMatrix((0), .ColIndex("EmpName")) = "Employee"
  .TextMatrix((0), .ColIndex("ApprovDate")) = "Approv Date"
  .TextMatrix((0), .ColIndex("Remarks")) = "Remarks"
  End With
With Fg
.TextMatrix((0), .ColIndex("selec")) = "Select"
.TextMatrix((0), .ColIndex("Name")) = "Item Name"
.TextMatrix((0), .ColIndex("ShowQty")) = "Qty"
.TextMatrix((0), .ColIndex("QtyAlaw")) = "Qty Avalibal"
.TextMatrix((0), .ColIndex("ExpQty")) = "Expnses Qty"
.TextMatrix((0), .ColIndex("showPrice")) = "Price"
.TextMatrix((0), .ColIndex("Total")) = "Total"
End With

 lbl(8).Caption = "By"
  
        lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
  
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



Private Sub ISButton4_Click()
Frame5.Visible = False
End Sub

Private Sub TxtCurrencyRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
           
           LostAllFocus
        End If
        
End Sub

Private Sub TxtCurrencyRate_KeyUp(KeyCode As Integer, Shift As Integer)
CalCuteCurrency
End Sub

Private Sub TxtManger_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtManger.text, EmpID
        DcbManger.BoundText = EmpID
    End If
End Sub

Private Sub TxtME_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtME.text, EmpID
        DcbManEmp.BoundText = EmpID
    End If
End Sub

'Public Sub retInfoCustomer()
' Dim EmpID As Integer
'Dim name As String
'Dim mobile As String
'Dim phone As String
'Dim boxmail As String
'Dim fax As String
'Dim mail As String
'Dim adress As String
'Dim ZipCode As String
'Dim DigCus As String
'    Dim fullcode As String
'    Dim map As String
'Dim entry As String
'Dim ResponsibleContact As String
'    Dim jobname As String
'        GetCustomerIDFromCode Me.TxtCustomer.text, EmpID, , fullcode, Me.DcbCustomer.text, name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus, jobname, entry, map, ResponsibleContact
'       '  Me.TxtCustomer = fullcode
       '  Me.TxtPersonCont.text = ResponsibleContact
'        Me.DcbCustomer.BoundText = EmpID
'      ' Me.TxtMobi.text = mobile
'        Me.TxtTel.text = phone
'       Me.TxtMap.text = map
'        Me.TxtEnter.text = entry
'        Me.DcbJobID.text = jobname
'        Me.Txtemail.text = mail
'        Me.TxtAdres.text = adress
'        'Me.txtboxzip.text = ZipCode
'
'        'Me.TxtTypeCustomer.text = val(DigCus) + 1
       ' DcboEmpName.BoundText = EmpID
    
'End Sub

'Private Sub TxtCustomer_KeyPress(KeyAscii As Integer)
'If KeyAscii = vbKeyReturn Then
''Me.DcbDelegate.BoundText = ""
'retInfoCustomer
'End If
'End Sub
Function GetQty(Optional ItemID As Integer = 0, Optional TransID As Double) As Double
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     Transaction_ID, ItemID, SUM(ExpQty) AS sm"
sql = sql & " From dbo.TblExchangeDet"
sql = sql & " Where (Transaction_ID =" & TransID & ") And (ItemID = " & ItemID & ")"
sql = sql & " GROUP BY Transaction_ID, ItemID"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
Rs7.MoveFirst
GetQty = IIf(IsNull(Rs7("sm").value), 0, Rs7("sm").value)
Else
GetQty = 0
End If
End Function
Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
        Frame5.Enabled = True
        Frame1.Enabled = True
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
   
   
    '    Accredit.Enabled = True
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  "
            'Me.menue(2).Enabled = True
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
Accredit.Enabled = False
        Case "E"
        Accredit.Enabled = False
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

 
Sub Retrive_PO10(Transaction_ID As Double)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
Dim i As Integer
    Dim row_count As Double
    Dim Num As Double
    
StrSQL = "SELECT     dbo.TblItems.HaveSerial AS Expr1, *"
StrSQL = StrSQL & " FROM         dbo.TblItems INNER JOIN"
StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesData ON dbo.Transaction_Details.countrisid = dbo.TblCountriesData.CountryID"
StrSQL = StrSQL & " Where (dbo.Transaction_Details.Transaction_ID = " & Transaction_ID & ")"
    
    Set rs2 = New ADODB.Recordset
    rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs2.RecordCount < 1 Then

        Exit Sub
    Else
    With Fg
  '  Total
    .rows = rs2.RecordCount + 1
    rs2.MoveFirst
    Dim k As Integer
    k = 0
    For i = .FixedRows To .rows - 1
    If GetQty(val(val(rs2("Item_ID").value)), Transaction_ID) < val(rs2("ShowQty").value) Then
    k = k + 1
    .TextMatrix(k, .ColIndex("itemid")) = IIf(IsNull(rs2("Item_ID").value), 0, rs2("Item_ID").value)
    .TextMatrix(k, .ColIndex("ShowQty")) = IIf(IsNull(rs2("ShowQty").value), 0, rs2("ShowQty").value)
    .TextMatrix(k, .ColIndex("showPrice")) = IIf(IsNull(rs2("showPrice").value), 0, rs2("showPrice").value)
   .TextMatrix(k, .ColIndex("QtyAlaw")) = val(.TextMatrix(i, .ColIndex("ShowQty"))) - GetQty(val(.TextMatrix(i, .ColIndex("itemid"))), Transaction_ID)
    
    If SystemOptions.UserInterface = ArabicInterface Then
    .TextMatrix(k, .ColIndex("Name")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
    Else
    .TextMatrix(k, .ColIndex("Name")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
    End If
  End If
    rs2.MoveNext
    Next i
    .AutoSize 0, .Cols - 1, False
    End With
    
     End If

 End Sub


Function Retrive_orders_data(Transaction_ID As Double, Optional str As String)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim row_count As Double
    Dim Num As Double

'    StrSQL = "Select * from transactions where Transaction_ID=" & Transaction_ID
    
StrSQL = "SELECT     QryTransactionsTotal.TransNet, dbo.Transactions.Transaction_ID"
StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & " dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID"
StrSQL = StrSQL & "  WHERE     ( requestOrOrder=0 and  dbo.Transactions.Transaction_ID = " & Transaction_ID & ")"
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 TxtPrice.text = 0
        Exit Function
    Else
    TxtPrice.text = IIf(IsNull(rs("TransNet").value), 1, (rs("TransNet").value))
       ' DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
       ' Me.Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), 1, rs("Currency_id").value)
        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
        'TxtLcNo.text = IIf(IsNull(rs("LcNo").value), "", (rs("LcNo").value))
    End If

 End Function


Private Sub TxtOrderNo_Change()
If Me.TxtModFlg = "N" Or TxtModFlg.text = "E" Then
TxtDes.text = ""
If val(CBoBasedON.ListIndex = 1) Then
If SystemOptions.UserInterface = ArabicInterface Then
TxtDes.text = "»‰«¡ ⁄·Ì ÿ·» ‘—«¡ »—ﬁ„ " & TxtOrderNo
Else
TxtDes.text = "Based on Request No " & TxtOrderNo
End If
ElseIf val(CBoBasedON.ListIndex) = 2 Then
Retrive_PO10 val(txtTransaction_ID.text)
ElseIf CBoBasedON.ListIndex = 3 Then
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
 If CheckVacationPayed(val(TxtOrderNo.text)) = True Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox " „ ⁄„· ”‰œ ’—› ·Â–Â «·Õ—ﬂ… „‰ ﬁ»· "
 Else
 MsgBox "It Was Paid"
 End If
 Exit Sub
 End If
 End If
 
Me.RetriveVacation val(TxtOrderNo.text)
If SystemOptions.UserInterface = ArabicInterface Then
TxtDes.text = TxtDes.text & "»‰«¡ ⁄·Ì „” Õﬁ«  ≈Ã«“…  »—ﬁ„ " & TxtOrderNo
Else
TxtDes.text = TxtDes.text & "Based on Vacation Entitlements " & TxtOrderNo
End If
ElseIf CBoBasedON.ListIndex = 4 Then

If SystemOptions.UserInterface = ArabicInterface Then
TxtDes.text = "»‰«¡⁄·Ï ‰Â«Ì… Œœ„…  »—ﬁ„ " & TxtOrderNo
Else
TxtDes.text = "Based on End Service     " & TxtOrderNo
End If

Me.RetriveEndSevice val(TxtOrderNo.text)
ElseIf CBoBasedON.ListIndex = 5 Then
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
 If CheckTenderPayed(val(TxtOrderNo.text)) = True Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox " „ ⁄„· ”‰œ ’—› ·Â–Â «·Õ—ﬂ… „‰ ﬁ»· "
 Else
 MsgBox "It Was Paid"
 End If
 Exit Sub
 End If
 
If SystemOptions.UserInterface = ArabicInterface Then
TxtDes.text = "»‰«¡⁄·Ï  ‰„Ê–Ã ‘—«¡ „‰«›”…  »—ﬁ„ " & TxtOrderNo
Else
TxtDes.text = "Based on  Tender Purchase Form      " & TxtOrderNo
End If

Me.RetriveTender val(TxtOrderNo.text)
End If
'If SystemOptions.UserInterface = ArabicInterface Then
'TxtDes.text = TxtDes.text & "»‰«¡ ⁄·Ì „” Õﬁ«  ≈Ã«“…  »—ﬁ„ " & TxtOrderNo
'Else
'TxtDes.text = TxtDes.text & "Based on Vacation Entitlements " & TxtOrderNo
'End If
End If
End If

End Sub
Function ChekPayment() As Boolean
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
Dim sql As String
ChekPayment = False
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "Select id from  End_of_service where id=" & val(txtTransaction_ID.text) & " and PaymPaid=1 "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekPayment = True
Else
ChekPayment = False
End If
End If
End Function
Sub RetrivOrder(Optional TransID As Integer)
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If TransID <> 0 Then
sql = " SELECT     Transaction_ID, Transaction_Serial, DeptID, Transaction_Date, Transaction_Type"
sql = sql & " From dbo.Transactions"
sql = sql & " Where(Transaction_ID = " & TransID & ")"
  Rs7.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
  If Rs7.RecordCount > 0 Then
  DcbDetpartment.BoundText = IIf(IsNull(Rs7("DeptID").value), "", Rs7("DeptID").value)
   ReqDate.value = IIf(IsNull(Rs7("Transaction_Date").value), Date, Rs7("Transaction_Date").value)
   End If
   End If
  
End Sub
Private Sub TxtOrderNo_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then

     
                If CBoBasedON.ListIndex = 0 Then
                ElseIf CBoBasedON.ListIndex = 1 Then
                
         '   TxtOrderNo.text = ""
         '   Order_no_search.show
         '   Order_no_search.RetrunType = 15
         '
            
                               FrmBuySearch.DealingForm = GridTransType.purchaserequest
                  FrmBuySearch.index = 19
                  If SystemOptions.UserInterface = ArabicInterface Then
                    FrmBuySearch.Caption = "«·»ÕÀ ⁄‰  ÿ·»  «·‘—«¡"
                    Else
                    FrmBuySearch.Caption = "Search Purchase Request"
                    End If
                   FrmBuySearch.show vbModal
               
               
               
                          ElseIf CBoBasedON.ListIndex = 2 Then
                
     Frame5.Visible = True
                  FrmBuySearch.DealingForm = GridTransType.purchaseOrderApproved
                  FrmBuySearch.index = 18
                  If SystemOptions.UserInterface = ArabicInterface Then
                   FrmBuySearch.Caption = "«·»ÕÀ ⁄‰  «Ê«„—  «·‘—«¡"
                   Else
                   FrmBuySearch.Caption = "Search Purchase Order"
                   End If
                   FrmBuySearch.show vbModal
               
               ''''//////////
         
ElseIf val(CBoBasedON.ListIndex) = 3 Then
Load FrmSearchVocationEntitlement
FrmSearchVocationEntitlement.index = 1
            FrmSearchVocationEntitlement.show
            
ElseIf val(CBoBasedON.ListIndex) = 4 Then
  Load FrmEnserviceSearch
            FrmEnserviceSearch.show
            FrmEnserviceSearch.index = 1
ElseIf val(CBoBasedON.ListIndex) = 5 Then
    Unload FrmInsurancesSearch
            FrmInsurancesSearch.BankInx = 606
            FrmInsurancesSearch.SendForm = 6
            FrmInsurancesSearch.show
 

               '''//////
            
        '    Order_no_search.lblSpecificsearch.Caption = val(cbobasedOn.ListIndex)
            End If
      

    End If

End Sub


Function LostAllFocus()
 TxtPrice.text = Format(TxtPrice.text, "#,##0.00")
TxtPriceE.text = Format(TxtPriceE.text, "#,##0.00")

End Function
 
 

Private Sub txtPrice_Change()
Me.lbl(3).Caption = WriteNo(Format(TxtPrice.text, "###.00"), 0)
End Sub

Private Sub TxtPrice_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
           
           LostAllFocus
        End If
End Sub


Private Sub TxtPriceE_Change()
'Me.lbl(3).Caption = WriteNo(Format(TxtPriceE.Text, "###.00").Text, 0, , , , , , val(DcbCurrency.BoundText))
Me.lbl(20).Caption = WriteNo(Format(TxtPriceE.text, "###.00"), 0, , , , , , val(DcbCurrency.BoundText))

End Sub

Private Sub TxtPriceE_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
           
           LostAllFocus
        End If
End Sub

Private Sub TxtPricee_GotFocus()
 
TxtPriceE.text = Format(TxtPriceE.text, "###.00")


End Sub
Private Sub TxtPrice_GotFocus()
TxtPrice.text = Format(TxtPrice.text, "###.00")
End Sub

Private Sub TxtPrice_KeyUp(KeyCode As Integer, Shift As Integer)


CalCuteCurrency
End Sub
Private Sub TxtPricee_LostFocus()
LostAllFocus
End Sub
Private Sub TxtPrice_LostFocus()
LostAllFocus
End Sub

Sub CalCuteCurrencyE()
'Exit Sub
If val(TxtCurrencyRate.text) = 0 Then
TxtCurrencyRate.text = ""
End If
TxtPrice.text = Round(val(Format(TxtPriceE.text, "###.00")) * val(TxtCurrencyRate.text), 2)

TxtPrice.text = Format(TxtPrice.text, "#,##0.00")

If val(TxtPriceE.text) = 0 Then TxtPrice.text = "": Exit Sub
 
'TxtPriceE.Text = Format(TxtPriceE.Text, "#,##0.00")

End Sub

Private Sub TxtPricee_Keyup(KeyCode As Integer, Shift As Integer)
CalCuteCurrencyE
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcbEmp.BoundText = EmpID
    End If

End Sub



Private Sub txtTransaction_ID_Change()
Dim Transaction_Type As Integer
Dim Transaction_ID As String
   ' Transaction_ID = get_transactionData("order_no", TxtOrderNo.text, "Transaction_ID", Transaction_Type)

    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
      Transaction_Type = 29
  TxtDes.text = ""
    If val(CBoBasedON.ListIndex) = 3 Then

    If val(Me.txtTransaction_ID.text) <> 0 Then
    Me.RetriveVacation (val(Me.txtTransaction_ID.text))
    Else
    Me.RetriveVacation (val(Me.TxtOrderNo.text))
    End If
   If SystemOptions.UserInterface = ArabicInterface Then
   TxtDes.text = TxtDes.text & "»‰«¡ ⁄·Ì „” Õﬁ«  ≈Ã«“…  »—ﬁ„ " & txtTransaction_ID
   Else
   TxtDes.text = TxtDes.text & "Based on Vacation Entitlements " & txtTransaction_ID
   End If
   ElseIf val(CBoBasedON.ListIndex) = 2 Then
Retrive_PO10 val(txtTransaction_ID.text)
    ElseIf val(CBoBasedON.ListIndex) = 4 Then
   
   If val(Me.txtTransaction_ID.text) <> 0 Then
   If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
If ChekPayment() = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Â–Â «·⁄„·Ì…  „ ”œ«œÂ« „”»ﬁ« Ì—ÃÏ «Œ Ì«— ⁄„·Ì… «Œ—Ï"
Else
MsgBox "This process is already paid"
End If
End If
End If
    Me.RetriveEndSevice (val(Me.txtTransaction_ID.text))
    Else
    Me.RetriveEndSevice (val(Me.TxtOrderNo.text))
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
     TxtDes.text = TxtDes.text & "»‰«¡ ⁄·Ï ‰Â«Ì… Œœ„…  »—ﬁ„ " & txtTransaction_ID
     Else
     TxtDes.text = TxtDes.text & "Based on End Service " & txtTransaction_ID
     End If
    Else
        Retrive_orders_data (val(Me.txtTransaction_ID.text))
        RetrivOrder (val(Me.txtTransaction_ID.text))
    End If
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
     Dim RsDetails1 As ADODB.Recordset
   Dim ContactTime As Date
    Dim i As Integer
    Dim StrSQL As String
      Frame5.Visible = False

    Me.Opt(0).value = False
     Me.Opt(1).value = False
      Me.Opt(2).value = False
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
 
 

    XPTxtID.text = IIf(IsNull(rs("Id").value), "", (rs("Id").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
 
    Me.TxtSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
   ''/
  DCboCashType121.ListIndex = IIf(IsNull(rs("FromType").value), -1, rs("FromType").value)
DCEmployee.BoundText = IIf(IsNull(rs("EmpID1").value), 0, rs("EmpID1").value)
DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
DCAccounts.BoundText = IIf(IsNull(rs("Account_Code").value), 0, rs("Account_Code").value)

   ''//
 Me.DcbCurrency.BoundText = IIf(IsNull(rs("CurrcyID").value), MainCurrency, rs("CurrcyID").value)
 Me.TxtPriceE.text = IIf(IsNull(rs("PriceE").value), 0, rs("PriceE").value)
 TxtPriceE.text = Format(TxtPriceE.text, "#,##0.00")
 Me.TxtCurrencyRate.text = IIf(IsNull(rs("Rate").value), 1, rs("Rate").value)
 DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
 Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
 Me.DcbEmp.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
 Me.DcbManEmp.BoundText = IIf(IsNull(rs("MempID").value), "", rs("MempID").value)
    Me.DcbManger.BoundText = IIf(IsNull(rs("ManagerID").value), "", rs("ManagerID").value)
    Me.DcbExchang.BoundText = val(IIf(IsNull(rs("TypeExch").value), "", rs("TypeExch").value))
Me.TxtDes.text = IIf(IsNull(rs("Des").value), "", rs("Des").value)
Me.TxtDes2.text = IIf(IsNull(rs("Des2").value), "", rs("Des2").value)
Me.TxtBankIBAN.text = IIf(IsNull(rs("BankIBAN").value), "", rs("BankIBAN").value)

 Me.DcbDetpartment.BoundText = IIf(IsNull(rs("DeptID").value), "", rs("DeptID").value)
ReqDate.value = IIf(IsNull(rs("ReqDate").value), Date, rs("ReqDate").value)
If Not IsNull(rs("basedOn").value) Then
CBoBasedON.ListIndex = IIf(IsNull(rs("basedOn").value), 1, rs("basedOn").value)
Else
CBoBasedON.ListIndex = 0
End If

Me.TxtOrderNo.text = IIf(IsNull(rs("OrderNo").value), "", rs("OrderNo").value)
Me.txtTransaction_ID.text = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)


Me.txtperson.text = IIf(IsNull(rs("ToPerson").value), "", rs("ToPerson").value)
Me.TxtPrice.text = IIf(IsNull(rs("Price").value), "", rs("Price").value)
TxtPrice.text = Format(TxtPrice.text, "#,##0.00")

TxtPriceE.text = Format(TxtPriceE.text, "#,##0.00")
Frame7.Visible = False
If DCboCashType121.ListIndex = 4 Then
Frame7.Visible = True
        If val(rs("salary_or_advance").value & "") = 1 Then
              Option4.value = True
        ElseIf (rs("salary_or_advance").value) = 0 Then
            Option5.value = True
       
        ElseIf (rs("salary_or_advance").value) = 1 Then
            Option5.value = True
       
        ElseIf (rs("salary_or_advance").value) = 2 Then
            Option6.value = True
       
        ElseIf (rs("salary_or_advance").value) = 3 Then
            Option7.value = True
       
        End If
 End If
 

 If val(rs("Type").value) = 0 Then
Me.Opt(0).value = True
ElseIf val(rs("Type").value) = 1 Then
Me.Opt(1).value = True
ElseIf val(rs("Type").value) = 2 Then
Me.Opt(2).value = True
End If
'       If IsNull(rs("posted").value) Then
'                                                   If SystemOptions.UserInterface = ArabicInterface Then
'                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
'                                                  Else
'                                                    Accredit.Caption = " send to Approval   "
'                                               End If
'                                               Accredit.Enabled = True
'  Else
''                                                   If SystemOptions.UserInterface = ArabicInterface Then
 '                                                   Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
 '                                                 Else
 ''                                                   Accredit.Caption = " sent to Approval   "
  '                                             End If
  '                                             Accredit.Enabled = False
  ' End If
  '
  '
  If val(Me.CBoBasedON.ListIndex) = 2 Then
   Set RsDetails = New ADODB.Recordset
 StrSQL = "SELECT     dbo.TblExchangeDet.ID, dbo.TblExchangeDet.ExhID, dbo.TblExchangeDet.Transaction_ID, dbo.TblExchangeDet.QtyAlaw, dbo.TblExchangeDet.ShipQty,"
 StrSQL = StrSQL & "                     dbo.TblExchangeDet.ShipPrice, dbo.TblExchangeDet.ExpQty, dbo.TblExchangeDet.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemCode,"
 StrSQL = StrSQL & "                     dbo.TblItems.ItemNamee"
 StrSQL = StrSQL & " FROM         dbo.TblExchangeDet LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblItems ON dbo.TblExchangeDet.ItemID = dbo.TblItems.ItemID"
 StrSQL = StrSQL & " Where (dbo.TblExchangeDet.ExhID =" & val(Me.XPTxtID.text) & " ) "
 RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Fg.Clear flexClearScrollable, flexClearEverything
   Fg.rows = Fg.FixedRows

   If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        Fg.rows = Fg.FixedRows + RsDetails.RecordCount

        For i = Me.Fg.FixedRows To Fg.rows - 1
        Fg.TextMatrix(i, Fg.ColIndex("itemid")) = IIf(IsNull(RsDetails("ItemID").value), "", RsDetails("ItemID").value)
        Fg.TextMatrix(i, Fg.ColIndex("ShowQty")) = IIf(IsNull(RsDetails("ShipQty").value), "", RsDetails("ShipQty").value)
           If SystemOptions.UserInterface = EnglishInterface Then
           Fg.TextMatrix(i, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemName").value), "", RsDetails("ItemName").value)
          Else
           Fg.TextMatrix(i, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemNamee").value), "", RsDetails("ItemNamee").value)
          End If
           Fg.TextMatrix(i, Fg.ColIndex("QtyAlaw")) = IIf(IsNull(RsDetails("QtyAlaw").value), "", RsDetails("QtyAlaw").value)
           Fg.TextMatrix(i, Fg.ColIndex("ExpQty")) = IIf(IsNull(RsDetails("ExpQty").value), 0, RsDetails("ExpQty").value)
           Fg.TextMatrix(i, Fg.ColIndex("showPrice")) = IIf(IsNull(RsDetails("ShipPrice").value), 0, RsDetails("ShipPrice").value)
           Fg.TextMatrix(i, Fg.ColIndex("Total")) = val(Fg.TextMatrix(i, Fg.ColIndex("showPrice"))) * val(Fg.TextMatrix(i, Fg.ColIndex("ExpQty")))
           Fg.cell(flexcpChecked, i, Fg.ColIndex("selec")) = flexChecked
            RsDetails.MoveNext
        Next i

    End If

    RsDetails.Close
    Set RsDetails = Nothing
  End If
  GetIBan
   '''''''''''''///////////////////////
'   Set RsDetails1 = New ADODB.Recordset
' StrSQL = "SELECT     dbo.TblRegDateDelgateDails.Id, dbo.TblRegDateDelgateDails.DelgID, dbo.TblRegDateDelgateDails.EmpID, dbo.TblRegDateDelgateDails.remark, "
'StrSQL = StrSQL & "                      dbo.TblRegDateDelgateDails.Type , dbo.TblCompo.name, dbo.TblCompo.namee, dbo.TblRegDateDelgateDails.Quantity"
'StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgateDails LEFT OUTER JOIN"
'  StrSQL = StrSQL & "                    dbo.TblCompo ON dbo.TblRegDateDelgateDails.EmpID = dbo.TblCompo.Id"
'
'StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.Type = 1) And (dbo.TblRegDateDelgateDails.DelgID = " & val(Me.XPTxtID.text) & " )"



' RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
'    fg2.Clear flexClearScrollable, flexClearEverything
'    fg2.Rows = fg2.FixedRows
'
'    If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
'        RsDetails1.MoveFirst
'        fg2.Rows = fg2.FixedRows + RsDetails1.RecordCount
'
'        For i = Me.fg2.FixedRows To fg2.Rows - 1
'        fg2.TextMatrix(i, fg2.ColIndex("Serial")) = i
'        fg2.TextMatrix(i, fg2.ColIndex("remarks")) = IIf(IsNull(RsDetails1("remark").value), "", RsDetails1("remark").value) ' RsDetails1("remark").value
'            fg2.TextMatrix(i, fg2.ColIndex("code")) = IIf(IsNull(RsDetails1("quantity").value), "", RsDetails1("quantity").value) 'RsDetails1("fullcode").value
'            If SystemOptions.UserInterface = EnglishInterface Then
'           fg2.TextMatrix(i, fg2.ColIndex("name")) = IIf(IsNull(RsDetails1("namee").value), "", RsDetails1("namee").value) 'RsDetails1("Emp_Namee").value
''           Else
 '          fg2.TextMatrix(i, fg2.ColIndex("name")) = IIf(IsNull(RsDetails1("name").value), "", RsDetails1("name").value) ' RsDetails1("emp_name").value
 '          End If
 '           fg2.TextMatrix(i, fg2.ColIndex("empid")) = RsDetails1("EmpID").value
 '           RsDetails1.MoveNext
 '       Next i
'
'    End If

'    RsDetails1.Close
'    Set RsDetails1 = Nothing
   
   
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
   
   
   
 '  fillapprovData
    fillapprovData
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

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
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
   GRID2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    GRID2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    GRID2.ColComboList(GRID2.ColIndex("Show")) = "..."
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
 GRID2.rows = 1
    End If
RsDetails.Close

End Function

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
Dim StrSQL1 As String
    'On Error GoTo ErrTrap

     If Me.TxtModFlg.text <> "R" Then
        If Me.DcbEmp.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ ÿ«·» «·’—›..!! "
            Else
            Msg = "Please Select Name"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcbEmp.SetFocus
           Sendkeys "{F4}"
            Exit Sub
        End If
   If Me.DcbExchang.BoundText = "" Then
   If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ   ⁄»«—…  «·’—›..!! "
        Else
        Msg = "Please select Type"
        End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            DcbExchang.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
If Me.TxtModFlg.text = "E" Then
                StrSQL1 = "Delete From TblExchangeDet Where ExhID=" & val(Me.XPTxtID.text)
 Cn.Execute StrSQL1, , adExecuteNoRecords
End If

        my_branch = val(Me.Dcbranch.BoundText)
Dim notserial1str As String

    If TxtSerial1.text = "" Then
    notserial1str = Voucher_coding(val(my_branch), XPDtbTrans.value, 58, 58)
 
                            If notserial1str = "error" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—›  ⁄œÌ  «·Õœ «·«ﬁ’Ì": Exit Sub
                                            Else
                                                MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                                            End If
                    
                            ElseIf notserial1str = "" Then
                                                    If SystemOptions.UserInterface = ArabicInterface Then
                                                        MsgBox "  ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ": Exit Sub
                                                    Else
                                                        MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                                                    End If
                    
                                Else
                                    TxtSerial1.text = notserial1str
                                End If
    
    End If





     Dim RsTest As New ADODB.Recordset

        Cn.BeginTrans
        BeginTrans = True
        
              If TxtModFlg.text = "N" Then


        '”·› ”«»ﬁ…
   


            XPTxtID.text = CStr(new_id("TblExchange", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
       ' ElseIf Me.TxtModFlg.text = "E" Then
       '     StrSQL = "Delete From TblRegDateDelgateDails Where DelgID=" & val(Me.XPTxtID.text)
       '     Cn.Execute StrSQL, , adExecuteNoRecords

        End If
           rs("ID").value = val(XPTxtID.text)
         rs("NoteSerial1").value = (Me.TxtSerial1.text)
         rs("CurrcyID").value = val((Me.DcbCurrency.BoundText))
         rs("RecordDate").value = XPDtbTrans.value
         rs("FromType").value = IIf(Me.DCboCashType121.ListIndex = -1, Null, Me.DCboCashType121.ListIndex)
     If val(DCboCashType121.ListIndex) = 0 Or val(DCboCashType121.ListIndex) = 1 Or val(DCboCashType121.ListIndex) = 2 Or val(DCboCashType121.ListIndex) = 3 Then
     rs("CusID").value = IIf(Me.DBCboClientName.BoundText = "", Null, val(Me.DBCboClientName.BoundText))
     End If
     If val(DCboCashType121.ListIndex) = 4 Then
     rs("EmpID1").value = IIf(Me.DCEmployee.BoundText = "", Null, val(Me.DCEmployee.BoundText))
     End If
     If val(DCboCashType121.ListIndex) = 5 Then
     rs("Account_Code").value = IIf(Me.DCAccounts.BoundText = "", Null, Me.DCAccounts.BoundText)
     End If
           rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
           rs("EmpID").value = IIf(Me.DcbEmp.BoundText = "", Null, Me.DcbEmp.BoundText)
           rs("ManagerID").value = IIf(Me.DcbManger.BoundText = "", Null, DcbManger.BoundText)
           rs("MempID").value = IIf(Me.DcbManEmp.BoundText = "", Null, DcbManEmp.BoundText)
           rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, DCboUserName.BoundText)
           rs("TypeExch").value = IIf(Me.DcbExchang.BoundText = "", Null, val(DcbExchang.BoundText))
           rs("Des").value = IIf(Me.TxtDes.text = "", "", Me.TxtDes.text)
           rs("Des2").value = IIf(Me.TxtDes2.text = "", "", Me.TxtDes2.text)
           rs("BankIBAN").value = IIf(Me.TxtBankIBAN.text = "", "", Me.TxtBankIBAN.text)
           
           rs("ToPerson").value = IIf(Me.txtperson.text = "", "", Me.txtperson.text)
           
         'new idea **********************
          TxtPrice.text = Format(TxtPrice.text, "###.00")
            rs("Price").value = val(IIf(Me.TxtPrice.text = "", 0, Me.TxtPrice.text))
          TxtPrice.text = Format(TxtPrice.text, "#,##0.00")
           'new idea **********************
           
           rs("OrderNo").value = val(IIf(Me.TxtOrderNo.text = "", "", Me.TxtOrderNo.text))
           rs("Transaction_ID").value = val(IIf(Me.txtTransaction_ID.text = "", "", Me.txtTransaction_ID.text))
           rs("DeptID").value = IIf(Me.DcbDetpartment.BoundText = "", Null, Me.DcbDetpartment.BoundText)
           rs("ReqDate").value = ReqDate.value
           rs("basedOn").value = CBoBasedON.ListIndex
           'new idea **********************
          TxtPriceE.text = Format(TxtPriceE.text, "###.00")
           rs("PriceE").value = val(IIf(Me.TxtPriceE.text = "", 0, val(Me.TxtPriceE.text)))
          TxtPriceE.text = Format(TxtPriceE.text, "#,##0.00")
           'new idea **********************
        Dim lblflag As Integer
                   If Option4.value = True Then
        lblflag = 1
       ElseIf Option5.value = True Then
        lblflag = 0

       ElseIf Option6.value = True Then
        lblflag = 2
      ElseIf Option7.value = True Then
        lblflag = 3
       End If
          rs("salary_or_advance").value = lblflag


        
        
           
           rs("Rate").value = val(IIf(Me.TxtCurrencyRate.text = "", 0, val(TxtCurrencyRate.text)))
 
          If Opt(0).value = True Then
           rs("Type").value = 0
         End If
         If Opt(1).value = True Then
          rs("Type").value = 1
        End If
        If Opt(2).value = True Then
          rs("Type").value = 2
        End If
        rs.update
     If val(Me.CBoBasedON.ListIndex) = 2 Then
           Set RsDetails = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblExchangeDet Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     For i = Me.Fg.FixedRows To Fg.rows - 1
       If Fg.cell(flexcpChecked, i, Fg.ColIndex("selec")) = flexChecked Then
           RsDetails.AddNew
           RsDetails("ExhID").value = val(XPTxtID.text)
           RsDetails("Transaction_ID").value = val(txtTransaction_ID.text)
           RsDetails("ItemID").value = Fg.TextMatrix(i, Fg.ColIndex("itemid"))
           RsDetails("ShipQty").value = val(Fg.TextMatrix(i, Fg.ColIndex("ShowQty")))
           RsDetails("QtyAlaw").value = val(Fg.TextMatrix(i, Fg.ColIndex("QtyAlaw")))
           RsDetails("ExpQty").value = val(Fg.TextMatrix(i, Fg.ColIndex("ExpQty")))
           RsDetails("ShipPrice").value = val(Fg.TextMatrix(i, Fg.ColIndex("showPrice")))
           RsDetails("selec").value = 1
           RsDetails.update
        End If
        
       Next i
 End If
  ''///////////'''''''''''''''''''''''''''''''
   '     Set RsDetails1 = New ADODB.Recordset
   ''    StrSQL = "SELECT     *  from dbo.TblRegDateDelgateDails Where (1 = -1)"
 '  RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
    
 '       For i = Me.fg2.FixedRows To fg2.Rows - 1
 '      If val(fg2.TextMatrix(i, fg2.ColIndex("EmpID"))) <> 0 Then
 '           RsDetails1.AddNew
 '           RsDetails1("DelgID").value = val(XPTxtID.text)
 '           RsDetails1("Type").value = 1
 '          RsDetails1("remark").value = fg2.TextMatrix(i, fg2.ColIndex("remarks"))
 '           RsDetails1("EmpID").value = val(fg2.TextMatrix(i, fg2.ColIndex("empid")))
 '   RsDetails1("quantity").value = val(fg2.TextMatrix(i, fg2.ColIndex("code")))
 '           RsDetails1.update
 '       End If
 '       Next i
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
     '   BeginTrans = False
     '   RsDetails.Close
        
     '   Set RsDetails = Nothing
     '   RsDetails1.Close
     '   Set RsDetails1 = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       If val(CBoBasedON.ListIndex) = 5 Then
           Cn.Execute "Update TblBankPledge4 set PayedPayment=1 where ID=" & val(TxtOrderNo.text) & ""
      End If
        Select Case Me.TxtModFlg.text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
Else
                Msg = " This is record Allredy saved " & CHR(13)
                Msg = Msg + "Yu need to enter another record"
End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                 MsgBox "Saved Successfullty", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        Msg = "Can Not Save this Data " & CHR(13)
        Msg = Msg + "Make sure the validity of the data and try again " & CHR(13)
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
     Msg = "Sorry...an error occurred while saving " & CHR(13)
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
            rs.Find "ID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

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
Dim StrSQL1 As String
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
        Msg = "You will be deleting data " & CHR(13)
        Msg = Msg + " Confirm Delete?"
End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                    Deletepost Me.Name, "TblExchange", "Id", val(DcbDetpartment.BoundText), val(Dcbranch.BoundText), val(XPTxtID.text), TxtSerial1
                    
                StrSQL = "Delete From TblExchange Where ID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                          StrSQL1 = "Delete From TblExchangeDet Where ExhID=" & val(Me.XPTxtID.text)
 Cn.Execute StrSQL1, , adExecuteNoRecords
                rs.MoveFirst
                If val(CBoBasedON.ListIndex) = 5 Then
           Cn.Execute "Update TblBankPledge4 set PayedPayment=null where ID=" & val(TxtOrderNo.text) & ""
               End If
              '  StrSQL1 = "Delete From TblDefinDetails Where IDDef=" & val(Me.XPTxtID.text)
 'Cn.Execute StrSQL1, , adExecuteNoRecords
                If rs.RecordCount < 1 Then
                    clear_all Me
                    ' Fg.Clear flexClearScrollable, flexClearEverything
           ' Fg.Rows = 2
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
         Msg = "This process is not available does not have any record"
        
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
        Msg = "Sorry an error occurred during the deletion " & CHR(13)
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
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
'  Dim rs1 As New ADODB.Recordset
' Dim i As Integer
'    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
'  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
'  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
'  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
'  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
'sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.name & "')"
'sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "
'
'    rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs1.RecordCount > 0 Then
'            currentdate = Now
'            For i = 1 To rs1.RecordCount
'              RSApproval.AddNew
'                RSApproval("ScreenName").value = Me.name
'                RSApproval("levelo").value = IIf(IsNull(rs1("levelo").value), Null, rs1("levelo").value)
'               RSApproval("EmpID").value = IIf(IsNull(rs1("EmpID").value), Null, rs1("EmpID").value)
'                RSApproval("levelorder").value = IIf(IsNull(rs1("levelorder").value), Null, rs1("levelorder").value)
'                 RSApproval("currorder").value = IIf(IsNull(rs1("currorder").value), Null, rs1("currorder").value)
'                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
'                   RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
'                RSApproval("Transaction_Date").value = Date
                
'                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.name), currentdate)
'               RSApproval("SendTime").value = currentdate
'
'                 If i = 1 Then
'                        RSApproval("Currcursor").value = 1
'                         RSApproval("FromUser").value = user_name
'                End If
'
'                RSApproval.update
'                rs1.MoveNext
'            Next i
'
'    End If
    
    

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
'                            Label11.BackColor = &H80FF80
'        Else
'                             If SystemOptions.UserInterface = ArabicInterface Then
'                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
'                            Else
'                                     Label11.Caption = "Currently required Approve"
'                            End If
'                 Label11.BackColor = &HFFFFC0
'        End If
'
'End If

'        Next Num
'Else
' GRID2.Rows = 1
'    End If
'RsDetails.Close
'
'End Function
'Private Sub ChekRepeat(Optional ind As Integer, Optional Row As Long, Optional ByRef bo As Boolean)
'    Dim i As Integer
'
'
'    With fg2
' bo = False
'        For i = .FixedRows To .Rows - 1
'If i <> Row Then
'            If val(.TextMatrix(i, .ColIndex("empid"))) = val(ind) Then
'             bo = True
'   End If
'            End If
'            Next i
'            End With
'        With Fg
' bo = False
'        For i = .FixedRows To .Rows - 1
'If i <> Row Then
'            If val(.TextMatrix(i, .ColIndex("empid"))) = val(ind) Then
'             bo = True
'             End If
'             Else
             
'            If val(ind) = val(Me.DcboEmpName.BoundText) Then
'              bo = True
'              End If
'   End If
'
'            Next i
'            End With
'        End Sub

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
    If SystemOptions.UserInterface = ArabicInterface Then
        .Create Me.hWnd, "    ‘«‘… ÿ·» ’—›  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        Else
         
        .Create Me.hWnd, "   Screen Request Exchange  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "New ..." & Wrap & "·Add a new process data" & Wrap & " Just click here", True
        End If
    End With

    With TTP
        If SystemOptions.UserInterface = ArabicInterface Then
        .Create Me.hWnd, "    ‘«‘… ÿ·» ’—›  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        Else
         
        .Create Me.hWnd, "   Screen Request Exchange  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), "Edite ..." & Wrap & "Edite Data of this  process " & Wrap & " Just click here", True
        End If
      
    End With

    With TTP
              If SystemOptions.UserInterface = ArabicInterface Then
        .Create Me.hWnd, "    ‘«‘… ÿ·» ’—›  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
         .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        Else
         
        .Create Me.hWnd, "   Screen Request Exchange  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Save ..." & Wrap & "Save Data of this  process " & Wrap & " Just click here", True
        End If

    End With

    With TTP
                  If SystemOptions.UserInterface = ArabicInterface Then
        .Create Me.hWnd, "    ‘«‘… ÿ·» ’—›  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        Else
        
        .Create Me.hWnd, "   Screen Request Exchange  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), "Retreat ..." & Wrap & "Retreat Data of this  process " & Wrap & " Just click here", True
        End If

    End With

    With TTP
         If SystemOptions.UserInterface = ArabicInterface Then
        .Create Me.hWnd, "    ‘«‘… ÿ·» ’—›  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
         .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        Else
       
        .Create Me.hWnd, "   Screen Request Exchange  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Delete ..." & Wrap & "Delete Data of this  process " & Wrap & " Just click here", True
        End If
    End With

    With TTP
       If SystemOptions.UserInterface = ArabicInterface Then
        .Create Me.hWnd, "    ‘«‘… ÿ·» ’—›  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
         .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        Else
         
        .Create Me.hWnd, "   Screen Request Exchange  ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Exit ..." & Wrap & "Exit  this  Screen " & Wrap & " Just click here", True
        End If
        
    End With

    With TTP
    If SystemOptions.UserInterface = ArabicInterface Then
        .Create Me.hWnd, "     ‘«‘…  ÿ·» ’—›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
      Else
         .Create Me.hWnd, "    Screen Request Exchange  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "First ..." & Wrap & "To move to the first record" & Wrap & " Just click here", True
      End If
    End With

    With TTP
    If SystemOptions.UserInterface = ArabicInterface Then
        .Create Me.hWnd, "     ‘«‘…  ÿ·» ’—›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        Else
                .Create Me.hWnd, "   Screen Request Exchange  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "Previous ..." & Wrap & "Moving to the previous record" & Wrap & " Just click here", True
        End If
    End With

    With TTP
    If SystemOptions.UserInterface = ArabicInterface Then
        .Create Me.hWnd, "     ‘«‘…  ÿ·» ’—›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
       Else
               .Create Me.hWnd, "    Screen Request Exchange  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "Next ..." & Wrap & "Moving to the next record" & Wrap & "Just click here", True
       End If
    End With

    With TTP
    If SystemOptions.UserInterface = ArabicInterface Then
        .Create Me.hWnd, "     ‘«‘…  ÿ·» ’—›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
      Else
              .Create Me.hWnd, "    Screen Request Exchange ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "Last ..." & Wrap & "Moving to the last record" & Wrap & " Just click here", True
      End If
    End With

    With TTP
    If SystemOptions.UserInterface = ArabicInterface Then
        .Create Me.hWnd, "    ‘«‘…  ÿ·» ’—›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        Else
         .Create Me.hWnd, "    Screen Request Exchange  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "Help ..." & Wrap & "To learn this window" & Wrap & "And how to handle them" & Wrap & "Just click here" & Wrap, True
        End If
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

 
'Private Sub XPDtbTransH_LostFocus()
'If Me.TxtModFlg.text <> "R" Then
'
'      XPDtbTrans.value = ToGregorianDate(XPDtbTransH.value)
'
'End If
'End Sub


Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.text <> "R" Then
TxtSerial1.text = ""
End If

End Sub

