VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmEmpsAdvanceRequest 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·» ”·›… ‰ﬁœÌ…"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12510
   Icon            =   "FrmEmpsAdvanceRequest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   12510
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtSearchCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   1185
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   13920
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   13440
      TabIndex        =   37
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
      TabIndex        =   35
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
      TabIndex        =   30
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtAdvanceValue 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4140
      TabIndex        =   2
      Top             =   1170
      Width           =   1395
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   1
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   12525
      _cx             =   22093
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
      Caption         =   "ÿ·» ”·›… ‰ﬁœÌ… "
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
         ButtonImage     =   "FrmEmpsAdvanceRequest.frx":038A
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
         ButtonImage     =   "FrmEmpsAdvanceRequest.frx":0724
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
         TabIndex        =   6
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
         ButtonImage     =   "FrmEmpsAdvanceRequest.frx":0ABE
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
         TabIndex        =   7
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
         ButtonImage     =   "FrmEmpsAdvanceRequest.frx":0E58
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
         Left            =   6000
         Picture         =   "FrmEmpsAdvanceRequest.frx":11F2
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
         Left            =   2280
         TabIndex        =   36
         Top             =   0
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   7740
      TabIndex        =   8
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   123600897
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DcboEmpName 
      Height          =   315
      Left            =   6720
      TabIndex        =   9
      Top             =   1185
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   1800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8520
      Width           =   8748
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
         TabIndex        =   11
         Top             =   0
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
         TabIndex        =   12
         Top             =   0
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
         TabIndex        =   13
         Top             =   0
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
         TabIndex        =   14
         Top             =   0
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
         TabIndex        =   15
         Top             =   0
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
         TabIndex        =   16
         Top             =   0
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   372
         Index           =   5
         Left            =   2880
         TabIndex        =   29
         Top             =   0
         Width           =   768
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
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   375
         Left            =   855
         TabIndex        =   91
         Top             =   0
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
         Index           =   9
         Left            =   1920
         TabIndex        =   92
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
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
      Height          =   312
      Left            =   8280
      TabIndex        =   17
      Top             =   8040
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
      TabIndex        =   18
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
      TabIndex        =   31
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
   Begin MSDataListLib.DataCombo Dcbranch 
      Bindings        =   "FrmEmpsAdvanceRequest.frx":4E5A
      Height          =   315
      Left            =   2640
      TabIndex        =   33
      Top             =   720
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   6375
      Left            =   0
      TabIndex        =   41
      Top             =   1560
      Width           =   12480
      _cx             =   22013
      _cy             =   11239
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
      Caption         =   "«·»Ì«‰« |Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "FrmEmpsAdvanceRequest.frx":4E6F
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5910
         Left            =   13125
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   45
         Width           =   12390
         _cx             =   21855
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
            TabIndex        =   43
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
            FormatString    =   $"FrmEmpsAdvanceRequest.frx":5209
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
            Left            =   8760
            RightToLeft     =   -1  'True
            TabIndex        =   84
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
            TabIndex        =   44
            Top             =   4560
            Visible         =   0   'False
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   5910
         Index           =   15
         Left            =   45
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   45
         Width           =   12390
         _cx             =   21855
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
         _GridInfo       =   $"FrmEmpsAdvanceRequest.frx":534C
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5880
            Index           =   16
            Left            =   15
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   15
            Width           =   12360
            _cx             =   21802
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
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
               Height          =   4980
               Index           =   0
               Left            =   5745
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   1320
               Width           =   6645
               Begin VB.TextBox TxtDiff 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   120
                  MaxLength       =   10
                  TabIndex        =   109
                  Top             =   3120
                  Width           =   1185
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÌœÊÌ"
                  Height          =   252
                  Index           =   2
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ— ﬁ”ÿ"
                  Height          =   252
                  Index           =   1
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "√Ê· ﬁ”ÿ"
                  Height          =   252
                  Index           =   0
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.TextBox TxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   -840
                  MaxLength       =   10
                  TabIndex        =   104
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.TextBox txtDiscountDES 
                  Alignment       =   1  'Right Justify
                  Height          =   975
                  Left            =   150
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   88
                  Top             =   3480
                  Width           =   3795
               End
               Begin VB.TextBox TxtDiscount 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2520
                  MaxLength       =   10
                  TabIndex        =   85
                  Top             =   3120
                  Width           =   1425
               End
               Begin VB.TextBox TxtPaymentCounts 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   4590
                  TabIndex        =   76
                  Top             =   720
                  Width           =   825
               End
               Begin VB.ComboBox CmbMonth 
                  Height          =   315
                  Left            =   4590
                  TabIndex        =   75
                  Text            =   "CmbMonth"
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.CheckBox ChkSaleryDis 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Œ’„ „‰ «·„— »  ·ﬁ«∆Ì«"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   4560
                  TabIndex        =   74
                  Top             =   2640
                  Value           =   1  'Checked
                  Visible         =   0   'False
                  Width           =   1965
               End
               Begin VB.ComboBox CboYear 
                  Height          =   315
                  Left            =   4590
                  TabIndex        =   73
                  Text            =   "CboYear"
                  Top             =   1800
                  Width           =   1095
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   435
                  Index           =   8
                  Left            =   4560
                  TabIndex        =   72
                  Top             =   2160
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
                  ButtonImage     =   "FrmEmpsAdvanceRequest.frx":5380
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   1965
                  Left            =   90
                  TabIndex        =   77
                  Top             =   1050
                  Width           =   4455
                  _cx             =   7858
                  _cy             =   3466
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
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmEmpsAdvanceRequest.frx":571A
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "›—ﬁ «·ﬂ”Ê—"
                  Height          =   285
                  Index           =   38
                  Left            =   1440
                  TabIndex        =   110
                  Top             =   3120
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   37
                  Left            =   4080
                  TabIndex        =   106
                  Top             =   240
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÊÌ„À·"
                  Height          =   540
                  Index           =   28
                  Left            =   5280
                  TabIndex        =   87
                  Top             =   3720
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÊÌŒ’„ „‰ «·”·› „»·€« Êﬁœ—…"
                  Height          =   555
                  Index           =   26
                  Left            =   4035
                  TabIndex        =   86
                  Top             =   3120
                  Width           =   2520
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·œ›⁄« "
                  Height          =   285
                  Index           =   9
                  Left            =   5550
                  TabIndex        =   82
                  Top             =   780
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «Ê· œ›⁄…"
                  Height          =   285
                  Index           =   10
                  Left            =   4860
                  TabIndex        =   81
                  Top             =   1170
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
                  TabIndex        =   80
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
                  Left            =   5730
                  TabIndex        =   79
                  Top             =   1470
                  Width           =   405
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”‰…"
                  Height          =   315
                  Index           =   12
                  Left            =   5730
                  TabIndex        =   78
                  Top             =   1800
                  Width           =   405
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  „«·Ì…"
               Height          =   1245
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   0
               Width           =   6045
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·ﬁ”«ÿ «·„ »ﬁÌ…"
                  Height          =   315
                  Index           =   41
                  Left            =   1680
                  TabIndex        =   116
                  Top             =   840
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   315
                  Index           =   39
                  Left            =   480
                  TabIndex        =   115
                  Top             =   840
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   405
                  Index           =   44
                  Left            =   3360
                  TabIndex        =   114
                  Top             =   480
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—’Ìœ «ÃÊ— „” Õﬁ…"
                  Height          =   195
                  Index           =   43
                  Left            =   3840
                  TabIndex        =   113
                  Top             =   480
                  Width           =   1965
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì  „Œ’’«  «Ã«“…"
                  Height          =   195
                  Index           =   17
                  Left            =   3960
                  TabIndex        =   70
                  Top             =   240
                  Width           =   1965
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·⁄ﬁœ «·„ »ﬁÌ…"
                  Height          =   285
                  Index           =   18
                  Left            =   1560
                  TabIndex        =   69
                  Top             =   600
                  Width           =   1605
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—’Ìœ «·”·› "
                  Height          =   285
                  Index           =   19
                  Left            =   1800
                  TabIndex        =   68
                  Top             =   240
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   285
                  Index           =   16
                  Left            =   -240
                  TabIndex        =   67
                  Top             =   600
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   20
                  Left            =   960
                  TabIndex        =   66
                  Top             =   600
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   21
                  Left            =   960
                  TabIndex        =   65
                  Top             =   240
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   22
                  Left            =   3240
                  TabIndex        =   64
                  Top             =   240
                  Width           =   525
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·„ÊŸ›"
               Height          =   1320
               Left            =   6165
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   0
               Width           =   6195
               Begin MSDataListLib.DataCombo DcboEmpDepartments 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   55
                  Top             =   720
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DBIssueDate 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   56
                  Top             =   360
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   115605505
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcboJobsType 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   57
                  Top             =   720
                  Width           =   1995
                  _ExtentX        =   3519
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboSpecifications 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   117
                  Top             =   360
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„— »…"
                  Height          =   285
                  Index           =   14
                  Left            =   1200
                  TabIndex        =   118
                  Top             =   360
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—« » «·«”«”Ì"
                  Height          =   285
                  Index           =   5
                  Left            =   4920
                  TabIndex        =   62
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
                  Height          =   285
                  Index           =   13
                  Left            =   3240
                  TabIndex        =   61
                  Top             =   360
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«œ«—…"
                  Height          =   285
                  Index           =   15
                  Left            =   2400
                  TabIndex        =   60
                  Top             =   720
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   23
                  Left            =   3960
                  TabIndex        =   59
                  Top             =   360
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊŸÌ›…"
                  Height          =   285
                  Index           =   24
                  Left            =   5280
                  TabIndex        =   58
                  Top             =   720
                  Width           =   645
               End
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   285
               Left            =   120
               TabIndex        =   83
               Top             =   5595
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
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
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   1860
               Left            =   0
               TabIndex        =   89
               Top             =   2085
               Width           =   5565
               _cx             =   9816
               _cy             =   3281
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
               Rows            =   3
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmEmpsAdvanceRequest.frx":57CB
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
            Begin MSDataListLib.DataCombo DcboEmpName2 
               Height          =   315
               Left            =   120
               TabIndex        =   101
               Top             =   1320
               Width           =   4665
               _ExtentX        =   8229
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„œÌ— «·„»«‘—"
               Height          =   270
               Index           =   35
               Left            =   4575
               TabIndex        =   103
               Top             =   1320
               Width           =   1140
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   285
               Index           =   33
               Left            =   4725
               TabIndex        =   100
               Top             =   15
               Width           =   1020
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·÷«„‰Ì‰"
               ForeColor       =   &H00FF0000&
               Height          =   510
               Index           =   31
               Left            =   1920
               TabIndex        =   90
               Top             =   1770
               Width           =   1635
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3480
               Index           =   62
               Left            =   2370
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   1560
               Width           =   540
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5880
            Index           =   9
            Left            =   15
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   15
            Width           =   12360
            _cx             =   21802
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
               Height          =   4425
               Left            =   3255
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   1275
               Width           =   660
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   3075
               Left            =   4095
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   1560
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3075
               Index           =   67
               Left            =   2310
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   1560
               Width           =   600
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   2940
               Index           =   68
               Left            =   3915
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   1995
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
               Height          =   3540
               Index           =   69
               Left            =   2910
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   1560
               Width           =   345
            End
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ê«›ﬁ… «·«œ«—…"
      Height          =   1470
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   93
      Top             =   5640
      Width           =   6120
      Begin VB.OptionButton opt_Notok 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "€Ì— „Ê«›ﬁ"
         Height          =   252
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt_ok 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Ê«›ﬁ"
         Height          =   252
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   98
         Top             =   240
         Width           =   732
      End
      Begin VB.TextBox txtReason 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   96
         Top             =   600
         Width           =   4632
      End
      Begin MSDataListLib.DataCombo DcboJobsType2 
         Height          =   312
         Left            =   2760
         TabIndex        =   94
         Top             =   240
         Width           =   1992
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”»» «·—›÷"
         Height          =   540
         Index           =   32
         Left            =   4440
         TabIndex        =   97
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„”„Ï «·ÊŸÌ›Ï"
         Height          =   285
         Index           =   36
         Left            =   4800
         TabIndex        =   95
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   0
      TabIndex        =   119
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   115605505
      CurrentDate     =   38784
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Ã„«·Ì «·„” Õﬁ«  ··„ÊŸ›"
      Height          =   195
      Index           =   42
      Left            =   0
      TabIndex        =   112
      Top             =   0
      Width           =   1965
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‘Â—"
      Height          =   315
      Index           =   40
      Left            =   0
      TabIndex        =   111
      Top             =   0
      Width           =   1605
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   285
      Index           =   34
      Left            =   0
      TabIndex        =   102
      Top             =   0
      Width           =   1005
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
      TabIndex        =   40
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
      Caption         =   "«·›—⁄"
      Height          =   255
      Index           =   29
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   780
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   12810
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   285
      Index           =   4
      Left            =   11310
      TabIndex        =   28
      Top             =   750
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   285
      Index           =   3
      Left            =   11430
      TabIndex        =   27
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… «·”·›…"
      Height          =   285
      Index           =   2
      Left            =   5430
      TabIndex        =   26
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   8670
      TabIndex        =   25
      Top             =   735
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   276
      Index           =   8
      Left            =   10920
      TabIndex        =   24
      Top             =   8040
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   312
      Index           =   7
      Left            =   2520
      TabIndex        =   23
      Top             =   8040
      Width           =   1068
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   312
      Index           =   6
      Left            =   600
      TabIndex        =   22
      Top             =   8040
      Width           =   972
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   21
      Top             =   7980
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   20
      Top             =   7980
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   13350
      TabIndex        =   19
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmEmpsAdvanceRequest"
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
Dim ScreenNameArabic As String
Dim ScreenNameEnglish As String
Public LongRow As Long

Function CheAdvanced(Optional advanceID As Double = 0) As Boolean
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
CheAdvanced = False
sql = "select AdvanceID from TblEmpAdvanceRequest where AdvanceID=" & advanceID & " and AccAproved=1"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CheAdvanced = True
Else
CheAdvanced = False
End If
End Function
Private Sub Accredit_Click()
    Dim BeginTrans As Boolean

 If val(XPTxtID.text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
      
      
If XPTxtID.text = "" Then Exit Sub
    Cn.BeginTrans
    BeginTrans = True

  '  If IsNull(rs("Posted")) Then
  '      rs("Posted") = user_id
  '      rs("PostedDate") = Time
  '  Else
  '      rs("Posted") = Null
  '     rs("PostedDate") = Time
  '  End If
  '
  '  rs.update
    
        SendTopost Me.Name, "TblEmpAdvanceRequest", "AdvanceID", val(DcboEmpDepartments.BoundText), val(Dcbranch.BoundText), val(XPTxtID.text), XPTxtID.text
    
   rs.Resync
   
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

    Cn.CommitTrans
    BeginTrans = False
'FillApprovedTable
    Retrive (val(Me.XPTxtID.text))
End Sub

Private Sub Cmd_Click(index As Integer)
Dim i As Integer
Dim Msg As String
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
            
              GRID2.Clear flexClearScrollable, flexClearEverything
    GRID2.rows = 1
               VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.rows = 2
            Me.DCboUserName.BoundText = user_id
            TxtPaymentCounts.text = 1
Dcbranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            
            Accredit.Enabled = True
                If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               
        Case 1
                    If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If
            
If CheAdvanced(val(Me.XPTxtID.text)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–Â «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ „œ›Ê⁄«  "
                    Else
                    Msg = " Can Not Edit this Process"
                    Msg = Msg & CHR(13) & " There is the Process of Payments "
                    
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
                
            If DoPremis(Do_Edit, Me.Name, True) = False Then
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
            VSFlexGrid1.rows = VSFlexGrid1.rows + 1
CuurentLogdata
        Case 2
    
            With FG
For i = 1 To .rows - 1
If .TextMatrix(i, .ColIndex("PartDate")) <> "" Then
DTPicker1.value = .TextMatrix(i, .ColIndex("PartDate"))
      If ChekPayedSalary(year(DTPicker1.value), Month(DTPicker1.value), val(Me.Dcbranch.BoundText)) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ Õ–› ﬁÌœ «·—Ê« »  ··‘Â— «·„Õœœ «Ê·«" & .TextMatrix(i, .ColIndex("PartDate"))
            Else
            MsgBox "Delete Salary Allocation JL" & .TextMatrix(i, .ColIndex("PartDate"))
            End If
            Exit Sub
            End If
   End If
    Next i
End With

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
            
      With FG
      For i = FG.FixedRows To FG.rows - 1
         If Opt(0).value = True And i = 1 Then
            FG.TextMatrix(i, FG.ColIndex("PartValue")) = val(FG.TextMatrix(i, FG.ColIndex("PartValue"))) + (val(TxtAdvanceValue.text) - val(TxtValue.text))
            End If
             If Opt(1).value = True And i = (FG.rows - 1) Then
            
            FG.TextMatrix(i, FG.ColIndex("PartValue")) = val(FG.TextMatrix(i, FG.ColIndex("PartValue"))) + (val(TxtAdvanceValue.text) - val(TxtValue.text))
            End If
            
        Next i
      End With
      Reline
With FG
If val(.TextMatrix(1, .ColIndex("PartNO"))) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  Ê“Ì⁄ «·”·›…"
Else
MsgBox "Please Advance Distribution"
End If
Exit Sub
End If
End With
If Round(val(TxtValue.text), 2) <> Round(val(TxtAdvanceValue.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "ÌÃ» «‰  ﬂÊ‰ ﬁÌ„… «·”·›…  ”«ÊÌ «Ã„«·Ì «·œ›⁄« "
Else
MsgBox "It must be advance value equal to the total amount of  payments"
End If
Exit Sub
End If
            my_branch = Me.Dcbranch.BoundText

            SaveData
CuurentLogdata
        Case 3
            Undo

        Case 4
        
                    If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If
            
    If CheAdvanced(val(Me.XPTxtID.text)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–Â «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ „œ›Ê⁄«  "
                    Else
                    Msg = " Can Not Delete this Process"
                    Msg = Msg & CHR(13) & " There is the Process of Payments "
                    
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
                
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
        General_Search.send_form = "advreq"
               Load General_Search
             General_Search.show

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 8
        If Opt(0).value = False And Opt(1).value = False And Opt(2).value = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
        Else
        MsgBox "Please Select Method Number of decimal"
        End If
        Exit Sub
        End If
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


'MySQL = " SELECT     dbo.TblEmpAdvanceRequest.AdvanceID AS AdvanceIDH, dbo.TblEmpAdvanceRequest.AdvanceValue, dbo.TblEmpAdvanceRequest.PaymentCounts, "
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequest.FirstDate, dbo.TblEmpAdvanceRequest.AdvanceDate, dbo.TblEmpAdvanceRequest.basicSalary, dbo.TblEmpAdvanceRequest.discount,"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequest.DiscountDES, dbo.TblEmpAdvanceRequest.EmpDue, dbo.TblEmpAdvanceRequest.Contractvalid,"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequest.oldAdvance, dbo.TblEmpAdvanceRequest.Posted, dbo.TblEmpAdvanceRequest.PostedDate, dbo.TblEmpAdvanceRequest.NoteSerial,"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequest.Approved, dbo.TblEmpAdvanceRequest.Transaction_ID, dbo.TblEmpAdvanceRequest.FirstMonthPayment,"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequest.FirstYearPayment, dbo.TblEmpAdvanceRequest.AutoDiscount, dbo.TblEmpAdvanceRequest.Emp_id,"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequest.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmpAdvanceRequest.gradeID,"
'MySQL = MySQL & "                      dbo.TblEmpGrades.name, dbo.TblEmpGrades.namee, dbo.TblEmpAdvanceRequest.jobID_approve, dbo.TblEmpAdvanceRequest.ok,"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequest.notok, dbo.TblEmpAdvanceRequest.reason, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequest.DeparmentID, TblEmpDepartments_2.DepartmentName, TblEmpDepartments_2.DepartmentNamee,"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequest.JobTypeID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmpAdvanceRequest.ManagerID,"
'MySQL = MySQL & "                      TblEmployee_1.Emp_Name AS mangerEmp_Name, TblEmployee_1.Emp_Namee AS mangerEmp_NameE, TblEmployee_1.Fullcode AS MangerFullcode,"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequestDetails2.LongContarct, dbo.TblEmpAdvanceRequestDetails2.salary, dbo.TblEmpAdvanceRequestDetails2.EmpID,"
'MySQL = MySQL & "                      TblEmployee_2.Emp_Name AS DetEmp_Name, TblEmployee_2.Fullcode AS DtFullcode, TblEmployee_2.Emp_Namee AS DetEmp_NameE,"
'MySQL = MySQL & "                      TblEmployee_2.DepartmentID, TblEmpDepartments_1.DepartmentName AS DetDepartmentName, TblEmpDepartments_1.DepartmentNamee AS DetDepartmentNameE,"
'MySQL = MySQL & "                       dbo.TblEmployee.NumEkama, dbo.TblEmployee.NumPoket, TblEmployee_2.JobTypeID AS DetJobTypeID, TblEmpJobsTypes_1.JobTypeName AS DetJobTypeName,"
'MySQL = MySQL & "                      TblEmpJobsTypes_1.JobTypeNamee AS DetJobTypeNameE, TblEmployee_2.BignDateWork"
'MySQL = MySQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequest LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequestDetails2 LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmpDepartments TblEmpDepartments_1 RIGHT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_2 LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_1 ON TblEmployee_2.JobTypeID = TblEmpJobsTypes_1.JobTypeID ON"
'MySQL = MySQL & "                      TblEmpDepartments_1.DeparmentID = TblEmployee_2.DepartmentID ON dbo.TblEmpAdvanceRequestDetails2.EmpID = TblEmployee_2.Emp_ID ON"
'MySQL = MySQL & "                      dbo.TblEmpAdvanceRequest.AdvanceID = dbo.TblEmpAdvanceRequestDetails2.AdvanceID ON"
'MySQL = MySQL & "                      dbo.TblEmployee.Emp_ID = dbo.TblEmpAdvanceRequest.Emp_id LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblEmpAdvanceRequest.ManagerID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TblEmpAdvanceRequest.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmpDepartments TblEmpDepartments_2 ON dbo.TblEmpAdvanceRequest.DeparmentID = TblEmpDepartments_2.DeparmentID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmpGrades ON dbo.TblEmpAdvanceRequest.gradeID = dbo.TblEmpGrades.gradeid LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblEmpAdvanceRequest.Branch_NO = dbo.TblBranchesData.branch_id"

MySQL = "SELECT  dbo.TblEmpAdvanceRequest.Balance,     dbo.TblEmpAdvanceRequest.AdvanceID AS AdvanceIDH, dbo.TblEmpAdvanceRequest.AdvanceValue, dbo.TblEmpAdvanceRequest.PaymentCounts, "
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequest.FirstDate, dbo.TblEmpAdvanceRequest.AdvanceDate, dbo.TblEmpAdvanceRequest.basicSalary, dbo.TblEmpAdvanceRequest.discount,"
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequest.DiscountDES, dbo.TblEmpAdvanceRequest.EmpDue, dbo.TblEmpAdvanceRequest.Contractvalid,"
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequest.oldAdvance, dbo.TblEmpAdvanceRequest.Posted, dbo.TblEmpAdvanceRequest.PostedDate, dbo.TblEmpAdvanceRequest.NoteSerial,"
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequest.Approved, dbo.TblEmpAdvanceRequest.Transaction_ID, dbo.TblEmpAdvanceRequest.FirstMonthPayment,"
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequest.FirstYearPayment, dbo.TblEmpAdvanceRequest.AutoDiscount, dbo.TblEmpAdvanceRequest.Emp_id,"
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequest.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmpAdvanceRequest.gradeID,"
MySQL = MySQL & "                                           dbo.TblEmpGrades.name, dbo.TblEmpGrades.namee, dbo.TblEmpAdvanceRequest.jobID_approve, dbo.TblEmpAdvanceRequest.ok,"
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequest.notok, dbo.TblEmpAdvanceRequest.reason, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
MySQL = MySQL & "                                           dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
MySQL = MySQL & "                                           dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequest.DeparmentID, TblEmpDepartments_2.DepartmentName, TblEmpDepartments_2.DepartmentNamee,"
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequest.JobTypeID, TblEmpJobsTypes_1.JobTypeName, TblEmpJobsTypes_1.JobTypeNamee, dbo.TblEmpAdvanceRequest.ManagerID,"
MySQL = MySQL & "                                           TblEmployee_1.Emp_Name AS mangerEmp_Name, TblEmployee_1.Emp_Namee AS mangerEmp_NameE, TblEmployee_1.Fullcode AS MangerFullcode,"
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequestDetails2.LongContarct, dbo.TblEmpAdvanceRequestDetails2.salary, dbo.TblEmpAdvanceRequestDetails2.EmpID,"
MySQL = MySQL & "                                           TblEmployee_2.Emp_Name AS DetEmp_Name, TblEmployee_2.Fullcode AS DtFullcode, TblEmployee_2.Emp_Namee AS DetEmp_NameE,"
MySQL = MySQL & "                                           TblEmployee_2.DepartmentID, TblEmpDepartments_1.DepartmentName AS DetDepartmentName, TblEmpDepartments_1.DepartmentNamee AS DetDepartmentNameE,"
MySQL = MySQL & "                                            dbo.TblEmployee.NumEkama, dbo.TblEmployee.NumPoket, TblEmployee_2.JobTypeID AS DetJobTypeID, TblEmpJobsTypes_1.JobTypeName AS DetJobTypeName,"
MySQL = MySQL & "                                           TblEmpJobsTypes_1.JobTypeNamee AS DetJobTypeNameE, TblEmployee_2.BignDateWork, dbo.TblUsers.UserName,"


MySQL = MySQL & " sign0 = " & GetUserSign(val(XPTxtID.text), Me.Name, 0)
MySQL = MySQL & " ,sign1 = " & GetUserSign(val(XPTxtID.text), Me.Name, 1)
MySQL = MySQL & " ,sign2 = " & GetUserSign(val(XPTxtID.text), Me.Name, 2)
MySQL = MySQL & " ,sign3 = " & GetUserSign(val(XPTxtID.text), Me.Name, 3)
MySQL = MySQL & " ,sign4 = " & GetUserSign(val(XPTxtID.text), Me.Name, 4)

MySQL = MySQL & "                     FROM         dbo.TblEmpAdvanceRequestDetails2 LEFT OUTER JOIN"
MySQL = MySQL & "                                           dbo.TblEmpDepartments TblEmpDepartments_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                                           dbo.TblEmployee TblEmployee_2 LEFT OUTER JOIN"
MySQL = MySQL & "                                           dbo.TblEmpJobsTypes TblEmpJobsTypes_1 ON TblEmployee_2.JobTypeID = TblEmpJobsTypes_1.JobTypeID ON"
MySQL = MySQL & "                                           TblEmpDepartments_1.DeparmentID = TblEmployee_2.DepartmentID ON dbo.TblEmpAdvanceRequestDetails2.EmpID = TblEmployee_2.Emp_ID RIGHT OUTER JOIN"
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequest INNER JOIN"
MySQL = MySQL & "                                           dbo.TblUsers ON dbo.TblEmpAdvanceRequest.UserID = dbo.TblUsers.UserID ON"
MySQL = MySQL & "                                           dbo.TblEmpAdvanceRequestDetails2.AdvanceID = dbo.TblEmpAdvanceRequest.AdvanceID LEFT OUTER JOIN"
MySQL = MySQL & "                                           dbo.TblEmployee ON dbo.TblEmpAdvanceRequest.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                                           dbo.TblEmployee TblEmployee_1 ON dbo.TblEmpAdvanceRequest.ManagerID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                                           dbo.TblEmpJobsTypes TblEmpJobsTypes_2 ON dbo.TblEmpAdvanceRequest.JobTypeID = TblEmpJobsTypes_2.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                                           dbo.TblEmpDepartments TblEmpDepartments_2 ON dbo.TblEmpAdvanceRequest.DeparmentID = TblEmpDepartments_2.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "                                           dbo.TblEmpGrades ON dbo.TblEmpAdvanceRequest.gradeID = dbo.TblEmpGrades.gradeid LEFT OUTER JOIN"
MySQL = MySQL & "                                           dbo.TblBranchesData ON dbo.TblEmpAdvanceRequest.Branch_NO = dbo.TblBranchesData.branch_id"
                      

MySQL = MySQL & "  Where (dbo.TblEmpAdvanceRequest.AdvanceID = " & val(XPTxtID.text) & ")"
 
        If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\AdvanceRequest.rpt"
             
        Else
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\AdvanceRequest.rpt"
              
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
        xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
         xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(FG.TextMatrix(Me.FG.FixedRows, FG.ColIndex("PartValue"))), val(FG.TextMatrix(Me.FG.FixedRows, FG.ColIndex("PartValue"))), 0)
 xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
   xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
           ''//////
   Dim xLogo As CRAXDRT.OLEObject
   Dim SqlT As String
   Dim i As Integer
   Dim EmpIDD As Long
   Dim xWidth As Integer
   Dim Rs4 As ADODB.Recordset
   Set Rs4 = New ADODB.Recordset
  SqlT = " SELECT        TOP (100) PERCENT dbo.TblUsers.Empid"
  SqlT = SqlT + "    FROM            dbo.ApprovalData INNER JOIN"
  SqlT = SqlT + "                      dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
  SqlT = SqlT + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (NOT (ApprovDate IS NULL)) AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
  SqlT = SqlT & " ORDER BY levelorder"
  Rs4.Open SqlT, Cn, adOpenStatic, adLockOptimistic, adCmdText
  xWidth = 300
  For i = 1 To Rs4.RecordCount
  EmpIDD = IIf(IsNull(Rs4("Empid").value), 0, Rs4("Empid").value)
            If Dir(App.path & "\" & SystemOptions.ImagesPath & "\sign" & EmpIDD & ".JPG") <> "" Then
    
            Set xLogo = xReport.Areas(4).Sections(1).AddPictureObject(App.path & "\" & SystemOptions.ImagesPath & "\sign" & EmpIDD & ".JPG", xWidth, 3000)
            xLogo.Width = 800
            xLogo.Height = 800
            xLogo.backcolor = vbWhite
            xLogo.BorderColor = 255
            xLogo.CloseAtPageBreak = True
           xWidth = xWidth + 1000
          End If
          Rs4.MoveNext
    Next i
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub FG_AfterEdit(ByVal row As Long, ByVal Col As Long)
Reline
End Sub

Private Sub FG_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With FG
Select Case .ColKey(Col)
Case "PartNO"
Cancel = True
Case "PartDate"
Cancel = True
Case "PartValue"
'Fg.ColComboList = ""
End Select
End With
End Sub
Sub Reline()

    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.FG
        For i = .FixedRows To .rows - 1
                If .TextMatrix(i, .ColIndex("PartDate")) <> "" Then
           Sm = Sm + val(.TextMatrix(i, .ColIndex("PartValue")))
           End If
           Next i
  
    End With
    TxtValue.text = Sm
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub opt_Notok_Click()
If opt_Notok.value = True Then
    lbl(32).Visible = True
    txtReason.Visible = True
End If

End Sub

Private Sub opt_ok_Click()
If opt_ok.value = True Then
    lbl(32).Visible = False
    txtReason.Visible = False
End If
End Sub

Private Sub TxtAdvanceValue_Change()
If Me.TxtModFlg.text = "E" Then
FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = FG.FixedRows
    End If
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
        FrmEmployeeSearch.lbltype = 3
       Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If

End Sub

Private Sub DcboEmpName_Click(Area As Integer)
  On Error Resume Next
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
        Dim StrTempAccountCode As String
     Dim AccountCode As String
        Dim endContractPerMonth As Double
        Dim balanceString As String
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth
        
          WriteCustomerBalPublic Account_code2, Balance
          
  lbl(22).Caption = val(Balance) '„Œ’’ «Ã«“…

          WriteCustomerBalPublic Account_code, Balance
          
  lbl(21).Caption = val(Balance) '–„„ „ÊŸ›Ì‰
  lbl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
        DBIssueDate.value = IssueDate
        DcboEmpDepartments.BoundText = DepID
        DcboSpecifications.BoundText = gradeID
        DcboJobsType.BoundText = JobTypeID
        lbl(23).Caption = GetSalaryEmployee(val(Me.DcboEmpName.BoundText))
      '  lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "1,13,16")
              StrTempAccountCode = get_EMPLOYEE_Account(val(DcboEmpName.BoundText), "Account_Code1")    'C?C??? C???E??E
      WriteCustomerBalPublic StrTempAccountCode, Balance, balanceString
     lbl(44).Caption = Balance ' —Ê« » „‘ Õﬁ…
     
    'End If

End Sub



Private Sub TxtValue_Change()
If Me.TxtModFlg.text <> "R" Then
txtDiff.text = val(TxtAdvanceValue.text) - val(TxtValue.text)
End If
End Sub

Public Sub VSFlexGrid1_AfterEdit(ByVal row As Long, ByVal Col As Long)
 
 Dim endContractPerMonth As Double
Dim StrAccountCode As String
Dim LngRow As Long
Dim StrSQL As String
Dim Rs1 As ADODB.Recordset
 With VSFlexGrid1

        Select Case .ColKey(Col)
 
            Case "name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(row, .ColIndex("id")) = StrAccountCode
                 .TextMatrix(row, .ColIndex("salary")) = GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(row, .ColIndex("id"))), "")
                 
                   get_employee_information val(.TextMatrix(row, .ColIndex("id"))), , , , , , , , endContractPerMonth
                   .TextMatrix(row, .ColIndex("LongContarct")) = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
                   StrSQL = "select * from TblEmployee where Emp_ID=" & val(StrAccountCode) & " "
                   Set Rs1 = New ADODB.Recordset
                Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If Rs1.RecordCount > 0 Then
                 .TextMatrix(row, .ColIndex("code")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                 End If
                  Case "code"
                    StrSQL = "select * from TblEmployee where Fullcode='" & .TextMatrix(row, .ColIndex("code")) & "' "
                   Set Rs1 = New ADODB.Recordset
                Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If Rs1.RecordCount > 0 Then
                 .TextMatrix(row, .ColIndex("id")) = IIf(IsNull(Rs1("Emp_ID").value), "", Rs1("Emp_ID").value)
                 If SystemOptions.UserInterface = ArabicInterface Then
                 
                 .TextMatrix(row, .ColIndex("name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                 Else
                 .TextMatrix(row, .ColIndex("name")) = IIf(IsNull(Rs1("Emp_Namee").value), "", Rs1("Emp_Namee").value)
                 End If
                 .TextMatrix(row, .ColIndex("salary")) = GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(row, .ColIndex("id"))), "")
                 get_employee_information val(.TextMatrix(row, .ColIndex("id"))), , , , , , , , endContractPerMonth
                   .TextMatrix(row, .ColIndex("LongContarct")) = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
                 End If
                  
                End Select
                 If row = .rows - 1 Then
    
            .rows = .rows + 1
        End If
                End With
          ReLineGrid
          


          
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
 With VSFlexGrid1

      
        Select Case .ColKey(Col)
            
            Case "salary"
               Cancel = True
            Case "LongContarct"
             Cancel = True
        End Select
        End With
End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.text <> "R" Then
    With Me.VSFlexGrid1

        Select Case .ColKey(.Col)

                 Case "code", "name"
              
                  LongRow = .row


   If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 27
        Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If
    
                               End Select
             End With
        End If
        
        
End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
Dim StrSQL As String
Dim Rs1 As ADODB.Recordset
Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

  With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "name"
            Set Rs1 = New ADODB.Recordset
                StrSQL = "select * from TblEmployee "
                Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = VSFlexGrid1.BuildComboList(Rs1, "Emp_Name", "Emp_ID")
                Else
                StrComboList = VSFlexGrid1.BuildComboList(Rs1, "Emp_Namee", "Emp_ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With
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

'Log Data
ScreenNameArabic = "ÿ·» ”·›… ‰ﬁœÌ…"
ScreenNameEnglish = "Cash Advance Request"
   
   RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"


    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

    With Me.FG
        .RowHeightMin = 300
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

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
       If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.DcboEmpName
    
   Dcombos.GetEmployees Me.DcboEmpName
   Dcombos.GetEmployees Me.DcboEmpName2
     
    Dcombos.GetBranches Me.Dcbranch

    Dcombos.GetEmpDepartments Me.DcboEmpDepartments
    Dcombos.GetEmpJobsTypes Me.DcboJobsType
    
        Dcombos.GetEmpJobsTypes Me.DcboJobsType2

    Dcombos.GetEmpGrades Me.DcboSpecifications
    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
    YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEmpAdvanceRequest     Order By AdvanceID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.text = "R"
    Retrive


 

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
    Exit Sub

ErrTrap:
End Sub


Function CuurentLogdata(Optional Currentmode As String)
  LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "—ﬁ„ «·ÿ·» " & XPTxtID.text & CHR(13) & "   «· «—ÌŒ   " & XPDtbTrans.value & CHR(13) & "  «·›—⁄  " & Dcbranch.text & CHR(13) & "   «”„ «·„ÊŸ›   " & DcboEmpName.text & CHR(13) & CHR(13) & "      ﬁÌ„… «·”·›…   " & TxtAdvanceValue.text & CHR(13) & " ⁄œœ «·œ›⁄«   " & TxtPaymentCounts.text
  LogTextA = LogTextA & CHR(13) & "  «·—« » «·«”«”Ï  " & val(lbl(23).Caption) & CHR(13) & CHR(13) & "   «—ÌŒ «· ⁄ÌÌ‰  " & DBIssueDate.value & CHR(13) & CHR(13) & "  «·ÊŸÌ›…  " & DcboJobsType.text & CHR(13) & "  «·«œ«—…  " & DcboEmpDepartments.text & CHR(13) & "  «·„— »…   " & DcboSpecifications.text & CHR(13) & "  ”·› ·„  ”œœ   " & val(lbl(21).Caption) & CHR(13) & "  «Ã„«·Ì «·„” Õﬁ«  ··„ÊŸ›  " & lbl(22).Caption & CHR(13) & "  „œ… «·⁄ﬁœ «·„ »ﬁÌ…  " & lbl(20).Caption
  LogTextA = LogTextA & CHR(13) & "   «·„œÌ— «·„»«‘—   " & DcboEmpName2.text & CHR(13)
  Dim i As Integer
  
For i = VSFlexGrid1.FixedRows To VSFlexGrid1.rows - 1
    LogTextA = LogTextA & "   «·÷«„‰ÌÌ‰   " & CHR(13)
    If VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("name")) <> "" Then
            LogTextA = LogTextA & CHR(13) & "   «·«”„   " & VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("name")) & CHR(13) & "    «·—« »   " & val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("salary"))) & CHR(13) & "   „œ… «·⁄ﬁœ «·„ »ﬁÌ…  " & VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("LongContarct")) & CHR(13)
    End If
Next
  
 LogTextA = LogTextA & "  „Ê«›ﬁ… «·«œ«—…   " & CHR(13)
 LogTextA = LogTextA & "  «·„”„Ï «·ÊŸÌ›ÌÏ  " & DcboJobsType2.text & CHR(13) & " „Ê«›ﬁ " & opt_ok.value & CHR(13) & "  €Ì— „Ê«›ﬁ  " & opt_Notok.value & CHR(13) & "   ”»» «·—›÷   " & txtReason.text
 LogTextA = LogTextA & CHR(13) & " ÿ—Ìﬁ… «·”œ«œ " & CHR(13)
 
 For i = FG.FixedRows To FG.rows - 1
     If val(FG.TextMatrix(i, FG.ColIndex("PartValue"))) <> 0 Then
            LogTextA = LogTextA & CHR(13) & "  —ﬁ„ «·œ›⁄…  " & FG.TextMatrix(i, FG.ColIndex("PartNO")) & CHR(13) & "   ﬁÌ„… «·œ›⁄…  " & val(FG.TextMatrix(i, FG.ColIndex("PartValue"))) & CHR(13) & "    «—ÌŒ «·”œ«œ " & FG.TextMatrix(i, FG.ColIndex("PartDate"))
    End If
Next
 LogTextA = LogTextA & CHR(13) & "   ÊÌŒ’„ „‰ «·”·› „»·€ Êﬁœ—…   " & TxtDiscount.text & CHR(13) & "   ÊÌ„À·  " & txtDiscountDES.text & CHR(13) & "   Õ—— »Ê«”ÿ…  " & DCboUserName.text
  
  
  
  
LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "  Request No.  " & XPTxtID & CHR(13) & " Date   " & XPDtbTrans.value & CHR(13) & "  Employee Name  " & DcboEmpName & CHR(13) & "      Value    " & TxtAdvanceValue & CHR(13) & "  Count    " & TxtPaymentCounts
LogTexte = LogTexte & CHR(13) & "   Basic Salary  " & val(lbl(23).Caption) & CHR(13) & CHR(13) & "  Begin Work Date  " & DBIssueDate.value & CHR(13) & CHR(13) & "  Job  " & DcboJobsType.text & CHR(13) & "  Department  " & DcboEmpDepartments.text & CHR(13) & "  class   " & DcboSpecifications.text & CHR(13) & "  Advances have not been paid   " & val(lbl(21).Caption) & CHR(13) & "  Total Dues  " & lbl(22).Caption & CHR(13) & "  Remain Period in Contract  " & lbl(20).Caption
LogTexte = LogTexte & CHR(13) & "   Direct Manager    " & DcboEmpName2.text & CHR(13)
  
  
For i = VSFlexGrid1.FixedRows To VSFlexGrid1.rows - 1
    LogTexte = LogTexte & "   Guarantors   " & CHR(13)
    If VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("name")) <> "" Then
            LogTexte = LogTexte & CHR(13) & "   Name   " & VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("name")) & CHR(13) & "    Salary   " & val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("salary"))) & CHR(13) & "   Remain Period In Contract  " & VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("LongContarct")) & CHR(13)
    End If
Next
  
 LogTexte = LogTexte & "  Managment Approve  " & CHR(13)
 LogTexte = LogTexte & "  job Title  " & DcboJobsType2.text & CHR(13) & "  Approve  " & opt_ok.value & CHR(13) & "  Not Approved   " & opt_Notok.value & CHR(13) & "   Refuse Reason     " & txtReason.text
 LogTexte = LogTexte & "  Payment Way   " & CHR(13)
 
 For i = FG.FixedRows To FG.rows - 1
     If val(FG.TextMatrix(i, FG.ColIndex("PartValue"))) <> 0 Then
            LogTexte = LogTexte & CHR(13) & "  Payment No.   " & FG.TextMatrix(i, FG.ColIndex("PartNO")) & CHR(13) & "   Payment Value    " & val(FG.TextMatrix(i, FG.ColIndex("PartValue"))) & CHR(13) & "   Payment Date  " & FG.TextMatrix(i, FG.ColIndex("PartDate"))
    End If
Next
 LogTexte = LogTexte & "   Discount From Advance Value    " & TxtDiscount.text & CHR(13) & "   Represent   " & txtDiscountDES.text & CHR(13) & "   Edit By  " & DCboUserName.text
  
  
  
   
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , val(TxtNoteSerial)
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , val(TxtNoteSerial)
    End If
    
End Function





Private Sub ChangeLang()
lbl(41).Caption = "No Payment"
    
    
    lbl(35).Caption = "Direct Manager"
    Frame3.Caption = "Managment Approve "
    opt_ok.Caption = "OK"
    opt_Notok.Caption = "Not Ok"
    lbl(32).Caption = "Refuse Reason"
    lbl(36).Caption = "Job Title"
    lbl(43).Caption = "Balance"
    lbl(38).Caption = "Diff"
    Opt(0).RightToLeft = False
    Opt(1).RightToLeft = False
    Opt(2).RightToLeft = False
    Opt(0).Caption = "Frist"
    Opt(1).Caption = "Last"
    Opt(2).Caption = "Manual"
    Accredit.Caption = "Send to Approv."
    lbl(37).Caption = "Method Number Decimal"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Label1.Visible = False
Cmd(9).Caption = "Print"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(8).Caption = "Calculate Date"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
lbl(31).Caption = "Guarantors"
    Me.Caption = " Advance Request"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(29).Caption = "Branch"
    lbl(3).Caption = "Employee"
    lbl(2).Caption = "Value"
    Frame1.Caption = "Data of Employee"
    lbl(5).Caption = "Salary"
    lbl(13).Caption = "Date  Appoin"
    lbl(24).Caption = "Position"
    lbl(15).Caption = "Mange"
    Frame2.Caption = "Data Financial"
    lbl(14).Caption = "Grade"
    lbl(19).Caption = "Advances not paid"
    lbl(18).Caption = "Remaining Duration Contract"
    lbl(17).Caption = "Total Emp Benefits"
    lbl(16).Caption = "Month"
  '  lbl(0).Caption = "Box"
    Fra(0).Caption = "Payments Method"
    lbl(28).Caption = "Represents"
    lbl(9).Caption = "Count"
    lbl(10).Caption = "Start"
    lbl(11).Caption = "Month"
    lbl(12).Caption = "Year"
    Cmd(8).Caption = "Calc Dates"
    ChkSaleryDis.Caption = "Auto Discount"
    lbl(26).Caption = "Deducted from the amount of advances"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    XPTab301.Caption = "Data|Accreditation status "
    lbl(6).Caption = "rec. count"
Label11.Caption = "Approval is Required"
    With Me.FG
        .TextMatrix(0, .ColIndex("PartNO")) = "NO"
        .TextMatrix(0, .ColIndex("PartValue")) = "Value"
        .TextMatrix(0, .ColIndex("PartDate")) = "Date"
        .TextMatrix(0, .ColIndex("Remark")) = "Status"

    End With
       With Me.GRID2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "LevelName"
        .TextMatrix(0, .ColIndex("EmpName")) = "EmpName"
.TextMatrix(0, .ColIndex("ApprovDate")) = "ApprovDate"
.TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
    End With
         With Me.VSFlexGrid1
         
         .TextMatrix(0, .ColIndex("code")) = "Code"
        .TextMatrix(0, .ColIndex("LineNo")) = "Serial"
        .TextMatrix(0, .ColIndex("name")) = "Name"
        .TextMatrix(0, .ColIndex("salary")) = "Salary"
        .TextMatrix(0, .ColIndex("LongContarct")) = "Remaining Duration Contract"

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
    
   RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
    
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtAdvanceValue_LostFocus()
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

        If val(TxtAdvanceValue.text) >= Mytot Then
            MsgBox "⁄›Ê« «·”·›…  ⁄œ  «·Õœ  «·„”„ÊÕ »Â ÊÂÊ 5 «÷⁄«› ﬁÌ„Â «·—« »  " & CHR(13) & "   —« » «·„ÊŸ›    " & MySal, vbOKOnly, App.Title
            Exit Sub
   
        End If
  
    End If
   
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = "ÿ·» ”·›… ‰ﬁœÌ…"
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
            TxtAdvanceValue.locked = True
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
            '        Me.Caption = "ÿ·» ”·›… ‰ﬁœÌ…( ÃœÌœ )"
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
            TxtAdvanceValue.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = "ÿ·» ”·›… ‰ﬁœÌ…(  ⁄œÌ· )"
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
            TxtAdvanceValue.locked = False
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
        MsgBox "«·œ›«⁄  «ﬂ»— „‰ «·Õœ ", vbOKOnly, App.Title
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
    Dim RsDev As ADODB.Recordset
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
             
    DcboEmpName2.BoundText = IIf(IsNull(rs("ManagerID").value), "", rs("ManagerID").value)
    DcboJobsType2.BoundText = IIf(IsNull(rs("jobID_approve").value), "", rs("jobID_approve").value)
    txtReason.text = IIf(IsNull(rs("reason").value), "", rs("reason").value)
    opt_ok.value = IIf(rs("ok").value = True, vbChecked, vbUnchecked)
    opt_Notok.value = IIf(rs("notok").value = True, vbChecked, vbUnchecked)
    XPTxtID.text = IIf(IsNull(rs("AdvanceID").value), "", (rs("AdvanceID").value))
    XPDtbTrans.value = IIf(IsNull(rs("AdvanceDate").value), Date, rs("AdvanceDate").value)
    Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    DcboEmpDepartments.BoundText = IIf(IsNull(rs("DeparmentID").value), "", rs("DeparmentID").value)
    DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)
    lbl(44).Caption = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
    DcboJobsType.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
    lbl(23).Caption = IIf(IsNull(rs("basicSalary").value), "", rs("basicSalary").value)
    lbl(22).Caption = IIf(IsNull(rs("EmpDue").value), "", rs("EmpDue").value)
    lbl(20).Caption = IIf(IsNull(rs("Contractvalid").value), "", rs("Contractvalid").value)
    lbl(21).Caption = IIf(IsNull(rs("oldAdvance").value), "", rs("oldAdvance").value)
    TxtDiscount.text = IIf(IsNull(rs("Discount").value), "", rs("Discount").value)
    txtDiscountDES.text = IIf(IsNull(rs("DiscountDES").value), "", rs("DiscountDES").value)
    DBIssueDate.value = IIf(IsNull(rs("DBIssueDate").value), Date, rs("DBIssueDate").value)
    txtDiff.text = IIf(IsNull(rs("DiffVal").value), 0, rs("DiffVal").value)
    If Not (IsNull(rs("MethodDeci").value)) Then
    If rs("MethodDeci").value = 0 Then
       Opt(0).value = True
    ElseIf rs("MethodDeci").value = 1 Then
       Opt(1).value = True
    ElseIf rs("MethodDeci").value = 2 Then
           rs("MethodDeci").value = 2
    End If
    End If
    Me.DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
  '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    Me.TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), "", rs("PaymentCounts").value)
    Me.CmbMonth.ListIndex = rs("FirstMonthPayment").value - 1
    Me.CboYear.text = rs("FirstYearPayment").value
    Me.ChkSaleryDis.value = IIf(rs("AutoDiscount").value = True, vbChecked, vbUnchecked)
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
    StrSQL = "Select * From  TblEmpAdvanceRequestDetails Where AdvanceID=" & val(XPTxtID.text)
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = FG.FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        FG.rows = FG.FixedRows + RsDetails.RecordCount

        For i = Me.FG.FixedRows To FG.rows - 1
            FG.TextMatrix(i, FG.ColIndex("PartNO")) = RsDetails("PartNO").value
            FG.TextMatrix(i, FG.ColIndex("PartValue")) = RsDetails("PartValue").value
            FG.TextMatrix(i, FG.ColIndex("PartDate")) = DisplayDate(CDate(RsDetails("PartDate").value))
            If Not IsNull(RsDetails("Remark").value) Then
            FG.TextMatrix(i, FG.ColIndex("Remark")) = RsDetails("Remark").value
            End If
            
            RsDetails.MoveNext
        Next i
FG.AutoSize 0, FG.Cols - 1, False
    End If

    RsDetails.Close
    Set RsDetails = Nothing
   StrSQL = "  SELECT     dbo.TblEmpAdvanceRequestDetails2.ID, dbo.TblEmpAdvanceRequestDetails2.AdvanceID, dbo.TblEmpAdvanceRequestDetails2.salary,"
   StrSQL = StrSQL & "                   dbo.TblEmpAdvanceRequestDetails2.LongContarct, dbo.TblEmpAdvanceRequestDetails2.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
   StrSQL = StrSQL & "                     dbo.TblEmployee.fullcode"
   StrSQL = StrSQL & "   FROM         dbo.TblEmpAdvanceRequestDetails2 LEFT OUTER JOIN"
   StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TblEmpAdvanceRequestDetails2.EmpID = dbo.TblEmployee.Emp_ID"
   StrSQL = StrSQL & "   Where (dbo.TblEmpAdvanceRequestDetails2.AdvanceID =" & val(XPTxtID.text) & ")"
Set RsDev = New ADODB.Recordset
  VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.rows = VSFlexGrid1.FixedRows
       RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsDev.RecordCount > 0) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid1
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
 
                 .TextMatrix(i, .ColIndex("LineNo")) = i
            
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("EmpID").value), 0, RsDev("EmpID").value)
                .TextMatrix(i, .ColIndex("LongContarct")) = IIf(IsNull(RsDev("LongContarct").value), "", RsDev("LongContarct").value)
                .TextMatrix(i, .ColIndex("salary")) = IIf(IsNull(RsDev("salary").value), "", RsDev("salary").value)
            
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
                Else
                
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
               End If
                          
                RsDev.MoveNext
            Next i
 
        End With
   End If
   Reline
    fillapprovData
    lbl(39).Caption = GetCountPayment()
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
          Msg = "Please Select Employee"
          End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboEmpName.SetFocus
          Sendkeys "{F4}"
            Exit Sub
        End If

   

        If CheckPartCal = False Then
            Exit Sub
        End If

        If CheckDate = False Then
            Exit Sub
        End If

        '”·› ”«»ﬁ…
        Dim RsTest As New ADODB.Recordset
        'Set RsTest = New ADODB.Recordset
        StrSQL = "SELECT dbo.TblEmpAdvanceRequest.AdvanceID, dbo.TblEmpAdvanceRequest.Emp_ID, dbo.TblEmpAdvanceRequestDetails.Payed, dbo.TblEmpAdvanceRequestDetails.PartValue FROM dbo.TblEmpAdvanceRequest INNER JOIN dbo.TblEmpAdvanceRequestDetails ON dbo.TblEmpAdvanceRequest.AdvanceID = dbo.TblEmpAdvanceRequestDetails.AdvanceID WHERE (dbo.TblEmpAdvanceRequestDetails.Payed IS NULL) AND (dbo.TblEmpAdvanceRequest.Emp_ID =" & Me.DcboEmpName.BoundText & ")"
        'RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        'If RsTest.RecordCount > 0 Then
        'MsgBox "«·„ÊŸ› " & DcboEmpName.text & "  ⁄·ÌÂ ”·› ”«»ﬁ… ·„  ”œœ »⁄œ"
        'RsTest.Close
        ' Exit Sub
        'End If

   '     If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.TxtAdvanceValue.text), Me.XPDtbTrans.value) = False Then
   '         Exit Sub
   '     End If

     '   CalCulateParts
     Reline
      If Opt(2).value = True Then
    If val(TxtValue.text) <> val(TxtAdvanceValue.text) Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "ÌÊÃœ  ›—ﬁ ›Ì «·ﬁÌ„ Ì—ÃÏ  ⁄œÌ·Â "
    Else
    MsgBox "There is a difference in values Please modify it"
    End If
    Exit Sub
    End If
   End If
 
        
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

        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("TblEmpAdvanceRequest", "AdvanceID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From TblEmpAdvanceRequestDetails Where AdvanceID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblEmpAdvanceRequestDetails2 Where AdvanceID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            

        End If

        rs("branch_no").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
        rs("AdvanceID").value = val(XPTxtID.text)
        rs("AdvanceDate").value = XPDtbTrans.value
        rs("Emp_ID").value = Me.DcboEmpName.BoundText
        rs("Balance").value = val(lbl(44).Caption)
        rs("DeparmentID").value = val(Me.DcboEmpDepartments.BoundText)
        rs("gradeID").value = val(Me.DcboSpecifications.BoundText)
        rs("JobTypeID").value = val(Me.DcboJobsType.BoundText)
        rs("basicSalary").value = val(lbl(23).Caption)
        rs("DBIssueDate").value = DBIssueDate.value
        rs("Discount").value = IIf(TxtDiscount.text = "", Null, val(TxtDiscount.text))
        rs("DiscountDES").value = IIf(txtDiscountDES.text = "", Null, (txtDiscountDES.text))
        rs("AdvanceValue").value = IIf(TxtAdvanceValue.text = "", Null, val(TxtAdvanceValue.text))
        rs("EmpDue").value = IIf(lbl(22).Caption = "", Null, val(lbl(22).Caption))
        rs("Contractvalid").value = IIf(lbl(20).Caption = "", Null, val(lbl(20).Caption))
        rs("oldAdvance").value = IIf(lbl(21).Caption = "", Null, val(lbl(21).Caption))
        rs("FirstDate").value = IIf(IsDate(FG.TextMatrix(Me.FG.FixedRows, FG.ColIndex("PartDate"))), FG.TextMatrix(Me.FG.FixedRows, FG.ColIndex("PartDate")), Null)
        rs("PaymentCounts").value = val(Me.TxtPaymentCounts.text)
        rs("AutoDiscount").value = IIf(Me.ChkSaleryDis.value = vbChecked, 1, 0)
        rs("FirstMonthPayment").value = Me.CmbMonth.ListIndex + 1
        rs("FirstYearPayment").value = val(Me.CboYear.text)
        rs("UserID").value = Me.DCboUserName.BoundText
        If Opt(0).value = True Then
        rs("MethodDeci").value = 0
        ElseIf Opt(1).value = True Then
        rs("MethodDeci").value = 1
        ElseIf Opt(2).value = True Then
        rs("MethodDeci").value = 2
        End If
        rs("DiffVal").value = val(Me.txtDiff.text)
        
         rs("jobID_approve").value = Me.DCboUserName.BoundText
         rs("ok").value = opt_ok.value
         rs("notok").value = opt_Notok.value
         rs("reason").value = txtReason.text
         rs("ManagerID").value = IIf(Me.DcboEmpName2.BoundText <> "", val(Me.DcboEmpName2.BoundText), Null)
        
        
        
'        rs("AdvanceType").value = 0
'        rs("RetrunID").value = Null
        rs.update
        Set RsDetails = New ADODB.Recordset
   '     RsDetails.Open "TblEmpAdvanceRequestDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
 StrSQL = "SELECT     dbo.TblEmpAdvanceRequestDetails.* from dbo.TblEmpAdvanceRequestDetails Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
        For i = Me.FG.FixedRows To FG.rows - 1
            RsDetails.AddNew
            RsDetails("AdvanceID").value = val(XPTxtID.text)
            RsDetails("PartNO").value = FG.TextMatrix(i, FG.ColIndex("PartNO"))
            RsDetails("PartValue").value = FG.TextMatrix(i, FG.ColIndex("PartValue"))
            RsDetails("PartDate").value = FG.TextMatrix(i, FG.ColIndex("PartDate"))
            RsDetails.update
        Next i
    ''///
                  Set RsDetails = New ADODB.Recordset
          StrSQL = "SELECT     *  from dbo.TblEmpAdvanceRequestDetails2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid1

   
       For i = .FixedRows To .rows - 1
       If val(.TextMatrix(i, .ColIndex("id"))) <> 0 Then
           RsDetails.AddNew
           RsDetails("AdvanceID").value = val(XPTxtID.text)
           RsDetails("EmpID").value = val(.TextMatrix(i, .ColIndex("id")))
           RsDetails("salary").value = val(.TextMatrix(i, .ColIndex("salary")))
           RsDetails("LongContarct").value = val(.TextMatrix(i, .ColIndex("LongContarct")))
                RsDetails.update
           End If
        Next i
        End With
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
    
        Select Case Me.TxtModFlg.text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
           Else
                   Msg = "This is Record Already Saved  " & CHR(13)
                Msg = Msg + "You Need Enter Another Record"
         End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
        End Select

        TxtModFlg.text = "R"
    End If

 CuurentLogdata

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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
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
  Msg = "Confirm Delete"
  End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
            CuurentLogdata ("D")
            
        Deletepost Me.Name, "TblEmpAdvanceRequest", "AdvanceID", val(DcboEmpDepartments.BoundText), val(Dcbranch.BoundText), val(XPTxtID.text), XPTxtID.text
                rs.delete
            '     CuurentLogdata ("D")
                 
               ' StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.Text)
               ' Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblEmpAdvanceRequestDetails Where AdvanceID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblEmpAdvanceRequestDetails2 Where AdvanceID=" & val(Me.XPTxtID.text)
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub



Function FillApprovedTable()
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   Dim StrSQL As String
'   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
 StrSQL = "SELECT     dbo.ApprovalData.* from dbo.ApprovalData Where (Transaction_ID = -1)"
   RSApproval.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   

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
Dim EmpID As Integer

 
 
    If Rs1.RecordCount > 0 Then
            currentdate = Now
            
            
                        GetApprovalDepartement val(DcboEmpDepartments.BoundText), UserID, EmpID
            
                            If UserID <> 0 Then
                           '***************************************
                                                 RSApproval.AddNew
                                        RSApproval("ScreenName").value = Me.Name
                                        RSApproval("levelo").value = 1
                                       RSApproval("EmpID").value = UserID
                                        RSApproval("levelorder").value = 1
                                         RSApproval("currorder").value = 1
                                          RSApproval("Transaction_ID").value = val(XPTxtID.text)
                                          RSApproval("NoteSerial").value = XPTxtID.text
                                        RSApproval("Transaction_Date").value = Date
                                        
                                          RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
                                       RSApproval("SendTime").value = currentdate
                        
                                 
                                                RSApproval("Currcursor").value = 1
                                                 RSApproval("FromUser").value = user_name
                                     
                                        
                                        RSApproval.update
                              End If
              
              
              
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
                
                                                 If i = 1 And UserID = 0 Then
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
 
'    GRID2.Rows = 0
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
 GRID2.rows = 1
    End If
RsDetails.Close

End Function

Function fillapprovDatax()
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
        GRID2.rows = RsDetails.RecordCount + 1
 

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
 GRID2.rows = 1
    End If
RsDetails.Close

End Function

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
  
    With VSFlexGrid1

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("id")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If
        Next i
 
    End With
  End Sub
Function GetCountPayment() As Integer
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim sql As String
sql = "SELECT     COUNT(AdvanceID) AS Cunt"

sql = sql & " From dbo.TblEmpAdvanceRequestDetails"
sql = sql & " WHERE     (AdvanceID = " & val(XPTxtID.text) & ") AND (PartDate >=" & SQLDate(Date, True) & " )"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
GetCountPayment = IIf(IsNull(Rs7("Cunt").value), 0, Rs7("Cunt").value)
Else
GetCountPayment = 0
End If
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
        .Create Me.hWnd, "ÿ·» ”·›… ‰ﬁœÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ”·›… ‰ﬁœÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ”·›… ‰ﬁœÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ”·›… ‰ﬁœÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ”·›… ‰ﬁœÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ”·›… ‰ﬁœÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ”·›… ‰ﬁœÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ”·›… ‰ﬁœÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ”·›… ‰ﬁœÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ”·›… ‰ﬁœÌ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ”·›… ‰ﬁœÌ…", 1, 15204351, -2147483630
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

Private Sub TxtAdvanceValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtAdvanceValue.text, 0)
End Sub

Private Function CheckDate() As Boolean
    Dim StrTemp As String
    Dim Msg  As String
 



    If year(Date) > val(Me.CboYear.text) Then ' ⁄«„ „÷Ï
'        Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ ÕÌÀ «‰Â ﬁ»·  «—ÌŒ «·ÌÊ„ ÊÂ–« ·« ÌÃÊ“...!!!"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        CheckDate = False
'        Exit Function
    ElseIf year(Date) = val(Me.CboYear.text) Then '‰›” «·⁄«„

'        If Month(Date) > val(Me.CmbMonth.ListIndex + 1) Then
'            Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ  ÕÌÀ «‰Â ﬁ»·  «—ÌŒ «·ÌÊ„ ÊÂ–« ·« ÌÃÊ“...!!!"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        CheckDate = False
'            Exit Function
'        End If
    End If

    CheckDate = True
End Function

Private Function CheckPartCal() As Boolean
    Dim Msg As String

    CheckPartCal = False

    If val(TxtAdvanceValue.text) = 0 Then
        Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·”·›…...!!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtAdvanceValue.SetFocus
        Exit Function
    End If

    If val(TxtPaymentCounts.text) = 0 Then
        Msg = "ÌÃ» «œŒ«· ⁄œœ „—«   ”œÌœ «·œ›⁄…...!!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtAdvanceValue.SetFocus
        Exit Function
    End If

    If CmbMonth.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ «Ê· ‘Â— · ”œÌœ «·œ›⁄…..!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CmbMonth.SetFocus
         Sendkeys "{F4}"
        Exit Function
    End If

    If CboYear.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ «Ê· ”‰… · ”œÌœ «·œ›⁄… ..!! "
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
    Dim Diff As Double
    Dim Sm As Double
    Sm = 0
    If CheckPartCal = False Then
        Exit Sub
    End If

    If CheckDate = False Then
        Exit Sub
    End If

    SngPartValue = val(Me.TxtAdvanceValue.text) / val(Me.TxtPaymentCounts.text)
    IntPartCounts = val(Me.TxtPaymentCounts.text)
    m_FirstDate = CDate(val(Me.CboYear.text) & "-" & Me.CmbMonth.ListIndex + 1 & "-01")

    With Me.FG
        .Clear flexClearScrollable, flexClearEverything
        .rows = .FixedRows + IntPartCounts
        .RowHeightMin = 300

        For i = 1 To IntPartCounts
            .TextMatrix(i, .ColIndex("PartNO")) = i
            .TextMatrix(i, .ColIndex("PartValue")) = Round(SngPartValue, 0)
            Sm = Sm + Round(SngPartValue, 0)
            .TextMatrix(i, .ColIndex("PartDate")) = DisplayDate(DateAdd("m", i - 1, m_FirstDate))
        Next i
TxtValue.text = Sm

    End With

End Sub

