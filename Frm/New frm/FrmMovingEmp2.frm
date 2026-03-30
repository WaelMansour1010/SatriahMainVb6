VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmMovingEmp2 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12480
   Icon            =   "FrmMovingEmp2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   12480
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   13920
      TabIndex        =   35
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   13440
      TabIndex        =   34
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
      TabIndex        =   32
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
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   9960
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
      TabIndex        =   2
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
      Caption         =   ""
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
         ButtonImage     =   "FrmMovingEmp2.frx":038A
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
         ButtonImage     =   "FrmMovingEmp2.frx":0724
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
         Left            =   1800
         TabIndex        =   5
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
         ButtonImage     =   "FrmMovingEmp2.frx":0ABE
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
         TabIndex        =   6
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
         ButtonImage     =   "FrmMovingEmp2.frx":0E58
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
         BackColor       =   &H00FFFFFF&
         Caption         =   "     ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   75
         Left            =   6720
         TabIndex        =   191
         Top             =   0
         Width           =   5205
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   5880
         Picture         =   "FrmMovingEmp2.frx":11F2
         Stretch         =   -1  'True
         Top             =   120
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
         TabIndex        =   33
         Top             =   0
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   7740
      TabIndex        =   7
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   180092929
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   870
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   8280
      Width           =   11145
      _cx             =   19659
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
         Left            =   9360
         TabIndex        =   9
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
         Left            =   8535
         TabIndex        =   10
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
         Left            =   7695
         TabIndex        =   11
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
         Left            =   6840
         TabIndex        =   12
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
         Left            =   5985
         TabIndex        =   13
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
         Left            =   2640
         TabIndex        =   14
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
         Left            =   3495
         TabIndex        =   15
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
         Left            =   1320
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
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
         TabIndex        =   37
         Top             =   60
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…  ’—ÌÕ Œ—ÊÃ"
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
         Index           =   11
         Left            =   4440
         TabIndex        =   181
         Top             =   75
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…  ›ÊÌ÷"
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
      Left            =   7110
      TabIndex        =   16
      Top             =   7860
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
      TabIndex        =   17
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
   Begin MSDataListLib.DataCombo Dcbranch 
      Bindings        =   "FrmMovingEmp2.frx":4E5A
      Height          =   315
      Left            =   4230
      TabIndex        =   30
      Top             =   720
      Width           =   2865
      _ExtentX        =   5054
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
      Height          =   6735
      Left            =   -60
      TabIndex        =   38
      Top             =   1140
      Width           =   12480
      _cx             =   22013
      _cy             =   11880
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
      Caption         =   " ’—ÌÕ Œ—ÊÃ „ƒﬁ |Õ«·Â «·«⁄ „«œ| ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…"
      Align           =   0
      CurrTab         =   2
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
      Picture(0)      =   "FrmMovingEmp2.frx":4E6F
      Flags(0)        =   2
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6270
         Left            =   -13035
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   45
         Width           =   12390
         _cx             =   21855
         _cy             =   11060
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
            Height          =   1590
            Left            =   -120
            TabIndex        =   40
            Tag             =   "1"
            Top             =   240
            Width           =   13230
            _cx             =   23336
            _cy             =   2805
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
            FormatString    =   $"FrmMovingEmp2.frx":5209
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
            TabIndex        =   96
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label Label111000 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   87
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
            TabIndex        =   41
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6270
         Index           =   15
         Left            =   -13335
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   45
         Width           =   12390
         _cx             =   21855
         _cy             =   11060
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
         _GridInfo       =   $"FrmMovingEmp2.frx":534C
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6240
            Index           =   16
            Left            =   15
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   15
            Width           =   12360
            _cx             =   21802
            _cy             =   11007
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
            Begin VB.TextBox TxtInterval 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3600
               MaxLength       =   10
               TabIndex        =   106
               Top             =   120
               Width           =   1395
            End
            Begin VB.TextBox TxtSearchCode 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9300
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txtRemark 
               Alignment       =   1  'Right Justify
               Height          =   2040
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   84
               Top             =   2145
               Width           =   11955
            End
            Begin VB.TextBox TxtDiscount 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   13410
               MaxLength       =   10
               TabIndex        =   82
               Top             =   2100
               Width           =   1425
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
               Height          =   3765
               Index           =   0
               Left            =   14145
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   360
               Width           =   6135
               Begin VB.TextBox TxtPaymentCounts 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   4110
                  MaxLength       =   2
                  TabIndex        =   75
                  Top             =   240
                  Width           =   825
               End
               Begin VB.ComboBox CmbMonth 
                  Height          =   315
                  Left            =   4110
                  Style           =   2  'Dropdown List
                  TabIndex        =   74
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
                  TabIndex        =   73
                  Top             =   2160
                  Value           =   1  'Checked
                  Visible         =   0   'False
                  Width           =   1965
               End
               Begin VB.ComboBox CboYear 
                  Height          =   315
                  Left            =   4110
                  Style           =   2  'Dropdown List
                  TabIndex        =   72
                  Top             =   1320
                  Width           =   1095
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   435
                  Index           =   8
                  Left            =   4080
                  TabIndex        =   71
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
                  ButtonImage     =   "FrmMovingEmp2.frx":5380
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   2325
                  Left            =   90
                  TabIndex        =   76
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
                  FormatString    =   $"FrmMovingEmp2.frx":571A
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
                  TabIndex        =   81
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
                  TabIndex        =   80
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
                  TabIndex        =   79
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
                  TabIndex        =   78
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
                  TabIndex        =   77
                  Top             =   1320
                  Width           =   405
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  „«·Ì…"
               Height          =   1005
               Left            =   14760
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   0
               Width           =   6015
               Begin MSDataListLib.DataCombo DcboSpecifications 
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   61
                  Top             =   360
                  Width           =   1995
                  _ExtentX        =   3519
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·„” Õﬁ«  ··„ÊŸ›"
                  Height          =   285
                  Index           =   17
                  Left            =   3960
                  TabIndex        =   69
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
                  TabIndex        =   68
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
                  TabIndex        =   67
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
                  TabIndex        =   66
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
                  TabIndex        =   65
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
                  TabIndex        =   64
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
                  TabIndex        =   63
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
                  TabIndex        =   62
                  Top             =   360
                  Width           =   1125
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «‰ﬁ·"
               Height          =   1545
               Left            =   13815
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   0
               Width           =   6105
               Begin MSDataListLib.DataCombo DcboEmpDepartments 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   52
                  Top             =   240
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DBIssueDate 
                  Height          =   315
                  Left            =   6480
                  TabIndex        =   53
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   180158465
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcboJobsType 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   54
                  Top             =   600
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DataCombo1 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   98
                  Top             =   240
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DTPicker1 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   100
                  Top             =   600
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   180158465
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DataCombo2 
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   101
                  Top             =   960
                  Width           =   3315
                  _ExtentX        =   5847
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»ÊŸÌ›…"
                  Height          =   285
                  Index           =   38
                  Left            =   5160
                  TabIndex        =   102
                  Top             =   960
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰ ÌÊ„"
                  Height          =   285
                  Index           =   37
                  Left            =   2280
                  TabIndex        =   99
                  Top             =   600
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»ÊŸÌ›…"
                  Height          =   285
                  Index           =   36
                  Left            =   5160
                  TabIndex        =   97
                  Top             =   600
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—« » «·«”«”Ì"
                  Height          =   285
                  Index           =   5
                  Left            =   6600
                  TabIndex        =   59
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
                  Height          =   285
                  Index           =   13
                  Left            =   6360
                  TabIndex        =   58
                  Top             =   360
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï ﬁ”„"
                  Height          =   285
                  Index           =   15
                  Left            =   2280
                  TabIndex        =   57
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   23
                  Left            =   6240
                  TabIndex        =   56
                  Top             =   480
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰ ﬁ”„"
                  Height          =   285
                  Index           =   24
                  Left            =   5280
                  TabIndex        =   55
                  Top             =   240
                  Width           =   645
               End
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   510
               Left            =   240
               TabIndex        =   86
               Top             =   795
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   900
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
            Begin MSComCtl2.DTPicker TxtExpectedouttime 
               Height          =   315
               Left            =   8940
               TabIndex        =   90
               Top             =   1005
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   180158467
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSComCtl2.DTPicker txtExpectedIntime 
               Height          =   375
               Left            =   8940
               TabIndex        =   91
               Top             =   1425
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   180158467
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSComCtl2.DTPicker txtActualouttime 
               Height          =   315
               Left            =   5880
               TabIndex        =   92
               Top             =   1005
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   180158467
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSComCtl2.DTPicker txtActualIntime 
               Height          =   375
               Left            =   5880
               TabIndex        =   93
               Top             =   1425
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   180158467
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSDataListLib.DataCombo DcboEmpName 
               Height          =   315
               Left            =   5880
               TabIndex        =   104
               Top             =   480
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcOutType 
               Height          =   315
               Left            =   5880
               TabIndex        =   109
               Top             =   120
               Width           =   4635
               _ExtentX        =   8176
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·«–‰"
               Height          =   285
               Index           =   33
               Left            =   10590
               TabIndex        =   110
               Top             =   120
               Width           =   1365
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„œ…"
               Height          =   285
               Index           =   2
               Left            =   4710
               TabIndex        =   108
               Top             =   120
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”«⁄Â"
               Height          =   285
               Index           =   29
               Left            =   2040
               TabIndex        =   107
               Top             =   150
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   285
               Index           =   3
               Left            =   10950
               TabIndex        =   105
               Top             =   495
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·⁄Êœ… «·›⁄·Ì"
               Height          =   255
               Index           =   35
               Left            =   7320
               TabIndex        =   95
               Top             =   1440
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·Œ—ÊÃ «·›⁄·Ì"
               Height          =   210
               Index           =   34
               Left            =   7320
               TabIndex        =   94
               Top             =   960
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·⁄Êœ… «·„ Êﬁ⁄"
               Height          =   375
               Index           =   32
               Left            =   10560
               TabIndex        =   89
               Top             =   1425
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·Œ—ÊÃ «·„ Êﬁ⁄"
               Height          =   450
               Index           =   31
               Left            =   10560
               TabIndex        =   88
               Top             =   1005
               Width           =   1605
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Â„…"
               Height          =   330
               Index           =   28
               Left            =   5160
               TabIndex        =   85
               Top             =   1785
               Width           =   600
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÊÌŒ’„ „‰ «·”·› „»·€« Êﬁœ—…"
               Height          =   330
               Index           =   26
               Left            =   12045
               TabIndex        =   83
               Top             =   1425
               Width           =   2280
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2520
               Index           =   62
               Left            =   2370
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   1155
               Width           =   540
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6240
            Index           =   9
            Left            =   15
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   15
            Width           =   12360
            _cx             =   21802
            _cy             =   11007
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
               Height          =   3270
               Left            =   1065
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   840
               Width           =   225
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   2250
               Left            =   1350
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1320
               Width           =   345
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2250
               Index           =   67
               Left            =   765
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   1320
               Width           =   195
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   2175
               Index           =   68
               Left            =   1290
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   1395
               Width           =   15
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
               Height          =   2340
               Index           =   69
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1320
               Width           =   105
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6270
         Index           =   0
         Left            =   45
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   45
         Width           =   12390
         _cx             =   21855
         _cy             =   11060
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
         Begin VB.Frame Frame8 
            BackColor       =   &H00E2E9E9&
            Height          =   495
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   199
            Top             =   150
            Width           =   8835
            Begin VB.OptionButton optAuthorization 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ›ÊÌ÷"
               Height          =   285
               Left            =   6630
               RightToLeft     =   -1  'True
               TabIndex        =   201
               Top             =   180
               Width           =   855
            End
            Begin VB.OptionButton optNotAuthorization 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·€«¡ «· ›ÊÌ÷"
               Height          =   285
               Left            =   4890
               RightToLeft     =   -1  'True
               TabIndex        =   200
               Top             =   180
               Width           =   1395
            End
            Begin MSComCtl2.DTPicker txtDateCancel 
               Height          =   315
               Left            =   1110
               TabIndex        =   202
               Top             =   120
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               Format          =   180158465
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·«·€«¡"
               Height          =   285
               Index           =   77
               Left            =   2910
               TabIndex        =   203
               Top             =   180
               Width           =   1005
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00E2E9E9&
            ClipControls    =   0   'False
            Height          =   1035
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   2700
            Width           =   12135
            Begin MSDataListLib.DataCombo DcbDept3 
               Height          =   315
               Left            =   7260
               TabIndex        =   193
               Top             =   240
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbKafelName 
               Height          =   315
               Left            =   90
               TabIndex        =   195
               Top             =   195
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDepartment2 
               Height          =   315
               Left            =   7260
               TabIndex        =   197
               Top             =   600
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁ”„"
               Height          =   225
               Index           =   5
               Left            =   11070
               TabIndex        =   198
               Top             =   660
               Width           =   885
            End
            Begin VB.Label Lab 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’«Õ» «·⁄„·"
               Height          =   360
               Index           =   19
               Left            =   3435
               RightToLeft     =   -1  'True
               TabIndex        =   196
               Top             =   255
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«œ«—…"
               Height          =   285
               Index           =   76
               Left            =   10605
               RightToLeft     =   -1  'True
               TabIndex        =   194
               Top             =   270
               Width           =   1335
            End
         End
         Begin VB.TextBox TxtRemarks2 
            Alignment       =   1  'Right Justify
            Height          =   480
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   171
            Top             =   5580
            Width           =   12075
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·„—ﬂ»…"
            ForeColor       =   &H00800000&
            Height          =   1425
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   166
            Top             =   3750
            Width           =   12135
            Begin VB.TextBox TxtBoardNO 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox TxtOperatorN 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   600
               Width           =   4215
            End
            Begin MSDataListLib.DataCombo DcbEquepment 
               Height          =   315
               Left            =   6270
               TabIndex        =   173
               Tag             =   "Õœœ «”„ «·„⁄œ…"
               Top             =   240
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbColor 
               Height          =   315
               Left            =   120
               TabIndex        =   175
               Tag             =   "Õœœ «”„ «·„⁄œ…"
               Top             =   960
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbModel 
               Height          =   315
               Left            =   6270
               TabIndex        =   176
               Tag             =   "Õœœ «”„ «·„⁄œ…"
               Top             =   960
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbTypeEqup 
               Height          =   315
               Left            =   6270
               TabIndex        =   179
               Tag             =   "Õœœ «”„ «·„⁄œ…"
               Top             =   570
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   435
               Left            =   120
               TabIndex        =   182
               TabStop         =   0   'False
               Top             =   240
               Width           =   2325
               _cx             =   4101
               _cy             =   767
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               Begin VB.TextBox txtNum4 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   0
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   190
                  Top             =   0
                  Width           =   300
               End
               Begin VB.TextBox txtLetter4 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1155
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   189
                  Top             =   0
                  Width           =   360
               End
               Begin VB.TextBox txtNum3 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   270
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   188
                  Top             =   0
                  Width           =   300
               End
               Begin VB.TextBox txtNum2 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   480
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   187
                  Top             =   0
                  Width           =   330
               End
               Begin VB.TextBox txtNum1 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   795
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   186
                  Top             =   0
                  Width           =   360
               End
               Begin VB.TextBox txtLetter3 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1440
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   185
                  Top             =   0
                  Width           =   315
               End
               Begin VB.TextBox txtLetter2 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1710
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   184
                  Top             =   0
                  Width           =   240
               End
               Begin VB.TextBox txtLetter1 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1935
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   183
                  Top             =   0
                  Width           =   285
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·„—ﬂ»…"
               Height          =   210
               Index           =   73
               Left            =   10875
               RightToLeft     =   -1  'True
               TabIndex        =   180
               Top             =   600
               Width           =   1170
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«··Ê‰"
               Height          =   210
               Index           =   72
               Left            =   4665
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Top             =   960
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÿ—«“"
               Height          =   210
               Index           =   71
               Left            =   10755
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   960
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„"
               Height          =   255
               Index           =   102
               Left            =   10695
               RightToLeft     =   -1  'True
               TabIndex        =   174
               Top             =   240
               Width           =   1350
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «··ÊÕ…"
               Height          =   285
               Index           =   66
               Left            =   4440
               TabIndex        =   170
               Top             =   240
               Width           =   1605
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—ﬁ„ «· ‘€Ì·Ì"
               Height          =   285
               Index           =   65
               Left            =   4320
               TabIndex        =   169
               Top             =   600
               Width           =   1725
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·„›Ê÷ ·Â"
            ForeColor       =   &H00800000&
            Height          =   1005
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   156
            Top             =   1680
            Width           =   12165
            Begin VB.TextBox TxtNumID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   600
               Width           =   3375
            End
            Begin VB.TextBox TxtNationality 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   240
               Width           =   3375
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9510
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   240
               Width           =   1065
            End
            Begin VB.TextBox TxtLeaderName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4800
               RightToLeft     =   -1  'True
               TabIndex        =   157
               Top             =   600
               Width           =   5775
            End
            Begin XtremeSuiteControls.RadioButton dcbLeaderType 
               Height          =   255
               Index           =   0
               Left            =   10680
               TabIndex        =   159
               Top             =   240
               Width           =   1335
               _Version        =   786432
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„ÊŸ›"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbLeaderID 
               Bindings        =   "FrmMovingEmp2.frx":57A5
               Height          =   315
               Left            =   4800
               TabIndex        =   160
               Top             =   240
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
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
            Begin XtremeSuiteControls.RadioButton dcbLeaderType 
               Height          =   255
               Index           =   1
               Left            =   10560
               TabIndex        =   161
               Top             =   600
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "€Ì— „ÊŸ›"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÂÊÌ…"
               Height          =   285
               Index           =   64
               Left            =   3600
               TabIndex        =   165
               Top             =   600
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ã‰”Ì…"
               Height          =   285
               Index           =   60
               Left            =   3480
               TabIndex        =   163
               Top             =   240
               Width           =   1245
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·„›Ê÷"
            ForeColor       =   &H00800000&
            Height          =   735
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   630
            Width           =   12135
            Begin VB.TextBox TxtComputerNo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   154
               Top             =   240
               Width           =   3375
            End
            Begin VB.TextBox TxtName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4800
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   240
               Width           =   6015
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Õ«”»"
               Height          =   285
               Index           =   63
               Left            =   3600
               TabIndex        =   155
               Top             =   240
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„"
               Height          =   285
               Index           =   61
               Left            =   11040
               TabIndex        =   153
               Top             =   240
               Width           =   765
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «‰ﬁ·"
            Height          =   1545
            Left            =   13815
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Top             =   0
            Width           =   6105
            Begin MSDataListLib.DataCombo DataCombo4 
               Height          =   315
               Left            =   3120
               TabIndex        =   136
               Top             =   240
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   6480
               TabIndex        =   137
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               Format          =   199819265
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo5 
               Height          =   315
               Left            =   3120
               TabIndex        =   138
               Top             =   600
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo6 
               Height          =   315
               Left            =   120
               TabIndex        =   139
               Top             =   240
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   315
               Left            =   120
               TabIndex        =   140
               Top             =   600
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   556
               _Version        =   393216
               Format          =   199819265
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo7 
               Height          =   315
               Left            =   1680
               TabIndex        =   141
               Top             =   960
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ ﬁ”„"
               Height          =   285
               Index           =   59
               Left            =   5280
               TabIndex        =   149
               Top             =   240
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               Height          =   285
               Index           =   58
               Left            =   6240
               TabIndex        =   148
               Top             =   480
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï ﬁ”„"
               Height          =   285
               Index           =   57
               Left            =   2280
               TabIndex        =   147
               Top             =   240
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
               Height          =   285
               Index           =   56
               Left            =   6360
               TabIndex        =   146
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—« » «·«”«”Ì"
               Height          =   285
               Index           =   55
               Left            =   6600
               TabIndex        =   145
               Top             =   360
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»ÊŸÌ›…"
               Height          =   285
               Index           =   54
               Left            =   5160
               TabIndex        =   144
               Top             =   600
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ ÌÊ„"
               Height          =   285
               Index           =   53
               Left            =   2280
               TabIndex        =   143
               Top             =   600
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»ÊŸÌ›…"
               Height          =   285
               Index           =   52
               Left            =   5160
               TabIndex        =   142
               Top             =   960
               Width           =   645
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  „«·Ì…"
            Height          =   1005
            Left            =   14760
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   0
            Width           =   6015
            Begin MSDataListLib.DataCombo DataCombo3 
               Height          =   315
               Left            =   3360
               TabIndex        =   126
               Top             =   360
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„— »…"
               Height          =   285
               Index           =   51
               Left            =   4800
               TabIndex        =   134
               Top             =   360
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               Height          =   285
               Index           =   50
               Left            =   3240
               TabIndex        =   133
               Top             =   720
               Width           =   525
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               Height          =   285
               Index           =   49
               Left            =   960
               TabIndex        =   132
               Top             =   360
               Width           =   525
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               Height          =   285
               Index           =   48
               Left            =   960
               TabIndex        =   131
               Top             =   720
               Width           =   525
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   285
               Index           =   47
               Left            =   -240
               TabIndex        =   130
               Top             =   720
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”·› ·„  ”œœ"
               Height          =   285
               Index           =   46
               Left            =   1800
               TabIndex        =   129
               Top             =   360
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œ… «·⁄ﬁœ «·„ »ﬁÌ…"
               Height          =   285
               Index           =   45
               Left            =   1560
               TabIndex        =   128
               Top             =   720
               Width           =   1365
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì «·„” Õﬁ«  ··„ÊŸ›"
               Height          =   285
               Index           =   44
               Left            =   3960
               TabIndex        =   127
               Top             =   720
               Width           =   1965
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
            Height          =   3765
            Index           =   1
            Left            =   14145
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   360
            Width           =   6135
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   4110
               Style           =   2  'Dropdown List
               TabIndex        =   117
               Top             =   1320
               Width           =   1095
            End
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈Œ’„ „‰ «·„— »  ·ﬁ«∆Ì«"
               Enabled         =   0   'False
               Height          =   315
               Left            =   3960
               TabIndex        =   116
               Top             =   2160
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   4110
               Style           =   2  'Dropdown List
               TabIndex        =   115
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               Height          =   345
               Left            =   4110
               MaxLength       =   2
               TabIndex        =   114
               Top             =   240
               Width           =   825
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   435
               Index           =   10
               Left            =   4080
               TabIndex        =   118
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
               ButtonImage     =   "FrmMovingEmp2.frx":57BA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   2325
               Left            =   90
               TabIndex        =   119
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
               FormatString    =   $"FrmMovingEmp2.frx":5B54
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
               Caption         =   "”‰…"
               Height          =   315
               Index           =   43
               Left            =   5250
               TabIndex        =   124
               Top             =   1320
               Width           =   405
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   315
               Index           =   42
               Left            =   5250
               TabIndex        =   123
               Top             =   990
               Width           =   405
            End
            Begin VB.Label Label2 
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
               TabIndex        =   122
               Top             =   2280
               Visible         =   0   'False
               Width           =   2595
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «Ê· œ›⁄…"
               Height          =   285
               Index           =   41
               Left            =   4380
               TabIndex        =   121
               Top             =   690
               Width           =   1665
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·œ›⁄« "
               Height          =   285
               Index           =   40
               Left            =   4830
               TabIndex        =   120
               Top             =   300
               Width           =   975
            End
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   13410
            MaxLength       =   10
            TabIndex        =   112
            Top             =   2100
            Width           =   1425
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   70
            Left            =   5400
            TabIndex        =   172
            Top             =   5220
            Width           =   1965
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Enabled         =   0   'False
            Height          =   2520
            Index           =   74
            Left            =   2370
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   1155
            Width           =   540
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
      TabIndex        =   36
      Top             =   3330
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   120
      Top             =   2520
      Width           =   6015
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Index           =   39
      Left            =   7020
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   12810
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ "
      Height          =   285
      Index           =   4
      Left            =   11520
      TabIndex        =   25
      Top             =   720
      Width           =   765
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   8790
      TabIndex        =   24
      Top             =   735
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   11685
      TabIndex        =   23
      Top             =   5715
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2760
      TabIndex        =   22
      Top             =   7980
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   1290
      TabIndex        =   21
      Top             =   8010
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   450
      TabIndex        =   20
      Top             =   7980
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   2220
      TabIndex        =   19
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   13350
      TabIndex        =   18
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmMovingEmp2"
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


Private Sub Label3_Click()

End Sub

Private Sub DcbCarModel_Click(Area As Integer)

End Sub

Private Sub DcbDept3_Change()
LoadDept
End Sub

Private Sub DcbDept3_Click(Area As Integer)
LoadDept
End Sub

Private Sub lblModel_Click()

End Sub


Private Sub Text7_Change()
      
End Sub

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

Private Sub Cmd_Click(index As Integer)

    ' On Error GoTo ErrTrap
    Select Case index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
Dim cCompanyInfo As New ClsCompanyInfo
            TxtModFlg.text = "N"
            clear_all Me
             Me.TxtName.text = cCompanyInfo.ArabCompanyName
             Me.TxtComputerNo.text = cCompanyInfo.ComputerNo
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
                                               
             
    GRID2.Clear flexClearScrollable, flexClearEverything
    GRID2.rows = 1
    
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
         Load FrmEmpAdvanceSearch
        FrmEmpAdvanceSearch.show

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
       Case 11

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report2 val(Me.XPTxtID.text)
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
MySQL = MySQL & "  SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
MySQL = MySQL & "  dbo.TblEmployee.Emp_Namee, dbo.TblUsers.UserName, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
MySQL = MySQL & "  dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeparmentID, dbo.TblEmpJobsTypes.JobTypeID,"
MySQL = MySQL & "  dbo.TblEmpPassOver2.AdvanceID, dbo.TblEmpPassOver2.[interval], dbo.TblEmpPassOver2.AdvanceDate, dbo.TblEmpPassOver2.Remark,"
MySQL = MySQL & "  dbo.TblEmpPassOver2.Expectedouttime, dbo.TblEmpPassOver2.ExpectedIntime, dbo.TblEmpPassOver2.Actualouttime, dbo.TblEmpPassOver2.ActualIntime,"
MySQL = MySQL & "  dbo.TblEmpPassOver2.OutTypeID, dbo.TblOutType.name, dbo.TblEmpPassOver2.PostedDate, dbo.TblEmpPassOver2.NoteSerial, dbo.TblEmpPassOver2.Approved,"
MySQL = MySQL & "  dbo.TblEmpPassOver2.Transaction_ID , dbo.TblEmpPassOver2.Posted, dbo.TblOutType.namee"
MySQL = MySQL & "  FROM         dbo.TblEmployee RIGHT OUTER JOIN"
MySQL = MySQL & "  dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
MySQL = MySQL & "  dbo.TblEmpPassOver2 LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblOutType ON dbo.TblEmpPassOver2.OutTypeID = dbo.TblOutType.id ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmpPassOver2.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblEmpDepartments ON dbo.TblEmpPassOver2.DeparmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblUsers ON dbo.TblEmpPassOver2.UserID = dbo.TblUsers.UserID ON dbo.TblEmployee.Emp_ID = dbo.TblEmpPassOver2.Emp_id LEFT OUTER JOIN"
MySQL = MySQL & "  dbo.TblBranchesData ON dbo.TblEmpPassOver2.Branch_NO = dbo.TblBranchesData.branch_id"
MySQL = MySQL & "  Where (dbo.TblEmpPassOver2.AdvanceID = " & val(XPTxtID.text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "PassOver.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "PassOver.rpt"
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
    
        StrReportTitle = "" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
     
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
   xReport.ParameterFields(10).AddCurrentValue CStr(FormatDateTime(txtActualouttime.value, vbLongTime))
   xReport.ParameterFields(11).AddCurrentValue CStr(FormatDateTime(txtActualIntime.value, vbLongTime))
      xReport.ParameterFields(12).AddCurrentValue CStr(FormatDateTime(TxtExpectedouttime.value, vbLongTime))
   xReport.ParameterFields(13).AddCurrentValue CStr(FormatDateTime(txtExpectedIntime.value, vbLongTime))
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
Sub RetriveCarsInfo(Optional CarID As Double = 0, Optional OperNo As String, Optional BoardNO As String, Optional Typ As Integer = 0)
If Me.TxtModFlg <> "R" Then
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from TblCarsData"
If Typ = 0 Then
sql = sql & "  Where FixedassetId = " & CarID & ""
ElseIf Typ = 1 Then
sql = sql & " where OperatorN='" & OperNo & "'"
ElseIf Typ = 2 Then
'BoardNO = Replace(Replace(BoardNO, Chr(10), ""), Chr(13), "")
' Sql = Sql & " where   REPLACE(BoardNO, ' ', '') = N'" & BoardNO & "'"
 sql = sql & " where  BoardNO = N'" & BoardNO & "'"
End If
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
If Typ <> 1 Then
Me.TxtOperatorN.text = IIf(IsNull(Rs3("OperatorN").value), "", Rs3("OperatorN").value)
End If
If Typ <> 2 Then
txtBoardNo.text = IIf(IsNull(Rs3("BoardNO").value), "", Rs3("BoardNO").value)
End If
If Typ <> 0 Then
DcbEquepment.BoundText = IIf(IsNull(Rs3("FixedassetId").value), 0, Rs3("FixedassetId").value)
End If
DcbTypeEqup.BoundText = IIf(IsNull(Rs3("CarsTypeId").value), 0, Rs3("CarsTypeId").value)
DcbModel.BoundText = IIf(IsNull(Rs3("VModel").value), 0, Rs3("VModel").value)
DcbColor.BoundText = IIf(IsNull(Rs3("VColor").value), 0, Rs3("VColor").value)
Else
DcbColor.BoundText = 0
DcbModel.BoundText = 0
DcbTypeEqup.BoundText = 0
If Typ <> 1 Then
TxtOperatorN.text = ""
End If
If Typ <> 2 Then
txtBoardNo.text = ""
End If
If Typ <> 0 Then
DcbEquepment.BoundText = 0
End If
End If
End If
End Sub
Function print_report2(Optional NoteSerial As String)
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = MySQL & " SELECT     dbo.TblEmpPassOver2.AdvanceID, dbo.TblEmpPassOver2.NoteSerial, dbo.TblEmpPassOver2.ComputerNo, dbo.TblEmpPassOver2.Name AS EmpName, "
MySQL = MySQL & "                      dbo.TblEmpPassOver2.EmpType, dbo.TblEmpPassOver2.DcbLeaderID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmpPassOver2.LeaderName,"
MySQL = MySQL & "                      dbo.TblEmpPassOver2.Nationality, dbo.TblEmpPassOver2.NumID, dbo.TblEmpPassOver2.Remarks2, dbo.TblEmpPassOver2.BoardNO, dbo.TblEmpPassOver2.OperatorN,"
MySQL = MySQL & "                      dbo.TblEmpPassOver2.ColorID, dbo.TblColor.name AS Colorname, dbo.TblColor.namee AS ColornameE, dbo.TblEmpPassOver2.ModelID, dbo.TblCarModels.Model,"
MySQL = MySQL & "                      dbo.TblCarModels.ModelE, dbo.TblEmpPassOver2.TypeEqupID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblEmpPassOver2.EquepmentID,"
MySQL = MySQL & "                      dbo.FixedAssets.Name AS EqupName, dbo.FixedAssets.namee AS EqupNameE, dbo.TblEmpPassOver2.AdvanceDate"
MySQL = MySQL & " FROM         dbo.TblEmpPassOver2 LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblEmpPassOver2.EquepmentID = dbo.FixedAssets.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TBLCarTypes ON dbo.TblEmpPassOver2.TypeEqupID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCarModels ON dbo.TblEmpPassOver2.ModelID = dbo.TblCarModels.Id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblColor ON dbo.TblEmpPassOver2.ColorID = dbo.TblColor.Id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblEmpPassOver2.DcbLeaderID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & " Where (dbo.TblEmpPassOver2.advanceID = " & val(XPTxtID.text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDelegating.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDelegating.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
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

Public Sub DcbEquepment_Change()
DcbEquepment_Click (0)
End Sub

Private Sub DcbEquepment_Click(Area As Integer)
RetriveCarsInfo val(DcbEquepment.BoundText), , , 0
End Sub

Private Sub DcbEquepment_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
        FrmCasrShearches.SendForm = "MovingEmp"
        FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub DcbLeaderID_Change()

       If val(DcbLeaderID.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbLeaderID.BoundText, EmpCode
    Text2.text = EmpCode
    Dim Nationality As String
    Dim NumEkama As String
   If Me.TxtModFlg = "R" Then Exit Sub
        get_employee_information val(Me.DcbLeaderID.BoundText), , , , , , , , , Nationality, , , , , NumEkama
        TxtNationality.text = Nationality
        TxtNumID.text = NumEkama
        
        Dim s As String
        Dim rsDummy As New ADODB.Recordset
        s = "Select * from TblEmployee  Where Emp_ID = " & val(DcbLeaderID.BoundText)
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            DcbDept3.BoundText = val(rsDummy!DepartmentID & "")
            DcbDepartment2.BoundText = val(rsDummy!DeptID2 & "")
            DcbKafelName.BoundText = Trim(rsDummy!KafelName & "")
        End If
End Sub

Private Sub DcbLeaderID_Click(Area As Integer)
DcbLeaderID_Change
End Sub

Private Sub DcbLeaderID_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 44
        FrmEmployeeSearch.show
    End If
End Sub

Private Sub dcbLeaderType_Click(index As Integer)
If dcbLeaderType(0).value = True Then
TxtSearchCode.Enabled = True
DcbLeaderID.Enabled = True
TxtLeaderName.Enabled = False
TxtLeaderName.text = ""
ElseIf dcbLeaderType(1).value = True Then
TxtSearchCode.Enabled = False
DcbLeaderID.Enabled = False
TxtLeaderName.Enabled = True
DcbLeaderID.BoundText = 0
TxtSearchCode.text = ""
End If
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub optNotAuthorization_Click()
    If optNotAuthorization Then
        txtDateCancel.Visible = True
        lbl(77).Visible = True
    Else
        txtDateCancel.Visible = False
        lbl(77).Visible = False
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text2.text, EmpID
        DcbLeaderID.BoundText = EmpID
    End If
End Sub

Private Sub TxtBoardNO_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , , txtBoardNo.text, 2
End If
End Sub

Private Sub txtLetter1_KeyPress(KeyAscii As Integer)
txtLetter1.text = ""
If Len(txtLetter1.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        txtLetter2.SetFocus
End Select

End Sub
Sub EmptyTxt()
Me.txtNum1.text = ""
Me.txtNum2.text = ""
Me.txtNum3.text = ""
Me.txtNum4.text = ""
Me.txtLetter1.text = ""
Me.txtLetter2.text = ""
Me.txtLetter3.text = ""
Me.txtLetter4.text = ""
End Sub
Private Sub txtLetter1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtLetter2_KeyPress(KeyAscii As Integer)
txtLetter2.text = ""
If Len(txtLetter2.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter3.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.text = ""
If Len(txtLetter3.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter4.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.text = ""
If Len(txtLetter4.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.text = ""
If Len(txtNum1.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum2.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.text = ""
If Len(txtNum2.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum3.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.text = ""
If Len(txtNum3.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum4.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.text = ""
If Len(txtNum4.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
Cal_Board

End Sub

Private Sub txtNum4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board

End Sub

Private Sub txtNum4_LostFocus()
txtBoardNo.SetFocus
End Sub

Private Sub TxtOperatorN_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , TxtOperatorN.text, , 1
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
        FrmEmployeeSearch.lbltype = 11
        Set FrmEmployeeSearch.RetrunFrm = Me

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
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth
        
          WriteCustomerBalPublic Account_code2, Balance
          
  lbl(22).Caption = val(Balance)

          WriteCustomerBalPublic Account_code, Balance
          
  lbl(21).Caption = val(Balance)
  lbl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
        DBIssueDate.value = IssueDate
        DcboEmpDepartments.BoundText = DepID
        DcboSpecifications.BoundText = gradeID
        DcboJobsType.BoundText = JobTypeID
        lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
        
    'End If

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
    Dim str As String
    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
    With Me.Fg
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
    
    Set Dcombos = New ClsDataCombos
    
    Dim RsOptions As ADODB.Recordset
    Set RsOptions = New ADODB.Recordset
    Dim optionStr As String
    optionStr = "select * from TblOptions"
    RsOptions.Open optionStr, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If RsOptions("ShowDriverOnly") = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            str = "  select   e.Emp_ID Emp_ID , e.Emp_Name   Emp_Name  from TblEmployee e, TblEmpJobsTypes  j"
        Else
            str = "  select   e.Emp_ID Emp_ID , e.Emp_NameE   Emp_NameE  from TblEmployee e, TblEmpJobsTypes  j"
        End If
        str = str & "   Where e.JobTypeID = j.JobTypeID"
        str = str & "     and  ( j.JobTypeName like '%”«∆ﬁ%'  or j.JobTypeNamee like '%driver%')"
       
        fill_combo DcbLeaderID, str
    Else
        Dcombos.GetEmployees Me.DcbLeaderID
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    str = " select id , Model  from TblCarModels"
    Else
    str = " select id , ModelE  from TblCarModels"
    End If
    fill_combo Me.DcbModel, str

If SystemOptions.UserInterface = ArabicInterface Then
    str = " select id , name from tblcolor "
    Else
    str = " select id , namee from tblcolor "
    End If

fill_combo Me.DcbColor, str
    
    
    
  Set Dcombos = New ClsDataCombos
   
   
   Dcombos.GetEmpDepartments DcbDept3
   
  
  
  
     
    
    str = "SELECT DISTINCT KafelName, KafelName AS KafelNames"
    str = str & " From dbo.TblEmployee"
    str = str & " WHERE     (NOT (KafelName IS NULL)) "
    fill_combo DcbKafelName, str
           
           
    
    Dcombos.GetTblCarsDataGroup DcbTypeEqup
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetOutType Me.DcOutType
    Dcombos.GetEquipments DcbEquepment
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetBranches Me.dcBranch

    Dcombos.GetEmpDepartments Me.DcboEmpDepartments
    
    Dcombos.GetNewDwpartMent DcbDepartment2
    
    Dcombos.GetEmpJobsTypes Me.DcboJobsType

    Dcombos.GetEmpGrades Me.DcboSpecifications
    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
    YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEmpPassOver2     Order By AdvanceID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.text = "R"
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
    Cmd(9).Caption = "Print"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Accredit.Caption = "Send To Approv."
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    Me.Caption = "Driving Letter "
    lbl(75).Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(39).Caption = "Branch"
    
    Frame5.Caption = "Delegate Data"
    lbl(61).Caption = "Name"
    lbl(63).Caption = "Account No."
    
    Frame6.Caption = "Delegate For Data"
    dcbLeaderType(0).Caption = "Employee"
    dcbLeaderType(0).RightToLeft = False
    dcbLeaderType(1).Caption = "Not Employee"
    dcbLeaderType(1).RightToLeft = False
    lbl(60).Caption = "Nationality"
    lbl(64).Caption = "ID No"
    
    Frame7.Caption = "Vehicle Data"
    lbl(102).Caption = "Name"
    lbl(66).Caption = "Car plate No."
    lbl(73).Caption = "Vehicle Type"
    lbl(65).Caption = "Operational number"
    lbl(71).Caption = "Class"
    lbl(72).Caption = "Color"
    
    lbl(70).Caption = "Notes"
    
    Cmd(11).Caption = "Print Delegation"
    
    XPTab301.Caption = "Basic Data"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"



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
  
    End If
   
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = " ’—ÌÕ Œ—ÊÃ „ƒﬁ "
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
            '        Me.Caption = " ’—ÌÕ Œ—ÊÃ „ƒﬁ ( ÃœÌœ )"
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
            '        Me.Caption = " ’—ÌÕ Œ—ÊÃ „ƒﬁ (  ⁄œÌ· )"
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
EmptyTxt
    XPTxtID.text = IIf(IsNull(rs("AdvanceID").value), "", val(rs("AdvanceID").value))
    XPDtbTrans.value = IIf(IsNull(rs("AdvanceDate").value), Date, rs("AdvanceDate").value)
    txtDateCancel.value = IIf(IsNull(rs("DateCancel").value), Date, rs("DateCancel").value)
    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
  
  DcOutType.BoundText = IIf(IsNull(rs("OutTypeID").value), "", rs("OutTypeID").value)
  DcboEmpDepartments.BoundText = IIf(IsNull(rs("DeparmentID").value), "", rs("DeparmentID").value)
  DcboJobsType.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
 txtremark.text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
  Me.DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    TxtInterval.text = IIf(IsNull(rs("Interval").value), "", rs("Interval").value)
  '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
  
  Me.DcbDepartment2.BoundText = IIf(IsNull(rs("DeptID2").value), "", rs("DeptID2").value)
   TxtExpectedouttime.value = IIf(IsNull(rs("Expectedouttime").value), "", rs("Expectedouttime").value)
 txtExpectedIntime.value = IIf(IsNull(rs("ExpectedIntime").value), "", rs("ExpectedIntime").value)
 txtActualouttime.value = IIf(IsNull(rs("Actualouttime").value), "", rs("Actualouttime").value)
 txtActualIntime.value = IIf(IsNull(rs("ActualIntime").value), "", rs("ActualIntime").value)

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
   Me.DcbDept3.BoundText = IIf(IsNull(rs("DeparmentID2").value), "", rs("DeparmentID2").value)
   Me.DcbDepartment2.BoundText = IIf(IsNull(rs("DeptID2").value), "", rs("DeptID2").value)
  Me.DcbKafelName.text = IIf(IsNull(rs("KafelName").value), "", rs("KafelName").value)
   
    
 ''//////////////
   Me.DcbColor.BoundText = IIf(IsNull(rs("ColorID").value), "", rs("ColorID").value)
   Me.DcbModel.BoundText = IIf(IsNull(rs("ModelID").value), "", rs("ModelID").value)
   Me.TxtOperatorN.text = IIf(IsNull(rs("OperatorN").value), "", rs("OperatorN").value)
   Me.txtBoardNo.text = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
   Me.TxtRemarks2.text = IIf(IsNull(rs("Remarks2").value), "", rs("Remarks2").value)
   Me.TxtNumID.text = IIf(IsNull(rs("NumID").value), "", rs("NumID").value)
   Me.TxtNationality.text = IIf(IsNull(rs("Nationality").value), "", rs("Nationality").value)
   Me.TxtLeaderName.text = IIf(IsNull(rs("LeaderName").value), "", rs("LeaderName").value)
   Me.TxtName.text = IIf(IsNull(rs("Name").value), "", rs("Name").value)
   Me.TxtComputerNo.text = IIf(IsNull(rs("ComputerNo").value), "", rs("ComputerNo").value)
   DcbLeaderID.BoundText = IIf(IsNull(rs("DcbLeaderID").value), "", rs("DcbLeaderID").value)
   Me.DcbTypeEqup.BoundText = IIf(IsNull(rs("TypeEqupID").value), "", rs("TypeEqupID").value)
   Me.DcbEquepment.BoundText = IIf(IsNull(rs("EquepmentID").value), "", rs("EquepmentID").value)
  If Not IsNull(rs("EmpType").value) Then
  If rs("EmpType").value = 1 Then
  dcbLeaderType(1).value = True
  Else
  dcbLeaderType(0).value = True
  End If
  Else
  dcbLeaderType(0).value = True
  End If
  
  
  If Not IsNull(rs("IsAuthorization").value) Then
        If rs("IsAuthorization").value = 1 Then
            optAuthorization.value = True
        Else
            optNotAuthorization.value = True
        End If
  Else
        optNotAuthorization.value = True
  End If
  
  
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
      '  If Me.DcboEmpName.BoundText = "" Then
      '      Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
      '      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      '      DcboEmpName.SetFocus
      '     'salah SendKeys "{F4}"
      '      Exit Sub
      '  End If

 
'
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("TblEmpPassOver2", "AdvanceID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
   '         StrSQL = "Delete From TblEmpPassOver2Details Where AdvanceID=" & val(Me.XPTxtID.text)
   '         Cn.Execute StrSQL, , adExecuteNoRecords

        End If

        rs("branch_no").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
 
        rs("AdvanceID").value = val(XPTxtID.text)
        rs("AdvanceDate").value = XPDtbTrans.value
        rs("DateCancel").value = txtDateCancel.value
        
        rs("Emp_ID").value = val(Me.DcboEmpName.BoundText)
     rs("DeparmentID").value = val(Me.DcboEmpDepartments.BoundText)
   rs("JobTypeID").value = val(Me.DcboJobsType.BoundText)
   rs("Remark").value = IIf(txtremark.text = "", Null, (txtremark.text))
  rs("interval").value = IIf(TxtInterval.text = "", Null, val(TxtInterval.text))
   rs("UserID").value = Me.DCboUserName.BoundText
    rs("OutTypeID").value = val(Me.DcOutType.BoundText)
 
    rs("Expectedouttime").value = TxtExpectedouttime.value
   rs("ExpectedIntime").value = txtExpectedIntime.value
   rs("Actualouttime").value = txtActualouttime.value
   rs("ActualIntime").value = txtActualIntime.value
 ''/////26 07 2016
rs("Remarks2").value = IIf(TxtRemarks2.text = "", Null, (TxtRemarks2.text))
rs("ComputerNo").value = IIf(TxtComputerNo.text = "", Null, (TxtComputerNo.text))
rs("Name").value = IIf(TxtName.text = "", Null, (TxtName.text))
rs("LeaderName").value = IIf(TxtLeaderName.text = "", Null, (TxtLeaderName.text))
rs("Nationality").value = IIf(TxtNationality.text = "", Null, (TxtNationality.text))
rs("DeptID2").value = val(Me.DcbDepartment2.BoundText)
rs("DcbLeaderID").value = IIf(val(DcbLeaderID.BoundText) = 0, Null, val(DcbLeaderID.BoundText))

rs("DeparmentID2") = val(Me.DcbDept3.BoundText)
 
 rs("DeptID2") = val(DcbDepartment2.BoundText)
 rs("KafelName").value = Trim(Me.DcbKafelName.text)


rs("NumID").value = IIf(TxtNumID.text = "", Null, (TxtNumID.text))
rs("BoardNO").value = IIf(txtBoardNo.text = "", Null, (txtBoardNo.text))
rs("OperatorN").value = IIf(TxtOperatorN.text = "", Null, (TxtOperatorN.text))
rs("DcbLeaderID").value = IIf(val(DcbLeaderID.BoundText) = 0, Null, val(DcbLeaderID.BoundText))
rs("ModelID").value = IIf(val(Me.DcbModel.BoundText) = 0, Null, val(DcbModel.BoundText))
rs("ColorID").value = IIf(val(Me.DcbColor.BoundText) = 0, Null, val(DcbColor.BoundText))
rs("TypeEqupID").value = IIf(val(Me.DcbTypeEqup.BoundText) = 0, Null, val(DcbTypeEqup.BoundText))
rs("EquepmentID").value = IIf(val(Me.DcbEquepment.BoundText) = 0, Null, val(DcbEquepment.BoundText))
If dcbLeaderType(1).value = True Then
rs("EmpType").value = 1
Else
rs("EmpType").value = Null
End If
         
If optAuthorization.value = True Then
    rs("IsAuthorization").value = 1
Else
    rs("IsAuthorization").value = Null
End If

         
         rs.update
        Cn.CommitTrans
        BeginTrans = False
    
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Data was saved  " & CHR(13)
                    Msg = Msg + "Do you want to add another record "
                End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox "Data was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Else
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
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
            Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì…  " & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Are you sure you want to delete this record"
        End If
            
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "This operation is not available due to lack of records"
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
        Msg = "Sorry , somthing went wrong while deleting data" & CHR(13)
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

Private Sub Cal_Board()
    txtBoardNo.text = txtLetter1.text & " " & txtLetter2.text & " " & txtLetter3.text & " " & txtLetter4.text & " " & txtNum1.text & " " & txtNum2.text & " " & txtNum3.text & " " & txtNum4.text
    RetriveCarsInfo , , txtBoardNo.text, 2
End Sub
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
        .Create Me.hWnd, " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ›ÊÌ÷ »ﬁÌ«œ… ”Ì«—…", 1, 15204351, -2147483630
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
        Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ ÕÌÀ «‰Â ﬁ»·  «—ÌŒ «·ÌÊ„ ÊÂ–« ·« ÌÃÊ“...!!!"
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
        Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·”·›…...!!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtInterval.SetFocus
        Exit Function
    End If

    If val(TxtPaymentCounts.text) = 0 Then
        Msg = "ÌÃ» «œŒ«· ⁄œœ „—«   ”œÌœ «·œ›⁄…...!!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtInterval.SetFocus
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

    If CheckPartCal = False Then
        Exit Sub
    End If

    If CheckDate = False Then
        Exit Sub
    End If

    SngPartValue = val(Me.TxtInterval.text) / val(Me.TxtPaymentCounts.text)
    IntPartCounts = val(Me.TxtPaymentCounts.text)
    m_FirstDate = CDate(val(Me.CboYear.text) & "-" & Me.CmbMonth.ListIndex + 1 & "-01")

    With Me.Fg
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

  Sub LoadDept()
     Dim Dcombos As New ClsDataCombos
     Dcombos.ClearMyDataCombo DcbDepartment2
     Dcombos.GetNewDwpartMent DcbDepartment2, True, val(DcbDept3.BoundText)
  End Sub

