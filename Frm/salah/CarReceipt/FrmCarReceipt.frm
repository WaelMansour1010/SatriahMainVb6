VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmCarReceipt 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13155
   Icon            =   "FrmCarReceipt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   13155
   Begin VB.TextBox TxtYear 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   115
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox TxtMonth 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   113
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox TxtDay 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   111
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox TxtSearchCode 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10680
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   1095
      Width           =   1335
   End
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
      Left            =   10680
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
      Width           =   13125
      _cx             =   23151
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
      Caption         =   "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…"
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCarReceipt.frx":038A
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCarReceipt.frx":0724
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
         ButtonImage     =   "FrmCarReceipt.frx":0ABE
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCarReceipt.frx":0E58
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
         Left            =   6240
         Picture         =   "FrmCarReceipt.frx":11F2
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
         TabIndex        =   33
         Top             =   0
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   8340
      TabIndex        =   7
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   224460801
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   2790
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9060
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
         Left            =   6375
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
         Left            =   5535
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
         Left            =   4680
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
         Left            =   3825
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
         Left            =   0
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
         Left            =   855
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
         Left            =   2760
         TabIndex        =   26
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
         TabIndex        =   36
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
      Left            =   8580
      TabIndex        =   16
      Top             =   8520
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
      Bindings        =   "FrmCarReceipt.frx":4E5A
      Height          =   315
      Left            =   3120
      TabIndex        =   30
      Top             =   1080
      Width           =   3375
      _ExtentX        =   5953
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
      Height          =   6975
      Left            =   240
      TabIndex        =   37
      Top             =   1440
      Width           =   12840
      _cx             =   22648
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
      Caption         =   "»Ì«‰«  «” ·«„ «·„⁄œÂ/«·”Ì«—…|Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "FrmCarReceipt.frx":4E6F
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6510
         Left            =   13485
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   45
         Width           =   12750
         _cx             =   22490
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
            Height          =   3630
            Left            =   120
            TabIndex        =   39
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
            FormatString    =   $"FrmCarReceipt.frx":5209
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
            TabIndex        =   51
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
            TabIndex        =   40
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6510
         Index           =   15
         Left            =   45
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   45
         Width           =   12750
         _cx             =   22490
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
         _GridInfo       =   $"FrmCarReceipt.frx":5355
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6480
            Index           =   16
            Left            =   15
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   15
            Width           =   12720
            _cx             =   22437
            _cy             =   11430
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
            Begin XtremeSuiteControls.GroupBox lblpart 
               Height          =   4305
               Left            =   0
               TabIndex        =   62
               Top             =   1665
               Width           =   12720
               _Version        =   786432
               _ExtentX        =   22437
               _ExtentY        =   7594
               _StockProps     =   79
               Caption         =   "√Ã“«¡ «·›Õ’ «·Œ«—ÃÌ "
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
               Begin VB.TextBox TxtMechanicalFaults 
                  Alignment       =   2  'Center
                  Height          =   1185
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   101
                  Top             =   3000
                  Width           =   4935
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H00E2E9E9&
                  Height          =   975
                  Left            =   6480
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   2880
                  Width           =   6255
                  Begin XtremeSuiteControls.CheckBox ChLicensePl 
                     Height          =   375
                     Left            =   4320
                     TabIndex        =   94
                     Top             =   480
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«··ÊÕ« "
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChReflecto 
                     Height          =   375
                     Left            =   2760
                     TabIndex        =   95
                     Top             =   480
                     Width           =   1455
                     _Version        =   786432
                     _ExtentX        =   2566
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·„—«Ì« «·⁄«ﬂ”…"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChWashers 
                     Height          =   375
                     Left            =   1440
                     TabIndex        =   96
                     Top             =   480
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "„«¡ «·„”«Õ« "
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChRemote 
                     Height          =   375
                     Left            =   120
                     TabIndex        =   97
                     Top             =   480
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "—Ì„Ê  ﬂ‰ —Ê·"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.Frame Frame6 
                  BackColor       =   &H00E2E9E9&
                  Height          =   975
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   1920
                  Width           =   6255
                  Begin XtremeSuiteControls.CheckBox ChBumper 
                     Height          =   375
                     Left            =   4320
                     TabIndex        =   89
                     Top             =   480
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·’œ«„« "
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChFireExt 
                     Height          =   375
                     Left            =   3000
                     TabIndex        =   90
                     Top             =   480
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "ÿ›«Ì… Õ—Ìﬁ"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChSeatB 
                     Height          =   375
                     Left            =   1560
                     TabIndex        =   91
                     Top             =   480
                     Width           =   1095
                     _Version        =   786432
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "√Õ“„… «·«„«‰"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChReserveT 
                     Height          =   375
                     Left            =   120
                     TabIndex        =   92
                     Top             =   480
                     Width           =   1335
                     _Version        =   786432
                     _ExtentX        =   2355
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·ﬂ›— «·«Õ Ì«ÿÌ"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.Frame Frame5 
                  BackColor       =   &H00E2E9E9&
                  Height          =   975
                  Left            =   6480
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   1920
                  Width           =   6255
                  Begin XtremeSuiteControls.CheckBox ChRecRad 
                     Height          =   375
                     Left            =   4200
                     TabIndex        =   84
                     Top             =   480
                     Width           =   1455
                     _Version        =   786432
                     _ExtentX        =   2566
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·„”Ã·/«·—«œÌÊ"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChWipers 
                     Height          =   375
                     Left            =   2760
                     TabIndex        =   85
                     Top             =   480
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·„”«Õ« "
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChTyres 
                     Height          =   375
                     Left            =   1560
                     TabIndex        =   86
                     Top             =   480
                     Width           =   1095
                     _Version        =   786432
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·«ÿ«—« "
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChParkingB 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   87
                     Top             =   480
                     Width           =   1095
                     _Version        =   786432
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "›—«„· «·Ìœ"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.Frame Frame4 
                  BackColor       =   &H00E2E9E9&
                  Height          =   975
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   1080
                  Width           =   6255
                  Begin XtremeSuiteControls.CheckBox ChFrontSeat 
                     Height          =   375
                     Left            =   3360
                     TabIndex        =   80
                     Top             =   480
                     Width           =   2175
                     _Version        =   786432
                     _ExtentX        =   3836
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·„ﬁ«⁄œ «·«„«„Ì…"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChBackSeat 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   81
                     Top             =   480
                     Width           =   1935
                     _Version        =   786432
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·„ﬁ«⁄œ «·Œ·›Ì…"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„ﬁ«⁄œ"
                     Height          =   285
                     Index           =   18
                     Left            =   2640
                     TabIndex        =   82
                     Top             =   120
                     Width           =   525
                  End
               End
               Begin VB.Frame Frame3 
                  BackColor       =   &H00E2E9E9&
                  Height          =   975
                  Left            =   6480
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   1080
                  Width           =   6255
                  Begin XtremeSuiteControls.CheckBox ChWindScreen 
                     Height          =   375
                     Left            =   3960
                     TabIndex        =   75
                     Top             =   480
                     Width           =   1575
                     _Version        =   786432
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·“Ã«Ã «·«„«„Ì"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChBackVew 
                     Height          =   375
                     Left            =   2040
                     TabIndex        =   76
                     Top             =   480
                     Width           =   1455
                     _Version        =   786432
                     _ExtentX        =   2566
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·“Ã«Ã «·Œ·›Ì"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChRearViewMirror 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   77
                     Top             =   480
                     Width           =   1695
                     _Version        =   786432
                     _ExtentX        =   2990
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "“Ã«Ã «·«»Ê«» «·Ã«‰»Ì…"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·“Ã«Ã"
                     Height          =   285
                     Index           =   11
                     Left            =   3000
                     TabIndex        =   78
                     Top             =   120
                     Width           =   645
                  End
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H00E2E9E9&
                  Height          =   975
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   240
                  Width           =   6255
                  Begin XtremeSuiteControls.CheckBox ChFlasher 
                     Height          =   375
                     Left            =   4440
                     TabIndex        =   70
                     Top             =   480
                     Width           =   1095
                     _Version        =   786432
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·›·‘—"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChFront 
                     Height          =   375
                     Left            =   2760
                     TabIndex        =   71
                     Top             =   480
                     Width           =   1335
                     _Version        =   786432
                     _ExtentX        =   2355
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·€„«“«  «·«„«„Ì…"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChBack 
                     Height          =   375
                     Left            =   360
                     TabIndex        =   72
                     Top             =   480
                     Width           =   1815
                     _Version        =   786432
                     _ExtentX        =   3201
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·€„«“«  «·Œ·›Ì…"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·€„«“« "
                     Height          =   285
                     Index           =   5
                     Left            =   2640
                     TabIndex        =   73
                     Top             =   120
                     Width           =   645
                  End
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E2E9E9&
                  Height          =   975
                  Left            =   6480
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   240
                  Width           =   6255
                  Begin XtremeSuiteControls.CheckBox ChHeadL 
                     Height          =   375
                     Left            =   4320
                     TabIndex        =   64
                     Top             =   480
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«·«‰Ê«— «·√„«„Ì…"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChTailL 
                     Height          =   255
                     Left            =   2880
                     TabIndex        =   65
                     Top             =   480
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "«·«‰Ê«— «·Œ·›Ì…"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChBackUpL 
                     Height          =   375
                     Left            =   1560
                     TabIndex        =   66
                     Top             =   480
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«‰Ê«— «·—ÌÊ”"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.CheckBox ChBrakeL 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   67
                     Top             =   480
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "«‰Ê«— «·›—«„·"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«‰Ê«—"
                     Height          =   285
                     Index           =   9
                     Left            =   3000
                     TabIndex        =   68
                     Top             =   120
                     Width           =   645
                  End
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·√⁄ÿ«· «·„Ìﬂ«‰ÌﬂÌ…"
                  Height          =   285
                  Index           =   2
                  Left            =   5040
                  TabIndex        =   100
                  Top             =   3480
                  Width           =   1365
               End
            End
            Begin XtremeSuiteControls.GroupBox lblData 
               Height          =   2025
               Left            =   0
               TabIndex        =   52
               Top             =   0
               Width           =   12720
               _Version        =   786432
               _ExtentX        =   22437
               _ExtentY        =   3572
               _StockProps     =   79
               Caption         =   "»Ì«‰«  «·„⁄œÂ/«·”Ì«—…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
               Begin VB.TextBox TxtTachometerReading 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  TabIndex        =   99
                  Top             =   1320
                  Width           =   4335
               End
               Begin VB.TextBox TxtGeneralShape 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   6600
                  TabIndex        =   98
                  Top             =   1320
                  Width           =   4695
               End
               Begin VB.TextBox TxtChassisNo 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   6600
                  TabIndex        =   61
                  Top             =   960
                  Width           =   4695
               End
               Begin VB.TextBox TxtPlateNo 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   6600
                  TabIndex        =   60
                  Top             =   600
                  Width           =   4695
               End
               Begin MSDataListLib.DataCombo DcbCarType 
                  Bindings        =   "FrmCarReceipt.frx":5389
                  Height          =   315
                  Left            =   120
                  TabIndex        =   53
                  Top             =   240
                  Width           =   4335
                  _ExtentX        =   7646
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
               Begin MSDataListLib.DataCombo DcbCarModel 
                  Bindings        =   "FrmCarReceipt.frx":539E
                  Height          =   315
                  Left            =   120
                  TabIndex        =   54
                  Top             =   600
                  Width           =   4335
                  _ExtentX        =   7646
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
               Begin MSDataListLib.DataCombo DcbColor 
                  Bindings        =   "FrmCarReceipt.frx":53B3
                  Height          =   315
                  Left            =   120
                  TabIndex        =   55
                  Top             =   960
                  Width           =   4335
                  _ExtentX        =   7646
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
               Begin MSDataListLib.DataCombo DcbProject 
                  Bindings        =   "FrmCarReceipt.frx":53C8
                  Height          =   315
                  Left            =   6600
                  TabIndex        =   59
                  Top             =   240
                  Width           =   4695
                  _ExtentX        =   8281
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
                  Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
                  Height          =   285
                  Index           =   12
                  Left            =   5160
                  TabIndex        =   109
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„ÊœÌ· «·„⁄œÂ/«·”Ì«—…"
                  Height          =   285
                  Index           =   13
                  Left            =   5160
                  TabIndex        =   108
                  Top             =   600
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "·Ê‰ «·„⁄œÂ/«·”Ì«—…"
                  Height          =   285
                  Index           =   14
                  Left            =   5160
                  TabIndex        =   107
                  Top             =   960
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «··ÊÕ…"
                  Height          =   285
                  Index           =   16
                  Left            =   11640
                  TabIndex        =   106
                  Top             =   600
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Êﬁ⁄ «·⁄„·"
                  Height          =   285
                  Index           =   17
                  Left            =   11640
                  TabIndex        =   105
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ÂÌﬂ·"
                  Height          =   285
                  Index           =   10
                  Left            =   11640
                  TabIndex        =   104
                  Top             =   960
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‘ﬂ· «·⁄«„"
                  Height          =   285
                  Index           =   19
                  Left            =   11400
                  TabIndex        =   103
                  Top             =   1320
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁ—«¡… «·⁄œ«œ"
                  Height          =   285
                  Index           =   20
                  Left            =   4560
                  TabIndex        =   102
                  Top             =   1320
                  Width           =   1605
               End
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   315
               Left            =   1200
               TabIndex        =   50
               Top             =   8415
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   556
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
               Height          =   4395
               Index           =   62
               Left            =   2430
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   1965
               Width           =   570
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6480
            Index           =   9
            Left            =   15
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   15
            Width           =   12720
            _cx             =   22437
            _cy             =   11430
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
               Height          =   4860
               Left            =   3315
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1365
               Width           =   690
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   3420
               Left            =   4185
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   1695
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3420
               Index           =   67
               Left            =   2370
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   1695
               Width           =   630
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   3240
               Index           =   68
               Left            =   4005
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   2175
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
               Height          =   3900
               Index           =   69
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   1695
               Width           =   315
            End
         End
      End
   End
   Begin MSDataListLib.DataCombo DcboEmpName 
      Height          =   315
      Left            =   7320
      TabIndex        =   57
      Top             =   1080
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label XPTxtCount1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   118
      Top             =   8760
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   -1
      TabIndex        =   117
      Top             =   -35640
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   -120
      TabIndex        =   116
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   114
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   112
      Top             =   720
      Width           =   135
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " √—ÌŒ ≈‰ Â«¡ «·«” „«—…"
      Height          =   285
      Index           =   3
      Left            =   5400
      TabIndex        =   110
      Top             =   735
      Width           =   1845
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” ·„"
      Height          =   285
      Index           =   15
      Left            =   12030
      TabIndex        =   58
      Top             =   1095
      Width           =   1005
   End
   Begin VB.Label LblBr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1140
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
      TabIndex        =   29
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   285
      Index           =   4
      Left            =   12030
      TabIndex        =   25
      Top             =   750
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   9510
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
      Left            =   11325
      TabIndex        =   23
      Top             =   8715
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2430
      TabIndex        =   22
      Top             =   8790
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   690
      TabIndex        =   21
      Top             =   8790
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   -5790
      TabIndex        =   20
      Top             =   9060
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   19
      Top             =   8820
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
Attribute VB_Name = "FrmCarReceipt"
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
Dim Date_to_Str As String
Dim str_to_date  As Date
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
Sub Ch()
Me.ChHeadL.value = vbUnchecked
Me.ChTailL.value = vbUnchecked
Me.ChBackUpL.value = vbUnchecked
Me.ChBrakeL.value = vbUnchecked
Me.ChFlasher.value = vbUnchecked
Me.ChFront.value = vbUnchecked
Me.ChBack.value = vbUnchecked
Me.ChWindScreen.value = vbUnchecked
Me.ChBackVew.value = vbUnchecked
Me.ChRearViewMirror.value = vbUnchecked
Me.ChFrontSeat.value = vbUnchecked
Me.ChBackSeat.value = vbUnchecked
Me.ChRecRad.value = vbUnchecked
Me.ChWipers.value = vbUnchecked
Me.ChTyres.value = vbUnchecked
Me.ChParkingB.value = vbUnchecked
Me.ChBumper.value = vbUnchecked
Me.ChFireExt.value = vbUnchecked
Me.ChSeatB.value = vbUnchecked
Me.ChReserveT.value = vbUnchecked
Me.ChLicensePl.value = vbUnchecked
Me.ChReflecto.value = vbUnchecked
Me.ChWashers.value = vbUnchecked
Me.ChRemote.value = vbUnchecked
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
        Ch
            
              GRID2.Clear flexClearScrollable, flexClearEverything
    GRID2.rows = 1
            Me.DCboUserName.BoundText = user_id
    '        TxtPaymentCounts.text = 1
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
            Load FrmCarReceptSearch
            FrmCarReceptSearch.show

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 8
        '    CalCulateParts
            
            
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


MySQL = " SELECT     dbo.TblCarReceipt.ID, dbo.TblCarReceipt.RecordDate, dbo.TblCarReceipt.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
        MySQL = MySQL & "              dbo.TblCarReceipt.EmpID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
               MySQL = MySQL & "         dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
             MySQL = MySQL & "           dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.EmpGroupDep.GroupName, dbo.TblCarReceipt.ProjectID, dbo.TblCarReceipt.Type,"
              MySQL = MySQL & "          dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblCarReceipt.Mod, dbo.TblCarModels.Model, dbo.TblCarReceipt.Colour, dbo.TblColor.name AS Color,"
                 MySQL = MySQL & "       dbo.TblColor.namee AS Colore, dbo.TblCarReceipt.PlateNo, dbo.TblCarReceipt.ChassisNo, dbo.TblCarReceipt.UserID, dbo.TblCarReceipt.GeneralShape,"
                  MySQL = MySQL & "      dbo.TblCarReceipt.TachometerReading, dbo.TblCarReceipt.TailL, dbo.TblCarReceipt.BackUpL, dbo.TblCarReceipt.HeadL, dbo.TblCarReceipt.BrakeL,"
                    MySQL = MySQL & "    dbo.TblCarReceipt.Flasher, dbo.TblCarReceipt.Front, dbo.TblCarReceipt.Back, dbo.TblCarReceipt.WindScreen, dbo.TblCarReceipt.BackVew,"
                    MySQL = MySQL & "    dbo.TblCarReceipt.RearViewMirror, dbo.TblCarReceipt.FrontSeat, dbo.TblCarReceipt.BackSeat, dbo.TblCarReceipt.RegRad, dbo.TblCarReceipt.Wipers,"
                     MySQL = MySQL & "   dbo.TblCarReceipt.Tyres, dbo.TblCarReceipt.Bumper, dbo.TblCarReceipt.ParkingB, dbo.TblCarReceipt.FireExt, dbo.TblCarReceipt.SeatB, dbo.TblCarReceipt.ReserveT,"
                     MySQL = MySQL & "   dbo.TblCarReceipt.LicensePl, dbo.TblCarReceipt.Reflecto, dbo.TblCarReceipt.Washers, dbo.TblCarReceipt.MechanicalFaults, dbo.TblCarReceipt.Remote,"
               MySQL = MySQL & "         dbo.TblCarReceipt.DateExp"
                      MySQL = MySQL & "   FROM         dbo.TblCarReceipt LEFT OUTER JOIN"
             MySQL = MySQL & "           dbo.TblColor ON dbo.TblCarReceipt.Colour = dbo.TblColor.Id LEFT OUTER JOIN"
                MySQL = MySQL & "        dbo.TblCarModels ON dbo.TblCarReceipt.Mod = dbo.TblCarModels.Id LEFT OUTER JOIN"
                 MySQL = MySQL & "       dbo.TBLCarTypes ON dbo.TblCarReceipt.Type = dbo.TBLCarTypes.id LEFT OUTER JOIN"
                   MySQL = MySQL & "     dbo.TblEmployee ON dbo.TblCarReceipt.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
                    MySQL = MySQL & "    dbo.TblBranchesData ON dbo.TblCarReceipt.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
                    MySQL = MySQL & "    dbo.EmpGroupDep ON dbo.TblCarReceipt.ProjectID = dbo.EmpGroupDep.GroupID"
                  MySQL = MySQL & "   Where (dbo.TblCarReceipt.id =" & val(XPTxtID.text) & ")"
 
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarReceipt.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCarReceipt.rpt"
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
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
    '    xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  xReport.ParameterFields(12).AddCurrentValue Date_to_Str
   
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



Private Sub DcbCarType_Click(Area As Integer)
Dim Dcombos As ClsDataCombos
      Set Dcombos = New ClsDataCombos
    
      If Me.DcbCarType.SelectedItem <> 0 Then
   Dcombos.GetTblCarModels Me.DcbCarModel, , Me.DcbCarType.SelectedItem
   End If
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DtExp_Change()
 
End Sub
Sub ConvertDateToHijri(ByVal dt As Date)
Dim TempDate As String
Dim lnth As Integer
Dim moth, dY, Yar As String
moth = Month(dt)
lnth = Len(moth)
If lnth = 1 Then
moth = "0" & moth
End If
dY = day(dt)
lnth = Len(dY)
If lnth = 1 Then
dY = "0" & dY
End If
Yar = year(dt)
lnth = Len(Yar)
If lnth = 1 Then
Yar = "0" & Yar
End If
Me.TxtMonth.text = moth
Me.TxtDay.text = dY
Me.Txtyear.text = Yar
 End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub TxtDay_Change()
If val(Me.TxtDay.text) > 31 Then
MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
Me.TxtDay.text = ""
Me.TxtDay.SetFocus
Exit Sub
End If
End Sub
Private Sub TxtMonth_Change()
If val(Me.TxtMonth.text) > 12 Then
MsgBox "Œÿ«¡ ›Ì «œŒ«· «·‘Â—"
Me.TxtMonth.text = ""
Me.TxtMonth.SetFocus
Exit Sub
End If

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub
Public Function TransDate(TheDate As Date, TypeTrans As Integer) As String
Dim TempDate As String, MD As Date, a As String
 If TypeTrans = 1 Then
VBA.Calendar = vbCalHijri
TempDate = CStr(TheDate)
TransDate = TempDate
VBA.Calendar = vbCalGreg
'Text1 = TransDate
Else
a = CStr(TheDate)
VBA.Calendar = vbCalHijri
MD = CDate(a)
VBA.Calendar = vbCalGreg
TransDate = CStr(Format(MD, "yyyy/mm/dd"))
'txtdateofenglish = TransDate
End If

End Function
Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 3
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
        Dim Project As Integer
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth, , , , Project
        
     '     WriteCustomerBalPublic Account_code2, Balance
          
 ' lbl(22).Caption = val(Balance)

     '     WriteCustomerBalPublic Account_Code, Balance
          
 ' lbl(21).Caption = val(Balance)
  'lbl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
        'DBIssueDate.value = issuedate
       ' DcboEmpDepartments.BoundText = depid
      '  DcboSpecifications.BoundText = gradeID
      '  DcboJobsType.BoundText = JobTypeID
       ' lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
        Me.DcbProject.BoundText = Project
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


    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

    'With Me.Fg
    '    .RowHeightMin = 300
    '    .WallPaper = GrdBack.Picture
    '    .AutoSize 0, .Cols - 1, False
    'End With

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
   Dcombos.GetEmpLocations Me.DcbProject
  Dcombos.GetTblColor Me.DcbColor
   Dcombos.GetTblCarsDataGroup Me.DcbCarType
    Dcombos.GetTblCarModels Me.DcbCarModel
    'Dcombos.GetEmpJobsTypes Me.DcboJobsType

    'Dcombos.GetEmpGrades Me.DcboSpecifications
    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
  '  YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblCarReceipt     Order By ID"
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
   ' Label1.Visible = False

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Prient"
    Cmd(6).Caption = "Exit"
    Cmd(0).Caption = "Prient"
    CmdHelp.Caption = "Help"
XPTab301.Caption = "Data of Car"
    Me.Caption = " Car Receipt   "
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lblBr.Caption = "Bramch"
    lbl(15).Caption = "Recipient'sName"
    lbldata.Caption = "Data of car"
    lbl(3).Caption = "ExpDate Form"
    lbl(17).Caption = "Location"
    lbl(12).Caption = "Type"
    lbl(13).Caption = "Mod"
    lbl(14).Caption = "Colour"
    lbl(20).Caption = "Tachometer Reading"
    lbl(16).Caption = "Plate No"
    lbl(10).Caption = "Chassis No"
    lbl(19).Caption = "General Shape"
    lblpart.Caption = "L-Visual Inspection"
    lbl(2).Caption = "Mechanical Faults"
    lbl(9).Caption = "H-Light"
    lbl(5).Caption = "Indicator"
    lbl(11).Caption = "Mirror"
    lbl(18).Caption = "Seat"
    Me.ChHeadL.RightToLeft = False
    Me.ChHeadL.Caption = "Head L"
    Me.ChTailL.RightToLeft = False
    Me.ChTailL.Caption = "Tail L"
    Me.ChBackUpL.RightToLeft = False
    Me.ChBackUpL.Caption = "Back-Up L"
    Me.ChBrakeL.RightToLeft = False
    Me.ChBrakeL.Caption = "Brake L"
    Me.ChFlasher.RightToLeft = False
    Me.ChFlasher.Caption = "Flasher"
    Me.ChFront.Caption = "Front"
    Me.ChFront.RightToLeft = False
    Me.ChBack.RightToLeft = False
    Me.ChBack.Caption = "Back"
    Me.ChWindScreen.RightToLeft = False
    Me.ChWindScreen.Caption = "WindScreen"
    Me.ChBackVew.RightToLeft = False
    Me.ChBackVew.Caption = "Back Vew"
    Me.ChRearViewMirror.RightToLeft = False
    Me.ChRearViewMirror.Caption = "Rear View Mirror"
    Me.ChFrontSeat.RightToLeft = False
    Me.ChFrontSeat.Caption = "Front Seat"
    Me.ChBackSeat.RightToLeft = False
    Me.ChBackSeat.Caption = "Back Seat"
    Me.ChRecRad.RightToLeft = False
    Me.ChRecRad.Caption = "Rec/Rad"
    Me.ChWipers.RightToLeft = False
    Me.ChWipers.Caption = "Wipers"
    Me.ChTyres.RightToLeft = False
    Me.ChTyres.Caption = "Tyres"
    Me.ChParkingB.RightToLeft = False
    Me.ChParkingB.Caption = "Parking B"
    Me.ChBumper.RightToLeft = False
    Me.ChBumper.Caption = "Bumper"
    Me.ChFireExt.RightToLeft = False
    Me.ChFireExt.Caption = "Fire Ext"
    Me.ChSeatB.RightToLeft = False
    Me.ChSeatB.Caption = "Seat B"
    Me.ChReserveT.RightToLeft = False
    Me.ChReserveT.Caption = "Reserve T"
    Me.ChLicensePl.RightToLeft = False
    Me.ChLicensePl.Caption = "License PL"
    Me.ChReflecto.RightToLeft = False
    Me.ChReflecto.Caption = "Reflec To"
    Me.ChWashers.RightToLeft = False
    Me.ChWashers.Caption = "Washers"
    Me.ChRemote.RightToLeft = False
    Me.ChRemote.Caption = "Remote Control"
 '   Fra(0).Caption = "payments Method"
 '   lbl(9).Caption = "Count"
 '   lbl(10).Caption = "Start"
 '   lbl(11).Caption = "Month"
 '   lbl(12).Caption = "Year"
 '   Cmd(8).Caption = "Calc Dates"
 '   ChkSaleryDis.Caption = "Auto Discount"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"

    'With Me.Fg
        '.TextMatrix(0, .ColIndex("PartNO")) = "NO"
      '  .TextMatrix(0, .ColIndex("PartValue")) = "Value"
     '   .TextMatrix(0, .ColIndex("PartDate")) = "Date"

  '  End With

End Sub

'Private Sub YearMonth()

 '   Dim i As Integer
   ' Dim IntDefIndex As Integer

   ' CmbMonth.Clear

 '   For i = 1 To 12
   '     CmbMonth.AddItem MonthName(i)
  '  Next

   ' CmbMonth.ListIndex = Month(Date) - 1
    'CboYear.Clear

   ' For i = 2010 To 2050
     '   CboYear.AddItem i

      '  If i = year(Date) Then
       '     IntDefIndex = CboYear.NewIndex
        'End If

'    Next

  '  CboYear.ListIndex = IntDefIndex
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

Private Sub TxtAdvanceValue_LostFocus()
    
   
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…"
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
            '        Me.Caption = "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…( ÃœÌœ )"
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
            '        Me.Caption = "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…(  ⁄œÌ· )"
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
         '   TxtAdvanceValue.Locked = False
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
Ch
    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount1.Caption = 0
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

    XPTxtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Date_to_Str = IIf(IsNull(rs("DateExp").value), Null, rs("DateExp").value)
 
    Me.ConvertDateToHijri Format(Date_to_Str, "yyyy/M/d")
    
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    DcbProject.BoundText = val(IIf(IsNull(rs("ProjectID").value), 0, rs("ProjectID").value))
    DcbCarType.BoundText = IIf(IsNull(rs("Type").value), "", rs("Type").value)
    DcbCarModel.BoundText = IIf(IsNull(rs("Mod").value), "", rs("Mod").value)
    DcbColor.BoundText = IIf(IsNull(rs("Colour").value), "", rs("Colour").value)
    TxtPlateNo.text = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
    TxtChassisNo.text = IIf(IsNull(rs("ChassisNo").value), "", rs("ChassisNo").value)
    TxtGeneralShape.text = IIf(IsNull(rs("GeneralShape").value), "", rs("GeneralShape").value)
    TxtTachometerReading.text = IIf(IsNull(rs("TachometerReading").value), "", rs("TachometerReading").value)
    TxtMechanicalFaults.text = IIf(IsNull(rs("MechanicalFaults").value), "", rs("MechanicalFaults").value)
    If rs("HeadL").value = True Then
    Me.ChHeadL.value = vbChecked
    Else
    Me.ChHeadL.value = vbUnchecked
    End If
    If rs("TailL").value = True Then
    Me.ChTailL.value = vbChecked
    Else
    Me.ChTailL.value = vbUnchecked
    End If
    If rs("BackUpL").value = True Then
    Me.ChBackUpL.value = vbChecked
    Else
    Me.ChBackUpL.value = vbUnchecked
    End If
    If rs("BrakeL").value = True Then
    Me.ChBrakeL.value = vbChecked
    Else
    Me.ChBrakeL.value = vbUnchecked
    End If
    If rs("Flasher").value = True Then
    Me.ChFlasher.value = vbChecked
    Else
    Me.ChFlasher.value = vbUnchecked
    End If
    If rs("Front").value = True Then
    Me.ChFront.value = vbChecked
      Else
      Me.ChFront.value = vbUnchecked
    End If
    If rs("Back").value = True Then
    Me.ChBack.value = vbChecked
    Else
    Me.ChBack.value = vbUnchecked
    End If
    If rs("WindScreen").value = True Then
    Me.ChWindScreen.value = vbChecked
    Else
    Me.ChWindScreen.value = vbUnchecked
    End If
    If rs("BackVew").value = True Then
    Me.ChBackVew.value = vbChecked
    Else
    Me.ChBackVew.value = vbUnchecked
    End If
    If rs("RearViewMirror").value = True Then
    Me.ChRearViewMirror.value = vbChecked
    Else
    Me.ChRearViewMirror.value = vbUnchecked
    End If
    If rs("FrontSeat").value = True Then
    Me.ChFrontSeat.value = vbChecked
     Else
     Me.ChFrontSeat.value = vbUnchecked
    End If
    If rs("BackSeat").value = True Then
    Me.ChBackSeat.value = vbChecked
    Else
    Me.ChBackSeat.value = vbUnchecked
    End If
    If rs("RegRad").value = True Then
    Me.ChRecRad.value = vbChecked
    Else
    Me.ChRecRad.value = vbUnchecked
    End If
    If rs("Wipers").value = True Then
    Me.ChWipers.value = vbChecked
    Else
    Me.ChWipers.value = vbUnchecked
    End If
    If rs("Tyres").value = True Then
    Me.ChTyres.value = vbChecked
    Else
    Me.ChTyres.value = vbUnchecked
    End If
    If rs("ParkingB").value = True Then
    Me.ChParkingB.value = vbChecked
    Else
    Me.ChParkingB.value = vbUnchecked
    End If
    If rs("Bumper").value = True Then
    Me.ChBumper.value = vbChecked
    Else
    Me.ChBumper.value = vbUnchecked
    End If
   
    If rs("FireExt").value = True Then
    Me.ChFireExt.value = vbChecked
    Else
    Me.ChFireExt.value = vbUnchecked
    End If
    If rs("SeatB").value = True Then
    Me.ChSeatB.value = vbChecked
    Else
    Me.ChSeatB.value = vbUnchecked
    End If
    If rs("ReserveT").value = True Then
    Me.ChReserveT.value = vbChecked
    Else
    Me.ChReserveT.value = vbUnchecked
    End If
    If rs("LicensePl").value = True Then
    Me.ChLicensePl.value = vbChecked
    Else
    Me.ChLicensePl.value = vbUnchecked
    End If
    If rs("Reflecto").value = True Then
    Me.ChReflecto.value = vbChecked
    Else
    Me.ChReflecto.value = vbUnchecked
    End If
    If rs("Washers").value = True Then
    Me.ChWashers.value = vbChecked
    Else
    Me.ChWashers.value = vbUnchecked
    End If
    If rs("Remote").value = True Then
    Me.ChRemote.value = vbChecked
    Else
    Me.ChRemote.value = vbUnchecked
    End If
    
       ' DcboEmpDepartments.BoundText = IIf(IsNull(rs("DeparmentID").value), "", rs("DeparmentID").value)

   ' DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)

    'DcboJobsType.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)

 '  lbl(23).Caption = IIf(IsNull(rs("basicSalary").value), "", rs("basicSalary").value)
 
 '   lbl(22).Caption = IIf(IsNull(rs("EmpDue").value), "", rs("EmpDue").value)
 '  lbl(20).Caption = IIf(IsNull(rs("Contractvalid").value), "", rs("Contractvalid").value)
 '  lbl(21).Caption = IIf(IsNull(rs("oldAdvance").value), "", rs("oldAdvance").value)
 '
'TxtDiscount.text = IIf(IsNull(rs("Discount").value), "", rs("Discount").value)
'txtDiscountDES.text = IIf(IsNull(rs("DiscountDES").value), "", rs("DiscountDES").value)

 

'    Me.DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
'    TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
  '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
'    Me.TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), "", rs("PaymentCounts").value)
 
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
   
   
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = "Select * From  TblEmpAdvanceRequestDetails Where AdvanceID=" & val(XPTxtID.text)
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    Fg.Clear flexClearScrollable, flexClearEverything
'    Fg.Rows = Fg.FixedRows
'
'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        Fg.Rows = Fg.FixedRows + RsDetails.RecordCount
'
'        For i = Me.Fg.FixedRows To Fg.Rows - 1
'            Fg.TextMatrix(i, Fg.ColIndex("PartNO")) = RsDetails("PartNO").value
'            Fg.TextMatrix(i, Fg.ColIndex("PartValue")) = RsDetails("PartValue").value
'            Fg.TextMatrix(i, Fg.ColIndex("PartDate")) = DisplayDate(CDate(RsDetails("PartDate").value))
'            RsDetails.MoveNext
'        Next i
'
'    End If

'    RsDetails.Close
'    Set RsDetails = Nothing
    
    fillapprovData
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount1.Caption = rs.RecordCount
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
Dim temp As Date
    'On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.DcboEmpName.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If

   

       ' If CheckPartCal = False Then
      '      Exit Sub
      '  End If

        'If CheckDate = False Then
         '   Exit Sub
       ' End If

        '”·› ”«»ﬁ…
       ' Dim RsTest As New ADODB.Recordset
       ' 'Set RsTest = New ADODB.Recordset
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

       ' CalCulateParts
    
 
        
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

            XPTxtID.text = CStr(new_id("TblCarReceipt", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
          '  StrSQL = "Delete From TblEmpAdvanceRequestDetails Where AdvanceID=" & val(Me.XPTxtID.text)
          '  Cn.Execute StrSQL, , adExecuteNoRecords

        End If
        Date_to_Str = Me.Txtyear.text
Date_to_Str = Date_to_Str & "/"
Date_to_Str = Date_to_Str & Me.TxtMonth.text
Date_to_Str = Date_to_Str & "/"
Date_to_Str = Date_to_Str & Me.TxtDay.text
   
      '  temp = CDate(Format(Date_to_Str, "DD/MM/YYY "))
        rs("ID").value = val(XPTxtID.text)
        rs("RecordDate").value = XPDtbTrans.value
        rs("DateExp").value = Date_to_Str
        rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
         rs("EmpID").value = IIf(Me.DcboEmpName.BoundText = "", Null, Me.DcboEmpName.BoundText)
          rs("ProjectID").value = IIf(Me.DcbProject.BoundText = "", Null, Me.DcbProject.BoundText) '
        rs("Type").value = IIf(Me.DcbCarType.BoundText = "", Null, Me.DcbCarType.BoundText)         '
       rs("PlateNo").value = TxtPlateNo.text
       rs("Mod").value = IIf(Me.DcbCarModel.BoundText = "", Null, Me.DcbCarModel.BoundText)
        rs("Colour").value = IIf(Me.DcbColor.BoundText = "", Null, Me.DcbColor.BoundText)
        rs("ChassisNo").value = TxtChassisNo.text
        rs("GeneralShape").value = TxtGeneralShape.text
        rs("TachometerReading").value = TxtTachometerReading.text
        rs("MechanicalFaults").value = Me.TxtMechanicalFaults.text
        If Me.ChHeadL.value = vbChecked Then
         rs("HeadL").value = 1
         Else
         rs("HeadL").value = 0
         End If
         If Me.ChTailL.value = vbChecked Then
          rs("TailL").value = 1
          Else
        rs("TailL").value = 0
        End If
        If Me.ChBackUpL.value = vbChecked Then
          rs("BackUpL").value = 1
        Else
        rs("BackUpL").value = 0
        End If
        If Me.ChBrakeL.value = vbChecked Then
         rs("BrakeL").value = 1
        Else
        rs("BrakeL").value = 0
        End If
        If Me.ChFlasher.value = vbChecked Then
         rs("Flasher").value = 1
        Else
        rs("Flasher").value = 0
        End If
        If Me.ChFront.value = vbChecked Then
          rs("Front").value = 1
        Else
        rs("Front").value = 0
        End If
        If Me.ChBack.value = vbChecked Then
         rs("Back").value = 1
        Else
        rs("Back").value = 0
        End If
        If Me.ChWindScreen.value = vbChecked Then
         rs("WindScreen").value = 1
        Else
        rs("WindScreen").value = 0
        End If
        If Me.ChBackVew.value = vbChecked Then
         rs("BackVew").value = 1
        Else
        rs("BackVew").value = 0
        End If
        If Me.ChRearViewMirror.value = vbChecked Then
         rs("RearViewMirror").value = 1
        Else
        rs("RearViewMirror").value = 0
        End If
        If Me.ChFrontSeat.value = vbChecked Then
         rs("FrontSeat").value = 1
        Else
        rs("FrontSeat").value = 0
        End If
        If Me.ChBackSeat.value = vbChecked Then
        rs("BackSeat").value = 1
        Else
        rs("BackSeat").value = 0
        End If
        If Me.ChRecRad.value = vbChecked Then
        rs("RegRad").value = 1
        Else
        rs("RegRad").value = 0
        End If
        If Me.ChWipers.value = vbChecked Then
        rs("Wipers").value = 1
        Else
        rs("Wipers").value = 0
        End If
        If Me.ChTyres.value = vbChecked Then
        rs("Tyres").value = 1
        Else
        rs("Tyres").value = 0
        End If
        If Me.ChParkingB.value = vbChecked Then
         rs("ParkingB").value = 1
        Else
        rs("ParkingB").value = 0
        End If
        If Me.ChBumper.value = vbChecked Then
        rs("Bumper").value = 1
        Else
        rs("Bumper").value = 0
        End If
       If Me.ChFireExt.value = vbChecked Then
       rs("FireExt").value = 1
        Else
        rs("FireExt").value = 0
        End If
        If Me.ChSeatB.value = vbChecked Then
        rs("SeatB").value = 1
        Else
        rs("SeatB").value = 0
        End If
        If Me.ChReserveT.value = vbChecked Then
         rs("ReserveT").value = 1
        Else
        rs("ReserveT").value = 0
        End If
        If Me.ChLicensePl.value = vbChecked Then
         rs("LicensePl").value = 1
        Else
        rs("LicensePl").value = 0
        End If
        If Me.ChReflecto.value = vbChecked Then
        rs("Reflecto").value = 1
        Else
        rs("Reflecto").value = 0
        End If
        If Me.ChWashers.value = xtpChecked Then
        rs("Washers").value = 1
        Else
        rs("Washers").value = 0
        End If
        If Me.ChRemote.value = xtpChecked Then
        rs("Remote").value = 1
        Else
        rs("Remote").value = 0
        End If
      '  rs("ManagerID").value = Me.DcmbManagerID.BoundText
      '  rs("JobID").value = val(Me.DcboJobsType.BoundText)
      '  rs("FirstDate").value = IIf(IsDate(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartDate"))), Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartDate")), Null)
      '  rs("PaymentCounts").value = val(Me.TxtPaymentCounts.text)
        rs("UserID").value = Me.DCboUserName.BoundText

        rs.update
     '   Set RsDetails = New ADODB.Recordset
     '   RsDetails.Open "TblEmp", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

     '   For i = Me.Fg.FixedRows To Fg.Rows - 1
     '       RsDetails.AddNew
     '       RsDetails("AdvanceID").value = val(XPTxtID.text)
     '       RsDetails("PartNO").value = Fg.TextMatrix(i, Fg.ColIndex("PartNO"))
     '       RsDetails("PartValue").value = Fg.TextMatrix(i, Fg.ColIndex("PartValue"))
     '       RsDetails("PartDate").value = Fg.TextMatrix(i, Fg.ColIndex("PartDate"))
     '       RsDetails.update
     '   Next i
    
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
     '   RsDetails.Close
     '   Set RsDetails = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount1.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

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
                    XPTxtCount1.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Ch
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
        .Create Me.hWnd, "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «” ·«„ «·„⁄œÂ/«·”Ì«—…", 1, 15204351, -2147483630
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
   
End Sub

Private Function CheckDate() As Boolean
 
End Function

'Private Function CheckPartCal() As Boolean
   ' Dim Msg As String

   ' CheckPartCal = False

   ' If val(TxtAdvanceValue.text) = 0 Then
      '  Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·”·›…...!!! "
       '' MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        'TxtAdvanceValue.SetFocus
      '  Exit Function
   ' End If

    'If val(TxtPaymentCounts.text) = 0 Then
      '  Msg = "ÌÃ» «œŒ«· ⁄œœ „—«   ”œÌœ «·œ›⁄…...!!! "
       ' MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
       ' TxtAdvanceValue.SetFocus
       ' Exit Function
    'End If

   ' If CmbMonth.ListIndex = -1 Then
  '      Msg = "ÌÃ»  ÕœÌœ «Ê· ‘Â— · ”œÌœ «·œ›⁄…..!! "
      '  MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
      '  CmbMonth.SetFocus
     '   SendKeys "{F4}"
       ' Exit Function
   ' End If

   ' If CboYear.ListIndex = -1 Then
     '   Msg = "ÌÃ»  ÕœÌœ «Ê· ”‰… · ”œÌœ «·œ›⁄… ..!! "
     '   MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
     '   CboYear.SetFocus
     '   SendKeys "{F4}"
      '  Exit Function
  '  End If

   ' CheckPartCal = True'
'End Function

'Private Sub CalCulateParts()
 '   Dim i As Integer
 '   Dim IntPartCounts As Integer
 '   Dim SngPartValue As Single
  '  Dim m_FirstDate As Date
'
  '  If CheckPartCal = False Then
 '       Exit Sub
  '  End If
'
   ' If CheckDate = False Then
   '     Exit Sub
 '   End If
'
   ' SngPartValue = val(Me.TxtAdvanceValue.text) / val(Me.TxtPaymentCounts.text)
 '   IntPartCounts = val(Me.TxtPaymentCounts.text)
  '  m_FirstDate = CDate(val(Me.CboYear.text) & "-" &   Me.CmbMonth.ListIndex + 1 & "-01"  )

   ' With Me.Fg
    '    .Clear flexClearScrollable, flexClearEverything
    '    .Rows = .FixedRows + IntPartCounts
     '   .RowHeightMin = 300

       ' For i = 1 To IntPartCounts
        '    .TextMatrix(i, .ColIndex("PartNO")) = i
         '   .TextMatrix(i, .ColIndex("PartValue")) = SngPartValue
         '   .TextMatrix(i, .ColIndex("PartDate")) = DisplayDate(DateAdd("m", i - 1, m_FirstDate))
    '    Next i
    
   ' End With

'End Sub

