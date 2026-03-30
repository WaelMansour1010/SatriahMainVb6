VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmDistriExpensItems 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«· ﬂ«·Ì› «· ﬁœÌ—Ì… ÿ»ﬁ« ··«’‰«›"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12840
   Icon            =   "FrmDistriExpensItems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   12840
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   13920
      TabIndex        =   34
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   13440
      TabIndex        =   33
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
      TabIndex        =   31
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
      Enabled         =   0   'False
      Height          =   315
      Left            =   10800
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
      Left            =   -360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   13155
      _cx             =   23204
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
      Caption         =   "«· ﬂ«·Ì› «· ﬁœÌ—Ì… ÿ»ﬁ« ··«’‰«› "
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
         Left            =   1425
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
         ButtonImage     =   "FrmDistriExpensItems.frx":038A
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
         Left            =   360
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
         ButtonImage     =   "FrmDistriExpensItems.frx":0724
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
         Left            =   1950
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
         ButtonImage     =   "FrmDistriExpensItems.frx":0ABE
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
         Left            =   885
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
         ButtonImage     =   "FrmDistriExpensItems.frx":0E58
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
         TabIndex        =   32
         Top             =   480
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   8220
      TabIndex        =   7
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
      Left            =   1230
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7260
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
      Left            =   9120
      TabIndex        =   16
      Top             =   6840
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
      Bindings        =   "FrmDistriExpensItems.frx":11F2
      Height          =   315
      Left            =   840
      TabIndex        =   30
      Top             =   720
      Width           =   5295
      _ExtentX        =   9340
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
      Height          =   5535
      Left            =   0
      TabIndex        =   37
      Top             =   1080
      Width           =   12840
      _cx             =   22648
      _cy             =   9763
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
      Caption         =   "Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "FrmDistriExpensItems.frx":1207
      Flags(0)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5070
         Left            =   13485
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   45
         Width           =   12750
         _cx             =   22490
         _cy             =   8943
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
            FormatString    =   $"FrmDistriExpensItems.frx":15A1
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
            TabIndex        =   40
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   5070
         Index           =   15
         Left            =   45
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   45
         Width           =   12750
         _cx             =   22490
         _cy             =   8943
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
         _GridInfo       =   $"FrmDistriExpensItems.frx":16ED
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5040
            Index           =   16
            Left            =   15
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   15
            Width           =   12720
            _cx             =   22437
            _cy             =   8890
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
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Height          =   2115
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   2760
               Width           =   12720
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   21
                  Left            =   120
                  TabIndex        =   64
                  Top             =   2520
                  Width           =   840
                  _ExtentX        =   1482
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
                  ButtonImage     =   "FrmDistriExpensItems.frx":1721
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid fg 
                  Height          =   1635
                  Left            =   0
                  TabIndex        =   65
                  Top             =   120
                  Width           =   12720
                  _cx             =   22437
                  _cy             =   2884
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
                  Rows            =   1
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmDistriExpensItems.frx":1CBB
                  ScrollTrack     =   0   'False
                  ScrollBars      =   2
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   12
                  Left            =   11760
                  TabIndex        =   70
                  Top             =   1680
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmDistriExpensItems.frx":1E11
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   13
                  Left            =   9960
                  TabIndex        =   71
                  Top             =   1680
                  Width           =   1650
                  _ExtentX        =   2910
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·ﬂ·"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmDistriExpensItems.frx":23AB
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin VB.Frame Frame11 
               Height          =   1785
               Left            =   4365
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   480
               Width           =   8085
               Begin VB.ListBox ListGroupSelected 
                  Height          =   1425
                  ItemData        =   "FrmDistriExpensItems.frx":2945
                  Left            =   120
                  List            =   "FrmDistriExpensItems.frx":294C
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   240
                  Width           =   3255
               End
               Begin VB.ListBox ListGroupAll 
                  Height          =   1425
                  ItemData        =   "FrmDistriExpensItems.frx":2963
                  Left            =   4200
                  List            =   "FrmDistriExpensItems.frx":296A
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   240
                  Width           =   3255
               End
               Begin VB.Label Label5 
                  Alignment       =   2  'Center
                  Caption         =   "<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   1080
                  Width           =   495
               End
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  Caption         =   "<<"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   840
                  Width           =   495
               End
               Begin VB.Label Label7 
                  Alignment       =   2  'Center
                  Caption         =   ">>"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   600
                  Width           =   495
               End
               Begin VB.Label Label8 
                  Alignment       =   2  'Center
                  Caption         =   ">"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   360
                  Width           =   495
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õœœ «·«’‰«›"
               Height          =   2925
               Index           =   11
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   0
               Width           =   12720
               Begin VB.TextBox TxtRemark 
                  Alignment       =   1  'Right Justify
                  Height          =   1605
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   600
                  Width           =   4095
               End
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6960
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   2400
                  Width           =   1455
               End
               Begin VB.CommandButton BtonAdd 
                  Caption         =   "«÷«›…"
                  Height          =   255
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   2520
                  Width           =   2055
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Œ «— «·„ÃÊ⁄Â „ÕœœÂ"
                  Height          =   210
                  Index           =   1
                  Left            =   8280
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   240
                  Width           =   4065
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’‰› „Õœœ  ≈Œ «— «·’‰›"
                  Height          =   210
                  Index           =   2
                  Left            =   7680
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   2400
                  Width           =   2865
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂ· «·«’‰«›"
                  Height          =   210
                  Index           =   0
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   2400
                  Value           =   -1  'True
                  Width           =   2385
               End
               Begin MSDataListLib.DataCombo DcItem1 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   72
                  Top             =   2400
                  Width           =   4695
                  _ExtentX        =   8281
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„Ã„Ê⁄«  «·„Œ «—…"
                  Height          =   285
                  Index           =   5
                  Left            =   4680
                  TabIndex        =   76
                  Top             =   240
                  Width           =   2205
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„·«ÕŸ« "
                  Height          =   285
                  Index           =   2
                  Left            =   360
                  TabIndex        =   75
                  Top             =   240
                  Width           =   2205
               End
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   510
               Left            =   0
               TabIndex        =   49
               Top             =   4680
               Visible         =   0   'False
               Width           =   1830
               _ExtentX        =   3228
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   8
               Left            =   0
               TabIndex        =   67
               Top             =   15000
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   688
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
               ButtonImage     =   "FrmDistriExpensItems.frx":297C
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   10
               Left            =   0
               TabIndex        =   68
               Top             =   -3720
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   688
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
               ButtonImage     =   "FrmDistriExpensItems.frx":2F16
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   11
               Left            =   -120
               TabIndex        =   69
               Top             =   33960
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   688
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
               ButtonImage     =   "FrmDistriExpensItems.frx":34B0
               DrawFocusRectangle=   0   'False
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5040
            Index           =   9
            Left            =   15
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   15
            Width           =   12720
            _cx             =   22437
            _cy             =   8890
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
               Height          =   3780
               Left            =   3345
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   1005
               Width           =   690
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   2670
               Left            =   4215
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   1365
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2670
               Index           =   67
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1365
               Width           =   570
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   2520
               Index           =   68
               Left            =   4035
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   1635
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
               Height          =   3000
               Index           =   69
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1365
               Width           =   375
            End
         End
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„"
      Height          =   285
      Index           =   3
      Left            =   11520
      TabIndex        =   55
      Top             =   720
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
      TabIndex        =   35
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
      Left            =   12810
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   285
      Index           =   4
      Left            =   6480
      TabIndex        =   25
      Top             =   720
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   9270
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
      Left            =   11805
      TabIndex        =   23
      Top             =   6915
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
      Top             =   6990
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
      Top             =   6990
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   20
      Top             =   6900
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   19
      Top             =   6900
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
Attribute VB_Name = "FrmDistriExpensItems"
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
Public LngCol As Double
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
'  '  Retrive (val(Me.XPTxtID.text))
'End Sub




Private Sub BtonAdd_Click()
If XPOptShowType(0).value = False And XPOptShowType(1).value = False And XPOptShowType(2).value = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "ÌÃ» «Œ Ì«— ‰Ê⁄ «·⁄„·ÌÂ «Ê·«"
Else
MsgBox "Please Select Type of Operation"
End If
Exit Sub
End If
Retrivetitems
End Sub
Sub Retrivetitems()
  Dim I As Integer
  Dim j As Integer
  Dim k As Integer
 
  Dim Msg As String
  Dim bool As Boolean
  Dim Rs1 As ADODB.Recordset
  Dim Sql As String
  bool = True
  
        Fg.Enabled = True
   
 
  With Fg
   If XPOptShowType(2).value = True Then
  If val(DcItem1.BoundText) = 0 Then
  If SystemOptions.UserInterface = ArabicInterface Then
  MsgBox "Ì—ÃÏ «Œ Ì«— «·’‰›"
  Else
  MsgBox "Please Select Item"
  End If
  DcItem1.SetFocus
  Exit Sub
  End If
   j = .Rows
.Rows = .Rows + 1
 Set Rs1 = New ADODB.Recordset
        For I = j To .Rows - 1
     
  Sql = " SELECT     dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, "
           Sql = Sql & "           dbo.TblUnites.UnitNamee , dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode, dbo.Groups.GroupNamee"
           Sql = Sql & "            FROM         dbo.Groups RIGHT OUTER JOIN"
           Sql = Sql & "            dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID LEFT OUTER JOIN"
           Sql = Sql & "            dbo.TblUnites RIGHT OUTER JOIN"
           Sql = Sql & "            dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
     Sql = Sql & "  Where (dbo.TblItems.ItemID =" & val(Me.DcItem1.BoundText) & ")"
     Rs1.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs1.RecordCount > 0 Then

              .TextMatrix(I, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), "", Rs1("ItemID").value)
              .TextMatrix(I, .ColIndex("ItemCode")) = IIf(IsNull(Rs1("ItemCode").value), "", Rs1("ItemCode").value)
 
             .TextMatrix(I, .ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
          
       If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
             .TextMatrix(I, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            Else
            .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
            .TextMatrix(I, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
      End If
End If
        
          
          
        Next I
        TxtSearchCode.Text = ""
        DcItem1.Text = ""
            
   End If
       
           If XPOptShowType(0).value = True Then


    Set Rs1 = New ADODB.Recordset
           Sql = " SELECT     dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, "
           Sql = Sql & "           dbo.TblUnites.UnitNamee , dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode, dbo.Groups.GroupNamee"
           Sql = Sql & "            FROM         dbo.Groups RIGHT OUTER JOIN"
           Sql = Sql & "            dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID LEFT OUTER JOIN"
           Sql = Sql & "            dbo.TblUnites RIGHT OUTER JOIN"
           Sql = Sql & "            dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
           Rs1.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then



j = .Rows
.Rows = .Rows + Rs1.RecordCount

        For I = j To .Rows - 1
     
              .TextMatrix(I, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), "", Rs1("ItemID").value)
              .TextMatrix(I, .ColIndex("ItemCode")) = IIf(IsNull(Rs1("ItemCode").value), "", Rs1("ItemCode").value)
 
             .TextMatrix(I, .ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
          
       If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
             .TextMatrix(I, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            Else
            .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
            .TextMatrix(I, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
      End If
        Rs1.MoveNext
        
        
        
        Next I

    End If
       
       
       
        End If
        Dim GROUPIDS As String
        
          If XPOptShowType(1).value = True Then
          

          For k = 1 To ListGroupSelected.ListCount

    Set Rs1 = New ADODB.Recordset
        '   sql = " SELECT * from  TblItems where GroupID =" & ListGroupSelected.ItemData(k - 1) & ""
        GROUPIDS = GetallChilddata(ListGroupSelected.ItemData(k - 1))
        If Len(GROUPIDS) > 2 Then GROUPIDS = Mid(GROUPIDS, 2, Len(GROUPIDS))
        Debug.Print GROUPIDS
        If GROUPIDS = "" Then GROUPIDS = ListGroupSelected.ItemData(k - 1)
          Sql = " SELECT     dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, "
           Sql = Sql & "           dbo.TblUnites.UnitNamee , dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode, dbo.Groups.GroupNamee"
           Sql = Sql & "            FROM         dbo.Groups RIGHT OUTER JOIN"
           Sql = Sql & "            dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID LEFT OUTER JOIN"
           Sql = Sql & "            dbo.TblUnites RIGHT OUTER JOIN"
           Sql = Sql & "            dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
           
        
       Sql = Sql & "  where dbo.TblItems.GroupID IN ( " & GROUPIDS & ")"
        
       '(GetallChilddata
           Rs1.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
 j = .Rows
.Rows = .Rows + Rs1.RecordCount

        For I = j To .Rows - 1
      .TextMatrix(I, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), "", Rs1("ItemID").value)
              .TextMatrix(I, .ColIndex("ItemCode")) = IIf(IsNull(Rs1("ItemCode").value), "", Rs1("ItemCode").value)
 
             .TextMatrix(I, .ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
          
       If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
             .TextMatrix(I, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            Else
            .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
            .TextMatrix(I, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
      End If
        Rs1.MoveNext
        
    
        Next I

    End If
       
       
         Next k
        End If

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
          

ListGroupSelected.Clear
 Fg.Clear flexClearScrollable, flexClearEverything
 Fg.Rows = 1

 XPOptShowType(1).value = False
            TxtModFlg.Text = "N"
            clear_all Me
  
            
         '     GRID2.Clear flexClearScrollable, flexClearEverything
    'GRID2.Rows = 1
            Me.DCboUserName.BoundText = user_id
          '  TxtPaymentCounts.text = 1
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

            TxtModFlg.Text = "E"
            
            Frame11.Enabled = True
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
            Load FrmDestriEpensItemSearch
            FrmDestriEpensItemSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.Text, , 200

        Case 12
         RemoveGridRow
            Case 13
             Fg.Clear flexClearScrollable, flexClearEverything
 Fg.Rows = 1
            
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

MySQL = "SELECT     dbo.TblDistriExpensItem.Remark, dbo.TblDistriExpensItem.RecordeDate, dbo.TblDistriExpensItem.Ind, dbo.TblDistriExpensItem.BranchID, "
MySQL = MySQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblDistriExpensItem.Selected, dbo.TblDistriExpensItemDet2.Account,"
MySQL = MySQL & "                      dbo.TblDistriExpensItemDet2.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblDistriExpensItemDet2.GroupID,"
MySQL = MySQL & "                      dbo.Groups.GroupName, dbo.Groups.GroupCode, dbo.Groups.GroupNamee, dbo.TblDistriExpensItemDet2.ID, dbo.TblDistriExpensItemDet3.IDDet,"
MySQL = MySQL & "                      dbo.TblDistriExpensItemDet3.TypeValue, dbo.TblDistriExpensItemDet2.Ind AS IndD2, dbo.TblDistriExpensItemDet3.Ind AS IndD3, dbo.TblDistriExpensItemDet3.Vlue,"
MySQL = MySQL & "                      dbo.TblDistriExpensItemDet3.Remark AS RemarkD3, dbo.TblDistriExpensItemDet3.Account_Code, dbo.ACCOUNTS.Account_Name,"
MySQL = MySQL & "                      dbo.ACCOUNTS.Account_NameEng"
MySQL = MySQL & " FROM         dbo.TblDistriExpensItemDet3 LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.ACCOUNTS ON dbo.TblDistriExpensItemDet3.Account_Code = dbo.ACCOUNTS.Account_Code RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblDistriExpensItemDet2 ON dbo.TblDistriExpensItemDet3.IDDet = dbo.TblDistriExpensItemDet2.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.Groups ON dbo.TblDistriExpensItemDet2.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblItems ON dbo.TblDistriExpensItemDet2.ItemID = dbo.TblItems.ItemID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblDistriExpensItem ON dbo.TblDistriExpensItemDet2.Ind = dbo.TblDistriExpensItem.Ind LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblDistriExpensItem.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblDistriExpensItem.ind = " & val(XPTxtID.Text) & ")"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDistributeExpensiveItems.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDistributeExpensiveItems.rpt"
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
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
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









Private Sub DcItem1_Change()
     Me.TxtSearchCode.Text = GetItemCode(val(Me.DcItem1.BoundText))
 End Sub

Private Sub DcItem1_Click(Area As Integer)
 DcItem1_Change
End Sub

Private Sub DcItem1_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 27
        FrmItemSearch.show vbModal
    End If
End Sub

Private Sub Fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With Fg
            If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
ReLineGrid
 
    End With
End Sub


Private Sub Fg_CellButtonClick(ByVal Row As Long, ByVal Col As Long)


    'On Error GoTo ErrTrap

    With Me.Fg

        Select Case .ColKey(Col)

                 Case "Account"
                  LngRow = Row

 LngCol = Col
             ' ItemProductionDate Row, Col, , 1
                Load FrmDistriItemAccount
                FrmDistriItemAccount.show vbModal

                    
                End Select
                End With
End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With Me.Fg

        Select Case .ColKey(Col)

                 Case "Account"
    
            .ColComboList(.ColIndex("Account")) = "..."
            End Select
           End With
End Sub





 


Private Sub Label5_Click()

    If ListGroupSelected.ListIndex > -1 Then
        ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
    End If

End Sub
Private Sub Label6_Click()
    ListGroupSelected.Clear
End Sub
Private Sub Label7_Click()
    Dim I As Integer
    If Me.XPOptShowType(1).value = True Then
    ListGroupSelected.Clear

    For I = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(I)
        ListGroupSelected.ItemData(I) = ListGroupAll.ItemData(I)
    Next I
End If
End Sub
Private Sub Label8_Click()
If Me.XPOptShowType(1).value = True Then
 If ListGroupAll.ListIndex > -1 Then
    ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
             
    ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)
            End If
            End If
End Sub



Private Sub ListGroupAll_Click()
 If XPOptShowType(1).value = True Then
        Frame11.Enabled = True
    Else
        Frame11.Enabled = False
    End If
End Sub

Private Sub TxtSearchCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        If TxtSearchCode.Text = "" Then
            Me.DcItem1.BoundText = ""
        Else
            Me.DcItem1.BoundText = GetItemID(Trim$(Me.TxtSearchCode.Text))
        End If
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
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
  'sa Frame10.Enabled = False
    Frame11.Enabled = False

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
    FillMylist
    Set Dcombos = New ClsDataCombos
      Dcombos.GetUsers Me.DCboUserName
  
    Dcombos.GetBranches Me.dcBranch
     Dcombos.GetItemsNames DcItem1, , , , True

   ' Dcombos.GetItemSGroups Me.DCbGroup, False



    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
    'YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblDistriExpensItem     Order By Ind"
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

    Me.Caption = "Distribution expenses on items "
    EleHeader.Caption = Me.Caption
    lbl(3).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(4).Caption = "Branch"
    lbl(5).Caption = "Groups Selected "
    lbl(2).Caption = "Remarks  "
    Fra(11).Caption = "Select Items "
    XPOptShowType(1).Caption = "A specific group chose Group"
    XPOptShowType(2).Caption = "A specific Item chose Item"
    XPOptShowType(0).Caption = "All Items"
    XPOptShowType(1).RightToLeft = False
    XPOptShowType(2).RightToLeft = False
    XPOptShowType(0).RightToLeft = False
   BtonAdd.Caption = "Add"
   Cmd(12).Caption = "Delete"
   Cmd(13).Caption = "Delete All"
   Accredit.Caption = "Accredite"
   XPTab301.Caption = "Distribution expenses on items "

   lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"

   With Me.Fg
        .TextMatrix(0, .ColIndex("serial")) = "NO"
        .TextMatrix(0, .ColIndex("Account")) = "Account"
        .TextMatrix(0, .ColIndex("ItemCode")) = "ItemCode"
         .TextMatrix(0, .ColIndex("ItemID")) = "ItemID"
        .TextMatrix(0, .ColIndex("ItemName")) = "ItemName"
         .TextMatrix(0, .ColIndex("GroupID")) = "GroupID"
        .TextMatrix(0, .ColIndex("GroupName")) = "GroupName"

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
        '    TxtAdvanceValue.Locked = False
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
        '    TxtAdvanceValue.Locked = False
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
    Dim RsDetails1 As ADODB.Recordset
    Dim RsDetails2 As ADODB.Recordset
   
    Frame11.Enabled = False
     
    Dim I As Integer
    Dim StrSQL As String
    ListGroupSelected.Clear

Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 1
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
            rs.find "Ind=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.Text = IIf(IsNull(rs("Ind").value), "", val(rs("Ind").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordeDate").value), Date, rs("RecordeDate").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
     Me.DcItem1.BoundText = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
  Me.txtRemark.Text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    If rs("Selected").value = 0 Then
XPOptShowType(0).value = True
End If
If rs("Selected").value = 1 Then
XPOptShowType(1).value = True
End If
If rs("Selected").value = 2 Then

XPOptShowType(2).value = True

End If

  '    If IsNull(rs("posted").value) Then
  '                                                 If SystemOptions.UserInterface = ArabicInterface Then
  '                                                 Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
  '                                               Else
  '                                                 Accredit.Caption = " send to Approval   "
  '                                             End If
  '                                             Accredit.Enabled = True
  'Else
  '                                                If SystemOptions.UserInterface = ArabicInterface Then
  '                                                  Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
  '                                                Else
  '                                                 Accredit.Caption = " sent to Approval   "
  '                                             End If
  '                                             Accredit.Enabled = False
  ' End If
  ''
   
    Set RsDetails = New ADODB.Recordset
StrSQL = " SELECT     dbo.TblDistriExpensItemDet2.ID, dbo.TblDistriExpensItemDet2.Ind, dbo.TblDistriExpensItemDet2.Account, dbo.TblDistriExpensItemDet2.GroupID,"
StrSQL = StrSQL & "                      dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblDistriExpensItemDet2.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
StrSQL = StrSQL & "                       dbo.TblItems.ItemNamee"
StrSQL = StrSQL & "  FROM         dbo.TblDistriExpensItemDet2 LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblItems ON dbo.TblDistriExpensItemDet2.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.Groups ON dbo.TblDistriExpensItemDet2.GroupID = dbo.Groups.GroupID"
StrSQL = StrSQL & "  Where (dbo.TblDistriExpensItemDet2.ind = " & val(XPTxtID.Text) & ")"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        With Fg
       Fg.Rows = Fg.FixedRows + RsDetails.RecordCount

        For I = Me.Fg.FixedRows To Fg.Rows - 1
             .TextMatrix(I, .ColIndex("serial")) = I
             .TextMatrix(I, .ColIndex("ItemID")) = IIf(IsNull(RsDetails("ItemID").value), "", RsDetails("ItemID").value)
              .TextMatrix(I, .ColIndex("id")) = IIf(IsNull(RsDetails("id").value), "", RsDetails("id").value)
             .TextMatrix(I, .ColIndex("Account1")) = IIf(IsNull(RsDetails("Account").value), "", RsDetails("Account").value)
             .TextMatrix(I, .ColIndex("ItemCode")) = IIf(IsNull(RsDetails("ItemCode").value), "", RsDetails("ItemCode").value)
             .TextMatrix(I, .ColIndex("GroupID")) = IIf(IsNull(RsDetails("GroupID").value), "", RsDetails("GroupID").value)
             .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(RsDetails("GroupName").value), "", RsDetails("GroupName").value)
       If SystemOptions.UserInterface = ArabicInterface Then
             .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(RsDetails("GroupName").value), "", RsDetails("GroupName").value)
             .TextMatrix(I, .ColIndex("ItemName")) = IIf(IsNull(RsDetails("ItemName").value), "", RsDetails("ItemName").value)
            Else
             .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(RsDetails("GroupNamee").value), "", RsDetails("GroupNamee").value)
             .TextMatrix(I, .ColIndex("ItemName")) = IIf(IsNull(RsDetails("ItemNamee").value), "", RsDetails("ItemNamee").value)
      End If
     
            RsDetails.MoveNext
        Next I
End With
    End If
   
  
    
   
       Set RsDetails1 = New ADODB.Recordset
StrSQL = "SELECT     dbo.TblDistriExpensItemDet1.ID, dbo.TblDistriExpensItemDet1.Ind, dbo.TblDistriExpensItemDet1.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee"
StrSQL = StrSQL & " FROM         dbo.TblDistriExpensItemDet1 LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Groups ON dbo.TblDistriExpensItemDet1.GroupID = dbo.Groups.GroupID"
StrSQL = StrSQL & " Where (dbo.TblDistriExpensItemDet1.ind = " & val(XPTxtID.Text) & ")"
RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
  For I = 0 To RsDetails1.RecordCount - 1
  ListGroupSelected.AddItem IIf(IsNull(RsDetails1("GroupName").value), "", RsDetails1("GroupName").value)
  ListGroupSelected.ItemData(I) = IIf(IsNull(RsDetails1("GroupID").value), "", RsDetails1("GroupID").value)
  
   RsDetails1.MoveNext
  
  Next I
  '''''''''''''''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
   
    RsDetails1.Close
    Set RsDetails1 = Nothing
    
     RsDetails.Close
    Set RsDetails = Nothing
   'sa fillapprovData
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Function linebreak(myString As String)
            
            
            If Len(myString) <> 0 Then
                    If right$(myString, 2) = Chr(10) Or right$(myString, 2) = Chr(13) Or right$(myString, 2) = vbCrLf Then
                            linebreak = left$(myString, Len(myString) - 1)
                    Else
                            linebreak = myString
                    End If
            End If
            
            
End Function


Private Sub SaveData()
  Dim astrSplit2tems2() As String
Dim astrSplitItems() As String
  Dim st As String
    Dim nElements As Integer
      Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
     Dim Sql As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
     Dim RsDetails2 As ADODB.Recordset
      
    Dim I As Integer
    Dim j As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap

'sa If Fg.Rows < 2 Then
'sa       If SystemOptions.UserInterface = ArabicInterface Then
'sa            Msg = " ·« ÊÃœ »Ì«‰«  " & Chr(13)
 'sa    Else
 'sa    Msg = "¬Not Found Data " & Chr(13)
 'sa    End If
   'sa         MsgBox Msg, vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 
          'sa  SendKeys "{F4}"
           ' Exit Sub
     'sa   End If

   
      
     

  Dim RsTest As New ADODB.Recordset
    
    

        Cn.BeginTrans
        BeginTrans = True
        If TxtModFlg.Text = "N" Then

            XPTxtID.Text = CStr(new_id("TblDistriExpensItem", "Ind", "", True))
    
        
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
          StrSQL = "Delete From TblDistriExpensItemDet1 Where Ind=" & val(Me.XPTxtID.Text)
           Cn.Execute StrSQL, , adExecuteNoRecords
   StrSQL = "Delete From TblDistriExpensItemDet2 Where Ind=" & val(Me.XPTxtID.Text)
           Cn.Execute StrSQL, , adExecuteNoRecords
              StrSQL = "Delete From TblDistriExpensItemDet3 Where Ind=" & val(Me.XPTxtID.Text)
           Cn.Execute StrSQL, , adExecuteNoRecords
        End If
        rs("Ind").value = val(XPTxtID.Text)
        rs("Remark").value = Me.txtRemark.Text
        rs("RecordeDate").value = XPDtbTrans.value
        rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
         rs("ItemID").value = IIf(Me.DcItem1.BoundText = "", Null, Me.DcItem1.BoundText)
        If XPOptShowType(0).value = True Then
rs("Selected").value = 0
End If
       If XPOptShowType(1).value = True Then
rs("Selected").value = 1
End If
       If XPOptShowType(2).value = True Then
rs("Selected").value = 2
End If
        rs("UserID").value = Me.DCboUserName.BoundText

        rs.update
       
''''''''''''''''''''''''//////////////////

   ''''''''''' /////////////////////////////''''''
   Set RsDetails2 = New ADODB.Recordset
     '   RsDetails2.Open "TblLink_Item_To_Store_Details3", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
     StrSQL = "SELECT     *  from dbo.TblDistriExpensItemDet1 Where (1 = -1)"
   RsDetails2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
             
             
        For I = 0 To ListGroupSelected.ListCount - 1
                  RsDetails2.AddNew
             RsDetails2("Ind").value = val(XPTxtID.Text)
             RsDetails2("GroupID").value = val(ListGroupSelected.ItemData(I))
                      RsDetails2.update
           
    Next I
        '''''''''///////////////////////////////////////////
              Set RsDetails2 = New ADODB.Recordset
     StrSQL = "SELECT     *  from dbo.TblDistriExpensItemDet3 Where (1 = -1)"
   RsDetails2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
           Set RsDetails = New ADODB.Recordset
     StrSQL = "SELECT     *  from dbo.TblDistriExpensItemDet2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
        For I = Me.Fg.FixedRows To Fg.Rows - 1
       If Fg.TextMatrix(I, Fg.ColIndex("ItemName")) <> "" Then
            RsDetails.AddNew
             RsDetails("Ind").value = val(XPTxtID.Text)
               RsDetails("Account").value = Fg.TextMatrix(I, Fg.ColIndex("Account1"))
             RsDetails("ItemID").value = val(Fg.TextMatrix(I, Fg.ColIndex("ItemID")))
             RsDetails("GroupID").value = val(Fg.TextMatrix(I, Fg.ColIndex("GroupID")))
              RsDetails.update
                If Fg.TextMatrix(I, Fg.ColIndex("Account1")) <> "" Then
          st = Fg.TextMatrix(I, Fg.ColIndex("Account1"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For j = 0 To nElements - 1
          RsDetails2.AddNew
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         RsDetails2("Ind").value = val(XPTxtID.Text)
         RsDetails2("IDDet").value = RsDetails("ID").value
         
         RsDetails2("Account_Code").value = (astrSplit2tems2(0))
         RsDetails2("TypeValue").value = astrSplit2tems2(1)
         RsDetails2("Vlue").value = astrSplit2tems2(2)
         RsDetails2("Remark").value = astrSplit2tems2(3)
          RsDetails2.update
         Next j
                  
          
          End If
              
                End If
        Next I
       
       
  
        Cn.CommitTrans
        BeginTrans = False
      
           RsDetails2.Close
         Set RsDetails2 = Nothing
     RsDetails.Close
         Set RsDetails = Nothing
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
            rs.find "Ind='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

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
                StrSQL = "Delete From TblDistriExpensItem Where Ind=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
   
                StrSQL1 = "Delete From TblDistriExpensItemDet1 Where Ind=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL1, , adExecuteNoRecords
               StrSQL1 = "Delete From TblDistriExpensItemDet2 Where Ind=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL1, , adExecuteNoRecords
                 StrSQL1 = "Delete From TblDistriExpensItemDet3 Where Ind=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL1, , adExecuteNoRecords
                If rs.RecordCount < 1 Then
                    clear_all Me
                        ListGroupSelected.Clear
  

                   Fg.Clear flexClearScrollable, flexClearEverything
                   Fg.Rows = 2
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



'Function FillApprovedTable()
' Dim RSApproval  As New ADODB.Recordset
'   Set RSApproval = New ADODB.Recordset
'   Dim currentdate As Date
'   Dim StrSQL As String
'   'RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'     StrSQL = "SELECT     *  from dbo.ApprovalData Where (1 = -1)"
'   RSApproval.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
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
'

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
Private Sub ReLineGrid()
    Dim I As Integer
    Dim IntCounter  As Integer
    
    IntCounter = 0

    With Fg

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("ItemName")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("serial")) = IntCounter
           
    
        End If
                

        Next I
 
    End With

End Sub
Function FillMylist()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim I As Integer
   

  Sql = " SELECT * from  Groups where GroupID>1"
 
    rs.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListGroupAll.Clear
    ListGroupSelected.Clear

    If rs.RecordCount > 0 Then

        For I = 1 To rs.RecordCount

            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupAll.AddItem IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
            Else
                ListGroupAll.AddItem IIf(IsNull(rs("GroupNamee").value), "", rs("GroupNamee").value)
            End If

            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("GroupID").value
            rs.MoveNext
        Next I

    End If

    rs.Close

End Function
Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "   Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
         .Create Me.hWnd, "   Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«› ", 1, 15204351, -2147483630
         .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "   Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   Ê“Ì⁄ «·„’—Ê›«  ⁄·Ï «·«’‰«› ", 1, 15204351, -2147483630
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
       
                'SaveData

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub



 Private Sub RemoveGridRow()

    With Me.Fg

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub
 Private Sub RemoveGridRowSpace()

    With Me.Fg

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub
 
Private Sub XPOptShowType_Click(Index As Integer)
 If XPOptShowType(1).value = True Then
        Frame11.Enabled = True
    Else
        Frame11.Enabled = False
    End If
End Sub

