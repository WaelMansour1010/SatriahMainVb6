VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLinkItemToStore 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "—»ÿ «·«’‰«› »«·„Œ«“‰"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12840
   Icon            =   "FrmLinkIteminStorefrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   12840
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   705
      Width           =   1320
   End
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
      Left            =   10320
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
      Caption         =   "—»ÿ «·«’‰«› »«·„Œ«“‰ "
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
         ButtonImage     =   "FrmLinkIteminStorefrm.frx":038A
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
         ButtonImage     =   "FrmLinkIteminStorefrm.frx":0724
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
         ButtonImage     =   "FrmLinkIteminStorefrm.frx":0ABE
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
         ButtonImage     =   "FrmLinkIteminStorefrm.frx":0E58
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
      Left            =   8340
      TabIndex        =   7
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   190906369
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
      Bindings        =   "FrmLinkIteminStorefrm.frx":11F2
      Height          =   315
      Left            =   3240
      TabIndex        =   30
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
      Height          =   5535
      Left            =   0
      TabIndex        =   37
      Top             =   1200
      Width           =   12720
      _cx             =   22437
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
      Caption         =   "—»ÿ «·«’‰«› »«·„Œ«“‰|Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "FrmLinkIteminStorefrm.frx":1207
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5070
         Left            =   13365
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   45
         Width           =   12630
         _cx             =   22278
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
            FormatString    =   $"FrmLinkIteminStorefrm.frx":15A1
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
         Width           =   12630
         _cx             =   22278
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
         _GridInfo       =   $"FrmLinkIteminStorefrm.frx":16ED
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
            Width           =   12600
            _cx             =   22225
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
               TabIndex        =   71
               Top             =   2760
               Width           =   12600
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   21
                  Left            =   120
                  TabIndex        =   72
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
                  ButtonImage     =   "FrmLinkIteminStorefrm.frx":1721
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8Ctl.VSFlexGrid fg 
                  Height          =   1635
                  Left            =   120
                  TabIndex        =   73
                  Top             =   120
                  Width           =   12360
                  _cx             =   21802
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
                  Cols            =   7
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmLinkIteminStorefrm.frx":1CBB
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
                  TabIndex        =   81
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
                  ButtonImage     =   "FrmLinkIteminStorefrm.frx":1DDB
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   13
                  Left            =   9960
                  TabIndex        =   82
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
                  ButtonImage     =   "FrmLinkIteminStorefrm.frx":2375
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin VB.Frame Frame11 
               Height          =   1545
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   480
               Width           =   5265
               Begin VB.ListBox ListGroupSelected 
                  Height          =   1230
                  ItemData        =   "FrmLinkIteminStorefrm.frx":290F
                  Left            =   120
                  List            =   "FrmLinkIteminStorefrm.frx":2916
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   240
                  Width           =   2295
               End
               Begin VB.ListBox ListGroupAll 
                  Height          =   1230
                  ItemData        =   "FrmLinkIteminStorefrm.frx":292D
                  Left            =   3120
                  List            =   "FrmLinkIteminStorefrm.frx":2934
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   240
                  Width           =   1935
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
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
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
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
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
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
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
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   360
                  Width           =   495
               End
            End
            Begin VB.TextBox TxtRemark 
               Alignment       =   1  'Right Justify
               Height          =   690
               Left            =   5895
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   62
               Top             =   2160
               Width           =   5835
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õœœ «·„Ã„Ê⁄Â"
               Height          =   2925
               Index           =   11
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   0
               Width           =   5625
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   2160
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.CommandButton BtonAdd 
                  Caption         =   "«÷«›…"
                  Height          =   255
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   2520
                  Width           =   2055
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Ã„Ê⁄… „ÕœœÂ ≈Œ «— «·„ÃÊ⁄Â"
                  Height          =   210
                  Index           =   1
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   240
                  Width           =   4065
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’‰› „Õœœ  ≈Œ «— «·’‰›"
                  Height          =   210
                  Index           =   2
                  Left            =   2520
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   2160
                  Visible         =   0   'False
                  Width           =   2865
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂ· «·«’‰«›"
                  Height          =   210
                  Index           =   0
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   2520
                  Value           =   -1  'True
                  Visible         =   0   'False
                  Width           =   2385
               End
               Begin MSDataListLib.DataCombo DcItem1 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   83
                  Top             =   2160
                  Visible         =   0   'False
                  Width           =   2295
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
            End
            Begin VB.Frame Frame10 
               Caption         =   "Õœœ «·„Œ«“‰"
               Height          =   2025
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   60
               Width           =   6735
               Begin VB.ListBox ListStoreall 
                  Height          =   1620
                  ItemData        =   "FrmLinkIteminStorefrm.frx":2946
                  Left            =   3720
                  List            =   "FrmLinkIteminStorefrm.frx":294D
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   240
                  Width           =   2895
               End
               Begin VB.ListBox ListStoreSelected 
                  Height          =   1620
                  ItemData        =   "FrmLinkIteminStorefrm.frx":295F
                  Left            =   120
                  List            =   "FrmLinkIteminStorefrm.frx":2966
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   240
                  Width           =   3015
               End
               Begin VB.Label LblSelect 
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
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label Label2 
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
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   600
                  Width           =   495
               End
               Begin VB.Label Label3 
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
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   840
                  Width           =   495
               End
               Begin VB.Label Label4 
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
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   1080
                  Width           =   495
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
               TabIndex        =   78
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
               ButtonImage     =   "FrmLinkIteminStorefrm.frx":297D
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   10
               Left            =   0
               TabIndex        =   79
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
               ButtonImage     =   "FrmLinkIteminStorefrm.frx":2F17
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   11
               Left            =   -120
               TabIndex        =   80
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
               ButtonImage     =   "FrmLinkIteminStorefrm.frx":34B1
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ…"
               Height          =   390
               Index           =   5
               Left            =   11400
               TabIndex        =   77
               Top             =   2400
               Width           =   1020
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5040
            Index           =   9
            Left            =   15
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   15
            Width           =   12600
            _cx             =   22225
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
               Left            =   3315
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   1005
               Width           =   660
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   2670
               Left            =   4155
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
               Left            =   2370
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
               Left            =   3975
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
               Left            =   2940
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
      Caption         =   "‰Ê⁄ «·”‰œ"
      Height          =   390
      Index           =   2
      Left            =   1785
      TabIndex        =   76
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”‰œ"
      Height          =   285
      Index           =   3
      Left            =   11520
      TabIndex        =   63
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
Attribute VB_Name = "FrmLinkItemToStore"
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
 Dim coun As Long
 Dim allgroupId As String

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
Sub ret()
 Dim i As Long
  Dim j As Long
  Dim k As Long
 
  Dim Msg As String
  Dim bool As Boolean
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
  Dim m As Long
  bool = True
  
      If ListStoreSelected.ListCount = 0 Then
       If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Õœœ     „Œ“‰ Ê«Õœ ⁄·Ï «·«ﬁ· " & CHR(13)
     Else
     Msg = "Select At Least One Store " & CHR(13)
     End If
            MsgBox Msg, vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 
            Sendkeys "{F4}"
            Exit Sub
        End If
  

 fg.Clear flexClearScrollable, flexClearEverything
 fg.rows = 1

 
    For i = 0 To ListStoreSelected.ListCount - 1
    If XPOptShowType(2).value = True Then
   Set Rs1 = New ADODB.Recordset
sql = "SELECT     dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, "
sql = sql & "                      dbo.TblItems.ItemID"
sql = sql & " FROM         dbo.TblItems INNER JOIN"
sql = sql & "  dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID"

sql = sql & " WHERE   (dbo.TblItems.ItemID =" & val(Me.DcItem1.BoundText) & ")"
Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs1.RecordCount > 0 Then
k = fg.rows
fg.rows = fg.rows + 1
     
       fg.TextMatrix(k, fg.ColIndex("serial")) = j
        fg.TextMatrix(k, fg.ColIndex("StoreName")) = ListStoreSelected.List(i)
        fg.TextMatrix(k, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i)
         fg.TextMatrix(k, fg.ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
         fg.TextMatrix(k, fg.ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), "", Rs1("ItemID").value)
         
         If SystemOptions.UserInterface = ArabicInterface Then
           fg.TextMatrix(k, fg.ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
              fg.TextMatrix(k, fg.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            Else
            fg.TextMatrix(k, fg.ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
            
                fg.TextMatrix(k, fg.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
            End If
   
   End If
        End If
       
           If XPOptShowType(0).value = True Then
        
    Set Rs1 = New ADODB.Recordset
sql = "SELECT     dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, "
sql = sql & "                      dbo.TblItems.ItemID"
sql = sql & " FROM         dbo.TblItems INNER JOIN"
sql = sql & "  dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID"

'sql = sql & " WHERE   (dbo.TblItems.ItemID =" & val(Me.DcItem1.BoundText) & ")"
Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText


    If Rs1.RecordCount > 0 Then
Rs1.MoveFirst
        For j = 1 To Rs1.RecordCount
               
   k = fg.rows
fg.rows = fg.rows + 1

           fg.TextMatrix(k, fg.ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
         fg.TextMatrix(k, fg.ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), "", Rs1("ItemID").value)
         
         If SystemOptions.UserInterface = ArabicInterface Then
           fg.TextMatrix(k, fg.ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
              fg.TextMatrix(k, fg.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            Else
            fg.TextMatrix(k, fg.ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
            
                fg.TextMatrix(k, fg.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
            End If
        fg.TextMatrix(k, fg.ColIndex("serial")) = k
        fg.TextMatrix(k, fg.ColIndex("StoreName")) = ListStoreSelected.List(i)
        fg.TextMatrix(k, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i)
        Rs1.MoveNext
        
       
        
        Next j

    End If
       
       
       
        End If
        Dim GROUPIDS As String
        
          If XPOptShowType(1).value = True Then
          For k = 1 To ListGroupSelected.ListCount

allgroupId = ""
    Set Rs1 = New ADODB.Recordset
        
        GROUPIDS = GetallChilddata(ListGroupSelected.ItemData(k - 1))
        If Len(GROUPIDS) > 2 Then GROUPIDS = mId(GROUPIDS, 2, Len(GROUPIDS))
        Debug.Print GROUPIDS
        If GROUPIDS = "" Then GROUPIDS = ListGroupSelected.ItemData(k - 1)
        allgroupId = allgroupId & "," & GROUPIDS
        sql = " SELECT * from  TblItems where GroupID IN ( " & GROUPIDS & ")"
      
           Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then

        For j = 1 To Rs1.RecordCount
        m = fg.rows
fg.rows = fg.rows + 1

            If SystemOptions.UserInterface = ArabicInterface Then
           
              fg.TextMatrix(m, fg.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            Else
                fg.TextMatrix(m, fg.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
            End If

         fg.TextMatrix(m, fg.ColIndex("ItemID")) = Rs1("ItemID").value
            fg.TextMatrix(m, fg.ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
              fg.TextMatrix(m, fg.ColIndex("GroupName")) = ListGroupSelected.List(k - 1)
       fg.TextMatrix(m, fg.ColIndex("serial")) = coun
        fg.TextMatrix(m, fg.ColIndex("StoreName")) = ListStoreSelected.List(i)
        fg.TextMatrix(m, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i)
        Rs1.MoveNext
        Next j

    End If
       
       
         Next k
        End If
    Next i
   
    
    ReLineGrid

End Sub
'ub retemp()
'Dim i As Integer
' Dim J As Integer
' Dim k As Integer
'
' Dim Msg As String
' Dim bool As Boolean
' Dim Rs1 As ADODB.Recordset
' Dim sql As String
' bool = True
'
'     If ListStoreSelected.ListCount = 0 Then
'      If SystemOptions.UserInterface = ArabicInterface Then
'           Msg = "Õœœ     „Œ“‰ Ê«Õœ ⁄·Ï «·«ﬁ· " & Chr(13)
'    Else
'    Msg = "Select At Least One Store " & Chr(13)
'    End If
'           MsgBox Msg, vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'
'           SendKeys "{F4}"
'           Exit Sub
'       End If
'       FG.Rows = 10000
'       FG.Enabled = True
'        'Set rs1 = New ADODB.Recordset
'        '  sql = " SELECT * from  TblItems "
'        ' rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
      'If (XPOptShowType(0).value = True Or XPOptShowType(1).value = True) And bool = True Then
      '   bool = False
      '             Fg.Rows = (ListStoreSelected.ListCount) * rs1.RecordCount
               
'Fg.Enabled = True
'Else
'If (XPOptShowType(0).value = True Or XPOptShowType(1).value = True) And (bool = False) Then
'                  Fg.Rows = Fg.Rows + ((ListStoreSelected.ListCount) * rs1.RecordCount)
                
'Fg.Enabled = True
'End If
'End If
    '   Fg.Rows = Fg.Rows + 1

 'If (XPOptShowType(2).value = True) And fg.Rows < 2 Then
 '
 '      Else
 '          fg.Rows = ListStoreSelected.ListCount + 1
 '      fg.Enabled = True
 '       End If
 
'   For i = 0 To ListStoreSelected.ListCount - 1
'   If XPOptShowType(2).value = True Then
'             coun = coun + 1
'      FG.TextMatrix(coun, FG.ColIndex("serial")) = coun
'       FG.TextMatrix(coun, FG.ColIndex("StoreName")) = ListStoreSelected.List(i)
'       FG.TextMatrix(coun, FG.ColIndex("StoreID")) = ListStoreSelected.ItemData(i)
'                FG.TextMatrix(coun, FG.ColIndex("ItemName")) = Me.DcItem1.text
'      FG.TextMatrix(coun, FG.ColIndex("ItemID")) = Me.DcItem1.BoundText
'
'       End If
'
'          If XPOptShowType(0).value = True Then
'
'
'   Set Rs1 = New ADODB.Recordset
'          sql = " SELECT * from  TblItems "
'          Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'   If Rs1.RecordCount > 0 Then
'
'       For J = 1 To Rs1.RecordCount
'               coun = coun + 1
'
'
'           If SystemOptions.UserInterface = ArabicInterface Then
'               FG.TextMatrix(coun, FG.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
'           Else
'               FG.TextMatrix(coun, FG.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
'           End If
'
'       FG.TextMatrix(coun, FG.ColIndex("ItemID")) = Rs1("ItemID").value
'       FG.TextMatrix(coun, FG.ColIndex("serial")) = coun
'       FG.TextMatrix(coun, FG.ColIndex("StoreName")) = ListStoreSelected.List(i)
'       FG.TextMatrix(coun, FG.ColIndex("StoreID")) = ListStoreSelected.ItemData(i)
'       Rs1.MoveNext
'
'
'
'       Next J

'   End If
       
       
       
'       End If
'       Dim GROUPIDS As String
'
'         If XPOptShowType(1).value = True Then
'         For k = 1 To ListGroupSelected.ListCount
'
'   Set Rs1 = New ADODB.Recordset
'       '   sql = " SELECT * from  TblItems where GroupID =" & ListGroupSelected.ItemData(k - 1) & ""
'       GROUPIDS = GetallChilddata(ListGroupSelected.ItemData(k - 1))
'       If Len(GROUPIDS) > 2 Then GROUPIDS = Mid(GROUPIDS, 2, Len(GROUPIDS))
'       Debug.Print GROUPIDS
'       If GROUPIDS = "" Then GROUPIDS = ListGroupSelected.ItemData(k - 1)
'       sql = " SELECT * from  TblItems where GroupID IN ( " & GROUPIDS & ")"
'
'      '(GetallChilddata
'          Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'   If Rs1.RecordCount > 0 Then
'
'       For J = 1 To Rs1.RecordCount
'oun = coun + 1
'           If SystemOptions.UserInterface = ArabicInterface Then
'
'             FG.TextMatrix(coun, FG.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
'           Else
'               FG.TextMatrix(coun, FG.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
'           End If
'
'        FG.TextMatrix(coun, FG.ColIndex("ItemID")) = Rs1("ItemID").value
''          FG.TextMatrix(coun, FG.ColIndex("GroupID")) = Rs1("GroupID").value
 '            FG.TextMatrix(coun, FG.ColIndex("GroupName")) = ListGroupSelected.List(k - 1)
 '     FG.TextMatrix(coun, FG.ColIndex("serial")) = coun
 '      FG.TextMatrix(coun, FG.ColIndex("StoreName")) = ListStoreSelected.List(i)
 '      FG.TextMatrix(coun, FG.ColIndex("StoreID")) = ListStoreSelected.ItemData(i)
 '      Rs1.MoveNext
 '      Next J
'
'   End If
       
       
'        Next k
'       End If
'   Next i
'
'   FG.Rows = coun + 1
'
'   ReLineGrid

'nd Sub


Private Sub BtonAdd_Click()
ret
End Sub

Private Sub Cmd_Click(index As Integer)

    ' On Error GoTo ErrTrap
    Select Case index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
          
ListStoreSelected.Clear
ListGroupSelected.Clear
 fg.Clear flexClearScrollable, flexClearEverything
 fg.rows = 1
 Frame10.Enabled = True
 XPOptShowType(1).value = False
            TxtModFlg.text = "N"
            clear_all Me
  Me.DcbOrderStatus.ListIndex = 1
            
         '     GRID2.Clear flexClearScrollable, flexClearEverything
    'GRID2.Rows = 1
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

            TxtModFlg.text = "E"
            Frame10.Enabled = True
            Frame11.Enabled = True
            Me.DCboUserName.BoundText = user_id

        Case 2
    
            Dim Msg As String

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

            Del_Trans

        Case 5
            Load FrmLinkIteminStoreSearch
            FrmLinkIteminStoreSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 12
         RemoveGridRow
            Case 13
             fg.Clear flexClearScrollable, flexClearEverything
 fg.rows = 1
            'sa coun = 0
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

MySQL = " SELECT     dbo.TblLink_Item_To_StoreH.Ind, dbo.TblLink_Item_To_StoreH.LinkType, dbo.TblLink_Item_To_StoreH.UserID, dbo.TblUsers.UserName,"
MySQL = MySQL & "                      dbo.TblLink_Item_To_StoreH.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblLink_Item_To_StoreH.RecordeDate,"
MySQL = MySQL & "                       dbo.TblLink_Item_To_StoreH.Remarks, dbo.TblLink_Item_To_StoreH.Selected, dbo.TblLink_Item_To_StoreH.Posted, dbo.TblLink_Item_To_Store_Details2.Ind AS Ind2,"
MySQL = MySQL & "                       dbo.TblLink_Item_To_Store_Details2.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblLink_Item_To_Store_Details2.ItemID,"
MySQL = MySQL & "                       dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblLink_Item_To_Store_Details2.LinkType AS linktype2, dbo.TblLink_Item_To_Store_Details2.GroupID,"
MySQL = MySQL & "                       dbo.Groups.GroupName , dbo.Groups.GroupNamee"
MySQL = MySQL & "  , dbo.TblItems.ItemCode  FROM         dbo.TblLink_Item_To_StoreH LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.TblLink_Item_To_StoreH.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblUsers ON dbo.TblLink_Item_To_StoreH.UserID = dbo.TblUsers.UserID RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.Groups RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblLink_Item_To_Store_Details2 ON dbo.Groups.GroupID = dbo.TblLink_Item_To_Store_Details2.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblItems ON dbo.TblLink_Item_To_Store_Details2.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblStore ON dbo.TblLink_Item_To_Store_Details2.StoreID = dbo.TblStore.StoreID ON"
MySQL = MySQL & "                       dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details2.Ind"
MySQL = MySQL & " Where (dbo.TblLink_Item_To_StoreH.Ind =" & val(XPTxtID.text) & ")"
'MySQL = MySQL & "Where (dbo.TblLink_Item_To_StoreH.Ind = " & val(XPTxtID.text) & ")"

 
 'MySQL = MySQL & "   Where (dbo.TblTreatment.id =" & val(XPTxtID.text) & ")"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepLinkingIteminStore.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepLinkingIteminStore.rpt"
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

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub


Private Sub txtDiscountDES_Change()

End Sub



Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim ItemID As Integer

    If KeyAscii = vbKeyReturn Then
        GetItemIDFromCode TxtSearchCode.text, ItemID
        DcItem1.BoundText = ItemID
    End If

End Sub

 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 9
        FrmEmployeeSearch.show
  
    End If

End Sub

Private Sub DcboEmpName_Click(Area As Integer)
   
  
 

End Sub



Private Sub Command1_Click()
  Dim i As Long
  Dim j As Long
  Dim k As Long
 
  Dim Msg As String
  Dim bool As Boolean
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
  bool = True
  
      If ListStoreSelected.ListCount = 0 Then
       If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Õœœ     „Œ“‰ Ê«Õœ ⁄·Ï «·«ﬁ· " & CHR(13)
     Else
     Msg = "Select At Least One Store " & CHR(13)
     End If
            MsgBox Msg, vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 
            Sendkeys "{F4}"
            Exit Sub
        End If
        fg.rows = 10000
        fg.Enabled = True
         'Set rs1 = New ADODB.Recordset
         '  sql = " SELECT * from  TblItems "
         ' rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
      'If (XPOptShowType(0).value = True Or XPOptShowType(1).value = True) And bool = True Then
      '   bool = False
      '             Fg.Rows = (ListStoreSelected.ListCount) * rs1.RecordCount
               
'Fg.Enabled = True
'Else
'If (XPOptShowType(0).value = True Or XPOptShowType(1).value = True) And (bool = False) Then
'                  Fg.Rows = Fg.Rows + ((ListStoreSelected.ListCount) * rs1.RecordCount)
                
'Fg.Enabled = True
'End If
'End If
    '   Fg.Rows = Fg.Rows + 1

 'If (XPOptShowType(2).value = True) And fg.Rows < 2 Then
 '
 '      Else
 '          fg.Rows = ListStoreSelected.ListCount + 1
 '      fg.Enabled = True
 '       End If
 
    For i = 1 To ListStoreSelected.ListCount
    If XPOptShowType(2).value = True Then
           coun = coun + 1
       fg.TextMatrix(count, fg.ColIndex("serial")) = coun
        fg.TextMatrix(coun, fg.ColIndex("StoreName")) = ListStoreSelected.List(i - 1)
        fg.TextMatrix(coun, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i - 1)
                 fg.TextMatrix(coun, fg.ColIndex("ItemName")) = Me.DcItem1.text
       fg.TextMatrix(coun, fg.ColIndex("ItemID")) = Me.DcItem1.BoundText
        End If
       
           If XPOptShowType(0).value = True Then
        

    Set Rs1 = New ADODB.Recordset
           sql = " SELECT * from  TblItems "
           Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then

        For j = 1 To Rs1.RecordCount
coun = coun + 1
            If SystemOptions.UserInterface = ArabicInterface Then
           
              fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            Else
                fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
            End If

         fg.TextMatrix(coun, fg.ColIndex("ItemID")) = Rs1("ItemID").value
           
                 
       fg.TextMatrix(coun, fg.ColIndex("serial")) = coun
        fg.TextMatrix(coun, fg.ColIndex("StoreName")) = ListStoreSelected.List(i - 1)
        fg.TextMatrix(coun, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i - 1)
        Rs1.MoveNext
        Next j

    End If
       
       
       
        End If
          If XPOptShowType(1).value = True Then
          For k = 1 To ListGroupSelected.ListCount

    Set Rs1 = New ADODB.Recordset
           sql = " SELECT * from  TblItems where GroupID =" & ListGroupSelected.ItemData(k - 1) & ""
           Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then

        For j = 1 To Rs1.RecordCount
coun = coun + 1
            If SystemOptions.UserInterface = ArabicInterface Then
           
              fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
            Else
                fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
            End If

         fg.TextMatrix(coun, fg.ColIndex("ItemID")) = Rs1("ItemID").value
            fg.TextMatrix(coun, fg.ColIndex("GroupID")) = Rs1("GroupID").value
              fg.TextMatrix(coun, fg.ColIndex("GroupName")) = ListGroupSelected.List(k - 1)
       fg.TextMatrix(coun, fg.ColIndex("serial")) = coun
        fg.TextMatrix(coun, fg.ColIndex("StoreName")) = ListStoreSelected.List(i - 1)
        fg.TextMatrix(coun, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i - 1)
        Rs1.MoveNext
        Next j

    End If
       
       
         Next k
        End If
    Next i
    If XPOptShowType(0).value = True Or XPOptShowType(1).value = True Then
    fg.rows = coun + 1
    End If
    ReLineGrid
End Sub

Private Sub Label2_Click()
    Dim i As Integer
    ListStoreSelected.Clear

    For i = 0 To ListStoreall.ListCount - 1
        ListStoreSelected.AddItem ListStoreall.List(i)
        ListStoreSelected.ItemData(i) = ListStoreall.ItemData(i)
    Next i

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
    Dim i As Integer
    If Me.XPOptShowType(1).value = True Then
    ListGroupSelected.Clear

    For i = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(i)
        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
    Next i
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
Private Sub Label4_Click()

    If ListStoreSelected.ListIndex > -1 Then
    
        ListStoreSelected.RemoveItem ListStoreSelected.ListIndex
    End If

End Sub
Private Sub Label3_Click()
    ListStoreSelected.Clear
End Sub

Private Sub LblSelect_Click()
If ListStoreall.ListIndex > -1 Then
    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
        
    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
            End If
End Sub

Private Sub ListGroupAll_Click()
 If XPOptShowType(1).value = True Then
        Frame11.Enabled = True
    Else
        Frame11.Enabled = False
    End If
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
 'Dim count As Integer
 coun = 0
     On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
   Frame10.Enabled = False
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
     If SystemOptions.UserInterface = EnglishInterface Then
       Me.DcbOrderStatus.AddItem "Cancel Link"
       Me.DcbOrderStatus.AddItem " link"
      Else
        Me.DcbOrderStatus.AddItem " ≈·€«¡«·—»ÿ"
       Me.DcbOrderStatus.AddItem " —»ÿ"
    End If
    Resize_Form Me
    AddTip
    FillMylist
    Set Dcombos = New ClsDataCombos
      Dcombos.GetUsers Me.DCboUserName
  
    Dcombos.GetBranches Me.Dcbranch
     Dcombos.GetItemsNames DcItem1, , , , True

   ' Dcombos.GetItemSGroups Me.DCbGroup, False



    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
    'YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblLink_Item_To_StoreH     Order By Ind"
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

    Me.Caption = "Linking Items With Stores "
    EleHeader.Caption = Me.Caption
    lbl(3).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(4).Caption = "Branch"
    lbl(2).Caption = "Type Linking"
    Frame10.Caption = "Select Store"
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
   XPTab301.Caption = "Linking Item"
lbl(5).Caption = "Remarks"
   lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"

   With Me.fg
        .TextMatrix(0, .ColIndex("serial")) = "NO"
        .TextMatrix(0, .ColIndex("StoreID")) = "StoreID"
        .TextMatrix(0, .ColIndex("StoreName")) = "StoreName"
         .TextMatrix(0, .ColIndex("ItemID")) = "ItemID"
        .TextMatrix(0, .ColIndex("ItemName")) = "ItemName"
         .TextMatrix(0, .ColIndex("GroupID")) = "GroupID"
        .TextMatrix(0, .ColIndex("GroupName")) = "GroupName"

    End With

End Sub

Private Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Long

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
    Dim RsDetails1 As ADODB.Recordset
    Dim RsDetails2 As ADODB.Recordset
    Frame10.Enabled = False
    Frame11.Enabled = False
     coun = 0
    Dim i As Long
    Dim StrSQL As String
    ListGroupSelected.Clear
    ListStoreSelected.Clear
fg.Clear flexClearScrollable, flexClearEverything
            fg.rows = 2
            fg.Enabled = True
 On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "Ind=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.text = IIf(IsNull(rs("Ind").value), "", val(rs("Ind").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordeDate").value), Date, rs("RecordeDate").value)
    Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
   Me.TxtRemark.text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
Me.DcbOrderStatus.ListIndex = rs("LinkType").value
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
StrSQL = " SELECT     dbo.TblLink_Item_To_StoreH.Ind, dbo.TblLink_Item_To_Store_Details2.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee,"
 StrSQL = StrSQL & "                       dbo.TblLink_Item_To_Store_Details2.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblLink_Item_To_Store_Details2.GroupID,"
 StrSQL = StrSQL & "                      dbo.Groups.GroupName , dbo.Groups.GroupNamee"
StrSQL = StrSQL & "  FROM         dbo.TblLink_Item_To_StoreH RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.Groups RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblLink_Item_To_Store_Details2 ON dbo.Groups.GroupID = dbo.TblLink_Item_To_Store_Details2.GroupID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblItems ON dbo.TblLink_Item_To_Store_Details2.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblStore ON dbo.TblLink_Item_To_Store_Details2.StoreID = dbo.TblStore.StoreID ON"
 StrSQL = StrSQL & "                      dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details2.Ind"
StrSQL = StrSQL & "  Where (dbo.TblLink_Item_To_StoreH.Ind = " & val(XPTxtID.text) & ")"

    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
       fg.rows = fg.FixedRows + RsDetails.RecordCount
If rs("Selected").value = 0 Then
XPOptShowType(0).value = True
End If
If rs("Selected").value = 1 Then
XPOptShowType(1).value = True
End If
If rs("Selected").value = 2 Then
DcItem1.BoundText = RsDetails("ItemID").value
XPOptShowType(2).value = True
Else
DcItem1.BoundText = ""

End If
        For i = Me.fg.FixedRows To fg.rows - 1
     
          fg.TextMatrix(i, fg.ColIndex("serial")) = i
           fg.TextMatrix(i, fg.ColIndex("ItemID")) = RsDetails("ItemID").value
            fg.TextMatrix(i, fg.ColIndex("StoreID")) = RsDetails("StoreID").value
          fg.TextMatrix(i, fg.ColIndex("ItemName")) = IIf(IsNull(RsDetails("ItemName").value), "", RsDetails("ItemName").value)
            fg.TextMatrix(i, fg.ColIndex("StoreName")) = RsDetails("StoreName").value
            If RsDetails("GroupID").value <> 0 Then
            fg.TextMatrix(i, fg.ColIndex("GroupID")) = RsDetails("GroupID").value
          fg.TextMatrix(i, fg.ColIndex("GroupName")) = IIf(IsNull(RsDetails("GroupName").value), "", RsDetails("GroupName").value)
          End If
     
            RsDetails.MoveNext
        Next i

    End If
     Set RsDetails2 = New ADODB.Recordset
 StrSQL = "  SELECT     dbo.TblLink_Item_To_Store_Details1.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblLink_Item_To_Store_Details1.Ind"
 
StrSQL = StrSQL & " FROM         dbo.TblLink_Item_To_Store_Details1 LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblStore ON dbo.TblLink_Item_To_Store_Details1.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblLink_Item_To_StoreH ON dbo.TblLink_Item_To_Store_Details1.Ind = dbo.TblLink_Item_To_StoreH.Ind"
 StrSQL = StrSQL & " Where (dbo.TblLink_Item_To_Store_Details1.Ind = " & val(XPTxtID.text) & ")"

  RsDetails2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
  For i = 0 To RsDetails2.RecordCount - 1
  ListStoreSelected.AddItem RsDetails2("StoreName").value
  ListStoreSelected.ItemData(i) = val(RsDetails2("StoreID").value)
  
   RsDetails2.MoveNext
  
  Next i
   RsDetails2.Close
    Set RsDetails2 = Nothing
   
       Set RsDetails1 = New ADODB.Recordset
StrSQL = " SELECT     dbo.TblLink_Item_To_Store_Details3.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblLink_Item_To_Store_Details3.Ind"
StrSQL = StrSQL & " FROM         dbo.Groups RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblLink_Item_To_Store_Details3 ON dbo.Groups.GroupID = dbo.TblLink_Item_To_Store_Details3.GroupID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblLink_Item_To_StoreH ON dbo.TblLink_Item_To_Store_Details3.ID = dbo.TblLink_Item_To_StoreH.Ind"
StrSQL = StrSQL & " WHERE     (dbo.TblLink_Item_To_Store_Details3.Ind = " & val(XPTxtID.text) & ")"
  RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
  For i = 0 To RsDetails1.RecordCount - 1
  ListGroupSelected.AddItem RsDetails1("GroupName").value
  ListGroupSelected.ItemData(i) = val(RsDetails1("GroupID").value)
  
   RsDetails1.MoveNext
  
  Next i
  '''''''''''''''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  
   
    RsDetails1.Close
    Set RsDetails1 = Nothing
    
     RsDetails.Close
    Set RsDetails = Nothing
    fillapprovData
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
 coun = 0

    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
     Dim sql As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
     Dim RsDetails2 As ADODB.Recordset
      
    Dim i As Long
    Dim j As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap

 If fg.rows < 2 Then
       If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " ·« ÊÃœ »Ì«‰«  " & CHR(13)
     Else
     Msg = "¬Not Found Data " & CHR(13)
     End If
            MsgBox Msg, vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 
            Sendkeys "{F4}"
            Exit Sub
        End If

   
      
     
If Me.DcbOrderStatus.ListIndex = 0 Then

For i = fg.FixedRows To fg.rows - 1
   sql = "update TblLink_Item_To_Store_Details2 set   LinkType=0   where  StoreID =" & val(fg.TextMatrix(i, fg.ColIndex("StoreID"))) & " and GroupID =" & val(fg.TextMatrix(i, fg.ColIndex("GroupID"))) & " and ItemID =" & val(fg.TextMatrix(i, fg.ColIndex("ItemID"))) & ""
                                    Cn.Execute sql
Next i
Else
  Dim RsTest As New ADODB.Recordset
    
    

        Cn.BeginTrans
        BeginTrans = True
        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("TblLink_Item_To_StoreH", "Ind", "", True))
    
        
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
          StrSQL = "Delete From TblLink_Item_To_Store_Details3 Where Ind=" & val(Me.XPTxtID.text)
           Cn.Execute StrSQL, , adExecuteNoRecords
   StrSQL = "Delete From TblLink_Item_To_Store_Details2 Where Ind=" & val(Me.XPTxtID.text)
           Cn.Execute StrSQL, , adExecuteNoRecords
              StrSQL = "Delete From TblLink_Item_To_Store_Details1 Where Ind=" & val(Me.XPTxtID.text)
           Cn.Execute StrSQL, , adExecuteNoRecords
        End If
        rs("Ind").value = val(XPTxtID.text)
        
        rs("RecordeDate").value = XPDtbTrans.value
        rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
         rs("LinkType").value = Me.DcbOrderStatus.ListIndex
        rs("Remarks").value = Me.TxtRemark.text
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
   Set RsDetails1 = New ADODB.Recordset
   '     RsDetails1.Open "TblLink_Item_To_Store_Details1", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     *  from dbo.TblLink_Item_To_Store_Details1 Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
        For i = 0 To ListStoreSelected.ListCount - 1
                  RsDetails1.AddNew
             RsDetails1("Ind").value = val(XPTxtID.text)
             RsDetails1("StoreID").value = val(ListStoreSelected.ItemData(i))
                      RsDetails1.update
           
    Next i
   ''''''''''' /////////////////////////////''''''
   Set RsDetails2 = New ADODB.Recordset
     '   RsDetails2.Open "TblLink_Item_To_Store_Details3", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
     StrSQL = "SELECT     *  from dbo.TblLink_Item_To_Store_Details3 Where (1 = -1)"
   RsDetails2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
             
             
        For i = 0 To ListGroupSelected.ListCount - 1
                  RsDetails2.AddNew
             RsDetails2("Ind").value = val(XPTxtID.text)
             RsDetails2("GroupID").value = val(ListGroupSelected.ItemData(i))
                      RsDetails2.update
           
    Next i
        '''''''''///////////////////////////////////////////
           Set RsDetails = New ADODB.Recordset
     '   RsDetails.Open "TblLink_Item_To_Store_Details2", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

     StrSQL = "SELECT     *  from dbo.TblLink_Item_To_Store_Details2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
        For i = Me.fg.FixedRows To fg.rows - 1
       If fg.TextMatrix(i, fg.ColIndex("StoreName")) <> "" Then
            RsDetails.AddNew
             RsDetails("Ind").value = val(XPTxtID.text)
             RsDetails("ItemID").value = val(fg.TextMatrix(i, fg.ColIndex("ItemID")))
                RsDetails("StoreID").value = val(fg.TextMatrix(i, fg.ColIndex("StoreID")))
                 RsDetails("GroupID").value = val(fg.TextMatrix(i, fg.ColIndex("GroupID")))
        
                RsDetails("LinkType").value = 1
                RsDetails.update
                End If
        Next i
       
       
  
        Cn.CommitTrans
        BeginTrans = False
          RsDetails1.Close
         Set RsDetails1 = Nothing
           RsDetails2.Close
         Set RsDetails2 = Nothing
     RsDetails.Close
         Set RsDetails = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    End If
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
            rs.Find "Ind='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

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
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
  
        '        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.Text)
        '        Cn.Execute StrSQL, , adExecuteNoRecords
        
                StrSQL1 = "Delete From TblLink_Item_To_Store_Details2 Where Ind=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL1, , adExecuteNoRecords
               StrSQL1 = "Delete From TblLink_Item_To_Store_Details1 Where Ind=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL1, , adExecuteNoRecords
                 StrSQL1 = "Delete From TblLink_Item_To_Store_Details3 Where Ind=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL1, , adExecuteNoRecords
                              rs.delete
        If rs.RecordCount > 0 Then
              rs.MoveFirst
         End If
         
             
                If rs.RecordCount < 1 Then
                    clear_all Me
                        ListGroupSelected.Clear
    ListStoreSelected.Clear

                   fg.Clear flexClearScrollable, flexClearEverything
                   fg.rows = 2
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
   'RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     
     StrSQL = "SELECT     *  from dbo.ApprovalData Where (1 = -1)"
   RSApproval.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
       

 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Long
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
Private Sub ReLineGrid()
    Dim i As Long
    Dim IntCounter  As Long
    
    IntCounter = 0

    With fg

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("StoreName")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("serial")) = IntCounter
           
    
        End If
                

        Next i
 
    End With

End Sub
Function FillMylist()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Long
    sql = " SELECT * from  TblStore"
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListStoreall.Clear
    ListStoreSelected.Clear

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
             
            If SystemOptions.UserInterface = ArabicInterface Then
                ListStoreall.AddItem IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
            Else
                ListStoreall.AddItem IIf(IsNull(rs("StoreNamee").value), "", rs("StoreNamee").value)
            End If

            ListStoreall.ItemData(ListStoreall.NewIndex) = rs("StoreID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

    'fil

  sql = " SELECT * from  Groups where GroupID>1"
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListGroupAll.Clear
    ListGroupSelected.Clear

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount

            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupAll.AddItem IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
            Else
                ListGroupAll.AddItem IIf(IsNull(rs("GroupNamee").value), "", rs("GroupNamee").value)
            End If

            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("GroupID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

End Function
Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "—»ÿ «·«’‰«› »«·„Œ«“‰ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "—»ÿ «·«’‰«› »«·„Œ«“‰ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " —»ÿ «·«’‰«› »«·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "—»ÿ «·«’‰«› »«·„Œ«“‰ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "—»ÿ «·«’‰«› »«·„Œ«“‰ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "—»ÿ «·«’‰«› »«·„Œ«“‰ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "—»ÿ «·«’‰«› »«·„Œ«“‰ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "—»ÿ «·«’‰«› »«·„Œ«“‰ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "—»ÿ «·«’‰«› »«·„Œ«“‰ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "—»ÿ «·«’‰«› »«·„Œ«“‰ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "—»ÿ «·«’‰«› »«·„Œ«“‰ ", 1, 15204351, -2147483630
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
       
                'SaveData

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

 Private Sub RemoveGridRow()
coun = coun - 1
    With Me.fg

        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With

    ReLineGrid
End Sub
 Private Sub RemoveGridRowSpace()

    With Me.fg

        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With

    ReLineGrid
End Sub
 
Private Sub XPOptShowType_Click(index As Integer)
 If XPOptShowType(1).value = True Then
        Frame11.Enabled = True
    Else
        Frame11.Enabled = False
    End If
End Sub

