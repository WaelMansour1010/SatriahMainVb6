VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLockedOrderMaintenance 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«ﬁ›«· ÿ·» «·’Ì«‰Â"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   FillColor       =   &H00C0E0FF&
   Icon            =   "FrmLockedOrderMaintenance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   11775
   Begin VB.TextBox TxtOrderNo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   720
      Width           =   1455
   End
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
      Left            =   12840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   3600
      Width           =   825
   End
   Begin VB.TextBox TxtDayPrice 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   51
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
      Width           =   11805
      _cx             =   20823
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
      Caption         =   "«ﬁ›«· ÿ·» «·’Ì«‰Â"
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
         ButtonImage     =   "FrmLockedOrderMaintenance.frx":038A
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
         ButtonImage     =   "FrmLockedOrderMaintenance.frx":0724
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
         ButtonImage     =   "FrmLockedOrderMaintenance.frx":0ABE
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
         ButtonImage     =   "FrmLockedOrderMaintenance.frx":0E58
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
         Left            =   3120
         Picture         =   "FrmLockedOrderMaintenance.frx":11F2
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
         Left            =   2160
         TabIndex        =   32
         Top             =   0
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   7380
      TabIndex        =   6
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   103153665
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   270
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5460
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
      Top             =   4800
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
      Bindings        =   "FrmLockedOrderMaintenance.frx":4E5A
      Height          =   315
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   3615
      Left            =   0
      TabIndex        =   36
      Top             =   1200
      Width           =   11760
      _cx             =   20743
      _cy             =   6376
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
      Picture(0)      =   "FrmLockedOrderMaintenance.frx":4E6F
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3150
         Left            =   12405
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   45
         Width           =   11670
         _cx             =   20585
         _cy             =   5556
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
            FormatString    =   $"FrmLockedOrderMaintenance.frx":5209
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
         Height          =   3150
         Index           =   15
         Left            =   45
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   45
         Width           =   11670
         _cx             =   20585
         _cy             =   5556
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
         _GridInfo       =   $"FrmLockedOrderMaintenance.frx":5355
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3120
            Index           =   16
            Left            =   15
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   15
            Width           =   11640
            _cx             =   20532
            _cy             =   5503
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
            Begin VB.Frame lblDataCli 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  "
               Height          =   3060
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   0
               Width           =   11580
               Begin VB.TextBox TXtCount 
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
                  Left            =   5640
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   2400
                  Width           =   705
               End
               Begin VB.ComboBox DcbDMY 
                  Height          =   315
                  ItemData        =   "FrmLockedOrderMaintenance.frx":5389
                  Left            =   4440
                  List            =   "FrmLockedOrderMaintenance.frx":5396
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   2400
                  Width           =   1215
               End
               Begin VB.TextBox TxtLocation 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   5400
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   960
                  Width           =   4815
               End
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   9360
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   1320
                  Width           =   855
               End
               Begin VB.TextBox TxtSearch 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
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
                  Left            =   9360
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   600
                  Width           =   855
               End
               Begin VB.TextBox TxtSearchCodeSuper 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   9360
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   240
                  Width           =   855
               End
               Begin VB.TextBox TxtRemark 
                  Alignment       =   1  'Right Justify
                  Height          =   1095
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   70
                  Top             =   1080
                  Width           =   4455
               End
               Begin VB.ComboBox DcbStatus 
                  Appearance      =   0  'Flat
                  Height          =   315
                  ItemData        =   "FrmLockedOrderMaintenance.frx":53A9
                  Left            =   5400
                  List            =   "FrmLockedOrderMaintenance.frx":53B3
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   1800
                  Width           =   975
               End
               Begin VB.CheckBox ChLock 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«ﬁ›«· «·ÿ·»"
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   2400
                  Width           =   1215
               End
               Begin VB.TextBox TxtOrderWork 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   675
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   64
                  Top             =   240
                  Width           =   4455
               End
               Begin MSComCtl2.DTPicker EndDate 
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   55
                  Top             =   2400
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   103153665
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal EndDateH 
                  Height          =   315
                  Left            =   180
                  TabIndex        =   56
                  Top             =   2400
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
               End
               Begin MSDataListLib.DataCombo DcboEmpNameSuper 
                  Height          =   315
                  Left            =   5400
                  TabIndex        =   62
                  Top             =   240
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbIqara 
                  Height          =   315
                  Left            =   5400
                  TabIndex        =   74
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
                  Top             =   600
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboEmpName 
                  Height          =   315
                  Left            =   5400
                  TabIndex        =   77
                  Top             =   1320
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker SatarDate 
                  Height          =   315
                  Left            =   8940
                  TabIndex        =   81
                  Top             =   1800
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   103153665
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal SatarDateH 
                  Height          =   315
                  Left            =   7440
                  TabIndex        =   82
                  Top             =   1800
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
               End
               Begin MSComCtl2.DTPicker LockDate 
                  Height          =   315
                  Left            =   8940
                  TabIndex        =   87
                  Top             =   2400
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   103153665
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal LockDateH 
                  Height          =   315
                  Left            =   7440
                  TabIndex        =   88
                  Top             =   2400
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "’·«ÕÌÂ «·ÿ·»"
                  Height          =   375
                  Index           =   12
                  Left            =   6120
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   2400
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ »œ«Ì…«·ÿ·»"
                  Height          =   375
                  Index           =   11
                  Left            =   10260
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   1800
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„Êﬁ⁄ «·⁄ﬁ«—"
                  Height          =   255
                  Index           =   18
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„ÊŸ› «·’Ì«‰Â"
                  Height          =   285
                  Index           =   5
                  Left            =   10410
                  TabIndex        =   78
                  Top             =   1335
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄ﬁ«—"
                  Height          =   255
                  Index           =   13
                  Left            =   10680
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   600
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ« "
                  Height          =   255
                  Index           =   10
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄„· «·„ÿ·Ê»"
                  Height          =   615
                  Index           =   2
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”ƒÊ· «·’Ì«‰Â"
                  Height          =   285
                  Index           =   3
                  Left            =   10410
                  TabIndex        =   63
                  Top             =   255
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ ‰Â«Ì… «·ÿ·»"
                  Height          =   375
                  Index           =   20
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   2400
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Õ«·… «·’Ì«‰Â"
                  Height          =   375
                  Index           =   17
                  Left            =   6120
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   1800
                  Width           =   1215
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   1830
               Index           =   62
               Left            =   2220
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   840
               Width           =   540
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3120
            Index           =   9
            Left            =   15
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   15
            Width           =   11640
            _cx             =   20532
            _cy             =   5503
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
               Height          =   2340
               Left            =   3045
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   705
               Width           =   630
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   1620
               Left            =   3855
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   840
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   1620
               Index           =   67
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   840
               Width           =   600
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   1560
               Index           =   68
               Left            =   3675
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   1095
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
               Height          =   1875
               Index           =   69
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   840
               Width           =   285
            End
         End
      End
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   0
      TabIndex        =   53
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
      Left            =   5760
      TabIndex        =   58
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
   End
   Begin MSComCtl2.DTPicker From 
      Height          =   315
      Left            =   12360
      TabIndex        =   60
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   103153665
      CurrentDate     =   38784
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«ﬁ›«· «·ÿ·» —ﬁ„"
      Height          =   285
      Index           =   9
      Left            =   4320
      TabIndex        =   68
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄„«—Â"
      Height          =   255
      Left            =   3660
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   0
      Width           =   855
   End
   Begin VB.Image img 
      Height          =   855
      Left            =   22680
      Picture         =   "FrmLockedOrderMaintenance.frx":53C3
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
      Picture         =   "FrmLockedOrderMaintenance.frx":63E7
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   720
   End
   Begin VB.Label lblBr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Left            =   1920
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
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   285
      Index           =   4
      Left            =   10680
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
      Left            =   8400
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
      Top             =   4875
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
      Top             =   5040
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
      Top             =   5070
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   19
      Top             =   5100
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   18
      Top             =   5100
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
Attribute VB_Name = "FrmLockedOrderMaintenance"
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


Sub SaveUoitInformation(Optional Unit As Integer = 0, Optional Status As Integer = 0, Optional cus As Double)
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL, Msg As String
 If Unit <> 0 Then
Msg = ""

    If SystemOptions.UserInterface = EnglishInterface Then
      Msg = Msg & "Work was guaranteed catch filtering No."
             Msg = Msg & Chr(13) & XPTxtID.Text
                  Msg = Msg & Chr(13)
      Msg = Msg & "  Date  Order "
      Msg = Msg & NourHijriCal1.value & " «·„Ê«›ﬁ " & XPDtbTrans.value
    Msg = Msg & Chr(13)
    Msg = Msg & "  Date lock Order "
      Msg = Msg & LockDateH.value & " «·„Ê«›ﬁ " & LockDate.value
    Msg = Msg & Chr(13)
     Msg = Msg & "  The case of maintenance "
      Msg = Msg & DcbStatus.Text
    Msg = Msg & Chr(13)
      Else
      Msg = Msg & "   „ ⁄„· «ﬁ›«· ÿ·» ’Ì«‰Â —ﬁ„   "
      Msg = Msg & XPTxtID.Text
      Msg = Msg & Chr(13)
      Msg = Msg & " » «—ÌŒ «·ÿ·» "
      Msg = Msg & NourHijriCal1.value & " «·„Ê«›ﬁ " & XPDtbTrans.value
    Msg = Msg & Chr(13)
    Msg = Msg & " » «—ÌŒ «·«ﬁ›«· "
      Msg = Msg & LockDateH.value & " «·„Ê«›ﬁ " & LockDate.value
    Msg = Msg & Chr(13)
     Msg = Msg & "  Õ«·… «·’Ì«‰Â "
      Msg = Msg & DcbStatus.Text
    Msg = Msg & Chr(13)
End If
        Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TblUnitNoInformation Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      RsDetails1.AddNew
      RsDetails1("CusID").value = cus
      RsDetails1("BranchId").value = Me.Dcbranch.BoundText
           RsDetails1("UnitNo").value = Unit
           RsDetails1("UnitStatus").value = Status
           RsDetails1("Des").value = Msg
           RsDetails1("RecDate").value = XPDtbTrans.value
           RsDetails1("RecDateH").value = NourHijriCal1.value
           RsDetails1("NoteID").value = Null
           RsDetails1("ContNo").value = Null
           RsDetails1("FilterNo").value = Null
           RsDetails1("OrderMaint").value = Null
           RsDetails1("LocOrderMaint").value = val(XPTxtID.Text)
           RsDetails1.update
           End If

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
    
Dcbranch.BoundText = Current_branch
  Me.DCboUserName.BoundText = user_id
  DcbStatus.ListIndex = 0
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
 'UnitsGrid.Rows = UnitsGrid.Rows + 1
 '           UnitsGrid.Enabled = True
            '
            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
    
            Dim Msg As String

            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Dcbranch.SetFocus
                SendKeys "{F4}"
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
            Load FrmLocedkOrderMaintAqarSearch
            FrmLocedkOrderMaintAqarSearch.show

        Case 6
            Unload Me

        Case 7
   

        Case 8
  
            
            
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


    MySQL = " SELECT     dbo.TblLockedOrderMaintenance.ID, dbo.TblLockedOrderMaintenance.RecDateH, dbo.TblLockedOrderMaintenance.RecDate, "
 MySQL = MySQL & "                     dbo.TblLockedOrderMaintenance.OrderNo, dbo.TblLockedOrderMaintenance.LocationIqar, dbo.TblLockedOrderMaintenance.Remark,"
MySQL = MySQL & "                      dbo.TblLockedOrderMaintenance.OrderWork, dbo.TblLockedOrderMaintenance.DMY, dbo.TblLockedOrderMaintenance.Status, dbo.TblLockedOrderMaintenance.Cont,"
MySQL = MySQL & "                      dbo.TblLockedOrderMaintenance.EndFateH, dbo.TblLockedOrderMaintenance.EndFate, dbo.TblLockedOrderMaintenance.Lock,"
MySQL = MySQL & "                     dbo.TblLockedOrderMaintenance.SatarDateH, dbo.TblLockedOrderMaintenance.SatarDate, dbo.TblLockedOrderMaintenance.LockDateH,"
MySQL = MySQL & "                      dbo.TblLockedOrderMaintenance.LockDate, dbo.TblLockedOrderMaintenance.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
MySQL = MySQL & "                      dbo.TblLockedOrderMaintenance.AqrID, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblLockedOrderMaintenance.EmpID, dbo.TblEmployee.Emp_Code,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,"
MySQL = MySQL & "                      dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1,"
MySQL = MySQL & "                    dbo.TblEmployee.Emp_Namee, dbo.TblLockedOrderMaintenance.SuperVM, TblEmployee_1.Emp_Code AS Emp_CodeSup,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name AS Emp_NameSup, TblEmployee_1.Emp_Name1 AS Emp_Name1Sup, TblEmployee_1.Emp_Name2 AS Emp_Name2Sup,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name3 AS Emp_Name3Sup, TblEmployee_1.Emp_Name4 AS Emp_Name4Sup, TblEmployee_1.Fullcode AS FullcodeSup,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee4 AS Emp_Namee4Sup, TblEmployee_1.Emp_Namee3 AS Emp_Namee3Sup, TblEmployee_1.Emp_Namee2 AS Emp_Namee2Sup,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee1 AS Emp_Namee1Sup, TblEmployee_1.Emp_Namee AS Emp_NameeSup"
MySQL = MySQL & " FROM         dbo.TblLockedOrderMaintenance LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblLockedOrderMaintenance.SuperVM = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblLockedOrderMaintenance.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAqar ON dbo.TblLockedOrderMaintenance.AqrID = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblLockedOrderMaintenance.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblLockedOrderMaintenance.id = " & val(XPTxtID.Text) & ")"



        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepLocorderMaintAqar.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepLocorderMaintAqar.rpt"
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
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    'xReport.ParameterFields(3).AddCurrentValue user_name
     '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtForRenter.text), "0.00"), 0, True, ".")
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



Private Sub DcboEmpNameSuper_Change()
DcboEmpNameSuper_Click (0)
End Sub

Private Sub DcboEmpNameSuper_Click(Area As Integer)

   If val(DcboEmpNameSuper.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpNameSuper.BoundText, EmpCode
    TxtSearchCodeSuper.Text = EmpCode
End Sub

Private Sub ENDDATE_Change()
If Me.TxtModFlg.Text <> "R" Then
         EndDateH.value = ToHijriDate(EndDate.value)
        
End If
End Sub

Private Sub ENDDATEH_LostFocus()

       If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            EndDate.value = ToGregorianDate(EndDateH.value)
            End If
End Sub









Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub LockDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         LockDateH.value = ToHijriDate(LockDate.value)
End If
End Sub

Private Sub LockDateH_LostFocus()
   If Me.TxtModFlg.Text <> "R" Then
             VBA.Calendar = vbCalGreg
           LockDate.value = ToGregorianDate(LockDateH.value)
           End If
End Sub

Private Sub NourHijriCal1_LostFocus()
      If Me.TxtModFlg.Text <> "R" Then
             VBA.Calendar = vbCalGreg
           XPDtbTrans.value = ToGregorianDate(NourHijriCal1.value)
           End If
                End Sub




Private Sub Form_Load()
3
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

   
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmployees Me.DcboEmpNameSuper
   ' Dcombos.getAkarUnit Me.DcbUnitType
  '  Dcombos.GetIqarUnit -2, 1, DcbUnitNo
    'Dcombos.GetCustomersSuppliers 1, Me.dcCustomer
     My_SQL = "select UserID,UserName From tblUsers "
  
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
   fill_combo DCboUserName, My_SQL
    Dcombos.GetIqar DcbIqara
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblLockedOrderMaintenance     Order By ID"
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






Private Sub SatarDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         SatarDateH.value = ToHijriDate(SatarDate.value)
End If
End Sub

Private Sub SatarDateH_LostFocus()
   If Me.TxtModFlg.Text <> "R" Then
             VBA.Calendar = vbCalGreg
           SatarDate.value = ToGregorianDate(SatarDateH.value)
           End If
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
               
          '  UnitsGrid.Clear flexClearScrollable, flexClearEverything
          '  UnitsGrid.Rows = 2
          '  UnitsGrid.Enabled = True
         
    
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



Private Sub TxtOrderNo_Change()
If Me.TxtModFlg <> "R" Then
If Me.TxtOrderNo.Text <> "" Then
RetriveOrder TxtOrderNo.Text
End If
End If
End Sub

Private Sub TxtOrderNo_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg <> "R" Then
If KeyCode = vbKeyF3 Then

 Load FrmOrderMaintAqarSearch
 FrmOrderMaintAqarSearch.m_RetrunType = 1
 FrmOrderMaintAqarSearch.show
End If

End If
End Sub

Private Sub TxtSearchCodeSuper_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCodeSuper.Text, EmpID
        DcboEmpNameSuper.BoundText = EmpID
    End If
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
Public Sub RetriveOrderDetails(Optional Lngid As Long = 0)

    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

  If Lngid <> 0 Then
 Set RsDetails = New ADODB.Recordset
 StrSQL = "select * From TblOrderMaintenanceDet     where (ORderID = " & val(Lngid) & ") "
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If RsDetails.RecordCount > 0 Then
    RsDetails.MoveFirst
    For i = 1 To RsDetails.RecordCount
    SaveUoitInformation val(IIf(IsNull(RsDetails("UnitNo").value), 0, val(RsDetails("UnitNo").value))), val(IIf(IsNull(RsDetails("UnitStatus").value), 0, val(RsDetails("UnitStatus").value))), val(IIf(IsNull(RsDetails("RenterID").value), 0, val(RsDetails("RenterID").value)))
    RsDetails.MoveNext
   Next i
    End If
    End If
   End Sub
Public Sub RetriveOrder(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
  Dim ContactTime As Date

  If Lngid <> 0 Then
 Set RsDetails = New ADODB.Recordset
 StrSQL = "select * From TblOrderMaintenance     where id = " & val(Lngid) & "and(Lock<>1)"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If RsDetails.EOF Or RsDetails.BOF Then
        Exit Sub
    Else

       


    XPTxtID.Text = IIf(IsNull(RsDetails("ID").value), "", val(RsDetails("ID").value))

     Dcbranch.BoundText = val(IIf(IsNull(RsDetails("BranchID").value), "", RsDetails("BranchID").value))
  Me.DcboEmpName.BoundText = val(IIf(IsNull(RsDetails("EmpID").value), "", RsDetails("EmpID").value))
    
    DcbIqara.BoundText = val(IIf(IsNull(RsDetails("AqrID").value), "", RsDetails("AqrID").value))
     Me.DcboEmpNameSuper.BoundText = val(IIf(IsNull(RsDetails("SuperVM").value), "", RsDetails("SuperVM").value))
 Me.txtLocation.Text = IIf(IsNull(RsDetails("LocationIqar").value), "", RsDetails("LocationIqar").value)

   Me.TxtOrderWork.Text = IIf(IsNull(RsDetails("Des").value), "", RsDetails("Des").value)
 
  Me.DcbDMY.ListIndex = val(IIf(IsNull(RsDetails("DMY").value), -1, RsDetails("DMY").value))
  Me.TxtCount.Text = val(IIf(IsNull(RsDetails("Cont").value), 0, RsDetails("Cont").value))
     EndDate.value = IIf(IsNull(RsDetails("EndFate").value), Date, RsDetails("EndFate").value)
    EndDateH.value = IIf(IsNull(RsDetails("EndFateH").value), "", RsDetails("EndFateH").value)
    SatarDate.value = IIf(IsNull(RsDetails("RecDate").value), Date, RsDetails("RecDate").value)
    SatarDateH.value = IIf(IsNull(RsDetails("RecDateH").value), "", RsDetails("RecDateH").value)
End If

End If
    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
  Dim ContactTime As Date
'UnitsGrid.Clear flexClearScrollable, flexClearEverything
    '        UnitsGrid.Rows = 2
    '        UnitsGrid.Enabled = True
         
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
    XPDtbTrans.value = IIf(IsNull(rs("RecDate").value), Date, rs("RecDate").value)
    NourHijriCal1.value = IIf(IsNull(rs("RecDateH").value), "", rs("RecDateH").value)
       LockDate.value = IIf(IsNull(rs("LockDate").value), Date, rs("LockDate").value)
    LockDateH.value = IIf(IsNull(rs("LockDateH").value), "", rs("LockDateH").value)
   Me.TxtOrderNo.Text = IIf(IsNull(rs("OrderNo").value), "", rs("OrderNo").value)
     Dcbranch.BoundText = val(IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value))
  Me.DcboEmpName.BoundText = val(IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value))
    
    DcbIqara.BoundText = val(IIf(IsNull(rs("AqrID").value), "", rs("AqrID").value))
     Me.DcboEmpNameSuper.BoundText = val(IIf(IsNull(rs("SuperVM").value), "", rs("SuperVM").value))
 Me.txtLocation.Text = IIf(IsNull(rs("LocationIqar").value), "", rs("LocationIqar").value)
 Me.txtremark.Text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
   Me.TxtOrderWork.Text = IIf(IsNull(rs("OrderWork").value), "", rs("OrderWork").value)
   Me.DcbStatus.ListIndex = val(IIf(IsNull(rs("Status").value), -1, rs("Status").value))
  Me.DcbDMY.ListIndex = val(IIf(IsNull(rs("DMY").value), -1, rs("DMY").value))
  Me.TxtCount.Text = val(IIf(IsNull(rs("Cont").value), 0, rs("Cont").value))
     EndDate.value = IIf(IsNull(rs("EndFate").value), Date, rs("EndFate").value)
    EndDateH.value = IIf(IsNull(rs("EndFateH").value), "", rs("EndFateH").value)
    SatarDate.value = IIf(IsNull(rs("SatarDate").value), Date, rs("SatarDate").value)
    SatarDateH.value = IIf(IsNull(rs("SatarDateH").value), "", rs("SatarDateH").value)
 If rs("Lock").value = True Then
 Me.ChLock.value = vbChecked
 Else
 Me.ChLock.value = vbUnchecked
 End If
    
    ' )
     ' TxtOFRenter.text = val(IIf(IsNull(rs("OFRenter").value), 0, rs("OFRenter").value))
    '
    ' TxtBillPrice.text = val(IIf(IsNull(rs("BillPrice").value), 0, rs("BillPrice").value))
    ' TxtAccountNo.text = IIf(IsNull(rs("AccountNo").value), "", rs("AccountNo").value)
  ' TxtDayLate.text = IIf(IsNull(rs("DayNo").value), "", rs("DayNo").value)
    ' TxtAmountDely.text = IIf(IsNull(rs("AmountDely").value), "", rs("AmountDely").value)

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
   
   
'    Set RsDetails = New ADODB.Recordset
'StrSQL = "SELECT     dbo.TblOrderMaintenanceDet.ORderID, dbo.TblOrderMaintenanceDet.TypeUnit, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblOrderMaintenanceDet.UnitNo, "
' StrSQL = StrSQL & "                     dbo.TblAqarDetai.unitno AS UnitNoName, dbo.TblOrderMaintenanceDet.UnitStatus, dbo.TblRentStatus.name AS NameStatus,"
'StrSQL = StrSQL & "                      dbo.TblRentStatus.namee AS NameStatusE, dbo.TblOrderMaintenanceDet.RenterID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
'StrSQL = StrSQL & "                      dbo.TblOrderMaintenanceDet.mobile , dbo.TblOrderMaintenanceDet.Ms"
'StrSQL = StrSQL & " FROM         dbo.TblOrderMaintenanceDet LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblOrderMaintenanceDet.RenterID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblRentStatus ON dbo.TblOrderMaintenanceDet.UnitStatus = dbo.TblRentStatus.id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblOrderMaintenanceDet.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblOrderMaintenanceDet.TypeUnit = dbo.TblAkarUnit.id"
'StrSQL = StrSQL & " Where (dbo.TblOrderMaintenanceDet.ORderID = " & val(XPTxtID.text) & ")"
'
'    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'
'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'       With Me.UnitsGrid
'      '  RsDetails.MoveFirst
'        .Rows = .FixedRows + RsDetails.RecordCount
'
'        For i = .FixedRows To .Rows - 1
'
'            .TextMatrix(i, .ColIndex("Ser")) = i
'            .TextMatrix(i, .ColIndex("id")) = val(RsDetails("UnitNo").value)
'            .TextMatrix(i, .ColIndex("unittype")) = RsDetails("TypeUnit").value
'              .TextMatrix(i, .ColIndex("mobile")) = RsDetails("Mobile").value
'              If RsDetails("Ms").value = True Then
'               .TextMatrix(i, .ColIndex("Ms")) = -1
'              Else
'              .TextMatrix(i, .ColIndex("Ms")) = 0
'              End If
'               .TextMatrix(i, .ColIndex("StatusId")) = val(RsDetails("UnitStatus").value)
'               .TextMatrix(i, .ColIndex("customerid")) = val(RsDetails("RenterID").value)
'                .TextMatrix(i, .ColIndex("unitno")) = RsDetails("UnitNoName").value
'                If SystemOptions.UserInterface = EnglishInterface Then
'
'                .TextMatrix(i, .ColIndex("nameunittype")) = RsDetails("namee").value
'                .TextMatrix(i, .ColIndex("customeridname")) = RsDetails("CusNamee").value
'               .TextMatrix(i, .ColIndex("Status")) = RsDetails("NameStatusE").value
'               Else
'
'                .TextMatrix(i, .ColIndex("nameunittype")) = RsDetails("name").value
'                .TextMatrix(i, .ColIndex("customeridname")) = RsDetails("CusName").value
'               .TextMatrix(i, .ColIndex("Status")) = RsDetails("NameStatus").value
'               End If
'            RsDetails.MoveNext
'
'        Next i
    'ReLineGridCount
'    ReLineGrid
    
'End With

'    End If

'    RsDetails.Close
'    Set RsDetails = Nothing
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
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim sql As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
 
   If Me.DcbIqara.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ   «”„ «·⁄ﬁ«—!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.DcbIqara.SetFocus
           ' SendKeys "{F4}"
            Exit Sub
        End If
   


        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then

            XPTxtID.Text = CStr(new_id("TblLockedOrderMaintenance", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
            StrSQL = "Delete From TblUnitNoInformation Where LocOrderMaint=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If
     
   

   
       rs("ID").value = val(XPTxtID.Text)
        rs("RecDate").value = XPDtbTrans.value
       rs("RecDateH").value = Me.NourHijriCal1.value
         rs("LockDate").value = LockDate.value
       rs("LockDateH").value = Me.LockDateH.value
       rs("OrderNo").value = Me.TxtOrderNo.Text
       rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
       rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)
       rs("EmpID").value = IIf(Me.DcboEmpName.BoundText = "", Null, Me.DcboEmpName.BoundText)
       rs("AqrID").value = IIf(Me.DcbIqara.BoundText = "", Null, Me.DcbIqara.BoundText)
          rs("SuperVM").value = IIf(Me.DcboEmpNameSuper.BoundText = "", Null, Me.DcboEmpNameSuper.BoundText)
         rs("LocationIqar").value = Me.txtLocation.Text
         rs("Remark").value = Me.txtremark.Text
         rs("OrderWork").value = Me.TxtOrderWork.Text
          rs("DMY").value = IIf(Me.DcbDMY.ListIndex = -1, Null, Me.DcbDMY.ListIndex)
           rs("Status").value = IIf(Me.DcbStatus.ListIndex = -1, Null, Me.DcbStatus.ListIndex)
     rs("Cont").value = val(Me.TxtCount.Text)
      rs("EndFate").value = EndDate.value
       rs("EndFateH").value = Me.EndDateH.value
       
       rs("SatarDate").value = SatarDate.value
       rs("SatarDateH").value = Me.SatarDateH.value
   If ChLock.value = vbChecked Then
    sql = "update TblOrderMaintenance set   LockDate =  " & SQLDate(LockDate.value, True) & " , Lock=1 ,LockDateH='" & LockDateH.value & "'  where ID =" & val(Me.TxtOrderNo.Text) & ""
                                    Cn.Execute sql
  rs("Lock").value = 1
   Else
   rs("Lock").value = 0
    End If
        rs.update
        '''''''''/////////////////////////////////
  
    '  Set RsDetails = New ADODB.Recordset
    '   StrSQL = "SELECT     *  from dbo.TblOrderMaintenanceDet Where (1 = -1)"
  ' RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

          
  '     For i = Me.UnitsGrid.FixedRows To UnitsGrid.Rows - 1
  ' If (UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("unitno"))) <> "" Then
  '  RsDetails.AddNew
  '                RsDetails("ORderID").value = val(XPTxtID.text)
  '               RsDetails("TypeUnit").value = val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("unittype")))
  '              RsDetails("UnitNo").value = val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("id")))
  '              RsDetails("RenterID").value = val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("customerid")))
  '              RsDetails("UnitStatus").value = val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("StatusId")))
  '              RsDetails("Mobile").value = UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("mobile"))
  '              If UnitsGrid.Cell(flexcpChecked, i, Me.UnitsGrid.ColIndex("Ms")) = flexChecked Then
  '      RsDetails("Ms").value = -1
  '      Else
  '      RsDetails("Ms").value = 0
  '         End If
  '
  '
  '       RsDetails.update
  '
  '    '   End If
  '    End If
  '      Next i
  '
  '      '''''''''''''''//////////////////////////
        If Me.TxtOrderNo.Text <> "" Then
 RetriveOrderDetails TxtOrderNo.Text
End If

      

    
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
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
Private Sub DcboEmpName_Click(Area As Integer)
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
End Sub
Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub
Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode

   ' dcsupplier.BoundText = ownerid
    'DcbUnitType_Change
End Sub
Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub
Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
End Sub

'
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
                  StrSQL = "Delete From TblUnitNoInformation Where LocOrderMaint=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblLockedOrderMaintenance Where ID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
' StrSQL1 = "Delete From TblOrderMaintenanceDet Where ORderID=" & val(Me.XPTxtID.text)
'            Cn.Execute StrSQL1, , adExecuteNoRecords
            '    If rs.RecordCount < 1 Then
             
           ' UnitsGrid.Clear flexClearScrollable, flexClearEverything
           ' UnitsGrid.Rows = 2
           ' UnitsGrid.Enabled = True
            
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
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
''        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
 '          If SystemOptions.UserInterface = ArabicInterface Then
 '           GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
 '         Else
 '            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
 '         End If
 '           If SystemOptions.UserInterface = ArabicInterface Then
 '           GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
 '           Else
 '           GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
 '           End If
 '           GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
 '         GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
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
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "≈ﬁ›«· ÿ·» ’Ì«‰Â    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈ﬁ›«· ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈ﬁ›«· ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈ﬁ›«· ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈ﬁ›«· ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈ﬁ›«· ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "≈ﬁ›«· ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈ﬁ›«· ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈ﬁ›«· ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈ﬁ›«· ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈ﬁ›«· ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
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

