VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmCommisRece 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘… «·⁄„Ê·«  «·„” Õﬁ…"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17265
   Icon            =   "frmCommisRece.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   17265
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   3840
      TabIndex        =   74
      Top             =   6600
      Width           =   5055
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   5
         Left            =   0
         Picture         =   "frmCommisRece.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   4
         Left            =   720
         Picture         =   "frmCommisRece.frx":07D5
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   3
         Left            =   1440
         Picture         =   "frmCommisRece.frx":0D2D
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   2
         Left            =   2880
         Picture         =   "frmCommisRece.frx":11E6
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   1
         Left            =   3240
         Picture         =   "frmCommisRece.frx":16B6
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "frmCommisRece.frx":1B57
         Height          =   555
         Index           =   0
         Left            =   4320
         Picture         =   "frmCommisRece.frx":8E89
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "frmCommisRece.frx":9430
         Height          =   555
         Index           =   6
         Left            =   3600
         Picture         =   "frmCommisRece.frx":10762
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "frmCommisRece.frx":10C03
         Height          =   555
         Index           =   7
         Left            =   2160
         Picture         =   "frmCommisRece.frx":17F35
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   18480
      TabIndex        =   34
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   18720
      TabIndex        =   33
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   15000
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
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      Left            =   18420
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
      Width           =   17595
      _cx             =   31036
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
      Caption         =   "‘«‘… «·⁄„Ê·«  «·„” Õﬁ… "
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
         ButtonImage     =   "frmCommisRece.frx":187C5
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
         ButtonImage     =   "frmCommisRece.frx":18B5F
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
         ButtonImage     =   "frmCommisRece.frx":18EF9
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
         ButtonImage     =   "frmCommisRece.frx":19293
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
         Left            =   6120
         Picture         =   "frmCommisRece.frx":1962D
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
         TabIndex        =   32
         Top             =   480
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   13020
      TabIndex        =   7
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   95748097
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
      Left            =   18720
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
      Left            =   18840
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
      Bindings        =   "frmCommisRece.frx":1D295
      Height          =   315
      Left            =   8520
      TabIndex        =   30
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
      Height          =   5535
      Left            =   0
      TabIndex        =   37
      Top             =   1080
      Width           =   17280
      _cx             =   30480
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
      Caption         =   "«·⁄„Ê·«  «·„” Õﬁ…|Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "frmCommisRece.frx":1D2AA
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5070
         Left            =   17925
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   45
         Width           =   17190
         _cx             =   30321
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
            FormatString    =   $"frmCommisRece.frx":1D644
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
         Width           =   17190
         _cx             =   30321
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
         _GridInfo       =   $"frmCommisRece.frx":1D790
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
            Width           =   17160
            _cx             =   30268
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
               Height          =   3795
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   1320
               Width           =   17160
               Begin VSFlex8Ctl.VSFlexGrid fg 
                  Height          =   2955
                  Left            =   0
                  TabIndex        =   56
                  Top             =   120
                  Width           =   17160
                  _cx             =   30268
                  _cy             =   5212
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
                  Cols            =   14
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmCommisRece.frx":1D7C6
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
                  TabIndex        =   60
                  Top             =   3240
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
                  ButtonImage     =   "frmCommisRece.frx":1D9E5
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   13
                  Left            =   9960
                  TabIndex        =   61
                  Top             =   3240
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
                  ButtonImage     =   "frmCommisRece.frx":1DF7F
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   285
                  Index           =   11
                  Left            =   360
                  TabIndex        =   70
                  Top             =   3240
                  Width           =   2325
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   285
                  Index           =   10
                  Left            =   3480
                  TabIndex        =   69
                  Top             =   3240
                  Width           =   645
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Height          =   1245
               Index           =   11
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   0
               Width           =   17160
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   10920
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   600
                  Width           =   855
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Œ «— «·›‰Ì"
                  Height          =   210
                  Index           =   1
                  Left            =   15240
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   600
                  Width           =   1905
               End
               Begin VB.OptionButton XPOptShowType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂ· «·›‰ÌÌ‰"
                  Height          =   210
                  Index           =   0
                  Left            =   14760
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   960
                  Value           =   -1  'True
                  Width           =   2385
               End
               Begin MSDataListLib.DataCombo DcItem1 
                  Height          =   315
                  Left            =   9600
                  TabIndex        =   62
                  Top             =   600
                  Width           =   4935
                  _ExtentX        =   8705
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DtpaFrom 
                  Height          =   315
                  Left            =   13080
                  TabIndex        =   64
                  Top             =   240
                  Width           =   2295
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   95748097
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker DateTo 
                  Height          =   315
                  Left            =   9600
                  TabIndex        =   66
                  Top             =   240
                  Width           =   2535
                  _ExtentX        =   4471
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   95748097
                  CurrentDate     =   45292
               End
               Begin ImpulseButton.ISButton BtonAdd 
                  Height          =   390
                  Left            =   8760
                  TabIndex        =   73
                  Top             =   600
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
                  ButtonImage     =   "frmCommisRece.frx":1E519
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "ÌÃ» «Œ Ì«— «· «—ÌŒ «Ê·« À„ «Œ Ì«— «·›‰ÌÌ‰"
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
                  Height          =   1020
                  Index           =   9
                  Left            =   0
                  TabIndex        =   68
                  Top             =   120
                  Width           =   8655
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï  «—ÌŒ"
                  Height          =   285
                  Index           =   5
                  Left            =   11850
                  TabIndex        =   67
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰  «—ÌŒ"
                  Height          =   285
                  Index           =   2
                  Left            =   15720
                  TabIndex        =   65
                  Top             =   255
                  Width           =   1365
               End
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   510
               Left            =   0
               TabIndex        =   49
               Top             =   4680
               Width           =   2460
               _ExtentX        =   4339
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
               TabIndex        =   57
               Top             =   15000
               Width           =   930
               _ExtentX        =   1640
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
               ButtonImage     =   "frmCommisRece.frx":1E8B3
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   10
               Left            =   0
               TabIndex        =   58
               Top             =   -3720
               Width           =   930
               _ExtentX        =   1640
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
               ButtonImage     =   "frmCommisRece.frx":1EE4D
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   11
               Left            =   -120
               TabIndex        =   59
               Top             =   33960
               Width           =   900
               _ExtentX        =   1588
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
               ButtonImage     =   "frmCommisRece.frx":1F3E7
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
            Width           =   17160
            _cx             =   30268
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
               Left            =   4545
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   1005
               Width           =   900
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   2670
               Left            =   5685
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   1365
               Width           =   1470
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2670
               Index           =   67
               Left            =   3240
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1365
               Width           =   750
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   2520
               Index           =   68
               Left            =   5445
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
               Left            =   3990
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1365
               Width           =   555
            End
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   315
      Index           =   14
      Left            =   3120
      TabIndex        =   72
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… «·ﬁÌœ"
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
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   285
      Index           =   12
      Left            =   7320
      TabIndex        =   71
      Top             =   720
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·Õ—ﬂ…"
      Height          =   285
      Index           =   3
      Left            =   16200
      TabIndex        =   54
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
      Left            =   18090
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
      Left            =   12240
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
      Left            =   13950
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
      Left            =   18870
      TabIndex        =   18
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmCommisRece"
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
 Dim coun  As Integer
 Dim Account_Code_dynamic As String
 
 Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim pre, pre1 As Double
    Dim Msg As String
    Dim i As Integer
  Dim d As Double
    Fg.Rows = 150
        Fg.Enabled = True
StrSQL = " SELECT     dbo.TblMaintenanceWork.name, dbo.TblMaintenanceWork.namee, dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.[Value],"
 StrSQL = StrSQL & "                     dbo.TblCardAuthorizationReformDetails.[count], dbo.TblCardAuthorizationReformDetails.PriceFitter, dbo.TblCardAuthorizationReformDetails.allocation,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.PerceTage, dbo.TblEmployee.WorkShop_Job,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.DateExit, dbo.TblCardAuthorizationReformDetails.TimOut, dbo.TblCardAuthorizationReformDetails.ID,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.finish, dbo.TblCardAuthorizationReformDetails.DateEnter, dbo.TblCardAuthorizationReformDetails.TimeEnter,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.workshop, dbo.TblCardAuthorizationReformDetails.supervisor, dbo.TblEmployee.Emp_ID,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.ID2, dbo.TblCardAuthorizationReformDetails.Mainte, dbo.TblCardAuthorizationReform.PlateNo,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.CarTypeID, dbo.TBLCarTypes.name AS namecar, dbo.TBLCarTypes.namee AS nameecar, dbo.TblCarModels.Model,"
StrSQL = StrSQL & "                      dbo.TblCarModels.ModelE, dbo.TblCardAuthorizationReform.CarModelID, dbo.TblCardAuthorizationReformDetails.nohours,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.bill, dbo.TblCardAuthorizationReformDetails.comp, dbo.TblCardAuthorizationReformDetails.fitter,"
 StrSQL = StrSQL & "                     dbo.TblCardAuthorizationReformDetails.payed"
StrSQL = StrSQL & " FROM         dbo.TblEmployee INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails ON dbo.TblEmployee.Emp_ID = dbo.TblCardAuthorizationReformDetails.EmpID INNER JOIN"
 StrSQL = StrSQL & "                     dbo.TblMaintenanceWork ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblMaintenanceWork.Id INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform ON dbo.TblCardAuthorizationReformDetails.ID = dbo.TblCardAuthorizationReform.ID INNER JOIN"
 StrSQL = StrSQL & "                     dbo.TBLCarTypes ON dbo.TblCardAuthorizationReform.CarTypeID = dbo.TBLCarTypes.id INNER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCarModels ON dbo.TblCardAuthorizationReform.CarModelID = dbo.TblCarModels.Id"
StrSQL = StrSQL & " WHERE     (dbo.TblCardAuthorizationReformDetails.Type = 0) AND (dbo.TblCardAuthorizationReformDetails.allocation = 0) AND "
 StrSQL = StrSQL & "                     (dbo.TblCardAuthorizationReformDetails.finish = 1)"
'StrSQL = StrSQL & "  Where (dbo.TblCardAuthorizationReformDetails.Type = 0)and (dbo.TblCardAuthorizationReformDetails.allocation <>1) And (dbo.TblEmployee.WorkShop_Job = 2) And (dbo.TblCardAuthorizationReformDetails.finish = 1)"
 'StrSQL = StrSQL & "                       (dbo.TblCardAuthorizationReformDetails.finish = 1)"
   

   ' BolBegine = False
    StrWhere = ""

 



If XPOptShowType(1).value = True Then
If Me.DcItem1.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReformDetails.EmpID =" & Me.DcItem1.BoundText & ""
Else
MsgBox "ÌÃ» «Œ Ì«— ›‰Ì "
Fg.Rows = coun + 1
Exit Sub
End If
Else
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReformDetails.EmpID <>0 "


End If

    'If Not IsNull(Me.DtpDateFrom.value) Then
       ' If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReformDetails.DateExit >=" & SQLDate(Me.DtpaFrom.value, True) & ""
       ' Else
    '        BolBegine = True
    '        StrWhere = " Where dbo.TblCardAuthorizationReform.RecordDate >=" DateExit SQLDate(Me.DtpDateFrom.value, True) & ""
    '    End If
    'End If

    'If Not IsNull(Me.DtpDateTo.value) Then
    '    If BolBegine = True Then
      ' StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReformDetails.DateExit <=" & SQLDate(Me.DateTo.value, True) & ""
    '    Else
    '        BolBegine = True
     '  StrWhere = " Where  dbo.TblCardAuthorizationReform.RecordDate <=" & SQLDate(Me.DateTo.value, True) & ""
    '    End If
    'End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblCardAuthorizationReformDetails.ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
           ' Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Fg.Rows = coun + 1
        'Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else

        With Me.Fg
        '    .Clear flexClearScrollable, flexClearEverything
           ' .Rows = .FixedRows
           ' .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                'Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
              '  Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = 0 To rs.RecordCount - 1
            d = 0
             coun = coun + 1
              .TextMatrix(coun, .ColIndex("type")) = IIf(IsNull(rs("namecar").value), "", rs("namecar").value)
                .TextMatrix(coun, .ColIndex("model")) = IIf(IsNull(rs("Model").value), "", rs("Model").value)
                .TextMatrix(coun, .ColIndex("plateno")) = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
             
               .TextMatrix(coun, .ColIndex("serial")) = coun
                
                
                .TextMatrix(coun, .ColIndex("Emp_id")) = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)
                
                .TextMatrix(coun, .ColIndex("ID_Aut")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("DateExit").value)) Then
                    .TextMatrix(coun, .ColIndex("DateOp")) = Format(rs("DateExit").value, "yyyy/M/d")
                End If
            pre = val(rs("Value").value) * val(rs("count").value)
                .TextMatrix(coun, .ColIndex("Total")) = pre
                .TextMatrix(coun, .ColIndex("PriceFitter")) = val(IIf(IsNull(rs("PriceFitter").value), 0, rs("PriceFitter").value))
                .TextMatrix(coun, .ColIndex("Fitter")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(coun, .ColIndex("Operation")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                .TextMatrix(coun, .ColIndex("PerceTage")) = IIf(IsNull(rs("PerceTage").value), "", rs("PerceTage").value)
                d = pre - val(rs("PriceFitter").value)
                d = val(IIf(IsNull(rs("PerceTage").value), 0, rs("PerceTage").value) * (d / 100))
                .TextMatrix(coun, .ColIndex("PerceTageValue")) = d
                .TextMatrix(coun, .ColIndex("net")) = d + val(rs("PriceFitter").value)
                rs.MoveNext
            Next i


            .AutoSize 0, .Cols - 1, False
         '  Me.lbl(11).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("PerceTageValue"), .Rows - 1, .ColIndex("PerceTageValue"))
        End With

    End If
 Fg.Rows = coun + 1
 ReLineGrid
End Sub
'Private Sub Accredit_Click()
'    Dim BeginTrans As Boolean
'
'    Cn.BeginTrans
'    BeginTrans = True
'
'    If IsNull(rs("Posted")) Then
'        rs("Posted") = user_id
''        rs("PostedDate") = Time
 '   Else
 '       rs("Posted") = Null
 '      rs("PostedDate") = Time
 '   End If
 '
 '   rs.update
 'If SystemOptions.UserInterface = ArabicInterface Then
 '   Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
'Else
'Accredit.Caption = "Sent To approval "
'End If

  '  Cn.CommitTrans
 '   BeginTrans = False
'FillApprovedTable
'    Retrive (val(Me.XPTxtID.text))
'End Sub



Private Sub BtonAdd_Click()
'  Dim i As Integer
'  Dim j As Integer
'  Dim k As Integer
 
'  Dim Msg As String
'  Dim bool As Boolean
'  Dim rs1 As ADODB.Recordset
'  Dim sql As String

GetData


End Sub

Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index
Case 14
ShowGL_cc Me.TxtNoteSerial.text, , 200, val(Me.TxtNoteID.text)

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If
          coun = 0

 Fg.Clear flexClearScrollable, flexClearEverything
 Fg.Rows = 2
 'Frame10.Enabled = True
 XPOptShowType(1).value = False
 XPOptShowType(0).value = True
            TxtModFlg.text = "N"
            clear_all Me
  'Me.DcbOrderStatus.ListIndex = 1
         XPOptShowType(0).value = True
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

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            'Frame10.Enabled = True
            'Frame11.Enabled = True
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
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
            Load FrmCommisSearch
            FrmCommisSearch.show

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 12
         RemoveGridRow
            Case 13
             Fg.Clear flexClearScrollable, flexClearEverything
 Fg.Rows = 2
            coun = 0
                 Case 9

            If DoPremis(Do_Print, Me.name, True) = False Then
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
MySQL = "SELECT     dbo.TblCommisReceDetails.ID_Aut, dbo.TblCommisReceDetails.DateOp, dbo.TblCommisReceDetails.Total, dbo.TblCommisReceDetails.Fitter, "
 MySQL = MySQL & "                     dbo.TblCommisReceDetails.Operation, dbo.TblCommisReceDetails.PerceTage, dbo.TblCommisReceDetails.PerceTageValue, dbo.TblCommisReceDetails.id2,"
 MySQL = MySQL & "                     dbo.TblCommisRece.id, dbo.TblCommisRece.FitterID, dbo.TblCommisRece.DateFrom, dbo.TblCommisRece.DateTo, dbo.TblCommisRece.RecordDate,"
MySQL = MySQL & "                      dbo.TblCommisRece.AllFit, dbo.TblCommisRece.LimitFit, dbo.TblCommisRece.UserID, dbo.TblCommisReceDetails.id AS idd, dbo.TblCommisReceDetails.PriceFitter,"
MySQL = MySQL & "                      dbo.TblCommisReceDetails.Emp_id , dbo.TblCommisReceDetails.plateno, dbo.TblCommisReceDetails.Type, dbo.TblCommisReceDetails.Model"
MySQL = MySQL & " FROM         dbo.TblCommisRece INNER JOIN"
MySQL = MySQL & "                      dbo.TblCommisReceDetails ON dbo.TblCommisRece.id = dbo.TblCommisReceDetails.id2"
MySQL = MySQL & " Where (dbo.TblCommisRece.id =" & val(XPTxtID.text) & ")"
'MySQL = MySQL & " Where (dbo.TblCommisRece.id =" & val(XPTxtID.text) & ")"

 
 'MySQL = MySQL & "   Where (dbo.TblTreatment.id =" & val(XPTxtID.text) & ")"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCommisRece.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCommisRece.rpt"
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


 


Private Sub Dcbranch_Change()
If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""

End Sub

'Private Sub ImgFavorites_Click()
'AddTofaforites Me.name, Me.Caption, Me.Caption

'End Sub

Private Sub menue_Click(Index As Integer)
showsforms Index
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim ItemID As Integer

    If KeyAscii = vbKeyReturn Then
        GetItemIDFromCode TxtSearchCode.text, ItemID
        DcItem1.BoundText = ItemID
    End If

End Sub

 
'
''Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

'    If KeyCode = vbKeyF3 Then
'        FrmEmployeeSearch.lbltype = 9
''        FrmEmployeeSearch.show
 '
 '   End If

'End Sub

'Private Sub DcboEmpName_Click(Area As Integer)
'   On Error Resume Next
''       If val(DcboEmpName.BoundText) = 0 Then Exit Sub
'
'    Dim EmpCode  As String
 
''    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
 '   TxtSearchCode.text = EmpCode
 '
 '  If Me.TxtModFlg = "R" Then Exit Sub
   
 ''
  '  Dim StrSQL As String
'
'
        
        
'        Dim issuedate As Date
'        Dim depid As Double
'        Dim specid As Double
''        Dim JobTypeID As Double
 '       Dim gradeID As Double
 '       Dim Account_code2 As String
 '          Dim Account_Code  As String
 '       Dim Balance As String
 '       Dim endContractPerMonth As Double
 '       Dim national As String
 '       Dim project As Integer
  '      Dim pasid As String
 ''     Dim iqamaid As String
  '    Dim placeiqama As String
  '    Dim endiq As String
  '      get_employee_information val(Me.DcboEmpName.BoundText), issuedate, depid, specid, JobTypeID, gradeID, Account_code2, Account_Code, endContractPerMonth, national, , , project, pasid, iqamaid, placeiqama, , endiq
        
  '    WriteCustomerBalPublic Account_code2, Balance
          
  'lbl(22).Caption = val(Balance)

  '       WriteCustomerBalPublic Account_Code, Balance
          
 ' lbl(21).Caption = val(Balance)
 ' l'bl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
     '   DBIssueDate.value = issuedate
  ' DcboEmpDepartments.BoundText = project
     '   DcboSpecifications.BoundText = gradeID
  '   Me.TxtIqFrom.text = placeiqama
  '   DcbEmpNation.text = national
  '      DcboJobsType.BoundText = JobTypeID
  '      TxtIqama.text = iqamaid
  '      Me.XpDtEnd.value = endiq
  '     TxtPas.text = pasid
        
     '   lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
        
    'End If

'End Sub



'Private Sub Command1_Click()
'  Dim i As Integer
'  Dim j As Integer
'  Dim k As Integer
 
'  Dim Msg As String
''  Dim bool As Boolean
 ' Dim rs1 As ADODB.Recordset
 ' Dim sql As String
 ' bool = True
 '
 '     If ListStoreSelected.ListCount = 0 Then
 '      If SystemOptions.UserInterface = ArabicInterface Then
 '           Msg = "Õœœ     „Œ“‰ Ê«Õœ ⁄·Ï «·«ﬁ· " & Chr(13)
 '    Else
 '    Msg = "Select At Least One Store " & Chr(13)
 '    End If
 '           MsgBox Msg, vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 '
 '           SendKeys "{F4}"
 '           Exit Sub
 '       End If
 '       fg.Rows = 10000
 '       fg.Enabled = True
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
 
 '   For i = 1 To ListStoreSelected.ListCount
 '   If XPOptShowType(2).value = True Then
 ''          coun = coun + 1
  '     fg.TextMatrix(count, fg.ColIndex("serial")) = coun
  ''      fg.TextMatrix(coun, fg.ColIndex("StoreName")) = ListStoreSelected.List(i - 1)
   '     fg.TextMatrix(coun, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i - 1)
   '              fg.TextMatrix(coun, fg.ColIndex("ItemName")) = Me.DcItem1.text
   '    fg.TextMatrix(coun, fg.ColIndex("ItemID")) = Me.DcItem1.BoundText
   '     End If
       
   '        If XPOptShowType(0).value = True Then
        

   ' Set rs1 = New ADODB.Recordset
   '        sql = " SELECT * from  TblItems "
   '        rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
''
 '   If rs1.RecordCount > 0 Then

 '       For j = 1 To rs1.RecordCount
'coun = coun + 1
'            If SystemOptions.UserInterface = ArabicInterface Then
           
'              fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(rs1("ItemName").value), "", rs1("ItemName").value)
'            Else
'                fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(rs1("ItemNamee").value), "", rs1("ItemNamee").value)
'            End If

'         fg.TextMatrix(coun, fg.ColIndex("ItemID")) = rs1("ItemID").value
           
                 
'       fg.TextMatrix(coun, fg.ColIndex("serial")) = coun
'        fg.TextMatrix(coun, fg.ColIndex("StoreName")) = ListStoreSelected.List(i - 1)
'        fg.TextMatrix(coun, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i - 1)
'        rs1.MoveNext
'        Next j

'    End If
       
       
       
'        End If
'          If XPOptShowType(1).value = True Then
'          For k = 1 To ListGroupSelected.ListCount

'    Set rs1 = New ADODB.Recordset
'           sql = " SELECT * from  TblItems where GroupID =" & ListGroupSelected.ItemData(k - 1) & ""
'           rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

'    If rs1.RecordCount > 0 Then

'        For j = 1 To rs1.RecordCount
'coun = coun + 1
'            If SystemOptions.UserInterface = ArabicInterface Then
           
'              fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(rs1("ItemName").value), "", rs1("ItemName").value)
'            Else
'                fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(rs1("ItemNamee").value), "", rs1("ItemNamee").value)
''            End If
'
'         fg.TextMatrix(coun, fg.ColIndex("ItemID")) = rs1("ItemID").value
''            fg.TextMatrix(coun, fg.ColIndex("GroupID")) = rs1("GroupID").value
 '             fg.TextMatrix(coun, fg.ColIndex("GroupName")) = ListGroupSelected.List(k - 1)
 '      fg.TextMatrix(coun, fg.ColIndex("serial")) = coun
 '       fg.TextMatrix(coun, fg.ColIndex("StoreName")) = ListStoreSelected.List(i - 1)
 '       fg.TextMatrix(coun, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i - 1)
 ''       rs1.MoveNext
  '      Next j
'
'    End If
       
       
'         Next k
'        End If
'    Next i
'    If XPOptShowType(0).value = True Or XPOptShowType(1).value = True Then
'    fg.Rows = coun + 1
'    End If
'    ReLineGrid
'End Sub

'Private Sub Label2_Click()
'    Dim i As Integer
'    ListStoreSelected.Clear
'''
' '   For i = 0 To ListStoreall.ListCount - 1
 '       ListStoreSelected.AddItem ListStoreall.List(i)
 '       ListStoreSelected.ItemData(i) = ListStoreall.ItemData(i)
 '   Next i
'
'End Sub
'Private Sub Label5_Click()
'
''    If ListGroupSelected.ListIndex > -1 Then
 '       ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
 '   End If
''
'End Sub
'Private Sub Label6_Click()
'    ListGroupSelected.Clear
'End Sub
'Private Sub Label7_Click()
'    Dim i As Integer
'    If Me.XPOptShowType(1).value = True Then
''    ListGroupSelected.Clear
'
'    For i = 0 To ListGroupAll.ListCount - 1
'        ListGroupSelected.AddItem ListGroupAll.List(i)
'        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
'    Next i
'End If
'End Sub
'Private Sub Label8_Click()
'If Me.XPOptShowType(1).value = True Then
'' If ListGroupAll.ListIndex > -1 Then
 ''   ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
             
  '  ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)
  '          End If
  '          End If
'End Sub
'Private Sub Label4_Click()
'
'    If ListStoreSelected.ListIndex > -1 Then
    
'        ListStoreSelected.RemoveItem ListStoreSelected.ListIndex
'    End If

'End Sub
'Private Sub Label3_Click()
'    ListStoreSelected.Clear
'End Sub

'Private Sub LblSelect_Click()
'If ListStoreall.ListIndex > -1 Then
'    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
'
'    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
'            End If
'End Sub

'Private Sub ListGroupAll_Click()
' If XPOptShowType(1).value = True Then
'        Frame11.Enabled = True
'    Else
'        Frame11.Enabled = False
'    End If
'End Sub

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
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
 'Dim count As Integer
 coun = 0
    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
 '  Frame10.Enabled = False
   ' Frame11.Enabled = False
Me.XPOptShowType(0).value = True
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
       'Me.DcbOrderStatus.AddItem "Cancel Link"
       'Me.DcbOrderStatus.AddItem " link"
      Else
       ' Me.DcbOrderStatus.AddItem " ≈·€«¡«·—»ÿ"
      ' Me.DcbOrderStatus.AddItem " —»ÿ"
    End If
    Resize_Form Me
    AddTip
   ' FillMylist
    Set Dcombos = New ClsDataCombos
      Dcombos.GetUsers Me.DCboUserName
  
    Dcombos.GetBranches Me.dcBranch
    ' Dcombos.GetItemsNames DcItem1, , , , True

    Dcombos.GetEmployees DcItem1



    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = False
    End If

    SetDtpickerDate Me.XPDtbTrans
    'YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblCommisRece     Order By id"
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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
'    Label1.Visible = False
Cmd(13).Caption = "Delete All"
Cmd(12).Caption = "Delete"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
lbl(10).Caption = "Total"
    Me.Caption = "Screen Commissions Receivable "
    lbl(9).Caption = "Must choose a date first and then chose technicians"
    EleHeader.Caption = Me.Caption
    lbl(3).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(4).Caption = "Branch"
    lbl(2).Caption = "From "
    lbl(5).Caption = "To "
    'Frame10.Caption = "Select Store"
    'Fra(11).Caption = "Select Items "
    'XPOptShowType(1).Caption = "A specific group chose Group"
    'XPOptShowType(2).Caption = "A specific Item chose Item"
    XPOptShowType(0).Caption = "All Technical"
    XPOptShowType(1).RightToLeft = False
XPOptShowType(1).Caption = "Select Technical"
    XPOptShowType(0).RightToLeft = False
   BtonAdd.Caption = "Add"
   
   Accredit.Caption = "Accredite"
   XPTab301.Caption = " Commissions Receivable"
'lbl(5).Caption = "Remarks"
   lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"

   With Me.Fg
        .TextMatrix(0, .ColIndex("serial")) = "NO"
        .TextMatrix(0, .ColIndex("net")) = "net"
        .TextMatrix(0, .ColIndex("PriceFitter")) = "PriceFitter"
        .TextMatrix(0, .ColIndex("ID_Aut")) = "No Matter"
        .TextMatrix(0, .ColIndex("DateOp")) = "Date of Operation"
         .TextMatrix(0, .ColIndex("Total")) = "Total"
        .TextMatrix(0, .ColIndex("Fitter")) = "Technical"
         .TextMatrix(0, .ColIndex("Operation")) = "Operation"
        .TextMatrix(0, .ColIndex("PerceTage")) = "Commission Rate"
    .TextMatrix(0, .ColIndex("PerceTageValue")) = "Commission Payable"
    .TextMatrix(0, .ColIndex("type")) = "Type"
         .TextMatrix(0, .ColIndex("model")) = "Model"
        .TextMatrix(0, .ColIndex("plateno")) = "Plate No"
    End With

End Sub

'Private Sub YearMonth()
'
'    Dim i As Integer
'    Dim IntDefIndex As Integer

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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim RsDetails2 As ADODB.Recordset
   
    
    Dim i As Integer
    Dim StrSQL As String
  
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
            rs.find "id=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.text = IIf(IsNull(rs("id").value), "", val(rs("id").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    dcBranch.BoundText = IIf(IsNull(rs("BrnchID").value), "", rs("BrnchID").value)
   Me.DcItem1.BoundText = IIf(IsNull(rs("FitterID").value), "", rs("FitterID").value)
   Me.DtpaFrom.value = IIf(IsNull(rs("DateFrom").value), Date, rs("DateFrom").value)
   Me.DateTo.value = IIf(IsNull(rs("DateTo").value), Date, rs("DateTo").value)
   TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
  TxtNoteID.text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)

   
   If rs("AllFit").value = 0 Then
   Me.XPOptShowType(0).value = False
   Else
   Me.XPOptShowType(0).value = True
   End If
     If rs("LimitFit").value = 0 Then
   Me.XPOptShowType(1).value = False
   Else
   Me.XPOptShowType(1).value = True
   End If
'Me.DcbOrderStatus.ListIndex = rs("LinkType").value
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
  '   If IsNull(rs("posted").value) Then
   '                                                If SystemOptions.UserInterface = ArabicInterface Then
   '                                                Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
   '                                              Else
   '                                                Accredit.Caption = " send to Approval   "
   '                                            End If
   '                                            Accredit.Enabled = True
  'Else
   '                                               If SystemOptions.UserInterface = ArabicInterface Then
  '                                                  Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
  ''                                                Else
   '                                                Accredit.Caption = " sent to Approval   "
   '                                            End If
   '                                            Accredit.Enabled = False
   'End If
   
   
    Set RsDetails = New ADODB.Recordset
StrSQL = " select * from TblCommisReceDetails where id2 = " & val(XPTxtID.text) & ""

    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
       Fg.Rows = Fg.FixedRows + RsDetails.RecordCount


        For i = Me.Fg.FixedRows To Fg.Rows - 1
     
          Fg.TextMatrix(i, Fg.ColIndex("serial")) = i
           Fg.TextMatrix(i, Fg.ColIndex("ID_Aut")) = RsDetails("ID_Aut").value
           Fg.TextMatrix(i, Fg.ColIndex("Emp_id")) = RsDetails("Emp_id").value
               Fg.TextMatrix(i, Fg.ColIndex("plateno")) = RsDetails("plateno").value
          Fg.TextMatrix(i, Fg.ColIndex("type")) = RsDetails("type").value
            Fg.TextMatrix(i, Fg.ColIndex("model")) = RsDetails("model").value
           
            Fg.TextMatrix(i, Fg.ColIndex("DateOp")) = RsDetails("DateOp").value
          Fg.TextMatrix(i, Fg.ColIndex("Total")) = RsDetails("Total").value
            Fg.TextMatrix(i, Fg.ColIndex("Fitter")) = RsDetails("Fitter").value
            
            Fg.TextMatrix(i, Fg.ColIndex("Operation")) = RsDetails("Operation").value
          Fg.TextMatrix(i, Fg.ColIndex("PerceTage")) = val(RsDetails("PerceTage").value)
      Fg.TextMatrix(i, Fg.ColIndex("PerceTageValue")) = val(RsDetails("PerceTageValue").value)
     Fg.TextMatrix(i, Fg.ColIndex("PriceFitter")) = val(RsDetails("PriceFitter").value)
     Fg.TextMatrix(i, Fg.ColIndex("net")) = val(RsDetails("PriceFitter").value) + val(RsDetails("PerceTageValue").value)
            RsDetails.MoveNext
        Next i

    End If

     RsDetails.Close
    Set RsDetails = Nothing
   ' fillapprovData
    ReLineGrid
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
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
            rs.find "id='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

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
  


Function createVoucher()
    Dim bankDes As String
    Dim AccountCode As String
 
    Dim Employee_account As String
    Dim NoteID As String
    Dim sql As String
 
    '//////////////////////////////////////Notes////////////////////////////////////
    Dim line_no As Integer
    Dim RsNotes As New ADODB.Recordset
    RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   
    If Me.TxtModFlg.text = "E" Then
                  
        sql = "Delete notes where NoteID=" & val(Me.TxtNoteID.text)
Else

    End If

    RsNotes.AddNew
    NoteID = CStr(TxtNoteID.text)
    RsNotes("NoteID").value = CStr(TxtNoteID.text)
                    
    bankDes = "”‰œ «” Õﬁ«ﬁ ⁄„Ê·«     " '& DcComponentType.text & Chr(13)
                       
    bankDes = bankDes & "  „‰ «·› —… " & DtpaFrom.value & "  «·Ï «·› —… : " & DateTo.value
    RsNotes("NoteType").value = 5151
    RsNotes("NoteDate").value = XPDtbTrans.value
    RsNotes("UserID").value = user_id
    RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '????? ?????
 '   RsNotes("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '????? ??? ?????
    RsNotes("numbering_type").value = sand_numbering_type(0) '??? ????? ??? ?????
    RsNotes("numbering_type1").value = sand_numbering_type(51) '??? ????? ??? ????????
    RsNotes("sanad_year").value = year(XPDtbTrans.value)
    RsNotes("sanad_month").value = Month(XPDtbTrans.value)
    RsNotes("note_value_by_characters").value = WriteNo(Format(val(lbl(11).Caption), "0.00"), 0, True, ".")
    'RsNotes("remark").value = TxtRemarks.text & bankDes
    RsNotes("Branch_no").value = val(Me.dcBranch.BoundText)
                
    RsNotes.update
                
    line_no = 1

    If Fg.Rows > 1 And val(lbl(11).Caption) > 0 Then
        Dim RsDev  As ADODB.Recordset
        Set RsDev = New ADODB.Recordset
        RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     
        AccountCode = Account_Code_dynamic
                       
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.dcBranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = AccountCode
        RsDev("Value").value = Round(val(Me.lbl(11).Caption), 2)
        RsDev("Credit_Or_Debit").value = 0
                    
        RsDev("RecordDate").value = Me.XPDtbTrans.value
        RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
        RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
                        
        RsDev("UserID").value = user_id
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                         
        RsDev.update
    End If

    ' ??????
          
    If Fg.Rows > 1 And val(lbl(11).Caption) > 0 Then
 
        Dim i  As Integer
        Dim LngDevID  As Long

        With Fg
 
            For i = .FixedRows To .Rows - 1

                If .TextMatrix(i, .ColIndex("Emp_ID")) <> "" And val(.TextMatrix(i, .ColIndex("net"))) > 0 Then
                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '???? ????? ??? ????
                    AccountCode = Employee_account
   
                    line_no = line_no + 1
  
                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, Round(val(.TextMatrix(i, .ColIndex("net"))), 2), 1, "" & bankDes & " „‰ «·«„— —ﬁ„ " & .TextMatrix(i, .ColIndex("ID_Aut")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , .TextMatrix(i, .ColIndex("net")), , , , bankDes, , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    
                    End If
         
                End If

            Next i

        End With
    
    End If

    updateNotesValueAndNobytext (val(NoteID))

ErrTrap:

End Function
Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String
Dim sql As String
 '   On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
    If val(Fg.TextMatrix(1, Fg.ColIndex("ID_Aut"))) = 0 Then
MsgBox "ÌÃ» «‰  Õ ÊÌ «·‘«‘… ⁄·Ï »Ì«‰« "
Exit Sub
End If
    '    If Me.DcbCarType.BoundText = "" Then
    '        Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄  «·„⁄œÂ/«·”Ì«—…!! "
      '      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    ''        Me.DcbCarType.SetFocus
     '  '     SendKeys "{F4}"
     '       Exit Sub
     '   End If
  'If Me.TxtCliientName.text = "" Then
  '          Msg = "ÌÃ» «œŒ«· «”„ «·⁄„Ì·!! "
  '          MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  '          Me.TxtCliientName.SetFocus
  '         ' SendKeys "{F4}"
  '          Exit Sub
  '      End If
   my_branch = val(Me.dcBranch.BoundText)
   
    

            Account_Code_dynamic = get_account_code_branch(78, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
              If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                   MsgBox "branch Not Created", vbCritical
                End If

                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»    „’—Ê›«   ’Ì«‰… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox " Maintenance  Expenses  Account Not Defined in this Branch", vbCritical
                    End If

                    GoTo ErrTrap

                End If
            End If
            
 
    If TxtNoteSerial.text = "" Then
        If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else

            If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                '                       TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
            End If
        End If
    End If

Dim TxtNoteSerial1str As String

    If TxtNoteSerial1.text = "" Then
    TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbTrans.value, 51, 5151)

                If TxtNoteSerial1str = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…     ”‰œ ⁄„Ê·«  „” Õﬁ…  ÃœÌœ… ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If TxtNoteSerial1str = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—…  «·’Ì«‰…  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                        '             txtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, DCPreFix.text)
                    End If
                End If
    End If
     






        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("TblCommisRece", "ID", "", True))
               TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
               
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
            Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
           
TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)

  TxtNoteSerial1 = Voucher_coding(val(my_branch), XPDtbTrans.value, 51, 5151)
            
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From TblCommisReceDetails Where ID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords

   StrSQL = "Delete notes where NoteID=" & val(Me.TxtNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(TxtNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords


        End If

        rs("BrnchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
 
        rs("id").value = val(XPTxtID.text)



rs("NoteID").value = CStr(TxtNoteID.text)
    rs("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '????? ?????
    rs("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '????? ??? ?????
    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
                 
         rs("FitterID").value = IIf(Me.DcItem1.BoundText = "", Null, Me.DcItem1.BoundText)
        rs("DateFrom").value = Me.DtpaFrom.value
        rs("DateTo").value = Me.DateTo.value
         rs("RecordDate").value = XPDtbTrans.value
        
        If XPOptShowType(0).value = True Then
        rs("AllFit").value = 1
        Else
        rs("AllFit").value = 0
        End If
        If XPOptShowType(1).value = True Then
         rs("LimitFit").value = 1
        Else
         rs("LimitFit").value = 0
        End If
         
      
      '
        rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)
'
        rs.update
        '''''''''/////////////////////////////////
        
      Set RsDetails = New ADODB.Recordset
       RsDetails.Open "TblCommisReceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
If Fg.Rows > 1 Then
          
       For i = Me.Fg.FixedRows To Fg.Rows - 1
         If val(Fg.TextMatrix(i, Fg.ColIndex("ID_Aut"))) <> 0 Then
           RsDetails.AddNew

  
                                    
          RsDetails("id2").value = val(XPTxtID.text)
 RsDetails("plateno").value = Fg.TextMatrix(i, Fg.ColIndex("plateno"))
          RsDetails("type").value = Fg.TextMatrix(i, Fg.ColIndex("type"))
        RsDetails("model").value = Fg.TextMatrix(i, Fg.ColIndex("model"))
          RsDetails("Emp_id").value = val(Fg.TextMatrix(i, Fg.ColIndex("Emp_id")))
          RsDetails("ID_Aut").value = val(Fg.TextMatrix(i, Fg.ColIndex("ID_Aut")))
        RsDetails("Total").value = val(Fg.TextMatrix(i, Fg.ColIndex("Total")))
        RsDetails("Fitter").value = Fg.TextMatrix(i, Fg.ColIndex("Fitter"))
        RsDetails("Operation").value = Fg.TextMatrix(i, Fg.ColIndex("Operation"))
       
        RsDetails("DateOp").value = IIf(IsDate(Fg.TextMatrix(i, Fg.ColIndex("DateOp"))), Fg.TextMatrix(i, Fg.ColIndex("DateOp")), Null)
        RsDetails("PriceFitter").value = val(IIf((Fg.TextMatrix(i, Fg.ColIndex("PriceFitter"))), Fg.TextMatrix(i, Fg.ColIndex("PriceFitter")), 0))
     RsDetails("PerceTage").value = val(IIf((Fg.TextMatrix(i, Fg.ColIndex("PerceTage"))) <> "", Fg.TextMatrix(i, Fg.ColIndex("PerceTage")), 0))
        RsDetails("PerceTageValue").value = val(IIf((Fg.TextMatrix(i, Fg.ColIndex("PerceTageValue"))), Fg.TextMatrix(i, Fg.ColIndex("PerceTageValue")), 0))
  
         RsDetails.update
           sql = "update TblCardAuthorizationReformDetails set   allocation=1  where ID=" & val(Fg.TextMatrix(i, Fg.ColIndex("ID_Aut"))) & ""
           Cn.Execute sql
        End If
        Next i
        End If


    
        Cn.CommitTrans
        BeginTrans = False
       RsDetails.Close
     
        Set RsDetails = Nothing
'createVoucher

        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
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
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
Dim StrSQL1 As String
Dim sql As String
Dim i As Integer
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
            If Fg.Rows > 1 Then
          
       For i = Me.Fg.FixedRows To Fg.Rows - 1
             sql = "update TblCardAuthorizationReformDetails set   allocation=0  where ID=" & val(Fg.TextMatrix(i, Fg.ColIndex("ID_Aut"))) & ""
  
                                    Cn.Execute sql
                                    Next i
                                    End If
                rs.delete
                
                
                   StrSQL = "Delete notes where NoteID=" & val(Me.TxtNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(TxtNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords


'                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.text)
'                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
   
                StrSQL1 = "Delete From TblCommisReceDetails Where id2=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL1, , adExecuteNoRecords
              
                    clear_all Me
                      '  ListGroupSelected.Clear
   ' ListStoreSelected.Clear

                   Fg.Clear flexClearScrollable, flexClearEverything
                   Fg.Rows = 2
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                 
                End If
           ' End If
        End If
   Retrive
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub



'Function FillApprovedTable()
' Dim RSApproval  As New ADODB.Recordset
'   Set RSApproval = New ADODB.Recordset
'   Dim currentdate As Date
'   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'

' Dim sql As String
'  Dim rs1 As New ADODB.Recordset
' Dim i As Integer
'    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
 ' sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
''  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
 ' sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
 ' sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
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

''                 If i = 1 Then
  '                      RSApproval("Currcursor").value = 1
 ''                        RSApproval("FromUser").value = user_name
  '              End If
  '
  '              RSApproval.update
  '              rs1.MoveNext
  '          Next i
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
''
 '   RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
' If Not (RsDetails.EOF Or RsDetails.BOF) Then
''        GRID2.Rows = RsDetails.RecordCount + 1
 

 '       For Num = 1 To RsDetails.RecordCount
 '
 '      GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
 '   If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
 '  GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
 '  Else
 '   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
 ''   End If
  '
  '      GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
  '         If SystemOptions.UserInterface = ArabicInterface Then
  '          GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
  '        Else
  '           GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
  '        End If
  '          If SystemOptions.UserInterface = ArabicInterface Then
  '          GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
  '          Else
  '          GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
  '          End If
  '          GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
  '        GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 '
 
'RsDetails.MoveNext
'If Num = RsDetails.RecordCount Then

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

'End If

  '      Next Num
'Else
' GRID2.Rows = 1
'    End If
'RsDetails.Close

'End Function


Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap
     .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
    End With
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
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
    Me.lbl(11).Caption = 0
    IntCounter = 0

    With Fg

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("Fitter")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("serial")) = IntCounter
           Me.lbl(11).Caption = Me.lbl(11).Caption + val(.TextMatrix(i, .ColIndex("net")))
    
        End If
                

        Next i
 
    End With

End Sub
'Function FillMylist()
'    Dim sql As String
'    Dim rs As New ADODB.Recordset
'    Dim Balance As Double
'    Dim i As Integer
'    sql = " SELECT * from  TblStore"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'  '  ListStoreall.Clear
'   ' ListStoreSelected.Clear
'
'    If rs.RecordCount > 0 Then
'
'        For i = 1 To rs.RecordCount
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'            '    ListStoreall.AddItem IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
'            Else
'             '   ListStoreall.AddItem IIf(IsNull(rs("StoreNamee").value), "", rs("StoreNamee").value)
'            End If
'
'          '  ListStoreall.ItemData(ListStoreall.NewIndex) = rs("StoreID").value
'            rs.MoveNext
'        Next i
'
'    End If
'
'    rs.Close
'
'    'fil
'
'  sql = " SELECT * from  Groups where GroupID>1"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    ListGroupAll.Clear
'    ListGroupSelected.Clear
'
'    If rs.RecordCount > 0 Then
'
'        For i = 1 To rs.RecordCount
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'                ListGroupAll.AddItem IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
'            Else
'                ListGroupAll.AddItem IIf(IsNull(rs("GroupNamee").value), "", rs("GroupNamee").value)
'            End If

'            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("GroupID").value
'            rs.MoveNext
'        Next i
'
'    End If
'
'    rs.Close

'End Function
Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "‘«‘… «·⁄„Ê·«  «·„” Õﬁ…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… «·⁄„Ê·«  «·„” Õﬁ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ‘«‘… «·⁄„Ê·«  «·„” Õﬁ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… «·⁄„Ê·«  «·„” Õﬁ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… «·⁄„Ê·«  «·„” Õﬁ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… «·⁄„Ê·«  «·„” Õﬁ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… «·⁄„Ê·«  «·„” Õﬁ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… «·⁄„Ê·«  «·„” Õﬁ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… «·⁄„Ê·«  «·„” Õﬁ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… «·⁄„Ê·«  «·„” Õﬁ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… «·⁄„Ê·«  «·„” Õﬁ… ", 1, 15204351, -2147483630
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



 Private Sub RemoveGridRow()
coun = coun - 1
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
      '  Frame11.Enabled = True
      Me.DcItem1.Enabled = True
    Else
       DcItem1.Enabled = False
    End If
End Sub

