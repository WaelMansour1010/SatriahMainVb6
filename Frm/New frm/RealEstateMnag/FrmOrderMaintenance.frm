VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmOrderMaintenance 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·» ’Ì«‰Â"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   FillColor       =   &H00C0E0FF&
   Icon            =   "FrmOrderMaintenance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   11790
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CommandButton CMDSENDSMS 
      Caption         =   "«—”«· —”«·Â"
      Height          =   255
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   88
      Top             =   6240
      Width           =   1215
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
      TabIndex        =   68
      Top             =   3600
      Width           =   825
   End
   Begin VB.TextBox TxtDayPrice 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9480
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
      Caption         =   "ÿ·» ’Ì«‰Â "
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
         ButtonImage     =   "FrmOrderMaintenance.frx":038A
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
         ButtonImage     =   "FrmOrderMaintenance.frx":0724
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
         ButtonImage     =   "FrmOrderMaintenance.frx":0ABE
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
         ButtonImage     =   "FrmOrderMaintenance.frx":0E58
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
         Left            =   3360
         Picture         =   "FrmOrderMaintenance.frx":11F2
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
      Top             =   6660
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
      Top             =   6240
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
      Bindings        =   "FrmOrderMaintenance.frx":4E5A
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
      Height          =   4815
      Left            =   0
      TabIndex        =   36
      Top             =   1200
      Width           =   11760
      _cx             =   20743
      _cy             =   8493
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
      Picture(0)      =   "FrmOrderMaintenance.frx":4E6F
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   4350
         Left            =   12405
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   45
         Width           =   11670
         _cx             =   20585
         _cy             =   7673
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
            FormatString    =   $"FrmOrderMaintenance.frx":5209
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
         Height          =   4350
         Index           =   15
         Left            =   45
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   45
         Width           =   11670
         _cx             =   20585
         _cy             =   7673
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
         _GridInfo       =   $"FrmOrderMaintenance.frx":5355
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4320
            Index           =   16
            Left            =   15
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   15
            Width           =   11640
            _cx             =   20532
            _cy             =   7620
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
               Caption         =   "»Ì«‰«  «·ÿ·»"
               Height          =   4335
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   0
               Width           =   11580
               Begin VB.TextBox TxtMobile 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   5280
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   3840
                  Width           =   975
               End
               Begin VB.CheckBox ChLock 
                  Alignment       =   1  'Right Justify
                  Caption         =   " „ «·«ﬁ›«· » «—ÌŒ"
                  Enabled         =   0   'False
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   9840
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   3840
                  Width           =   1575
               End
               Begin VB.ComboBox DcbDMY 
                  Height          =   315
                  ItemData        =   "FrmOrderMaintenance.frx":5389
                  Left            =   6960
                  List            =   "FrmOrderMaintenance.frx":5396
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   3480
                  Width           =   1815
               End
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
                  Left            =   8760
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   3480
                  Width           =   1065
               End
               Begin VB.TextBox txtDes 
                  Alignment       =   1  'Right Justify
                  Height          =   675
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   74
                  Top             =   960
                  Width           =   10335
               End
               Begin VB.TextBox TxtSearchCodeSuper 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   9720
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   615
                  Width           =   735
               End
               Begin VB.TextBox TxtSearch 
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
                  Left            =   9720
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   240
                  Width           =   735
               End
               Begin VB.TextBox TxtLocation 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   615
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   240
                  Width           =   4215
               End
               Begin MSComCtl2.DTPicker EndDate 
                  Height          =   315
                  Left            =   3720
                  TabIndex        =   56
                  Top             =   3480
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   103153665
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker FilterDate 
                  Height          =   315
                  Left            =   8520
                  TabIndex        =   57
                  Top             =   3840
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
                  Left            =   1980
                  TabIndex        =   58
                  Top             =   3480
                  Width           =   1695
                  _ExtentX        =   2778
                  _ExtentY        =   556
               End
               Begin Dynamic_Byte.NourHijriCal FilterDateH 
                  Height          =   315
                  Left            =   6960
                  TabIndex        =   59
                  Top             =   3840
                  Width           =   1515
                  _ExtentX        =   2778
                  _ExtentY        =   556
               End
               Begin MSDataListLib.DataCombo DcbIqara 
                  Height          =   315
                  Left            =   5280
                  TabIndex        =   65
                  Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
                  Top             =   240
                  Width           =   4455
                  _ExtentX        =   7858
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboEmpNameSuper 
                  Height          =   315
                  Left            =   7440
                  TabIndex        =   72
                  Top             =   600
                  Width           =   2295
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VSFlex8UCtl.VSFlexGrid UnitsGrid 
                  Height          =   1725
                  Left            =   120
                  TabIndex        =   76
                  Top             =   1680
                  Width           =   11445
                  _cx             =   20188
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
                  Cols            =   27
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmOrderMaintenance.frx":53A9
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”‰œ «·’—›"
                     Height          =   1050
                     Index           =   40
                     Left            =   -1680
                     TabIndex        =   87
                     Top             =   0
                     Width           =   1440
                  End
               End
               Begin MSDataListLib.DataCombo DcboEmpName 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   81
                  Top             =   3840
                  Width           =   2655
                  _ExtentX        =   4683
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton Cmdd 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   85
                  Top             =   3840
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   661
                  ButtonPositionImage=   1
                  Caption         =   "«—”«· ·„”ƒÊ· «·’Ì«‰Â"
                  BackColor       =   14871017
                  ForeColor       =   -2147483635
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
                  ColorToggledText=   -2147483635
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   13
                  Left            =   240
                  TabIndex        =   86
                  Top             =   3480
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
                  ButtonImage     =   "FrmOrderMaintenance.frx":57C6
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÃÊ«· «·„”ƒÊ·"
                  Height          =   285
                  Index           =   9
                  Left            =   6360
                  TabIndex        =   83
                  Top             =   600
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„ÊŸ› «·’Ì«‰Â"
                  Height          =   285
                  Index           =   5
                  Left            =   5760
                  TabIndex        =   82
                  Top             =   3855
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ê’› «·’Ì«‰Â"
                  Height          =   255
                  Index           =   2
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   1080
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”ƒÊ· «·’Ì«‰Â"
                  Height          =   285
                  Index           =   3
                  Left            =   10410
                  TabIndex        =   73
                  Top             =   615
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„Êﬁ⁄ «·⁄ﬁ«—"
                  Height          =   255
                  Index           =   18
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ «‰Â«¡ «·ÿ·»"
                  Height          =   375
                  Index           =   20
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   3480
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "’·«ÕÌÂ «·ÿ·»"
                  Height          =   375
                  Index           =   17
                  Left            =   10200
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   3480
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄ﬁ«—"
                  Height          =   255
                  Index           =   13
                  Left            =   10680
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
               Height          =   2490
               Index           =   62
               Left            =   2220
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   1185
               Width           =   540
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4320
            Index           =   9
            Left            =   15
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   15
            Width           =   11640
            _cx             =   20532
            _cy             =   7620
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
               Height          =   3240
               Left            =   3045
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   930
               Width           =   630
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   2205
               Left            =   3855
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   1185
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2205
               Index           =   67
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1185
               Width           =   600
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   2160
               Index           =   68
               Left            =   3675
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   1470
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
               Height          =   2565
               Index           =   69
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1185
               Width           =   285
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
      Left            =   5760
      TabIndex        =   63
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
   End
   Begin MSComCtl2.DTPicker From 
      Height          =   315
      Left            =   12360
      TabIndex        =   67
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   103153665
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker TimeOrder 
      Height          =   315
      Left            =   3120
      TabIndex        =   69
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   103153666
      CurrentDate     =   38784
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Êﬁ  «·ÿ·»"
      Height          =   285
      Index           =   14
      Left            =   4500
      TabIndex        =   70
      Top             =   720
      Width           =   1005
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
      Picture         =   "FrmOrderMaintenance.frx":5D60
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
      Picture         =   "FrmOrderMaintenance.frx":6D84
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
      Left            =   10800
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
      Top             =   6315
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
      Top             =   6270
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
      Top             =   6270
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   19
      Top             =   6300
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   18
      Top             =   6300
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
Attribute VB_Name = "FrmOrderMaintenance"
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
      Msg = Msg & "Work was Order Maintenance No."
             Msg = Msg & Chr(13) & XPTxtID.Text
             Msg = Msg & "  Date of Order   "
       Msg = Msg & NourHijriCal1.value & "corresponding " & XPDtbTrans.value
       Msg = Msg & Chr(13)
       Msg = Msg & "  Maintenance Officer "
       Msg = Msg & DcboEmpNameSuper.Text
       Msg = Msg & Chr(13)
        Msg = Msg & " FITNESS Order   "
       Msg = Msg & TxtCount.Text & " " & DcbDMY.Text
       Msg = Msg & Chr(13)
      Else
      Msg = Msg & "   „ ⁄„· ÿ·» ’Ì«‰Â »—ﬁ„   "
       Msg = Msg & XPTxtID.Text
       Msg = Msg & Chr(13)
       Msg = Msg & "  «—ÌŒ «·ÿ·»   "
       Msg = Msg & NourHijriCal1.value & "«·„Ê«›ﬁ" & XPDtbTrans.value
       Msg = Msg & Chr(13)
       Msg = Msg & "  „”ƒÊ· «·’Ì«‰Â   "
       Msg = Msg & DcboEmpNameSuper.Text
       Msg = Msg & Chr(13)
        Msg = Msg & "   ’·«ÕÌ… «·ÿ·»   "
       Msg = Msg & TxtCount.Text & " " & DcbDMY.Text
       Msg = Msg & Chr(13)
End If
        Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TblUnitNoInformation Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      RsDetails1.AddNew
      RsDetails1("BranchId").value = Dcbranch.BoundText
      RsDetails1("CusID").value = cus
           RsDetails1("UnitNo").value = Unit
           RsDetails1("UnitStatus").value = Status
           RsDetails1("Des").value = Msg
           RsDetails1("RecDate").value = XPDtbTrans.value
           RsDetails1("RecDateH").value = NourHijriCal1.value
           RsDetails1("NoteID").value = Null
           RsDetails1("ContNo").value = Null
           RsDetails1("FilterNo").value = Null
           RsDetails1("OrderMaint").value = val(XPTxtID.Text)
           RsDetails1("LocOrderMaint").value = Null
           RsDetails1.update
           End If

   End Sub
Private Sub RemoveGridRow()

    With Me.UnitsGrid

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
    XPDtbTrans_Change
    FilterDate_Change
    ENDDATE_Change
Dcbranch.BoundText = Current_branch
  Me.DCboUserName.BoundText = user_id
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
 UnitsGrid.Rows = UnitsGrid.Rows + 1
            UnitsGrid.Enabled = True
          
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
SendMessage (1)

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
          Load FrmOrderMaintAqarSearch
            FrmOrderMaintAqarSearch.show

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
        
        SendMessage (2)
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


    MySQL = " SELECT     dbo.TblOrderMaintenanceDet.ORderID, dbo.TblOrderMaintenanceDet.TypeUnit, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblOrderMaintenanceDet.UnitNo, "
   MySQL = MySQL & "                   dbo.TblAqarDetai.unitno AS UnitNoName, dbo.TblOrderMaintenanceDet.UnitStatus, dbo.TblRentStatus.name AS NameStatus,"
   MySQL = MySQL & "                   dbo.TblRentStatus.namee AS NameStatusE, dbo.TblOrderMaintenanceDet.RenterID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
  MySQL = MySQL & "                    dbo.TblOrderMaintenanceDet.Mobile, dbo.TblOrderMaintenanceDet.Ms, dbo.TblOrderMaintenance.ID, dbo.TblOrderMaintenance.RecDateH,"
  MySQL = MySQL & "                    dbo.TblOrderMaintenance.RecDate, dbo.TblOrderMaintenance.TimOrder, dbo.TblOrderMaintenance.BranchID, dbo.TblBranchesData.branch_name,"
  MySQL = MySQL & "                    dbo.TblBranchesData.branch_namee, dbo.TblOrderMaintenance.EmpID, TblEmployee_2.Emp_Code, TblEmployee_2.Emp_Name, TblEmployee_2.Emp_Name1,"
 MySQL = MySQL & "                     TblEmployee_2.Emp_Name2, TblEmployee_2.Emp_Name3, TblEmployee_2.Emp_Name4, TblEmployee_2.Fullcode, TblEmployee_2.Emp_Namee4,"
 MySQL = MySQL & "                     TblEmployee_2.Emp_Namee3, TblEmployee_2.Emp_Namee2, TblEmployee_2.Emp_Namee1, TblEmployee_2.Emp_Namee, dbo.TblOrderMaintenance.SuperVM,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_Code AS Emp_CodeSup, TblEmployee_1.Emp_Name AS Emp_NameSup, TblEmployee_1.Emp_Name1 AS Emp_Name1Sup,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_Name2 AS Emp_Name2Sup, TblEmployee_1.Emp_Name3 AS Emp_Name3Sup, TblEmployee_1.Emp_Name4 AS Emp_Name4Sup,"
 MySQL = MySQL & "                     TblEmployee_1.Fullcode AS FullcodeSup, TblEmployee_1.Emp_Namee4 AS Emp_Namee4Sup, TblEmployee_1.Emp_Namee3 AS Emp_Namee3Sup,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_Namee2 AS Emp_Namee2Sup, TblEmployee_1.Emp_Namee1 AS Emp_Namee1Sup, TblEmployee_1.Emp_Namee AS Emp_NameeSup,"
MySQL = MySQL & "                      dbo.TblOrderMaintenance.AqrID, dbo.TblAqar.aqarname, dbo.TblOrderMaintenance.LocationIqar, dbo.TblOrderMaintenance.Des, dbo.TblOrderMaintenance.DMY,"
MySQL = MySQL & "                      dbo.TblOrderMaintenance.Cont, dbo.TblOrderMaintenance.EndFateH, dbo.TblOrderMaintenance.EndFate, dbo.TblOrderMaintenance.Lock,"
MySQL = MySQL & "                      dbo.TblOrderMaintenance.LockDateH , dbo.TblOrderMaintenance.LockDate"
MySQL = MySQL & " FROM         dbo.TblAqarDetai RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblOrderMaintenance ON TblEmployee_2.Emp_ID = dbo.TblOrderMaintenance.EmpID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee TblEmployee_1 ON dbo.TblOrderMaintenance.SuperVM = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblAqar ON dbo.TblOrderMaintenance.AqrID = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblOrderMaintenance.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblOrderMaintenanceDet ON dbo.TblOrderMaintenance.ID = dbo.TblOrderMaintenanceDet.ORderID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers ON dbo.TblOrderMaintenanceDet.RenterID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblRentStatus ON dbo.TblOrderMaintenanceDet.UnitStatus = dbo.TblRentStatus.id ON dbo.TblAqarDetai.Id = dbo.TblOrderMaintenanceDet.UnitNo LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAkarUnit ON dbo.TblOrderMaintenanceDet.TypeUnit = dbo.TblAkarUnit.id"
MySQL = MySQL & " Where (dbo.TblOrderMaintenance.id = " & val(XPTxtID.Text) & ")"



        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderMaintAqar.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderMaintAqar.rpt"
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




Private Sub CMDSENDSMS_Click()
'0 manual
'1 save
'2 Print

SendMessage (0)
End Sub
Function SendMessage(currentOpt As Integer)
            Dim subject As String
            Dim Msg As String
            Dim msgstatus As Boolean
           Dim CompanyName As String
           Dim cOptions As ClsCompanyInfo
           Set cOptions = New ClsCompanyInfo
           Dim companyphone As String
           Dim Opt As Integer
            Dim CurrentMessage As String
            Dim t As String
    CurrentMessage = ComposMessage(Me.Name, 0, "", "", Opt)
  If Opt = currentOpt Then
  
  
  
      CompanyName = cOptions.ArabCompanyName '& CHR(13) & CurrentBranchName
     companyphone = cOptions.Company_Mobile
  '
  Dim i As Integer
       For i = Me.UnitsGrid.FixedRows To UnitsGrid.Rows - 1
       
                 If (UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("unitno"))) <> "" Then
                
 '
 '··„‘—›
DoEvents
 Msg = "   „ «‰‘«¡ ÿ·» ’Ì«‰… ··ÊÕœ…  " & (UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("unitno"))) & Chr(13) & "  »«·⁄ﬁ«—   " & DcbIqara.Text & "."
t = sendMessageM("user", "password", Msg, "", GetEmployeeNumber(val(DcboEmpNameSuper.BoundText)))
 DoEvents
 
'  «·„” √Ã—

 Msg = "   „ «‰‘«¡ ÿ·» ’Ì«‰… ··ÊÕœ…  " & (UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("unitno"))) & Chr(13) & "  »«·⁄ﬁ«—   " & DcbIqara.Text & Chr(13) & "  ··„·«ÕŸ«  " & " 0552641000 "
t = sendMessageM("user", "password", Msg, "", UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("mobile")))




DoEvents
' ··„«·ﬂ
 Msg = "   „ «‰‘«¡ ÿ·» ’Ì«‰… ··ÊÕœ…  " & (UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("unitno"))) & Chr(13) & "  »«·⁄ﬁ«—   " & DcbIqara.Text
t = sendMessageM("user", "password", Msg, "", GetEmployeeNumber(getownerId(val(DcbIqara.BoundText))))
 
DoEvents


               
               
                End If
  
        Next i
  
MsgBox " „ «·«—”«·"
     
     
     End If
 
End Function



Private Sub DcbDMY_Change()
DcbDMY_Click
End Sub

Private Sub DcbDMY_Click()
If val(Me.DcbDMY.ListIndex) = 0 Then
EndDate.value = DateAdd("d", val(Me.TxtCount.Text), Date)
ElseIf val(Me.DcbDMY.ListIndex) = 1 Then
EndDate.value = DateAdd("M", val(Me.TxtCount.Text), Date)
ElseIf val(Me.DcbDMY.ListIndex) = 2 Then
EndDate.value = DateAdd("Yyyy", val(Me.TxtCount.Text), Date)
End If
EndDateH.value = ToHijriDate(EndDate.value)
End Sub

Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode
    retlocatin val(DcbIqara.BoundText), str
    txtLocation.Text = str
   ' dcsupplier.BoundText = ownerid
    'DcbUnitType_Change
End Sub




Private Sub DcbIqara_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmAqarSearch
FrmAqarSearch.m_RetrunType = 2020
FrmAqarSearch.show


End If
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
End Sub

Private Sub DcboEmpNameSuper_Change()
DcboEmpNameSuper_Click (0)
End Sub

Private Sub DcboEmpNameSuper_Click(Area As Integer)

   If val(DcboEmpNameSuper.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpNameSuper.BoundText, EmpCode
    TxtSearchCodeSuper.Text = EmpCode


Dim Mobile As String

        get_employee_information val(Me.DcboEmpNameSuper.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , , , Mobile
  txtmobile.Text = Mobile
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









Private Sub FilterDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         FilterDateH.value = ToHijriDate(FilterDate.value)
        
End If
End Sub

Private Sub FilterDateH_LostFocus()
  If Me.TxtModFlg.Text <> "R" Then
             VBA.Calendar = vbCalGreg
           FilterDate.value = ToGregorianDate(FilterDateH.value)
           End If
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
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

   
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmployees Me.DcboEmpNameSuper
    Dcombos.GetIqar DcbIqara
   ' Dcombos.getAkarUnit Me.DcbUnitType
  '  Dcombos.GetIqarUnit -2, 1, DcbUnitNo
    'Dcombos.GetCustomersSuppliers 1, Me.dcCustomer
     My_SQL = "select UserID,UserName From tblUsers "
  
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
   fill_combo DCboUserName, My_SQL
    
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblOrderMaintenance     Order By ID"
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
    XPDtbTrans_Change
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




Private Sub txtCount_Change()
DcbDMY_Click
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








Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub

Private Sub TxtSearchCodeSuper_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCodeSuper.Text, EmpID
        DcboEmpNameSuper.BoundText = EmpID
    End If
End Sub

Private Sub UnitsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With UnitsGrid
               
    

        Select Case .ColKey(Col)
         Case "customeridname"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("customerid"), False, True)
                .TextMatrix(Row, .ColIndex("customerid")) = StrAccountCode
                
                 StrSQL = "select * from TblCustemers  Where ( CusID= " & val(StrAccountCode) & " )"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rs.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("mobile")) = IIf(IsNull(rs("Cus_mobile").value), "", rs("Cus_mobile").value)
                End If
 Case "nameunittype"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("unittype"), False, True)
                .TextMatrix(Row, .ColIndex("unittype")) = StrAccountCode


Case "unitno"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
         StrSQL = " SELECT     dbo.TblAqarDetai.Status, dbo.TblRentStatus.name, dbo.TblRentStatus.namee, dbo.TblAqarDetai.Id, dbo.TblAqarDetai.Aqarid, dbo.TblAqarDetai.customerid, "
StrSQL = StrSQL & "                      dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.fullcode"
StrSQL = StrSQL & " FROM         dbo.TblAqarDetai LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblAqarDetai.customerid = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblRentStatus ON dbo.TblAqarDetai.Status = dbo.TblRentStatus.id"
StrSQL = StrSQL & " Where (dbo.TblAqarDetai.id =" & val(StrAccountCode) & ")"
 rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If rs.RecordCount > 0 Then
  If SystemOptions.UserInterface = ArabicInterface Then
  .TextMatrix(Row, .ColIndex("customeridname")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
  Else
  .TextMatrix(Row, .ColIndex("customeridname")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
  End If
  .TextMatrix(Row, .ColIndex("mobile")) = IIf(IsNull(rs("Cus_mobile").value), "", rs("Cus_mobile").value)
  .TextMatrix(Row, .ColIndex("customerid")) = IIf(IsNull(rs("customerid").value), "", rs("customerid").value)
                .TextMatrix(Row, .ColIndex("Status")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                .TextMatrix(Row, .ColIndex("StatusId")) = IIf(IsNull(rs("Status").value), "", rs("Status").value)
                End If
                
          '    Case "Status"
 'StrAccountCode = .ComboData
 '               LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("StatusId"), False, True)
 '
'                .TextMatrix(Row, .ColIndex("StatusId")) = StrAccountCode
'
    
           If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

  End Select
End With
ReLineGrid
End Sub

Private Sub UnitsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With UnitsGrid

      
        Select Case .ColKey(Col)
      
       Case "mobile"
             .ComboList = ""
              Case "customeridname"
             .ComboList = ""
           
               Case "Status"
             .ComboList = ""
   
        End Select

    End With
End Sub

Private Sub UnitsGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With UnitsGrid

        Select Case .ColKey(Col)
                   Case "customeridname"
                StrSQL = "select * from TblCustemers  Where ( Type=56  or CustomerandVendor=1 )"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                     StrComboList = UnitsGrid.BuildComboList(rs, "CusName", "CusID")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "CusNamee", "CusID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
 
        '    Case "Status"
        '        StrSQL = "select * from TblRentStatus"
        '        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    StrComboList = UnitsGrid.BuildComboList(rs, "name", "id")
'                Else
'                    StrComboList = UnitsGrid.BuildComboList(rs, "namee", "id")
'                End If
'
'                If StrComboList <> "" Then
'                    StrComboList = "|" & StrComboList
'                End If
'                 .ComboList = StrComboList
                 
            Case "nameunittype"
             .TextMatrix(Row, .ColIndex("unitno")) = ""
                StrSQL = "select * from TblAkarUnit"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = UnitsGrid.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
   
             Case "unitno"
            If val(DcbIqara.BoundText) = 0 Then
            MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄ﬁ«—"
            DcbIqara.SetFocus
            Exit Sub
            End If
                StrSQL = "select * from dbo.TblAqarDetai  where  Aqarid=" & val(DcbIqara.BoundText) & " and unittype=" & val(.TextMatrix(Row, .ColIndex("unittype")))
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = UnitsGrid.BuildComboList(rs, "unitno", "id")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "unitno", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 


 Case "namerentType"
                StrSQL = "select * from TblRentType"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = UnitsGrid.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
   
    


          
   


        End Select

    End With
    ReLineGrid
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
    Dim i As Integer
    Dim IntCounter As Integer
   ''''///
     With UnitsGrid

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("nameunittype")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
             
                  End If
                  
            

        Next i

    End With
    End Sub
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
  Dim ContactTime As Date
UnitsGrid.Clear flexClearScrollable, flexClearEverything
            UnitsGrid.Rows = 2
            UnitsGrid.Enabled = True
         
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
        If Not IsNull(rs("TimOrder").value) Then
      ContactTime = FormatDateTime(rs("TimOrder").value, vbShortTime)
        Me.TimeOrder.value = ContactTime
        End If
     Dcbranch.BoundText = val(IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value))
  Me.DcboEmpName.BoundText = val(IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value))
    
    DcbIqara.BoundText = val(IIf(IsNull(rs("AqrID").value), "", rs("AqrID").value))
     Me.DcboEmpNameSuper.BoundText = val(IIf(IsNull(rs("SuperVM").value), "", rs("SuperVM").value))
 Me.txtLocation.Text = IIf(IsNull(rs("LocationIqar").value), "", rs("LocationIqar").value)
   Me.TxtDes.Text = IIf(IsNull(rs("Des").value), "", rs("Des").value)
     Me.txtmobile.Text = IIf(IsNull(rs("Mobile").value), "", rs("Mobile").value)
  Me.DcbDMY.ListIndex = val(IIf(IsNull(rs("DMY").value), -1, rs("DMY").value))
  Me.TxtCount.Text = val(IIf(IsNull(rs("Cont").value), 0, rs("Cont").value))
     EndDate.value = IIf(IsNull(rs("EndFate").value), Date, rs("EndFate").value)
    EndDateH.value = IIf(IsNull(rs("EndFateH").value), "", rs("EndFateH").value)
    FilterDate.value = IIf(IsNull(rs("LockDate").value), Date, rs("LockDate").value)
    FilterDateH.value = IIf(IsNull(rs("LockDateH").value), "", rs("LockDateH").value)
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
   
   
    Set RsDetails = New ADODB.Recordset
StrSQL = "SELECT     dbo.TblOrderMaintenanceDet.ORderID, dbo.TblOrderMaintenanceDet.TypeUnit, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblOrderMaintenanceDet.UnitNo, "
 StrSQL = StrSQL & "                     dbo.TblAqarDetai.unitno AS UnitNoName, dbo.TblOrderMaintenanceDet.UnitStatus, dbo.TblRentStatus.name AS NameStatus,"
StrSQL = StrSQL & "                      dbo.TblRentStatus.namee AS NameStatusE, dbo.TblOrderMaintenanceDet.RenterID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenanceDet.mobile , dbo.TblOrderMaintenanceDet.Ms"
StrSQL = StrSQL & " FROM         dbo.TblOrderMaintenanceDet LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblOrderMaintenanceDet.RenterID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblRentStatus ON dbo.TblOrderMaintenanceDet.UnitStatus = dbo.TblRentStatus.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblOrderMaintenanceDet.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblOrderMaintenanceDet.TypeUnit = dbo.TblAkarUnit.id"
StrSQL = StrSQL & " Where (dbo.TblOrderMaintenanceDet.ORderID = " & val(XPTxtID.Text) & ")"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.UnitsGrid
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To .Rows - 1
   'IIf(IsNull(RsDetails("NameStatus").value), "", RsDetails("NameStatus").value) = 1
            .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("id")) = val(IIf(IsNull(RsDetails("UnitNo").value), 0, RsDetails("UnitNo").value))
            .TextMatrix(i, .ColIndex("unittype")) = val(IIf(IsNull(RsDetails("TypeUnit").value), 0, RsDetails("TypeUnit").value))
              .TextMatrix(i, .ColIndex("mobile")) = IIf(IsNull(RsDetails("Mobile").value), "", RsDetails("Mobile").value)
              If RsDetails("Ms").value = True Then
               .TextMatrix(i, .ColIndex("Ms")) = -1
              Else
              .TextMatrix(i, .ColIndex("Ms")) = 0
              End If
               .TextMatrix(i, .ColIndex("StatusId")) = val(IIf(IsNull(RsDetails("UnitStatus").value), 0, RsDetails("UnitStatus").value))
               .TextMatrix(i, .ColIndex("customerid")) = val(IIf(IsNull(RsDetails("RenterID").value), 0, RsDetails("RenterID").value))
                .TextMatrix(i, .ColIndex("unitno")) = IIf(IsNull(RsDetails("UnitNoName").value), "", RsDetails("UnitNoName").value)
                If SystemOptions.UserInterface = EnglishInterface Then
                                            
                .TextMatrix(i, .ColIndex("nameunittype")) = IIf(IsNull(RsDetails("namee").value), "", RsDetails("namee").value)
                .TextMatrix(i, .ColIndex("customeridname")) = IIf(IsNull(RsDetails("CusNamee").value), "", RsDetails("CusNamee").value)
               .TextMatrix(i, .ColIndex("Status")) = IIf(IsNull(RsDetails("NameStatusE").value), "", RsDetails("NameStatusE").value)
               Else
                
                .TextMatrix(i, .ColIndex("nameunittype")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
                .TextMatrix(i, .ColIndex("customeridname")) = IIf(IsNull(RsDetails("CusName").value), "", RsDetails("CusName").value)
               .TextMatrix(i, .ColIndex("Status")) = IIf(IsNull(RsDetails("NameStatus").value), "", RsDetails("NameStatus").value)
               End If
            RsDetails.MoveNext
         
        Next i
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

            XPTxtID.Text = CStr(new_id("TblOrderMaintenance", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
            StrSQL = "Delete From TblOrderMaintenanceDet Where ORderID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblUnitNoInformation Where OrderMaint=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If
     
   

   
       rs("ID").value = val(XPTxtID.Text)
        rs("RecDate").value = XPDtbTrans.value
       rs("RecDateH").value = Me.NourHijriCal1.value
       rs("TimOrder").value = FormatDateTime(Me.TimeOrder.value, vbShortTime)
       rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
       rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)
       rs("EmpID").value = IIf(Me.DcboEmpName.BoundText = "", Null, Me.DcboEmpName.BoundText)
       rs("AqrID").value = IIf(Me.DcbIqara.BoundText = "", Null, Me.DcbIqara.BoundText)
          rs("SuperVM").value = IIf(Me.DcboEmpNameSuper.BoundText = "", Null, Me.DcboEmpNameSuper.BoundText)
         rs("LocationIqar").value = Me.txtLocation.Text
         rs("Des").value = Me.TxtDes.Text
          rs("DMY").value = IIf(Me.DcbDMY.ListIndex = -1, Null, Me.DcbDMY.ListIndex)
     rs("Cont").value = val(Me.TxtCount.Text)
      rs("EndFate").value = EndDate.value
       rs("EndFateH").value = Me.EndDateH.value
       rs("EndFateH").value = Me.EndDateH.value
       rs("Mobile").value = Me.txtmobile.Text
       rs("LockDateH").value = Me.FilterDateH.value
   If ChLock.value = vbChecked Then
  rs("Lock").value = -1
   Else
   rs("Lock").value = 0
    End If
        rs.update
        '''''''''/////////////////////////////////
        Dim Temp As Integer
        Temp = -1
      Set RsDetails = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblOrderMaintenanceDet Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

          
       For i = Me.UnitsGrid.FixedRows To UnitsGrid.Rows - 1
   If (UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("unitno"))) <> "" Then
    RsDetails.AddNew
                  RsDetails("ORderID").value = val(XPTxtID.Text)
                 RsDetails("TypeUnit").value = val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("unittype")))
                RsDetails("UnitNo").value = val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("id")))
                RsDetails("RenterID").value = val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("customerid")))
                RsDetails("UnitStatus").value = val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("StatusId")))
                RsDetails("Mobile").value = UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("mobile"))
                If UnitsGrid.Cell(flexcpChecked, i, Me.UnitsGrid.ColIndex("Ms")) = flexChecked Then
        RsDetails("Ms").value = -1
        Else
        RsDetails("Ms").value = 0
           End If
         
        
         RsDetails.update
      
      '   End If
     
SaveUoitInformation val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("id"))), val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("StatusId"))), val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("customerid")))
      End If
        Next i
        
        '''''''''''''''//////////////////////////
        
 
      

    
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
Sub retlocatin(Optional ID As Integer, Optional ByRef str As String)
   Dim rs As ADODB.Recordset
   Dim str1 As String
    Dim StrSQL As String
    str1 = ""
    Set rs = New ADODB.Recordset
StrSQL = "  SELECT     dbo.TblAqar.cityid, dbo.TblAqar.Aqarid, dbo.TblAqar.heyid, dbo.TblAqar.schemeid, dbo.tblSchemes.name, dbo.tblSchemes.namee,"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments.GovernmentName , dbo.TblCountriesGovernmentsCities.CityName"
StrSQL = StrSQL & " FROM         dbo.tblSchemes RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernmentsCities RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblCountriesGovernmentsCities.CityID = dbo.TblAqar.heyid ON dbo.TblCountriesGovernments.GovernmentID = dbo.TblAqar.cityid ON"
StrSQL = StrSQL & "                      dbo.tblSchemes.id = dbo.TblAqar.schemeid"
StrSQL = StrSQL & " Where (dbo.TblAqar.Aqarid= " & ID & ")"
 rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If rs.RecordCount > 0 Then
 str1 = "„œÌ‰…:   "
              str1 = str1 + IIf(IsNull(rs("GovernmentName").value), "·«ÌÊÃœ", rs("GovernmentName").value) + " "
              str1 = str1 + "ÕÌ :   "
                str1 = str1 + IIf(IsNull(rs("CityName").value), "·«ÌÊÃœ", rs("CityName").value) + " "
                  str1 = str1 + "„Œÿÿ : "
                str1 = str1 + IIf(IsNull(rs("name").value), "·«ÌÊÃœ", rs("name").value) + " "
                str = str1
                End If
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
                StrSQL = "Delete From TblOrderMaintenance Where ID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
 StrSQL1 = "Delete From TblOrderMaintenanceDet Where ORderID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL1, , adExecuteNoRecords
             StrSQL = "Delete From TblUnitNoInformation Where OrderMaint=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                If rs.RecordCount < 1 Then
             
            UnitsGrid.Clear flexClearScrollable, flexClearEverything
            UnitsGrid.Rows = 2
            UnitsGrid.Enabled = True
            
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

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
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
        .Create Me.hWnd, "ÿ·» ’Ì«‰Â    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» ’Ì«‰Â", 1, 15204351, -2147483630
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

