VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmWaiverSettlement3 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18720
   FillColor       =   &H00C0E0FF&
   Icon            =   "FrmWaiverSettlement3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   18720
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
      Height          =   375
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   78
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
      Height          =   375
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   9240
      Width           =   1095
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
      Left            =   13350
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   1380
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox TxtDayPrice 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   12390
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   1050
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox TxtOrder 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   720
      Width           =   1515
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   16110
      RightToLeft     =   -1  'True
      TabIndex        =   50
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
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   9210
      Width           =   1695
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
      Width           =   18615
      _cx             =   32835
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
      Caption         =   " ’›Ì… Ê ‰«“· ⁄‰ «·⁄ﬁœ  "
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
      Begin VB.CheckBox chkoutCondition 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "‘—Êÿ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   120
         Width           =   1215
      End
      Begin VB.CheckBox chkoutflow 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Â—Ê»"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox TxtContNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
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
         ButtonImage     =   "FrmWaiverSettlement3.frx":038A
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
         ButtonImage     =   "FrmWaiverSettlement3.frx":0724
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
         ButtonImage     =   "FrmWaiverSettlement3.frx":0ABE
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
         ButtonImage     =   "FrmWaiverSettlement3.frx":0E58
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
         Left            =   4200
         Picture         =   "FrmWaiverSettlement3.frx":11F2
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
      Left            =   10230
      TabIndex        =   6
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   196280321
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   1110
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9660
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
      Left            =   7980
      TabIndex        =   15
      Top             =   9240
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
      Bindings        =   "FrmWaiverSettlement3.frx":4E5A
      Height          =   315
      Left            =   2640
      TabIndex        =   29
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
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
      Height          =   7815
      Left            =   0
      TabIndex        =   36
      Top             =   1200
      Width           =   18660
      _cx             =   32914
      _cy             =   13785
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
      Caption         =   "»Ì«‰« |New Tab|„’«—Ì› «Œ—Ï"
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
      Picture(0)      =   "FrmWaiverSettlement3.frx":4E6F
      Flags(1)        =   2
      Begin VB.Frame LblWork 
         BackColor       =   &H00E2E9E9&
         Height          =   7350
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   130
         Top             =   45
         Width           =   18570
         Begin VSFlex8Ctl.VSFlexGrid fg 
            Height          =   1740
            Left            =   120
            TabIndex        =   131
            Top             =   240
            Width           =   17985
            _cx             =   31724
            _cy             =   3069
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
            FormatString    =   $"FrmWaiverSettlement3.frx":5209
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
         Begin ImpulseButton.ISButton Cmdd 
            Height          =   375
            Left            =   360
            TabIndex        =   132
            Top             =   2040
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "«Õ”»"
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Ã„Ê⁄"
            Height          =   285
            Index           =   24
            Left            =   8280
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   2160
            Width           =   2970
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Height          =   285
            Index           =   12
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   133
            Top             =   2160
            Width           =   9570
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7350
         Left            =   -19215
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   45
         Width           =   18570
         _cx             =   32755
         _cy             =   12965
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
         Begin VB.Frame lblDataCli 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·„” «Ã—"
            Height          =   3540
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   0
            Width           =   11775
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   960
               Width           =   2355
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   159
               Top             =   600
               Width           =   2355
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   960
               Width           =   2595
            End
            Begin VB.TextBox TxtAmountDely 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   157
               Top             =   1680
               Width           =   2355
            End
            Begin VB.TextBox TxtDayLate 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   156
               Top             =   1680
               Width           =   2595
            End
            Begin VB.TextBox Text2 
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
               Left            =   3600
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   240
               Width           =   825
            End
            Begin VB.TextBox TxtDayPricen 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   154
               Top             =   2040
               Width           =   2955
            End
            Begin VB.TextBox TxtWaterPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   2040
               Width           =   2595
            End
            Begin VB.TextBox TxtActualDays 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   1320
               Width           =   2355
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   1320
               Width           =   2595
            End
            Begin VB.TextBox TxtDayPricentotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   2400
               Width           =   2955
            End
            Begin VB.TextBox TxtWaterPriceotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Top             =   2400
               Width           =   2595
            End
            Begin VB.TextBox TxtRentValuePayed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   2760
               Width           =   2955
            End
            Begin VB.TextBox txtWaterPayed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   147
               Top             =   2760
               Width           =   2595
            End
            Begin VB.TextBox TxtService 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   146
               Top             =   2400
               Width           =   2355
            End
            Begin VB.TextBox txtTelandNetPayed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   145
               Top             =   2760
               Width           =   2355
            End
            Begin VB.TextBox txtServicePrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   144
               Top             =   2040
               Width           =   2355
            End
            Begin VB.TextBox txtRemainService 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   143
               Top             =   3120
               Width           =   2355
            End
            Begin VB.TextBox txtRemainWater 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   3120
               Width           =   2595
            End
            Begin VB.TextBox txtRemainRent 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   3120
               Width           =   2955
            End
            Begin MSComCtl2.DTPicker EndDate 
               Height          =   315
               Left            =   8880
               TabIndex        =   161
               Top             =   1320
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   196280321
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   8880
               TabIndex        =   162
               Top             =   1680
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Format          =   196280321
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal EndDateH 
               Height          =   315
               Left            =   7380
               TabIndex        =   163
               Top             =   1320
               Width           =   1455
               _ExtentX        =   2778
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal NourHijriCal2 
               Height          =   315
               Left            =   7380
               TabIndex        =   164
               Top             =   1680
               Width           =   1455
               _ExtentX        =   2778
               _ExtentY        =   556
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   315
               Left            =   120
               TabIndex        =   165
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
               Top             =   240
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo3 
               Height          =   315
               Left            =   5070
               TabIndex        =   166
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
               Top             =   240
               Width           =   5235
               _ExtentX        =   9234
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo4 
               Height          =   315
               Left            =   3600
               TabIndex        =   167
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   600
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo5 
               Height          =   315
               Left            =   7320
               TabIndex        =   168
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   600
               Width           =   2985
               _ExtentX        =   5265
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker StartDate 
               Height          =   315
               Left            =   8880
               TabIndex        =   169
               Top             =   960
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   196280321
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal StartDateh 
               Height          =   315
               Left            =   7380
               TabIndex        =   170
               Top             =   960
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄ﬁ«—"
               Height          =   255
               Index           =   51
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   197
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ÊÕœÂ"
               Height          =   195
               Index           =   50
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   196
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«· √„Ì‰"
               Height          =   255
               Index           =   49
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   195
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "‰Â«Ì… «·«ÌÃ«—"
               Height          =   375
               Index           =   48
               Left            =   10320
               RightToLeft     =   -1  'True
               TabIndex        =   194
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «· ’›Ì…"
               Height          =   375
               Index           =   47
               Left            =   10320
               RightToLeft     =   -1  'True
               TabIndex        =   193
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·Õ”«»"
               Height          =   255
               Index           =   46
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   192
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ›Ê« Ì— ﬂÂ—»«¡"
               Height          =   375
               Index           =   45
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   191
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„»·€ «· «ŒÌ—"
               Height          =   255
               Index           =   22
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   190
               Top             =   360
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ «Ì«„ «·Œ’„"
               Height          =   255
               Index           =   21
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   189
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «·„” √Ã—"
               Height          =   285
               Index           =   1
               Left            =   10800
               RightToLeft     =   -1  'True
               TabIndex        =   188
               Top             =   240
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ê⁄ «·ÊÕœ…"
               Height          =   195
               Index           =   0
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Top             =   600
               Width           =   990
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»œ«Ì… «·«ÌÃ«—"
               Height          =   375
               Index           =   26
               Left            =   10260
               RightToLeft     =   -1  'True
               TabIndex        =   186
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·«ÌÃ«— «·ÌÊ„Ì"
               Height          =   255
               Index           =   28
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   185
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„Ì«Â «·ÌÊ„Ì"
               Height          =   255
               Index           =   29
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   184
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «Ì«„ «·”ﬂ‰"
               Height          =   255
               Index           =   31
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   183
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄ﬁœ ·„œ…"
               Height          =   375
               Index           =   17
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   182
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «·„ÿ·Ê» «ÌÃ«—"
               Height          =   375
               Index           =   33
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   181
               Top             =   2400
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «·„ÿ·Ê» „Ì«…"
               Height          =   375
               Index           =   34
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   180
               Top             =   2400
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„”œœ «ÌÃ«—"
               Height          =   375
               Index           =   35
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   179
               Top             =   2760
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " „”œœ „Ì«…"
               Height          =   375
               Index           =   36
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Top             =   2760
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «·„ÿ·Ê» Œœ„« "
               Height          =   375
               Index           =   37
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   2400
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„”œœ Œœ„« "
               Height          =   375
               Index           =   38
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   2760
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·Œœ„«  «·ÌÊ„Ì"
               Height          =   255
               Index           =   39
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   175
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„ »ﬁÌ Œœ„« "
               Height          =   375
               Index           =   40
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   174
               Top             =   3120
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " „ »ﬁÌ „Ì«…"
               Height          =   375
               Index           =   41
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   173
               Top             =   3120
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„ »ﬁÌ «ÌÃ«—"
               Height          =   375
               Index           =   42
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   172
               Top             =   3120
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„»·€ «·Œ’„"
               Height          =   255
               Index           =   43
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   1680
               Width           =   975
            End
         End
         Begin VB.TextBox TxtInsurance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   330
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   705
            Width           =   2355
         End
         Begin VB.TextBox TxtBillPrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   3810
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   705
            Width           =   2595
         End
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   1230
            Left            =   13440
            TabIndex        =   38
            Tag             =   "1"
            Top             =   600
            Width           =   11430
            _cx             =   20161
            _cy             =   2170
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
            FormatString    =   $"FrmWaiverSettlement3.frx":5344
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
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   6465
            Index           =   0
            Left            =   420
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   750
            Width           =   11445
            _cx             =   20188
            _cy             =   11404
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
            Begin VB.ComboBox DcbPeriodsID 
               Height          =   315
               ItemData        =   "FrmWaiverSettlement3.frx":5490
               Left            =   7575
               List            =   "FrmWaiverSettlement3.frx":549D
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   915
               Width           =   1140
            End
            Begin VB.TextBox TxtPeriods 
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
               Height          =   285
               Left            =   8820
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   960
               Width           =   1110
            End
            Begin VB.TextBox TxtPaymentCount 
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
               Height          =   285
               Left            =   8820
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   255
               Width           =   1110
            End
            Begin VB.CheckBox chkDivWater 
               Alignment       =   1  'Right Justify
               Caption         =   " ﬁ”Ì„ «·„Ì«Â ⁄·Ï «·œ›⁄« "
               ForeColor       =   &H00FF0000&
               Height          =   585
               Left            =   2490
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   135
               Width           =   2115
            End
            Begin VB.CheckBox chkDivElectric 
               Alignment       =   1  'Right Justify
               Caption         =   " ﬁ”Ì„ «·ﬂÂ—»«¡ ⁄·Ï «·œ›⁄« "
               ForeColor       =   &H00FF0000&
               Height          =   585
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   135
               Width           =   2385
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               Caption         =   "ÌœÊÌ"
               Height          =   180
               Index           =   2
               Left            =   1620
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   1125
               Width           =   1125
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               Caption         =   "«Œ— ﬁ”ÿ"
               Height          =   180
               Index           =   3
               Left            =   3105
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   1125
               Width           =   1140
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               Caption         =   "√Ê· ﬁ”ÿ"
               Height          =   180
               Index           =   4
               Left            =   4470
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   1125
               Width           =   1140
            End
            Begin MSComCtl2.DTPicker FristPaymentDate 
               Height          =   345
               Left            =   4710
               TabIndex        =   88
               TabStop         =   0   'False
               Top             =   255
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               Format          =   196280323
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal FirstInstallDateH 
               Height          =   285
               Left            =   6210
               TabIndex        =   89
               Top             =   255
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   503
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   465
               Index           =   20
               Left            =   495
               TabIndex        =   90
               Top             =   840
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   820
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈÷«›…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmWaiverSettlement3.frx":54B0
               DrawFocusRectangle=   0   'False
            End
            Begin C1SizerLibCtl.C1Tab TabMain 
               Height          =   5115
               Left            =   60
               TabIndex        =   91
               Top             =   1305
               Width           =   11370
               _cx             =   20055
               _cy             =   9022
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
               FrontTabForeColor=   16711680
               Caption         =   "«·œ›⁄«  |«·œ›⁄«  ﬁ»· «· ⁄œÌ·| Ê«—ÌŒ «· ⁄œÌ·«  ⁄·Ï «·œ›⁄« "
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
               Begin C1SizerLibCtl.C1Elastic ELe 
                  Height          =   4740
                  Index           =   12
                  Left            =   45
                  TabIndex        =   92
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   11280
                  _cx             =   19897
                  _cy             =   8361
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
                  Begin VB.CommandButton cmdSavePayment 
                     Caption         =   "Õ›Ÿ  ⁄œÌ·«  «·œ›⁄« "
                     Height          =   600
                     Left            =   8145
                     RightToLeft     =   -1  'True
                     TabIndex        =   93
                     Top             =   3630
                     Width           =   2055
                  End
                  Begin VSFlex8UCtl.VSFlexGrid FgItems 
                     Height          =   4740
                     Index           =   1
                     Left            =   13095
                     TabIndex        =   94
                     Top             =   1500
                     Width           =   11190
                     _cx             =   19738
                     _cy             =   8361
                     Appearance      =   2
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
                     Rows            =   50
                     Cols            =   5
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmWaiverSettlement3.frx":584A
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
                  Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
                     Height          =   3315
                     Left            =   -225
                     TabIndex        =   95
                     Top             =   -75
                     Width           =   11235
                     _cx             =   19817
                     _cy             =   5847
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
                     Rows            =   12
                     Cols            =   62
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmWaiverSettlement3.frx":590A
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
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "€Ì— „”œœ"
                     Height          =   330
                     Index           =   36
                     Left            =   1350
                     RightToLeft     =   -1  'True
                     TabIndex        =   99
                     Top             =   3915
                     Width           =   1455
                  End
                  Begin VB.Label LblNotPayed 
                     Alignment       =   2  'Center
                     Caption         =   "0"
                     Height          =   255
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   98
                     Top             =   3885
                     Width           =   1635
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·œ›⁄« "
                     Height          =   750
                     Index           =   34
                     Left            =   5940
                     RightToLeft     =   -1  'True
                     TabIndex        =   97
                     Top             =   3915
                     Width           =   1980
                  End
                  Begin VB.Label LblTotalQasts 
                     Alignment       =   2  'Center
                     Caption         =   "0"
                     Height          =   465
                     Left            =   4860
                     RightToLeft     =   -1  'True
                     TabIndex        =   96
                     Top             =   3765
                     Width           =   1650
                  End
               End
               Begin C1SizerLibCtl.C1Elastic ELe 
                  Height          =   4740
                  Index           =   11
                  Left            =   12015
                  TabIndex        =   100
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   11280
                  _cx             =   19897
                  _cy             =   8361
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
                  Begin VSFlex8UCtl.VSFlexGrid GridInstallments2 
                     Height          =   3540
                     Left            =   0
                     TabIndex        =   101
                     Top             =   0
                     Width           =   11235
                     _cx             =   19817
                     _cy             =   6244
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
                     Rows            =   12
                     Cols            =   61
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmWaiverSettlement3.frx":62A3
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
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "€Ì— „”œœ"
                     Height          =   1080
                     Index           =   71
                     Left            =   1350
                     RightToLeft     =   -1  'True
                     TabIndex        =   105
                     Top             =   3540
                     Width           =   1455
                  End
                  Begin VB.Label LblNotPayed2 
                     Alignment       =   2  'Center
                     Caption         =   "0"
                     Height          =   990
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   104
                     Top             =   3540
                     Width           =   1635
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«Ã„«·Ì «·œ›⁄« "
                     Height          =   1080
                     Index           =   72
                     Left            =   5955
                     RightToLeft     =   -1  'True
                     TabIndex        =   103
                     Top             =   3540
                     Width           =   1950
                  End
                  Begin VB.Label LblTotalQasts2 
                     Alignment       =   2  'Center
                     Caption         =   "0"
                     Height          =   990
                     Left            =   4860
                     RightToLeft     =   -1  'True
                     TabIndex        =   102
                     Top             =   3540
                     Width           =   1650
                  End
               End
               Begin C1SizerLibCtl.C1Elastic ELe 
                  Height          =   4740
                  Index           =   13
                  Left            =   12315
                  TabIndex        =   106
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   11280
                  _cx             =   19897
                  _cy             =   8361
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
                  Begin VSFlex8UCtl.VSFlexGrid grdHistory 
                     Height          =   14565
                     Left            =   5565
                     TabIndex        =   107
                     Top             =   210
                     Width           =   5715
                     _cx             =   10081
                     _cy             =   25691
                     Appearance      =   2
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
                     Rows            =   50
                     Cols            =   4
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmWaiverSettlement3.frx":6C0E
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
               End
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·› —Â »Ì‰ «·œ›⁄« "
               Height          =   180
               Index           =   11
               Left            =   9810
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   915
               Width           =   1470
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «Ê· œ›⁄Â"
               Height          =   645
               Index           =   9
               Left            =   7575
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   255
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ «·œ›⁄« "
               Height          =   645
               Index           =   8
               Left            =   10185
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   255
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Index           =   44
               Left            =   5340
               TabIndex        =   108
               Top             =   1125
               Width           =   2010
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«· √„Ì‰"
            Height          =   255
            Index           =   16
            Left            =   2610
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   705
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ›Ê« Ì— ﬂÂ—»«¡"
            Height          =   375
            Index           =   18
            Left            =   6450
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   4080
            Visible         =   0   'False
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
            Visible         =   0   'False
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   7350
         Index           =   15
         Left            =   -19515
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   45
         Width           =   18570
         _cx             =   32755
         _cy             =   12965
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
         _GridInfo       =   $"FrmWaiverSettlement3.frx":6CAD
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   7320
            Index           =   16
            Left            =   15
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   15
            Width           =   18540
            _cx             =   32703
            _cy             =   12912
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
            Begin VB.Frame Frame3 
               Caption         =   "«·ﬂÂ—»«¡"
               Height          =   975
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   213
               Top             =   4680
               Width           =   15195
               Begin VB.TextBox txtTotalCounter 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00808080&
                  Height          =   345
                  Left            =   2460
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   221
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtServiceCounter 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   3990
                  RightToLeft     =   -1  'True
                  TabIndex        =   220
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtPrevBalance 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   5550
                  RightToLeft     =   -1  'True
                  TabIndex        =   219
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtR 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00808080&
                  Height          =   345
                  Left            =   7140
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   218
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   8670
                  RightToLeft     =   -1  'True
                  TabIndex        =   217
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtDiff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00808080&
                  Height          =   345
                  Left            =   10290
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   216
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtLastInvoiceRead2 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   12060
                  RightToLeft     =   -1  'True
                  TabIndex        =   215
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtLastInvoiceRead 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   13890
                  RightToLeft     =   -1  'True
                  TabIndex        =   214
                  Top             =   540
                  Width           =   945
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   375
                  Index           =   70
                  Left            =   2460
                  RightToLeft     =   -1  'True
                  TabIndex        =   229
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œœ„… «·⁄œ«œ"
                  Height          =   375
                  Index           =   66
                  Left            =   3990
                  RightToLeft     =   -1  'True
                  TabIndex        =   228
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—’Ìœ ”«»ﬁ"
                  Height          =   375
                  Index           =   65
                  Left            =   5250
                  RightToLeft     =   -1  'True
                  TabIndex        =   227
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "R"
                  Height          =   285
                  Index           =   64
                  Left            =   7410
                  RightToLeft     =   -1  'True
                  TabIndex        =   226
                  Top             =   240
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄— «·ÊÕœ…"
                  Height          =   375
                  Index           =   63
                  Left            =   8700
                  RightToLeft     =   -1  'True
                  TabIndex        =   225
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·›—ﬁ"
                  Height          =   255
                  Index           =   61
                  Left            =   10500
                  RightToLeft     =   -1  'True
                  TabIndex        =   224
                  Top             =   240
                  Width           =   555
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄œ«œ ›Ì «Œ— ›« Ê—Â"
                  Height          =   405
                  Index           =   59
                  Left            =   13290
                  RightToLeft     =   -1  'True
                  TabIndex        =   223
                  Top             =   240
                  Width           =   1785
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄œ«œ ⁄‰œ Œ—ÊÃ «·„” √Ã—"
                  Height          =   555
                  Index           =   60
                  Left            =   11490
                  RightToLeft     =   -1  'True
                  TabIndex        =   222
                  Top             =   240
                  Width           =   1935
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "«Ì«„ ‰«ﬁ’…"
               Height          =   1245
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   206
               Top             =   6030
               Width           =   2895
               Begin VB.TextBox txtDaysValueIncomplete 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000010&
                  Height          =   345
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   211
                  Top             =   870
                  Width           =   1380
               End
               Begin VB.TextBox txtDayCountIncomplete 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000010&
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   90
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   210
                  Top             =   510
                  Width           =   1380
               End
               Begin VB.TextBox txtDayValueIncomplete 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000010&
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   90
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   208
                  Top             =   150
                  Width           =   1380
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·«Ì«„ «·‰«ﬁ’…"
                  Height          =   255
                  Index           =   56
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   212
                  Top             =   900
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «Ì«„"
                  Height          =   255
                  Index           =   58
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   209
                  Top             =   570
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·ÌÊ„"
                  Height          =   255
                  Index           =   53
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   207
                  Top             =   240
                  Width           =   915
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "«Ì«„ “Ì«œ…"
               Height          =   1185
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   199
               Top             =   4860
               Width           =   2925
               Begin VB.TextBox txtDaysValueIncrease 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000010&
                  Height          =   345
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   204
                  Top             =   780
                  Width           =   1380
               End
               Begin VB.TextBox txtDayCountInc 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000010&
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   90
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   203
                  Top             =   450
                  Width           =   1380
               End
               Begin VB.TextBox txtDayValueInc 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000010&
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   90
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   201
                  Top             =   120
                  Width           =   1380
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·«Ì«„ “Ì«œ…"
                  Height          =   255
                  Index           =   54
                  Left            =   1350
                  RightToLeft     =   -1  'True
                  TabIndex        =   205
                  Top             =   900
                  Width           =   1425
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «Ì«„"
                  Height          =   255
                  Index           =   52
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   202
                  Top             =   540
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·ÌÊ„"
                  Height          =   255
                  Index           =   57
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   200
                  Top             =   180
                  Width           =   915
               End
            End
            Begin VB.TextBox txtTotal2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   345
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   4590
               Visible         =   0   'False
               Width           =   1890
            End
            Begin VB.TextBox txtTotal1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   345
               Left            =   3780
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   4140
               Visible         =   0   'False
               Width           =   1380
            End
            Begin VB.TextBox TxtAccountNo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   2370
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   450
               Visible         =   0   'False
               Width           =   2280
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
               Left            =   7455
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   90
               Width           =   855
            End
            Begin VB.TextBox TxtContractDays 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   5625
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   810
               Visible         =   0   'False
               Width           =   2685
            End
            Begin VB.TextBox TxtNet 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   375
               Left            =   13020
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   6990
               Width           =   1950
            End
            Begin VB.TextBox TxtOFRenter 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Enabled         =   0   'False
               Height          =   360
               Left            =   13020
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   6570
               Width           =   1950
            End
            Begin VB.TextBox TxtForRenter 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Enabled         =   0   'False
               Height          =   360
               Left            =   13020
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   6150
               Width           =   1950
            End
            Begin VSFlex8Ctl.VSFlexGrid grd 
               Height          =   3300
               Left            =   90
               TabIndex        =   112
               Top             =   1200
               Width           =   18375
               _cx             =   32411
               _cy             =   5821
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
               GridLines       =   3
               GridLinesFixed  =   2
               GridLineWidth   =   5
               Rows            =   2
               Cols            =   24
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmWaiverSettlement3.frx":6CE3
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
            Begin MSDataListLib.DataCombo DcbIqara 
               Height          =   315
               Left            =   4200
               TabIndex        =   120
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
               Top             =   90
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcCustomer 
               Height          =   315
               Left            =   11145
               TabIndex        =   121
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
               Top             =   90
               Width           =   5220
               _ExtentX        =   9208
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbUnitNo 
               Height          =   315
               Left            =   5625
               TabIndex        =   122
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   450
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbUnitType 
               Height          =   315
               Left            =   13455
               TabIndex        =   123
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   450
               Width           =   2910
               _ExtentX        =   5133
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker FilterDate 
               Height          =   315
               Left            =   14985
               TabIndex        =   135
               Top             =   780
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               Format          =   196280321
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal FilterDateH 
               Height          =   315
               Left            =   13575
               TabIndex        =   136
               Top             =   780
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «· ’›Ì…"
               Height          =   375
               Index           =   20
               Left            =   16530
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ÊÕœÂ"
               Height          =   195
               Index           =   15
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   510
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄ﬁ«—"
               Height          =   255
               Index           =   13
               Left            =   8460
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   90
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·Õ”«»"
               Height          =   255
               Index           =   19
               Left            =   4650
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   450
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «·„” √Ã—"
               Height          =   285
               Index           =   5
               Left            =   16875
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   90
               Width           =   870
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ê⁄ «·ÊÕœ…"
               Height          =   195
               Index           =   15
               Left            =   16635
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   450
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄ﬁœ ·„œ…"
               Height          =   375
               Index           =   32
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   810
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂ «»…"
               ForeColor       =   &H8000000D&
               Height          =   315
               Index           =   11
               Left            =   11070
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   6990
               Width           =   1905
            End
            Begin VB.Label lbll 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   300
               Index           =   0
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   6990
               Width           =   8160
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "’«›Ì «·Õ”«»"
               Height          =   300
               Index           =   9
               Left            =   16425
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   6990
               Width           =   2040
            End
            Begin VB.Label lbll 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   285
               Index           =   11
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   6570
               Width           =   8160
            End
            Begin VB.Label lbll 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   300
               Index           =   9
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   6150
               Width           =   8160
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂ «»…"
               ForeColor       =   &H8000000D&
               Height          =   300
               Index           =   5
               Left            =   11040
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   6570
               Width           =   1935
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂ «»…"
               ForeColor       =   &H8000000D&
               Height          =   300
               Index           =   3
               Left            =   11100
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   6150
               Width           =   1875
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„»·€ «·„” Õﬁ  ··„” √Ã— »⁄œ «· ’›ÌÂ —ﬁ„«"
               Height          =   300
               Index           =   2
               Left            =   14550
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   6570
               Width           =   3915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„»·€ «·„” Õﬁ  ⁄·Ï «·„” √Ã— »⁄œ «· ’›ÌÂ —ﬁ„«"
               Height          =   300
               Index           =   10
               Left            =   14085
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   6150
               Width           =   4380
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„‹‹‹‹‹‹‹‹‹·«ÕŸ‹‹‹‹‹‹‹‹‹«  «· ’‹‹›‹‹Ì‹‹‹‹‹…"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Index           =   23
               Left            =   27630
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   2985
               Width           =   7035
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   4200
               Index           =   62
               Left            =   3660
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   2010
               Width           =   945
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   7320
            Index           =   9
            Left            =   15
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   15
            Width           =   18540
            _cx             =   32703
            _cy             =   12912
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
               Height          =   5490
               Left            =   4845
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   1590
               Width           =   1035
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   3720
               Left            =   6150
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   2010
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3720
               Index           =   67
               Left            =   3435
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   2010
               Width           =   945
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   3660
               Index           =   68
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   2505
               Width           =   45
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
               Height          =   4350
               Index           =   69
               Left            =   4380
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   2010
               Width           =   465
            End
         End
      End
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   0
      TabIndex        =   52
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
      Left            =   8610
      TabIndex        =   54
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
   End
   Begin MSComCtl2.DTPicker From 
      Height          =   315
      Left            =   12360
      TabIndex        =   70
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   196280321
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   375
      Left            =   0
      TabIndex        =   72
      TabStop         =   0   'False
      ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
      Top             =   720
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
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
      BackStyle       =   0
      ButtonImage     =   "FrmWaiverSettlement3.frx":70B3
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«Ì«„ “Ì«œ…"
      Height          =   255
      Index           =   55
      Left            =   1830
      RightToLeft     =   -1  'True
      TabIndex        =   198
      Top             =   6840
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ ««·ﬁÌœ"
      Height          =   255
      Index           =   25
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   9240
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "⁄ﬁœ —ﬁ„"
      Height          =   255
      Index           =   14
      Left            =   1740
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄„«—Â"
      Height          =   255
      Left            =   3660
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   0
      Width           =   855
   End
   Begin VB.Image img 
      Height          =   855
      Left            =   22680
      Picture         =   "FrmWaiverSettlement3.frx":74B0
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
      Picture         =   "FrmWaiverSettlement3.frx":84D4
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   720
   End
   Begin VB.Label lblBr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Left            =   4800
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
      Caption         =   "—ﬁ„ «· ’›ÌÂ"
      Height          =   285
      Index           =   4
      Left            =   17430
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
      Left            =   11250
      TabIndex        =   23
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   10725
      TabIndex        =   22
      Top             =   9315
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2310
      TabIndex        =   21
      Top             =   9270
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   570
      TabIndex        =   20
      Top             =   9270
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   -30
      TabIndex        =   19
      Top             =   9300
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1620
      TabIndex        =   18
      Top             =   9300
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
Attribute VB_Name = "FrmWaiverSettlement3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim UonitStatus As Integer
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String
Dim FlagContrNew As Boolean
Dim FlagContrNew2 As Boolean


Private Sub cmdSavePayment_Click()
  Dim Msg As String
'    If mchkAllowEditPaymentCont Then
'        TxtModFlg = "E"
'    End If
    
    If (ChkRenew Or checkContractTransactions(val(TxtContNo.Text))) And mchkAllowEditPaymentCont Then
        mCanEdit = True
        
    Else
        mCanEdit = False
    End If
    
    If ChkRenew.value = vbChecked And Not mchkAllowEditPaymentCont Then
        MsgBox "·« Ì„ﬂ‰  ⁄œÌ· Â–« «·⁄ﬁœ ·«‰… „Ãœœ "
        Exit Sub
    End If


    If checkContractTransactions(val(TxtContNo.Text)) = True And Not mchkAllowEditPaymentCont Then
        MsgBox "ÌÊÃœ Õ—ﬂ«  „ﬁ»Ê÷«  ⁄·Ï Â–« «·⁄ﬁœ Ê·«Ì„ﬂ‰  ⁄œÌ·…", vbCritical
        Exit Sub
    
    End If
    
            Dim s As String
        Dim RsDetails2 As New ADODB.Recordset
        s = "Select * from TblContractInstallmentsHist Where ContNo = " & Trim(TxtContNo.Text)
        
 
    
        RsDetails2.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        'If Not RsDetails2.EOF Then
            RsDetails2.AddNew
            RsDetails2!UserID = user_id
            RsDetails2!EditDate = Date
            RsDetails2!ContNo = val(TxtContNo)
            RsDetails2.update
        'End If
       
       SaveGridPayment False
       SaveGridPayment True
       MsgBox " „ Õ›Ÿ  ⁄œÌ·«  «·œ›⁄« "
       RetriveOldPayment
       
    
End Sub


Private Sub FirstInstallDateH_GotFocus()

hijriorJerojian = 0
End Sub
Private Sub FirstInstallDateH_LostFocus()
        
        If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            FristPaymentDate.value = ToGregorianDate(FirstInstallDateH.value)
               
        End If

End Sub

Private Sub FristPaymentDate_Change()
If Me.TxtModFlg.Text <> "R" Then
     
         FirstInstallDateH.value = ToHijriDate(FristPaymentDate.value)
       
End If
End Sub

Private Sub FristPaymentDate_GotFocus()
hijriorJerojian = 1
End Sub

Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With Me.GridInstallments
Select Case .ColKey(Col)
Case "NetWater"
.TextMatrix(Row, .ColIndex("Water")) = .TextMatrix(Row, .ColIndex("NetWater"))
Case "NetElectric"
.TextMatrix(Row, .ColIndex("Electric")) = .TextMatrix(Row, .ColIndex("NetElectric"))
End Select
End With

ReLineGrid
End Sub
Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

'   If (ChkRenew Or checkContractTransactions(val(TxtContNo.Text))) And mchkAllowEditPaymentCont Then
'        mCanEdit = True
'
'    Else
'        mCanEdit = False
'    End If
'
'    If ChkRenew.value = vbChecked And Not mchkAllowEditPaymentCont Then
'        MsgBox "·« Ì„ﬂ‰  ⁄œÌ· Â–« «·⁄ﬁœ ·«‰… „Ãœœ "
'        Exit Sub
'    End If
'
'    If checkContractTransactions(val(TxtContNo.Text)) = True And Not mchkAllowEditPaymentCont Then
'        MsgBox "ÌÊÃœ Õ—ﬂ«  „ﬁ»Ê÷«  ⁄·Ï Â–« «·⁄ﬁœ Ê·«Ì„ﬂ‰  ⁄œÌ·…", vbCritical
'        Exit Sub
'
'    End If

    If (Me.TxtModFlg.Text = "R" And GridInstallments.ColKey(Col) <> "PrintJE" And GridInstallments.ColKey(Col) <> "Print" And GridInstallments.ColKey(Col) <> "RecalcVAt") Then
        If Not mchkAllowEditPaymentCont Then
            Cancel = True
        End If
    Else

    
    
    With GridInstallments
 If (Opt(4).value = True Or Opt(3).value = True) And .ColKey(Col) <> "Print" Then
 Cancel = True
 ElseIf GridInstallments.ColKey(Col) = "RecalcVAt" And val(TxtFATValue.Text) <> 0 Then
 Cancel = True
 Else
 Cancel = False
 End If
 
    '     If .ColKey(Col) <> "Status" And .ColKey(Col) <> "TelandNet" And .ColKey(Col) <> "Insurance" Then
   
   '      If Opt(0).value = True Then Cancel = True: Exit Sub
   '     Cancel = True
   '
   '     End If
 
        
    End With
  End If
End Sub

Private Sub GridInstallments_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With GridInstallments
Select Case .ColKey(Col)
Case "Print"
PeintInstalMent val(.TextMatrix(Row, .ColIndex("InstallNo")))
Case "PrintJE"
ShowGL_cc .TextMatrix(Row, .ColIndex("NoteSerial")), , 200
Case "RecalcVAt"
RecalcVAt Row
createVoucher2 (Row)
End Select
End With
End Sub
Private Sub GridInstallments_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With GridInstallments
Select Case .ColKey(Col)
Case "Print"
.ColComboList(.ColIndex("Print")) = "..."

Case "RecalcVAt"
.ColComboList(.ColIndex("RecalcVAt")) = "..."
Case "PrintJE"
.ColComboList(.ColIndex("PrintJE")) = "..."

End Select
End With
End Sub
''//
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
  '«·„” √Ã—
 Msg = "  „ ⁄„·    ’›Ì…  " & "  ··ÊÕœ… —ﬁ„   " & DcbUnitNo.Text & "    ··⁄ﬁ«— —ﬁ„ " & CHR(13) & DcbIqara.Text & " √„·Ì‰ —÷«ﬂ„ "
t = sendMessageM("user", "password", Msg, "", GetCustomerNumber(dcCustomer.BoundText))

'„” √Ã—

DoEvents
 Msg = "  „ ⁄„·    ’›Ì…  " & "  ··ÊÕœ… —ﬁ„   " & DcbUnitNo.Text & CHR(13) & "    ··⁄ﬁ«— —ﬁ„ " & DcbIqara.Text
t = sendMessageM("user", "password", Msg, "", GetCustomerNumber(getownerId(DcbIqara.BoundText)))



DoEvents



MsgBox " „ «·«—”«·"
     
     
     End If
 
End Function


Sub GetUonitStatus()
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL As String
 
       Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT   Status  from  TblAqarDetai where id =" & val(DcbUnitNo.BoundText) & ""
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails1.RecordCount > 0 Then
 UonitStatus = val(IIf(IsNull(RsDetails1("Status").value), "", RsDetails1("Status").value))
   End If
   End Sub
  Sub SaveUoitInformation()
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL, Msg As String
Msg = ""

 
    If SystemOptions.UserInterface = EnglishInterface Then
      Msg = Msg & "Action filter and a waiver of the decade No."
             Msg = Msg & CHR(13) & XPTxtID.Text
               Msg = Msg & CHR(13)
         Msg = Msg & " Dtae  "
      Msg = Msg & NourHijriCal1.value & "corresponding to" & XPDtbTrans.value
      Msg = Msg & CHR(13)
        Msg = Msg & " The amount due from Renter  "
      Msg = Msg & TxtForRenter.Text
      Msg = Msg & CHR(13)
        Msg = Msg & "  The amount due to Renter "
      Msg = Msg & TxtOFRenter.Text
      Msg = Msg & CHR(13)
       Msg = Msg & "  Net "
      Msg = Msg & txtnet.Text
      Msg = Msg & CHR(13)
      Else
      Msg = Msg & "   „ ⁄„·  ’›ÌÂ Ê ‰«“· »—ﬁ„  "
      Msg = Msg & XPTxtID.Text
      Msg = Msg & CHR(13)
         Msg = Msg & " » «—ÌŒ  "
      Msg = Msg & NourHijriCal1.value & "«·„Ê«›ﬁ" & XPDtbTrans.value
      Msg = Msg & CHR(13)
        Msg = Msg & " «·„»·€ «·„” Õﬁ ⁄·Ï «·„” «Ã—  "
      Msg = Msg & TxtForRenter.Text
      Msg = Msg & CHR(13)
        Msg = Msg & "  «·„»·€ «·„” Õﬁ ··„” √Ã—  "
      Msg = Msg & TxtOFRenter.Text
      Msg = Msg & CHR(13)
       Msg = Msg & "  «·’«›Ì "
      Msg = Msg & txtnet.Text
      Msg = Msg & CHR(13)

End If
        Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TblUnitNoInformation Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      RsDetails1.AddNew
      RsDetails1("CusID").value = val(dcCustomer.BoundText)
      RsDetails1("BranchId").value = val(Dcbranch.BoundText)
      
           RsDetails1("UnitNo").value = val(DcbUnitNo.BoundText)
           RsDetails1("UnitStatus").value = UonitStatus
           RsDetails1("Des").value = Msg
           RsDetails1("RecDate").value = XPDtbTrans.value
           RsDetails1("RecDateH").value = NourHijriCal1.value
           RsDetails1("NoteID").value = Null
           RsDetails1("ContNo").value = Null
           RsDetails1("FilterNo").value = val(XPTxtID.Text)
           RsDetails1("OrderMaint").value = Null
           RsDetails1("LocOrderMaint").value = Null
           RsDetails1.update

   End Sub
   

''//''
Sub RetriveIqarCOmpenet()
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
    
Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
               
    Set RsDetails = New ADODB.Recordset

     
StrSQL = " SELECT     TOP 100 PERCENT dbo.TblAqrCompenetDet.IDAqComp, dbo.TblAqrCompenetDet.Name AS NameT, dbo.TblAqrCompenetDet.Price, dbo.TblAqrCompenetDet.ID, "
 StrSQL = StrSQL & "                     dbo.TblAqrCompenet.Namee, dbo.TblAqrCompenet.Name, dbo.TblAqrCompenetDet.Namee AS NameET"
StrSQL = StrSQL & ",TblAqrCompenetDet.Accountsus  FROM         dbo.TblAqrCompenet LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAqrCompenetDet ON dbo.TblAqrCompenet.ID = dbo.TblAqrCompenetDet.IDAqComp"
StrSQL = StrSQL & " ORDER BY dbo.TblAqrCompenetDet.IDAqComp"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
     
     Dim j, k As Integer
   Dim temp As Integer
j = 0
temp = -1
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.Fg
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails.RecordCount
k = 0
        For i = .FixedRows To .Rows - 1
    j = j + 1
    k = k + 1
    
    If temp = val(IIf(IsNull(RsDetails("IDAqComp").value), 0, RsDetails("IDAqComp").value)) Then
            .TextMatrix(k, .ColIndex("serial")) = j
           
            .TextMatrix(k, .ColIndex("iditem")) = val(IIf(IsNull(RsDetails("id").value), "", RsDetails("id").value))
               If SystemOptions.UserInterface = EnglishInterface Then
       .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("NameeT").value), "", RsDetails("NameT").value)
       Else
       .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("NameT").value), "", RsDetails("NameT").value)
    End If
           .TextMatrix(k, .ColIndex("price")) = val(IIf(IsNull(RsDetails("Price").value), "", RsDetails("Price").value))
            
            Else
           
            .Rows = .Rows + 1
            .TextMatrix(k, .ColIndex("iditem")) = 0
             .TextMatrix(k, .ColIndex("id")) = val(IIf(IsNull(RsDetails("IDAqComp").value), "", RsDetails("IDAqComp").value)) 'val(RsDetails("IDAqComp").value)
           .TextMatrix(k, .ColIndex("serial")) = ""
                  
                 If SystemOptions.UserInterface = EnglishInterface Then
       .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("Namee").value), "", RsDetails("Namee").value)
       Else
       .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("Name").value), "", RsDetails("Name").value)
    End If
             .TextMatrix(k, .ColIndex("price")) = ""
              .Cell(flexcpBackColor, k, 1, k, 7) = &H80C0FF
             
             k = k + 1
         '    j = j + 1
       '   .Cell(flexcpBackColor, k, 1, k, 7) = &H80C0FF
               .TextMatrix(k, .ColIndex("serial")) = j
              If SystemOptions.UserInterface = EnglishInterface Then
       .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("NameeT").value), "", RsDetails("NameT").value)
       Else
       .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("NameT").value), "", RsDetails("NameT").value)
    End If
             .TextMatrix(k, .ColIndex("price")) = val(IIf(IsNull(RsDetails("Price").value), "", RsDetails("Price").value))
              temp = val(IIf(IsNull(RsDetails("IDAqComp").value), 0, RsDetails("IDAqComp").value))
              .TextMatrix(k, .ColIndex("iditem")) = val(IIf(IsNull(RsDetails("id").value), "", RsDetails("id").value))
               .TextMatrix(k, .ColIndex("Accountsus")) = (IIf(IsNull(RsDetails("Accountsus").value), "", RsDetails("Accountsus").value))
              
               .TextMatrix(k, .ColIndex("id")) = 0
          '  j = 0
           End If
            RsDetails.MoveNext
         
        Next i
    ReLineGridCount
End With
    End If

    RsDetails.Close
    Set RsDetails = Nothing
End Sub
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
   lbl(12).Caption = 0
    IntCounter = 0

    With Fg

        For i = .FixedRows To .Rows - 1
 
               If Fg.TextMatrix(i, Fg.ColIndex("Accountsus")) <> "" Then
                                    '  If val(.TextMatrix(i, .ColIndex("count"))) <> 0 Then
                             .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("price")) * val(.TextMatrix(i, .ColIndex("count"))))
                               lbl(12).Caption = val(lbl(12).Caption) + val(.TextMatrix(i, .ColIndex("total")))
                        
                           ' End If
    
   End If
 
        Next i
        
        Dim totals As String
        'totals = val(txtRemainWater) + val(txtRemainRent) + val(txtRemainService)
        
     '   TxtForRenter.text = val(lbl(12).Caption) + val(TxtBillPrice)
 
 TxtForRenter.Text = 0
  TxtOFRenter.Text = 0
 TxtOFRenter.Text = val(Me.TxtInsurance.Text)
 
 TxtForRenter.Text = Round(val(TxtForRenter.Text) + val(TxtAmountDely) + val(lbl(12).Caption) + val(txtTotalCounter), 3)
 
 
' If totals > 0 Then
' TxtForRenter.Text = Round(val(TxtForRenter.Text) + val(totals), 3)
'
' Else
' TxtOFRenter = Round(val(TxtOFRenter) + val(Abs(totals)), 3)
' End If
 TxtForRenter = Round(val(txttotal1) + val(lbl(12).Caption) + val(txtTotalCounter), 3)
 TxtOFRenter = Round(val(txttotal2), 3)
 txtnet.Text = Round(val(TxtForRenter.Text) - val(TxtOFRenter.Text), 3)


 ReLineGridCount
    End With
    

'Me.Lbtotal.Caption = val(Me.LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
End Sub
Private Sub ReLineGridCount()
    Dim i As Integer
    Dim IntCounter  As Integer

    IntCounter = 0

    With Fg

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("serial")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("serial")) = IntCounter
      Else
      IntCounter = 0
      
        '.TextMatrix(i, .ColIndex("serial")) = IntCounter
            End If

        Next i
 
    End With
    

'Me.Lbtotal.Caption = val(Me.LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
End Sub
'Edit.Caption = "Sent To approval "
'End If

Public Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index
'ddddddddddddddddd
     Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
        XPDtbTrans.value = Date
    FilterDate.value = Date
    FilterDateH.value = ToHijriDate(Date)
    NourHijriCal1.value = ToHijriDate(Date)
 RetriveIqarCOmpenet
Dcbranch.BoundText = Current_branch
  Me.DCboUserName.BoundText = user_id
  
  ReLineGrid
  Grd.Rows = 1
 ' FG.Rows = 1
  
        Case 1
                   If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
 Fg.Rows = Fg.Rows + 1
            Fg.Enabled = True
            '
            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id




        Case 2
                       If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
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
                   If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
            Load FrmIqarWaiverSet
          ' Wael
            FrmIqarWaiverSet.Show vbModal



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
           Case 20
        If Txtorder <> "" Then
'RtriveInfoOrbon val(TxtNotID.Text)
End If
        If FlagContrNew2 = False Then
        If TxtNoteserial.Text <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
Else
MsgBox "Please Delete JE"
End If
Exit Sub
End If
End If
            If Me.TxtModFlg.Text <> "R" Then
                If Opt(4).value = False And Opt(3).value = False And Opt(2).value = False Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ «Œ Ì«— ÿ—ﬁ… Ã»— «·ﬂ”Ê—"
                    Else
                        MsgBox "Please Select Method Number of decimal"
                    End If
                    Exit Sub
                End If
'                If val(TxtTotalContract.Text) < val(TxtMiniRentValue.Text) Then
'                    MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ «·«Ã«— «ﬁ· „‰ «ﬁ· ﬁÌ„…  «ÃÌ—ÌÂ"
'                    TxtTotalContract.SetFocus
'                    Exit Sub
'                End If
'                If val(TxtPaymentCount) = 0 Then
'                    MsgBox "·«»œ „‰  ÕœÌœ «·› —… »Ì‰ «·œ›⁄« "
'                    TxtPaymentCount.SetFocus
'                    'SendKeys "{F4}"
'                     Exit Sub
'                End If
 Dim MSGType As Integer
                If CheckJE() = True Then
                 MSGType = MsgBox("”Ê›  „ Õ–› ﬁÌœ «·œ›⁄«  ", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
                 If MSGType = vbNo Then
                 Exit Sub
                 End If
                End If
                DeleteJE
                Calculations
            End If
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Calculations(Optional WithMsg As Boolean = True)
'    On Error GoTo ErrTrap
    Dim SngAllValue As Single
    Dim i  As Integer
    Dim IntNoOFQast As Integer
    Dim IntRes As Integer
    Dim SngOnePor As Single
    Dim FirstDate As Date
    Dim PreDate As Date
    Dim NewDate As Date
    Dim DateInterval As String
    Dim NewDateH As String
    Dim PreDateH As String
    Dim InstalNew As Double
    Dim DateNumber As Integer
    Dim Msg As String
    
Dim watervalue As Double
Dim Electricity As Double
    If TxtPaymentCount.Text = "" Then
   
            Msg = "ÌÃ» ≈œŒ«· ⁄œœ «·√ﬁ”«ÿ"

                        If WithMsg = True Then
                            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            TxtPaymentCount.SetFocus
                        End If

            Exit Sub
  End If
  
 If chkDivWater.value = vbChecked Then
 If val(TxtPaymentCount.Text) > 0 Then
 watervalue = Round(val(TxtWater.Text) \ val(TxtPaymentCount.Text), 2)
 Else
 watervalue = 0
 End If
 Else
 watervalue = val(TxtWater.Text)
 End If

 If chkDivElectric.value = vbChecked Then
  If val(TxtPaymentCount.Text) > 0 Then
 Electricity = Round(val(TxtElectricity.Text) \ val(TxtPaymentCount.Text), 2)
 Else
 Electricity = 0
 End If
Else
Electricity = val(TxtElectricity.Text)
 End If



    If DcbPeriodsID.ListIndex = -1 Then
   
            Msg = "ÌÃ» ≈œŒ«·   «·› —… »Ì‰ «·«ﬁ”«ÿ"

                        If WithMsg = True Then
                            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            DcbPeriodsID.SetFocus
                        End If

            Exit Sub
  End If
  
        If Not IsNumeric(TxtPaymentCount.Text) Then
            Msg = " ⁄œœ «·√ﬁ”«ÿ ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"

                    If WithMsg = True Then
                        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                         TxtPaymentCount.SetFocus
                    End If

            Exit Sub
        End If
    SngAllValue = Round((val(TxtTotalContract)) / val(TxtPaymentCount), 2)
    SngAllValue = SngAllValue + val(watervalue) + val(Electricity) + val(TxtEnternet)
    IntNoOFQast = val(TxtPaymentCount)
    SngOnePor = SngAllValue

   ' If val(Me.TxtPaymentCount.text) > 0 Then
   '  '   IntNoOFQast = SngAllValue \ val(Me.TxtPaymentCount.text)
   '  ' SngOnePor = val(Me.TxtPaymentCount.text)
   '     SngOnePor = SngAllValue / IntNoOFQast
   ' Else
   '     SngOnePor = SngAllValue / IntNoOFQast
   ' End If
'
    If DcbPeriodsID.ListIndex = 0 Then
        DateInterval = "d"
    ElseIf DcbPeriodsID.ListIndex = 1 Then
        DateInterval = "M"
    ElseIf DcbPeriodsID.ListIndex = 2 Then
        DateInterval = "yyyy"
    End If

    NewDate = FristPaymentDate.value
    NewDateH = FirstInstallDateH.value
     
     DateNumber = val(TxtPeriods.Text)

    'End If
    
   If FlagContrNew2 = True Then
  InstalNew = InstalNo + 1
   End If
    Dim notpayed As Double
    notpayed = 0
    With Me.GridInstallments
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + IntNoOFQast

        For i = 1 To IntNoOFQast

            DoEvents
            If FlagContrNew2 = False Then
            .TextMatrix(i, .ColIndex("InstallNo")) = i
            Else
            .TextMatrix(i, .ColIndex("InstallNo")) = InstalNew
            .TextMatrix(i, .ColIndex("TempInstal")) = i
             InstalNew = InstalNew + 1
            End If
              .TextMatrix(i, .ColIndex("Countsofall")) = val(TxtPeriods.Text)
           
            
            If i = 1 Then
           ''// ''19 08 2015 NetRent
           .TextMatrix(i, .ColIndex("Rent1")) = Round((val(TxtTotalContract.Text)) / val(TxtPaymentCount.Text), 2)
           .TextMatrix(i, .ColIndex("RentArbon")) = val(TxtRetValue2.Text)
           .TextMatrix(i, .ColIndex("VATArboon")) = val(TxtFATValue2.Text)
           .TextMatrix(i, .ColIndex("NetRent")) = val(.TextMatrix(i, .ColIndex("Rent1")))
           .TextMatrix(i, .ColIndex("VATValue")) = val(TxtFATValue.Text) / val(TxtPaymentCount.Text)
           .TextMatrix(i, .ColIndex("Commissions1")) = val(TxtCommiValue.Text)
           .TextMatrix(i, .ColIndex("CommissionsArbon")) = val(TxtCommValue2.Text)
           .TextMatrix(i, .ColIndex("NetCommissions")) = val(TxtCommiValue.Text) - val(TxtCommValue2.Text)
           .TextMatrix(i, .ColIndex("ServiceArbon")) = val(TxtServce.Text)
           
           .TextMatrix(i, .ColIndex("Insurance1")) = val(TxtInsuranceValue.Text)
           .TextMatrix(i, .ColIndex("InsuranceArbon")) = val(TxtInstrunceValue2.Text)
           .TextMatrix(i, .ColIndex("NetInsurance")) = val(TxtInsuranceValue.Text) - val(TxtInstrunceValue2.Text)
           
           If chkDivWater.value = vbChecked Then
             .TextMatrix(i, .ColIndex("Water1")) = Round((val(TxtWater.Text)) / IntNoOFQast, 2)
             .TextMatrix(i, .ColIndex("WaterArbon")) = val(TxtWaterValue2.Text)
           Else
           .TextMatrix(i, .ColIndex("WaterArbon")) = val(TxtWaterValue2.Text)
            .TextMatrix(i, .ColIndex("Water1")) = val(TxtWater.Text)
          End If
    .TextMatrix(i, .ColIndex("NetWater")) = val(.TextMatrix(i, .ColIndex("Water1"))) '- val(.TextMatrix(i, .ColIndex("WaterArbon")))
    .TextMatrix(i, .ColIndex("Water")) = val(.TextMatrix(i, .ColIndex("NetWater")))
      If chkDivElectric.value = vbChecked Then
             .TextMatrix(i, .ColIndex("Electric1")) = Round((val(TxtElectricity.Text)) / IntNoOFQast, 2)
             .TextMatrix(i, .ColIndex("ElectricArbon")) = val(TxtElectricityValue2.Text)
           Else
            .TextMatrix(i, .ColIndex("Electric1")) = val(TxtElectricity.Text)
             .TextMatrix(i, .ColIndex("ElectricArbon")) = val(TxtElectricityValue2.Text)
            
          End If
               .TextMatrix(i, .ColIndex("NetElectric")) = val(.TextMatrix(i, .ColIndex("Electric1"))) ' - val(.TextMatrix(i, .ColIndex("WaterArbon")))
           .TextMatrix(i, .ColIndex("Electric")) = val(.TextMatrix(i, .ColIndex("NetElectric")))
           ''//
            .TextMatrix(i, .ColIndex("TelandNet")) = val(TxtPhone)
             .TextMatrix(i, .ColIndex("RentValue")) = Round(((val(TxtTotalContract.Text)) / IntNoOFQast), 2)
              .TextMatrix(i, .ColIndex("Commissions")) = val(TxtCommiValue)
             .TextMatrix(i, .ColIndex("Insurance")) = val(TxtInsuranceValue)
             .TextMatrix(i, .ColIndex("Commissions")) = val(TxtCommiValue)
           .TextMatrix(i, .ColIndex("Value")) = Round(SngOnePor, Decimal_Places1) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtPhone.Text) + val(.TextMatrix(i, .ColIndex("VATValue")))
         
         
         
         
         .TextMatrix(i, .ColIndex("hijri")) = hijriorJerojian
If chkDivWater.value = vbChecked Then
    .TextMatrix(i, .ColIndex("Water")) = Round((val(TxtWater) / IntNoOFQast), 2)
 Else
    .TextMatrix(i, .ColIndex("Water")) = val(TxtWater)
 End If

 If chkDivElectric.value = vbChecked Then
 .TextMatrix(i, .ColIndex("Electric")) = Round((val(TxtElectricity) / IntNoOFQast), 2)
 Else
 .TextMatrix(i, .ColIndex("Electric")) = val(TxtElectricity)
 End If
           
      
         
            
            
            Else
    '        .TextMatrix(i, .ColIndex("Value")) = Round(SngOnePor, Decimal_Places1)
            
'            If chkDivWater.value = vbChecked Then
'    .TextMatrix(i, .ColIndex("Water")) = val(TxtWater) / IntNoOFQast
' Else
'    .TextMatrix(i, .ColIndex("Water")) = 0
' End If
       If chkDivWater.value = vbChecked Then
             .TextMatrix(i, .ColIndex("Water1")) = Round((val(TxtWater.Text)) / IntNoOFQast, 2)
             '.TextMatrix(i, .ColIndex("WaterArbon")) = val(TxtWaterValue2.text)
           Else
           .TextMatrix(i, .ColIndex("WaterArbon")) = 0
            .TextMatrix(i, .ColIndex("Water1")) = 0
          End If
          .TextMatrix(i, .ColIndex("VATValue")) = val(TxtFATValue.Text) / IntNoOFQast
    .TextMatrix(i, .ColIndex("NetWater")) = val(.TextMatrix(i, .ColIndex("Water1")))
    .TextMatrix(i, .ColIndex("Water")) = val(.TextMatrix(i, .ColIndex("NetWater")))
    
             .TextMatrix(i, .ColIndex("RentValue")) = Round((val(TxtTotalContract)) / IntNoOFQast, 2)
 If chkDivElectric.value = vbChecked Then
 .TextMatrix(i, .ColIndex("Electric")) = Round(val(TxtElectricity) / IntNoOFQast, 2)
 Else
 .TextMatrix(i, .ColIndex("Electric")) = 0
 End If
 
   If chkDivElectric.value = vbChecked Then
             .TextMatrix(i, .ColIndex("Electric1")) = Round((val(TxtElectricity.Text)) / IntNoOFQast, 2)
            ' .TextMatrix(i, .ColIndex("ElectricArbon")) = val(TxtElectricityValue2.text)
           Else
            .TextMatrix(i, .ColIndex("Electric1")) = 0
           '  .TextMatrix(i, .ColIndex("ElectricArbon")) = 0
            
          End If
               .TextMatrix(i, .ColIndex("NetElectric")) = val(.TextMatrix(i, .ColIndex("Electric1"))) ' - val(.TextMatrix(i, .ColIndex("WaterArbon")))
          
          .TextMatrix(i, .ColIndex("Electric")) = val(.TextMatrix(i, .ColIndex("NetElectric")))
            End If
            
          
            
            If i = 1 Then
                NewDate = NewDate
                NewDateH = NewDateH
            
            Else
                PreDate = CDate(Trim(.TextMatrix(i - 1, .ColIndex("Due_Date"))))
                
                If hijriorJerojian = 1 Then 'jorijan
                NewDate = DateAdd(DateInterval, DateNumber, PreDate)
                NewDateH = ToHijriDate(NewDate)
                End If
                
                     PreDateH = (Trim(.TextMatrix(i - 1, .ColIndex("Due_DateH"))))
     
If hijriorJerojian = 0 Then 'hijri
                NewDateH = (DateAdd(DateInterval, DateNumber, PreDateH))
NewDate = ToGregorianDate(NewDateH)
End If
                
                
                
            End If
   
   
   Dim currentvalue  As Double
    Dim increasrate  As Double
    
   .TextMatrix(i, .ColIndex("Value")) = val(.TextMatrix(i, .ColIndex("RentValue"))) + val(.TextMatrix(i, .ColIndex("VATValue"))) + val(.TextMatrix(i, .ColIndex("Commissions"))) + val(.TextMatrix(i, .ColIndex("Insurance"))) + val(.TextMatrix(i, .ColIndex("NetWater"))) + val(.TextMatrix(i, .ColIndex("NetElectric"))) + val(.TextMatrix(i, .ColIndex("TelandNet")))
  
   
   If lblnew.Visible = True Then
 ' currentvalue = .TextMatrix(i, .ColIndex("Value"))
 '  increasrate = currentvalue * val(TxtIncresYearRate) / 100
 '  currentvalue = currentvalue + increasrate
 '    .TextMatrix(i, .ColIndex("Value")) = currentvalue
   End If
   
  
   
            .TextMatrix(i, .ColIndex("Due_Date")) = Format(NewDate, "yyyy/MM/dd")
            .TextMatrix(i, .ColIndex("Due_DateH")) = Format(NewDateH, "yyyy/MM/dd")
                   If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then
           notpayed = notpayed + val(.TextMatrix(i, .ColIndex("Value")))
        End If
        
        
        
     
            
            Due_Date = Format(NewDate, "yyyy/M/d")
        Next i
LblNotPayed.Caption = notpayed
         .AutoSize 1, .Cols - 1, False
        Me.LblTotalQasts.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With
 ReLineGrid

 
    'BolQastCal = True
    Exit Sub
ErrTrap:
End Sub
Sub DeleteJE()
Dim i As Integer
Dim StrSQL As String
With GridInstallments
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("NoteId"))) <> 0 Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(.TextMatrix(i, .ColIndex("NoteId"))) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "delete From Notes where NoteID =" & val(.TextMatrix(i, .ColIndex("NoteId"))) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute "Update TblContractInstallments set NoteID=null ,NoteSerial=null where id=" & val(.TextMatrix(i, .ColIndex("Installid"))) & " "
FindRec val(Me.TxtContNo.Text)
End If
Next i
End With

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
        
    Dim s As String
    s = " SELECT DISTINCT"
    s = s & "    dbo.TblFiterWaiver.ID,"
    s = s & "        dbo.TblFiterWaiver.RecordDateH,"
    s = s & "        dbo.TblFiterWaiver.RecordDate,"
    s = s & "        dbo.TblFiterWaiver.BranchID,"
    s = s & "        dbo.TblFiterWaiver.BulidID,"
    s = s & "        dbo.TblAqar.aqarname,"
    s = s & "        dbo.TblFiterWaiver.RenterID,"
    s = s & "        dbo.TblCustemers.CusName,"
    s = s & "        dbo.TblCustemers.CusNamee,"
    s = s & "        dbo.TblFiterWaiver.ApartmentID,"
    s = s & "        dbo.TblAqarDetai.unitno,"
    s = s & "        dbo.TblFiterWaiver.EndDateH,"
    s = s & "        dbo.TblFiterWaiver.EndDate,"
    s = s & "        dbo.TblFiterWaiver.FilterDate,"
    s = s & "        dbo.TblFiterWaiver.FilterDateH,"
    s = s & "        t2.BillPrice,"
    s = s & "        dbo.TblFiterWaiver.AccountNo,"
    s = s & "        dbo.TblFiterWaiver.AmountDely,"
    s = s & "        dbo.TblFiterWaiver.DayNo,"
    s = s & "        dbo.TblFiterWaiver.UserID,"
    s = s & "        dbo.TblFiterWaiver.OFRenter,"
    s = s & "        dbo.TblFiterWaiver.ForRenter,"
    s = s & "        dbo.TblFiterWaiver.unittype,"
    s = s & "        dbo.TblAkarUnit.name         AS nameUnt,"
    s = s & "        dbo.TblAkarUnit.namee,"
    s = s & "        dbo.TblFiterWaiver.ContNo,"
    s = s & "        dbo.TblFiterWaiver.ContractNo,"
    s = s & "        dbo.TblFiterWaiver.NoteID,"
    s = s & "        dbo.TblFiterWaiver.NoteSerial,"
    s = s & "        dbo.TblFiterWaiver.ContractDays,"
    s = s & "        dbo.TblFiterWaiver.WaterPrice,"
    s = s & "        dbo.TblFiterWaiver.ActualDays,"
    s = s & "        dbo.TblFiterWaiver.DayPricen,"
    s = s & "        T2.WaterPriceotal,"
    s = s & "        T2.ServicePrice,"
    s = s & "        T2.DayPricentotal,"
    s = s & "        T2.Service,"
    s = s & "        T2.WaterPayed,"
    s = s & "        T2.RentValuePayed,"
    s = s & "        T2.OldRent TelandNetPayed,"
    s = s & "        T2.RemainWater,"
    s = s & "        T2.RemainRent,"
    s = s & "        T2.RemainService,"
    s = s & "        T2.Insurance,"
    s = s & "        T2.outflow,"
    s = s & "        T2.StartDate,"
    s = s & "        T2.StartDateh,"
    s = s & "        T2.TotalStill,"
    s = s & "        T2.RemainCommissions,"
    s = s & "        T2.NoDaye,"
    s = s & "        dbo.TblFiterWaiver.outCondition,"
    s = s & "        dbo.TblFiterWaiver.DaysValueIncrease,"
    s = s & "        dbo.TblFiterWaiver.DaysValueIncomplete,"
    s = s & "        dbo.TblFiterWaiver.DayValueInc,"
    s = s & "        dbo.TblFiterWaiver.DayCountInc,"
    s = s & "        dbo.TblFiterWaiver.DayValueIncomplete,"
    s = s & "        dbo.TblFiterWaiver.DayCountIncomplete,"
    s = s & "        dbo.TblFiterWaiver.Efflux,"
    s = s & "        dbo.TblFiterWaiver.ValDay,"
    s = s & "        dbo.TblFiterWaiver.Discount,"
    s = s & "        dbo.TblFiterWaiver.totalcollected,"
    s = s & "        dbo.TblFiterWaiver.totalpayed,"
    s = s & "        dbo.TblFiterWaiver.LegalIssue,"
    s = s & "        dbo.TblFiterWaiver.net"
    s = s & " From dbo.TblAkarUnit"
    s = s & "        RIGHT OUTER JOIN dbo.TblFiterWaiver"
    s = s & "             ON  dbo.TblAkarUnit.id = dbo.TblFiterWaiver.unittype"
    s = s & "        LEFT OUTER JOIN dbo.TblAqarDetai"
    s = s & "             ON  dbo.TblFiterWaiver.ApartmentID = dbo.TblAqarDetai.Id"
    s = s & "        LEFT OUTER JOIN dbo.TblCustemers"
    s = s & "             ON  dbo.TblFiterWaiver.RenterID = dbo.TblCustemers.CusID"
    s = s & "        LEFT OUTER JOIN dbo.TblAqar"
    s = s & "             ON  dbo.TblFiterWaiver.BulidID = dbo.TblAqar.Aqarid"

            
    s = s & "        LEFT OUTER JOIN dbo.TblFiterWaiverDet2 T2"
    s = s & "             ON  T2.MasterID = dbo.TblFiterWaiver.ID"
  

    s = s & "        Where (dbo.TblFiterWaiver.id = " & val(XPTxtID.Text) & ")"
    
    
  'db_createOrUpdateviewSQL "VwFiterWaiver", s

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepFilterWiaver.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepFilterWiaver.rpt"
        End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

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
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng

        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
       xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtForRenter.Text), "0.00"), 0, True, ".")
          xReport.ParameterFields(5).AddCurrentValue WriteNo(Format(val(TxtOFRenter.Text), "0.00"), 0, True, ".")
          xReport.ParameterFields(6).AddCurrentValue WriteNo(Format(val(txtnet.Text), "0.00"), 0, True, ".")
          xReport.ParameterFields(7).AddCurrentValue (lbl(12).Caption)
          xReport.ParameterFields(8).AddCurrentValue WriteNo(Format(val(lbl(12).Caption), "0.00"), 0, True, ".")
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

Private Sub Cmdd_Click()


'TxtForRenter.text = 0
'TxtOFRenter.text = 0
'TxtForRenter.text = val(TxtForRenter.text) + val(TxtBillPrice.text) + val(TxtAmountDely.text) + val(lbl(12).Caption)

'TxtNet.text = val(TxtOFRenter.text) - val(TxtForRenter.text)



End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub


Private Sub Command9_Click()
    ShowGL_cc Me.TxtNoteserial.Text, , 200
End Sub

Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then: Exit Sub

    Dim EmpCode  As String
 Dim ownerid As Integer
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode
   ' dcsupplier.BoundText = ownerid
    'DcbUnitType_Change
End Sub


Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteserial.Text = ""
    TxtNoteserial1.Text = ""
    
End Sub

Private Sub dcCustomer_Click(Area As Integer)
    dcCustomer_Change
End Sub


Private Sub dcCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmAqarSearch
FrmIqarContractSearch.m_RetrunType = 1
FrmIqarContractSearch.Show


End If
End Sub

Private Sub ENDDATE_Change()
If Me.TxtModFlg.Text <> "R" Then
         EndDateH.value = ToHijriDate(EndDate.value)
'         IntMintsCount = (DateDiff("d", EndDate, FilterDate))
'Me.TxtDayLate.text = IntMintsCount
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
Dim pricrday As Double
         FilterDateH.value = ToHijriDate(FilterDate.value)
         Dim IntMintsCount As Integer
         RetriveOrder
         GetContract val(Txtorder)


            If IntMintsCount > 0 Then
            pricrday = val(TxtDayPrice.Text) / IntMintsCount
           ' TxtAmountDely.text = pricrday * val(TxtDayLate.text)
            Else
            pricrday = 0
            ' TxtAmountDely.text = 0
            End If
dcCustomer_Change

End If

End Sub


Private Sub FilterDateH_LostFocus()
 If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            FilterDate.value = ToGregorianDate(FilterDateH.value)
                     Dim IntMintsCount, pricrday As Integer
IntMintsCount = (DateDiff("d", EndDate, FilterDate))
'Me.TxtDayLate.text = IntMintsCount
IntMintsCount = (DateDiff("d", From, EndDate))
If IntMintsCount <> 0 Then
pricrday = val(TxtDayPrice.Text) / IntMintsCount
End If
'TxtAmountDely.text = pricrday * val(TxtDayLate.text)
FilterDate_Change
            End If
End Sub


Function CALCdISCOUNT() As Double
CALCdISCOUNT = Round(val(TxtWaterPrice) * val(TxtDayLate), 2) + Round(val(TxtDayPricen) * val(TxtDayLate), 2) + Round(val(txtServicePrice) * val(TxtDayLate), 2)


End Function




 

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()
RetriveOrder
GetContract val(Txtorder)
'
'Load FrmIqarContractSearch
'FrmIqarContractSearch.fg.TextMatrix(fg.Row, fg.ColIndex("NoteSerial")) = Me.TxtOrder
'FrmIqarContractSearch.m_RetrunType = 2
'FrmIqarContractSearch.show vbModal

Load FrmIqarContractSearch
'FrmIqarContractSearch.fg.TextMatrix(fg.Row, fg.ColIndex("NoteSerial"))=me.Text15
FrmIqarContractSearch.m_RetrunType = 2
FrmIqarContractSearch.Show vbModal
End Sub

Private Sub lblDataCli_DragDrop(Source As Control, X As Single, Y As Single)
''''
End Sub

Private Sub NourHijriCal1_LostFocus()
      If Me.TxtModFlg.Text <> "R" Then
             VBA.Calendar = vbCalGreg
           XPDtbTrans.value = ToGregorianDate(NourHijriCal1.value)
           End If
End Sub


Public Sub RetriveOrder(Optional order_no As String = "", Optional serial As Integer)
   
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
  '  If serial = 1 Then
  StrSQL = "Select * from TblContract  where    ContNo='" & val(Me.TxtContNo) & "'"
  '  Else
 '   StrSQL = "Select * from TblContract  where    NoteSerial1='" & order_no & "'"
  '  End If
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
     DcbIqara.BoundText = IIf(IsNull(rs("Iqar").value), "", rs("Iqar").value)
     DcbUnitType.BoundText = IIf(IsNull(rs("unittype").value), "", rs("unittype").value)
     DcbUnitNo.BoundText = val(IIf(IsNull(rs("UnitNo").value), "", rs("UnitNo").value))
     
     dcCustomer.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
       Txtorder.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
         EndDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    EndDateH.value = IIf(IsNull(rs("TodateH").value), "", rs("TodateH").value)
    
          startDate.value = IIf(IsNull(rs("StrDate").value), Date, rs("StrDate").value)
   StartDateh.value = IIf(IsNull(rs("FromdateH").value), "", rs("FromdateH").value)
   Dim IntMintsCount   As Integer
IntMintsCount = (DateDiff("d", EndDate, FilterDate))
'Me.TxtDayLate.text = IntMintsCount
IntMintsCount = (DateDiff("d", From, EndDate))
    TxtActualDays.Text = (DateDiff("d", startDate.value, FilterDate.value))
TxtContractDays.Text = (DateDiff("d", (rs("StrDate").value), (rs("EndDate").value)))
'TxtContractDays.Text = val(val(TxtContractDays.Text) * 30)
    'datediff("m",date(FromdateH),date(TodateH))
   
    TxtAccountNo.Text = Me.Text15.Text
       ' TxtActualDays.Text = (DateDiff("d", startDate, FilterDate))
' TxtContractDays.text = (DateDiff("d", CDate(rs("Fromdateh").value), CDate(rs("todateH").value)))
'TxtContractDays.text = (DateDiff("d", startDate, EndDate))

    
    TxtDayLate = val(TxtContractDays.Text) - val(TxtActualDays.Text)


       From.value = IIf(IsNull(rs("StrDate").value), Date, rs("StrDate").value)
         
     Me.TxtDayPrice.Text = IIf(IsNull(rs("TotalContract").value), "", rs("TotalContract").value)
    TxtInsurance.Text = IIf(IsNull(rs("InsuranceValue").value), "", rs("InsuranceValue").value)
     TxtBillPrice.Text = 0 ' IIf(IsNull(rs("Electricity").value), "", rs("Electricity").value)
TxtService.Text = IIf(IsNull(rs("phone").value), "", rs("phone").value)
 Dim WaterPayed     As Double
Dim RentValuePayed As Double
Dim TelandNetPayed As Double

         If val(TxtContractDays.Text) > 0 Then

      Me.TxtDayPricen.Text = Round(IIf(IsNull(rs("TotalContract").value), 0, rs("TotalContract").value) / val(TxtContractDays.Text), 2)
        Me.txtServicePrice.Text = Round(IIf(IsNull(rs("phone").value), 0, rs("phone").value) / val(TxtContractDays.Text), 2)
     
     TxtWaterPrice.Text = Round(IIf(IsNull(rs("water").value), 0, rs("water").value) / val(TxtContractDays.Text), 2)
      TxtDayPricentotal = val(Me.TxtDayPricen.Text) * val(TxtActualDays.Text)
    TxtWaterPriceotal = val(Me.TxtWaterPrice.Text) * val(TxtActualDays.Text)
     TxtService = val(Me.txtServicePrice.Text) * val(TxtActualDays.Text)
     getActualpayedToContract val(TxtContNo), RentValuePayed, WaterPayed, TelandNetPayed
      TxtRentValuePayed.Text = RentValuePayed
      txtWaterPayed.Text = WaterPayed
       txtTelandNetPayed.Text = TelandNetPayed
       
     txtRemainRent.Text = val(TxtDayPricentotal) - val(TxtRentValuePayed)
       txtRemainWater.Text = val(TxtWaterPriceotal) - val(txtWaterPayed.Text)
       txtRemainService.Text = val(TxtService) - val(txtTelandNetPayed.Text)
       
      End If
      
      
         ' tod.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    End If
ReLineGrid
    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    
   
    Exit Sub
ErrTrap:

End Sub

Private Sub TxtActualDays_Change()
  TxtDayPricentotal = val(Me.TxtDayPricen.Text) * val(TxtActualDays.Text)
    TxtWaterPriceotal = val(Me.TxtWaterPrice.Text) * val(TxtActualDays.Text)
     TxtService = val(Me.txtServicePrice.Text) * val(TxtActualDays.Text)
End Sub

Private Sub TxtAmountDely_Change()
'TxtForRenter.text = val(TxtForRenter.text) + val(Me.TxtAmountDely.text)
End Sub

Private Sub TxtBillPrice_Change()
'TxtForRenter.text = val(TxtForRenter.text) + val(Me.TxtBillPrice.text)
ReLineGrid

End Sub

Private Sub TxtContNo_Change()
Dim ID As Long
'If Cek(val(TxtContNo.text), ID) = True Then
'Retrive ID
'Else
'RetriveOrder val(TxtContNo.text)
' ReLineGrid
'End If
End Sub
Public Function chek(Optional ContNo As Long = 0, Optional ByRef WaviStNo As Long) As Boolean
Dim Rs5 As ADODB.Recordset
Dim sql As String
Set Rs5 = New ADODB.Recordset
sql = "select id from TblFiterWaiver where ContNo=" & ContNo & " "
     If SystemOptions.usertype <> UserAdminAll Then
             sql = sql & "   and BranchID=" & Current_branch & " "
    End If
    Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
chek = True
WaviStNo = IIf(IsNull(Rs5("id").value), 0, Rs5("id").value)
Retrive WaviStNo
Else
Cmd_Click (0)
'RetriveOrder val(txtContNo.text)
chek = False
End If
End Function

Private Sub TxtDayLate_Change()
 TxtAmountDely.Text = CALCdISCOUNT
 ReLineGrid
 
End Sub

Private Sub TxtDayLate_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtDayLate.Text, 1)

 
End Sub





Private Sub Fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid
End Sub



Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With Fg

        Select Case .ColKey(Col)
            
            Case "total"
               Cancel = True
       
        End Select

    End With

    
End Sub

Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
des = " ’›Ì… ⁄ﬁœ «ÌÃ«— —ﬁ„ " & Txtorder & " · " & dcCustomer.Text
Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
tablename = "TblFiterWaiver"
Filedname = "ID"
ContNo = XPTxtID

Notevalue = 0


                     If Me.TxtModFlg = "N" Then
                                 CreateNotes NoteID, (XPDtbTrans.value), val(Dcbranch.BoundText), 60, Notevalue, NoteSerial, XPTxtID, tablename, Filedname, ContNo, des, NourHijriCal1.value
                                     TxtNoteID.Text = NoteID
                                    TxtNoteserial.Text = NoteSerial
                    Else
                                      If TxtNoteID.Text = "" Or TxtNoteserial.Text = "" Then
                                    CreateNotes NoteID, (XPDtbTrans.value), val(Dcbranch.BoundText), 60, Notevalue, NoteSerial, TxtNoteserial1, tablename, Filedname, ContNo, des, NourHijriCal1.value
                                                       TxtNoteID.Text = NoteID
                                                  TxtNoteserial.Text = NoteSerial
                                    Else
                                                  sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                  sql = sql & ",NoteSerial1='" & XPTxtID & "'"
                                                    sql = sql & " where NoteID=" & val(TxtNoteID.Text)
                                                     Cn.Execute sql
                                                     
                                       End If
                         
                    End If
ReLineGrid
CREATE_VOUCHER_GE val(TxtNoteID.Text), val(Dcbranch.BoundText), user_id, XPDtbTrans.value
rs.Resync adAffectCurrent


End Function



Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
 
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        

 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
     
    my_branch = BranchID

 
'        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
'GoTo ll
            
  
            StrTempDes = " ’›Ì… ⁄ﬁœ «ÌÃ«— —ﬁ„    " & TxtNoteserial1 & "  ··„” √Ã—   " & dcCustomer.Text & " ··ÊÕœ… " & DcbUnitNo.Text
            LngDevNO = LngDevNO + 1
'Notevalue = val(TxtTotalContract) + val(TxtPayAmini) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone) + val(TxtEnternet)
Notevalue = 0
 
 Dim Account_Code_dynamic80 As String
  Dim Account_Code_dynamic81 As String
   Dim Account_Code_dynamic82 As String
    Dim Account_Code_dynamic83 As String
     Dim Account_Code_dynamic84 As String
      Dim Account_Code_dynamic85 As String
      
   Account_Code_dynamic80 = get_account_code_branch(80, my_branch)
            Account_Code_dynamic81 = get_account_code_branch(81, my_branch)
            Account_Code_dynamic82 = get_account_code_branch(82, my_branch)
            Account_Code_dynamic83 = get_account_code_branch(83, my_branch)
            Account_Code_dynamic84 = get_account_code_branch(84, my_branch)
            Account_Code_dynamic85 = get_account_code_branch(85, my_branch)
            
'll:
   LngDevNO = 0
  
 
 If val(txtRemainRent.Text) < 0 Then
       '«·⁄„Ì· œ«∆‰
       Notevalue = Abs(val(txtRemainRent.Text))
   LngDevNO = LngDevNO + 1
   StrTempAccountCode = Account_Code_dynamic80
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… «·«ÌÃ«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
        LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «·«ÌÃ«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  

            
            
  End If
  
  '*************
   If val(txtRemainWater.Text) > 0 Then
       '«·⁄„Ì· œ«∆‰
       Notevalue = Abs(val(txtRemainWater.Text))
   LngDevNO = LngDevNO + 1
   StrTempAccountCode = Account_Code_dynamic83
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… «·„Ì«… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
        LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «·„Ì«… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  

            
            
  End If
  
   
   If val(txtRemainService.Text) < 0 Then
       '«·⁄„Ì· œ«∆‰
       Notevalue = Abs(val(txtRemainService.Text))
   LngDevNO = LngDevNO + 1
   StrTempAccountCode = Account_Code_dynamic85
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… «·Œœ„«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
        LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «·Œœ„«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  

            
            
  End If
   
  
   
   
'*************************************************
If val(txtRemainRent.Text) > 0 Then
       '«·⁄„Ì· „œÌ‰
       Notevalue = Abs(val(txtRemainRent.Text))
   LngDevNO = LngDevNO + 1
   
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
      
      If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… «·«ÌÃ«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  
  StrTempAccountCode = Account_Code_dynamic80
           
        LngDevNO = LngDevNO + 1
  
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «·«ÌÃ«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If

            
            
  End If
  
  '*************
   If val(txtRemainWater.Text) > 0 Then
       '«·⁄„Ì· „œÌ‰
              Notevalue = Abs(val(txtRemainWater.Text))
   LngDevNO = LngDevNO + 1
   
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
      
      If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… «·„Ì«Â ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  
             StrTempAccountCode = Account_Code_dynamic83

        LngDevNO = LngDevNO + 1
  
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «·„Ì«… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
            
  End If
  
   
   If val(txtRemainService.Text) > 0 Then
       '«·⁄„Ì· „œÌ‰
        
       '«·⁄„Ì· „œÌ‰
              Notevalue = Abs(val(txtRemainService.Text))
   LngDevNO = LngDevNO + 1
   
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
      
      If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… «·Œœ„«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  
             StrTempAccountCode = Account_Code_dynamic85

        LngDevNO = LngDevNO + 1
  
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «·Œœ„«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
              
  End If
   
   
   
   If val(TxtInsurance.Text) > 0 Then
               
               Notevalue = Abs(val(TxtInsurance.Text))
   LngDevNO = LngDevNO + 1
  
                 If SystemOptions.CreateInsuranceAccountForCustomers Then
    StrTempAccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText), "InsuranceAccount")
 Else
 StrTempAccountCode = Account_Code_dynamic82
  End If
        
        
      If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… «· √„Ì‰ «·„” —œ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
 
 
           StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
 
          
          
        LngDevNO = LngDevNO + 1
  
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «· √„Ì‰ «·„” —œ  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
              
  End If
   
   
   
'**************************************************
     
     
     
If val(TxtBillPrice.Text) > 0 Then
       '  «·ﬂÂ—»«¡
       Notevalue = Abs(val(TxtBillPrice.Text))
       
               LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… ›Ê« Ì—  «·ﬂÂ—»«¡ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
   LngDevNO = LngDevNO + 1
   StrTempAccountCode = Account_Code_dynamic84
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„…  ›Ê« Ì—  «·ﬂÂ—»«¡ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            

  

            
            
  End If
     
     
     
'**************************************************************************’Ì«‰…
     
   '**************************************************
     
If val(TxtAmountDely.Text) > 0 Then
       '  «·Œ’Ê„« 
       Notevalue = (val(TxtAmountDely.Text))
       
               LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„…    Œ’„ «Ì«„ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
   LngDevNO = LngDevNO + 1
   
   
   If val(TxtDayPricen.Text) * val(TxtDayLate) > 0 Then
   
   Notevalue = Round(val(TxtDayPricen.Text) * val(TxtDayLate), 2)
     StrTempAccountCode = Account_Code_dynamic80
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„…   Œ’„ «Ì«„  ﬁÌ„… «ÌÃ«—  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
   End If
   
 
            
         LngDevNO = LngDevNO + 1
   
   
   If val(TxtWaterPrice.Text) * val(TxtDayLate) > 0 Then
   Notevalue = Round(val(TxtWaterPrice.Text) * val(TxtDayLate), 2)
     StrTempAccountCode = Account_Code_dynamic83
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„…   Œ’„ «Ì«„  ﬁÌ„… „Ì«Â  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
   End If
   

  
        LngDevNO = LngDevNO + 1
   
   
   If val(TxtService.Text) * val(TxtDayLate) > 0 Then
   Notevalue = Round(val(TxtService.Text) * val(TxtDayLate), 2)
     StrTempAccountCode = Account_Code_dynamic85
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„…   Œ’„ «Ì«„  ﬁÌ„… Œœ„«   ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
   End If
   

            
            
  End If
     
     
     
'**************************************************************************’Ì«‰…
  
     
     If val(lbl(12).Caption) > 0 Then
       '  «·⁄„Ì·
       Notevalue = Abs(val(lbl(12).Caption))
       
               LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… ›Ê« Ì—  «·ﬂÂ—»«¡ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            

            

  

            
            
  End If
     


     
     
     
          For i = Me.Fg.FixedRows To Fg.Rows - 1
    
                  If val(Fg.TextMatrix(i, Fg.ColIndex("total"))) > 0 And Fg.TextMatrix(i, Fg.ColIndex("Accountsus")) <> "" Then
              Notevalue = val(Fg.TextMatrix(i, Fg.ColIndex("total")))
            
               LngDevNO = LngDevNO + 1
   StrTempAccountCode = Fg.TextMatrix(i, Fg.ColIndex("Accountsus"))
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„…  ’Ì«‰…    " & Fg.TextMatrix(i, Fg.ColIndex("group")), general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
                   
                      End If
         
    
  
        Next i
  
ErrTrap:
End Function



Private Sub Form_Load()

    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String, MySQL As String
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic



   Dim s As String, MySQL As String
    s = " SELECT"
s = s & "        TblFiterWaiverDe.IDFItWaiv,TblAqar.aqarname,"
s = s & "        dbo.TblFiterWaiverDe.Remark,"
s = s & "        Sum(dbo.TblFiterWaiverDe.Price * dbo.TblFiterWaiverDe.[Count]) AS Total ,"
s = s & "        dbo.TblAqrCompenetDet.Name AS nameDet"
       
s = s & " From dbo.TblFiterWaiverDe"
s = s & "        LEFT OUTER JOIN dbo.TblAqrCompenetDet"
s = s & "             ON  dbo.TblFiterWaiverDe.IDItem = dbo.TblAqrCompenetDet.ID"
s = s & "        LEFT OUTER JOIN dbo.TblAqrCompenet"
s = s & "             ON  dbo.TblFiterWaiverDe.GroupID = dbo.TblAqrCompenet.ID"
            
s = s & "             LEFT OUTER JOIN tblFiterWaiver ON TblFiterWaiverDe.IDFItWaiv =tblFiterWaiver.ID"
s = s & "             LEFT OUTER JOIN TblAqar ON TblAqar.Aqarid =  tblFiterWaiver.BulidID"

s = s & " Where count <> 0"
s = s & " GROUP BY dbo.TblFiterWaiverDe.Remark,dbo.TblAqrCompenetDet.Name,TblFiterWaiverDe.IDFItWaiv,TblAqar.aqarname"
    

db_createOrUpdateviewSQL "View_WaiverExpens", s


       s = " SELECT   TotalExp = (SELECT SUM(COUNT * Price) TotalExpe FROM TblFiterWaiverDe),"
s = s & "       CountContract = (SELECT COUNT(*) FROM TblContract ), "
s = s & "                          Cashing = SUM("
s = s & "                   Case Notes.NoteCashingType"
s = s & "                        WHEN 0 THEN (Note_Value)"
s = s & "                        ELSE 0"
s = s & "                   End"
s = s & "               ),"
s = s & "               Commission        = SUM("
s = s & "                   Case Notes.CashingType"
s = s & "                        WHEN 12 THEN (Note_Value)"
s = s & "                        ELSE 0"
s = s & "                   End"
s = s & "               ),"
s = s & "               Arbon             = SUM(CASE Notes.CashingType WHEN 9 THEN (Note_Value) ELSE 0 END),"
s = s & "               ValueTransfer     = SUM("
s = s & "                   Case Notes.NoteCashingType"
s = s & "                        WHEN 3 THEN (Note_Value)"
s = s & "                        ELSE 0"
s = s & "                   End"
s = s & "               )"
s = s & "        From Notes "
db_createOrUpdateviewSQL "View_TotalsIq", s

MySQL = " SELECT     dbo.Notes.NoteID,tu.UserName, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value, dbo.Notes.NoteDateH,"
MySQL = MySQL & "                       dbo.Notes.ContractNo, dbo.Notes.ContNo, dbo.Notes.commission, dbo.Notes.rent, dbo.Notes.Water, dbo.Notes.FilterID, dbo.Notes.FIlterTotal, dbo.Notes.Instrunce,"
MySQL = MySQL & "                       dbo.Notes.comX, dbo.Notes.ComY, dbo.Notes.CommissionOut, dbo.Notes.NoteOrBonID, dbo.Notes.comXold, dbo.Notes.ComYold, dbo.Notes.NoteOrBonValue,"
MySQL = MySQL & "                       dbo.Notes.NoteOrBonSereal, dbo.Notes.Telephone, dbo.Notes.CashingType, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
MySQL = MySQL & "                       dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.Notes.renterName, dbo.Notes.NoteCashingType, dbo.Notes.BankName, dbo.Notes.DueDate,"
MySQL = MySQL & "                       dbo.Notes.ChqueNum, dbo.Notes.Remark, dbo.Notes.Remark2, dbo.Notes.ToPriodDateH, dbo.Notes.FrmPriodDateH, dbo.Notes.ToPriodDate, dbo.Notes.FrmPriodDate,"
MySQL = MySQL & "                       dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAqarDetai.Id, dbo.TblAqarDetai.unitno,"
MySQL = MySQL & "                       dbo.TblAqarDetai.unittype, dbo.TblAqarDetai.Aqarid, TblAqar_1.aqarname, TblAkarUnit_2.name, TblAkarUnit_2.namee, dbo.Notes.akarid,"
                      MySQL = MySQL & " TblAqar_1.aqarname AS aqarname2, dbo.Notes.unittype AS unittype2, TblAkarUnit_1.name AS name2, TblAkarUnit_1.namee AS namee2, dbo.Notes.Electricity,"
MySQL = MySQL & "                       dbo.Notes.BankID, dbo.BanksData.BankNamee, dbo.BanksData.BankName AS BankName2, dbo.TblNotesSales.rate, dbo.TblNotesSales.valu,"
MySQL = MySQL & "                       dbo.TblNotesSales.Type, dbo.TblNotesSales.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.Notes.Servce,"
 MySQL = MySQL & "                      dbo.Notes.RemaiValue, dbo.ContracttBillInstallmentsDone.WaterPayed, dbo.ContracttBillInstallmentsDone.RentValuePayed,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.CommissionsPayed, dbo.ContracttBillInstallmentsDone.InsurancePayed, dbo.ContracttBillInstallmentsDone.ElectricPayed,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.TelandNetPayed, dbo.ContracttBillInstallmentsDone.RecordDate, dbo.ContracttBillInstallmentsDone.RecordDateH,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.total, dbo.ContracttBillInstallmentsDone.[Value], dbo.ContracttBillInstallmentsDone.InstallNo,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.VATPayed, dbo.ContracttBillInstallmentsDone.VATValue, dbo.ContracttBillInstallmentsDone.ActVAT,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.Commisionvalue , dbo.ContracttBillInstallmentsDone.OldValuePayed, dbo.ContracttBillInstallmentsDone.PaymentType"
MySQL = MySQL & " FROM         dbo.ContracttBillInstallmentsDone RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.Notes ON dbo.ContracttBillInstallmentsDone.NoteID = dbo.Notes.NoteID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblNotesSales LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee ON dbo.TblNotesSales.EmpID = dbo.TblEmployee.Emp_ID ON dbo.Notes.NoteID = dbo.TblNotesSales.NoteID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_1 ON dbo.Notes.unittype = TblAkarUnit_1.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqar TblAqar_1 ON dbo.Notes.akarid = TblAqar_1.Aqarid LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqarDetai LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_2 ON dbo.TblAqarDetai.unittype = TblAkarUnit_2.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqar TblAqar_2 ON dbo.TblAqarDetai.Aqarid = TblAqar_2.Aqarid ON dbo.Notes.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID"
MySQL = MySQL & "                                     LEFT OUTER JOIN dbo.TblUsers AS tu"
MySQL = MySQL & "                                   ON  dbo.Notes.UserID = tu.UserID"
'Where (dbo.Notes.NoteID = 4441)
MySQL = MySQL & " Where "
MySQL = MySQL & " (dbo.Notes.NoteType = 4)"
MySQL = MySQL & "        AND ISNULL(contNo, 0) <> 0"

db_createOrUpdateviewSQL "View_Waiver", MySQL



    
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
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.GetIqarUnit -2, 1, DcbUnitNo
    Dcombos.GetCustomersSuppliers 56, Me.dcCustomer
     My_SQL = "select UserID,UserName From tblUsers "
  
  

    SetDtpickerDate Me.XPDtbTrans
   fill_combo DCboUserName, My_SQL
    Dcombos.GetIqar DcbIqara
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblFiterWaiver "
      If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
      StrSQL = StrSQL & "   where BranchID=" & Current_branch & "     Order By ID"
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

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

    Me.Caption = "Filter waiver"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
   Me.lblBr.Caption = "Branch"
   Me.lblDataCli.Caption = "Data Of Renter"
  Me.Label1(5).Caption = "Renter"
  lbl(13).Caption = "Iqar"
    lbl(14).Caption = "Based ContNo"
    Label1(15) = "Type"
    lbl(15).Caption = "UnitNo"
    lbl(16) = "Insurance"
lbl(17).Caption = "End Date"
lbl(18).Caption = "Electricity"
lbl(19).Caption = "AccountNo"
lbl(20).Caption = "DateFiltering"
  lbl(21).Caption = "No Day"
    lbl(22).Caption = "LatePrice"
    lbl(23).Caption = "Remarks Filtering "
       lbl(24).Caption = "Total "
       Cmdd.Caption = "Calcu"
       lbl(3).Caption = "Writing"
       lbl(5).Caption = "Writing"
       lbl(11).Caption = "Writing"
       lbl(10).Caption = " Amount owed From the tenant after liquidation"
       lbl(2).Caption = " Amount owed to the tenant after liquidation"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
lbl(9).Caption = "Net"
XPTab301.Caption = "Data"
    With Me.Fg
        .TextMatrix(0, .ColIndex("serial")) = "NO"
        .TextMatrix(0, .ColIndex("group")) = "Name"
        .TextMatrix(0, .ColIndex("price")) = "Price"
         .TextMatrix(0, .ColIndex("count")) = "Count"
.TextMatrix(0, .ColIndex("total")) = "Total"
.TextMatrix(0, .ColIndex("remark")) = "Remark"
    End With


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

Private Sub TxtDayPricentotal_Change()
txtRemainRent.Text = val(TxtDayPricentotal.Text) - val(txtRemainRent.Text)
End Sub

Private Sub txtDaysValueIncomplete_Change()
CalcTotal
End Sub

Private Sub txtDaysValueIncrease_Change()
CalcTotal
End Sub

Private Sub TxtForRenter_Change()
If val(TxtForRenter.Text) > 0 Then
lbll(9).Caption = WriteNo(Round(Me.TxtForRenter.Text, 3), 0)
Else
lbll(9).Caption = ""
End If
txtnet.Text = Round(val(TxtForRenter.Text) - val(TxtOFRenter.Text), 3)
'TxtNet.text = val(Me.TxtForRenter.text) - val(Me.TxtOFRenter.text)
End Sub




Private Sub TxtInsurance_Change()
TxtOFRenter.Text = val(TxtOFRenter.Text) + val(TxtInsurance.Text)
End Sub

Private Sub txtLastInvoiceRead_Change()
CalcTotalCounter
End Sub
Private Sub CalcTotalCounter()
TxtDiff = Round(val(txtLastInvoiceRead2) - val(txtLastInvoiceRead), 2)
txtR = Round(val(TxtPrice) * val(TxtDiff), 2)
txtTotalCounter = Round(val(txtR) + val(txtPrevBalance) + val(txtServiceCounter), 2)
ReLineGrid
End Sub

Private Sub txtLastInvoiceRead2_Change()
CalcTotalCounter
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
               
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
         
    
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





Private Sub txtNet_Change()

If val(TxtOFRenter) > 0 Then
lbll(0).Caption = WriteNo(Round(val(Me.txtnet.Text), 3), 0)
Else
lbll(0).Caption = ""
End If


End Sub

Private Sub TxtOFRenter_Change()

If val(TxtOFRenter) > 0 Then
lbll(11).Caption = WriteNo(Round(val(Me.TxtOFRenter.Text), 3), 0)
Else
lbll(11).Caption = ""
End If


End Sub

Private Sub TxtOrder_Change()
ReLineGrid
'RetriveOrder TxtOrder, 0
End Sub

Private Sub TxtOrder_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmIqarContractSearch
'FrmIqarContractSearch.fg.TextMatrix(fg.Row, fg.ColIndex("NoteSerial"))=me.Text15
FrmIqarContractSearch.m_RetrunType = 2
FrmIqarContractSearch.Show


End If
End Sub

Private Sub txtPrevBalance_Change()
CalcTotalCounter
End Sub

Private Sub txtPrice_Change()
CalcTotalCounter

End Sub

Private Sub TxtService_Change()
txtRemainService.Text = val(TxtService.Text) - val(txtTelandNetPayed.Text)
End Sub

Private Sub txtServiceCounter_Change()
CalcTotalCounter
End Sub

Private Sub TxtWaterPriceotal_Change()
txtRemainWater.Text = val(TxtWaterPriceotal.Text) - val(txtWaterPayed.Text)
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
            rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    TxtContNo.Text = IIf(IsNull(rs("ContNo").value), "", val(rs("ContNo").value))
   
   Me.Txtorder.Text = IIf(IsNull(rs("ContractNo").value), "", (rs("ContractNo").value))
    Me.TxtNoteID.Text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
  Me.TxtNoteserial.Text = IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value)
   

    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    NourHijriCal1.value = IIf(IsNull(rs("RecordDateH").value), "", rs("RecordDateH").value)
     Dcbranch.BoundText = val(IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value))
    dcCustomer.BoundText = val(IIf(IsNull(rs("RenterID").value), "", rs("RenterID").value))
    
    DcbIqara.BoundText = val(IIf(IsNull(rs("BulidID").value), "", rs("BulidID").value))
      DcbUnitType.BoundText = val(IIf(IsNull(rs("unittype").value), "", rs("unittype").value))
    
    DcbUnitNo.BoundText = val(IIf(IsNull(rs("ApartmentID").value), "", rs("ApartmentID").value))
    TxtInsurance.Text = val(IIf(IsNull(rs("Insurance").value), 0, rs("Insurance").value))
     EndDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    EndDateH.value = IIf(IsNull(rs("EndDateH").value), "", rs("EndDateH").value)
    FilterDate.value = IIf(IsNull(rs("FilterDate").value), Date, rs("FilterDate").value)
    FilterDateH.value = IIf(IsNull(rs("FilterDateH").value), "", rs("FilterDateH").value)
    '
     
    txtLastInvoiceRead.Text = val(IIf(IsNull(rs("LastInvoiceRead").value), 0, rs("LastInvoiceRead").value))
    txtLastInvoiceRead2.Text = val(IIf(IsNull(rs("LastInvoiceRead2").value), 0, rs("LastInvoiceRead2").value))
    TxtDiff.Text = val(IIf(IsNull(rs("Diff").value), 0, rs("Diff").value))
    TxtPrice.Text = val(IIf(IsNull(rs("Price").value), 0, rs("Price").value))
    txtR.Text = val(IIf(IsNull(rs("R").value), 0, rs("R").value))
    txtPrevBalance.Text = val(IIf(IsNull(rs("PrevBalance").value), 0, rs("PrevBalance").value))
    txtServiceCounter.Text = val(IIf(IsNull(rs("ServiceCounter").value), 0, rs("ServiceCounter").value))
    txtTotalCounter.Text = val(IIf(IsNull(rs("TotalCounter").value), 0, rs("TotalCounter").value))
     
     
     TxtForRenter.Text = val(IIf(IsNull(rs("ForRenter").value), 0, rs("ForRenter").value))
      TxtOFRenter.Text = val(IIf(IsNull(rs("OFRenter").value), 0, rs("OFRenter").value))
    '
     TxtBillPrice.Text = val(IIf(IsNull(rs("BillPrice").value), 0, rs("BillPrice").value))
     Me.txtnet.Text = val(IIf(IsNull(rs("net").value), 0, rs("net").value))
     TxtAccountNo.Text = IIf(IsNull(rs("AccountNo").value), "", rs("AccountNo").value)
   TxtDayLate.Text = IIf(IsNull(rs("DayNo").value), "", rs("DayNo").value)
     TxtAmountDely.Text = IIf(IsNull(rs("AmountDely").value), "", rs("AmountDely").value)
'*******************************************************************************************
 



TxtContractDays.Text = IIf(IsNull(rs("ContractDays").value), 0, rs("ContractDays").value)
TxtActualDays.Text = IIf(IsNull(rs("ActualDays").value), 0, rs("ActualDays").value)
TxtWaterPrice.Text = IIf(IsNull(rs("WaterPrice").value), 0, rs("WaterPrice").value)
TxtDayPricen.Text = IIf(IsNull(rs("DayPricen").value), 0, rs("DayPricen").value)
 
txtServicePrice.Text = IIf(IsNull(rs("ServicePrice").value), 0, rs("ServicePrice").value)
TxtWaterPriceotal.Text = IIf(IsNull(rs("WaterPriceotal").value), 0, rs("WaterPriceotal").value)
TxtDayPricentotal.Text = IIf(IsNull(rs("DayPricentotal").value), 0, rs("DayPricentotal").value)
TxtService.Text = IIf(IsNull(rs("Service").value), 0, rs("Service").value)
txtWaterPayed.Text = IIf(IsNull(rs("WaterPayed").value), 0, rs("WaterPayed").value)
TxtRentValuePayed.Text = IIf(IsNull(rs("RentValuePayed").value), 0, rs("RentValuePayed").value)
txtTelandNetPayed.Text = IIf(IsNull(rs("TelandNetPayed").value), 0, rs("TelandNetPayed").value)
txtRemainWater.Text = IIf(IsNull(rs("RemainWater").value), 0, rs("RemainWater").value)
txtRemainRent.Text = IIf(IsNull(rs("RemainRent").value), 0, rs("RemainRent").value)
txtRemainService.Text = IIf(IsNull(rs("RemainService").value), 0, rs("RemainService").value)
txtDaysValueIncrease.Tag = ""
txtDaysValueIncomplete.Tag = ""
txtDaysValueIncrease.Text = IIf(IsNull(rs("DaysValueIncrease").value), 0, rs("DaysValueIncrease").value)
txtDaysValueIncomplete.Text = IIf(IsNull(rs("DaysValueIncomplete").value), 0, rs("DaysValueIncomplete").value)

txtDayValueInc.Text = IIf(IsNull(rs("DayValueInc").value), 0, rs("DayValueInc").value)
txtDayCountInc.Text = IIf(IsNull(rs("DayCountInc").value), 0, rs("DayCountInc").value)
txtDayValueIncomplete.Text = IIf(IsNull(rs("DayValueIncomplete").value), 0, rs("DayValueIncomplete").value)
txtDayCountIncomplete.Text = IIf(IsNull(rs("DayCountIncomplete").value), 0, rs("DayCountIncomplete").value)




 If IsNull(rs.Fields("outflow").value) Then
 chkoutflow.value = vbUnchecked
 Else
 chkoutflow.value = vbChecked
   
 End If
 
  If IsNull(rs.Fields("outCondition").value) Then
 chkoutCondition.value = vbUnchecked
 Else
 chkoutCondition.value = vbChecked
   
 End If
 
 
'*******************************************************************************************

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)

   
    Set RsDetails = New ADODB.Recordset

StrSQL = "SELECT     dbo.TblFiterWaiverDe.IDFItWaiv, dbo.TblFiterWaiverDe.[Count], dbo.TblFiterWaiverDe.Remark, dbo.TblFiterWaiverDe.Price, "
 StrSQL = StrSQL & "                     dbo.TblAqrCompenetDet.Name AS nameDet, dbo.TblFiterWaiverDe.IDItem, dbo.TblAqrCompenet.Name, dbo.TblFiterWaiverDe.GroupID"
StrSQL = StrSQL & ",TblFiterWaiverDe.Accountsus   FROM         dbo.TblFiterWaiverDe LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblAqrCompenetDet ON dbo.TblFiterWaiverDe.IDItem = dbo.TblAqrCompenetDet.ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblAqrCompenet ON dbo.TblFiterWaiverDe.GroupID = dbo.TblAqrCompenet.ID"
StrSQL = StrSQL & "  Where (dbo.TblFiterWaiverDe.IDFItWaiv = " & val(XPTxtID.Text) & ")"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
 
   Dim temp, k, j As Integer
j = 0
temp = -1
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.Fg
      '  RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails.RecordCount
k = 0
        For i = .FixedRows To .Rows - 1
    j = j + 1
    k = k + 1
   
    If val(RsDetails("IDItem").value) <> 0 And val(RsDetails("GroupID").value) = 0 Then
            .TextMatrix(k, .ColIndex("serial")) = j
            .TextMatrix(k, .ColIndex("id")) = 0
             .TextMatrix(k, .ColIndex("iditem")) = val(IIf(IsNull(RsDetails("IDItem").value), 0, RsDetails("IDItem").value))
            .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("nameDet").value), "", RsDetails("nameDet").value)
             .TextMatrix(k, .ColIndex("price")) = val(IIf(IsNull(RsDetails("Price").value), 0, RsDetails("Price").value))
              .TextMatrix(k, .ColIndex("remark")) = IIf(IsNull(RsDetails("Remark").value), "", RsDetails("Remark").value)
               .TextMatrix(k, .ColIndex("count")) = val(IIf(IsNull(RsDetails("Count").value), Null, RsDetails("Count").value))
               .TextMatrix(k, .ColIndex("total")) = val(RsDetails("Count").value) * val(RsDetails("Price").value)
                .TextMatrix(k, .ColIndex("Accountsus")) = IIf(IsNull((RsDetails("Accountsus").value)), "", (RsDetails("Accountsus").value))
               
   Else
   
   If val(RsDetails("IDItem").value) = 0 And val(RsDetails("GroupID").value) <> 0 Then
           
               .TextMatrix(k, .ColIndex("id")) = val(IIf(IsNull(RsDetails("GroupID").value), 0, RsDetails("GroupID").value))
               .TextMatrix(k, .ColIndex("iditem")) = 0
           .TextMatrix(k, .ColIndex("serial")) = ""
            .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("Name").value), "", RsDetails("Name").value)
             .TextMatrix(k, .ColIndex("price")) = ""
                    .TextMatrix(k, .ColIndex("remark")) = ""
               .TextMatrix(k, .ColIndex("count")) = ""
                      .Cell(flexcpBackColor, k, 1, k, 7) = &H80C0FF
      '    .TextMatrix(k, .ColIndex("Accountsus")) = (RsDetails("Accountsussub").value)
    .TextMatrix(k, .ColIndex("Accountsus")) = IIf(IsNull((RsDetails("Accountsus").value)), "", (RsDetails("Accountsus").value))
            j = 0
          End If
           End If
            RsDetails.MoveNext
         
        Next i
    'ReLineGridCount
    
    Cmdd_Click
End With

    End If

    RsDetails.Close
    Set RsDetails = Nothing
    '//////////////////////////////////////////
  '  Set RsDetails1 = New ADODB.Recordset

    
    StrSQL = "Select *  From TblFiterWaiverDet2 Where MasterID=" & val(Me.XPTxtID.Text)
    loadgrid StrSQL, Grd, True, True
CalcTotal
  '  RsDetails1.Close
  '  Set RsDetails1 = Nothing
    
  '  fillapprovData
    'ReLineGrid
  ' GetContract val(TxtOrder)
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub TotalItemPrice()
 
End Sub
Public Sub GetContract(ByVal mContractNo As Long)
    Dim s As String, mCount As Long
    Dim rsDummyCount  As New ADODB.Recordset
    s = "Select Count(*) CC from TblContract Where NoteSerial1 = " & mContractNo
    rsDummyCount.Open s, Cn, adOpenKeyset, adLockReadOnly
    If Not rsDummyCount.EOF Then
        mCount = val(rsDummyCount!CC & "")
    End If
    Dim rsDummy  As New ADODB.Recordset
    s = "Select * from TblContract Where NoteSerial1 = " & mContractNo & " Order By ContNo "
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    Grd.Rows = 1
 
    Dim mActualDays As Double
    Dim mContractDays As Double
    Dim mDayPricentotal As Double
    Dim mDayPricen As Double
    Dim WaterPayed     As Double
    Dim RentValuePayed As Double
    Dim TelandNetPayed As Double
    Dim mRentValuePayed As Double
    Dim mWaterPrice As Double
    Dim mWaterPriceotal As Double, mRemainRent As Double, mRemainWater As Double, mRemainDays As Double
    Dim mService As Double
    Dim mServicePrice As Double
    Dim mRemainService As Double
    Dim mDayLate As Double
    Dim mAmountDely As Double
    

  Dim CommissionsPayed  As Double
  Dim InsurancePayed    As Double
  
  Dim ElectricPayed   As Double
  
  Dim payed As Double
  Dim VATPayed As Double
  
    Dim mDaysValue As Double
    Dim mTotalDept As Double
    Dim mTotalRight As Double
    Dim mTotalDaysValue As Double
    Dim mTotalDept2 As Double
    Dim mTotalRight2 As Double
    Dim mElictricPrice As Double
    Dim mRemElictricPrice  As Double
    Dim TotalOldValue As Double, RemainCommissions As Double
    Dim DaysValueIncrease As Double
    Dim DaysValueIncomplete As Double
    txttotal1 = 0
    txttotal1 = 0
    txtWaterPayed = ""
    TxtDayLate = ""
    TxtActualDays = ""
    txtTelandNetPayed = ""
    TxtRentValuePayed = ""
    TxtDayPricentotal = ""
    TxtInsurance = ""
    txtRemainService = ""
    TxtBillPrice = ""
    txtRemainWater = ""
    txtRemainRent = ""
    Do While Not rsDummy.EOF
        If rsDummy!StrDate = FilterDate.value Then
            mActualDays = 1
        Else
            mActualDays = (DateDiff("d", rsDummy!StrDate & "", FilterDate.value))
        End If
        
        If rsDummy!StrDate = rsDummy!EndDate Then
            mContractDays = 1
        Else
            mContractDays = (DateDiff("d", rsDummy!StrDate, rsDummy!EndDate))
        End If
        

        If val(mContractDays) <> 0 Then
            mDayPricen = Round(IIf(IsNull(rsDummy("TotalContract").value), 0, rsDummy("TotalContract").value) / val(mContractDays), 2)
        End If
        
        mDayPricentotal = val(mDayPricen) * val(mActualDays)
        
        
        payed = getinsttPayedTocontract2(val(rsDummy!ContNo & ""), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed, TotalOldValue, , , VATPayed)

        mRentValuePayed = RentValuePayed
        mRemainRent = RentValuePayed
        mRemainWater = WaterPayed
        RemainCommissions = CommissionsPayed
        mRemElictricPrice = ElectricPayed
       
        mRemainService = TelandNetPayed
        mDayLate = val(mContractDays) - val(mActualDays)
        mAmountDely = Round(val(mWaterPrice) * val(mDayLate), 2) + Round(val(mDayPricen) * val(mDayLate), 2) + Round(val(mServicePrice) * val(mDayLate), 2)
   
    
     
        
        With Grd
            .AddItem 1
             If mCount = .Rows - 1 Then
                mRemainDays = (DateDiff("d", FilterDate.value, rsDummy!EndDate & ""))
                If val(mContractDays) <> 0 Then
                    mDaysValue = Round((IIf(IsNull(rsDummy("TotalValue").value), 0, rsDummy("TotalValue").value) / val(mContractDays)) * Abs(mRemainDays), 2)
                    mDaysValue = Round((IIf(IsNull(rsDummy("TotalValue").value), 0, rsDummy("TotalValue").value) / val(mContractDays)) * mRemainDays, 2)
                End If
             
             End If
            mTotalDept = mRemainRent + mRemainWater + mRemElictricPrice + RemainCommissions + mRemainService + val(TotalOldValue) + IIf(mDaysValue > 0, val(mDaysValue), 0)
            mTotalRight = val(val(rsDummy!InsuranceValue & "") + IIf(mDaysValue < 0, Abs(val(mDaysValue)), 0))
            DaysValueIncrease = DaysValueIncrease + IIf(mDaysValue < 0, Abs(val(mDaysValue)), 0)
            DaysValueIncomplete = DaysValueIncomplete + IIf(mDaysValue > 0, Abs(val(mDaysValue)), 0)
            
            
            mTotalDept2 = mTotalDept2 + mRemainRent + mRemainWater + RemainCommissions + mRemElictricPrice + mRemainService + val(TotalOldValue)
            mTotalRight2 = mTotalRight2 + val(val(rsDummy!InsuranceValue & ""))
            .TextMatrix(.Rows - 1, .ColIndex("ContNo")) = rsDummy!ContNo & ""
            .TextMatrix(.Rows - 1, .ColIndex("StartDate")) = rsDummy!StrDate & ""
            .TextMatrix(.Rows - 1, .ColIndex("StartDateh")) = rsDummy!Fromdateh & ""
            .TextMatrix(.Rows - 1, .ColIndex("EndDate")) = rsDummy!EndDate & ""
            .TextMatrix(.Rows - 1, .ColIndex("EndDateH")) = rsDummy!todateH & ""
            .TextMatrix(.Rows - 1, .ColIndex("RemainRent")) = mRemainRent
            .TextMatrix(.Rows - 1, .ColIndex("RemainWater")) = mRemainWater
            .TextMatrix(.Rows - 1, .ColIndex("BillPrice")) = mRemElictricPrice
            .TextMatrix(.Rows - 1, .ColIndex("RemainService")) = mRemainService
            .TextMatrix(.Rows - 1, .ColIndex("insurance")) = InsurancePayed
            '.TextMatrix(.Rows - 1, .ColIndex("DaysValue")) = rsDummy!InsuranceValue & ""
            .TextMatrix(.Rows - 1, .ColIndex("RemainDays")) = mRemainDays
            .TextMatrix(.Rows - 1, .ColIndex("DaysValue")) = mDaysValue
            .TextMatrix(.Rows - 1, .ColIndex("RemainCommissions")) = RemainCommissions
            
            .TextMatrix(.Rows - 1, .ColIndex("OldRent")) = TotalOldValue
            '.TextMatrix(.Rows - 1, .ColIndex("DaysValue")) = mDayPricentotal
            .TextMatrix(.Rows - 1, .ColIndex("RentValuePayed")) = RentValuePayed
            .TextMatrix(.Rows - 1, .ColIndex("WaterPayed")) = WaterPayed
            .TextMatrix(.Rows - 1, .ColIndex("TelandNetPayed")) = TelandNetPayed
            .TextMatrix(.Rows - 1, .ColIndex("ActualDays")) = mContractDays
            .TextMatrix(.Rows - 1, .ColIndex("DayLate")) = mDayLate
            .TextMatrix(.Rows - 1, .ColIndex("AmountDely")) = mAmountDely
            .TextMatrix(.Rows - 1, .ColIndex("TotalDept")) = mTotalDept
            .TextMatrix(.Rows - 1, .ColIndex("TotalRight")) = mTotalRight
            .TextMatrix(.Rows - 1, .ColIndex("TotalStill")) = RentValuePayed + CommissionsPayed + WaterPayed + ElectricPayed + TelandNetPayed + TotalOldValue
            
            txtWaterPayed = val(txtWaterPayed) + WaterPayed
            TxtDayLate = val(TxtDayLate) + WaterPayed
            TxtActualDays = val(TxtActualDays) + mActualDays
            txtTelandNetPayed = val(txtTelandNetPayed) + TelandNetPayed
            TxtRentValuePayed = val(TxtRentValuePayed) + RentValuePayed
            TxtDayPricentotal = val(TxtDayPricentotal) + mDayPricentotal
            TxtInsurance = InsurancePayed
            txtRemainService = val(txtRemainService) + mRemainService
            TxtBillPrice = val(TxtBillPrice) + mRemElictricPrice
            txtRemainWater = val(txtRemainWater) + mRemainWater
            txtRemainRent = val(txtRemainRent) + mRemainRent
                         
        End With
    
        rsDummy.MoveNext
    Loop
        With Grd
           .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
            .IsSubtotal(.Rows - 1) = True
            Dim SngTotal As Single
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainRent"), .Rows - 1, .ColIndex("RemainRent"))
            .TextMatrix(.Rows - 1, .ColIndex("RemainRent")) = SngTotal
        
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainWater"), .Rows - 1, .ColIndex("RemainWater"))
            .TextMatrix(.Rows - 1, .ColIndex("RemainWater")) = SngTotal
            
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillPrice"), .Rows - 1, .ColIndex("BillPrice"))
                    .TextMatrix(.Rows - 1, .ColIndex("BillPrice")) = SngTotal
                    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainService"), .Rows - 1, .ColIndex("RemainService"))
                    .TextMatrix(.Rows - 1, .ColIndex("RemainService")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TelandNetPayed"), .Rows - 1, .ColIndex("TelandNetPayed"))
                    .TextMatrix(.Rows - 1, .ColIndex("TelandNetPayed")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainCommissions"), .Rows - 1, .ColIndex("RemainCommissions"))
                    .TextMatrix(.Rows - 1, .ColIndex("RemainCommissions")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalStill"), .Rows - 1, .ColIndex("TotalStill"))
                    .TextMatrix(.Rows - 1, .ColIndex("TotalStill")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OldRent"), .Rows - 1, .ColIndex("OldRent"))
                    .TextMatrix(.Rows - 1, .ColIndex("OldRent")) = SngTotal
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
        .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
     End With
    txtDaysValueIncrease.Tag = Round(DaysValueIncrease, 2)
    txtDaysValueIncomplete.Tag = Round(DaysValueIncomplete, 2)
    
    txtDaysValueIncrease = Round(DaysValueIncrease, 2)
    txtDaysValueIncomplete = Round(DaysValueIncomplete, 2)
    txtDayValueInc = Round(IIf(mDaysValue < 0, Abs(val(mDaysValue)), 0) / IIf(Abs(mRemainDays), Abs(mRemainDays), 1), 2)
    txtDayValueIncomplete = Round(IIf(mDaysValue > 0, Abs(val(mDaysValue)), 0) / IIf(Abs(mRemainDays), Abs(mRemainDays), 1), 2)
    
    txtDayCountIncomplete = Round(IIf(mRemainDays > 0, Abs(val(mRemainDays)), 0), 2)
    txtDayCountInc = Round(IIf(mRemainDays < 0, Abs(val(mRemainDays)), 0), 2)
    txttotal1 = val(mTotalDept2) + val(txtDaysValueIncrease)
    txttotal2 = val(mTotalRight2) + val(txtDaysValueIncomplete)
    ReLineGrid
End Sub
Private Sub CalcTotal()
  
  
  With Grd
           If .IsSubtotal(.Rows - 1) = True Then
           .RemoveItem .Rows - 1
           End If
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
                .IsSubtotal(.Rows - 1) = True
            
            
            Dim SngTotal As Single
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainRent"), .Rows - 1, .ColIndex("RemainRent"))
            .TextMatrix(.Rows - 1, .ColIndex("RemainRent")) = SngTotal
        
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainWater"), .Rows - 1, .ColIndex("RemainWater"))
            .TextMatrix(.Rows - 1, .ColIndex("RemainWater")) = SngTotal
            
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillPrice"), .Rows - 1, .ColIndex("BillPrice"))
                    .TextMatrix(.Rows - 1, .ColIndex("BillPrice")) = SngTotal
                    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainService"), .Rows - 1, .ColIndex("RemainService"))
                    .TextMatrix(.Rows - 1, .ColIndex("RemainService")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TelandNetPayed"), .Rows - 1, .ColIndex("TelandNetPayed"))
                    .TextMatrix(.Rows - 1, .ColIndex("TelandNetPayed")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainCommissions"), .Rows - 1, .ColIndex("RemainCommissions"))
                    .TextMatrix(.Rows - 1, .ColIndex("RemainCommissions")) = SngTotal
            
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("insurance"), .Rows - 1, .ColIndex("insurance"))
                    .TextMatrix(.Rows - 1, .ColIndex("insurance")) = SngTotal
                    
               txttotal2 = SngTotal + val(txtDaysValueIncomplete)
          
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OldRent"), .Rows - 1, .ColIndex("OldRent"))
                    .TextMatrix(.Rows - 1, .ColIndex("OldRent")) = SngTotal
                    
  SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalStill"), .Rows - 1, .ColIndex("TotalStill"))
                    .TextMatrix(.Rows - 1, .ColIndex("TotalStill")) = SngTotal
          txttotal1 = SngTotal + val(txtDaysValueIncrease)
          
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
        .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
     End With
    
    
ReLineGrid
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
        If Me.dcCustomer.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ   «”„ «·„” «Ã—!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          '  Me.dcCustomer.SetFocus
           ' SendKeys "{F4}"
            Exit Sub
        End If
   If Me.DcbIqara.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ   «”„ «·⁄„«—Â!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.DcbIqara.SetFocus
           ' SendKeys "{F4}"
            Exit Sub
        End If
   


        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then

            XPTxtID.Text = CStr(new_id("TblFiterWaiver", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
          StrSQL = "Delete From TblUnitNoInformation Where FilterNo=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From TblFiterWaiverDe Where IDFItWaiv=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
               StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords


        End If
       rs("ID").value = val(XPTxtID.Text)
           rs("ContNo").value = val(TxtContNo.Text)
             rs("ContractNo").value = (Txtorder.Text)
             
       rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
       rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)
       rs("RenterID").value = IIf(Me.dcCustomer.BoundText = "", Null, Me.dcCustomer.BoundText)
       rs("BulidID").value = IIf(Me.DcbIqara.BoundText = "", Null, Me.DcbIqara.BoundText)
       rs("unittype").value = IIf(Me.DcbUnitType.BoundText = "", Null, Me.DcbUnitType.BoundText)
       rs("ApartmentID").value = IIf(Me.DcbUnitNo.BoundText = "", Null, Me.DcbUnitNo.BoundText)
       rs("RecordDate").value = XPDtbTrans.value
       rs("RecordDateH").value = Me.NourHijriCal1.value
       rs("Insurance").value = val(Me.TxtInsurance.Text)
       rs("net").value = val(Me.txtnet.Text)
       rs("ForRenter").value = val(Me.TxtForRenter.Text)
       rs("OFRenter").value = val(Me.TxtOFRenter.Text)
      ''
       rs("EndDate").value = EndDate.value
       rs("EndDateH").value = Me.EndDateH.value
       rs("FilterDate").value = FilterDate.value
       rs("FilterDateH").value = Me.FilterDateH.value
       rs("BillPrice").value = val(Me.TxtBillPrice.Text)
       rs("AccountNo").value = Me.TxtAccountNo.Text
       rs("DayNo").value = val(Me.TxtDayLate.Text)
       rs("AmountDely").value = val(Me.TxtAmountDely.Text)
       
        rs("LastInvoiceRead").value = val(Me.txtLastInvoiceRead.Text)
        rs("LastInvoiceRead2").value = val(Me.txtLastInvoiceRead2.Text)
        rs("Diff").value = val(Me.TxtDiff.Text)
        rs("Price").value = val(Me.TxtPrice.Text)
        rs("R").value = val(Me.txtR.Text)
        rs("PrevBalance").value = val(Me.txtPrevBalance.Text)
        rs("ServiceCounter").value = val(Me.txtServiceCounter.Text)
        rs("TotalCounter").value = val(Me.txtTotalCounter.Text)

   
     
            
            
    '***************************************************************************
   rs("ContractDays").value = val(Me.TxtContractDays.Text)
   rs("ActualDays").value = val(Me.TxtActualDays.Text)
rs("WaterPrice").value = val(Me.TxtWaterPrice.Text)
rs("DayPricen").value = val(Me.TxtDayPricen.Text)

rs("ServicePrice").value = val(Me.txtServicePrice.Text)
rs("WaterPriceotal").value = val(Me.TxtWaterPriceotal.Text)
rs("DayPricentotal").value = val(Me.TxtDayPricentotal.Text)
rs("Service").value = val(Me.TxtService.Text)
rs("WaterPayed").value = val(Me.txtWaterPayed.Text)
rs("RentValuePayed").value = val(Me.TxtRentValuePayed.Text)
rs("TelandNetPayed").value = val(Me.txtTelandNetPayed.Text)
rs("RemainWater").value = val(Me.txtRemainWater.Text)
rs("RemainRent").value = val(Me.txtRemainRent.Text)
rs("RemainService").value = val(Me.txtRemainService.Text)
    '***************************************************************************
   rs("outflow").value = IIf(chkoutflow.value = vbUnchecked, Null, 1)
   rs("outCondition").value = IIf(chkoutCondition.value = vbUnchecked, Null, 1)
   
   rs("DaysValueIncrease").value = val(Me.txtDaysValueIncrease.Text)
   rs("DaysValueIncomplete").value = val(Me.txtDaysValueIncomplete.Text)

    rs("DayValueInc").value = val(Me.txtDayValueInc.Text)
    rs("DayCountInc").value = val(Me.txtDayCountInc.Text)
    rs("DayValueIncomplete").value = val(Me.txtDayValueIncomplete.Text)
    rs("DayCountIncomplete").value = val(Me.txtDayCountIncomplete.Text)
    

     
        
        rs.update
        '''''''''/////////////////////////////////
        Dim temp As Integer
        temp = -1
      Set RsDetails = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblFiterWaiverDe Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

          
       For i = Me.Fg.FixedRows To Fg.Rows - 1
    
       If Fg.TextMatrix(i, Fg.ColIndex("group")) <> "" Then
   
           RsDetails.AddNew
        
           If val(Fg.TextMatrix(i, Fg.ColIndex("iditem"))) = 0 Then
           RsDetails("GroupID").value = val(Fg.TextMatrix(i, Fg.ColIndex("id")))
            RsDetails("IDItem").value = 0
             RsDetails("IDFItWaiv").value = val(XPTxtID.Text)
    
              RsDetails("Count").value = 0
           RsDetails("price").value = 0
          RsDetails("Remark").value = ""
        
          '    temp = val(fg.TextMatrix(i, fg.ColIndex("id")))
           Else
           RsDetails("IDItem").value = val(Fg.TextMatrix(i, Fg.ColIndex("iditem")))
             RsDetails("GroupID").value = 0
                  RsDetails("IDFItWaiv").value = val(XPTxtID.Text)
                           RsDetails("Count").value = val(Fg.TextMatrix(i, Fg.ColIndex("count")))
           RsDetails("price").value = val(Fg.TextMatrix(i, Fg.ColIndex("price")))
                    
                RsDetails("Remark").value = Fg.TextMatrix(i, Fg.ColIndex("remark"))
        RsDetails("Accountsus").value = Fg.TextMatrix(i, Fg.ColIndex("Accountsus"))
        
        
           End If
         
        
         RsDetails.update
      '  End If
      '   End If
      End If
        Next i

    StrSQL = "Delete From TblFiterWaiverDet2 Where MasterID=" & val(Me.XPTxtID.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    

    
    StrSQL = "Select *  From TblFiterWaiverDet2 Where MasterID=" & val(Me.XPTxtID.Text)
    
   
 
 
    
    saveGrid StrSQL, Grd, "ContNo", "id", "MasterID", val(Me.XPTxtID.Text)
 

                
        '''''''''''''''//////////////////////////
       GetUonitStatus
SaveUoitInformation
 
      

    'Dim StrSql As String
    Dim Rs7 As ADODB.Recordset
   
    
        Cn.CommitTrans
        BeginTrans = False
    '    RsDetails.Close
        Set RsDetails = Nothing
        Set RsDetails1 = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        If SystemOptions.NoCreatJLInRentContract = False Then
         createVoucher
        End If
      updateNotesValueAndNobytext (val(TxtNoteID.Text))
      Dim j As Long
      Dim mContID  As Long
            For j = 1 To Grd.Rows - 1
                mContID = val(Grd.TextMatrix(j, Grd.ColIndex("ContNo")))
                  Cn.Execute "  update TblAqarDetai  Set ContID=0 , FilterDateH='" & FilterDateH.value & "'  ,FilterDate=" & SQLDate(FilterDate.value, True) & " ,Status = 0   ,customerid=null  Where id =" & val(DcbUnitNo.BoundText)
                  Cn.Execute "  update TblContract  Set ContID=0, EndContract = 1    Where ContNo =" & val(mContID) & " and CusID=" & val(dcCustomer.BoundText) & " and UnitNo=" & val(DcbUnitNo.BoundText) & ""
                 StrSQL = " SELECT     dbo.TblIqrMerg.UntID"
                 StrSQL = StrSQL & "          FROM         dbo.TblIqrMerg INNER JOIN"
                 StrSQL = StrSQL & "          dbo.TblContract ON dbo.TblIqrMerg.Cont = dbo.TblContract.ContNo"
                 StrSQL = StrSQL & " Where (dbo.TblIqrMerg.cont = " & val(mContID) & ") And (dbo.TblContract.CusID =" & val(dcCustomer.BoundText) & ")"
                ' StrSQL = StrSQL & "  WHERE     (Cont <= " & val(TxtContNo.Text) & ") and CusID=" & val(dcCustomer.BoundText) & ""
                 Set Rs7 = New ADODB.Recordset
                 Rs7.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                 If Rs7.RecordCount > 0 Then
                 Rs7.MoveFirst
                 For i = 1 To Rs7.RecordCount
                  Cn.Execute "  update TblAqarDetai  Set ContID=0,Status = 0   ,customerid=null  Where id =" & IIf(IsNull(Rs7("UntID").value), 0, Rs7("UntID").value)
                  Rs7.MoveNext
                  Next i
                  End If
             Next j
             
        Select Case Me.TxtModFlg.Text

            Case "N"
                Msg = "  „ Õ›Ÿ »«‰«  Â–Â «·⁄„·… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ï ≈÷«›… »«‰«  √Œ—Ï"

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
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
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
            rs.Find "ID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

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
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
           Cn.Execute "  update TblAqarDetai  Set Status = 1   ,customerid=" & val(dcCustomer.BoundText) & "  Where id =" & val(DcbUnitNo.BoundText)
             Cn.Execute "  update TblContract  Set EndContract = null    Where ContNo =" & val(TxtContNo.Text)
             
             
                rs.delete
                StrSQL = "Delete From TblFiterWaiver Where ID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
                
                StrSQL = "Delete From TblFiterWaiverDet2 Where MasterID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
 StrSQL1 = "Delete From TblFiterWaiverDe Where IDFItWaiv=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL1, , adExecuteNoRecords
            StrSQL = "Delete From TblUnitNoInformation Where FilterNo=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
                  StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
        StrSQL = "Delete From NOTES Where NoteID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
                If rs.RecordCount < 1 Then
             
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
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
  Dim EmpID As Integer
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.Text, EmpID
        dcCustomer.BoundText = EmpID
    End If
End Sub

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 1215
        FrmCustemerSearch.Show vbModal

    End If
 

If KeyCode = vbKeyF5 Then
'reloadCombos
End If
End Sub
Private Sub dcCustomer_Change()
   If val(dcCustomer.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcCustomer.BoundText, EmpCode
    Me.Text15.Text = EmpCode
End Sub
'Function reloadCombos()
'Dim Dcombos As ClsDataCombos
'
' Set Dcombos = New ClsDataCombos
'Dcombos.GetCustomersSuppliers 1, Me.dcCustomer
'    Dcombos.GetCustomersSuppliers 2, Me.dcsupplier
'   Dcombos.GetIqar DcbIqara
'    Dcombos.getAkarUnit Me.DcbUnitType
'  'Dcombos.GetIqarUnit 1, DcbUnitNo
'  Dcombos.GetBranches Dcbranch
'
'  Dcombos.GetSalesRepData Me.DcboEmp
  

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
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hwnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
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


Function RtriveInfoOrbon(Optional NotID As Double = 0) As Boolean
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL As String
Dim Total As Double
RtriveInfoOrbon = True
       Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT   *  from  Notes where NoteID =" & NotID & ""
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails1.RecordCount > 0 Then
   Total = (IIf(IsNull(RsDetails1("allowdate").value), Date, RsDetails1("allowdate").value))
   If allowdate.value < ContDate.value And SystemOptions.AllowOrbonDate = False Then
   RtriveInfoOrbon = False
   Exit Function
   End If
   TxtNotSreail1.Text = val(IIf(IsNull(RsDetails1("NoteSerial1").value), "", RsDetails1("NoteSerial1").value))
 TxtNotVal.Text = val(IIf(IsNull(RsDetails1("Note_Value").value), TxtNotVal.Text, RsDetails1("Note_Value").value))
 Total = val(IIf(IsNull(RsDetails1("Note_Value").value), "", RsDetails1("Note_Value").value))
 DcbIqara.BoundText = val(IIf(IsNull(RsDetails1("akarid").value), "", RsDetails1("akarid").value))
 DcbUnitType.BoundText = val(IIf(IsNull(RsDetails1("UnitType").value), "", RsDetails1("UnitType").value))
 DcbUnitNo.BoundText = val(IIf(IsNull(RsDetails1("UnitNo").value), "", RsDetails1("UnitNo").value))
 TxtTotalContract.Text = val(IIf(IsNull(RsDetails1("rent").value) Or RsDetails1("rent").value = 0, TxtTotalContract.Text, RsDetails1("rent").value))
 TxtCommiValue.Text = val(IIf(IsNull(RsDetails1("commission").value) Or RsDetails1("commission").value = 0, TxtCommiValue.Text, RsDetails1("commission").value))
 TxtInsuranceValue.Text = val(IIf(IsNull(RsDetails1("Instrunce").value) Or RsDetails1("Instrunce").value = 0, TxtInsuranceValue.Text, RsDetails1("Instrunce").value))
 TxtWater.Text = val(IIf(IsNull(RsDetails1("Water").value) Or RsDetails1("Water").value = 0, TxtWater.Text, RsDetails1("Water").value))
 TxtElectricity.Text = val(IIf(IsNull(RsDetails1("Electricity").value) Or RsDetails1("Electricity").value = 0, TxtElectricity.Text, RsDetails1("Electricity").value))
 TxtPhone.Text = val(IIf(IsNull(RsDetails1("Servce").value) Or RsDetails1("Servce").value = 0, TxtPhone.Text, RsDetails1("Servce").value))
 TxtFATValue2.Text = val(IIf(IsNull(RsDetails1("VAT").value), 0, RsDetails1("VAT").value))

If val(TxtCommiValue.Text) <= Total Then

Me.TxtCommValue2.Text = val(Me.TxtCommiValue.Text)

Else
Me.TxtCommValue2.Text = Total
End If
Total = Total - val(TxtCommValue2.Text)
'''//////////
If val(TxtPhone.Text) <= Total Then
Me.TxtServce.Text = Me.TxtPhone.Text
Else
Me.TxtServce.Text = Total
End If
Total = Total - val(TxtServce.Text)

''////////
If val(TxtInsuranceValue.Text) <= Total Then
Me.TxtInstrunceValue2.Text = Me.TxtInsuranceValue.Text
ElseIf Total > 0 Then
Me.TxtInstrunceValue2.Text = Total
Else
Me.TxtInstrunceValue2.Text = 0
End If
Total = Total - val(TxtInstrunceValue2.Text)
''//
If val(TxtWater.Text) <= Total Then
If chkDivWater.value = vbChecked Then
Me.TxtWaterValue2.Text = Round(val(Me.TxtWater.Text) / val(TxtPaymentCount.Text), 2)
Else
Me.TxtWaterValue2.Text = Me.TxtWater.Text
End If
ElseIf Total > 0 Then
Me.TxtWaterValue2.Text = Total
Else
Me.TxtWaterValue2.Text = 0
End If
Total = Total - val(TxtWaterValue2.Text)
''//
''//
If val(TxtElectricity.Text) <= Total Then
If chkDivElectric.value = vbChecked Then
Me.TxtElectricityValue2.Text = Round(val(Me.TxtElectricity.Text) / val(TxtPaymentCount.Text), 2)
Else
Me.TxtElectricityValue2.Text = Me.TxtElectricity.Text
End If
ElseIf Total > 0 Then
Me.TxtElectricityValue2.Text = Total
Else
Me.TxtElectricityValue2.Text = 0
End If
''//
Total = Total - val(TxtElectricityValue2.Text)
If val(TxtTotalContract.Text) <= Total Then
Me.TxtRetValue2.Text = Me.TxtTotalContract.Text
ElseIf Total > 0 Then
Me.TxtRetValue2.Text = Total
Else
Me.TxtRetValue2.Text = 0
End If


   End If
End Function

Function CheckJE() As Boolean
Dim i As Integer
CheckJE = False
With GridInstallments
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("NoteId"))) <> 0 Then
CheckJE = True
Exit Function
End If
Next i
End With
End Function
