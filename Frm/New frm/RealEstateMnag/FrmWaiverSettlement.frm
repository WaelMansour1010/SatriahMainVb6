VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmWaiverSettlement 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18615
   FillColor       =   &H00C0E0FF&
   Icon            =   "FrmWaiverSettlement.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   18615
   Begin VB.TextBox txtLegalAffairs 
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
      Left            =   150
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   271
      Top             =   510
      Width           =   2985
   End
   Begin VB.CheckBox chkIsLegalAffairs 
      Alignment       =   1  'Right Justify
      Caption         =   "«Õ«·… ··‘∆Ê‰ «·ﬁ«‰Ê‰Ì…"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4110
      RightToLeft     =   -1  'True
      TabIndex        =   270
      Top             =   540
      Width           =   2055
   End
   Begin VB.TextBox txtManualNo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   14970
      RightToLeft     =   -1  'True
      TabIndex        =   268
      Top             =   1260
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CheckBox ChkAccredit 
      Alignment       =   1  'Right Justify
      Caption         =   " „ «· ÊÀÌﬁ"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2190
      RightToLeft     =   -1  'True
      TabIndex        =   267
      Top             =   1230
      Width           =   1455
   End
   Begin VB.TextBox TxtElectAccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   3780
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   265
      Top             =   1200
      Width           =   2025
   End
   Begin VB.TextBox TXTNewNO 
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
      Left            =   7470
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   263
      Top             =   1260
      Width           =   2985
   End
   Begin VB.CommandButton CMDSENDSMS 
      Caption         =   "«—”«· —”«·Â"
      Height          =   375
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   9630
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
      Height          =   375
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   9450
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
      TabIndex        =   58
      Top             =   2160
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox TxtDayPrice 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   12390
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   1830
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox TxtOrder 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6780
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   510
      Width           =   1515
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   16110
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   510
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
      Left            =   12990
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   9450
      Width           =   1575
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
      Height          =   495
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   18615
      _cx             =   32835
      _cy             =   873
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
         TabIndex        =   64
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
         TabIndex        =   63
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox TxtContNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   60
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
         ButtonImage     =   "FrmWaiverSettlement.frx":038A
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
         ButtonImage     =   "FrmWaiverSettlement.frx":0724
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
         ButtonImage     =   "FrmWaiverSettlement.frx":0ABE
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
         ButtonImage     =   "FrmWaiverSettlement.frx":0E58
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
         Picture         =   "FrmWaiverSettlement.frx":11F2
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
      Left            =   13920
      TabIndex        =   6
      Top             =   510
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   200146945
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   3690
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9450
      Width           =   8175
      _cx             =   14420
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
      Left            =   15390
      TabIndex        =   15
      Top             =   9480
      Width           =   2280
      _ExtentX        =   4022
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
      Bindings        =   "FrmWaiverSettlement.frx":4E5A
      Height          =   315
      Left            =   8970
      TabIndex        =   29
      Top             =   510
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
      Top             =   1650
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
      Picture(0)      =   "FrmWaiverSettlement.frx":4E6F
      Flags(1)        =   2
      Begin VB.Frame LblWork 
         BackColor       =   &H00E2E9E9&
         Height          =   7350
         Left            =   19605
         RightToLeft     =   -1  'True
         TabIndex        =   117
         Top             =   45
         Width           =   18570
         Begin VSFlex8Ctl.VSFlexGrid fg 
            Height          =   6300
            Left            =   120
            TabIndex        =   118
            Top             =   240
            Width           =   17985
            _cx             =   31724
            _cy             =   11112
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
            FormatString    =   $"FrmWaiverSettlement.frx":5209
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
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   240
            Left            =   13545
            TabIndex        =   228
            Top             =   6780
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   423
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› ”ÿ—"
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
            ButtonImage     =   "FrmWaiverSettlement.frx":5344
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   240
            Left            =   11880
            TabIndex        =   229
            Top             =   6780
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   423
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› «·ﬂ·"
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
            ButtonImage     =   "FrmWaiverSettlement.frx":58DE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
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
            TabIndex        =   120
            Top             =   6870
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
            TabIndex        =   119
            Top             =   6870
            Width           =   9570
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7350
         Left            =   19305
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
            TabIndex        =   126
            Top             =   0
            Width           =   11775
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   146
               Top             =   960
               Width           =   2355
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   145
               Top             =   600
               Width           =   2355
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   144
               Top             =   960
               Width           =   2595
            End
            Begin VB.TextBox TxtAmountDely 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   143
               Top             =   1680
               Width           =   2355
            End
            Begin VB.TextBox TxtDayLate 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   142
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
               TabIndex        =   141
               Top             =   240
               Width           =   825
            End
            Begin VB.TextBox TxtDayPricen 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   140
               Top             =   2040
               Width           =   2955
            End
            Begin VB.TextBox TxtWaterPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   2040
               Width           =   2595
            End
            Begin VB.TextBox TxtActualDays 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   1320
               Width           =   2355
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   137
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
               TabIndex        =   136
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
               TabIndex        =   135
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
               TabIndex        =   134
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
               TabIndex        =   133
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
               TabIndex        =   132
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
               TabIndex        =   131
               Top             =   2760
               Width           =   2355
            End
            Begin VB.TextBox txtServicePrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   130
               Top             =   2040
               Width           =   2355
            End
            Begin VB.TextBox txtRemainService 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   3120
               Width           =   2355
            End
            Begin VB.TextBox txtRemainWater 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   3120
               Width           =   2595
            End
            Begin VB.TextBox txtRemainRent 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   3120
               Width           =   2955
            End
            Begin MSComCtl2.DTPicker EndDate 
               Height          =   315
               Left            =   8880
               TabIndex        =   147
               Top             =   1320
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   249757697
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   8880
               TabIndex        =   148
               Top             =   1680
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Format          =   249757697
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal EndDateH 
               Height          =   315
               Left            =   7380
               TabIndex        =   149
               Top             =   1320
               Width           =   1455
               _ExtentX        =   2778
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal NourHijriCal2 
               Height          =   315
               Left            =   7380
               TabIndex        =   150
               Top             =   1680
               Width           =   1455
               _ExtentX        =   2778
               _ExtentY        =   556
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   315
               Left            =   120
               TabIndex        =   151
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
               TabIndex        =   152
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
               TabIndex        =   153
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
               TabIndex        =   154
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
               TabIndex        =   155
               Top             =   960
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   249823233
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal StartDateh 
               Height          =   315
               Left            =   7380
               TabIndex        =   156
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
               TabIndex        =   183
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
               TabIndex        =   182
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
               TabIndex        =   181
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
               TabIndex        =   180
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
               TabIndex        =   179
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
               TabIndex        =   178
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
               TabIndex        =   177
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
               TabIndex        =   176
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
               TabIndex        =   175
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
               TabIndex        =   174
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
               TabIndex        =   173
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
               TabIndex        =   172
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
               TabIndex        =   171
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
               TabIndex        =   170
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
               TabIndex        =   169
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
               TabIndex        =   168
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
               TabIndex        =   167
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
               TabIndex        =   166
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
               TabIndex        =   165
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
               TabIndex        =   164
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
               TabIndex        =   163
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
               TabIndex        =   162
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
               TabIndex        =   161
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
               TabIndex        =   160
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
               TabIndex        =   159
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
               TabIndex        =   158
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
               TabIndex        =   157
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
            TabIndex        =   101
            Top             =   705
            Width           =   2355
         End
         Begin VB.TextBox TxtBillPrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   3810
            RightToLeft     =   -1  'True
            TabIndex        =   100
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
            FormatString    =   $"FrmWaiverSettlement.frx":5E78
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
            TabIndex        =   66
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
               ItemData        =   "FrmWaiverSettlement.frx":5FC4
               Left            =   7575
               List            =   "FrmWaiverSettlement.frx":5FD1
               RightToLeft     =   -1  'True
               TabIndex        =   74
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
               TabIndex        =   73
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
               TabIndex        =   72
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
               TabIndex        =   71
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
               TabIndex        =   70
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
               TabIndex        =   69
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
               TabIndex        =   68
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
               TabIndex        =   67
               Top             =   1125
               Width           =   1140
            End
            Begin MSComCtl2.DTPicker FristPaymentDate 
               Height          =   345
               Left            =   4710
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   255
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               Format          =   247267331
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal FirstInstallDateH 
               Height          =   285
               Left            =   6210
               TabIndex        =   76
               Top             =   255
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   503
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   465
               Index           =   20
               Left            =   495
               TabIndex        =   77
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
               ButtonImage     =   "FrmWaiverSettlement.frx":5FE4
               DrawFocusRectangle=   0   'False
            End
            Begin C1SizerLibCtl.C1Tab TabMain 
               Height          =   5115
               Left            =   60
               TabIndex        =   78
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
                  TabIndex        =   79
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
                     TabIndex        =   80
                     Top             =   3630
                     Width           =   2055
                  End
                  Begin VSFlex8UCtl.VSFlexGrid FgItems 
                     Height          =   4740
                     Index           =   1
                     Left            =   13095
                     TabIndex        =   81
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
                     FormatString    =   $"FrmWaiverSettlement.frx":637E
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
                     TabIndex        =   82
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
                     FormatString    =   $"FrmWaiverSettlement.frx":643E
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
                     TabIndex        =   86
                     Top             =   3915
                     Width           =   1455
                  End
                  Begin VB.Label LblNotPayed 
                     Alignment       =   2  'Center
                     Caption         =   "0"
                     Height          =   255
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
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
                     TabIndex        =   84
                     Top             =   3915
                     Width           =   1980
                  End
                  Begin VB.Label LblTotalQasts 
                     Alignment       =   2  'Center
                     Caption         =   "0"
                     Height          =   465
                     Left            =   4860
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   3765
                     Width           =   1650
                  End
               End
               Begin C1SizerLibCtl.C1Elastic ELe 
                  Height          =   4740
                  Index           =   11
                  Left            =   12015
                  TabIndex        =   87
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
                     TabIndex        =   88
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
                     FormatString    =   $"FrmWaiverSettlement.frx":6DD7
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
                     TabIndex        =   92
                     Top             =   3540
                     Width           =   1455
                  End
                  Begin VB.Label LblNotPayed2 
                     Alignment       =   2  'Center
                     Caption         =   "0"
                     Height          =   990
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   91
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
                     TabIndex        =   90
                     Top             =   3540
                     Width           =   1950
                  End
                  Begin VB.Label LblTotalQasts2 
                     Alignment       =   2  'Center
                     Caption         =   "0"
                     Height          =   990
                     Left            =   4860
                     RightToLeft     =   -1  'True
                     TabIndex        =   89
                     Top             =   3540
                     Width           =   1650
                  End
               End
               Begin C1SizerLibCtl.C1Elastic ELe 
                  Height          =   4740
                  Index           =   13
                  Left            =   12315
                  TabIndex        =   93
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
                     TabIndex        =   94
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
                     FormatString    =   $"FrmWaiverSettlement.frx":7742
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
               TabIndex        =   98
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
               TabIndex        =   97
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
               TabIndex        =   96
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
               TabIndex        =   95
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
            TabIndex        =   103
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
            TabIndex        =   102
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
            TabIndex        =   48
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
         Left            =   45
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
         _GridInfo       =   $"FrmWaiverSettlement.frx":77E1
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
            Begin VB.TextBox TxtForRenterB 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Enabled         =   0   'False
               Height          =   360
               Left            =   12990
               RightToLeft     =   -1  'True
               TabIndex        =   275
               Top             =   5460
               Width           =   1950
            End
            Begin VB.TextBox txtDiscount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00808080&
               Height          =   345
               Left            =   12990
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   273
               Top             =   5880
               Width           =   1965
            End
            Begin VB.CommandButton Command2 
               Caption         =   "› Õ ”‰œ «·ﬁ»÷"
               Height          =   315
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   262
               Top             =   5400
               Width           =   2295
            End
            Begin VB.CommandButton Command1 
               Caption         =   "› Õ ”‰œ «·’—›"
               Height          =   315
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   261
               Top             =   5400
               Width           =   2295
            End
            Begin VB.Frame Frame5 
               Enabled         =   0   'False
               Height          =   495
               Left            =   7530
               RightToLeft     =   -1  'True
               TabIndex        =   250
               Top             =   30
               Width           =   2145
               Begin VB.OptionButton RdRTypeDate 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„Ì·«œÌ"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   1
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   252
                  Top             =   150
                  Width           =   855
               End
               Begin VB.OptionButton RdRTypeDate 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ÂÃ—Ì"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   0
                  Left            =   1170
                  RightToLeft     =   -1  'True
                  TabIndex        =   251
                  Top             =   150
                  Width           =   735
               End
            End
            Begin VB.Frame Frame4 
               Enabled         =   0   'False
               Height          =   495
               Left            =   9750
               RightToLeft     =   -1  'True
               TabIndex        =   247
               Top             =   30
               Width           =   2355
               Begin VB.OptionButton ComResid 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”ﬂ‰Ì"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   0
                  Left            =   1110
                  RightToLeft     =   -1  'True
                  TabIndex        =   249
                  Top             =   180
                  Width           =   975
               End
               Begin VB.OptionButton ComResid 
                  Alignment       =   1  'Right Justify
                  Caption         =   " Ã«—Ì"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   1
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   248
                  Top             =   180
                  Width           =   975
               End
            End
            Begin VB.CheckBox ChkCalcLastPayment 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Õ ”«» »‰«¡« ⁄·Ï «Œ— œ›⁄… "
               Height          =   255
               Left            =   9720
               RightToLeft     =   -1  'True
               TabIndex        =   246
               Top             =   480
               Width           =   2535
            End
            Begin VB.CheckBox chkTypeMonthCalc 
               Alignment       =   1  'Right Justify
               Caption         =   "«Õ ”«» «·‘Â— 30 ÌÊ„ Õ«·… «·ÂÃ—Ì"
               Height          =   255
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   236
               Top             =   900
               Width           =   2805
            End
            Begin VB.TextBox txtTotalinsuranceS 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000010&
               Enabled         =   0   'False
               Height          =   345
               Left            =   3000
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   234
               Top             =   6960
               Width           =   1380
            End
            Begin VB.TextBox txtOldInsurance 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000010&
               Enabled         =   0   'False
               Height          =   345
               Left            =   3030
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   232
               Top             =   6390
               Width           =   1380
            End
            Begin VB.TextBox txtTotalLastDays 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000010&
               Enabled         =   0   'False
               Height          =   345
               Left            =   3030
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   231
               Top             =   5790
               Width           =   1380
            End
            Begin VB.TextBox TxtForRenter 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Enabled         =   0   'False
               Height          =   360
               Left            =   12990
               RightToLeft     =   -1  'True
               TabIndex        =   218
               Top             =   6180
               Width           =   1950
            End
            Begin VB.TextBox TxtOFRenter 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Enabled         =   0   'False
               Height          =   360
               Left            =   12990
               RightToLeft     =   -1  'True
               TabIndex        =   217
               Top             =   6540
               Width           =   1950
            End
            Begin VB.TextBox TxtNet 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   12990
               RightToLeft     =   -1  'True
               TabIndex        =   216
               Top             =   6930
               Width           =   1950
            End
            Begin VB.Frame Frame1 
               Caption         =   "«Ì«„ “Ì«œ…"
               Height          =   1185
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   209
               Top             =   4680
               Width           =   2925
               Begin VB.TextBox txtDayValueInc 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Height          =   345
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   212
                  Top             =   120
                  Width           =   1380
               End
               Begin VB.TextBox txtDayCountInc 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Height          =   345
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   211
                  Top             =   450
                  Width           =   1380
               End
               Begin VB.TextBox txtDaysValueIncrease 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Height          =   345
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   210
                  Top             =   780
                  Width           =   1380
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·ÌÊ„"
                  Height          =   255
                  Index           =   57
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   215
                  Top             =   180
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «Ì«„"
                  Height          =   255
                  Index           =   52
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   214
                  Top             =   540
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·«Ì«„ “Ì«œ…"
                  Height          =   255
                  Index           =   54
                  Left            =   1350
                  RightToLeft     =   -1  'True
                  TabIndex        =   213
                  Top             =   900
                  Width           =   1425
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "«Ì«„ ‰«ﬁ’…"
               Height          =   1245
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   202
               Top             =   6030
               Width           =   2895
               Begin VB.TextBox txtDayValueIncomplete 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000B&
                  Height          =   345
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   205
                  Top             =   150
                  Width           =   1380
               End
               Begin VB.TextBox txtDayCountIncomplete 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000B&
                  Height          =   345
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   204
                  Top             =   510
                  Width           =   1380
               End
               Begin VB.TextBox txtDaysValueIncomplete 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000B&
                  Height          =   345
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   203
                  Top             =   870
                  Width           =   1380
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·ÌÊ„"
                  Height          =   255
                  Index           =   53
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   208
                  Top             =   240
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «Ì«„"
                  Height          =   255
                  Index           =   58
                  Left            =   1860
                  RightToLeft     =   -1  'True
                  TabIndex        =   207
                  Top             =   570
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·«Ì«„ «·‰«ﬁ’…"
                  Height          =   255
                  Index           =   56
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   206
                  Top             =   900
                  Width           =   1335
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "«·ﬂÂ—»«¡"
               Height          =   975
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   185
               Top             =   4350
               Width           =   15375
               Begin VB.TextBox txtTotalCounterNet 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00808080&
                  Height          =   345
                  Left            =   60
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   241
                  Top             =   540
                  Width           =   1575
               End
               Begin VB.TextBox TxtVAt2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00808080&
                  Height          =   345
                  Left            =   1680
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   240
                  Top             =   540
                  Width           =   1245
               End
               Begin VB.TextBox TxtVAtPercent 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2940
                  RightToLeft     =   -1  'True
                  TabIndex        =   238
                  Text            =   "5"
                  Top             =   540
                  Width           =   375
               End
               Begin VB.TextBox txtLastInvoiceRead 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   13890
                  RightToLeft     =   -1  'True
                  TabIndex        =   193
                  Top             =   540
                  Width           =   945
               End
               Begin VB.TextBox txtLastInvoiceRead2 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   12060
                  RightToLeft     =   -1  'True
                  TabIndex        =   192
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtDiff 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00808080&
                  Height          =   345
                  Left            =   10560
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   191
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   9390
                  RightToLeft     =   -1  'True
                  TabIndex        =   190
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtR 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00808080&
                  Height          =   345
                  Left            =   8190
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   189
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtPrevBalance 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   6960
                  RightToLeft     =   -1  'True
                  TabIndex        =   188
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtServiceCounter 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   5700
                  RightToLeft     =   -1  'True
                  TabIndex        =   187
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.TextBox txtTotalCounter 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00808080&
                  Height          =   345
                  Left            =   4500
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   186
                  Top             =   540
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·’«›Ì »⁄œ ﬁ.„"
                  Height          =   375
                  Index           =   75
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   242
                  Top             =   210
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬁÌ„… «·„÷«›…"
                  Height          =   375
                  Index           =   74
                  Left            =   1740
                  RightToLeft     =   -1  'True
                  TabIndex        =   239
                  Top             =   180
                  Width           =   1065
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁ.„"
                  Height          =   375
                  Index           =   73
                  Left            =   2490
                  RightToLeft     =   -1  'True
                  TabIndex        =   237
                  Top             =   180
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄œ«œ ⁄‰œ Œ—ÊÃ «·„” √Ã—"
                  Height          =   555
                  Index           =   60
                  Left            =   11490
                  RightToLeft     =   -1  'True
                  TabIndex        =   201
                  Top             =   240
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄œ«œ ›Ì «Œ— ›« Ê—Â"
                  Height          =   405
                  Index           =   59
                  Left            =   13290
                  RightToLeft     =   -1  'True
                  TabIndex        =   200
                  Top             =   240
                  Width           =   1785
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·›—ﬁ"
                  Height          =   255
                  Index           =   61
                  Left            =   10740
                  RightToLeft     =   -1  'True
                  TabIndex        =   199
                  Top             =   240
                  Width           =   555
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄— «·ÊÕœ…"
                  Height          =   375
                  Index           =   63
                  Left            =   9300
                  RightToLeft     =   -1  'True
                  TabIndex        =   198
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   285
                  Index           =   64
                  Left            =   8430
                  RightToLeft     =   -1  'True
                  TabIndex        =   197
                  Top             =   240
                  Width           =   585
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—’Ìœ ”«»ﬁ"
                  Height          =   375
                  Index           =   65
                  Left            =   6690
                  RightToLeft     =   -1  'True
                  TabIndex        =   196
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œœ„… «·⁄œ«œ"
                  Height          =   375
                  Index           =   66
                  Left            =   5700
                  RightToLeft     =   -1  'True
                  TabIndex        =   195
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   375
                  Index           =   70
                  Left            =   4350
                  RightToLeft     =   -1  'True
                  TabIndex        =   194
                  Top             =   240
                  Width           =   1005
               End
            End
            Begin VB.TextBox txtTotal2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   345
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   4440
               Visible         =   0   'False
               Width           =   1890
            End
            Begin VB.TextBox txtTotal1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   345
               Left            =   3780
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   5070
               Visible         =   0   'False
               Width           =   1380
            End
            Begin VB.TextBox TxtAccountNo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   300
               RightToLeft     =   -1  'True
               TabIndex        =   106
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
               Left            =   5385
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   90
               Width           =   855
            End
            Begin VB.TextBox TxtContractDays 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   3555
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   810
               Visible         =   0   'False
               Width           =   2685
            End
            Begin VSFlex8Ctl.VSFlexGrid grd 
               Height          =   3150
               Left            =   90
               TabIndex        =   99
               Top             =   1200
               Width           =   18165
               _cx             =   32041
               _cy             =   5556
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
               Cols            =   26
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmWaiverSettlement.frx":7817
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
               Left            =   2130
               TabIndex        =   107
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
               Left            =   12195
               TabIndex        =   108
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
               Top             =   90
               Width           =   4170
               _ExtentX        =   7355
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbUnitNo 
               Height          =   315
               Left            =   3555
               TabIndex        =   109
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
               TabIndex        =   110
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
               TabIndex        =   121
               Top             =   780
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               Format          =   205717505
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal FilterDateH 
               Height          =   315
               Left            =   13575
               TabIndex        =   122
               Top             =   780
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
            End
            Begin MSComCtl2.DTPicker txtLastInstalldate 
               Height          =   315
               Left            =   9960
               TabIndex        =   243
               Top             =   885
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               Format          =   205717505
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal txtInstalldateH 
               Height          =   315
               Left            =   8550
               TabIndex        =   244
               Top             =   885
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„»·€ «·„” Õﬁ  ⁄·Ï «·„” √Ã— »⁄œ «· ’›ÌÂ —ﬁ„«"
               Height          =   300
               Index           =   81
               Left            =   14055
               RightToLeft     =   -1  'True
               TabIndex        =   276
               Top             =   5460
               Width           =   4380
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·Œ’„"
               Height          =   285
               Index           =   80
               Left            =   17400
               RightToLeft     =   -1  'True
               TabIndex        =   274
               Top             =   5850
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ ¬Œ— œ›⁄…"
               Height          =   375
               Index           =   76
               Left            =   11310
               RightToLeft     =   -1  'True
               TabIndex        =   245
               Top             =   900
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì «· √„Ì‰"
               Height          =   285
               Index           =   72
               Left            =   2940
               RightToLeft     =   -1  'True
               TabIndex        =   235
               Top             =   6720
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " √„Ì‰ ”«»ﬁ"
               Height          =   375
               Index           =   71
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   233
               Top             =   6120
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ «Ì«„ «·”ﬂ‰ ·«Œ— ⁄ﬁœ"
               Height          =   465
               Index           =   62
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   230
               Top             =   5370
               Width           =   1275
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„»·€ «·„” Õﬁ  ⁄·Ï «·„” √Ã— »⁄œ «· ’›ÌÂ —ﬁ„«"
               Height          =   300
               Index           =   10
               Left            =   14055
               RightToLeft     =   -1  'True
               TabIndex        =   227
               Top             =   6180
               Width           =   4380
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„»·€ «·„” Õﬁ  ··„” √Ã— »⁄œ «· ’›ÌÂ —ﬁ„«"
               Height          =   300
               Index           =   2
               Left            =   14520
               RightToLeft     =   -1  'True
               TabIndex        =   226
               Top             =   6540
               Width           =   3915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂ «»…"
               ForeColor       =   &H8000000D&
               Height          =   300
               Index           =   3
               Left            =   11070
               RightToLeft     =   -1  'True
               TabIndex        =   225
               Top             =   6180
               Width           =   1875
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂ «»…"
               ForeColor       =   &H8000000D&
               Height          =   300
               Index           =   5
               Left            =   11010
               RightToLeft     =   -1  'True
               TabIndex        =   224
               Top             =   6570
               Width           =   1935
            End
            Begin VB.Label lbll 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   300
               Index           =   9
               Left            =   4710
               RightToLeft     =   -1  'True
               TabIndex        =   223
               Top             =   6180
               Width           =   7860
            End
            Begin VB.Label lbll 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   285
               Index           =   11
               Left            =   4410
               RightToLeft     =   -1  'True
               TabIndex        =   222
               Top             =   6630
               Width           =   8160
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "’«›Ì «·Õ”«»"
               Height          =   300
               Index           =   9
               Left            =   16395
               RightToLeft     =   -1  'True
               TabIndex        =   221
               Top             =   6930
               Width           =   2040
            End
            Begin VB.Label lbll 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   210
               Index           =   0
               Left            =   4410
               RightToLeft     =   -1  'True
               TabIndex        =   220
               Top             =   6990
               Width           =   8160
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂ «»…"
               ForeColor       =   &H8000000D&
               Height          =   315
               Index           =   11
               Left            =   11040
               RightToLeft     =   -1  'True
               TabIndex        =   219
               Top             =   6930
               Width           =   1905
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «· ’›Ì…"
               Height          =   375
               Index           =   20
               Left            =   16530
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ÊÕœÂ"
               Height          =   195
               Index           =   15
               Left            =   6330
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   510
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄ﬁ«—"
               Height          =   255
               Index           =   13
               Left            =   6390
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   90
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·Õ”«»"
               Height          =   255
               Index           =   19
               Left            =   2670
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   480
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
               TabIndex        =   113
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
               TabIndex        =   112
               Top             =   450
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄ﬁœ ·„œ…"
               Height          =   375
               Index           =   32
               Left            =   6330
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   810
               Visible         =   0   'False
               Width           =   1005
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
               TabIndex        =   52
               Top             =   2985
               Width           =   7035
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   7320
            Index           =   9
            Left            =   15
            TabIndex        =   42
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
               TabIndex        =   44
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
               TabIndex        =   43
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
               TabIndex        =   47
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
               TabIndex        =   46
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
               TabIndex        =   45
               Top             =   2010
               Width           =   465
            End
         End
      End
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   0
      TabIndex        =   51
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
      Left            =   12270
      TabIndex        =   53
      Top             =   510
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
   End
   Begin MSComCtl2.DTPicker From 
      Height          =   315
      Left            =   12360
      TabIndex        =   57
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   243466241
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   375
      Left            =   6300
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
      Top             =   510
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
      ButtonImage     =   "FrmWaiverSettlement.frx":7C38
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
   End
   Begin MSDataListLib.DataCombo dcCustomer2 
      Height          =   315
      Left            =   1020
      TabIndex        =   253
      Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
      Top             =   870
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbUnitNo2 
      Height          =   315
      Left            =   8250
      TabIndex        =   255
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   870
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbIqara2 
      Height          =   315
      Left            =   13890
      TabIndex        =   257
      Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
      Top             =   870
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbUnitType2 
      Height          =   315
      Left            =   11160
      TabIndex        =   259
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   870
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "”»» «·«Õ«·…"
      Height          =   285
      Index           =   4
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   272
      Top             =   540
      Width           =   810
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «· ’›Ì… «·ÌœÊÌ"
      Height          =   255
      Index           =   79
      Left            =   17220
      RightToLeft     =   -1  'True
      TabIndex        =   269
      Top             =   1290
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ Õ”«» «·ﬂÂ—»«¡"
      Height          =   435
      Index           =   48
      Left            =   5700
      RightToLeft     =   -1  'True
      TabIndex        =   266
      Top             =   1230
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·—ﬁ„ «·„ÊÕœ"
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   79
      Left            =   10470
      RightToLeft     =   -1  'True
      TabIndex        =   264
      Top             =   1260
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «·ÊÕœ…"
      Height          =   195
      Index           =   3
      Left            =   12660
      RightToLeft     =   -1  'True
      TabIndex        =   260
      Top             =   930
      Width           =   1080
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄ﬁ«—"
      Height          =   255
      Index           =   78
      Left            =   17700
      RightToLeft     =   -1  'True
      TabIndex        =   258
      Top             =   900
      Width           =   810
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·ÊÕœÂ"
      Height          =   195
      Index           =   77
      Left            =   10140
      RightToLeft     =   -1  'True
      TabIndex        =   256
      Top             =   930
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "  ’›Ì… ⁄ﬁÊœ „” √Ã—"
      Height          =   285
      Index           =   2
      Left            =   6090
      RightToLeft     =   -1  'True
      TabIndex        =   254
      Top             =   885
      Width           =   1470
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«Ì«„ “Ì«œ…"
      Height          =   255
      Index           =   55
      Left            =   1830
      RightToLeft     =   -1  'True
      TabIndex        =   184
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
      Left            =   14370
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   9480
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "⁄ﬁœ —ﬁ„"
      Height          =   255
      Index           =   14
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   510
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄„«—Â"
      Height          =   255
      Left            =   3660
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   0
      Width           =   855
   End
   Begin VB.Image img 
      Height          =   855
      Left            =   22680
      Picture         =   "FrmWaiverSettlement.frx":8035
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
      Picture         =   "FrmWaiverSettlement.frx":9059
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   720
   End
   Begin VB.Label lblBr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Left            =   11130
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   570
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
      Top             =   510
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   14940
      TabIndex        =   23
      Top             =   510
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   17775
      TabIndex        =   22
      Top             =   9555
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2400
      TabIndex        =   21
      Top             =   9540
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   450
      TabIndex        =   20
      Top             =   9540
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   -180
      TabIndex        =   19
      Top             =   9210
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1500
      TabIndex        =   18
      Top             =   9570
      Width           =   645
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
Attribute VB_Name = "FrmWaiverSettlement"
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

 



Private Sub ChkCalcLastPayment_Click()
    If Me.TxtModFlg.text <> "R" Then


        Dim IntMintsCount As Integer
        RetriveOrder
        GetContract val(TxtOrder), val(TxtContNo)
    End If
End Sub

Private Sub chkTypeMonthCalc_Click()
If Me.TxtModFlg.text <> "R" Then
    RetriveOrder
    GetContract val(TxtOrder), val(TxtContNo)
End If
End Sub

Private Sub Cmd_DeleteAll_Click()
If Me.TxtModFlg.text <> "R" Then


 FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 2

End If
End Sub
Private Sub RemoveGridRow()

    With Me.FG
'MsgBox .Row
        If .Row <= 0 Then
                .rows = 2
        Exit Sub
        Else
        .RemoveItem .Row
        End If
    End With
End Sub
Private Sub Cmd_DeleteRow_Click()
If Me.TxtModFlg.text <> "R" Then

RemoveGridRow

End If
End Sub

Private Sub Command1_Click()
FrmPayments2.show
FrmPayments2.Cmd_Click (0)
FrmPayments2.DCboCashType.ListIndex = 6
                FrmPayments2.TxtFilterNo.text = val(TxtOrder.text)
                FrmPayments2.TXtFilter.text = val(XPTxtID.text)
                FrmPayments2.GetWonerID (val(DcbIqara.BoundText))
                FrmPayments2.DBCboClientName.BoundText = val(dcCustomer.BoundText)
                ''//
                FrmPayments2.DcbIqara.BoundText = val(DcbIqara.BoundText)
                FrmPayments2.DcbUnitType.BoundText = val(DcbUnitType.BoundText)
                FrmPayments2.DcbUnitNo.BoundText = val(DcbUnitNo.BoundText)
                
                FrmPayments2.XPTxtVal.text = Abs(val(txtNet.text))
            
End Sub

Private Sub Command2_Click()
FrmCashing1.show
FrmCashing1.Cmd_Click (0)
FrmCashing1.DCboCashType.ListIndex = 10
  
                    FrmCashing1.DcbIqara.BoundText = (val(DcbIqara.BoundText))
                    FrmCashing1.TxtFilterNo.text = val(Me.XPTxtID)
                    
              '      FrmCashing1.TXtFilter.Text = val(fg.TextMatrix(fg.Row, fg.ColIndex("net")))
                    FrmCashing1.DBCboClientName.BoundText = val(dcCustomer.BoundText)
                    FrmCashing1.DcbUnitNo.text = val(DcbUnitNo.BoundText)
                    FrmCashing1.Dtaefilter.value = XPDtbTrans.value
                    FrmCashing1.XPTxtVal.text = Abs(val(txtNet.text))
                    FrmCashing1.txtTotalinsuranceS.text = Abs(val(txtTotalinsuranceS.text))
                    FrmCashing1.GetWonerID ((val(DcbIqara.BoundText)))
                    
                     FrmCashing1.DcbUnitType.BoundText = val(DcbUnitType.BoundText)
                     
                    
                    Dim str As String
                    str = " ’›Ì…  ··ÊÕœ… " & "   " & FrmCashing1.DcbUnitType.text & " -" & FrmCashing1.DcbUnitNo.text & " ··⁄ﬁ«— " & FrmCashing1.DcbIqara.text
                 '   str = str & " ··ÊÕœ… " & FrmCashing1.DcbUnitType.Text
                    
                    FrmCashing1.txtRemarks.text = str
                    FrmCashing1.XPMTxtRemarks.text = str
FrmCashing1.TxtFilterNo_Change

End Sub

Private Sub DcbUnitNo2_Click(Area As Integer)
    If DcbUnitNo2.text <> "" Then TxtOrder = ""
     RetriveOrder
     GetContract val(TxtOrder), val(TxtContNo)
End Sub

Private Sub dcCustomer2_Click(Area As Integer)
    If dcCustomer2.text <> "" Then TxtOrder = ""
     RetriveOrder
     GetContract val(TxtOrder), val(TxtContNo)
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With FG

        Select Case .ColKey(Col)
 
            Case "group"
             .TextMatrix(Row, .ColIndex("group")) = ""
                StrSQL = "select * from TblAqrCompenetDet"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = .BuildComboList(rs, "Name", "ID")
               'If SystemOptions.UserInterface = ArabicInterface Then
               '    StrComboList = .BuildComboList(rs, "Name", "ID")
                'lse
                    'StrComboList = .BuildComboList(rs, "Emp_Namee", "Emp_ID")
               'End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            End Select
        End With
End Sub

Private Sub FirstInstallDateH_LostFocus()
        
        If Me.TxtModFlg.text <> "R" Then
              VBA.Calendar = vbCalGreg
            FristPaymentDate.value = ToGregorianDate(FirstInstallDateH.value)
               
        End If

End Sub

Private Sub FristPaymentDate_Change()
If Me.TxtModFlg.text <> "R" Then
     
         FirstInstallDateH.value = ToHijriDate(FristPaymentDate.value)
       
End If
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
           Dim opt As Integer
            Dim CurrentMessage As String
            Dim t As String
    CurrentMessage = ComposMessage(Me.Name, 0, "", "", opt)
  If opt = currentOpt Then
  
      CompanyName = cOptions.ArabCompanyName '& CHR(13) & CurrentBranchName
     companyphone = cOptions.Company_Mobile
  '«·„” √Ã—
 Msg = "  „ ⁄„·    ’›Ì…  " & "  ··ÊÕœ… —ﬁ„   " & DcbUnitNo.text & "    ··⁄ﬁ«— —ﬁ„ " & CHR(13) & DcbIqara.text & " √„·Ì‰ —÷«ﬂ„ "
t = sendMessageM("user", "password", Msg, "", GetCustomerNumber(dcCustomer.BoundText))

'„” √Ã—

DoEvents
 Msg = "  „ ⁄„·    ’›Ì…  " & "  ··ÊÕœ… —ﬁ„   " & DcbUnitNo.text & CHR(13) & "    ··⁄ﬁ«— —ﬁ„ " & DcbIqara.text
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
   Sub SaveNotes()
   Dim NoteIDs As String
   Dim NoteID As Double
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL, Msg As String
        Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  Notes Where (1 = -1)"
      RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      NoteIDs = CStr(new_id("Notes", "NoteID", "", True))
      NoteID = NoteIDs
      RsDetails1.AddNew
      RsDetails1("CusID").value = val(dcCustomer.BoundText)
      RsDetails1("branch_no").value = val(dcBranch.BoundText)
      RsDetails1("NoteType").value = -1
      RsDetails1("NoteID").value = NoteID
      RsDetails1("akarid").value = val(Me.DcbIqara.BoundText)
      RsDetails1.Fields("UnitType").value = val(DcbUnitType.BoundText)
      RsDetails1.Fields("UnitNo").value = val(DcbUnitNo.BoundText)
      RsDetails1("NoteDate").value = XPDtbTrans.value
      RsDetails1("FilterID2").value = val(XPTxtID.text)
      RsDetails1("txtOldInsurance").value = val(txtOldInsurance.text)
    With grd
      RsDetails1("RemainRent").value = val(.TextMatrix(.rows - 1, .ColIndex("RemainRent")))
      RsDetails1("RemainWater").value = val(.TextMatrix(.rows - 1, .ColIndex("RemainWater")))
      RsDetails1("BillPrice").value = val(.TextMatrix(.rows - 1, .ColIndex("BillPrice")))
      RsDetails1("RemainCommissions").value = val(.TextMatrix(.rows - 1, .ColIndex("RemainCommissions")))
      RsDetails1("OldRent").value = val(.TextMatrix(.rows - 1, .ColIndex("OldRent")))
      RsDetails1("RemainService").value = val(.TextMatrix(.rows - 1, .ColIndex("RemainService")))
      RsDetails1("insurance").value = val(.TextMatrix(.rows - 1, .ColIndex("insurance")))
   End With
      RsDetails1.update

   End Sub
   
  Sub SaveUoitInformation()
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL, Msg As String
Msg = ""

 
    If SystemOptions.UserInterface = EnglishInterface Then
      Msg = Msg & "Action filter and a waiver of the decade No."
             Msg = Msg & CHR(13) & XPTxtID.text
               Msg = Msg & CHR(13)
         Msg = Msg & " Dtae  "
      Msg = Msg & NourHijriCal1.value & "corresponding to" & XPDtbTrans.value
      Msg = Msg & CHR(13)
        Msg = Msg & " The amount due from Renter  "
      Msg = Msg & TxtForRenter.text
      Msg = Msg & CHR(13)
        Msg = Msg & "  The amount due to Renter "
      Msg = Msg & TxtOFRenter.text
      Msg = Msg & CHR(13)
       Msg = Msg & "  Net "
      Msg = Msg & txtNet.text
      Msg = Msg & CHR(13)
      Else
      Msg = Msg & "   „ ⁄„·  ’›ÌÂ Ê ‰«“· »—ﬁ„  "
      Msg = Msg & XPTxtID.text
      Msg = Msg & CHR(13)
         Msg = Msg & " » «—ÌŒ  "
      Msg = Msg & NourHijriCal1.value & "«·„Ê«›ﬁ" & XPDtbTrans.value
      Msg = Msg & CHR(13)
        Msg = Msg & " «·„»·€ «·„” Õﬁ ⁄·Ï «·„” «Ã—  "
      Msg = Msg & TxtForRenter.text
      Msg = Msg & CHR(13)
        Msg = Msg & "  «·„»·€ «·„” Õﬁ ··„” √Ã—  "
      Msg = Msg & TxtOFRenter.text
      Msg = Msg & CHR(13)
       Msg = Msg & "  «·’«›Ì "
      Msg = Msg & txtNet.text
      Msg = Msg & CHR(13)

End If
        Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TblUnitNoInformation Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      RsDetails1.AddNew
      RsDetails1("CusID").value = val(dcCustomer.BoundText)
      RsDetails1("BranchId").value = val(dcBranch.BoundText)
      
           RsDetails1("UnitNo").value = val(DcbUnitNo.BoundText)
           RsDetails1("UnitStatus").value = UonitStatus
           RsDetails1("Des").value = Msg
           RsDetails1("RecDate").value = XPDtbTrans.value
           RsDetails1("RecDateH").value = NourHijriCal1.value
           RsDetails1("NoteID").value = Null
           RsDetails1("ContNo").value = Null
           RsDetails1("FilterNo").value = val(XPTxtID.text)
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
    
FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 2
            FG.Enabled = True
               
    Set RsDetails = New ADODB.Recordset

     
StrSQL = " SELECT     TOP 100 PERCENT dbo.TblAqrCompenetDet.IDAqComp, dbo.TblAqrCompenetDet.Name AS NameT, dbo.TblAqrCompenetDet.Price, dbo.TblAqrCompenetDet.ID, "
 StrSQL = StrSQL & "                     dbo.TblAqrCompenet.Namee, dbo.TblAqrCompenet.Name, dbo.TblAqrCompenetDet.Namee AS NameET"
StrSQL = StrSQL & ",TblAqrCompenetDet.Accountsus  FROM         dbo.TblAqrCompenet LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAqrCompenetDet ON dbo.TblAqrCompenet.ID = dbo.TblAqrCompenetDet.IDAqComp"
StrSQL = StrSQL & " ORDER BY dbo.TblAqrCompenetDet.IDAqComp"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
     
     Dim j, k As Integer
   Dim Temp As Integer
j = 0
Temp = -1
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.FG
      '  RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount
k = 0
        For i = .FixedRows To .rows - 1
    j = j + 1
    k = k + 1
    
            If Temp = val(IIf(IsNull(RsDetails("IDAqComp").value), 0, RsDetails("IDAqComp").value)) Then
                .TextMatrix(k, .ColIndex("serial")) = j
           
                .TextMatrix(k, .ColIndex("iditem")) = val(IIf(IsNull(RsDetails("id").value), "", RsDetails("id").value))
                If SystemOptions.UserInterface = EnglishInterface Then
                    .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("NameeT").value), "", RsDetails("NameT").value)
                Else
                    .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("NameT").value), "", RsDetails("NameT").value)
                End If
                .TextMatrix(k, .ColIndex("price")) = val(IIf(IsNull(RsDetails("Price").value), "", RsDetails("Price").value))
            
            Else
           
                .rows = .rows + 1
                .TextMatrix(k, .ColIndex("iditem")) = 0
                .TextMatrix(k, .ColIndex("id")) = val(IIf(IsNull(RsDetails("IDAqComp").value), "", RsDetails("IDAqComp").value)) 'val(RsDetails("IDAqComp").value)
                .TextMatrix(k, .ColIndex("serial")) = ""
                 If SystemOptions.UserInterface = EnglishInterface Then
                    .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("Namee").value), "", RsDetails("Namee").value)
                Else
                    .TextMatrix(k, .ColIndex("group")) = IIf(IsNull(RsDetails("Name").value), "", RsDetails("Name").value)
                End If
                .TextMatrix(k, .ColIndex("price")) = ""
                .cell(flexcpBackColor, k, 1, k, 7) = &H80C0FF
             
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
                Temp = val(IIf(IsNull(RsDetails("IDAqComp").value), 0, RsDetails("IDAqComp").value))
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

    With FG

        For i = .FixedRows To .rows - 1
 
             '  If fg.TextMatrix(i, fg.ColIndex("Accountsus")) <> "" Then
                                    '  If val(.TextMatrix(i, .ColIndex("count"))) <> 0 Then
                             .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("price"))) * val(.TextMatrix(i, .ColIndex("count")))
                               lbl(12).Caption = val(lbl(12).Caption) + val(.TextMatrix(i, .ColIndex("total")))
                        
                           ' End If
    
   'End If
 
        Next i
        
        Dim totals As String
        'totals = val(txtRemainWater) + val(txtRemainRent) + val(txtRemainService)
        
     '   TxtForRenter.text = val(lbl(12).Caption) + val(TxtBillPrice)
 
 TxtForRenter.text = 0
  TxtOFRenter.text = 0
 'TxtOFRenter.Text = val(Me.TxtInsurance.Text) + val(Me.txtOldInsurance.Text)
 '·Â
 
 
 
 If val(txtInsurance.text) = 0 Then
    TxtForRenter.text = Round(val(txtInsurance.text) + val(Me.txtOldInsurance.text) + val(TxtAmountDely) + val(lbl(12).Caption) + val(txtTotalCounterNet), 3) - val(txtDiscount)
Else
    TxtForRenter.text = Round(val(txtInsurance.text) + val(TxtAmountDely) + val(lbl(12).Caption) + val(txtTotalCounterNet), 3) - val(txtDiscount)
End If
 
 
' If totals > 0 Then
' TxtForRenter.Text = Round(val(TxtForRenter.Text) + val(totals), 3)
'
' Else
' TxtOFRenter = Round(val(TxtOFRenter) + val(Abs(totals)), 3)
' End If
'⁄·ÌÂ  ”«·„
 TxtForRenterB = Round(val(txtTotal1) + val(lbl(12).Caption) + val(txtTotalCounterNet), 3)
 TxtForRenter = Round(val(txtTotal1) + val(lbl(12).Caption) + val(txtTotalCounterNet), 3) - val(txtDiscount)
'·Â ”«·„
 TxtOFRenter = Round(val(txtTotal2), 3) + Round(val(Me.txtOldInsurance.text), 3)
 '’«›Ì
 txtNet.text = Round(val(TxtForRenter.text) - val(TxtOFRenter.text), 3)


 ReLineGridCount
    End With
    

'Me.Lbtotal.Caption = val(Me.LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
End Sub
Private Sub ReLineGridCount()
    Dim i As Integer
    Dim IntCounter  As Integer

    IntCounter = 0

    With FG

        For i = .FixedRows To .rows - 1

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

            TxtModFlg.text = "N"
            clear_all Me
        XPDtbTrans.value = Date
    FilterDate.value = Date
    FilterDateH.value = ToHijriDate(Date)
    NourHijriCal1.value = ToHijriDate(Date)
' RetriveIqarCOmpenet
dcBranch.BoundText = Current_branch
  Me.DCboUserName.BoundText = user_id
  
  ReLineGrid
  grd.rows = 1
  TxtVAtPercent = 5
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
 FG.rows = FG.rows + 1
            FG.Enabled = True
            '
            TxtModFlg.text = "E"
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
          
            FrmIqarWaiverSet.show vbModal



        Case 6
            Unload Me

        Case 7
   

        Case 8
  
            
            
                 Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text)
        
        
            End If
           Case 20
        If TxtOrder <> "" Then
'RtriveInfoOrbon val(TxtNotID.Text)
End If
        If FlagContrNew2 = False Then
        If TxtNoteSerial.text <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
Else
MsgBox "Please Delete JE"
End If
Exit Sub
End If
End If
            If Me.TxtModFlg.text <> "R" Then
                If opt(4).value = False And opt(3).value = False And opt(2).value = False Then
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
                 MSGType = MsgBox("”Ê›  „ Õ–› ﬁÌœ «·œ›⁄«  ", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
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
        
End Sub
Sub DeleteJE()


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
    s = s & "        dbo.TblFiterWaiver.EndDate,TblFiterWaiver.LastInstalldate,"
    s = s & "        dbo.TblFiterWaiver.FilterDate,"
    s = s & "        dbo.TblFiterWaiver.FilterDateH,"
    s = s & "        t2.BillPrice,"
    s = s & "        dbo.TblFiterWaiver.AccountNo,"
    s = s & "        dbo.TblFiterWaiver.AmountDely,"
    s = s & "        dbo.TblFiterWaiver.DayNo,"
    s = s & "        dbo.TblFiterWaiver.UserID,"
    s = s & "        dbo.TblFiterWaiver.OFRenter,"
    s = s & "        dbo.TblFiterWaiver.ForRenter,"
    s = s & "        dbo.TblFiterWaiver.ForRenterB,"
    s = s & "        dbo.TblFiterWaiver.Discount,"
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
    s = s & "        dbo.TblFiterWaiver.outflow,"
    
    
   

    
   s = s & "        dbo.TblFiterWaiver.LastInvoiceRead ServicePrice  ,"
   s = s & "        dbo.TblFiterWaiver.LastInvoiceRead2 WaterPriceotal,"
   s = s & "        dbo.TblFiterWaiver.Diff RentValuePayed,"
   s = s & "        dbo.TblFiterWaiver.Price NoDaye,"
   s = s & "        dbo.TblFiterWaiver.R ValDay,"
   s = s & "        dbo.TblFiterWaiver.PrevBalance totalpayed,"
   s = s & "        dbo.TblFiterWaiver.ServiceCounter totalcollected,"
   s = s & "        dbo.TblFiterWaiver.TotalCounter net,"
    
    s = s & "        dbo.TblFiterWaiver.DayPricen,"
'    s = s & "        T2.WaterPriceotal,"
'    s = s & "        T2.ServicePrice,"
'    s = s & "        T2.DayPricentotal,"
    s = s & "        T2.Service,"
    s = s & "        T2.WaterPayed,"
   ' s = s & "        T2.RentValuePayed,"
    s = s & "        T2.OldRent TelandNetPayed,"
    s = s & "        T2.RemainWater,"
    s = s & "        T2.RemainRent,"
    s = s & "        T2.RemainService,"
    s = s & "        T2.Insurance,"
    's = s & "        T2.outflow,"
    s = s & "        T2.StartDate,"
    s = s & "        T2.StartDateh,"
    s = s & "        T2.TotalStill,"
    s = s & "        T2.RemainCommissions,TblContract.UnitElectric,TblContract.NewNO,TblContract.Accredit,"
    's = s & "        T2.NoDaye,"
    s = s & "        dbo.TblFiterWaiver.outCondition,"
    s = s & "        dbo.TblFiterWaiver.DaysValueIncrease,"
    s = s & "        dbo.TblFiterWaiver.DaysValueIncomplete,"
    s = s & "        dbo.TblFiterWaiver.DayValueInc,"
    s = s & "        dbo.TblFiterWaiver.DayCountInc,"
    s = s & "        dbo.TblFiterWaiver.DayValueIncomplete,"
    s = s & "        dbo.TblFiterWaiver.DayCountIncomplete,"
    s = s & "        dbo.TblFiterWaiver.Efflux,"
   ' s = s & "        dbo.TblFiterWaiver.ValDay,"
    s = s & "        dbo.TblFiterWaiver.Discount," & txtTotalCounterNet & "   WaterPrice ,"
 '   s = s & "        dbo.TblFiterWaiver.totalcollected,"
  '  s = s & "        dbo.TblFiterWaiver.totalpayed,"
    s = s & "        dbo.TblFiterWaiver.LegalIssue,"
 '   s = s & "        dbo.TblFiterWaiver.net"
 s = s & " TotalMaint = (Select "
 s = s & "        SUM([Count] *  dbo.TblFiterWaiverDe.Price)"
     
    s = s & "        FROM         dbo.TblFiterWaiverDe "
    s = s & "        Where TblFiterWaiverDe.IDFItWaiv = TblFiterWaiver.ID)"
    s = s & " From dbo.TblAkarUnit"
    s = s & "        RIGHT OUTER JOIN dbo.TblFiterWaiver"
    s = s & "             ON  dbo.TblAkarUnit.id = dbo.TblFiterWaiver.unittype"
    s = s & "        LEFT OUTER JOIN dbo.TblAqarDetai"
    s = s & "             ON  dbo.TblFiterWaiver.ApartmentID = dbo.TblAqarDetai.Id"
    s = s & "        LEFT OUTER JOIN dbo.TblCustemers"
    s = s & "             ON  dbo.TblFiterWaiver.RenterID = dbo.TblCustemers.CusID"
    s = s & "        LEFT OUTER JOIN dbo.TblAqar"
    s = s & "             ON  dbo.TblFiterWaiver.BulidID = dbo.TblAqar.Aqarid"
    s = s & "        LEFT OUTER JOIN dbo.TblContract"
    s = s & "             ON  dbo.TblFiterWaiver.ContNo= dbo.TblContract.ContNo"
    

            
    s = s & "        LEFT OUTER JOIN dbo.TblFiterWaiverDet2 T2"
    s = s & "             ON  T2.MasterID = dbo.TblFiterWaiver.ID"
  

    s = s & "        Where (dbo.TblFiterWaiver.id = " & val(XPTxtID.text) & ")"
    
    
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If



    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData





    s = "SELECT     dbo.TblFiterWaiverDe.IDFItWaiv, dbo.TblFiterWaiverDe.[Count], dbo.TblFiterWaiverDe.Remark, dbo.TblFiterWaiverDe.Price, "
    s = s & "                     dbo.TblAqrCompenetDet.Name AS nameDet, dbo.TblFiterWaiverDe.IDItem, dbo.TblAqrCompenet.Name, dbo.TblFiterWaiverDe.GroupID "
    s = s & ",TblFiterWaiverDe.Accountsus   FROM         dbo.TblFiterWaiverDe LEFT OUTER JOIN "
    s = s & "                      dbo.TblAqrCompenetDet ON dbo.TblFiterWaiverDe.IDItem = dbo.TblAqrCompenetDet.ID LEFT OUTER JOIN "
    s = s & "                      dbo.TblAqrCompenet ON dbo.TblFiterWaiverDe.GroupID = dbo.TblAqrCompenet.ID "
    s = s & "  Where (dbo.TblFiterWaiverDe.IDFItWaiv = " & val(XPTxtID.text) & ")"
       
     
       
    

   Set RsData = New ADODB.Recordset
    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText


    


    If s <> "" Then
        Dim RsData2  As New ADODB.Recordset
        
         
      '  RsData2.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        xReport.OpenSubreport("aa").Database.SetDataSource RsData
  
    End If
    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng

        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
       xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtForRenter.text), "0.00"), 0, True, ".")
          xReport.ParameterFields(5).AddCurrentValue WriteNo(Format(val(TxtOFRenter.text), "0.00"), 0, True, ".")
          xReport.ParameterFields(6).AddCurrentValue WriteNo(Format(val(txtNet.text), "0.00"), 0, True, ".")
          xReport.ParameterFields(7).AddCurrentValue (lbl(12).Caption)
          xReport.ParameterFields(8).AddCurrentValue WriteNo(Format(val(lbl(12).Caption), "0.00"), 0, True, ".")
          xReport.ParameterFields(9).AddCurrentValue "" & txtTotalLastDays & ""
          xReport.ParameterFields(10).AddCurrentValue "" & txtOldInsurance & ""
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , s

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
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub


Private Sub Command9_Click()
    ShowGL_cc Me.TxtNoteSerial.text, , 200, val(TXTNoteID.text)
End Sub

Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then: Exit Sub

    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.text = EmpCode
   ' dcsupplier.BoundText = ownerid
    'DcbUnitType_Change
End Sub


Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    
End Sub

Private Sub dcCustomer_Click(Area As Integer)
    dcCustomer_Change
End Sub


Private Sub dcCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmAqarSearch
FrmIqarContractSearch.m_RetrunType = 1
FrmIqarContractSearch.show


End If
End Sub

Private Sub ENDDATE_Change()
If Me.TxtModFlg.text <> "R" Then
         EndDateH.value = ToHijriDate(EndDate.value)
'         IntMintsCount = (DateDiff("d", EndDate, FilterDate))
'Me.TxtDayLate.text = IntMintsCount
End If
End Sub

Private Sub ENDDATEH_LostFocus()

       If Me.TxtModFlg.text <> "R" Then
              VBA.Calendar = vbCalGreg
            EndDate.value = ToGregorianDate(EndDateH.value)
            End If
End Sub

Private Sub FilterDate_Change()
If Me.TxtModFlg.text <> "R" Then
Dim pricrday As Double
         FilterDateH.value = ToHijriDate(FilterDate.value)
         Dim IntMintsCount As Integer
         RetriveOrder
         GetContract val(TxtOrder), val(TxtContNo)


            If IntMintsCount > 0 Then
            pricrday = val(TxtDayPrice.text) / IntMintsCount
           ' TxtAmountDely.text = pricrday * val(TxtDayLate.text)
            Else
            pricrday = 0
            ' TxtAmountDely.text = 0
            End If
dcCustomer_Change

End If

End Sub


Private Sub FilterDateH_LostFocus()
If Me.TxtModFlg.text <> "R" Then
Dim pricrday As Double
           VBA.Calendar = vbCalGreg
           FilterDate.value = ToGregorianDate(FilterDateH.value)
           
         Dim IntMintsCount As Integer
         RetriveOrder
         GetContract val(TxtOrder), val(TxtContNo)


            If IntMintsCount > 0 Then
            pricrday = val(TxtDayPrice.text) / IntMintsCount
           ' TxtAmountDely.text = pricrday * val(TxtDayLate.text)
            Else
            pricrday = 0
            ' TxtAmountDely.text = 0
            End If
dcCustomer_Change

End If
End Sub


Function CALCdISCOUNT() As Double
CALCdISCOUNT = Round(val(TxtWaterPrice) * val(TxtDayLate), 2) + Round(val(TxtDayPricen) * val(TxtDayLate), 2) + Round(val(txtServicePrice) * val(TxtDayLate), 2)


End Function




 

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()

'
'Load FrmIqarContractSearch
'FrmIqarContractSearch.fg.TextMatrix(fg.Row, fg.ColIndex("NoteSerial")) = Me.TxtOrder
'FrmIqarContractSearch.m_RetrunType = 2
'FrmIqarContractSearch.show vbModal

Load FrmIqarContractSearch
'FrmIqarContractSearch.fg.TextMatrix(fg.Row, fg.ColIndex("NoteSerial"))=me.Text15
FrmIqarContractSearch.m_RetrunType = 2
FrmIqarContractSearch.show vbModal


Dim IntMintsCount As Integer
RetriveOrder val(TxtContNo)
GetContract (TxtOrder), val(TxtContNo)

End Sub

Private Sub lblDataCli_DragDrop(Source As Control, X As Single, Y As Single)
''''
End Sub

Private Sub NourHijriCal1_LostFocus()
      If Me.TxtModFlg.text <> "R" Then
             VBA.Calendar = vbCalGreg
           XPDtbTrans.value = ToGregorianDate(NourHijriCal1.value)
           End If
End Sub


Public Sub RetriveOrder(Optional order_no As String = "", Optional serial As Integer, Optional ByVal mCustId As Long = 0, Optional ByVal mUnitId As Long = 0)
   mCustId = val(dcCustomer2.BoundText)
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
  '  If serial = 1 Then
  If val(TxtOrder) = 0 And mCustId = 0 And val(order_no) = 0 Then Exit Sub
StrSQL = " SELECT distinct contNo,"
StrSQL = StrSQL & "         TblContractInstallments.Installdate , TblContractInstallments.InstalldateH"
StrSQL = StrSQL & " From dbo.TblContractInstallments"
StrSQL = StrSQL & "        Left Outer JOIN ContracttBillInstallmentsDone T2"
StrSQL = StrSQL & "             ON  T2.istallid = TblContractInstallments.ID"
StrSQL = StrSQL & " WHERE     "

If val(order_no) = 0 Then
    StrSQL = StrSQL & "  contNo IN ("
Else
    StrSQL = StrSQL & " contNo =" & val(order_no) & " or contNo IN ("
End If

StrSQL = StrSQL & " SELECT  TblContract.ContNo FROM TblContract WHERE "
StrSQL = StrSQL & " Installdate <= " & SQLDate(FilterDate.value, True) & " and "
'If val(TxtOrder) <> 0 Then
   ' StrSQL = StrSQL & " (    (NoteSerial1 = '" & Trim(TxtOrder) & "' Or ContNo = " & val(order_no) & "  )     and ( 1 =1  "
'Else

If SystemOptions.WaiverSetByContract Then
    StrSQL = StrSQL & " (    (  NoteSerial1 = '" & Trim(TxtOrder) & "' Or ContNo = " & val(order_no) & "  )    and ( 1 =1    "
Else
    StrSQL = StrSQL & "  (( 1 =1  "
End If
   
'End If

        If Not SystemOptions.WaiverSetByContract Then
            If val(DcbUnitType2.BoundText) <> 0 Then
                StrSQL = StrSQL & " And  unittype = " & val(DcbUnitType2.BoundText)
            End If
    
            If val(DcbIqara2.BoundText) <> 0 Then
                StrSQL = StrSQL & " And  Iqar = " & val(DcbIqara2.BoundText)
            End If
    
            If val(DcbUnitNo2.BoundText) <> 0 Then
                StrSQL = StrSQL & " And UnitNo = " & val(DcbUnitNo2.BoundText)
            End If
    
            If mCustId <> 0 Then
                StrSQL = StrSQL & " And CusID = " & val(dcCustomer2.BoundText)
            
            End If
        End If
        StrSQL = StrSQL & "Or 1 =-1)"
        StrSQL = StrSQL & " ))"
        




StrSQL = StrSQL & " Order By"
StrSQL = StrSQL & " TblContractInstallments.Installdate desc"
Set rs = New ADODB.Recordset

rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.EOF Then
    txtLastInstalldate.Visible = True
    txtInstalldateH.Visible = True
    lbl(76).Visible = True
    StrSQL = " SELECT Top 1 contNo,"
StrSQL = StrSQL & "         TblContractInstallments.Installdate , TblContractInstallments.InstalldateH"
StrSQL = StrSQL & " From dbo.TblContractInstallments"
StrSQL = StrSQL & "        Left Outer JOIN ContracttBillInstallmentsDone T2"
StrSQL = StrSQL & "             ON  T2.istallid = TblContractInstallments.ID"
StrSQL = StrSQL & " WHERE     "

If val(order_no) = 0 Then
    StrSQL = StrSQL & "  contNo IN ("
Else
    StrSQL = StrSQL & " contNo =" & val(order_no) & " or contNo IN ("
End If

StrSQL = StrSQL & " SELECT  TblContract.ContNo FROM TblContract WHERE "
StrSQL = StrSQL & " Installdate >= " & SQLDate(Trim(rs!installdate & ""), True) & " and "

If SystemOptions.WaiverSetByContract Then
    If val(TxtOrder) <> 0 Then
        StrSQL = StrSQL & "   (  (NoteSerial1 = '" & Trim(TxtOrder) & "' Or ContNo = " & val(order_no) & " )     and ( 1 =1  "
    Else
        StrSQL = StrSQL & "  (( 1 =1  "
    End If
Else
    StrSQL = StrSQL & "  (( 1 =1  "
End If

        If Not SystemOptions.WaiverSetByContract Then
            If val(DcbUnitType2.BoundText) <> 0 Then
                StrSQL = StrSQL & " And  unittype = " & val(DcbUnitType2.BoundText)
            End If
    
            If val(DcbIqara2.BoundText) <> 0 Then
                StrSQL = StrSQL & " And  Iqar = " & val(DcbIqara2.BoundText)
            End If
    
            If val(DcbUnitNo2.BoundText) <> 0 Then
                StrSQL = StrSQL & " And UnitNo = " & val(DcbUnitNo2.BoundText)
            End If
    
            If mCustId <> 0 Then
                StrSQL = StrSQL & " And CusID = " & val(dcCustomer2.BoundText)
            
            End If
        End If
        StrSQL = StrSQL & "Or 1 =-1)"
        StrSQL = StrSQL & " ))"
        



    
    StrSQL = StrSQL & " Order By"
    StrSQL = StrSQL & " TblContractInstallments.Installdate "
    Set rs = New ADODB.Recordset

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        txtLastInstalldate.value = DateAdd("D", 0, rs!installdate & "")
        txtInstalldateH.value = DateAdd("D", 0, rs!installdateh & "")
    End If
Else
    txtLastInstalldate.Visible = False
    txtInstalldateH.Visible = False
    lbl(76).Visible = False
End If
    
    
  StrSQL = " SELECT *  FROM TblContract WHERE "

If val(TxtOrder) <> 0 Then
    'StrSQL = StrSQL & "  (dbo.TblContract.EndContract IS NULL)        "
    StrSQL = StrSQL & "  (1 = 1         "
    If SystemOptions.WaiverSetByContract Then
        StrSQL = StrSQL & " and    (NoteSerial1 = '" & Trim(TxtOrder) & "' Or ContNo = " & val(order_no) & " )     and ( 1 =1   "
        'StrSQL = StrSQL & "          and ( 1 =1  "
    End If
Else
    StrSQL = StrSQL & "  (( 1 =1  "
End If
        If Not SystemOptions.WaiverSetByContract Then
            If val(DcbUnitType2.BoundText) <> 0 Then
                StrSQL = StrSQL & " And  unittype = " & val(DcbUnitType2.BoundText)
            End If
    
            If val(DcbIqara2.BoundText) <> 0 Then
                StrSQL = StrSQL & " And  Iqar = " & val(DcbIqara2.BoundText)
            End If
    
            If val(DcbUnitNo2.BoundText) <> 0 Then
                StrSQL = StrSQL & " And UnitNo = " & val(DcbUnitNo2.BoundText)
            End If
    
            If mCustId <> 0 Then
                StrSQL = StrSQL & " And CusID = " & val(dcCustomer2.BoundText)
            
            End If
        End If
        StrSQL = StrSQL & " Or 1 =-1"
        StrSQL = StrSQL & " )"
        
    
'    If val(TxtContNo.Text) <> 0 Then
'       StrSQL = "Select * from TblContract  where   ( ContNo=" & val(TxtContNo.Text) & ""
'        If mUnitId <> 0 Then
'            StrSQL = StrSQL & " Or UnitNo = " & val(DcbUnitNo2.BoundText)
'        End If
'        If mCustId <> 0 Then
'            StrSQL = StrSQL & " Or CusID = " & val(dcCustomer2.BoundText)
'        End If
'       StrSQL = StrSQL & ")"
'    Else
'        StrSQL = "Select * from TblContract  where    ( NoteSerial1='" & (Me.TxtOrder) & "'"
'        If mUnitId <> 0 Then
'            StrSQL = StrSQL & " Or UnitNo = " & val(DcbUnitNo2.BoundText)
'        End If
'        If mCustId <> 0 Then
'            StrSQL = StrSQL & " Or CusID = " & val(dcCustomer2.BoundText)
'        End If
'        StrSQL = StrSQL & ")"
'
'    End If
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
     DcbIqara.BoundText = IIf(IsNull(rs("Iqar").value), "", rs("Iqar").value)
     DcbUnitType.BoundText = IIf(IsNull(rs("unittype").value), "", rs("unittype").value)
     DcbUnitNo.BoundText = val(IIf(IsNull(rs("UnitNo").value), "", rs("UnitNo").value))
     TxtElectAccount = rs!UnitElectric & ""
     TXTNewNO.text = IIf(IsNull(rs("NewNO").value), "", rs("NewNO").value)
     
            
    If IsNull(rs.Fields("Accredit").value) Then
        ChkAccredit.value = vbUnchecked
    Else
        ChkAccredit.value = vbChecked
      
    End If
           
     dcCustomer.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
     'dcCustomer2.BoundText = IIf(IsNull(rs("CusID2").value), "", rs("CusID2").value)
       'TxtOrder.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    EndDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    EndDateH.value = IIf(IsNull(rs("TodateH").value), "", rs("TodateH").value)
    
    StartDate.value = IIf(IsNull(rs("StrDate").value), Date, rs("StrDate").value)
    StartDateh.value = IIf(IsNull(rs("FromdateH").value), "", rs("FromdateH").value)
    Dim IntMintsCount   As Integer
    IntMintsCount = (DateDiff("d", EndDate, FilterDate))
    'Me.TxtDayLate.text = IntMintsCount
    IntMintsCount = (DateDiff("d", From, EndDate))
    TxtActualDays.text = (DateDiff("d", StartDate.value, FilterDate.value))
TxtContractDays.text = (DateDiff("d", (rs("StrDate").value), (rs("EndDate").value)))
'TxtContractDays.Text = val(val(TxtContractDays.Text) * 30)
    'datediff("m",date(FromdateH),date(TodateH))
   
    TxtAccountNo.text = Me.Text15.text
       ' TxtActualDays.Text = (DateDiff("d", startDate, FilterDate))
' TxtContractDays.text = (DateDiff("d", CDate(rs("Fromdateh").value), CDate(rs("todateH").value)))
'TxtContractDays.text = (DateDiff("d", startDate, EndDate))

    
    TxtDayLate = val(TxtContractDays.text) - val(TxtActualDays.text)
 
     If Not IsNull(rs.Fields("ComResid").value) Then
        If rs.Fields("ComResid").value = 1 Then
            ComResid(1).value = True
        Else
            ComResid(0).value = True
        End If
   Else
        ComResid(0).value = True
   End If

  If Not IsNull(rs("TypeDate").value) Then
        If rs("TypeDate").value = 1 Then
            RdRTypeDate(1).value = True
        Else
            RdRTypeDate(0).value = True
        End If
    Else
        RdRTypeDate(0).value = True
    End If

       From.value = IIf(IsNull(rs("StrDate").value), Date, rs("StrDate").value)
         
     Me.TxtDayPrice.text = IIf(IsNull(rs("TotalContract").value), "", rs("TotalContract").value)
    txtInsurance.text = IIf(IsNull(rs("InsuranceValue").value), "", rs("InsuranceValue").value)
     TxtBillPrice.text = 0 ' IIf(IsNull(rs("Electricity").value), "", rs("Electricity").value)
TxtService.text = IIf(IsNull(rs("phone").value), "", rs("phone").value)
 Dim WaterPayed     As Double
Dim RentValuePayed As Double
Dim TelandNetPayed As Double

         If val(TxtContractDays.text) > 0 Then

      Me.TxtDayPricen.text = Round(IIf(IsNull(rs("TotalContract").value), 0, rs("TotalContract").value) / val(TxtContractDays.text), 2)
        Me.txtServicePrice.text = Round(IIf(IsNull(rs("phone").value), 0, rs("phone").value) / val(TxtContractDays.text), 2)
     
      
      
     TxtWaterPrice.text = Round(IIf(IsNull(rs("water").value), 0, rs("water").value) / val(TxtContractDays.text), 2)
      TxtDayPricentotal = val(Me.TxtDayPricen.text) * val(TxtActualDays.text)
    TxtWaterPriceotal = val(Me.TxtWaterPrice.text) * val(TxtActualDays.text)
     TxtService = val(Me.txtServicePrice.text) * val(TxtActualDays.text)
     getActualpayedToContract val(TxtContNo), RentValuePayed, WaterPayed, TelandNetPayed
      TxtRentValuePayed.text = RentValuePayed
      txtWaterPayed.text = WaterPayed
       txtTelandNetPayed.text = TelandNetPayed
       
     txtRemainRent.text = val(TxtDayPricentotal) - val(TxtRentValuePayed)
       txtRemainWater.text = val(TxtWaterPriceotal) - val(txtWaterPayed.text)
       txtRemainService.text = val(TxtService) - val(txtTelandNetPayed.text)
       
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
  TxtDayPricentotal = val(Me.TxtDayPricen.text) * val(TxtActualDays.text)
    TxtWaterPriceotal = val(Me.TxtWaterPrice.text) * val(TxtActualDays.text)
     TxtService = val(Me.txtServicePrice.text) * val(TxtActualDays.text)
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
ReLineGrid
If Me.TxtModFlg.text <> "R" Then


    Dim IntMintsCount As Integer
    RetriveOrder
    GetContract (TxtOrder), val(TxtContNo)
End If
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
sql = "select TblFiterWaiver.id from TblFiterWaiver Left Outer join  TblFiterWaiverDet2 On TblFiterWaiverDet2.MasterId = TblFiterWaiver.Id  where TblFiterWaiverDet2.ContNo=" & ContNo & " "
sql = sql & " AND iSnULL(TblFiterWaiver.ApartmentID2,0) =" & val(Me.DcbUnitNo2.BoundText)
sql = sql & " AND iSnULL(TblFiterWaiver.unittype2,0) = " & val(Me.DcbUnitType2.BoundText)
sql = sql & " AND iSnULL(TblFiterWaiver.BulidID2,0) = " & val(Me.DcbIqara2.BoundText)
sql = sql & " AND iSnULL(TblFiterWaiver.RenterID2,0) = " & val(Me.dcCustomer2.BoundText)

      
     If SystemOptions.usertype <> UserAdminAll Then
             sql = sql & "   and TblFiterWaiver.BranchID=" & Current_branch & " "
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

Private Sub txtDayCountInc_Change()
calctextTotal
End Sub

Private Sub txtDayCountIncomplete_Change()
calctextTotal
End Sub

Private Sub TxtDayLate_Change()
 TxtAmountDely.text = CALCdISCOUNT
 ReLineGrid
 
End Sub

Private Sub TxtDayLate_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtDayLate.text, 1)

 
End Sub





Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
Dim StrComboList As String
    'Dim Rs2 As ADODB.Recordset
On Error GoTo ErrTrap
    With FG
               
  
  
Select Case .ColKey(Col)
  
Case "group"
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Id"), False, True)
                .TextMatrix(Row, .ColIndex("Id")) = StrAccountCode
                s = "Select * from TblAqrCompenetDet Where Id = " & val(StrAccountCode)
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then
                    .TextMatrix(Row, .ColIndex("Accountsus")) = rsDummy!Accountsus & ""
                    .TextMatrix(Row, .ColIndex("iditem")) = StrAccountCode
                    .TextMatrix(Row, .ColIndex("price")) = rsDummy!Price & ""
                    
                End If
Case "price", "count"
    .TextMatrix(Row, .ColIndex("total")) = Round(val(.TextMatrix(Row, .ColIndex("count"))) * val(.TextMatrix(Row, .ColIndex("price"))), 2)
End Select
  If Row = .rows - 1 Then
    
            .rows = .rows + 1
    End If
End With
ReLineGrid
ErrTrap:
End Sub



Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With FG

        Select Case .ColKey(Col)
            
            Case "total"
               Cancel = True
            Case "price", "count", "remark", "serial"
            .ComboList = ""
        End Select

    End With


'With Grid
'
'   Select Case .ColKey(Col)
'        Case "Qun"
'        .ComboList = ""
'           Case "NoteNo"
'        .ComboList = ""
'        Case "DayMeter"
'        .ComboList = ""
'        Case "Name"
'       ' Cancel = True
'        End Select
'
'    End With

    
End Sub

Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
des = " ’›Ì… ⁄ﬁœ «ÌÃ«— —ﬁ„ " & TxtOrder & " · " & dcCustomer.text
Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
tablename = "TblFiterWaiver"
Filedname = "ID"
ContNo = XPTxtID

Notevalue = 0


                     If Me.TxtModFlg = "N" Then
                                 CreateNotes NoteID, (XPDtbTrans.value), val(dcBranch.BoundText), -1, Notevalue, NoteSerial, XPTxtID, tablename, Filedname, ContNo, des, NourHijriCal1.value
                                     TXTNoteID.text = NoteID
                                    TxtNoteSerial.text = NoteSerial
                    Else
                                      If TXTNoteID.text = "" Or TxtNoteSerial.text = "" Then
                                    CreateNotes NoteID, (XPDtbTrans.value), val(dcBranch.BoundText), -1, Notevalue, NoteSerial, TxtNoteSerial1, tablename, Filedname, ContNo, des, NourHijriCal1.value
                                                       TXTNoteID.text = NoteID
                                                  TxtNoteSerial.text = NoteSerial
                                    Else
                                                  sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                  sql = sql & ",NoteSerial1='" & XPTxtID & "'"
                                                    sql = sql & " where NoteID=" & val(TXTNoteID.text)
                                                     Cn.Execute sql
                                                     
                                       End If
                         
                    End If
ReLineGrid
CREATE_VOUCHER_GE val(TXTNoteID.text), val(dcBranch.BoundText), user_id, XPDtbTrans.value
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
            
  
            StrTempDes = " ’›Ì… ⁄ﬁœ «ÌÃ«— —ﬁ„    " & TxtNoteSerial1 & "  ··„” √Ã—   " & dcCustomer.text & " ··ÊÕœ… " & DcbUnitNo.text
            LngDevNO = LngDevNO + 1
'Notevalue = val(TxtTotalContract) + val(TxtPayAmini) + val(TxtCommiValue) + val(TxtInsuranceValue) + val(TxtWater) + val(TxtElectricity) + val(TxtPhone) + val(TxtEnternet)
Notevalue = 0
 
 Dim Account_Code_dynamic80 As String
  Dim Account_Code_dynamic81 As String
   Dim Account_Code_dynamic82 As String
    Dim Account_Code_dynamic83 As String
     Dim Account_Code_dynamic84 As String
      Dim Account_Code_dynamic85 As String
      
            Account_Code_dynamic80 = get_account_code_branch(86, my_branch)
            
            Account_Code_dynamic81 = get_account_code_branch(81, my_branch)
            Account_Code_dynamic82 = get_account_code_branch(82, my_branch)
            Account_Code_dynamic83 = get_account_code_branch(83, my_branch)
            Account_Code_dynamic84 = get_account_code_branch(84, my_branch)
            Account_Code_dynamic85 = get_account_code_branch(85, my_branch)
            
'll:
   LngDevNO = 0
  
 
 If val(txtDaysValueIncomplete.text) > 0 Then '«Ì«„ ‰«ﬁ’…
       '«·⁄„Ì· œ«∆‰
       Notevalue = Abs(val(txtDaysValueIncomplete.text))
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
  ''*************
  
  '99999999999999999999999999999999999999999999999
  If val(txtDaysValueIncrease.text) > 0 Then '«Ì«„ “Ì«œ…
       '«·⁄„Ì· œ«∆‰
       Notevalue = Abs(val(txtDaysValueIncrease.text))
   
        LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… «·«ÌÃ«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  

  LngDevNO = LngDevNO + 1
   StrTempAccountCode = Account_Code_dynamic80
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„… «·«ÌÃ«— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
                      
            
  End If
  
  '99999999999999999999999999999999999999999999999
  
   If val(txtRemainWater.text) > 0 And 0 = 1 Then
       '«·⁄„Ì· œ«∆‰
       Notevalue = Abs(val(txtRemainWater.text))
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
  
   
   If val(txtRemainService.text) < 0 And 0 = 1 Then
       '«·⁄„Ì· œ«∆‰
       Notevalue = Abs(val(txtRemainService.text))
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
If val(txtRemainRent.text) > 0 And 0 = 1 Then
       '«·⁄„Ì· „œÌ‰
       Notevalue = Abs(val(txtRemainRent.text))
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
   If val(txtRemainWater.text) > 0 And 0 = 1 Then
       '«·⁄„Ì· „œÌ‰
              Notevalue = Abs(val(txtRemainWater.text))
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
  
   
   If val(txtRemainService.text) > 0 And 0 = 1 Then
       '«·⁄„Ì· „œÌ‰
        
       '«·⁄„Ì· „œÌ‰
              Notevalue = Abs(val(txtRemainService.text))
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
   
   
   
   If val(txtTotalinsuranceS.text) > 0 Then
               
               Notevalue = Abs(val(txtTotalinsuranceS.text))
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
     
     
     
If val(txtTotalCounterNet.text) > 0 Then
       '  «·ﬂÂ—»«¡
       Notevalue = Abs(val(txtTotalCounterNet.text))
       
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
     
If val(TxtAmountDely.text) > 0 And 0 = 1 Then
       '  «·Œ’Ê„« 
       Notevalue = (val(TxtAmountDely.text))
       
               LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„…    Œ’„ «Ì«„ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
   LngDevNO = LngDevNO + 1
   
   
   If val(TxtDayPricen.text) * val(TxtDayLate) > 0 And 0 = 1 Then
   
   Notevalue = Round(val(TxtDayPricen.text) * val(TxtDayLate), 2)
     StrTempAccountCode = Account_Code_dynamic80
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„…   Œ’„ «Ì«„  ﬁÌ„… «ÌÃ«—  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
   End If
   
 
            
         LngDevNO = LngDevNO + 1
   
   
   If val(TxtWaterPrice.text) * val(TxtDayLate) > 0 And 0 = 1 Then
   Notevalue = Round(val(TxtWaterPrice.text) * val(TxtDayLate), 2)
     StrTempAccountCode = Account_Code_dynamic83
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„…   Œ’„ «Ì«„  ﬁÌ„… „Ì«Â  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
   End If
   

  
        LngDevNO = LngDevNO + 1
   
   
   If val(TxtService.text) * val(TxtDayLate) > 0 And 0 = 1 Then
   Notevalue = Round(val(TxtService.text) * val(TxtDayLate), 2)
     StrTempAccountCode = Account_Code_dynamic85
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & "       ﬁÌ„…   Œ’„ «Ì«„  ﬁÌ„… Œœ„«   ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
   End If
   

            
            
  End If
     
     
     
'**************************************************************************’Ì«‰…
  
     
     If val(lbl(12).Caption) > 0 And 0 = 1 Then
       '  «·⁄„Ì·
       Notevalue = Abs(val(lbl(12).Caption))
       
               LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & "       ﬁÌ„… ›Ê« Ì—  «·ﬂÂ—»«¡ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            

            

  

            
            
  End If
     


     
     Dim mDiscr As String
     
          For i = Me.FG.FixedRows To FG.rows - 1
    
                  If val(FG.TextMatrix(i, FG.ColIndex("total"))) > 0 And FG.TextMatrix(i, FG.ColIndex("Accountsus")) <> "" Then
              Notevalue = val(FG.TextMatrix(i, FG.ColIndex("total")))
                mDiscr = Trim(FG.TextMatrix(i, FG.ColIndex("group")))
            
                       LngDevNO = LngDevNO + 1
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & mDiscr, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
    
    
               LngDevNO = LngDevNO + 1
   StrTempAccountCode = FG.TextMatrix(i, FG.ColIndex("Accountsus"))
  
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & mDiscr, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
                   
                      End If
         
    
  
        Next i
  
ErrTrap:
End Function



Private Sub Form_Load()

    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

If SystemOptions.SpecialVersion = True Then
lbl(25).Visible = False
TxtNoteSerial.Visible = False
Command1.Visible = False
Command2.Visible = False
Command9.Visible = False
End If
   


If SystemOptions.CanEditLegalAffairs Then
    chkIsLegalAffairs.Enabled = True
    txtLegalAffairs.Enabled = True
Else
    chkIsLegalAffairs.Enabled = False
    txtLegalAffairs.Enabled = False

End If
    
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

   
    Dcombos.GetBranches Me.dcBranch
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.getAkarUnit Me.DcbUnitType2
    
    Dcombos.GetIqarUnit -2, 1, DcbUnitNo
    Dcombos.GetIqarUnit -2, 1, DcbUnitNo2
    
    Dcombos.GetCustomersSuppliers 56, Me.dcCustomer
    Dcombos.GetCustomersSuppliers 56, Me.dcCustomer2
     My_SQL = "select UserID,UserName From tblUsers "


  

    SetDtpickerDate Me.XPDtbTrans
   fill_combo DCboUserName, My_SQL
    Dcombos.GetIqar DcbIqara
    Dcombos.GetIqar DcbIqara2
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblFiterWaiver "
      If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
      StrSQL = StrSQL & "   where BranchID=" & Current_branch & "     Order By ID"
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        Me.TxtModFlg.text = "R"
            

grd.ColComboList(grd.ColIndex("TypeDate")) = "#0; ÂÃ—Ì|#1; „Ì·«œÌ"
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
    With Me.FG
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
txtRemainRent.text = val(TxtDayPricentotal.text) - val(txtRemainRent.text)
End Sub

Private Sub txtDaysValueIncomplete_Change()
CalcTotal
End Sub

Private Sub txtDaysValueIncrease_Change()
CalcTotal
End Sub

Private Sub txtDayValueInc_Change()
calctextTotal
End Sub

Private Sub txtDayValueIncomplete_Change()
calctextTotal
End Sub

Private Sub TxtForRenter_Change()
If val(TxtForRenter.text) > 0 Then
lbll(9).Caption = WriteNo(Round(Me.TxtForRenter.text, 3), 0)
Else
lbll(9).Caption = ""
End If
txtNet.text = Round(val(TxtForRenter.text) - val(TxtOFRenter.text), 3)
'TxtNet.text = val(Me.TxtForRenter.text) - val(Me.TxtOFRenter.text)
End Sub




Private Sub TxtInsurance_Change()
TxtOFRenter.text = val(TxtOFRenter.text) + val(txtInsurance.text)
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

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
               
            FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 2
            FG.Enabled = True
         
    
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
lbll(0).Caption = WriteNo(Round(val(Me.txtNet.text), 3), 0)
Else
lbll(0).Caption = ""
End If


End Sub

Private Sub TxtOFRenter_Change()

If val(TxtOFRenter) > 0 Then
lbll(11).Caption = WriteNo(Round(val(Me.TxtOFRenter.text), 3), 0)
Else
lbll(11).Caption = ""
End If


End Sub

Private Sub TxtOrder_Change()
ReLineGrid
If Me.TxtModFlg.text <> "R" Then


    Dim IntMintsCount As Integer
    If TxtOrder <> "" Then
        RetriveOrder TxtOrder
    Else
        RetriveOrder
    End If
    GetContract (TxtOrder), val(TxtContNo)
End If
'RetriveOrder TxtOrder, 0
End Sub

Private Sub TxtOrder_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmIqarContractSearch
'FrmIqarContractSearch.fg.TextMatrix(fg.Row, fg.ColIndex("NoteSerial"))=me.Text15
FrmIqarContractSearch.m_RetrunType = 2
FrmIqarContractSearch.show


End If
End Sub

Private Sub txtPrevBalance_Change()
CalcTotalCounter
End Sub

Private Sub txtPrice_Change()
CalcTotalCounter
End Sub

Private Sub TxtService_Change()
txtRemainService.text = val(TxtService.text) - val(txtTelandNetPayed.text)
End Sub

Private Sub txtServiceCounter_Change()
CalcTotalCounter
End Sub


Private Sub TxtVAtPercent_Change()
txtVat2.text = (val(txtTotalCounter) - val(txtPrevBalance)) * val(TxtVAtPercent) / 100
txtTotalCounterNet = val(txtVat2) + val(txtTotalCounter)
End Sub

Private Sub TxtWaterPriceotal_Change()
txtRemainWater.text = val(TxtWaterPriceotal.text) - val(txtWaterPayed.text)
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
    Dim i As Integer
    Dim StrSQL As String
 
FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 2
            FG.Enabled = True
         
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
TxtVAtPercent = 5
    XPTxtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    TxtContNo.text = IIf(IsNull(rs("ContNo").value), "", val(rs("ContNo").value))
    TxtManualNO.text = rs("ManualNo") & ""
           
   Me.TxtOrder.text = IIf(IsNull(rs("ContractNo").value), "", (rs("ContractNo").value))
    Me.TXTNoteID.text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
  Me.TxtNoteSerial.text = IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value)
   

    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    NourHijriCal1.value = IIf(IsNull(rs("RecordDateH").value), "", rs("RecordDateH").value)
     dcBranch.BoundText = val(IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value))
    dcCustomer.BoundText = val(IIf(IsNull(rs("RenterID").value), "", rs("RenterID").value))
    
    dcCustomer2.BoundText = val(IIf(IsNull(rs("RenterID2").value), "", rs("RenterID2").value))
    
    DcbIqara.BoundText = val(IIf(IsNull(rs("BulidID").value), "", rs("BulidID").value))
      DcbUnitType.BoundText = val(IIf(IsNull(rs("unittype").value), "", rs("unittype").value))
      
      DcbUnitType2.BoundText = val(IIf(IsNull(rs("unittype2").value), "", rs("unittype2").value))
      DcbIqara2.BoundText = val(IIf(IsNull(rs("BulidID2").value), "", rs("BulidID2").value))
    
   
     TxtElectAccount = rs!UnitElectric & ""
     TXTNewNO.text = IIf(IsNull(rs("NewNO").value), "", rs("NewNO").value)
     
            
    If IsNull(rs.Fields("Accredit").value) Then
        ChkAccredit.value = vbUnchecked
    Else
        ChkAccredit.value = vbChecked
      
    End If
    
    If IsNull(rs.Fields("isLegalAffairs").value) Then
        chkIsLegalAffairs.value = vbUnchecked
    Else
        If rs.Fields("isLegalAffairs").value = 0 Then
            chkIsLegalAffairs.value = vbUnchecked
        Else
            chkIsLegalAffairs.value = vbChecked
        End If
      
    End If
    
    txtLegalAffairs.text = Trim(rs("LegalAffairs").value & "")
        
        
    
    DcbUnitNo.BoundText = val(IIf(IsNull(rs("ApartmentID").value), "", rs("ApartmentID").value))
    DcbUnitNo2.BoundText = val(IIf(IsNull(rs("ApartmentID2").value), "", rs("ApartmentID2").value))
    txtOldInsurance.text = val(IIf(IsNull(rs("Insurance").value), 0, rs("Insurance").value))
     EndDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    EndDateH.value = IIf(IsNull(rs("EndDateH").value), "", rs("EndDateH").value)
    FilterDate.value = IIf(IsNull(rs("FilterDate").value), Date, rs("FilterDate").value)
    FilterDateH.value = IIf(IsNull(rs("FilterDateH").value), "", rs("FilterDateH").value)
    '
    txtOldInsurance.text = Round(IIf(IsNull(rs("Insurance").value), 0, rs("Insurance").value), 2)
    txtTotalinsuranceS.text = Round(IIf(IsNull(rs("TotalinsuranceS").value), 0, rs("TotalinsuranceS").value), 2)
    
     TxtForRenter.text = val(IIf(IsNull(rs("ForRenter").value), 0, rs("ForRenter").value))
     TxtForRenterB.text = val(IIf(IsNull(rs("ForRenterB").value), 0, rs("ForRenterB").value))
     txtDiscount.text = val(IIf(IsNull(rs("Discount").value), 0, rs("Discount").value))
      TxtOFRenter.text = val(IIf(IsNull(rs("OFRenter").value), 0, rs("OFRenter").value))
    '
     TxtBillPrice.text = val(IIf(IsNull(rs("BillPrice").value), 0, rs("BillPrice").value))
     Me.txtNet.text = val(IIf(IsNull(rs("net").value), 0, rs("net").value))
     TxtAccountNo.text = IIf(IsNull(rs("AccountNo").value), "", rs("AccountNo").value)
   TxtDayLate.text = IIf(IsNull(rs("DayNo").value), "", rs("DayNo").value)
     TxtAmountDely.text = IIf(IsNull(rs("AmountDely").value), "", rs("AmountDely").value)
'*******************************************************************************************

  
   
    txtLastInvoiceRead.text = val(IIf(IsNull(rs("LastInvoiceRead").value), 0, rs("LastInvoiceRead").value))
    txtLastInvoiceRead2.text = val(IIf(IsNull(rs("LastInvoiceRead2").value), 0, rs("LastInvoiceRead2").value))
    txtDiff.text = val(IIf(IsNull(rs("Diff").value), 0, rs("Diff").value))
    TxtPrice.text = val(IIf(IsNull(rs("Price").value), 0, rs("Price").value))
    txtR.text = val(IIf(IsNull(rs("R").value), 0, rs("R").value))
    txtPrevBalance.text = val(IIf(IsNull(rs("PrevBalance").value), 0, rs("PrevBalance").value))
    txtServiceCounter.text = val(IIf(IsNull(rs("ServiceCounter").value), 0, rs("ServiceCounter").value))
    txtTotalCounter.text = val(IIf(IsNull(rs("TotalCounter").value), 0, rs("TotalCounter").value))
   
    TxtVAtPercent.text = IIf(IsNull(rs("VAtPercent").value), "", rs("VAtPercent").value)
    txtVat2.text = IIf(IsNull(rs("VAt2").value), "", rs("VAt2").value)
    txtTotalCounterNet.text = IIf(IsNull(rs("TotalCounterNet").value), "", rs("TotalCounterNet").value)
    If val(txtTotalCounterNet) = 0 Then
        TxtVAtPercent.text = 5
        
    End If
 
   If Not IsNull(rs.Fields("ComResid").value) Then
        If rs.Fields("ComResid").value = 1 Then
            ComResid(1).value = True
        Else
            ComResid(0).value = True
        End If
   Else
        ComResid(0).value = True
   End If

  If Not IsNull(rs("TypeDate").value) Then
        If rs("TypeDate").value = 1 Then
            RdRTypeDate(1).value = True
        Else
            RdRTypeDate(0).value = True
        End If
    Else
        RdRTypeDate(0).value = True
    End If


    If IsNull(rs("CalcLastPayment").value) Then
        ChkCalcLastPayment.value = vbUnchecked
    
    Else
        If (rs("CalcLastPayment").value) = vbFalse Then
            ChkCalcLastPayment.value = vbUnchecked
        Else
            ChkCalcLastPayment.value = vbChecked
        End If
    End If

    txtLastInstalldate.value = IIf(IsNull(rs("LastInstalldate").value), Date, rs("LastInstalldate").value)
    txtInstalldateH.value = IIf(IsNull(rs("InstalldateH").value), "", rs("InstalldateH").value)


TxtContractDays.text = IIf(IsNull(rs("ContractDays").value), 0, rs("ContractDays").value)
TxtActualDays.text = IIf(IsNull(rs("ActualDays").value), 0, rs("ActualDays").value)
TxtWaterPrice.text = IIf(IsNull(rs("WaterPrice").value), 0, rs("WaterPrice").value)
TxtDayPricen.text = IIf(IsNull(rs("DayPricen").value), 0, rs("DayPricen").value)
 
txtServicePrice.text = IIf(IsNull(rs("ServicePrice").value), 0, rs("ServicePrice").value)
TxtWaterPriceotal.text = IIf(IsNull(rs("WaterPriceotal").value), 0, rs("WaterPriceotal").value)
TxtDayPricentotal.text = IIf(IsNull(rs("DayPricentotal").value), 0, rs("DayPricentotal").value)
TxtService.text = IIf(IsNull(rs("Service").value), 0, rs("Service").value)
txtWaterPayed.text = IIf(IsNull(rs("WaterPayed").value), 0, rs("WaterPayed").value)
TxtRentValuePayed.text = IIf(IsNull(rs("RentValuePayed").value), 0, rs("RentValuePayed").value)
txtTelandNetPayed.text = IIf(IsNull(rs("TelandNetPayed").value), 0, rs("TelandNetPayed").value)
txtRemainWater.text = IIf(IsNull(rs("RemainWater").value), 0, rs("RemainWater").value)
txtRemainRent.text = IIf(IsNull(rs("RemainRent").value), 0, rs("RemainRent").value)
txtRemainService.text = IIf(IsNull(rs("RemainService").value), 0, rs("RemainService").value)
txtDaysValueIncrease.Tag = ""
txtDaysValueIncomplete.Tag = ""
txtDaysValueIncrease.text = IIf(IsNull(rs("DaysValueIncrease").value), 0, rs("DaysValueIncrease").value)
txtDaysValueIncomplete.text = IIf(IsNull(rs("DaysValueIncomplete").value), 0, rs("DaysValueIncomplete").value)

txtDayValueInc.text = IIf(IsNull(rs("DayValueInc").value), 0, rs("DayValueInc").value)
txtDayCountInc.text = IIf(IsNull(rs("DayCountInc").value), 0, rs("DayCountInc").value)
txtDayValueIncomplete.text = IIf(IsNull(rs("DayValueIncomplete").value), 0, rs("DayValueIncomplete").value)
txtDayCountIncomplete.text = IIf(IsNull(rs("DayCountIncomplete").value), 0, rs("DayCountIncomplete").value)




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

   If IsNull(rs.Fields("TypeMonthCalc").value) Then
 chkTypeMonthCalc.value = vbUnchecked
 Else
 chkTypeMonthCalc.value = vbChecked
   
 End If

'*******************************************************************************************

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)

   
    Set RsDetails = New ADODB.Recordset

StrSQL = "SELECT     dbo.TblFiterWaiverDe.IDFItWaiv, dbo.TblFiterWaiverDe.[Count], dbo.TblFiterWaiverDe.Remark, dbo.TblFiterWaiverDe.Price, "
 StrSQL = StrSQL & "                     dbo.TblAqrCompenetDet.Name AS nameDet, dbo.TblFiterWaiverDe.IDItem, dbo.TblAqrCompenet.Name, dbo.TblFiterWaiverDe.GroupID"
StrSQL = StrSQL & ",TblFiterWaiverDe.Accountsus   FROM         dbo.TblFiterWaiverDe LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblAqrCompenetDet ON dbo.TblFiterWaiverDe.IDItem = dbo.TblAqrCompenetDet.ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblAqrCompenet ON dbo.TblFiterWaiverDe.GroupID = dbo.TblAqrCompenet.ID"
StrSQL = StrSQL & "  Where (dbo.TblFiterWaiverDe.IDFItWaiv = " & val(XPTxtID.text) & ")"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
 
   Dim Temp, k, j As Integer
j = 0
Temp = -1
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.FG
      '  RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount
k = 0
        For i = .FixedRows To .rows - 1
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
                      .cell(flexcpBackColor, k, 1, k, 7) = &H80C0FF
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

    
    StrSQL = "Select *,IsNull(TypeDate,1) TypeDate  From TblFiterWaiverDet2 Where MasterID=" & val(Me.XPTxtID.text)
    loadgrid StrSQL, grd, True, True
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

Public Function DateDiffH(ByVal mInterval As String, ByVal mDate1 As String, ByVal mDate2 As String) As Double
Dim mDateDiff1 As Date
Dim mDateDiff2 As Date

mDateDiff1 = ToGregorianDate(mDate1)
mDateDiff2 = ToGregorianDate(mDate2)
If chkTypeMonthCalc.value = vbChecked Then
    DateDiffH = Days360(mDate1, mDate2)
Else
    DateDiffH = (DateDiff("d", mDateDiff1, mDateDiff2))
End If

End Function

Public Function Days360(ByVal StartDate As Date, ByVal EndDate As Date, Optional ByVal Method As Boolean = False) As Long
 
    Dim lMonths As Long
    Dim lStartDay As Long
    Dim lEndDay As Long
    Dim FebruaryAdjustment As Long
 
    lStartDay = day(StartDate)
    lEndDay = day(EndDate)
    
    If Not Method Then
    
        If lStartDay > 30 Then
            StartDate = DateAdd("d", -1, StartDate)
        End If
        
        If (lEndDay = 31) And (lStartDay < 30) Then
            EndDate = DateAdd("d", 1, EndDate)
        ElseIf (lEndDay = 31) And (lStartDay >= 30) Then
            EndDate = DateAdd("d", -1, EndDate)
        End If
        
        
        If IsLastDayInFebruary(StartDate) Then
            FebruaryAdjustment = 30 - day(StartDate)
        End If
        
    Else

        If lStartDay > 30 Then
            StartDate = DateAdd("d", -1, StartDate)
        End If
        
        If lEndDay > 30 Then
            EndDate = DateAdd("d", -1, EndDate)
        End If
 
    End If
    
    lStartDay = day(StartDate)
    lEndDay = day(EndDate)
    
    lMonths = DateDiff("M", StartDate, EndDate)
    
    Days360 = (lMonths * 30) + (lEndDay - lStartDay) - FebruaryAdjustment
 
 
End Function

Private Function IsLastDayInFebruary(ByVal dt As Date) As Boolean
    Dim tmpDate As Date
    tmpDate = DateAdd("d", 1, dt)
    If day(tmpDate) = 1 And Month(tmpDate) = 3 Then
        IsLastDayInFebruary = True
    Else
        IsLastDayInFebruary = False
    End If
End Function

Public Function DateDiffH2(ByVal mInterval As String, ByVal mDate1 As String, ByVal mDate2 As String) As Double
Dim mDateDiff1 As Date
Dim mDateDiff2 As Date
Dim mYear1 As Long
Dim mYear2 As Long
Dim mDay1 As Long
Dim mDay2 As Long
Dim mMonthes As Long

mDay1 = 30 - day(mDate1)
mDay2 = day(mDate2)

If DateDiff("m", mDate1, mDate2) > 1 Then
    mMonthes = DateDiff("m", mDate1, mDate2) - 1
ElseIf DateDiff("m", mDate1, mDate2) >= 2 Then
    mMonthes = 0
ElseIf DateDiff("m", mDate1, mDate2) < 0 Then
    mMonthes = DateDiff("m", mDate1, mDate2)
    
    DateDiffH2 = (mMonthes * 30) + (mDay1 + mDay2)
    Exit Function
End If
If Month(mDate1) = Month(mDate2) And year(mDate1) = year(mDate2) Then
    DateDiffH2 = day(mDate2) - day(mDate1)
    Exit Function
End If

'DateDiffH2 = (DateDiff("d", mDateDiff1, mDateDiff2))
DateDiffH2 = (mMonthes * 30) + mDay1 + mDay2

End Function
Function kh_count_day(Mydate_Max As Date, Mydate_Min As Date)

If IsDate(Mydate_Max) And CDate(Mydate_Min) Then

    kh_count_day = Mydate_Max - Mydate_Min

End If

End Function
Public Sub GetContract(ByVal mContractNo As String, Optional ByVal mTransID As Long = 0)
  On Error Resume Next
    Dim mCustId As Long
    Dim mUnitId As Long
    Dim mIqar As Long
    Dim mUnittype As Long
    
    mCustId = val(dcCustomer2.BoundText)
    mUnitId = val(DcbUnitNo2.BoundText)
    mIqar = val(DcbIqara2.BoundText)
    mUnittype = val(DcbUnitType2.BoundText)
    
    If SystemOptions.WaiverSetByContract And mContractNo <> 0 Then
        mCustId = 0: mUnitId = 0: mIqar = 0: mUnittype = 0
    End If
    If val(mContractNo) = 0 And mCustId = 0 And val(mTransID) = 0 Then Exit Sub
 
    Dim s As String, mCount As Long
    Dim rsDummyCount  As New ADODB.Recordset
    If mContractNo = 0 And mTransID = 0 And mIqar = 0 Then Exit Sub
    If val(mContractNo) <> 0 And mCustId = 0 Then
    
        s = "Select Count(*) CC from TblContract Where "
        If Me.TxtModFlg.text = "R" Then
            s = s & "    (dbo.TblContract.EndContract IS NULL)      and "
        End If
        s = s & "    (    (NoteSerial1 = '" & Trim(mContractNo) & "' Or ContNo = " & val(mTransID) & " )   and ( 1 =1  "
        
        
        If Not SystemOptions.WaiverSetByContract Or mContractNo = 0 Then
            If mUnittype <> 0 Then
                s = s & " And  unittype = " & val(mUnittype)
            End If
    
            If mIqar <> 0 Then
                s = s & " And  Iqar = " & val(mIqar)
            End If
    
            If mUnitId <> 0 Then
                s = s & " And UnitNo = " & val(DcbUnitNo2.BoundText)
            End If
    
            If mCustId <> 0 Then
                s = s & " And CusID = " & val(dcCustomer2.BoundText)
            
            End If
        End If
        s = s & "Or 1 =-1)"
        s = s & " )"
        If mUnittype = 0 And mIqar = 0 Then
            s = "Select Count(*) CC from TblContract Where  "
            If Me.TxtModFlg.text = "R" Then
                s = s & "    (dbo.TblContract.EndContract IS NULL)      and "
            End If
 

        
            s = s & " (NoteSerial1 = '" & Trim(mContractNo) & "'      )    "
        End If
    Else
        s = "Select Count(*) CC from TblContract Where "
        s = s & " (( 1 =1  "
        
        If Not SystemOptions.WaiverSetByContract Or mContractNo = 0 Then
            If mUnittype <> 0 Then
                s = s & " And  unittype = " & val(mUnittype)
            End If
    
            If mIqar <> 0 Then
                s = s & " And  Iqar = " & val(mIqar)
            End If
    
            If mUnitId <> 0 Then
                s = s & " And UnitNo = " & val(DcbUnitNo2.BoundText)
            End If
    
            If mCustId <> 0 Then
                s = s & " And CusID = " & val(dcCustomer2.BoundText)
            
            End If
        End If
        s = s & "Or 1 =-1)"
        s = s & " )"
        If mUnittype = 0 And mIqar = 0 Then
                s = "Select Count(*) CC from TblContract Where "
                If Me.TxtModFlg.text = "R" Then
                    s = s & "    (dbo.TblContract.EndContract IS NULL)      and "
                End If

            
            
            s = s & "    ContNo = " & mTransID
        End If
        
        
    End If
    
    rsDummyCount.Open s, Cn, adOpenKeyset, adLockReadOnly
    If Not rsDummyCount.EOF Then
        mCount = val(rsDummyCount!CC & "")
    End If
   Dim rsDummy  As New ADODB.Recordset
   If mContractNo <> 0 And mCustId = 0 Then
        s = "Select OldInsurance From TblContract Where "
        
         If Me.TxtModFlg.text = "R" Then
            s = s & "    (dbo.TblContract.EndContract IS NULL)      and "
        End If
          
        s = s & "    ((NoteSerial1 = '" & Trim(mContractNo) & "' Or ContNo = " & val(mTransID) & " )  and ( 1 =1  "
        If Not SystemOptions.WaiverSetByContract Or mContractNo = 0 Then
            If mUnittype <> 0 Then
                s = s & " And  unittype = " & val(mUnittype)
            End If
    
            If mIqar <> 0 Then
                s = s & " And  Iqar = " & val(mIqar)
            End If
    
            If mUnitId <> 0 Then
                s = s & " And UnitNo = " & val(DcbUnitNo2.BoundText)
            End If
    
            If mCustId <> 0 Then
                s = s & " And CusID = " & val(dcCustomer2.BoundText)
            
            End If
        End If
        s = s & "Or 1 =-1)"
        s = s & " )"
        
        If mUnittype = 0 And mIqar = 0 Then
                s = "Select OldInsurance From TblContract Where "
                If Me.TxtModFlg.text = "R" Then
                    s = s & "    (dbo.TblContract.EndContract IS NULL)      and "
                End If
            
            's = s & "     NoteSerial1 = '" & Trim(mContractNo) & "'"
           s = s & "         (NoteSerial1 = '" & Trim(mContractNo) & "' Or ContNo = " & val(mTransID) & " )"
        End If
        s = s & " And IsNull(OldInsurance,0) <> 0  Order By ContNo Desc"
    Else
        s = "Select OldInsurance From TblContract Where "
        s = s & " (ContNo = " & mTransID & "   and ( 1 =1  "
        If Not SystemOptions.WaiverSetByContract Or mContractNo = 0 Then
            If mUnittype <> 0 Then
                s = s & " And  unittype = " & val(mUnittype)
            End If
    
            If mIqar <> 0 Then
                s = s & " And  Iqar = " & val(mIqar)
            End If
    
            If mUnitId <> 0 Then
                s = s & " And UnitNo = " & val(DcbUnitNo2.BoundText)
            End If
    
            If mCustId <> 0 Then
                s = s & " And CusID = " & val(dcCustomer2.BoundText)
            
            End If
        End If
        s = s & "Or 1 =-1)"
        s = s & " )"
        
        If mUnittype = 0 And mIqar = 0 Then
            s = "Select OldInsurance From TblContract Where "
            s = s & " ContNo = " & mTransID & " "
        End If
        s = s & " And IsNull(OldInsurance,0) <> 0  Order By ContNo Desc"
    End If
   rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
   txtOldInsurance = ""
   If Not rsDummy.EOF Then
        txtOldInsurance = IIf(IsNull(rsDummy("OldInsurance").value), "", rsDummy("OldInsurance").value)
   End If
   
  

    
    
    
    Set rsDummy = New ADODB.Recordset
    If mContractNo <> 0 And mCustId = 0 Then
        s = "Select * from TblContract Where "
        If Me.TxtModFlg.text = "R" Then
            s = s & "    (dbo.TblContract.EndContract IS NULL)      and "
        End If
            
        
    '    s = s & " (NoteSerial1 = '" & Trim(mContractNo) & "'  " & " and ( 1 =1  "
         s = s & "         (NoteSerial1 = '" & Trim(mContractNo) & "' Or ContNo = " & val(mTransID) & " )  and ( 1 =1  "
        If Not SystemOptions.WaiverSetByContract Or mContractNo = 0 Then
            If mUnittype <> 0 Then
                s = s & " And  unittype = " & val(mUnittype)
            End If
    
            If mIqar <> 0 Then
                s = s & " And  Iqar = " & val(mIqar)
            End If
    
            If mUnitId <> 0 Then
                s = s & " And UnitNo = " & val(DcbUnitNo2.BoundText)
            End If
    
            If mCustId <> 0 Then
                s = s & " And CusID = " & val(dcCustomer2.BoundText)
            
            End If
        End If
        s = s & "Or 1 =-1"
        s = s & " )"
        If mUnittype = 0 And mIqar = 0 Then
            s = "Select * from TblContract Where "
            If Me.TxtModFlg.text = "R" Then
                    s = s & "    (dbo.TblContract.EndContract IS NULL)      and "
            End If
            
           ' s = s & " NoteSerial1 = '" & Trim(mContractNo) & "'  "
            s = s & "         (NoteSerial1 = '" & Trim(mContractNo) & "' Or ContNo = " & val(mTransID) & " ) "
        End If
        s = s & " Order By ContNo "
    Else
        s = "Select * from TblContract Where "
        If Me.TxtModFlg.text = "R" Then
                    s = s & "    (dbo.TblContract.EndContract IS NULL)      and "
        End If
            
        's = s & " (ContNo = " & mTransID & "   and ( 1 =1  "
        s = s & " ( ( 1 =1  "
        If Not SystemOptions.WaiverSetByContract Or mContractNo = 0 Then
            If mUnittype <> 0 Then
                s = s & " And  unittype = " & val(mUnittype)
            End If
    
            If mIqar <> 0 Then
                s = s & " And  Iqar = " & val(mIqar)
            End If
    
            If mUnitId <> 0 Then
                s = s & " And UnitNo = " & val(DcbUnitNo2.BoundText)
            End If
    
            If mCustId <> 0 Then
                s = s & " And CusID = " & val(dcCustomer2.BoundText)
            
            End If
        End If
        s = s & " Or 1 =-1)"
        s = s & " )"
        If mUnittype = 0 And mIqar = 0 Then
            s = "Select * from TblContract Where "
            If Me.TxtModFlg.text = "R" Then
                    s = s & "    (dbo.TblContract.EndContract IS NULL)      and "
                End If
            
            s = s & " and  ContNo = " & mTransID
        End If
        s = s & " Order By ContNo "
    End If
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    grd.rows = 1
 
    
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
    Dim mTotalLastDays As Double
    Dim OldInsurance As Double
    Dim mTotalContract As Double
    txtTotal1 = 0
    txtTotal1 = 0
    txtWaterPayed = ""
    TxtDayLate = ""
    TxtActualDays = ""
    txtTelandNetPayed = ""
    TxtRentValuePayed = ""
    TxtDayPricentotal = ""
    txtInsurance = ""
    txtRemainService = ""
    TxtBillPrice = ""
    txtRemainWater = ""
    txtRemainRent = ""
    Dim mDateConvert As Date
    txtLastInstalldate.Visible = True
    Do While Not rsDummy.EOF
        If rsDummy("TypeDate").value = 0 Then
            If ChkCalcLastPayment.value = vbUnchecked Then
                mRemainDays = (DateDiffH("d", FilterDateH.value, rsDummy!todateH & ""))
            Else
                mRemainDays = (DateDiffH("d", FilterDateH.value, txtInstalldateH.value))
            End If
             'mDateConvert = ToGregorianDate(rsDummy!toDateH & "")
           '  mRemainDays = (DateDiff("d", FilterDate.value, mDateConvert))
         '    DateDiffH22
            If rsDummy!fromdateH = FilterDateH.value Then
                mActualDays = 0
            Else
         '   mActualDays = kh_count_day(FilterDateH.value, rsDummy!FromdateH & "")
                mActualDays = (DateDiffH("d", rsDummy!fromdateH & "", FilterDateH.value))
               
                
            End If
            
            If rsDummy!fromdateH = rsDummy!todateH Then
                mContractDays = 1
            Else
                
           
                    mContractDays = (DateDiffH("d", rsDummy!fromdateH & "", rsDummy!todateH & ""))
                
                
            End If
            
            
        Else
            If ChkCalcLastPayment.value = vbUnchecked Then
                mRemainDays = (DateDiff("d", FilterDate.value, rsDummy!EndDate & ""))

            Else
                mRemainDays = (DateDiff("d", FilterDate.value, txtLastInstalldate.value))
            End If
            If rsDummy!StrDate = FilterDate.value Then
                mActualDays = 0
            Else
                mActualDays = (DateDiff("d", rsDummy!StrDate & "", FilterDate.value))
            End If
            
          
        End If
        If rsDummy!StrDate = rsDummy!EndDate Then
            mContractDays = 0
        Else
            mContractDays = (DateDiff("d", rsDummy!StrDate, rsDummy!EndDate))
        End If
        If val(mContractDays) <> 0 Then
           'mContractDays = mContractDays + 1
            mDayPricen = Round(IIf(IsNull(rsDummy("TotalContract").value), 0, rsDummy("TotalContract").value) / val(mContractDays), 2)
        End If
        
        mDayPricentotal = val(mDayPricen) * val(mActualDays)
        
        If ChkCalcLastPayment.value = vbUnchecked Then
            payed = getinsttPayedTocontract2(val(rsDummy!ContNo & ""), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed, TotalOldValue, , , VATPayed)
        Else
            Dim mTypeDate As Boolean
            mTypeDate = IIf(RdRTypeDate(0), True, False)
            If txtLastInstalldate.Visible = True Then
                If Not mTypeDate Then
                    payed = getinsttPayedTocontract2(IIf(IsNull(rsDummy("ContNo").value), 0, val(rsDummy!ContNo & "")), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed, TotalOldValue, , , VATPayed, CStr(txtLastInstalldate.value), mTypeDate)
                Else
                    payed = getinsttPayedTocontract2(IIf(IsNull(rsDummy("ContNo").value), 0, val(rsDummy!ContNo & "")), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed, TotalOldValue, , , VATPayed, CStr(txtInstalldateH.value), mTypeDate)
                End If
            Else
                If Not mTypeDate Then
                    payed = getinsttPayedTocontract2(IIf(IsNull(rsDummy("ContNo").value), 0, val(rsDummy!ContNo & "")), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed, TotalOldValue, , , VATPayed, CStr(FilterDate.value), mTypeDate)
                Else
                    payed = getinsttPayedTocontract2(IIf(IsNull(rsDummy("ContNo").value), 0, val(rsDummy!ContNo & "")), RentValuePayed, CommissionsPayed, InsurancePayed, WaterPayed, ElectricPayed, TelandNetPayed, TotalOldValue, , , VATPayed, CStr(FilterDateH.value), mTypeDate)
                End If
            End If
        End If

        mRentValuePayed = RentValuePayed
        mRemainRent = RentValuePayed
        mRemainWater = WaterPayed
        RemainCommissions = CommissionsPayed
        mRemElictricPrice = ElectricPayed
       
        mRemainService = TelandNetPayed
        mDayLate = val(mContractDays) - val(mActualDays)
        mAmountDely = Round(val(mWaterPrice) * val(mDayLate), 2) + Round(val(mDayPricen) * val(mDayLate), 2) + Round(val(mServicePrice) * val(mDayLate), 2)
   
    
     
        
        With grd
            .AddItem 1
             If mCount = .rows - 1 Then
               
                If val(mContractDays) <> 0 Then
                    mTotalContract = Round(val(rsDummy!TotalContract & "") + val(rsDummy!phone & "") + val(rsDummy!Electricity & "") + val(rsDummy!Water & ""), 2)
                    mDaysValue = Round((mTotalContract / val(mContractDays)) * Abs(mRemainDays), 2)
                   
                    mDaysValue = Round((mTotalContract / val(mContractDays)) * mRemainDays, 2)
                End If
             
             End If
            mTotalDept = mRemainRent + mRemainWater + mRemElictricPrice + RemainCommissions + mRemainService + val(TotalOldValue) + IIf(mDaysValue > 0, val(mDaysValue), 0)
            mTotalRight = val(val(rsDummy!InsuranceValue & "") + IIf(mDaysValue < 0, Abs(val(mDaysValue)), 0))
            DaysValueIncrease = DaysValueIncrease + IIf(mDaysValue < 0, Abs(val(mDaysValue)), 0)
            DaysValueIncomplete = DaysValueIncomplete + IIf(mDaysValue > 0, Abs(val(mDaysValue)), 0)
            
            
            mTotalDept2 = mTotalDept2 + mRemainRent + mRemainWater + RemainCommissions + mRemElictricPrice + mRemainService + val(TotalOldValue)
            mTotalRight2 = mTotalRight2 + val(val(rsDummy!InsuranceValue & ""))
            
            .TextMatrix(.rows - 1, .ColIndex("TypeDate")) = rsDummy!TypeDate & ""
            .TextMatrix(.rows - 1, .ColIndex("ContNo")) = rsDummy!ContNo & ""
            .TextMatrix(.rows - 1, .ColIndex("StartDate")) = rsDummy!StrDate & ""
            .TextMatrix(.rows - 1, .ColIndex("StartDateh")) = rsDummy!fromdateH & ""
            .TextMatrix(.rows - 1, .ColIndex("EndDate")) = rsDummy!EndDate & ""
            .TextMatrix(.rows - 1, .ColIndex("EndDateH")) = rsDummy!todateH & ""
            .TextMatrix(.rows - 1, .ColIndex("RemainRent")) = mRemainRent
            .TextMatrix(.rows - 1, .ColIndex("RemainWater")) = mRemainWater
            .TextMatrix(.rows - 1, .ColIndex("BillPrice")) = mRemElictricPrice
            .TextMatrix(.rows - 1, .ColIndex("RemainService")) = mRemainService
            .TextMatrix(.rows - 1, .ColIndex("insurance")) = InsurancePayed
            '.TextMatrix(.Rows - 1, .ColIndex("DaysValue")) = rsDummy!InsuranceValue & ""
            .TextMatrix(.rows - 1, .ColIndex("RemainDays")) = mRemainDays
            .TextMatrix(.rows - 1, .ColIndex("DaysValue")) = mDaysValue
            .TextMatrix(.rows - 1, .ColIndex("RemainCommissions")) = RemainCommissions
            
            .TextMatrix(.rows - 1, .ColIndex("OldRent")) = TotalOldValue
            '.TextMatrix(.Rows - 1, .ColIndex("DaysValue")) = mDayPricentotal
            .TextMatrix(.rows - 1, .ColIndex("RentValuePayed")) = RentValuePayed
            .TextMatrix(.rows - 1, .ColIndex("WaterPayed")) = WaterPayed
            .TextMatrix(.rows - 1, .ColIndex("TelandNetPayed")) = TelandNetPayed
            .TextMatrix(.rows - 1, .ColIndex("ActualDays")) = mContractDays
            .TextMatrix(.rows - 1, .ColIndex("DayLate")) = mDayLate
            .TextMatrix(.rows - 1, .ColIndex("AmountDely")) = mAmountDely
            .TextMatrix(.rows - 1, .ColIndex("TotalDept")) = mTotalDept
            .TextMatrix(.rows - 1, .ColIndex("TotalRight")) = mTotalRight
            .TextMatrix(.rows - 1, .ColIndex("TotalStill")) = RentValuePayed + CommissionsPayed + WaterPayed + ElectricPayed + TelandNetPayed + TotalOldValue
'
'            If val(mContractDays) <> 0 Then
'               'mContractDays = mContractDays + 1
'
'                mDayPricen = Round(IIf(IsNull(rsDummy("TotalContract").value), 0, rsDummy("TotalContract").value) / val(mContractDays), 2)
'                mDayPricen = Round(Grd.TextMatrix(Grd.Rows - 1, Grd.ColIndex("TotalStill")) / val(mContractDays), 2)
'            End If
'            If ChkCalcLastPayment.value = vbChecked Then
'                mDayPricentotal = val(mDayPricen) * val(mActualDays)
'            End If

      
            
            
            txtWaterPayed = val(txtWaterPayed) + WaterPayed
            TxtDayLate = val(TxtDayLate) + WaterPayed
            TxtActualDays = val(TxtActualDays) + mActualDays
            txtTelandNetPayed = val(txtTelandNetPayed) + TelandNetPayed
            TxtRentValuePayed = val(TxtRentValuePayed) + RentValuePayed
            TxtDayPricentotal = val(TxtDayPricentotal) + mDayPricentotal
            txtInsurance = InsurancePayed
            txtRemainService = val(txtRemainService) + mRemainService
            TxtBillPrice = val(TxtBillPrice) + mRemElictricPrice
            txtRemainWater = val(txtRemainWater) + mRemainWater
            txtRemainRent = val(txtRemainRent) + mRemainRent
            If .rows - 1 = mCount Then
                If rsDummy("TypeDate").value = 0 Then
                    mTotalLastDays = (DateDiffH("d", .TextMatrix(.rows - 1, .ColIndex("StartDateh")), FilterDateH.value))
                Else
                    mTotalLastDays = (DateDiff("d", .TextMatrix(.rows - 1, .ColIndex("StartDate")), FilterDate.value))
                End If
                mTotalLastDays = mTotalLastDays + 1
                
            End If
        End With

        
        rsDummy.MoveNext
    Loop
        With grd
           .rows = .rows + 1
            .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
            .IsSubtotal(.rows - 1) = True
            Dim SngTotal As Single
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainRent"), .rows - 1, .ColIndex("RemainRent"))
            .TextMatrix(.rows - 1, .ColIndex("RemainRent")) = SngTotal
        
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainWater"), .rows - 1, .ColIndex("RemainWater"))
            .TextMatrix(.rows - 1, .ColIndex("RemainWater")) = SngTotal
            
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillPrice"), .rows - 1, .ColIndex("BillPrice"))
                    .TextMatrix(.rows - 1, .ColIndex("BillPrice")) = SngTotal
                    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainService"), .rows - 1, .ColIndex("RemainService"))
                    .TextMatrix(.rows - 1, .ColIndex("RemainService")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TelandNetPayed"), .rows - 1, .ColIndex("TelandNetPayed"))
                    .TextMatrix(.rows - 1, .ColIndex("TelandNetPayed")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainCommissions"), .rows - 1, .ColIndex("RemainCommissions"))
                    .TextMatrix(.rows - 1, .ColIndex("RemainCommissions")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalStill"), .rows - 1, .ColIndex("TotalStill"))
                    .TextMatrix(.rows - 1, .ColIndex("TotalStill")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OldRent"), .rows - 1, .ColIndex("OldRent"))
                    .TextMatrix(.rows - 1, .ColIndex("OldRent")) = SngTotal
    .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
        .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
        .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
        .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
     End With
    txtDaysValueIncrease.Tag = Round(DaysValueIncrease, 2)
    txtDaysValueIncomplete.Tag = Round(DaysValueIncomplete, 2)
    txtTotalLastDays = mTotalLastDays
    txtDaysValueIncrease = Round(DaysValueIncrease, 2)
    txtDaysValueIncomplete = Round(DaysValueIncomplete, 2)
    txtDayValueInc = Round(IIf(mDaysValue < 0, Abs(val(mDaysValue)), 0) / IIf(Abs(mRemainDays), Abs(mRemainDays), 1), 2)
    txtDayValueIncomplete = Round(IIf(mDaysValue > 0, Abs(val(mDaysValue)), 0) / IIf(Abs(mRemainDays), Abs(mRemainDays), 1), 2)
    txtTotalLastDays = mTotalLastDays
    txtDayCountIncomplete = Round(IIf(mRemainDays > 0, Abs(val(mRemainDays)), 0), 2)
    txtDayCountInc = Round(IIf(mRemainDays < 0, Abs(val(mRemainDays)), 0), 2)
    txtTotal1 = val(mTotalDept2) + val(txtDaysValueIncrease)
    txtTotal2 = val(mTotalRight2) + val(txtDaysValueIncomplete)
    CalcTotal
    ReLineGrid
End Sub
Private Sub calctextTotal()
txtDaysValueIncomplete = val(txtDayCountIncomplete) * val(txtDayValueIncomplete)
txtDaysValueIncrease = val(txtDayCountInc) * val(txtDayValueInc)
End Sub
Private Sub CalcTotal()
  
Dim mTotalLastDays As Double
  txtTotalinsuranceS = 0
  With grd
           
           
           If .IsSubtotal(.rows - 1) = True Then
                .RemoveItem .rows - 1
           End If
           
            Dim i As Long
            For i = 1 To grd.rows - 1
                If .TextMatrix(i, .ColIndex("StartDate")) <> "" Then
                    
                    If val(.TextMatrix(i, .ColIndex("TypeDate"))) = 0 Then
                         mTotalLastDays = (DateDiffH("d", .TextMatrix(i, .ColIndex("StartDateh")), FilterDateH.value))
                    Else
                        mTotalLastDays = (DateDiff("d", .TextMatrix(i, .ColIndex("StartDate")), FilterDate.value))
                    End If
                End If
              
                   txtTotalinsuranceS = val(txtTotalinsuranceS) + val(.TextMatrix(i, .ColIndex("insurance")))
            Next
           
                .rows = .rows + 1
                .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
                .IsSubtotal(.rows - 1) = True
            
            
            
            
            Dim SngTotal As Single
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainRent"), .rows - 1, .ColIndex("RemainRent"))
            .TextMatrix(.rows - 1, .ColIndex("RemainRent")) = SngTotal
        
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainWater"), .rows - 1, .ColIndex("RemainWater"))
            .TextMatrix(.rows - 1, .ColIndex("RemainWater")) = SngTotal
            
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillPrice"), .rows - 1, .ColIndex("BillPrice"))
                    .TextMatrix(.rows - 1, .ColIndex("BillPrice")) = SngTotal
                    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainService"), .rows - 1, .ColIndex("RemainService"))
                    .TextMatrix(.rows - 1, .ColIndex("RemainService")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TelandNetPayed"), .rows - 1, .ColIndex("TelandNetPayed"))
                    .TextMatrix(.rows - 1, .ColIndex("TelandNetPayed")) = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("RemainCommissions"), .rows - 1, .ColIndex("RemainCommissions"))
                    .TextMatrix(.rows - 1, .ColIndex("RemainCommissions")) = SngTotal
            
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("insurance"), .rows - 1, .ColIndex("insurance"))
                    .TextMatrix(.rows - 1, .ColIndex("insurance")) = SngTotal
            
            If SngTotal <> 0 Then txtOldInsurance = 0
                    
               txtTotal2 = SngTotal + val(txtDaysValueIncomplete)
          
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OldRent"), .rows - 1, .ColIndex("OldRent"))
                    .TextMatrix(.rows - 1, .ColIndex("OldRent")) = SngTotal
                    
  SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalStill"), .rows - 1, .ColIndex("TotalStill"))
                    .TextMatrix(.rows - 1, .ColIndex("TotalStill")) = SngTotal
          txtTotal1 = SngTotal + val(txtDaysValueIncrease)
          
    .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
        .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
        .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
        .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
     End With
    txtTotalLastDays = mTotalLastDays
    If val(txtTotalinsuranceS) = 0 Then
        txtTotalinsuranceS = val(txtTotalinsuranceS) + val(txtOldInsurance)
    Else
        txtTotalinsuranceS = val(txtTotalinsuranceS)
    End If
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


    If Me.TxtModFlg.text <> "R" Then
        If Me.dcCustomer.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ   «”„ «·„” «Ã—!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
          '  Me.dcCustomer.SetFocus
           ' SendKeys "{F4}"
            Exit Sub
        End If
   If Me.DcbIqara.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ   «”„ «·⁄„«—Â!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.DcbIqara.SetFocus
           ' SendKeys "{F4}"
            Exit Sub
        End If
   
        If Trim(TxtManualNO.text) = "" Then
            TxtManualNO.text = XPTxtID
            
        End If
        

        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("TblFiterWaiver", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
          StrSQL = "Delete From TblUnitNoInformation Where FilterNo=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From TblFiterWaiverDe Where IDFItWaiv=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
               StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
         StrSQL = "Delete From Notes Where FilterID2 =" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords


        End If
       rs("ID").value = val(XPTxtID.text)
           rs("ContNo").value = val(TxtContNo.text)
             rs("ContractNo").value = (TxtOrder.text)
        rs("ManualNo") = Trim(TxtManualNO.text)
       rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
       rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)
       rs("RenterID").value = IIf(Me.dcCustomer.BoundText = "", Null, Me.dcCustomer.BoundText)
       rs("RenterID2").value = IIf(Me.dcCustomer2.BoundText = "", Null, Me.dcCustomer2.BoundText)
       rs("BulidID").value = IIf(Me.DcbIqara.BoundText = "", Null, Me.DcbIqara.BoundText)
       rs("unittype").value = IIf(Me.DcbUnitType.BoundText = "", Null, Me.DcbUnitType.BoundText)
       rs("ApartmentID").value = IIf(Me.DcbUnitNo.BoundText = "", Null, Me.DcbUnitNo.BoundText)
       rs("ApartmentID2").value = IIf(Me.DcbUnitNo2.BoundText = "", Null, Me.DcbUnitNo2.BoundText)
       
       rs("unittype2").value = IIf(Me.DcbUnitType2.BoundText = "", Null, Me.DcbUnitType2.BoundText)
       rs("BulidID2").value = IIf(Me.DcbIqara2.BoundText = "", Null, Me.DcbIqara2.BoundText)
       
      
    
   
     
        rs.Fields("UnitElectric").value = IIf(Me.TxtElectAccount.text <> "", val(TxtElectAccount.text), Null)
            
        rs("Accredit").value = IIf(ChkAccredit.value = vbUnchecked, Null, 1)
        rs("NewNO").value = (TXTNewNO)
    
       
       
       rs("RecordDate").value = XPDtbTrans.value
       rs("RecordDateH").value = Me.NourHijriCal1.value
       rs("Insurance").value = val(Me.txtOldInsurance.text)
       rs("net").value = val(Me.txtNet.text)
       rs("ForRenter").value = val(Me.TxtForRenter.text)
       
       rs("Discount").value = val(Me.txtDiscount.text)
       rs("ForRenterB").value = val(Me.TxtForRenterB.text)
       
       
       rs("OFRenter").value = val(Me.TxtOFRenter.text)
       rs("TotalinsuranceS").value = val(Me.txtTotalinsuranceS.text)
       
      ''
       rs("EndDate").value = EndDate.value
       rs("EndDateH").value = Me.EndDateH.value
       rs("FilterDate").value = FilterDate.value
       rs("FilterDateH").value = Me.FilterDateH.value
       rs("BillPrice").value = val(Me.TxtBillPrice.text)
       rs("AccountNo").value = Me.TxtAccountNo.text
       rs("DayNo").value = val(Me.TxtDayLate.text)
       rs("AmountDely").value = val(Me.TxtAmountDely.text)
       rs("Insurance").value = val(Me.txtOldInsurance.text)
        If RdRTypeDate(1).value = True Then
            rs("TypeDate").value = 1
        Else
            rs("TypeDate").value = 0
        End If

        If ComResid(1).value = True Then
            rs.Fields("ComResid").value = 1
        Else
            rs.Fields("ComResid").value = 0
        End If
        If ChkCalcLastPayment.value = vbChecked Then
            rs.Fields("CalcLastPayment").value = 1
        Else
            rs.Fields("CalcLastPayment").value = 0
        End If
        
       
        rs("LastInstalldate").value = txtLastInstalldate.value
        rs("InstalldateH").value = Me.txtInstalldateH.value
        
        
        rs("VAtPercent").value = val(Me.TxtVAtPercent.text)
         rs("VAt2").value = val(Me.txtVat2.text)
        rs("TotalCounterNet").value = val(Me.txtTotalCounterNet.text)
          
    '***************************************************************************
   rs("ContractDays").value = val(Me.TxtContractDays.text)
   rs("ActualDays").value = val(Me.TxtActualDays.text)
rs("WaterPrice").value = val(Me.TxtWaterPrice.text)
rs("DayPricen").value = val(Me.TxtDayPricen.text)

     rs("LastInvoiceRead").value = val(Me.txtLastInvoiceRead.text)
        rs("LastInvoiceRead2").value = val(Me.txtLastInvoiceRead2.text)
        rs("Diff").value = val(Me.txtDiff.text)
        rs("Price").value = val(Me.TxtPrice.text)
        rs("R").value = val(Me.txtR.text)
        rs("PrevBalance").value = val(Me.txtPrevBalance.text)
        rs("ServiceCounter").value = val(Me.txtServiceCounter.text)
        rs("TotalCounter").value = val(Me.txtTotalCounter.text)




rs("ServicePrice").value = val(Me.txtServicePrice.text)
rs("WaterPriceotal").value = val(Me.TxtWaterPriceotal.text)
rs("DayPricentotal").value = val(Me.TxtDayPricentotal.text)
rs("Service").value = val(Me.TxtService.text)
rs("WaterPayed").value = val(Me.txtWaterPayed.text)
rs("RentValuePayed").value = val(Me.TxtRentValuePayed.text)
rs("TelandNetPayed").value = val(Me.txtTelandNetPayed.text)
rs("RemainWater").value = val(Me.txtRemainWater.text)
rs("RemainRent").value = val(Me.txtRemainRent.text)
rs("RemainService").value = val(Me.txtRemainService.text)
    '***************************************************************************
   rs("outflow").value = IIf(chkoutflow.value = vbUnchecked, Null, 1)
   rs("outCondition").value = IIf(chkoutCondition.value = vbUnchecked, Null, 1)
   rs("TypeMonthCalc").value = IIf(chkTypeMonthCalc.value = vbUnchecked, Null, 1)
   
   rs("DaysValueIncrease").value = val(Me.txtDaysValueIncrease.text)
   rs("DaysValueIncomplete").value = val(Me.txtDaysValueIncomplete.text)

    rs("DayValueInc").value = val(Me.txtDayValueInc.text)
    rs("DayCountInc").value = val(Me.txtDayCountInc.text)
    rs("DayValueIncomplete").value = val(Me.txtDayValueIncomplete.text)
    rs("DayCountIncomplete").value = val(Me.txtDayCountIncomplete.text)
    
    
        rs("LegalAffairs").value = Trim(Me.txtLegalAffairs.text)
        If chkIsLegalAffairs.value = vbChecked Then
            rs.Fields("isLegalAffairs").value = 1
        Else
            rs.Fields("isLegalAffairs").value = 0
        End If
        
        rs.update
        '''''''''/////////////////////////////////
        Dim Temp As Integer
        Temp = -1
      Set RsDetails = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblFiterWaiverDe Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

          
       For i = Me.FG.FixedRows To FG.rows - 1
    
       If FG.TextMatrix(i, FG.ColIndex("group")) <> "" Then
   
           RsDetails.AddNew
        
           If val(FG.TextMatrix(i, FG.ColIndex("iditem"))) = 0 Then
           RsDetails("GroupID").value = val(FG.TextMatrix(i, FG.ColIndex("id")))
            RsDetails("IDItem").value = 0
             RsDetails("IDFItWaiv").value = val(XPTxtID.text)
    
              RsDetails("Count").value = 0
           RsDetails("price").value = 0
          RsDetails("Remark").value = ""
        
          '    temp = val(fg.TextMatrix(i, fg.ColIndex("id")))
           Else
           RsDetails("IDItem").value = val(FG.TextMatrix(i, FG.ColIndex("iditem")))
             RsDetails("GroupID").value = 0
                  RsDetails("IDFItWaiv").value = val(XPTxtID.text)
                           RsDetails("Count").value = val(FG.TextMatrix(i, FG.ColIndex("count")))
           RsDetails("price").value = val(FG.TextMatrix(i, FG.ColIndex("price")))
                    
                RsDetails("Remark").value = FG.TextMatrix(i, FG.ColIndex("remark"))
        RsDetails("Accountsus").value = FG.TextMatrix(i, FG.ColIndex("Accountsus"))
        
        
           End If
         
        
         RsDetails.update
      '  End If
      '   End If
      End If
        Next i

    StrSQL = "Delete From TblFiterWaiverDet2 Where MasterID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    

    StrSQL = "Select Max(ID) MaxID From TblFiterWaiverDet2 "
    Dim rsDummy As New ADODB.Recordset
    Dim mLastIndex As Long
    rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        mLastIndex = val(rsDummy!MaxID & "")
    End If
    StrSQL = "Select *  From TblFiterWaiverDet2 Where MasterID= -1"
     
    
   
 
 
    
    saveGrid StrSQL, grd, "ContNo", "Index" & mLastIndex, "MasterID", val(Me.XPTxtID.text)
 

                
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
'        If SystemOptions.NoCreatJLInRentContract = False And val(TxtNet.text) <> 0 Then
'         createVoucher
'        End If
        
         If val(txtNet.text) <> 0 Then
         createVoucher
        End If
         
        
      updateNotesValueAndNobytext (val(TXTNoteID.text))
      Dim j As Long
      Dim mContID  As Long
            For j = 1 To grd.rows - 1
                mContID = val(grd.TextMatrix(j, grd.ColIndex("ContNo")))
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
             
SaveNotes
        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »«‰«  Â–Â «·⁄„·… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ï ≈÷«›… »«‰«  √Œ—Ï"

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
  Dim StrSQL1 As String
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
           Cn.Execute "  update TblAqarDetai  Set Status = 1   ,customerid=" & val(dcCustomer.BoundText) & "  Where id =" & val(DcbUnitNo.BoundText)
             Cn.Execute "  update TblContract  Set EndContract = null    Where ContNo =" & val(TxtContNo.text)
             
             
                rs.delete
                StrSQL = "Delete From TblFiterWaiver Where ID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
                
                StrSQL = "Delete From TblFiterWaiverDet2 Where MasterID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
 StrSQL1 = "Delete From TblFiterWaiverDe Where IDFItWaiv=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL1, , adExecuteNoRecords
            StrSQL = "Delete From TblUnitNoInformation Where FilterNo=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
                  StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
         StrSQL = "Delete From Notes Where FilterID2 =" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
        StrSQL = "Delete From NOTES Where NoteID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
                If rs.RecordCount < 1 Then
             
            FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 2
            FG.Enabled = True
            
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
        GetIqarCode TxtSearch.text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.text, EmpID
        dcCustomer.BoundText = EmpID
    End If
End Sub

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 1215
        FrmCustemerSearch.show vbModal

    End If
 

If KeyCode = vbKeyF5 Then
'reloadCombos
End If
End Sub

Private Sub dcCustomer_Change()
If Me.TxtModFlg.text <> "R" Then
   If val(dcCustomer.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , val(dcCustomer.BoundText), CStr(EmpCode)
    Me.Text15.text = EmpCode
End If
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
        .Create Me.hWnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’›ÌÂ Ê ‰«“· ⁄‰ «·⁄ﬁœ", 1, 15204351, -2147483630
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


Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.text <> "R" Then
         NourHijriCal1.value = ToHijriDate(XPDtbTrans.value)
End If
End Sub


Function RtriveInfoOrbon(Optional NotID As Double = 0) As Boolean
  
  
End Function

Function CheckJE() As Boolean
Dim i As Integer
CheckJE = False
With GridInstallments
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("NoteId"))) <> 0 Then
CheckJE = True
Exit Function
End If
Next i
End With
End Function

Private Sub txtLastInvoiceRead_Change()
CalcTotalCounter
End Sub
Private Sub CalcTotalCounter()
txtDiff = Round(val(txtLastInvoiceRead2) - val(txtLastInvoiceRead), 2)
txtR = Round(val(TxtPrice) * val(txtDiff), 2)
txtTotalCounter = Round(val(txtR) + val(txtPrevBalance) + val(txtServiceCounter), 2)
TxtVAtPercent_Change

ReLineGrid
End Sub
Private Sub CalcVatValue()
    
End Sub

Private Sub txtLastInvoiceRead2_Change()
CalcTotalCounter
End Sub


