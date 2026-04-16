VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormEmpMoveDepartment 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·» ‰ﬁ· „ÊŸ›"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   Icon            =   "formempmovedepartment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   10215
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame9 
      Caption         =   "»Ì«‰«  „Õ«”»Ì…"
      Height          =   735
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   81
      Top             =   6240
      Width           =   7335
      Begin VB.TextBox TxtNoteID 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command9 
         Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "ﬂ‘› Õ”«»"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         Height          =   195
         Index           =   35
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.TextBox TxtSearchCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   1185
      Width           =   1335
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
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   7560
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
      Width           =   10245
      _cx             =   18071
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
      Caption         =   "ÿ·» ‰ﬁ· „ÊŸ›  "
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
         Left            =   1320
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
         ButtonImage     =   "formempmovedepartment.frx":038A
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
         ButtonImage     =   "formempmovedepartment.frx":0724
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
         Left            =   1920
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
         ButtonImage     =   "formempmovedepartment.frx":0ABE
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
         Left            =   720
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
         ButtonImage     =   "formempmovedepartment.frx":0E58
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
         Left            =   4320
         Picture         =   "formempmovedepartment.frx":11F2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label lblb 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2400
         TabIndex        =   32
         Top             =   0
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   5340
      TabIndex        =   7
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   65273857
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DcboEmpName 
      Height          =   315
      Left            =   4200
      TabIndex        =   8
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
      Left            =   1710
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7020
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         Left            =   225
         TabIndex        =   14
         Top             =   75
         Visible         =   0   'False
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
         Left            =   1080
         TabIndex        =   15
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
         Left            =   1935
         TabIndex        =   16
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
         Left            =   3840
         TabIndex        =   28
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
         Left            =   3000
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
      Left            =   6420
      TabIndex        =   17
      Top             =   5880
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
   Begin MSDataListLib.DataCombo Dcbranch 
      Bindings        =   "formempmovedepartment.frx":4E5A
      Height          =   315
      Left            =   240
      TabIndex        =   29
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
      Height          =   4215
      Left            =   0
      TabIndex        =   37
      Top             =   1560
      Width           =   10200
      _cx             =   17992
      _cy             =   7435
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
      Caption         =   "«·»Ì«‰«  «·«”«”Ì…|Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "formempmovedepartment.frx":4E6F
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3750
         Left            =   10845
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   45
         Width           =   10110
         _cx             =   17833
         _cy             =   6615
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
            Height          =   3030
            Left            =   120
            TabIndex        =   39
            Tag             =   "1"
            Top             =   240
            Width           =   10950
            _cx             =   19315
            _cy             =   5345
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
            FormatString    =   $"formempmovedepartment.frx":5209
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
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   3360
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
         Height          =   3750
         Index           =   15
         Left            =   45
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   45
         Width           =   10110
         _cx             =   17833
         _cy             =   6615
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
         _GridInfo       =   $"formempmovedepartment.frx":534C
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3720
            Index           =   16
            Left            =   15
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   15
            Width           =   10080
            _cx             =   17780
            _cy             =   6562
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
            Begin VB.TextBox TxtDiff 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3264
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   3375
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txtreson 
               Alignment       =   1  'Right Justify
               Height          =   660
               Left            =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   64
               Top             =   2400
               Width           =   8760
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   330
               Left            =   -165
               TabIndex        =   50
               Top             =   3180
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   582
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
            Begin MSDataListLib.DataCombo DcmbFromDepart 
               Height          =   315
               Left            =   4995
               TabIndex        =   52
               Top             =   570
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcmbToDepart 
               Height          =   315
               Left            =   90
               TabIndex        =   53
               Top             =   570
               Width           =   3945
               _ExtentX        =   6959
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcmbFromProject 
               Height          =   285
               Left            =   4995
               TabIndex        =   55
               Top             =   945
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcmbToProject 
               Height          =   285
               Left            =   90
               TabIndex        =   56
               Top             =   945
               Width           =   3945
               _ExtentX        =   6959
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcmbManagerID 
               Height          =   285
               Left            =   4995
               TabIndex        =   57
               Top             =   1740
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboJobsType 
               Height          =   315
               Left            =   4995
               TabIndex        =   58
               Top             =   1335
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DBIssueDate 
               Height          =   285
               Left            =   2775
               TabIndex        =   59
               Top             =   1740
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               _Version        =   393216
               Format          =   65273857
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DcmbToJob 
               Height          =   315
               Left            =   90
               TabIndex        =   61
               Top             =   1380
               Width           =   3945
               _ExtentX        =   6959
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DateDiv 
               Height          =   300
               Left            =   4830
               TabIndex        =   76
               Top             =   3375
               Visible         =   0   'False
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   529
               _Version        =   393216
               Format          =   65273857
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DcbFrmBarnch 
               Height          =   315
               Left            =   4995
               TabIndex        =   77
               Top             =   120
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbToBarnch 
               Height          =   315
               Left            =   90
               TabIndex        =   78
               Top             =   120
               Width           =   3945
               _ExtentX        =   6959
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DBToDate 
               Height          =   285
               Left            =   120
               TabIndex        =   87
               Top             =   1800
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               _Version        =   393216
               Format          =   65273857
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " √—ÌŒ «·‰ﬁ· «·Ì"
               Height          =   465
               Index           =   5
               Left            =   1470
               TabIndex        =   88
               Top             =   1800
               Width           =   1065
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï ›—⁄"
               Height          =   345
               Left            =   3795
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   120
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ ›—⁄"
               Height          =   345
               Index           =   0
               Left            =   8625
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   120
               Width           =   1230
            End
            Begin VB.Label lblfd 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ «·«œ«—…"
               Height          =   345
               Left            =   8595
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   570
               Width           =   1230
            End
            Begin VB.Label lblm 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„œÌ— "
               Height          =   345
               Left            =   8595
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   1740
               Width           =   1230
            End
            Begin VB.Label lblfj 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ ÊŸÌ›…"
               Height          =   345
               Left            =   8595
               TabIndex        =   69
               Top             =   1380
               Width           =   1230
            End
            Begin VB.Label lblfp 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ „Êﬁ⁄"
               Height          =   345
               Left            =   8595
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   960
               Width           =   1230
            End
            Begin VB.Label lblreson 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”»»  «·‰ﬁ·"
               Height          =   420
               Left            =   8895
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   2325
               Width           =   1050
            End
            Begin VB.Label lbltj 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï ÊŸÌ›…"
               Height          =   345
               Left            =   3765
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   1380
               Width           =   1080
            End
            Begin VB.Label lbltp 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï „Êﬁ⁄"
               Height          =   345
               Left            =   3765
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   960
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " √—ÌŒ «·‰ﬁ· „‰"
               Height          =   465
               Index           =   13
               Left            =   4125
               TabIndex        =   60
               Top             =   1740
               Width           =   600
            End
            Begin VB.Label lbltd 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï «·«œ«—…"
               Height          =   345
               Left            =   3765
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   570
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2550
               Index           =   62
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   1110
               Width           =   555
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3720
            Index           =   9
            Left            =   15
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   15
            Width           =   10080
            _cx             =   17780
            _cy             =   6562
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
               Height          =   2790
               Left            =   2616
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   795
               Width           =   540
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   1935
               Left            =   3312
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   990
               Width           =   840
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   1935
               Index           =   67
               Left            =   1860
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   990
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   1860
               Index           =   68
               Left            =   3165
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1260
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
               Height          =   2250
               Index           =   69
               Left            =   2340
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   990
               Width           =   285
            End
         End
      End
   End
   Begin VB.Label XPTxtCount1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   360
      TabIndex        =   74
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   0
      TabIndex        =   73
      Top             =   -2040
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2280
      TabIndex        =   72
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—« » «·«”«”Ì"
      Height          =   285
      Index           =   9
      Left            =   2880
      TabIndex        =   67
      Top             =   1200
      Width           =   1110
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   66
      Top             =   1200
      Width           =   870
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
   Begin VB.Label lblbranch 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄ «·ﬁ«∆„ »«·⁄„·Ì…"
      Height          =   255
      Index           =   0
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   780
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   285
      Index           =   4
      Left            =   9030
      TabIndex        =   27
      Top             =   750
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   285
      Index           =   3
      Left            =   9030
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
      Left            =   6270
      TabIndex        =   25
      Top             =   735
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   9165
      TabIndex        =   24
      Top             =   5955
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   3120
      TabIndex        =   23
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   1080
      TabIndex        =   22
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   21
      Top             =   6780
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   20
      Top             =   6780
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
Attribute VB_Name = "FormEmpMoveDepartment"
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

Private Sub Accredit_Click()
 If val(XPTxtID.Text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
         
   ' Dim BeginTrans As Boolean
'
'    Cn.BeginTrans
'    BeginTrans = True
'
'    If IsNull(rs("Posted")) Then
'        rs("Posted") = user_id
'        rs("PostedDate") = Time
'    Else
        'rs("Posted") = Null
'       rs("PostedDate") = Time
'    End If
   
'    rs.update
' If SystemOptions.UserInterface = ArabicInterface Then
'    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
'Else
'Accredit.Caption = "Sent To approval "
'End If

 '   Cn.CommitTrans
 '   BeginTrans = False
'FillApprovedTable
'    Retrive (val(Me.XPTxtID.Text))
    Dim sql As String
    Dim BeginTrans As Boolean

    SendTopost Me.Name, "TblMoveEmp1", "Id", val(DcmbFromDepart.BoundText), val(Dcbranch.BoundText), val(XPTxtID.Text), XPTxtID.Text
    rs.Resync
    
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

  Retrive (val(XPTxtID.Text))
End Sub
Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData left  JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.Rows = RsDetails.RecordCount + 1
 

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
 GRID2.Rows = 1
    End If
RsDetails.Close

End Function
 
 Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
 lbl(2).Caption = 0
            
              GRID2.Clear flexClearScrollable, flexClearEverything
    GRID2.Rows = 1
            Me.DCboUserName.BoundText = user_id
           ' TxtPaymentCounts.text = 1
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
 If ScreenAproved(val(XPTxtID.Text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not edit.This process associated with approvals"
         End If
         Exit Sub
       End If
       
  
  
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
             
        If ScreenAproved(val(XPTxtID.Text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «·Õ–›.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not delete.This process associated with approvals"
         End If
         Exit Sub
       End If
  
  
            Del_Trans

        Case 5
        General_Search.send_form = "empmov"
            Load General_Search
           General_Search.show
        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.Text, , 200

            
            
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


 
MySQL = "   SELECT dbo.TblMoveEmp1.BranchID, dbo.TblMoveEmp1.ID, dbo.TblMoveEmp1.RecordDate, dbo.TblMoveEmp1.EmpID, dbo.TblMoveEmp1.FromDepart, dbo.TblMoveEmp1.ToDepart,"
MySQL = MySQL + "                     dbo.TblMoveEmp1.ManagerID, dbo.TblMoveEmp1.moveDate, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
MySQL = MySQL + "                     TblEmployee_1.Emp_Code AS mangercode, TblEmployee_1.Emp_Name AS mangername, TblEmployee_1.Emp_Namee AS mangernamee,"
 MySQL = MySQL + "                    TblEmpDepartments_2.DepartmentName AS deptom, TblEmpDepartments_2.DepartmentNamee AS deptomE, TblEmpDepartments_1.DepartmentName AS depfrom,"
   MySQL = MySQL + "                  TblEmpDepartments_1.DepartmentNamee AS depfrome, dbo.TblMoveEmp1.JobID, TblEmpJobsTypes_1.JobTypeName, TblEmpJobsTypes_1.JobTypeNamee,"
   MySQL = MySQL + "                  TblEmployee_1.Emp_Name AS Namea, TblEmployee_1.Emp_Name1 AS Namea1, TblEmployee_1.Emp_Name2 AS Namea2, TblEmployee_1.Emp_Name3 AS Namea3,"
    MySQL = MySQL + "                 TblEmployee_1.Emp_Name4 AS Namea4, TblEmployee_1.Emp_Namee4 AS Namee4, TblEmployee_1.Emp_Namee3 AS Namee3, TblEmployee_1.Emp_Namee2 AS Namee2,"
   MySQL = MySQL + "                  TblEmployee_1.Emp_Namee1 AS Namee1, dbo.TblMoveEmp1.ProjectFrom, dbo.TblMoveEmp1.ProjectTo, dbo.TblMoveEmp1.basicSalary, dbo.TblMoveEmp1.Reson,"
     MySQL = MySQL + "                EmpGroupDep_2.GroupName AS frmdep, EmpGroupDep_1.GroupName AS todep, dbo.TblMoveEmp1.JobTo, TblEmpJobsTypes_1.JobTypeName AS namejob,"
       MySQL = MySQL + "              TblEmpJobsTypes_1.JobTypeNamee AS nameejob, dbo.TblMoveEmp1.UserID, dbo.TblMoveEmp1.posted, TblEmployee_2.Emp_Code, TblEmployee_2.Emp_Name,"
       MySQL = MySQL + "              TblEmployee_2.Emp_Namee, TblEmpJobsTypes_2.JobTypeName AS frmjob, TblEmpJobsTypes_2.JobTypeNamee AS frmjobE,"
    MySQL = MySQL + "                 TblEmployee_2.BignDateWork AS BignDateWork1, TblEmployee_2.Fullcode"
MySQL = MySQL + "   FROM     dbo.TblEmployee AS TblEmployee_1 RIGHT OUTER JOIN"
   MySQL = MySQL + "                 dbo.TblMoveEmp1 LEFT OUTER JOIN"
     MySQL = MySQL + "                dbo.TblEmpJobsTypes AS TblEmpJobsTypes_1 ON dbo.TblMoveEmp1.JobTo = TblEmpJobsTypes_1.JobTypeID LEFT OUTER JOIN"
       MySQL = MySQL + "              dbo.EmpGroupDep AS EmpGroupDep_1 ON dbo.TblMoveEmp1.ProjectTo = EmpGroupDep_1.GroupID LEFT OUTER JOIN"
        MySQL = MySQL + "             dbo.EmpGroupDep AS EmpGroupDep_2 ON dbo.TblMoveEmp1.ProjectFrom = EmpGroupDep_2.GroupID LEFT OUTER JOIN"
       MySQL = MySQL + "              dbo.TblBranchesData ON dbo.TblMoveEmp1.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
        MySQL = MySQL + "             dbo.TblEmployee AS TblEmployee_2 ON dbo.TblMoveEmp1.EmpID = TblEmployee_2.Emp_ID ON TblEmployee_1.Emp_ID = dbo.TblMoveEmp1.ManagerID LEFT OUTER JOIN"
         MySQL = MySQL + "            dbo.TblEmpDepartments AS TblEmpDepartments_1 ON dbo.TblMoveEmp1.FromDepart = TblEmpDepartments_1.DeparmentID LEFT OUTER JOIN"
           MySQL = MySQL + "          dbo.TblEmpDepartments AS TblEmpDepartments_2 ON dbo.TblMoveEmp1.ToDepart = TblEmpDepartments_2.DeparmentID LEFT OUTER JOIN"
          MySQL = MySQL + "           dbo.TblEmpJobsTypes AS TblEmpJobsTypes_2 ON dbo.TblMoveEmp1.JobID = TblEmpJobsTypes_2.JobTypeID"



    MySQL = MySQL & " Where (dbo.TblMoveEmp1.id = " & val(XPTxtID.Text) & ")"
 
        If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\empmove.rpt"
             '  StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_EmpMoveDepartment.rpt"
        Else
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\empmove.rpt"
               ' StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_EmpMoveDepartment.rpt"
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
  '      xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
         xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
  '  xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
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
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub



Private Sub Command8_Click()
Dim StrTempAccountCode As String
Dim FirstPeriod As Date
        getFirstPeriodDateInthisYear FirstPeriod
       StrTempAccountCode = get_EMPLOYEE_Account(val(DcboEmpName.BoundText), "Account_Code1")    '«·«ÃÊ— «·„” Õﬁ…
       ShowReport StrTempAccountCode, DcboEmpName.Text, FirstPeriod, Date
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub


Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub

 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 13
       ' Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If

End Sub
Function GETlASTiSSUEDATE(Emp_id As Integer, Optional novalue As Boolean) As Date
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
  sql = "SELECT     MAX(todate) AS MaxDate from dbo.TblEmpHolidaysDetails WHERE     (Emp_ID = " & Emp_id & ")"
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    If Not IsNull(rs("MaxDate").value) Then
 GETlASTiSSUEDATE = IIf(IsNull(rs("MaxDate").value), Date, rs("MaxDate").value)
novalue = False
Else
 GETlASTiSSUEDATE = Date
 novalue = True
 End If
 Else
 GETlASTiSSUEDATE = Date
 novalue = True
    End If

End Function
Private Sub DcboEmpName_Click(Area As Integer)
     On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub
Dim novalue As Boolean
    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
    
   If Me.TxtModFlg = "R" Then Exit Sub
   
   
    Dim StrSQL As String
      
        
        Dim IssueDate As Date
        Dim DepID As Double
        Dim specid As Double
        Dim JobTypeID As Double
        Dim gradeID As Double
        Dim Account_code2 As String
        Dim gropid As Integer
        Dim Account_code  As String
        Dim Balance As String
        Dim manger As Integer
        Dim BranchID As Integer
        Dim DateMoveNo As Date
        Dim endContractPerMonth As Double
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth, , manger, , gropid, , , , , , , , , , , , , , , DateMoveNo, , , , , , , , , , , BranchID
        
          WriteCustomerBalPublic Account_code2, Balance
            WriteCustomerBalPublic Account_code, Balance
            Me.DcboJobsType.BoundText = JobTypeID
        DBIssueDate.value = IssueDate
       ' DBToDate.Value =ToDate
        Me.DcbFrmBarnch.BoundText = BranchID
        Me.DcmbFromDepart.BoundText = DepID
        Me.dcmbFromProject.BoundText = gropid
        If gropid = 0 Then
        dcmbFromProject.Enabled = True
        Else
       dcmbFromProject.Enabled = False
        
        End If
        
       Me.DcmbManagerID.BoundText = manger
       Me.txtDiff.Text = DateDiff("d", DateMoveNo, XPDtbTrans.value)
   '   Me.Dcbranch.BoundText
   lbl(2).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
      DateDiv.value = GETlASTiSSUEDATE(val(Me.DcboEmpName.BoundText), novalue)
      If novalue = False Then
      Me.txtDiff.Text = DateDiff("d", DateDiv, XPDtbTrans.value)
      End If
      
      
    'End If

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
     Dcombos.GetUsers Me.DCboUserName
     Dcombos.GetEmployees Me.DcboEmpName
     Dcombos.GetBranches Me.Dcbranch
     Dcombos.GetBranches Me.DcbFrmBarnch
     Dcombos.GetBranches Me.DcbToBarnch
     Dcombos.GetEmpDepartments Me.DcmbFromDepart
     Dcombos.GetEmpDepartments Me.DcmbToDepart
     Dcombos.GetEmployees Me.DcmbManagerID
     Dcombos.GetEmpJobsTypes Me.DcboJobsType
   Dcombos.GetEmpJobsTypes Me.DcmbToJob
    Dcombos.GetEmpLocations Me.dcmbFromProject ' location
   Dcombos.GetEmpLocations Me.dcmbToProject ' location
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
    
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblMoveEmp1    Order By ID"
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
        Command8.Caption = "Acc.Statement"
'Me.Label1.Caption = "Duration"
Frame9.Caption = "Accounting"
Label1(35).Caption = "No.GL"
    Command9.Caption = "Print GL"
    Set Me.XPBtnMove(3).ButtonImage = XPic
   ' Label1.Visible = False
   Label2.Caption = "To Branch"
   Label1(0).Caption = "From Branch"
Accredit.Caption = "Send to Approval"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
XPTab301.Caption = "Data|Approved"
    Me.Caption = "Employee Transfer"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(3).Caption = "Employee"
    lbl(2).Caption = "value"
   ' lbl(0).Caption = "Box"
   ' Fra(0).Caption = "payments Method"
    lbl(9).Caption = "Basic Salary"
    lbl(13).Caption = "Transfer Datae"
     lblBranch(0).Caption = "Branch"
     Me.lblfp.Caption = "From Loca"
     Me.lbltp.Caption = "To Loca"
     Me.lblfd.Caption = "From Dept"
     Me.lbltd.Caption = "To Dept"
     Me.lblfj.Caption = "From Job"
     Me.lbltj.Caption = "To Job"
     Me.LblM.Caption = "Manager"
     Me.lblreson.Caption = "Reson of transfer"

   ' lbl(10).Caption = "Start"
   ' lbl(11).Caption = "Month"
   ' lbl(12).Caption = "Year"
  '  Cmd(8).Caption = "Calc Dates"
 '   ChkSaleryDis.Caption = "Auto Discount"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
  '  me.
'lbl(24).Caption = "From Job"

    Label11.Caption = "Approval Requested By"
    
    With GRID2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
    End With

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

'Private Sub TxtAdvanceValue_LostFocus()
  '  Dim StrSQL As String
  '  Dim Mytot As String
  '  Dim MySal As String
  '  Exit Sub
 '   Dim Myrs As New ADODB.Recordset
    'StrSQL =
   ' Myrs.Open "SELECT * From TblEmployee  where EmpID=" & val(DcboEmpName.BoundText), Cn, adOpenStatic, adLockReadOnly

   ' If Not Myrs.EOF And Not IsNull(Myrs!Emp_Salary) Then
    '    MySal = Myrs!Emp_Salary
     '   Mytot = val(MySal) * 5
'
      '  If val(TxtAdvanceValue.text) >= Mytot Then
        '    MsgBox "⁄›Ê« «·”·›…  ⁄œ  «·Õœ  «·„”„ÊÕ »Â ÊÂÊ 5 «÷⁄«› ﬁÌ„Â «·—« »  " & Chr(13) & "   —« » «·„ÊŸ›    " & MySal, vbOKOnly, App.Title
        '    Exit Sub
   
  '      End If
  '
  '  End If
   
'End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            '        Me.Caption = "ÿ·» ‰ﬁ· „ÊŸ›"
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
            '        Me.Caption = "ÿ·» ‰ﬁ· „ÊŸ›( ÃœÌœ )"
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
         '  TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = "ÿ·» ‰ﬁ· „ÊŸ›(  ⁄œÌ· )"
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
    Dim i As Integer
    Dim StrSQL As String

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        Me.XPTxtCount1.Caption = 0
    Me.XPTxtCount1.Caption = 0
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
    Me.TxtNoteID.Text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    XPTxtID.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcmbManagerID.BoundText = IIf(IsNull(rs("ManagerID").value), "", rs("ManagerID").value)
    DcboJobsType.BoundText = IIf(IsNull(rs("JobID").value), "", rs("JobID").value)
    DcmbToJob.BoundText = IIf(IsNull(rs("JobTo").value), "", rs("JobTo").value)
    DcmbFromDepart.BoundText = IIf(IsNull(rs("FromDepart").value), "", rs("FromDepart").value)
    DcmbToDepart.BoundText = IIf(IsNull(rs("ToDepart").value), "", rs("ToDepart").value)
    dcmbToProject.BoundText = IIf(IsNull(rs("ProjectTo").value), "", rs("ProjectTo").value)
    dcmbFromProject.BoundText = IIf(IsNull(rs("ProjectFrom").value), "", rs("ProjectFrom").value)
    DBIssueDate.value = IIf(IsNull(rs("moveDate").value), "", rs("moveDate").value)
     
     dbTodate.value = IIf(IsNull(rs("ToDate").value), "", rs("ToDate").value)
        
    'DBToDate
    lbl(2).Caption = IIf(IsNull(rs("basicSalary").value), "", rs("basicSalary").value)
    TxtReson.Text = IIf(IsNull(rs("reson").value), "", rs("reson").value)
    Me.txtDiff.Text = IIf(IsNull(rs("DiffDate").value), "", rs("DiffDate").value) ' rs("DiffDate").value
    DcbFrmBarnch.BoundText = IIf(IsNull(rs("FroBranchID").value), "", rs("FroBranchID").value)
    DcbToBarnch.BoundText = IIf(IsNull(rs("ToBranchID").value), "", rs("ToBranchID").value)

 '   lbl(22).Caption = IIf(IsNull(rs("EmpDue").value), "", rs("EmpDue").value)
 '  lbl(20).Caption = IIf(IsNull(rs("Contractvalid").value), "", rs("Contractvalid").value)
 '  lbl(21).Caption = IIf(IsNull(rs("oldAdvance").value), "", rs("oldAdvance").value)
 

Me.DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    'TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
  '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
  
 
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
    StrSQL = "Select * From  TblMoveEmp1 Where ID=" & val(XPTxtID.Text)
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    

    RsDetails.Close
    Set RsDetails = Nothing
    
    fillapprovData
    
   Me.XPTxtCurrent1.Caption = rs.AbsolutePosition
Me.XPTxtCount1.Caption = rs.RecordCount
  
    Exit Sub
ErrTrap:
End Sub
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
Dim Balance As String
des = "‰ﬁ· „ÊŸ› »—ﬁ„ —ﬁ„" & XPTxtID.Text & " ··„ÊŸ› " & DcboEmpName.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim depit_side As String
Dim BranchID As Integer
Dim sql As String
tablename = "TblMoveEmp1"
Filedname = "ID"
NoteSerial1 = val(XPTxtID.Text)
Notevalue = 0
notytype = 9051
BranchID = val(Dcbranch.BoundText)
NoteDate = (XPDtbTrans.value)
  depit_side = get_EMPLOYEE_Account(DcboEmpName.BoundText, "Account_Code1")
                WriteCustomerBalPublic depit_side, Balance
                Notevalue = val(Balance)
                 depit_side = get_EMPLOYEE_Account(DcboEmpName.BoundText, "Account_Code2")
                WriteCustomerBalPublic depit_side, Balance
                Notevalue = Notevalue + val(Balance)
                
 
If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TxtNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial
                                     Else
                                                 If TxtNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TxtNoteID.Text = NoteID
                                                                TxtNoteSerial.Text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TxtNoteID.Text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TxtNoteID.Text), BranchID, user_id, NoteDate
 rs.Resync adAffectCurrent
     End If

End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords

    Dim LngDevID As Long
    Dim Msg As String
    Dim FromBrnchID     As Integer
    Dim ToBrnchID     As Integer
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim total_value As Double
    Dim depit_side As String
     Dim CURRENT_LINE As Double
     Dim Balance As String
     notes_id = general_noteid
   FromBrnchID = val(Me.DcbFrmBarnch.BoundText)
   ToBrnchID = val(Me.DcbToBarnch.BoundText)
   LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

Dim lineno As Integer
If SystemOptions.UserInterface = ArabicInterface Then
Msg = " ÿ·» ‰ﬁ· „ÊŸ› »—ﬁ„ " + XPTxtID.Text
Else
Msg = " Move  Employee " + XPTxtID.Text

End If
lineno = 1
''******************************************«ÃÊ— „” Õﬁ…
    
                depit_side = get_EMPLOYEE_Account(DcboEmpName.BoundText, "Account_Code1")
                CURRENT_LINE = setfoxy_Line
                WriteCustomerBalPublic depit_side, Balance
                total_value = val(Balance)
                
                If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " «ÃÊ— „” Õﬁ…  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , FromBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                      If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + "«ÃÊ— „” Õﬁ…  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , ToBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
                
               depit_side = get_EMPLOYEE_Account(DcboEmpName.BoundText, "Account_Code2")
                CURRENT_LINE = setfoxy_Line
                WriteCustomerBalPublic depit_side, Balance
                total_value = val(Balance)
                
                If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " „Œ’’«  «Ã«“…  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , FromBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                      If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + "„Œ’’«  «Ã«“… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , ToBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
            ''///////////
            depit_side = get_account_code_branch(55, Me.DcbToBarnch.Text)

                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " „’—Ê› „Œ’’ «Ã«“… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , ToBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                      If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + "„’—Ê› „Œ’’ «Ã«“…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , FromBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
           '''//////////////
                        depit_side = get_EMPLOYEE_Account(DcboEmpName.BoundText, "Account_Code4")
                CURRENT_LINE = setfoxy_Line
                WriteCustomerBalPublic depit_side, Balance
                total_value = val(Balance)
                
                If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " „Œ’’ ‰Â«Ì… «·Œœ„…  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , FromBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                      If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + " „Œ’’ ‰Â«Ì… «·Œœ„… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , ToBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
           
           '//////////////
                  depit_side = get_account_code_branch(56, Me.DcbToBarnch.Text)
                CURRENT_LINE = setfoxy_Line
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " „’—Ê› „Œ’’ ‰Â«Ì… «·Œœ„… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , ToBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                      If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + "„’—Ê› „Œ’’ ‰Â«Ì… «·Œœ„…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , FromBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
            '''/////////
                        depit_side = get_EMPLOYEE_Account(DcboEmpName.BoundText, "Account_Code5")
                         total_value = val(Balance)
                CURRENT_LINE = setfoxy_Line
                WriteCustomerBalPublic depit_side, Balance
               
                
                If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " „Œ’’  –«ﬂ—  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , FromBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                      If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + " „Œ’’  –«ﬂ—  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , ToBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
    
                          depit_side = get_account_code_branch(94, Me.DcbToBarnch.Text)
                
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " „’—Ê› „Œ’’  –«ﬂ— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , ToBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                      If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + "„’—Ê› „Œ’’  –«ﬂ— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , FromBrnchID, , , , , , , , val(DcboEmpName.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:

'********************************************************************
  End Function
Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim sql As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

 '   On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
        Else
        Msg = "Please Select Employee"
       End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboEmpName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

   
 
'
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then
            XPTxtID.Text = CStr(new_id("TblMoveEmp1", "ID", "", True))
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
   StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        End If
        If val(Me.dcmbToProject.BoundText) <> 0 Then
sql = "update TblEmployee set   GroupID =" & val(Me.dcmbToProject.BoundText) & "  where Emp_ID =" & val(Me.DcboEmpName.BoundText) & ""
                                    Cn.Execute sql
                                    End If
                           If val(Me.DcmbToDepart.BoundText) <> 0 Then
 sql = "update TblEmployee set   DepartmentID = " & val(Me.DcmbToDepart.BoundText) & "   where Emp_ID =" & val(Me.DcboEmpName.BoundText) & ""
                                    Cn.Execute sql
                                    End If
                              If val(Me.DcmbToJob.BoundText) <> 0 Then
   sql = "update TblEmployee set   JobTypeID =" & val(Me.DcmbToJob.BoundText) & "  where Emp_ID =" & val(Me.DcboEmpName.BoundText) & ""
                                  Cn.Execute sql
                                  End If

   sql = "update TblEmployee set   DateMoveno =" & SQLDate(Me.XPDtbTrans.value, True) & "  where Emp_ID =" & val(Me.DcboEmpName.BoundText) & ""
                                    Cn.Execute sql
  If val(DcbToBarnch.BoundText) <> 0 Then
  If val(DcbToBarnch.BoundText) <> val(DcbFrmBarnch.BoundText) Then
 sql = "update TblEmployee set   BranchID =" & val(Me.DcbToBarnch.BoundText) & "  where Emp_ID =" & val(Me.DcboEmpName.BoundText) & ""
 Cn.Execute sql
 End If
 End If
        rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
        rs("FromDepart").value = IIf(Me.DcmbFromDepart.BoundText = "", Null, Me.DcmbFromDepart.BoundText)
        rs("ToDepart").value = IIf(Me.DcmbToDepart.BoundText = "", Null, Me.DcmbToDepart.BoundText)
        rs("ID").value = val(XPTxtID.Text)
        rs("RecordDate").value = XPDtbTrans.value
        rs("EmpID").value = IIf(Me.DcboEmpName.BoundText = "", Null, Me.DcboEmpName.BoundText) 'Me.DcboEmpName.BoundText
        rs("ManagerID").value = IIf(Me.DcmbManagerID.BoundText = "", Null, Me.DcmbManagerID.BoundText) 'val(Me.DcmbManagerID.BoundText)
        rs("JobID").value = IIf(Me.DcboJobsType.BoundText = "", Null, Me.DcboJobsType.BoundText) 'val(Me.DcboJobsType.BoundText)
        rs("moveDate").value = Me.DBIssueDate.value
        rs("ToDate").value = Me.dbTodate.value
        rs("JobTo").value = IIf(Me.DcmbToJob.BoundText = "", Null, Me.DcmbToJob.BoundText) 'val(Me.DcmbToJob.BoundText)
        rs("ProjectTo").value = IIf(Me.dcmbToProject.BoundText = "", Null, Me.dcmbToProject.BoundText) 'val(Me.dcmbToProject.BoundText)
        rs("ProjectFrom").value = IIf(Me.dcmbFromProject.BoundText = "", Null, Me.dcmbFromProject.BoundText) 'val(Me.dcmbFromProject.BoundText)
        rs("reson").value = Me.TxtReson.Text
        
        rs("DiffDate").value = Me.txtDiff.Text
        rs("basicSalary").value = val(lbl(2).Caption)
        rs("FroBranchID").value = IIf(Me.DcbFrmBarnch.BoundText = "", Null, Me.DcbFrmBarnch.BoundText)
        rs("ToBranchID").value = IIf(Me.DcbToBarnch.BoundText = "", Null, Me.DcbToBarnch.BoundText)
      rs.update
        Cn.CommitTrans
        BeginTrans = False
If SystemOptions.AllowIndirectCost = True Then
        createVoucher
  End If
'        RsDetails.Close
        Set RsDetails = Nothing
      Me.XPTxtCurrent1.Caption = rs.AbsolutePosition
       Me.XPTxtCount1.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
          Else
             Msg = " Saved  " & CHR(13)
                Msg = Msg + "Âyou need new transaction"
                
          End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
 
If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
Else
MsgBox "Update success", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End If
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

    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
                
    Deletepost Me.Name, "TblMoveEmp1", "Id", val(DcmbFromDepart.BoundText), val(Dcbranch.BoundText), val(XPTxtID.Text), XPTxtID.Text
    rs.delete
    
                rs.MoveFirst



                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                 Me.XPTxtCurrent1.Caption = 0
                Me.XPTxtCount1.Caption = 0
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
                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.Text)
                   RSApproval("NoteSerial").value = val(Me.XPTxtID.Text)
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
        .Create Me.hwnd, "ÿ·» ‰ﬁ· „ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "ÿ·» ‰ﬁ· „ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "ÿ·» ‰ﬁ· „ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "ÿ·» ‰ﬁ· „ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "ÿ·» ‰ﬁ· „ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "ÿ·» ‰ﬁ· „ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hwnd, "ÿ·» ‰ﬁ· „ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "ÿ·» ‰ﬁ· „ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "ÿ·» ‰ﬁ· „ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "ÿ·» ‰ﬁ· „ÊŸ›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "ÿ·» ‰ﬁ· „ÊŸ›", 1, 15204351, -2147483630
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



