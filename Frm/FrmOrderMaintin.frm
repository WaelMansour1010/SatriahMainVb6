VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmOrderMaintin 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«„— ‘€· ’Ì«‰…"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17610
   Icon            =   "FrmOrderMaintin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   17610
   Begin VB.TextBox MaintPlan 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   194
      Top             =   720
      Width           =   915
   End
   Begin VB.OptionButton BaisedOn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»‰«¡ ⁄·Ï Œÿ… ’Ì«‰… —ﬁ„ "
      Height          =   315
      Index           =   1
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   193
      Top             =   720
      Width           =   2025
   End
   Begin VB.OptionButton BaisedOn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»‰«¡ ⁄·Ï ÿ·» —ﬁ„"
      Height          =   315
      Index           =   0
      Left            =   7230
      RightToLeft     =   -1  'True
      TabIndex        =   192
      Top             =   720
      Width           =   1545
   End
   Begin VB.ComboBox DcbType 
      Height          =   315
      Left            =   11340
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   720
      Width           =   1275
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   18960
      TabIndex        =   36
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   18480
      TabIndex        =   35
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   19080
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   19200
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtOrder 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6030
      MaxLength       =   10
      TabIndex        =   0
      Top             =   720
      Width           =   1155
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   15480
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   18180
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   3
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
      Caption         =   "   «„— ‘€· ’Ì«‰…  "
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
         ButtonImage     =   "FrmOrderMaintin.frx":038A
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
         ButtonImage     =   "FrmOrderMaintin.frx":0724
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
         ButtonImage     =   "FrmOrderMaintin.frx":0ABE
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
         TabIndex        =   7
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
         ButtonImage     =   "FrmOrderMaintin.frx":0E58
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
         Left            =   5880
         Picture         =   "FrmOrderMaintin.frx":11F2
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
         TabIndex        =   34
         Top             =   0
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   13380
      TabIndex        =   8
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   131334145
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   3390
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9060
      Width           =   11385
      _cx             =   20082
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
         Left            =   10680
         TabIndex        =   10
         Top             =   60
         Width           =   645
         _ExtentX        =   1138
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
         Left            =   9960
         TabIndex        =   11
         Top             =   60
         Width           =   645
         _ExtentX        =   1138
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
         Left            =   9255
         TabIndex        =   12
         Top             =   60
         Width           =   645
         _ExtentX        =   1138
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
         Left            =   8520
         TabIndex        =   13
         Top             =   30
         Width           =   645
         _ExtentX        =   1138
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
         Left            =   7785
         TabIndex        =   14
         Top             =   60
         Width           =   645
         _ExtentX        =   1138
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
         Left            =   120
         TabIndex        =   15
         Top             =   60
         Width           =   645
         _ExtentX        =   1138
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
         Left            =   2775
         TabIndex        =   16
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "«·„—›ﬁ« "
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
         Left            =   7080
         TabIndex        =   27
         Top             =   60
         Width           =   645
         _ExtentX        =   1138
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
         Left            =   6360
         TabIndex        =   38
         Top             =   60
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…"
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
         Left            =   5160
         TabIndex        =   186
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… «·›« Ê—…"
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
         Index           =   12
         Left            =   4440
         TabIndex        =   187
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… 1"
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
         Index           =   14
         Left            =   3720
         TabIndex        =   198
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄… 2"
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
         Index           =   15
         Left            =   1800
         TabIndex        =   219
         Top             =   60
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "Œÿ… «·’Ì«‰Â"
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
         Index           =   16
         Left            =   840
         TabIndex        =   220
         Top             =   60
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ·» «·’Ì«‰Â"
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
      Left            =   14940
      TabIndex        =   17
      Top             =   9120
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   18360
      TabIndex        =   18
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
      Left            =   18720
      TabIndex        =   29
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
      Bindings        =   "FrmOrderMaintin.frx":4E5A
      Height          =   315
      Left            =   8820
      TabIndex        =   31
      Top             =   720
      Width           =   1995
      _ExtentX        =   3519
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
   Begin C1SizerLibCtl.C1Elastic C1Elastic6 
      Height          =   7770
      Left            =   0
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1155
      Width           =   17610
      _cx             =   31062
      _cy             =   13705
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
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   7815
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   17640
         _cx             =   31115
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483633
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483632
         TabOutlineColor =   -2147483633
         FrontTabForeColor=   -2147483630
         Caption         =   "»Ì«‰«  «·’Ì«‰…|”‰œ«  «·’—›|„·Õﬁ« |Õ«·Â «·«⁄ „«œ|»Ì«‰«   ›’Ì·Ì…"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   1
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   0   'False
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   0   'False
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
         Picture(0)      =   "FrmOrderMaintin.frx":4E6F
         Flags(3)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   7350
            Left            =   18285
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   45
            Width           =   17550
            _cx             =   30956
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
            Begin VSFlex8UCtl.VSFlexGrid vchrgrid 
               Height          =   5685
               Left            =   120
               TabIndex        =   44
               Top             =   120
               Width           =   17325
               _cx             =   30559
               _cy             =   10028
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
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmOrderMaintin.frx":5209
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
                  Index           =   51
                  Left            =   0
                  TabIndex        =   45
                  Top             =   5880
                  Width           =   1440
               End
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”‰œ«  «·„‰’—›… ··«„—"
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   35
               Left            =   7680
               TabIndex        =   49
               Top             =   120
               Width           =   3120
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   " ÕœÌÀ"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   11280
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   120
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì  «·”‰œ« "
               Height          =   285
               Index           =   57
               Left            =   4440
               TabIndex        =   47
               Top             =   6240
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   58
               Left            =   240
               TabIndex        =   46
               Top             =   6240
               Width           =   3765
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7350
            Index           =   15
            Left            =   45
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   45
            Width           =   17550
            _cx             =   30956
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
            _GridInfo       =   $"FrmOrderMaintin.frx":53D5
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   7320
               Index           =   16
               Left            =   15
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   15
               Width           =   17520
               _cx             =   30903
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
               Begin VB.TextBox TxtLastKM 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   240
                  MaxLength       =   10
                  TabIndex        =   190
                  Top             =   2280
                  Width           =   1575
               End
               Begin VB.TextBox TxtCurrKM 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   240
                  MaxLength       =   10
                  TabIndex        =   188
                  Top             =   2640
                  Width           =   1575
               End
               Begin VB.TextBox TxtInitialNotes 
                  Alignment       =   1  'Right Justify
                  Height          =   465
                  Left            =   3480
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   184
                  Top             =   2160
                  Width           =   5655
               End
               Begin VB.TextBox TxtDeptNotes 
                  Alignment       =   1  'Right Justify
                  Height          =   465
                  Left            =   10440
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   182
                  Top             =   2160
                  Width           =   5775
               End
               Begin VB.TextBox TxtDiscount 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   19290
                  MaxLength       =   10
                  TabIndex        =   144
                  Top             =   2100
                  Width           =   1425
               End
               Begin VB.Frame Fra 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
                  Height          =   3765
                  Index           =   0
                  Left            =   18825
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   360
                  Width           =   6135
                  Begin VB.TextBox TxtPaymentCounts 
                     Alignment       =   2  'Center
                     Height          =   345
                     Left            =   4110
                     MaxLength       =   2
                     TabIndex        =   136
                     Top             =   240
                     Width           =   825
                  End
                  Begin VB.ComboBox CmbMonth 
                     Height          =   315
                     Left            =   4110
                     Style           =   2  'Dropdown List
                     TabIndex        =   135
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
                     TabIndex        =   134
                     Top             =   2160
                     Value           =   1  'Checked
                     Visible         =   0   'False
                     Width           =   1965
                  End
                  Begin VB.ComboBox CboYear 
                     Height          =   315
                     Left            =   4110
                     Style           =   2  'Dropdown List
                     TabIndex        =   133
                     Top             =   1320
                     Width           =   1095
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   435
                     Index           =   8
                     Left            =   4080
                     TabIndex        =   137
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
                     ButtonImage     =   "FrmOrderMaintin.frx":540B
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VSFlex8UCtl.VSFlexGrid Fg 
                     Height          =   2325
                     Index           =   0
                     Left            =   90
                     TabIndex        =   138
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
                     FormatString    =   $"FrmOrderMaintin.frx":57A5
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
                     TabIndex        =   143
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
                     TabIndex        =   142
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
                     TabIndex        =   141
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
                     TabIndex        =   140
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
                     TabIndex        =   139
                     Top             =   1320
                     Width           =   405
                  End
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  „«·Ì…"
                  Height          =   1005
                  Left            =   18000
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   0
                  Width           =   6015
                  Begin MSDataListLib.DataCombo DcboSpecifications 
                     Height          =   315
                     Left            =   3360
                     TabIndex        =   123
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
                     TabIndex        =   131
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
                     TabIndex        =   130
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
                     TabIndex        =   129
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
                     TabIndex        =   128
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
                     TabIndex        =   127
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
                     TabIndex        =   126
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
                     TabIndex        =   125
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
                     TabIndex        =   124
                     Top             =   360
                     Width           =   1125
                  End
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  «‰ﬁ·"
                  Height          =   1545
                  Left            =   18495
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   0
                  Width           =   6105
                  Begin MSDataListLib.DataCombo DcboEmpDepartments 
                     Height          =   315
                     Left            =   3120
                     TabIndex        =   108
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
                     TabIndex        =   109
                     Top             =   360
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   124256257
                     CurrentDate     =   38784
                  End
                  Begin MSDataListLib.DataCombo DcboJobsType 
                     Height          =   315
                     Left            =   3120
                     TabIndex        =   110
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
                     TabIndex        =   111
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
                     TabIndex        =   112
                     Top             =   600
                     Width           =   1935
                     _ExtentX        =   3413
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   124256257
                     CurrentDate     =   38784
                  End
                  Begin MSDataListLib.DataCombo DataCombo2 
                     Height          =   315
                     Left            =   1680
                     TabIndex        =   113
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
                     TabIndex        =   121
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
                     TabIndex        =   120
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
                     TabIndex        =   119
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
                     TabIndex        =   118
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
                     TabIndex        =   117
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
                     TabIndex        =   116
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
                     TabIndex        =   115
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
                     TabIndex        =   114
                     Top             =   240
                     Width           =   645
                  End
               End
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6600
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   120
                  Width           =   735
               End
               Begin VB.TextBox TxtCost 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   240
                  MaxLength       =   10
                  TabIndex        =   105
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.TextBox txtDes 
                  Alignment       =   1  'Right Justify
                  Height          =   585
                  Left            =   3480
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   104
                  Top             =   480
                  Width           =   3855
               End
               Begin VB.Frame Frame3 
                  BackColor       =   &H00E2E9E9&
                  Height          =   1215
                  Left            =   840
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   6120
                  Width           =   10575
                  Begin VB.TextBox Text2 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   5880
                     RightToLeft     =   -1  'True
                     TabIndex        =   92
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.TextBox reciverRemarks 
                     Alignment       =   1  'Right Justify
                     Height          =   360
                     Left            =   2640
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   91
                     Top             =   600
                     Width           =   4215
                  End
                  Begin VB.CheckBox ended 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " „ «‰Â«¡ «·’Ì«‰…"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Left            =   5760
                     RightToLeft     =   -1  'True
                     TabIndex        =   90
                     Top             =   960
                     Width           =   2175
                  End
                  Begin MSComCtl2.DTPicker endmaintenanceDate 
                     Height          =   315
                     Left            =   8040
                     TabIndex        =   93
                     Top             =   240
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   123600897
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker endmaintenanceTime 
                     Height          =   315
                     Left            =   8040
                     TabIndex        =   94
                     Top             =   720
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   123600898
                     CurrentDate     =   38784
                  End
                  Begin MSDataListLib.DataCombo reciverid 
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   95
                     Top             =   240
                     Width           =   3195
                     _ExtentX        =   5636
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSComCtl2.DTPicker RecmaintenanceDate 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   96
                     Top             =   240
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   123600897
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker RecmaintenanceTime 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   97
                     Top             =   720
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   123600898
                     CurrentDate     =   38784
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " «—ÌŒ «·«‰ Â«¡"
                     Height          =   285
                     Index           =   43
                     Left            =   9330
                     TabIndex        =   103
                     Top             =   255
                     Width           =   1005
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Êﬁ  «·«‰ Â«¡"
                     Height          =   285
                     Index           =   44
                     Left            =   9420
                     TabIndex        =   102
                     Top             =   720
                     Width           =   885
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·„” ·„"
                     Height          =   285
                     Index           =   45
                     Left            =   6870
                     TabIndex        =   101
                     Top             =   255
                     Width           =   1125
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„·«ÕŸ«  «·„” ·„"
                     Height          =   285
                     Index           =   46
                     Left            =   6840
                     TabIndex        =   100
                     Top             =   600
                     Width           =   1125
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " «—ÌŒ «· ”·Ì„"
                     Height          =   405
                     Index           =   47
                     Left            =   1410
                     TabIndex        =   99
                     Top             =   255
                     Width           =   1125
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Êﬁ  «· ”·Ì„"
                     Height          =   405
                     Index           =   48
                     Left            =   1620
                     TabIndex        =   98
                     Top             =   720
                     Width           =   885
                  End
               End
               Begin VB.TextBox txtnote 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   11520
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   88
                  Top             =   6960
                  Width           =   4695
               End
               Begin VB.Frame Frame4 
                  BackColor       =   &H00E2E9E9&
                  Height          =   1095
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   0
                  Width           =   9015
                  Begin VB.TextBox TxtBoardNO 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   2520
                     TabIndex        =   72
                     Top             =   600
                     Width           =   1635
                  End
                  Begin VB.TextBox TxtOperatorN 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   2520
                     TabIndex        =   71
                     Top             =   240
                     Width           =   1635
                  End
                  Begin MSDataListLib.DataCombo DcbEquepment 
                     Height          =   315
                     Left            =   5280
                     TabIndex        =   73
                     Top             =   240
                     Width           =   3135
                     _ExtentX        =   5530
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcbBranchFrom 
                     Height          =   315
                     Left            =   5280
                     TabIndex        =   74
                     Top             =   600
                     Width           =   3135
                     _ExtentX        =   5530
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic7 
                     Height          =   435
                     Left            =   120
                     TabIndex        =   75
                     TabStop         =   0   'False
                     Top             =   600
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
                        TabIndex        =   83
                        Top             =   0
                        Width           =   300
                     End
                     Begin VB.TextBox txtLetter4 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   1155
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   82
                        Top             =   0
                        Width           =   360
                     End
                     Begin VB.TextBox txtNum3 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   270
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   81
                        Top             =   0
                        Width           =   300
                     End
                     Begin VB.TextBox txtNum2 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   480
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   80
                        Top             =   0
                        Width           =   330
                     End
                     Begin VB.TextBox txtNum1 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   795
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   79
                        Top             =   0
                        Width           =   360
                     End
                     Begin VB.TextBox txtLetter3 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   1440
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   78
                        Top             =   0
                        Width           =   315
                     End
                     Begin VB.TextBox txtLetter2 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   1710
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   77
                        Top             =   0
                        Width           =   240
                     End
                     Begin VB.TextBox txtLetter1 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   1935
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   76
                        Top             =   0
                        Width           =   285
                     End
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„⁄œÂ"
                     Height          =   285
                     Index           =   29
                     Left            =   7800
                     TabIndex        =   87
                     Top             =   240
                     Width           =   1125
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÃÂ…"
                     Height          =   285
                     Index           =   2
                     Left            =   7800
                     TabIndex        =   86
                     Top             =   600
                     Width           =   1125
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·—ﬁ„ «· ‘€Ì·Ì"
                     Height          =   285
                     Index           =   66
                     Left            =   4200
                     TabIndex        =   85
                     Top             =   240
                     Width           =   1005
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «··ÊÕ…"
                     Height          =   285
                     Index           =   67
                     Left            =   4080
                     TabIndex        =   84
                     Top             =   600
                     Width           =   1005
                  End
               End
               Begin VB.Frame Frame5 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·„ «·„⁄œ…"
                  Height          =   975
                  Left            =   3360
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   1080
                  Width           =   6855
                  Begin VB.TextBox Text6 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   4710
                     RightToLeft     =   -1  'True
                     TabIndex        =   66
                     Top             =   240
                     Width           =   1065
                  End
                  Begin VB.TextBox TxtDrievName 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   65
                     Top             =   600
                     Width           =   5655
                  End
                  Begin XtremeSuiteControls.RadioButton ChDrievType 
                     Height          =   255
                     Index           =   0
                     Left            =   5280
                     TabIndex        =   67
                     Top             =   240
                     Width           =   1455
                     _Version        =   786432
                     _ExtentX        =   2566
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "„ÊŸ›"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcbDrievID 
                     Bindings        =   "FrmOrderMaintin.frx":5830
                     Height          =   315
                     Left            =   120
                     TabIndex        =   68
                     Top             =   240
                     Width           =   4575
                     _ExtentX        =   8070
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
                  Begin XtremeSuiteControls.RadioButton ChDrievType 
                     Height          =   255
                     Index           =   1
                     Left            =   5280
                     TabIndex        =   69
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
               End
               Begin VB.Frame Frame6 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁ«∆œ «·„⁄œ…"
                  Height          =   975
                  Left            =   10320
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   1080
                  Width           =   7215
                  Begin VB.TextBox TxtLeaderName 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   60
                     Top             =   600
                     Width           =   5775
                  End
                  Begin VB.TextBox Text1 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   4830
                     RightToLeft     =   -1  'True
                     TabIndex        =   59
                     Top             =   240
                     Width           =   1065
                  End
                  Begin XtremeSuiteControls.RadioButton ChLeaderType 
                     Height          =   255
                     Index           =   0
                     Left            =   5640
                     TabIndex        =   61
                     Top             =   240
                     Width           =   1455
                     _Version        =   786432
                     _ExtentX        =   2566
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "„ÊŸ›"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcbLeaderID 
                     Bindings        =   "FrmOrderMaintin.frx":5845
                     Height          =   315
                     Left            =   120
                     TabIndex        =   62
                     Top             =   240
                     Width           =   4575
                     _ExtentX        =   8070
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
                  Begin XtremeSuiteControls.RadioButton ChLeaderType 
                     Height          =   255
                     Index           =   1
                     Left            =   5640
                     TabIndex        =   63
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
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H00E2E9E9&
                  Height          =   1095
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   0
                  Width           =   9015
                  Begin VB.TextBox TxtJiha 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   120
                     MaxLength       =   10
                     TabIndex        =   55
                     Top             =   600
                     Width           =   8175
                  End
                  Begin VB.TextBox TxtEquepmentName 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   120
                     MaxLength       =   10
                     TabIndex        =   54
                     Top             =   240
                     Width           =   8175
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„⁄œÂ"
                     Height          =   285
                     Index           =   54
                     Left            =   7800
                     TabIndex        =   57
                     Top             =   240
                     Width           =   1125
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÃÂÂ"
                     Height          =   285
                     Index           =   26
                     Left            =   7800
                     TabIndex        =   56
                     Top             =   600
                     Width           =   1125
                  End
               End
               Begin VB.ComboBox DcbStutsMaint 
                  Height          =   315
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   480
                  Width           =   1575
               End
               Begin ImpulseButton.ISButton Accredit 
                  Height          =   630
                  Left            =   0
                  TabIndex        =   145
                  Top             =   6555
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   1111
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
               Begin MSDataListLib.DataCombo DcboEmpName 
                  Height          =   315
                  Left            =   3480
                  TabIndex        =   146
                  Top             =   120
                  Width           =   3075
                  _ExtentX        =   5424
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker startmaintenanceTime 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   147
                  Top             =   1920
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   147456002
                  CurrentDate     =   38784
               End
               Begin VSFlex8UCtl.VSFlexGrid Fgpart 
                  Height          =   1365
                  Left            =   120
                  TabIndex        =   148
                  Top             =   4680
                  Width           =   17325
                  _cx             =   30559
                  _cy             =   2408
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
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmOrderMaintin.frx":585A
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
                     Left            =   0
                     TabIndex        =   149
                     Top             =   2400
                     Width           =   1440
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid gridMaintenance 
                  Height          =   1365
                  Left            =   120
                  TabIndex        =   150
                  Top             =   3000
                  Width           =   17325
                  _cx             =   30559
                  _cy             =   2408
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
                  Cols            =   21
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmOrderMaintin.frx":59AE
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
                     Index           =   42
                     Left            =   0
                     TabIndex        =   151
                     Top             =   2400
                     Width           =   1440
                  End
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   13
                  Left            =   16680
                  TabIndex        =   152
                  Top             =   4320
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   476
                  ButtonStyle     =   1
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
                  ButtonImage     =   "FrmOrderMaintin.frx":5C8C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   10
                  Left            =   16680
                  TabIndex        =   153
                  Top             =   6120
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   476
                  ButtonStyle     =   1
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
                  ButtonImage     =   "FrmOrderMaintin.frx":6226
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSComCtl2.DTPicker EnterDate 
                  Height          =   285
                  Left            =   240
                  TabIndex        =   154
                  Top             =   840
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   503
                  _Version        =   393216
                  Format          =   147456001
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker EnterTime 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   155
                  Top             =   1200
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   147456002
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker startmaintenanceDate 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   156
                  Top             =   1560
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   147456001
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker expectedEndDate 
                  Height          =   315
                  Left            =   12120
                  TabIndex        =   195
                  Top             =   6360
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   147456001
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker expectedEndtime 
                  Height          =   315
                  Left            =   12120
                  TabIndex        =   196
                  Top             =   6720
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   147456002
                  CurrentDate     =   41640
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·Œ—ÊÃ «·„ Êﬁ⁄"
                  Height          =   285
                  Index           =   33
                  Left            =   12120
                  TabIndex        =   197
                  Top             =   6135
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ— ﬂÌ·Ê „ —"
                  Height          =   285
                  Index           =   71
                  Left            =   1920
                  TabIndex        =   191
                  Top             =   2280
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂÌ·Ê „ — «·Õ«·Ì"
                  Height          =   285
                  Index           =   70
                  Left            =   1920
                  TabIndex        =   189
                  Top             =   2640
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ«  «·ﬂ‘› «·„»œ∆Ì"
                  Height          =   405
                  Index           =   69
                  Left            =   9120
                  TabIndex        =   185
                  Top             =   2160
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ«  «·ﬁ”„"
                  Height          =   285
                  Index           =   68
                  Left            =   16320
                  TabIndex        =   183
                  Top             =   2280
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   2520
                  Index           =   62
                  Left            =   4410
                  RightToLeft     =   -1  'True
                  TabIndex        =   174
                  Top             =   1155
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”ƒÊ· «·’Ì«‰…"
                  Height          =   285
                  Index           =   3
                  Left            =   7350
                  TabIndex        =   173
                  Top             =   135
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ﬂ·›…"
                  Height          =   285
                  Index           =   31
                  Left            =   2130
                  TabIndex        =   172
                  Top             =   120
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Êﬁ  »œ«Ì… «·’Ì«‰…"
                  Height          =   285
                  Index           =   32
                  Left            =   1920
                  TabIndex        =   171
                  Top             =   1920
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ê’› «·’Ì«‰…"
                  Height          =   285
                  Index           =   34
                  Left            =   7320
                  TabIndex        =   170
                  Top             =   600
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’Ì«‰… «·„ÿ·Ê»…"
                  ForeColor       =   &H00FF0000&
                  Height          =   450
                  Index           =   41
                  Left            =   7440
                  TabIndex        =   169
                  Top             =   2760
                  Width           =   2640
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ«  «·›‰Ï"
                  Height          =   330
                  Index           =   49
                  Left            =   16320
                  TabIndex        =   168
                  Top             =   6960
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÿ⁄ €Ì«— Œ«—ÃÌ…"
                  ForeColor       =   &H00FF0000&
                  Height          =   210
                  Index           =   52
                  Left            =   7440
                  TabIndex        =   167
                  Top             =   4440
                  Width           =   3360
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì ﬁÿ⁄ «·€Ì«—"
                  Height          =   285
                  Index           =   28
                  Left            =   15240
                  TabIndex        =   166
                  Top             =   6120
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   285
                  Index           =   53
                  Left            =   13440
                  TabIndex        =   165
                  Top             =   6120
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·’Ì«‰…"
                  Height          =   285
                  Index           =   55
                  Left            =   4200
                  TabIndex        =   164
                  Top             =   4440
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   285
                  Index           =   56
                  Left            =   0
                  TabIndex        =   163
                  Top             =   4440
                  Width           =   3765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì «·⁄«„"
                  Height          =   285
                  Index           =   59
                  Left            =   15000
                  TabIndex        =   162
                  Top             =   6480
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   285
                  Index           =   60
                  Left            =   13440
                  TabIndex        =   161
                  Top             =   6480
                  Width           =   1605
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ œŒÊ· «·’Ì«‰…"
                  Height          =   285
                  Index           =   61
                  Left            =   1890
                  TabIndex        =   160
                  Top             =   855
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’Ì«‰…"
                  Height          =   285
                  Index           =   63
                  Left            =   2130
                  TabIndex        =   159
                  Top             =   480
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Êﬁ  œŒÊ· «·’Ì«‰…"
                  Height          =   285
                  Index           =   64
                  Left            =   1890
                  TabIndex        =   158
                  Top             =   1200
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ »œ«Ì…«·’Ì«‰…"
                  Height          =   285
                  Index           =   65
                  Left            =   1890
                  TabIndex        =   157
                  Top             =   1560
                  Width           =   1365
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   7350
            Left            =   18585
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   45
            Width           =   17550
            _cx             =   30956
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   7350
               Left            =   0
               TabIndex        =   177
               TabStop         =   0   'False
               Top             =   0
               Width           =   17550
               _cx             =   30956
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
               Begin VB.CommandButton showAll 
                  Caption         =   "⁄—÷ «·ﬂ·"
                  Height          =   360
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   178
                  Top             =   6495
                  Width           =   1560
               End
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid13 
                  Height          =   5955
                  Left            =   135
                  TabIndex        =   179
                  Top             =   360
                  Width           =   17280
                  _cx             =   30480
                  _cy             =   10504
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
                  BackColor       =   16777215
                  ForeColor       =   -2147483640
                  BackColorFixed  =   14871017
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483633
                  BackColorAlternate=   16777088
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483633
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483633
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   -1  'True
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   1
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
                  FormatString    =   $"FrmOrderMaintin.frx":67C0
                  ScrollTrack     =   -1  'True
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
               Begin ImpulseButton.ISButton removeRow 
                  Height          =   390
                  Left            =   14790
                  TabIndex        =   180
                  Top             =   6495
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› ”ÿ— "
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
                  ButtonImage     =   "FrmOrderMaintin.frx":686E
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton clearGridBtn 
                  Height          =   390
                  Left            =   13080
                  TabIndex        =   181
                  Top             =   6495
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·ﬂ·"
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
                  ButtonImage     =   "FrmOrderMaintin.frx":6E08
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   7350
            Left            =   18885
            TabIndex        =   176
            TabStop         =   0   'False
            Top             =   45
            Width           =   17550
            _cx             =   30956
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
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   7350
            Left            =   19185
            TabIndex        =   199
            TabStop         =   0   'False
            Top             =   45
            Width           =   17550
            _cx             =   30956
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
            Begin VB.TextBox alarms 
               Alignment       =   1  'Right Justify
               Height          =   1185
               Left            =   120
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   218
               Top             =   5160
               Width           =   8295
            End
            Begin VB.CheckBox separatedreport1 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   12600
               RightToLeft     =   -1  'True
               TabIndex        =   216
               Top             =   2400
               Width           =   255
            End
            Begin VB.TextBox mangercomment 
               Alignment       =   1  'Right Justify
               Height          =   1185
               Left            =   8640
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   215
               Top             =   5160
               Width           =   8415
            End
            Begin VB.TextBox alarmsPeriod 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   8520
               MaxLength       =   4000
               TabIndex        =   214
               Top             =   6480
               Width           =   7155
            End
            Begin VB.TextBox carendperiod1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6120
               MaxLength       =   4000
               TabIndex        =   211
               Top             =   2400
               Width           =   2715
            End
            Begin VB.TextBox report1des1 
               Alignment       =   1  'Right Justify
               Height          =   1785
               Left            =   120
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   207
               Top             =   2760
               Width           =   16935
            End
            Begin VB.TextBox carendperiod 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6120
               MaxLength       =   4000
               TabIndex        =   206
               Top             =   120
               Width           =   2715
            End
            Begin VB.CheckBox separatedreport 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   12600
               RightToLeft     =   -1  'True
               TabIndex        =   205
               Top             =   120
               Width           =   255
            End
            Begin VB.TextBox report1des 
               Alignment       =   1  'Right Justify
               Height          =   1785
               Left            =   120
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   201
               Top             =   480
               Width           =   16935
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Ê’Ì«  „œÌ— «·Ê—‘…"
               Height          =   405
               Index           =   81
               Left            =   13320
               TabIndex        =   217
               Top             =   4680
               Width           =   3525
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·› —… «·„ÿ·Ê»…"
               Height          =   405
               Index           =   80
               Left            =   15120
               TabIndex        =   213
               Top             =   6600
               Width           =   1845
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ‰»ÌÂ«  «·„ÿ·Ê»…"
               Height          =   405
               Index           =   79
               Left            =   6360
               TabIndex        =   212
               Top             =   4800
               Width           =   1845
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Êﬁ  «·«› —«÷Ì ··«’·«Õ"
               Height          =   405
               Index           =   78
               Left            =   8880
               TabIndex        =   210
               Top             =   2520
               Width           =   1845
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ﬁ—Ì— „‰›’·"
               Height          =   405
               Index           =   77
               Left            =   11880
               TabIndex        =   209
               Top             =   2520
               Width           =   1845
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «’·«Õ «·„·Õﬁ"
               Height          =   405
               Index           =   76
               Left            =   15120
               TabIndex        =   208
               Top             =   2520
               Width           =   1845
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Êﬁ  «·«› —«÷Ì ··«’·«Õ"
               Height          =   405
               Index           =   74
               Left            =   8880
               TabIndex        =   204
               Top             =   240
               Width           =   1845
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ﬁ—Ì— „‰›’·"
               Height          =   405
               Index           =   73
               Left            =   11880
               TabIndex        =   203
               Top             =   240
               Width           =   1845
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «’·«Õ «·„⁄œ…"
               Height          =   405
               Index           =   72
               Left            =   15120
               TabIndex        =   202
               Top             =   240
               Width           =   1845
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   45
               Index           =   75
               Left            =   240
               TabIndex        =   200
               Top             =   6240
               Width           =   3765
            End
         End
      End
   End
   Begin MSDataListLib.DataCombo DCMaintenanceTypes 
      Height          =   315
      Left            =   30
      TabIndex        =   221
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «·’Ì«‰Â"
      Height          =   315
      Index           =   82
      Left            =   1890
      RightToLeft     =   -1  'True
      TabIndex        =   222
      Top             =   780
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·«„—"
      Height          =   285
      Index           =   50
      Left            =   12600
      TabIndex        =   39
      Top             =   720
      Width           =   765
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
      TabIndex        =   37
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
      Left            =   10920
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   780
      Width           =   375
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   17850
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·«„—"
      Height          =   285
      Index           =   4
      Left            =   16710
      TabIndex        =   26
      Top             =   750
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   14640
      TabIndex        =   25
      Top             =   735
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   16245
      TabIndex        =   24
      Top             =   9195
      Width           =   1380
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2430
      TabIndex        =   23
      Top             =   9150
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   570
      TabIndex        =   22
      Top             =   9150
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   -150
      TabIndex        =   21
      Top             =   9180
      Width           =   615
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1500
      TabIndex        =   20
      Top             =   9180
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   18510
      TabIndex        =   19
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmOrderMaintin"
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
Dim starteorktimeFocus As Boolean
Dim EnterDateFOCUS As Boolean
Dim EnterTimeFOCUS As Boolean
Dim EndDateFOCUS As Boolean
Dim EndTimeFOCUS As Boolean
Dim RecmaintenanceDateFOCUS As Boolean
Dim ERecmaintenanceTimeFOCUS As Boolean


Private Sub ReLineGrid()
    Dim I As Integer
    Dim IntCounter As Integer
    Dim summ As Double
   ''''///
   summ = 0
   lbl(56).Caption = 0
     With gridMaintenance

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("id")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("Ser")) = IntCounter
             summ = summ + val(.TextMatrix(I, .ColIndex("Total")))
                  End If
        Next I

    End With
    lbl(56).Caption = summ
    summ = 0
IntCounter = 0
lbl(58).Caption = 0
        With Me.vchrgrid
        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("Transaction_ID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("Ser")) = IntCounter
                .TextMatrix(I, .ColIndex("Total")) = val(.TextMatrix(I, .ColIndex("ShowQty"))) * val(.TextMatrix(I, .ColIndex("OperPrice")))
                summ = summ + val(.TextMatrix(I, .ColIndex("Total")))
             
                  End If
        Next I
    End With
    lbl(58).Caption = summ
    summ = 0
    IntCounter = 0
    lbl(53).Caption = 0
           With Me.Fgpart
        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("PartName")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("Ser")) = IntCounter
              summ = summ + val(.TextMatrix(I, .ColIndex("Total")))
             
                  End If
        Next I
    End With
    lbl(53).Caption = summ
    lbl(60).Caption = val(lbl(56).Caption) + val(lbl(58).Caption) + val(lbl(53).Caption) + val(TxtCost.Text)
    End Sub

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

'    Cn.CommitTrans
'    BeginTrans = False
'FillApprovedTable
'    Retrive (val(Me.XPTxtID.text))
'End Sub

Private Sub RemoveGridRow()
If Me.TxtModFlg.Text <> "R" Then
    With Me.gridMaintenance

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End If
    ReLineGrid
End Sub
Private Sub RemoveGridRowFgpart()
If Me.TxtModFlg.Text <> "R" Then
    With Me.Fgpart

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End If
    ReLineGrid
End Sub

Public Sub BaisedOn_Click(Index As Integer)
    DcbType.ListIndex = 0
    If BaisedOn(0).value = True Then
        TxtOrder.Enabled = True
        MaintPlan.Enabled = False
        MaintPlan.Text = ""
    ElseIf BaisedOn(1).value = True Then
        TxtOrder.Enabled = False
        MaintPlan.Enabled = True
        TxtOrder.Text = ""
    End If
End Sub

Private Sub ChDrievType_Click(Index As Integer)
If ChDrievType(0).value = True Then
Text6.Enabled = True
DcbDrievID.Enabled = True
TxtDrievName.Enabled = False
TxtDrievName.Text = ""
ElseIf ChDrievType(1).value = True Then
Text6.Enabled = False
DcbDrievID.Enabled = False
TxtDrievName.Enabled = True
DcbDrievID.BoundText = 0
Text6.Text = ""
End If
End Sub

Private Sub ChLeaderType_Click(Index As Integer)
If ChLeaderType(0).value = True Then
Text1.Enabled = True
DcbLeaderID.Enabled = True
TxtLeaderName.Enabled = False
TxtLeaderName.Text = ""
ElseIf ChLeaderType(1).value = True Then
Text1.Enabled = False
DcbLeaderID.Enabled = False
TxtLeaderName.Enabled = True
DcbLeaderID.BoundText = 0
Text1.Text = ""
End If
End Sub

Public Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            ReLineGrid
            starteorktimeFocus = False
            EnterTimeFOCUS = False
            EnterDateFOCUS = False
            EndDateFOCUS = True
            EnterDateFOCUS = True
            RecmaintenanceDateFOCUS = True
ERecmaintenanceTimeFOCUS = True

             ChLeaderType_Click (0)
             lbl_Click (0)
             Me.ChLeaderType(0).value = True
             ChDrievType_Click (0)
             Me.ChDrievType(0).value = True
            Me.DCboUserName.BoundText = user_id
           ' TxtPaymentCounts.text = 1
Dcbranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            Accredit.Enabled = True
            DCMaintenanceTypes.Enabled = False
    gridMaintenance.Clear flexClearScrollable, flexClearEverything
    vchrgrid.Clear flexClearScrollable, flexClearEverything
     Fgpart.Clear flexClearScrollable, flexClearEverything
gridMaintenance.Rows = 2
Fgpart.Rows = 2
vchrgrid.Rows = 2
            gridMaintenance.Enabled = True
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
   RecmaintenanceDateFOCUS = True
ERecmaintenanceTimeFOCUS = True

            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id
            gridMaintenance.Rows = gridMaintenance.Rows + 1
            gridMaintenance.Enabled = True
            Fgpart.Rows = Fgpart.Rows + 1
        Case 2
    
            Dim Msg As String
If Me.TxtModFlg = "N" Then

If EnterDateFOCUS = False Then

        If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«»œ „‰ «œŒ«·  «—ÌŒ   «·œŒÊ·  "
                Else
                    Msg = "íMust enter  Start Work time"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Screen.MousePointer = vbDefault
                Exit Sub
     
End If


If EnterTimeFOCUS = False Then

        If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«»œ „‰ «œŒ«· Êﬁ    «·œŒÊ·  "
                Else
                    Msg = "íMust enter  Start Work time"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Screen.MousePointer = vbDefault
                Exit Sub
     
End If


If starteorktimeFocus = False Then

        If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«»œ „‰ «œŒ«· Êﬁ  »œ«Ì… «·’Ì«‰…  "
                Else
                    Msg = "íMust enter  Start Work time"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Screen.MousePointer = vbDefault
                Exit Sub
     
End If


If SystemOptions.SAVEMAINTENANCEJOBWITHORDERORPLANONLY = True Then

If TxtOrder.Text = "" And MaintPlan.Text = "" Then

        If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "«„— «·‘€· » «¡ ⁄·Ì ÿ·» «Ê Œÿ… ›ﬁÿ"
                Else
                    Msg = "íMust enter order number or plan number"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Screen.MousePointer = vbDefault
                Exit Sub
       
       
End If

End If



End If

If EndDateFOCUS = False Then

        If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«»œ „‰ «œŒ«·  «—ÌŒ «·«‰ Â«¡    "
                Else
                    Msg = "Must enter  End Date"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Screen.MousePointer = vbDefault
                Exit Sub
     
End If

If EndTimeFOCUS = False Then

        If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«»œ „‰ «œŒ«· Êﬁ  «·«‰ Â«¡    "
                Else
                    Msg = "Must enter  End Time"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Screen.MousePointer = vbDefault
                Exit Sub
     
End If
If val(DcbStutsMaint.ListIndex) = -1 Then
MsgBox "Ì—ÃÏ  ÕœÌœ Õ«·… «·’«Ì‰…"
Exit Sub
End If
            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Dcbranch.SetFocus
           '     SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If


If ERecmaintenanceTimeFOCUS = False Then

        If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«»œ „‰ «œŒ«· Êﬁ    «· ”·Ì„  "
                Else
                    Msg = "íMust enter  Start Work time"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Screen.MousePointer = vbDefault
                Exit Sub
     
End If


If RecmaintenanceDateFOCUS = False Then

        If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«»œ „‰ «œŒ«·  «—ÌŒ   «· ”·Ì„  "
                Else
                    Msg = "íMust enter  Start Work time"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
 
     Load FrmSearchOrderMainten
    FrmSearchOrderMainten.show

Case 6
            Unload Me

        Case 7
'            ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 8
'            CalCulateParts
            
            
         Case 9, 11

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                print_report val(Me.XPTxtID.Text)
            End If
    Case 10
       RemoveGridRowFgpart
   Case 12

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                print_report2 val(Me.XPTxtID.Text)
            End If
   Case 13
       RemoveGridRow
       
       
Case 14
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                print_report2 val(Me.XPTxtID.Text), 1
            End If
Case 15
    Load FrmCarsPlan
    FrmCarsPlan.show
Case 16
    Load FrmRequerMainten
    FrmRequerMainten.show
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
    MySQL = " SELECT     dbo.TblOrderMaint.ID, dbo.TblOrderMaint.RecordDate, dbo.TblOrderMaint.BranchID, TblBranchesData_2.branch_name, TblBranchesData_2.branch_namee, "
    MySQL = MySQL & "                  dbo.TblOrderMaint.UserID, dbo.TblOrderMaint.EquepID, FixedAssets_2.Name, FixedAssets_2.namee, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Name1,"
    MySQL = MySQL & "                  TblEmployee_1.Emp_Name2, TblEmployee_1.Emp_Name3, TblEmployee_1.Emp_Name4, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee,"
    MySQL = MySQL & "                  TblEmployee_1.Emp_Namee1, TblEmployee_1.Emp_Namee2, TblEmployee_1.Emp_Namee3, dbo.TblOrderMaint.TypeMaint, dbo.TblOrderMaint.Jiha,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.Remarks, dbo.TblOrderMaint.Cost, dbo.TblOrderMaint.Des, dbo.TblOrderMaint.startmaintenanceTime, dbo.TblOrderMaint.endmaintenanceTime,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.RecmaintenanceTime, dbo.TblOrderMaint.RecmaintenanceDate, dbo.TblOrderMaint.reciverRemarks, dbo.TblOrderMaint.TechNote,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.reciverid, TblEmployee_1.Emp_Name AS ReciEmp_Name, TblEmployee_1.Emp_Name1 AS ReciEmp_Name1,"
    MySQL = MySQL & "                  TblEmployee_1.Emp_Name2 AS ReciEmp_Name2, TblEmployee_1.Emp_Name3 AS ReciEmp_Name3, TblEmployee_1.Fullcode AS ReciFullcode,"
    MySQL = MySQL & "                  TblEmployee_1.Emp_Namee4 AS ReciEmp_Namee4, TblEmployee_1.Emp_Namee3 AS ReciEmp_Namee3, TblEmployee_1.Emp_Namee2 AS ReciEmp_Namee2,"
    MySQL = MySQL & "                  TblEmployee_1.Emp_Namee1 AS ReciEmp_Namee1, TblEmployee_1.Emp_Namee AS RecieEmp_Namee, dbo.TblOrderMaint.endmaintenanceDate,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.ended, dbo.TblOrderMaint.ReqMainID, TblEmployee_1.Emp_Namee4, dbo.tblordermaintenancetypes.Qty,"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.Remarks AS RemarksDet, dbo.tblordermaintenancetypes.ID AS IDDet, dbo.tblordermaintenancetypes.ORderID,"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.maintenanceid, TblMaintenanceType_2.name AS nameMType, TblMaintenanceType_2.namee AS nameMTypeE,"
    MySQL = MySQL & "                  TblMaintenanceType_2.id AS MainID, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemNamee, dbo.Transaction_Details.showPrice,"
    MySQL = MySQL & "                  dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.ID AS TrnsID, dbo.TblOrderMaint.LeaderName, dbo.TblOrderMaint.LeaderType,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.DrievType, dbo.TblOrderMaint.DrievName, dbo.TblOrderMaint.EquepmentName, dbo.TblOrderMaint.Total, dbo.TblOrderMaint.DcbBranchFrom,"
    MySQL = MySQL & "                  TblBranchesData_1.branch_name AS Frombranch_name, TblBranchesData_1.branch_namee AS Frombranch_nameE, dbo.TblOrderMaint.LeaderID,"
    MySQL = MySQL & "                  TblEmployee_3.Emp_Name AS LeaderEmp_Name, TblEmployee_3.Fullcode AS LeaderFullcode, TblEmployee_3.Emp_Namee AS LeaderEmp_NameE,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.SuperVisor, TblEmployee_2.Emp_Name AS SuperEmp_Name, TblEmployee_2.Fullcode AS SuperFullcode,"
    MySQL = MySQL & "                  TblEmployee_2.Emp_Namee AS SuperEmp_NameE, dbo.TblOrderMaint.DrievID, TblEmployee_4.Emp_Name AS DevEmp_Name,"
    MySQL = MySQL & "                   TblEmployee_4.Fullcode AS DevFullcode, TblEmployee_4.Emp_Namee AS DevEmp_NameE, dbo.tblordermaintenancetypes.LocaMaint,"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.Company, dbo.tblordermaintenancetypes.Price, dbo.tblordermaintenancetypes.Total AS TotalDet, dbo.tblordermaintenancetypes.BillNo,"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.CusMobile, dbo.tblordermaintenancetypes.PartName, dbo.tblordermaintenancetypes.CusID, dbo.TblCustemers.CusName,"
    MySQL = MySQL & "                  dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.tblordermaintenancetypes.EmpID, TblEmployee_5.Emp_Name AS FiterEmp_Name,"
    MySQL = MySQL & "                  TblEmployee_5.Fullcode AS FiterFullcode, TblEmployee_5.Emp_Namee AS FiterEmp_NameE, dbo.tblordermaintenancetypes.SuperID,"
    MySQL = MySQL & "                  TblEmployee_6.Emp_Name AS SuperEmp_NameDet, TblEmployee_6.Fullcode AS SuperFullcodeDet, TblEmployee_6.Emp_Namee AS SuperEmp_NameDetE,"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.Transaction_ID AS Transaction_IDH, dbo.tblordermaintenancetypes.Transaction_IDDet, dbo.Transactions.Transaction_Date,"
    MySQL = MySQL & "                  dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_HijriDate,"
    MySQL = MySQL & "                  dbo.Transactions.TransactionComment, dbo.Transactions.OpOrderID, dbo.Transactions.OldOpOrderID, dbo.Transaction_Details.OperPrice,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.TotalSand, dbo.TblOrderMaint.TotalSpare, dbo.TblOrderMaint.TotalMaint, dbo.tblordermaintenancetypes.PartID, FixedAssets_1.code,"
    MySQL = MySQL & "                  FixedAssets_1.Name AS EqupName, FixedAssets_1.namee AS EqupNameE, dbo.tblordermaintenancetypes.GroupID, TblMaintenanceType_1.name AS GroupName,"
    MySQL = MySQL & "                   TblMaintenanceType_1.namee AS GroupNameE, dbo.TblItems.ItemName, dbo.tblordermaintenancetypes.TypeTrans"
    MySQL = MySQL & "     FROM         dbo.TblEmployee TblEmployee_3 RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblOrderMaint LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.Transactions ON dbo.TblOrderMaint.ID = dbo.Transactions.OpOrderID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee TblEmployee_4 ON dbo.TblOrderMaint.DrievID = TblEmployee_4.Emp_ID ON"
    MySQL = MySQL & "                  TblEmployee_2.Emp_ID = dbo.TblOrderMaint.SuperVisor LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee TblEmployee_1 ON dbo.TblOrderMaint.reciverid = TblEmployee_1.Emp_ID ON"
    MySQL = MySQL & "                  TblEmployee_3.Emp_ID = dbo.TblOrderMaint.LeaderID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblBranchesData TblBranchesData_1 ON dbo.TblOrderMaint.DcbBranchFrom = TblBranchesData_1.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.FixedAssets FixedAssets_2 ON dbo.TblOrderMaint.EquepID = FixedAssets_2.id LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblBranchesData TblBranchesData_2 ON dbo.TblOrderMaint.BranchID = TblBranchesData_2.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee TblEmployee_6 RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblCustemers RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.FixedAssets FixedAssets_1 RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblMaintenanceType TblMaintenanceType_1 ON dbo.tblordermaintenancetypes.GroupID = TblMaintenanceType_1.id ON"
    MySQL = MySQL & "                  FixedAssets_1.id = dbo.tblordermaintenancetypes.PartID ON dbo.TblCustemers.CusID = dbo.tblordermaintenancetypes.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmpDepartments ON dbo.tblordermaintenancetypes.DeptID = dbo.TblEmpDepartments.DeparmentID ON"
    MySQL = MySQL & "                  TblEmployee_6.Emp_ID = dbo.tblordermaintenancetypes.SuperID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee TblEmployee_5 ON dbo.tblordermaintenancetypes.EmpID = TblEmployee_5.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & "                   dbo.TblMaintenanceType TblMaintenanceType_2 ON dbo.tblordermaintenancetypes.maintenanceid = TblMaintenanceType_2.id ON"
    MySQL = MySQL & "                  dbo.TblOrderMaint.ID = dbo.tblordermaintenancetypes.OrderID"
    MySQL = MySQL & " Where (dbo.TblOrderMaint.id = " & val(XPTxtID.Text) & ") "
   'And (dbo.Transactions.Transaction_Type = 19)"
   'And ((dbo.Transactions.OldOpOrderID = " & val(XPTxtID.text) & ") or (dbo.Transactions.OpOrderID = " & val(XPTxtID.text) & "))"
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderMainten.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderMaintenE.rpt"
        End If
    
      '  If SystemOptions.UserInterface = ArabicInterface Then
      '      StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderMainten1.rpt"
      '  Else
      '      StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderMaintenE1.rpt"
      '  End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
      If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
      End If
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
     
  '      xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

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
Function print_report2(Optional NoteSerial As String, Optional reportno As Integer)
    Dim Fullcode As String
    Dim OperatorNo As String
    Dim BoardNO As String
    Dim OwnerName As String
    Dim OwnerName2 As String
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    MySQL = " SELECT   separatedreport,separatedreport1,mangercomment,alarms,alarmsPeriod, report1des, report1des1,carendperiod,carendperiod1,  dbo.TblOrderMaint.EnterDate , dbo.TblOrderMaint.EnterTime,   dbo.TblOrderMaint.ID, dbo.TblOrderMaint.RecordDate, dbo.TblOrderMaint.BranchID, TblBranchesData_2.branch_name, TblBranchesData_2.branch_namee, "
    MySQL = MySQL & "                  dbo.TblOrderMaint.UserID, dbo.TblOrderMaint.EquepID, FixedAssets_2.Name, FixedAssets_2.namee, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Name1,"
    MySQL = MySQL & "                  TblEmployee_1.Emp_Name2, TblEmployee_1.Emp_Name3, TblEmployee_1.Emp_Name4, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee,"
    MySQL = MySQL & "                  TblEmployee_1.Emp_Namee1, TblEmployee_1.Emp_Namee2, TblEmployee_1.Emp_Namee3, dbo.TblOrderMaint.TypeMaint, dbo.TblOrderMaint.Jiha,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.Remarks, dbo.TblOrderMaint.Cost, dbo.TblOrderMaint.Des, dbo.TblOrderMaint.startmaintenanceTime, dbo.TblOrderMaint.endmaintenanceTime,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.RecmaintenanceTime, dbo.TblOrderMaint.RecmaintenanceDate, dbo.TblOrderMaint.reciverRemarks, dbo.TblOrderMaint.TechNote,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.reciverid, TblEmployee_1.Emp_Name AS ReciEmp_Name, TblEmployee_1.Emp_Name1 AS ReciEmp_Name1,"
    MySQL = MySQL & "                  TblEmployee_1.Emp_Name2 AS ReciEmp_Name2, TblEmployee_1.Emp_Name3 AS ReciEmp_Name3, TblEmployee_1.Fullcode AS ReciFullcode,"
    MySQL = MySQL & "                  TblEmployee_1.Emp_Namee4 AS ReciEmp_Namee4, TblEmployee_1.Emp_Namee3 AS ReciEmp_Namee3, TblEmployee_1.Emp_Namee2 AS ReciEmp_Namee2,"
    MySQL = MySQL & "                  TblEmployee_1.Emp_Namee1 AS ReciEmp_Namee1, TblEmployee_1.Emp_Namee AS RecieEmp_Namee, dbo.TblOrderMaint.endmaintenanceDate,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.ended, dbo.TblOrderMaint.ReqMainID, TblEmployee_1.Emp_Namee4, dbo.tblordermaintenancetypes.Qty,"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.Remarks AS RemarksDet, dbo.tblordermaintenancetypes.ID AS IDDet, dbo.tblordermaintenancetypes.maintenanceid,"
    MySQL = MySQL & "                  TblMaintenanceType_1.name AS nameMType, TblMaintenanceType_1.namee AS nameMTypeE, TblMaintenanceType_1.id AS MainID,"
    MySQL = MySQL & "                  dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemNamee, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ShowQty,"
    MySQL = MySQL & "                  dbo.Transaction_Details.ID AS TrnsID, dbo.TblOrderMaint.LeaderName, dbo.TblOrderMaint.LeaderType, dbo.TblOrderMaint.DrievType, dbo.TblOrderMaint.DrievName,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.EquepmentName, dbo.TblOrderMaint.Total, dbo.TblOrderMaint.DcbBranchFrom, TblBranchesData_1.branch_name AS Frombranch_name,"
    MySQL = MySQL & "                  TblBranchesData_1.branch_namee AS Frombranch_nameE, dbo.TblOrderMaint.LeaderID, TblEmployee_3.Emp_Name AS LeaderEmp_Name,"
    MySQL = MySQL & "                  TblEmployee_3.Fullcode AS LeaderFullcode, TblEmployee_3.Emp_Namee AS LeaderEmp_NameE, dbo.TblOrderMaint.SuperVisor,"
    MySQL = MySQL & "                  TblEmployee_2.Emp_Name AS SuperEmp_Name, TblEmployee_2.Fullcode AS SuperFullcode, TblEmployee_2.Emp_Namee AS SuperEmp_NameE,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.DrievID, TblEmployee_4.Emp_Name AS DevEmp_Name, TblEmployee_4.Fullcode AS DevFullcode,"
    MySQL = MySQL & "                  TblEmployee_4.Emp_Namee AS DevEmp_NameE, dbo.tblordermaintenancetypes.LocaMaint, dbo.tblordermaintenancetypes.Company,"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.Price, dbo.tblordermaintenancetypes.Total AS TotalDet, dbo.tblordermaintenancetypes.BillNo,"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.CusMobile, dbo.tblordermaintenancetypes.PartName, dbo.tblordermaintenancetypes.CusID, dbo.TblCustemers.CusName,"
    MySQL = MySQL & "                  dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.tblordermaintenancetypes.EmpID, TblEmployee_5.Emp_Name AS FiterEmp_Name,"
    MySQL = MySQL & "                  TblEmployee_5.Fullcode AS FiterFullcode, TblEmployee_5.Emp_Namee AS FiterEmp_NameE, dbo.tblordermaintenancetypes.SuperID,"
    MySQL = MySQL & "                  TblEmployee_6.Emp_Name AS SuperEmp_NameDet, TblEmployee_6.Fullcode AS SuperFullcodeDet, TblEmployee_6.Emp_Namee AS SuperEmp_NameDetE,"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.Transaction_ID AS Transaction_IDH, dbo.tblordermaintenancetypes.Transaction_IDDet, dbo.Transaction_Details.OperPrice,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.TotalSand, dbo.TblOrderMaint.TotalSpare, dbo.TblOrderMaint.TotalMaint, dbo.tblordermaintenancetypes.PartID, FixedAssets_1.code,"
    MySQL = MySQL & "                  FixedAssets_1.Name AS EqupName, FixedAssets_1.namee AS EqupNameE, dbo.tblordermaintenancetypes.GroupID, TblMaintenanceType_2.name AS GroupName,"
    MySQL = MySQL & "                  TblMaintenanceType_2.namee AS GroupNameE, dbo.TblItems.ItemName, dbo.tblordermaintenancetypes.TypeTrans, dbo.tblordermaintenancetypes.Head_Details,"
    MySQL = MySQL & "                  dbo.TblOrderMaint.InitialNotes, dbo.TblOrderMaint.DeptNotes, TblCarsData_2.Fullcode AS CarFullcode, TblCarsData_2.BoardNO, TblCarsData_2.OwnerName,"
    MySQL = MySQL & "                  TblCarsData_2.OwnerName2, TblCarsData_2.OperatorN, TblCarsData_2.LastKMCounter, TblCarsData_2.Model, TblCarsData_2.VModel,"
    MySQL = MySQL & "                  dbo.TblCarModels.Model AS ModelName, dbo.TblCarModels.ModelE AS ModelNameE, TblCarsData_1.Chesis, TblCarsData_1.fixedAssetid,"
    MySQL = MySQL & "                  TblCarsData_1.Fullcode AS PartCaraFullCode, TblCarsData_1.LicenseNO, TblCarsData_1.BoardNO AS PartBoardNO, TblCarsData_1.OperatorN AS PartOperatorN,"
    MySQL = MySQL & "                  TblCarsData_1.OwnerName AS PartOwnerName, TblCarsData_1.OwnerName2 AS PartOwnerName2, dbo.TblOrderMaint.CurrKM, dbo.TblOrderMaint.LastKM,"
    MySQL = MySQL & "                  dbo.Transaction_Details.Head_Details AS TransHead_Details, dbo.Transaction_Details.OrderNo, dbo.tblordermaintenancetypes.ORderID,"
    MySQL = MySQL & "                  dbo.Transaction_Details.GroupIDMint, dbo.Transaction_Details.MintID, dbo.Transaction_Details.EqupID, dbo.Transactions.Transaction_Serial,"
    MySQL = MySQL & "                  dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, dbo.Transactions.TransactionComment,"
    MySQL = MySQL & "                  dbo.TblItems.Fullcode AS ItemFullcode, dbo.TblItems.barCodeNO"
    MySQL = MySQL & "           FROM         dbo.TblMaintenanceType TblMaintenanceType_2 RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmpDepartments RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblItems RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.Transactions INNER JOIN"
    MySQL = MySQL & "                  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes ON dbo.Transaction_Details.MintID = dbo.tblordermaintenancetypes.maintenanceid AND"
    MySQL = MySQL & "                  dbo.Transaction_Details.OrderNo = dbo.tblordermaintenancetypes.ORderID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblCustemers ON dbo.tblordermaintenancetypes.CusID = dbo.TblCustemers.CusID ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblMaintenanceType TblMaintenanceType_1 ON dbo.tblordermaintenancetypes.GroupID = TblMaintenanceType_1.id LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.FixedAssets FixedAssets_1 LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblCarsData TblCarsData_1 ON FixedAssets_1.id = TblCarsData_1.fixedAssetid ON dbo.tblordermaintenancetypes.PartID = FixedAssets_1.id ON"
    MySQL = MySQL & "                  dbo.TblEmpDepartments.DeparmentID = dbo.tblordermaintenancetypes.DeptID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee TblEmployee_6 ON dbo.tblordermaintenancetypes.SuperID = TblEmployee_6.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee TblEmployee_5 ON dbo.tblordermaintenancetypes.EmpID = TblEmployee_5.Emp_ID ON"
    MySQL = MySQL & "                  TblMaintenanceType_2.id = dbo.tblordermaintenancetypes.maintenanceid RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblCarsData TblCarsData_2 LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblCarModels ON TblCarsData_2.VModel = dbo.TblCarModels.Id RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.FixedAssets FixedAssets_2 ON TblCarsData_2.fixedAssetid = FixedAssets_2.id RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee TblEmployee_3 RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblOrderMaint LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee TblEmployee_4 ON dbo.TblOrderMaint.DrievID = TblEmployee_4.Emp_ID ON"
    MySQL = MySQL & "                  TblEmployee_2.Emp_ID = dbo.TblOrderMaint.SuperVisor LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee TblEmployee_1 ON dbo.TblOrderMaint.reciverid = TblEmployee_1.Emp_ID ON"
    MySQL = MySQL & "                  TblEmployee_3.Emp_ID = dbo.TblOrderMaint.LeaderID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblBranchesData TblBranchesData_1 ON dbo.TblOrderMaint.DcbBranchFrom = TblBranchesData_1.branch_id ON"
    MySQL = MySQL & "                  FixedAssets_2.id = dbo.TblOrderMaint.EquepID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblBranchesData TblBranchesData_2 ON dbo.TblOrderMaint.BranchID = TblBranchesData_2.branch_id ON"
    MySQL = MySQL & "                  dbo.tblordermaintenancetypes.OrderID = dbo.TblOrderMaint.ID"
    MySQL = MySQL & " Where (dbo.TblOrderMaint.id = " & val(XPTxtID.Text) & ") "
    
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderMainten2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderMainten2E.rpt"
        End If
        
        If reportno = 1 Then
        
        
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderMainten3.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOrderMainten3E.rpt"
        End If
        
        
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
       End If
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
     
  '      xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
 
    End If

    If VSFlexGrid13.Rows >= 2 Then
       RetriveInformation val(VSFlexGrid13.TextMatrix(1, VSFlexGrid13.ColIndex("PartID"))), Fullcode, OperatorNo, BoardNO, OwnerName, OwnerName2
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue GetDesMaint(val(XPTxtID.Text), 0)
    xReport.ParameterFields(5).AddCurrentValue GetDesMaint(val(XPTxtID.Text), 1)
    xReport.ParameterFields(6).AddCurrentValue Fullcode
    xReport.ParameterFields(7).AddCurrentValue OperatorNo
    xReport.ParameterFields(8).AddCurrentValue BoardNO
    xReport.ParameterFields(9).AddCurrentValue OwnerName
    xReport.ParameterFields(10).AddCurrentValue OwnerName2
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Sub RetriveInformation(Optional ID As Double, Optional ByRef Fullcode As String, Optional ByRef OperatorN As String, Optional ByRef BoardNO As String, Optional ByRef OwnerName As String, Optional ByRef OwnerName2 As String)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     fixedAssetid, Fullcode,OperatorN, BoardNO, Name, OwnerName, OwnerName2"
sql = sql & " From dbo.TblCarsData"
sql = sql & " WHERE     (fixedAssetid = " & ID & ")  "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
Fullcode = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
OperatorN = IIf(IsNull(rs2("OperatorN").value), "", rs2("OperatorN").value)
BoardNO = IIf(IsNull(rs2("BoardNO").value), "", rs2("BoardNO").value)
OwnerName = IIf(IsNull(rs2("OwnerName").value), "", rs2("OwnerName").value)
OwnerName2 = IIf(IsNull(rs2("OwnerName2").value), "", rs2("OwnerName2").value)
Else
OwnerName2 = "'"
OwnerName = ""
BoardNO = ""
OperatorN = ""
Fullcode = ""
End If
End Sub
Function GetDesMaint(Optional ID As Double, Optional TrnsType As Integer = 0) As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "SELECT     Des"
sql = sql & " From dbo.TblOrderMaint"
sql = sql & " WHERE      (ID = " & GetMinID(ID, val(DcbEquepment.BoundText), TrnsType) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetDesMaint = IIf(IsNull(rs2("Des").value), "", rs2("Des").value)
Else
GetDesMaint = ""
End If
End Function
Function GetMinID(Optional ID As Double, Optional EquepID As Double, Optional TrnsType As Integer = 0) As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
If TrnsType <> 0 Then
sql = " SELECT     min(ID) AS MinID"
Else
sql = " SELECT     max(ID) AS MinID"
End If
sql = sql & " From dbo.TblOrderMaint"
sql = sql & " Where (EquepID = " & EquepID & ")"
If TrnsType = 0 Then
sql = sql & " And (ID < " & ID & ")"
Else
sql = sql & " And (ID > " & ID & ")"
End If
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMinID = IIf(IsNull(rs2("MinID").value), 0, rs2("MinID").value)
Else
GetMinID = 0
End If
End Function

Private Sub CmdHelp_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments XPTxtID, "0703201701"

End Sub

Private Sub DcbDrievID_Change()
DcbDrievID_Click (0)
End Sub

Private Sub DcbDrievID_Click(Area As Integer)
    If val(DcbDrievID.BoundText) = 0 Then Exit Sub
      Dim EmpCode  As String
      GetEmployeeIDFromCode , , DcbDrievID.BoundText, EmpCode
      Text6.Text = EmpCode
End Sub

Private Sub DcbDrievID_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 41
        FrmEmployeeSearch.show
    End If
End Sub

Public Sub DcbEquepment_Change()
    DcbEquepment_Click (0)
End Sub

Private Sub DcbEquepment_Click(Area As Integer)
On Error Resume Next
RetriveCarsInfo val(DcbEquepment.BoundText), , , 0
If Me.TxtModFlg.Text <> "R" Then
Retrive_CarParts
End If
End Sub

Private Sub DcbEquepment_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
          FrmCasrShearches.SendForm = "OrderMaintin"
          FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub DcbLeaderID_Change()
DcbLeaderID_Click (0)
End Sub

Private Sub DcbLeaderID_Click(Area As Integer)
      If val(DcbLeaderID.BoundText) = 0 Then Exit Sub
      Dim EmpCode  As String
      GetEmployeeIDFromCode , , DcbLeaderID.BoundText, EmpCode
      Text1.Text = EmpCode
End Sub

Private Sub DcbLeaderID_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 40
        FrmEmployeeSearch.show
    End If
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 42
        FrmEmployeeSearch.show
    End If
End Sub

Private Sub DcbType_Change()
TxtOrder.Visible = False
'lbl(33).Visible = False
Frame4.Visible = False
Frame7.Visible = False
If val(DcbType.ListIndex) = 0 Then
Frame4.Visible = True
TxtOrder.Visible = True
'lbl(33).Visible = True
Else
Frame7.Visible = True
End If
End Sub

Private Sub DcbType_Click()
DcbType_Change
End Sub

Private Sub DCMaintenanceTypes_Click(Area As Integer)
Dim s As String
Dim rsDummy As New ADODB.Recordset
    If Trim(MaintPlan) <> "" And val(DCMaintenanceTypes.BoundText) <> 0 Then
        s = "Select Id From TblOrderMaint Where MaintenanceTypesLineNo = " & val(DCMaintenanceTypes.BoundText) & " and  MaintPlan = " & MaintPlan & " and Id <> " & val(XPTxtID)
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not rsDummy.EOF Then
            MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ ’Ì«‰… ¬Œ— ÕÌÀ «‰ Â–« «·—ﬁ„ „” Œœ„ ›Ï Õ—ﬂ… —ﬁ„ " & rsDummy!ID
            DCMaintenanceTypes.BoundText = ""
            DCMaintenanceTypes.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub ended_Click()
If ended.value = vbChecked Then
EndDateFOCUS = False
EndTimeFOCUS = False
Else
EndDateFOCUS = True
EndTimeFOCUS = True
End If

End Sub

Private Sub endmaintenanceDate_Click()
EndDateFOCUS = True
End Sub

Private Sub endmaintenanceTime_Click()
EndTimeFOCUS = True
End Sub

Private Sub EnterDate_Change()
EnterDateFOCUS = True

End Sub

Private Sub EnterTime_Change()
EnterTimeFOCUS = True

End Sub

Private Sub Fgpart_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With Me.Fgpart
Select Case .ColKey(Col)
  Case "qty"
              .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("qty"))) * val(.TextMatrix(Row, .ColIndex("Price")))
      Case "Price"
            .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("qty"))) * val(.TextMatrix(Row, .ColIndex("Price")))
End Select
           If Row = .Rows - 1 Then
            .Rows = .Rows + 1
           End If
End With
ReLineGrid
End Sub

Private Sub Fgpart_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fgpart
        Select Case .ColKey(Col)
       Case "Remarks"
             .ComboList = ""
             Case "qty"
             .ComboList = ""
                Case "Price"
             .ComboList = ""
                Case "Total"
             Cancel = True
                Case "Company"
             .ComboList = ""
                Case "BillNo"
             .ComboList = ""
                Case "CusMobile"
             .ComboList = ""
             
        End Select
    End With
End Sub

Private Sub gridMaintenance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
   Dim rs2 As ADODB.Recordset

    With gridMaintenance
        Select Case .ColKey(Col)
        Case "DepartmentName"
                 StrAccountCode = .ComboData
                  LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("DeptID"), False, True)
                  .TextMatrix(Row, .ColIndex("DeptID")) = StrAccountCode
        Case "supervisor"
                 StrAccountCode = .ComboData
                  LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("SuperID"), False, True)
                  .TextMatrix(Row, .ColIndex("SuperID")) = StrAccountCode
           Case "fitter"
                 StrAccountCode = .ComboData
                  LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("EmpID"), False, True)
                  .TextMatrix(Row, .ColIndex("EmpID")) = StrAccountCode
           Case "Group"
              StrAccountCode = .ComboData
              LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("GroupID"), False, True)
              .TextMatrix(Row, .ColIndex("GroupID")) = StrAccountCode
              .TextMatrix(Row, .ColIndex("id")) = 0
              .TextMatrix(Row, .ColIndex("name")) = ""
              .TextMatrix(Row, .ColIndex("Total")) = 0
              .TextMatrix(Row, .ColIndex("qty")) = 0
                
         Case "name"
              StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
              StrSQL = "Select * from TblMaintenanceType where id=" & val(.TextMatrix(Row, .ColIndex("id"))) & ""
            Set rs2 = New ADODB.Recordset
             rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If rs2.RecordCount > 0 Then
            .TextMatrix(Row, .ColIndex("Price")) = IIf(IsNull(rs2("Valuee").value), 0, rs2("Valuee").value)
            Else
            .TextMatrix(Row, .ColIndex("Price")) = 0
            End If
            .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("qty"))) * val(.TextMatrix(Row, .ColIndex("Price")))
        Case "CusName"
               StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CusID"), False, True)
                .TextMatrix(Row, .ColIndex("CusID")) = StrAccountCode
      Case "qty"
              .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("qty"))) * val(.TextMatrix(Row, .ColIndex("Price")))
      Case "Price"
            .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("qty"))) * val(.TextMatrix(Row, .ColIndex("Price")))
       Case "LocaMaint"
       .TextMatrix(Row, .ColIndex("CusName")) = ""
       .TextMatrix(Row, .ColIndex("CusID")) = ""
       
           End Select
                      If Row = .Rows - 1 Then
                     .Rows = .Rows + 1
                     End If
   End With
   ReLineGrid
End Sub

Private Sub gridMaintenance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With gridMaintenance
        Select Case .ColKey(Col)
       Case "Remarks"
             .ComboList = ""
       Case "QuickSearch"
             .ComboList = ""
             Case "qty"
             .ComboList = ""
                Case "Price"
             .ComboList = ""
                Case "Total"
             Cancel = True
                Case "Company"
             .ComboList = ""
         Case "CusName"
            If val(.TextMatrix(.Row, .ColIndex("LocaMaint"))) <> 2 Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "ÌÃ» «‰ ÌﬂÊ‰ „ﬂ«‰ «·’Ì«‰… Œ«—ÃÌ"
         Else
         MsgBox "It must be an external maintenance type"
         End If
         Cancel = True
         Exit Sub
         Else
         Cancel = False
         End If
        End Select
    End With
End Sub

Private Sub gridMaintenance_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim StrSQL As String
Dim StrComboList As String
Dim rs As New ADODB.Recordset
With Me.gridMaintenance

        Select Case .ColKey(Col)
        Case "QuickSearch"
              .TextMatrix(Row, .ColIndex("id")) = 0
              .TextMatrix(Row, .ColIndex("name")) = ""
        Case "DepartmentName"
              StrSQL = "SELECT  DISTINCT    dbo.TblEmpDepartments.DeparmentID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.Dpeterial, "
              StrSQL = StrSQL & "                    dbo.TblEmpDepartments.DeptBr"
              StrSQL = StrSQL & "  FROM         dbo.SuperTech INNER JOIN"
              StrSQL = StrSQL & "                     dbo.TblEmpDepartments ON dbo.SuperTech.DeparmentID = dbo.TblEmpDepartments.DeparmentID"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = .BuildComboList(rs, "DepartmentNamee", "DeparmentID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
                
        
     Case "supervisor"
         If val(.TextMatrix(Row, .ColIndex("DeptID"))) = 0 Then
         If SystemOptions.UserInterface = ArabicInterface Then
                     MsgBox "ÌÃ» «Œ Ì«— «·ﬁ”„ «Ê·«"
            Else
            MsgBox "Please Select Department "
            End If
           Exit Sub
           Else

            StrSQL = " SELECT DISTINCT "
            StrSQL = StrSQL & "                 dbo.Technicians1.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_ID AS Expr1,"
            StrSQL = StrSQL & "                dbo.SuperTech.id , dbo.SuperTech.DeparmentID"
            StrSQL = StrSQL & " FROM         dbo.Technicians1 INNER JOIN"
            StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Technicians1.Emp_ID = dbo.TblEmployee.Emp_ID INNER JOIN"
            StrSQL = StrSQL & "                     dbo.SuperTech ON dbo.Technicians1.DeparmentID = dbo.SuperTech.ID"
            StrSQL = StrSQL & " Where (dbo.SuperTech.DeparmentID =" & val(.TextMatrix(Row, .ColIndex("DeptID"))) & ")"

                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rs.RecordCount > 0 Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Emp_Name", "Emp_ID")
                Else
                    StrComboList = .BuildComboList(rs, "Emp_Namee", "Emp_ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 End If
         
                 .ComboList = StrComboList
               End If
    Case "fitter"

    If val(.TextMatrix(Row, .ColIndex("DeptID"))) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "ÌÃ» «Œ Ì«— «·ﬁ”„ «Ê·«"
      Else
      MsgBox "Please Select Department"
      End If
      Exit Sub
    Else
    If val(.TextMatrix(Row, .ColIndex("SuperID"))) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "ÌÃ» «Œ Ì«— «·„‘—› «Ê·«"
    Else
    MsgBox "Please Select Supervisor "
    End If
   Exit Sub
   Else
  
StrSQL = " SELECT     dbo.Technicians1.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Technicians1.Emp_ID1,"
StrSQL = StrSQL & "                       dbo.SuperTech.id , dbo.SuperTech.DeparmentID"
StrSQL = StrSQL & "  FROM         dbo.Technicians1 INNER JOIN"
 StrSQL = StrSQL & "                      dbo.SuperTech ON dbo.Technicians1.DeparmentID = dbo.SuperTech.ID INNER JOIN"
 StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Technicians1.Emp_ID1 = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "  Where (dbo.SuperTech.DeparmentID =" & val(.TextMatrix(Row, .ColIndex("DeptID"))) & ") And (dbo.Technicians1.Emp_id =" & val(.TextMatrix(Row, .ColIndex("SuperID"))) & ")"
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Emp_Name", "Emp_ID1")
                Else
                    StrComboList = .BuildComboList(rs, "Emp_Namee", "Emp_ID1")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
            
End If
End If
       Case "Group"
                StrSQL = "select * from TblMaintenanceType  where MainType=1"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
       Case "name"
                StrSQL = "select * from TblMaintenanceType  "
                StrSQL = StrSQL & " where ( MainType =0 or MainType is null) "
                If val(.TextMatrix(Row, .ColIndex("GroupID"))) <> 0 Then
               StrSQL = StrSQL & "  and  FollowID=" & val(.TextMatrix(Row, .ColIndex("GroupID"))) & "   "
               End If
              If (.TextMatrix(Row, .ColIndex("QuickSearch"))) <> "" Then
              StrSQL = StrSQL & " and( name like '%" & (.TextMatrix(Row, .ColIndex("QuickSearch"))) & "%'  or namee like '%" & (.TextMatrix(Row, .ColIndex("QuickSearch"))) & "%' )"
              ' StrSQL = StrSQL & "  and  name=" & (.TextMatrix(Row, .ColIndex("QuickSearch"))) & "   "
               End If
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
         Case "CusName"
      
                StrSQL = "select * from TblCustemers  "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "CusName", "CusID")
                Else
                    StrComboList = .BuildComboList(rs, "CusNamee", "CusID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList

       
        End Select

End With

End Sub

Private Sub Label2_Click()
newret
End Sub

Private Sub lbl_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
lbl(60).Caption = val(lbl(56).Caption) + val(lbl(58).Caption) + val(lbl(53).Caption) + val(TxtCost.Text)
End If
End Sub

Private Sub MaintPlan_Validate(Cancel As Boolean)
    Dim s As String, rsDummy As New ADODB.Recordset
       If Trim(MaintPlan) <> "" Then
        s = "Select Planid From TblCarMaintenancePlan Where Planid = " & val(MaintPlan)
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If rsDummy.EOF Then
            MaintPlan = ""
            MsgBox "Ì—ÃÏ «Œ Ì«— Œÿ… √Œ—Ï ÕÌÀ «‰ Â–« «·—ﬁ„ €Ì— „ÊÃÊœ" & MaintPlan
            
            MaintPlan.SetFocus
            
            Exit Sub
        End If
    End If
    RetriveMaintPlan MaintPlan.Text
    
End Sub

Private Sub reciverid_Change()
reciverid_Click (0)
End Sub

Private Sub reciverid_Click(Area As Integer)
     If val(reciverid.BoundText) = 0 Then Text2.Text = "": Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , reciverid.BoundText, EmpCode
    Text2.Text = EmpCode
End Sub

Private Sub reciverid_GotFocus()
RecmaintenanceDateFOCUS = False
ERecmaintenanceTimeFOCUS = False
End Sub

Private Sub reciverid_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 43
        FrmEmployeeSearch.show
    End If
End Sub

Private Sub RecmaintenanceDate_GotFocus()
RecmaintenanceDateFOCUS = True

End Sub

Private Sub RecmaintenanceTime_GotFocus()
ERecmaintenanceTimeFOCUS = True
End Sub

Private Sub startmaintenanceTime_GotFocus()
starteorktimeFocus = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text1.Text, EmpID
        DcbLeaderID.BoundText = EmpID
    End If
End Sub





Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text2.Text, EmpID
        reciverid.BoundText = EmpID
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text6.Text, EmpID
        DcbDrievID.BoundText = EmpID
    End If
End Sub

Private Sub TxtBoardNO_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , , TxtBoardNO.Text, 2
End If
End Sub

Private Sub TxtCost_Change()
If Me.TxtModFlg.Text <> "R" Then
lbl(60).Caption = val(lbl(56).Caption) + val(lbl(58).Caption) + val(lbl(53).Caption) + val(TxtCost.Text)
End If
End Sub



Private Sub TxtOperatorN_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , TxtOperatorN.Text, , 1
End If
End Sub

Private Sub TxtOrder_Change()
    If val(TxtOrder.Text) = 0 Then Exit Sub
    RetriveOrder TxtOrder.Text
If Me.TxtOrder.Text <> "" Then
Frame4.Enabled = False
Else
Frame4.Enabled = True
End If
End Sub
Sub EmptyTxt()
Me.txtNum1.Text = ""
Me.txtNum2.Text = ""
Me.txtNum3.Text = ""
Me.txtNum4.Text = ""
Me.txtLetter1.Text = ""
Me.txtLetter2.Text = ""
Me.txtLetter3.Text = ""
Me.txtLetter4.Text = ""
End Sub

Private Sub TxtOrder_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then
 'wael
    FrmSearchRequerMainten.lbltype = 1
    Load FrmSearchRequerMainten
    FrmSearchRequerMainten.show
 
End If

End Sub

'Private Sub ImgFavorites_Click()
'AddTofaforites Me.name, Me.Caption, Me.Caption
'
'End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub


Private Sub DcboEmpName_Click(Area As Integer)
On Error Resume Next
      If val(DcboEmpName.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
    End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
With gridMaintenance
   If SystemOptions.UserInterface = ArabicInterface Then
                .ColComboList(.ColIndex("LocaMaint")) = "#1;  œ«Œ·Ì|#2; Œ«—ÃÌ"
                .ColComboList(.ColIndex("Head_Details")) = "#1;  „⁄œÂ|#2; „·Õﬁ"
   
   ElseIf SystemOptions.UserInterface = EnglishInterface Then
               .ColComboList(.ColIndex("LocaMaint")) = "#1;Internal |#2; External "
               .ColComboList(.ColIndex("Head_Details")) = "#1;Head |#2; Follow "
   End If
End With

DB_CreateField "TblOrderMaint", "MaintPlan", adInteger, adColNullable, , , , False, True
DB_CreateField "TblOrderMaint", "BaisedOn", adBoolean, adColNullable, , , "", False, True


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
     If SystemOptions.UserInterface = EnglishInterface Then
          With DcbStutsMaint
     .Clear
     .AddItem "Current Reform"
     .AddItem "Ready"
     .AddItem "Exit"
     End With

        Me.DcbType.AddItem "Internal"
        Me.DcbType.AddItem "External"
        Else
          Me.DcbType.AddItem "œ«Œ·Ì"
        Me.DcbType.AddItem "Œ«—ÃÌ"
     With DcbStutsMaint
     .Clear
     .AddItem "Ã«—Ì «·«’·«Õ"
     .AddItem "Ã«Â“"
     .AddItem "Œ—Ã"
     End With
    End If
      Dim str  As String
       If SystemOptions.UserInterface = ArabicInterface Then
      str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Namee"
   Else
   str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Name"
   End If
    str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
 If SystemOptions.ShowDriverOnly = True Then
    str = str & "     where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1) "
End If
    fill_combo DcbLeaderID, str
    fill_combo DcbDrievID, str
    
    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEquipments DcbEquepment
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmployees Me.reciverid
    Dcombos.GetBranches Me.DcbBranchFrom
     If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
    
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblOrderMaint     Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.Text = "R"
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
    clearGridBtn.Caption = "Delete All"
    showAll.Caption = "Show All"
    removeRow.Caption = "Delete"
    Label1.Visible = False
    lbl(63).Caption = "Situation"
    lbl(61).Caption = "Date Entry "
    lbl(64).Caption = "Time Entry "
    Cmd(9).Caption = "Print"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    ended.Caption = "Ended"
    Cmd(6).Caption = "Exit"
    Cmd(12).Caption = "Print"
    CmdHelp.Caption = "Help"
    lbl(70).Caption = "Current KM"
    lbl(71).Caption = "Last KM"
    Me.Caption = "Ordered Maintenance"
    EleHeader.Caption = Me.Caption
    lbl(49).Caption = "Tech.Remark"
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
  lbl(3).Caption = "SuperVisor"
    lbl(33).Caption = "Based On"
    lbl(39).Caption = "Branch"
    lbl(29).Caption = "Machine"
    lbl(2).Caption = "Type"
lbl(31).Caption = "Cost"
lbl(28).Caption = "Repair Content   "
XPTab301.Caption = " Data| Exp| Parts"
lbl(26).Caption = "From'"
Accredit.Caption = "Send To Approve"
lbl(32).Caption = "Time Begining "
lbl(65).Caption = "Date Begining "
lbl(34).Caption = "Desc."
lbl(59).Caption = "Gen.Total"
lbl(50).Caption = "Type Order"
Frame5.Caption = "Deliver Car"
ChDrievType(0).RightToLeft = False
ChDrievType(1).RightToLeft = False
ChDrievType(0).Caption = "Employee"
lbl(28).Caption = "Total"
lbl(55).Caption = "Total"
lbl(52).Caption = "Spare parts"
lbl(57).Caption = "Total"
ChDrievType(1).Caption = "Non"
Frame6.Caption = "Leader "
ChLeaderType(0).RightToLeft = False
ChLeaderType(1).RightToLeft = False
ChLeaderType(0).Caption = "Employee"
ChLeaderType(1).Caption = "Non"
Cmd(11).Caption = "Print Bill"
lbl(69).Caption = "Initial Notes"
lbl(68).Caption = "Section Notes"
lbl(67).Caption = "Plate No."
lbl(66).Caption = "Oper.No"
 With Me.gridMaintenance
       .TextMatrix(0, .ColIndex("Ser")) = "Ser"
       .TextMatrix(0, .ColIndex("DepartmentName")) = "Department Name "
       .TextMatrix(0, .ColIndex("supervisor")) = "SuperVisor "
       .TextMatrix(0, .ColIndex("fitter")) = "Technical"
       .TextMatrix(0, .ColIndex("name")) = "Maintainance Type"
       .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
       .TextMatrix(0, .ColIndex("qty")) = "Quantity"
       .TextMatrix(0, .ColIndex("Price")) = "Value"
       .TextMatrix(0, .ColIndex("Total")) = "Total"
       .TextMatrix(0, .ColIndex("LocaMaint")) = "Location"
       .TextMatrix(0, .ColIndex("CusName")) = "Vendor"
       .TextMatrix(0, .ColIndex("Company")) = "Company"
       .TextMatrix(0, .ColIndex("Group")) = "Group"
       .TextMatrix(0, .ColIndex("Head_Details")) = "Type"
    End With
 With Me.Fgpart
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("PartName")) = "Part Name "
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
       .TextMatrix(0, .ColIndex("qty")) = "Quantity"
       .TextMatrix(0, .ColIndex("Price")) = "Value"
       .TextMatrix(0, .ColIndex("Total")) = "Total"
       .TextMatrix(0, .ColIndex("CusMobile")) = "Mobile"
       .TextMatrix(0, .ColIndex("BillNo")) = "Bill No."
       .TextMatrix(0, .ColIndex("company")) = "Company"
    End With
    
 With Me.vchrgrid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Order No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
         .TextMatrix(0, .ColIndex("View")) = "View"
        .TextMatrix(0, .ColIndex("TransactionComment")) = "Remarks"
        .TextMatrix(0, .ColIndex("ItemName")) = "Part Name"
        .TextMatrix(0, .ColIndex("ShowQty")) = "Quantity"
        .TextMatrix(0, .ColIndex("OperPrice")) = "Value"
        .TextMatrix(0, .ColIndex("Total")) = "Total"
    End With
    Cmd(10).Caption = "Delete"
XPTab301.TabCaption(0) = "Data"
XPTab301.TabCaption(1) = "Payment vouchers"
lbl(41).Caption = "Required Maintainance"
Cmd(13).Caption = "Delete"
lbl(35).Caption = "Bill FOr Order"
Label2.Caption = "Refresh"
lbl(43).Caption = "Date"
lbl(44).Caption = "Time"
lbl(45).Caption = "Receiver"
lbl(46).Caption = "Receiver's note"
lbl(47).Caption = "Receive Date"
lbl(48).Caption = "Receive Time"

    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"



End Sub

'

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

'

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

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
      

            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub
'
'

Private Sub vchrgrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid
End Sub
Private Sub vchrgrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With vchrgrid
Select Case .ColKey(Col)
Case "NoteSerial1"
Cancel = True
Case "Transaction_Date"
Cancel = True
Case "ItemName"
Cancel = True
Case "ShowQty"
Cancel = True
Case "Total"
Cancel = True
Case "TransactionComment"
Cancel = True
End Select
End With
End Sub

Private Sub vchrgrid_Click()
    With vchrgrid

        Select Case .Col
            Case 10

           If checkApility("FrmOut") = False Then
                        Exit Sub
                    End If
               
               FrmOut.Retrive val(.TextMatrix(.Row, .ColIndex("Transaction_ID")))

         
        End Select

    End With
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
Public Sub RetriveOrder(Optional order_no As String = "")
   If Me.TxtModFlg.Text <> "R" Then
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    StrSQL = "Select * from TblRequerMainten  where    ID='" & val(order_no) & "' and (StatusMaint=0 or StatusMaint=3)"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        DcbEquepment.BoundText = IIf(IsNull(rs("EquepID").value), "", rs("EquepID").value)
        txtDes.Text = IIf(IsNull(rs("Des").value), "", rs("Des").value)
        txtnote.Text = IIf(IsNull(rs("remarks").value), "", rs("remarks").value)
        DcbBranchFrom.BoundText = IIf(IsNull(rs("BranchIDTo").value), "", rs("BranchIDTo").value)
        DcbLeaderID.BoundText = IIf(IsNull(rs("LeaderID").value), "", rs("LeaderID").value)
        DcbDrievID.BoundText = IIf(IsNull(rs("DrievID").value), "", rs("DrievID").value)
        TxtLeaderName.Text = IIf(IsNull(rs("LeaderName").value), "", rs("LeaderName").value)
        TxtDrievName.Text = IIf(IsNull(rs("DrievName").value), "", rs("DrievName").value)
        TxtInitialNotes.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
        
          expectedEndDate.value = IIf(IsNull(rs("expectedEndDate").value), Date, rs("expectedEndDate").value)
    
 If Not IsNull(rs("expectedEndtime").value) Then
      expectedEndtime = FormatDateTime(rs("expectedEndtime").value, vbShortTime)
        Me.expectedEndtime.value = expectedEndtime
    End If
       
       
    


   If Not (IsNull(rs("DrievType").value)) Then
    If rs("DrievType").value = 1 Then
    ChDrievType(1).value = True
    Else
    ChDrievType(0).value = True
    End If
    Else
    ChDrievType(0).value = True
    End If
    If Not (IsNull(rs("LeaderType").value)) Then
    If rs("LeaderType").value = 1 Then
    ChLeaderType(1).value = True
    Else
    ChLeaderType(0).value = True
    End If
    Else
     ChLeaderType(0).value = True
    End If
    
     Else
     DcbBranchFrom.BoundText = 0
     DcbLeaderID.BoundText = 0
     DcbDrievID.BoundText = 0
     TxtLeaderName.Text = ""
     TxtDrievName.Text = ""
     DcbEquepment.BoundText = 0
     txtDes.Text = ""
     txtnote.Text = ""
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If
   
    Exit Sub
    End If
ErrTrap:

End Sub


Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    EmptyTxt
    Dim I As Integer
    Dim StrSQL As String

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
'DB_CreateField "TblOrderMaint", "MaintenanceTypesLineNo", adInteger, adColNullable, , , , False, True

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
gridMaintenance.Clear flexClearScrollable, flexClearEverything
            gridMaintenance.Rows = 2
vchrgrid.Clear flexClearScrollable, flexClearEverything
            vchrgrid.Rows = 1
  Fgpart.Clear flexClearScrollable, flexClearEverything
            Fgpart.Rows = 1
            
            
            
 TxtCurrKM.Text = IIf(IsNull(rs("CurrKM").value), "", rs("CurrKM").value)
 TxtLastKM.Text = IIf(IsNull(rs("LastKM").value), "", rs("LastKM").value)
 TxtOrder.Text = IIf(IsNull(rs("reqmainID").value), "", rs("reqmainID").value)
 TxtDeptNotes.Text = IIf(IsNull(rs("DeptNotes").value), "", rs("DeptNotes").value)
 TxtInitialNotes.Text = IIf(IsNull(rs("InitialNotes").value), "", rs("InitialNotes").value)
 XPTxtID.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
 XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
 Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
 Me.DcboEmpName.BoundText = IIf(IsNull(rs("SuperVisor").value), "", rs("SuperVisor").value)
 DcbEquepment.BoundText = IIf(IsNull(rs("EquepID").value), "", rs("EquepID").value)
 'txtRemark.text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
 TxtJiha.Text = IIf(IsNull(rs("Jiha").value), "", rs("Jiha").value)
 TxtCost.Text = IIf(IsNull(rs("Cost").value), "", rs("Cost").value)
 DcbType.ListIndex = val(IIf(IsNull(rs("TypeMaint").value), -1, rs("TypeMaint").value))
 txtDes.Text = IIf(IsNull(rs("Des").value), "", rs("Des").value)
 endmaintenanceDate.value = IIf(IsNull(rs("endmaintenanceDate").value), Date, rs("endmaintenanceDate").value)
 RecmaintenanceDate.value = IIf(IsNull(rs("RecmaintenanceDate").value), Date, rs("RecmaintenanceDate").value)
 reciverid.BoundText = IIf(IsNull(rs("reciverid").value), "", rs("reciverid").value)
 reciverRemarks.Text = IIf(IsNull(rs("reciverRemarks").value), "", rs("reciverRemarks").value)
 txtnote.Text = IIf(IsNull(rs("TechNote").value), "", rs("TechNote").value)
 Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
 lbl(60).Caption = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
 DcbStutsMaint.ListIndex = IIf(IsNull(rs("StutsMaint").value), -1, rs("StutsMaint").value)
 EnterDate.value = IIf(IsNull(rs("EnterDate").value), Date, rs("EnterDate").value)
startmaintenanceDate.value = IIf(IsNull(rs("startmaintenanceDate").value), Date, rs("startmaintenanceDate").value)
lbl(56).Caption = IIf(IsNull(rs("TotalMaint").value), 0, rs("TotalMaint").value)
lbl(53).Caption = IIf(IsNull(rs("TotalSpare").value), 0, rs("TotalSpare").value)
lbl(58).Caption = IIf(IsNull(rs("TotalSand").value), 0, rs("TotalSand").value)
TxtBoardNO.Text = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
TxtOperatorN.Text = IIf(IsNull(rs("OperatorN").value), "", rs("OperatorN").value)
'******************************************
mangercomment.Text = IIf(IsNull(rs("mangercomment").value), "", rs("mangercomment").value)
alarms.Text = IIf(IsNull(rs("alarms").value), "", rs("alarms").value)
alarmsPeriod.Text = IIf(IsNull(rs("alarmsPeriod").value), "", rs("alarmsPeriod").value)
report1des.Text = IIf(IsNull(rs("report1des").value), "", rs("report1des").value)
report1des1.Text = IIf(IsNull(rs("report1des1").value), "", rs("report1des1").value)
carendperiod.Text = IIf(IsNull(rs("carendperiod").value), "", rs("carendperiod").value)
carendperiod1.Text = IIf(IsNull(rs("carendperiod1").value), "", rs("carendperiod1").value)

MaintPlan = IIf(IsNull(rs("MaintPlan").value), "", rs("MaintPlan").value)
If val(MaintPlan) <> 0 Then
    DCMaintenanceTypes.Enabled = True
    loadDcbMaint 0, val(MaintPlan.Text)

    DCMaintenanceTypes.BoundText = val(rs!MaintenanceTypesLineNo & "")
Else
    DCMaintenanceTypes.BoundText = 0
    DCMaintenanceTypes.Enabled = False
End If

Dim mIsBaisedOn As Boolean
If IsNull(rs("BaisedOn").value) Then

    mIsBaisedOn = True
Else
    mIsBaisedOn = rs!BaisedOn = 0
End If
BaisedOn(0) = mIsBaisedOn
BaisedOn(1) = Not mIsBaisedOn





        If IsNull(rs("separatedreport").value) Then
              separatedreport.value = vbUnchecked
        Else
                If rs("separatedreport").value = 0 Then
                separatedreport.value = vbUnchecked
                Else
                separatedreport.value = vbChecked
                End If
        End If



        If IsNull(rs("separatedreport1").value) Then
              separatedreport1.value = vbUnchecked
        Else
                If rs("separatedreport1").value = 0 Then
                separatedreport1.value = vbUnchecked
                Else
                separatedreport1.value = vbChecked
                End If
        End If


 
 



        If IsNull(rs("ended").value) Then
              ended.value = vbUnchecked
        Else
                If rs("ended").value = 0 Then
                ended.value = vbUnchecked
                Else
                ended.value = vbChecked
                End If
        End If
    Dim startmaintenanceTime As Date
    Dim endmaintenanceTime As Date
    Dim RecmaintenanceTime As Date
   If Not IsNull(rs("startmaintenanceTime").value) Then
         startmaintenanceTime = FormatDateTime(rs("startmaintenanceTime").value, vbShortTime)
         Me.startmaintenanceTime.value = startmaintenanceTime
   End If
   If Not IsNull(rs("endmaintenanceTime").value) Then
        endmaintenanceTime = FormatDateTime(rs("endmaintenanceTime").value, vbShortTime)
        Me.endmaintenanceTime.value = endmaintenanceTime
   End If
   If Not IsNull(rs("RecmaintenanceTime").value) Then
        RecmaintenanceTime = FormatDateTime(rs("RecmaintenanceTime").value, vbShortTime)
        Me.RecmaintenanceTime.value = RecmaintenanceTime
   End If
    If Not IsNull(rs("EnterTime").value) Then
         startmaintenanceTime = FormatDateTime(rs("EnterTime").value, vbShortTime)
         Me.EnterTime.value = startmaintenanceTime
   End If
   ''///////////////////////
   ''04 05 2016
   Me.DcbBranchFrom.BoundText = IIf(IsNull(rs("DcbBranchFrom").value), "", rs("DcbBranchFrom").value)
   Me.DcbLeaderID.BoundText = IIf(IsNull(rs("LeaderID").value), "", rs("LeaderID").value)
   Me.TxtLeaderName.Text = IIf(IsNull(rs("LeaderName").value), "", rs("LeaderName").value)
   Me.TxtDrievName.Text = IIf(IsNull(rs("DrievName").value), "", rs("DrievName").value)
   Me.TxtEquepmentName.Text = IIf(IsNull(rs("EquepmentName").value), "", rs("EquepmentName").value)
   Me.DcbDrievID.BoundText = IIf(IsNull(rs("DrievID").value), "", rs("DrievID").value)
   If Not IsNull(rs("LeaderType").value) Then
   If rs("LeaderType").value = 1 Then
   ChLeaderType(1).value = True
   Else
   ChLeaderType(0).value = True
   End If
   Else
   ChLeaderType(0).value = True
   End If
   
    If Not IsNull(rs("DrievType").value) Then
   If rs("DrievType").value = 1 Then
   ChDrievType(1).value = True
   Else
   ChDrievType(0).value = True
   End If
   Else
   ChDrievType(0).value = True
   End If
     '**************************************************
      
  Set RsDetails = New ADODB.Recordset
StrSQL = " SELECT     dbo.TblMaintenanceType.name, dbo.TblMaintenanceType.namee, dbo.tblordermaintenancetypes.Qty, dbo.tblordermaintenancetypes.Remarks, "
StrSQL = StrSQL & "                      dbo.tblordermaintenancetypes.maintenanceid, dbo.tblordermaintenancetypes.ORderID, dbo.tblordermaintenancetypes.Total, dbo.tblordermaintenancetypes.Price,"
StrSQL = StrSQL & "                      dbo.tblordermaintenancetypes.Company, dbo.tblordermaintenancetypes.LocaMaint, dbo.tblordermaintenancetypes.TypeTrans, dbo.tblordermaintenancetypes.DeptID,"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.tblordermaintenancetypes.EmpID, TblEmployee_2.Emp_Name,"
StrSQL = StrSQL & "                      TblEmployee_2.Fullcode, TblEmployee_2.Emp_Namee, dbo.tblordermaintenancetypes.SuperID, TblEmployee_1.Emp_Name AS SubEmp_Name,"
StrSQL = StrSQL & "                      TblEmployee_1.Fullcode AS SupFullcode, TblEmployee_1.Emp_Namee AS SubEmp_NameE, dbo.tblordermaintenancetypes.CusID, dbo.TblCustemers.CusName,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.tblordermaintenancetypes.GroupID, TblMaintenanceType_1.name AS GroupName,"
StrSQL = StrSQL & "                      TblMaintenanceType_1.namee AS GroupNameE ,dbo.tblordermaintenancetypes.Head_Details "
StrSQL = StrSQL & " FROM         dbo.tblordermaintenancetypes LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblMaintenanceType TblMaintenanceType_1 ON dbo.tblordermaintenancetypes.GroupID = TblMaintenanceType_1.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.tblordermaintenancetypes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.tblordermaintenancetypes.SuperID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.tblordermaintenancetypes.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblMaintenanceType ON dbo.tblordermaintenancetypes.maintenanceid = dbo.TblMaintenanceType.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.tblordermaintenancetypes.DeptID = dbo.TblEmpDepartments.DeparmentID"
StrSQL = StrSQL & " WHERE     (dbo.tblordermaintenancetypes.ORderID = " & val(XPTxtID.Text) & ") And (dbo.tblordermaintenancetypes.TypeTrans = 0) "
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.gridMaintenance
        .Rows = .FixedRows + RsDetails.RecordCount
        For I = .FixedRows To .Rows - 1
             .TextMatrix(I, .ColIndex("Ser")) = I
             .TextMatrix(I, .ColIndex("Head_Details")) = (IIf(IsNull(RsDetails("Head_Details").value), 0, RsDetails("Head_Details").value) + 1)
             .TextMatrix(I, .ColIndex("GroupID")) = (IIf(IsNull(RsDetails("GroupID").value), 0, RsDetails("GroupID").value))
             .TextMatrix(I, .ColIndex("id")) = (IIf(IsNull(RsDetails("maintenanceid").value), 0, RsDetails("maintenanceid").value))
             .TextMatrix(I, .ColIndex("qty")) = (IIf(IsNull(RsDetails("Qty").value), 0, RsDetails("Qty").value))
             .TextMatrix(I, .ColIndex("Remarks")) = (IIf(IsNull(RsDetails("Remarks").value), 0, RsDetails("Remarks").value))
             .TextMatrix(I, .ColIndex("DeptID")) = (IIf(IsNull(RsDetails("DeptID").value), 0, RsDetails("DeptID").value))
             .TextMatrix(I, .ColIndex("CusID")) = (IIf(IsNull(RsDetails("CusID").value), 0, RsDetails("CusID").value))
             .TextMatrix(I, .ColIndex("Price")) = (IIf(IsNull(RsDetails("Price").value), 0, RsDetails("Price").value))
             .TextMatrix(I, .ColIndex("Total")) = (IIf(IsNull(RsDetails("Total").value), 0, RsDetails("Total").value))
             .TextMatrix(I, .ColIndex("EmpID")) = (IIf(IsNull(RsDetails("EmpID").value), 0, RsDetails("EmpID").value))
             .TextMatrix(I, .ColIndex("company")) = (IIf(IsNull(RsDetails("company").value), "", RsDetails("company").value))
             .TextMatrix(I, .ColIndex("SuperID")) = (IIf(IsNull(RsDetails("SuperID").value), 0, RsDetails("SuperID").value))
             .TextMatrix(I, .ColIndex("LocaMaint")) = (IIf(IsNull(RsDetails("LocaMaint").value), 0, RsDetails("LocaMaint").value))
    If SystemOptions.UserInterface = ArabicInterface Then
              .TextMatrix(I, .ColIndex("Group")) = (IIf(IsNull(RsDetails("GroupName").value), "", RsDetails("GroupName").value))
              .TextMatrix(I, .ColIndex("supervisor")) = (IIf(IsNull(RsDetails("SubEmp_Name").value), "", RsDetails("SubEmp_Name").value))
              .TextMatrix(I, .ColIndex("fitter")) = (IIf(IsNull(RsDetails("Emp_Name").value), "", RsDetails("Emp_Name").value))
              .TextMatrix(I, .ColIndex("CusName")) = (IIf(IsNull(RsDetails("CusName").value), "", RsDetails("CusName").value))
              .TextMatrix(I, .ColIndex("DepartmentName")) = (IIf(IsNull(RsDetails("DepartmentName").value), "", RsDetails("DepartmentName").value))
              .TextMatrix(I, .ColIndex("name")) = IIf(IsNull(RsDetails("name").value), "", RsDetails("name").value)
    Else
               .TextMatrix(I, .ColIndex("Group")) = (IIf(IsNull(RsDetails("GroupNameE").value), "", RsDetails("GroupNameE").value))
              .TextMatrix(I, .ColIndex("supervisor")) = (IIf(IsNull(RsDetails("SubEmp_NameE").value), "", RsDetails("SubEmp_NameE").value))
              .TextMatrix(I, .ColIndex("fitter")) = (IIf(IsNull(RsDetails("Emp_Namee").value), "", RsDetails("Emp_Namee").value))
              .TextMatrix(I, .ColIndex("CusName")) = (IIf(IsNull(RsDetails("CusNamee").value), "", RsDetails("CusNamee").value))
              .TextMatrix(I, .ColIndex("name")) = IIf(IsNull(RsDetails("namee").value), "", RsDetails("namee").value)
              .TextMatrix(I, .ColIndex("DepartmentName")) = (IIf(IsNull(RsDetails("DepartmentNamee").value), "", RsDetails("DepartmentNamee").value))
    End If
            RsDetails.MoveNext
         
       Next I
    ReLineGrid
End With
    End If
 ''/////////////////////////////////////////////
   Set RsDetails = New ADODB.Recordset
StrSQL = "SELECT     Qty, Remarks, ORderID, Total, Price, Company, TypeTrans, BillNo, CusMobile, PartName"
StrSQL = StrSQL & " From dbo.tblordermaintenancetypes"
StrSQL = StrSQL & " WHERE     (dbo.tblordermaintenancetypes.ORderID = " & val(XPTxtID.Text) & ") And (dbo.tblordermaintenancetypes.TypeTrans = 1) "
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
       With Me.Fgpart
        .Rows = .FixedRows + RsDetails.RecordCount
        For I = .FixedRows To .Rows - 1
              .TextMatrix(I, .ColIndex("Ser")) = I
              .TextMatrix(I, .ColIndex("qty")) = (IIf(IsNull(RsDetails("Qty").value), 0, RsDetails("Qty").value))
              .TextMatrix(I, .ColIndex("Remarks")) = (IIf(IsNull(RsDetails("Remarks").value), 0, RsDetails("Remarks").value))
              .TextMatrix(I, .ColIndex("Price")) = (IIf(IsNull(RsDetails("Price").value), 0, RsDetails("Price").value))
              .TextMatrix(I, .ColIndex("Total")) = (IIf(IsNull(RsDetails("Total").value), 0, RsDetails("Total").value))
              .TextMatrix(I, .ColIndex("PartName")) = (IIf(IsNull(RsDetails("PartName").value), "", RsDetails("PartName").value))
              .TextMatrix(I, .ColIndex("company")) = (IIf(IsNull(RsDetails("company").value), "", RsDetails("company").value))
              .TextMatrix(I, .ColIndex("CusMobile")) = (IIf(IsNull(RsDetails("CusMobile").value), "", RsDetails("CusMobile").value))
               .TextMatrix(I, .ColIndex("BillNo")) = (IIf(IsNull(RsDetails("BillNo").value), "", RsDetails("BillNo").value))
            RsDetails.MoveNext
       Next I
    ReLineGrid
End With
    End If
 ''//////////////////////////////////////////////
 '##############################################################################################################################################
     
    Dim rs_det As ADODB.Recordset
    Set rs_det = New ADODB.Recordset
    
    StrSQL = " SELECT     dbo.tblordermaintenancetypes.ID, dbo.tblordermaintenancetypes.PartID, dbo.FixedAssets.code, dbo.FixedAssets.Name, dbo.FixedAssets.namee"
    StrSQL = StrSQL & "     FROM         dbo.tblordermaintenancetypes LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.FixedAssets ON dbo.tblordermaintenancetypes.PartID = dbo.FixedAssets.id RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblOrderMaint ON dbo.tblordermaintenancetypes.ORderID = dbo.TblOrderMaint.ID"
    StrSQL = StrSQL & " Where tblordermaintenancetypes.TypeTrans = 2 and TblOrderMaint.ID = " & val(XPTxtID.Text) & " "
    
    rs_det.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    VSFlexGrid13.Clear
    VSFlexGrid13.Rows = 1
    If rs_det.RecordCount > 0 Then
        rs_det.MoveFirst
        With VSFlexGrid13
            .Rows = rs_det.RecordCount + 1
            For I = 1 To .Rows - 1
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs_det("ID").value), 0, rs_det("ID").value)
                .TextMatrix(I, .ColIndex("PartID")) = IIf(IsNull(rs_det("PartID").value), 0, rs_det("PartID").value)
                .TextMatrix(I, .ColIndex("Code")) = IIf(IsNull(rs_det("code").value), "", rs_det("code").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs_det("Name").value), "", rs_det("Name").value)
                Else
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs_det("namee").value), "", rs_det("namee").value)
                End If
                rs_det.MoveNext
            Next
         End With
    End If
'################################################################################################################################################

newret
ReLineGrid

        
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Function newret()
  Dim RsDetails1 As New ADODB.Recordset
Dim StrSQL As String
Dim I As Integer
vchrgrid.Clear flexClearScrollable, flexClearEverything
            vchrgrid.Rows = 2
            
            
StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_HijriDate, dbo.Transactions.TransactionComment, dbo.Transactions.OpOrderID,"
StrSQL = StrSQL & "                      dbo.Transactions.OldOpOrderID, dbo.Transaction_Details.OperPrice, dbo.Transaction_Details.ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.Item_ID,"
StrSQL = StrSQL & "                      dbo.TblItems.itemname , dbo.TblItems.ItemNamee, dbo.TblItems.fullcode , dbo.Transaction_Details.showPrice"
StrSQL = StrSQL & " FROM         dbo.TblItems RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_Type = 19) And ((dbo.Transactions.OldOpOrderID = " & val(XPTxtID.Text) & ") or (dbo.Transactions.OpOrderID = " & val(XPTxtID.Text) & "))"
    RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    

    If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
       With Me.vchrgrid
      '  RsDetails1.MoveFirst
        .Rows = .FixedRows + RsDetails1.RecordCount

        For I = .FixedRows To .Rows - 1
   'IIf(IsNull(RsDetails1("NameStatus").value), "", RsDetails1("NameStatus").value) = 1
            .TextMatrix(I, .ColIndex("Ser")) = I
            .TextMatrix(I, .ColIndex("Transaction_Date")) = (IIf(IsNull(RsDetails1("Transaction_Date").value), "", RsDetails1("Transaction_Date").value))
            .TextMatrix(I, .ColIndex("NoteSerial1")) = val(IIf(IsNull(RsDetails1("NoteSerial1").value), "", RsDetails1("NoteSerial1").value))
            .TextMatrix(I, .ColIndex("TransactionComment")) = (IIf(IsNull(RsDetails1("TransactionComment").value), "", RsDetails1("TransactionComment").value))
            .TextMatrix(I, .ColIndex("Transaction_ID")) = (IIf(IsNull(RsDetails1("Transaction_ID").value), 0, RsDetails1("Transaction_ID").value))
            .TextMatrix(I, .ColIndex("ID")) = (IIf(IsNull(RsDetails1("ID").value), 0, RsDetails1("ID").value))
            .TextMatrix(I, .ColIndex("OperPrice")) = (IIf(IsNull(RsDetails1("ShowPrice").value), 0, RsDetails1("ShowPrice").value))
            .TextMatrix(I, .ColIndex("ShowQty")) = (IIf(IsNull(RsDetails1("ShowQty").value), 0, RsDetails1("ShowQty").value))
            .TextMatrix(I, .ColIndex("Item_ID")) = (IIf(IsNull(RsDetails1("Item_ID").value), 0, RsDetails1("Item_ID").value))
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(I, .ColIndex("ItemName")) = (IIf(IsNull(RsDetails1("ItemName").value), "", RsDetails1("ItemName").value))
            Else
            .TextMatrix(I, .ColIndex("ItemName")) = (IIf(IsNull(RsDetails1("ItemNamee").value), "", RsDetails1("ItemNamee").value))
            End If
            RsDetails1.MoveNext
         
        Next I
    'ReLineGridCount
    ReLineGrid
    
End With
End If
End Function
Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim rsDummy As New ADODB.Recordset
    Dim I As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String
    Dim s As String
    
    On Error GoTo ErrTrap
    If Trim(TxtOrder) <> "" Then
        s = "Select Id From TblOrderMaint Where reqmainid = " & val(TxtOrder) & " and Id <> " & val(XPTxtID)
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not rsDummy.EOF Then
            MsgBox "Ì—ÃÏ «Œ Ì«— —ﬁ„ ÿ·» ¬Œ— ÕÌÀ «‰ Â–« «·—ﬁ„ „” Œœ„ ›Ï Õ—ﬂ… —ﬁ„ " & rsDummy!ID
            TxtOrder.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(MaintPlan) <> "" Then
        s = "Select Planid From TblCarMaintenancePlan Where Planid = " & val(MaintPlan)
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If rsDummy.EOF Then
            MsgBox "Ì—ÃÏ «Œ Ì«— Œÿ… √Œ—Ï ÕÌÀ «‰ Â–« «·—ﬁ„ €Ì— „ÊÃÊœ" & MaintPlan
            MaintPlan.SetFocus
            Exit Sub
        End If
    End If
     
    
    If Trim(MaintPlan) <> "" Then
        s = "Select Id From TblOrderMaint Where MaintenanceTypesLineNo = " & val(DCMaintenanceTypes.BoundText) & " and  MaintPlan = " & val(MaintPlan) & " and Id <> " & val(XPTxtID)
        If rsDummy.State = 1 Then rsDummy.Close
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not rsDummy.EOF Then
            MsgBox "Ì—ÃÏ «Œ Ì«— Œÿ… √Œ—Ï ÕÌÀ «‰ Â–« «·—ﬁ„ „” Œœ„ ›Ï Õ—ﬂ… —ﬁ„ " & rsDummy!ID
            MaintPlan.SetFocus
            Exit Sub
        End If
    End If
    
     
    
    If Me.TxtModFlg.Text <> "R" Then
    If val(DcbType.ListIndex) = -1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·«„—"
    Else
    MsgBox "Please Select Type Order"
    End If
    DcbType.SetFocus
    Exit Sub
    End If
        If val(Me.DcboEmpName.BoundText) = 0 Or DcboEmpName.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ  «”„ «·„”ƒÊ·..!! "
       Else
       Msg = "Please select Maintenance Manager"
       End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboEmpName.SetFocus
           ' SendKeys "{F4}"
            Exit Sub
        End If

If val(DcbType.ListIndex) = 0 Then
        If Me.DcbEquepment.Text = "" Or val(Me.DcbEquepment.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ  «”„ «·„⁄œÂ..!! "
        Else
        Msg = "Please select Equipment"
        End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcbEquepment.SetFocus
           'SendKeys "{F4}"
            Exit Sub
        End If
Else
        If Me.TxtEquepmentName.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ  «œŒ«·  «·„⁄œÂ..!! "
        Else
        Msg = "Please eneter Equipment"
        End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtEquepmentName.SetFocus
           'SendKeys "{F4}"
            Exit Sub
        End If
End If
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then

            XPTxtID.Text = CStr(new_id("TblOrderMaint", "ID", "", True))
      
        
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
           StrSQL = "Delete From tblordermaintenancetypes Where ORderID=" & val(Me.XPTxtID.Text)
             Cn.Execute StrSQL, , adExecuteNoRecords
 
        End If

        rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
        rs("TechNote").value = txtnote.Text
        rs("DeptNotes").value = TxtDeptNotes.Text
        rs("InitialNotes").value = TxtInitialNotes.Text
        rs("ID").value = val(XPTxtID.Text)
        rs("RecordDate").value = XPDtbTrans.value
        rs("startmaintenanceDate").value = startmaintenanceDate.value
        rs("SuperVisor").value = Me.DcboEmpName.BoundText
        rs("EquepID").value = val(Me.DcbEquepment.BoundText)
        rs("TypeMaint").value = val(Me.DcbType.ListIndex)
        'rs("JobTypeID").value = val(Me.DcboJobsType.BoundText)
        'rs("Remarks").value = IIf(txtRemark.text = "", Null, (txtRemark.text))
        rs("Total").value = val(lbl(60).Caption)
        rs("Jiha").value = IIf(TxtJiha.Text = "", Null, TxtJiha.Text)
        rs("Cost").value = IIf(TxtCost.Text = "", Null, val(TxtCost.Text))
        rs("UserID").value = Me.DCboUserName.BoundText
     If ended.value = vbChecked Then
        rs("ended").value = 1
     Else
        rs("ended").value = 0
    End If
   rs("Des").value = IIf(txtDes.Text = "", Null, (txtDes.Text))
   rs("EnterTime").value = FormatDateTime(Me.EnterTime.value, vbShortTime)
   rs("startmaintenanceTime").value = FormatDateTime(Me.startmaintenanceTime.value, vbShortTime)
   rs("endmaintenanceTime").value = FormatDateTime(Me.endmaintenanceTime.value, vbShortTime)
   rs("RecmaintenanceTime").value = FormatDateTime(Me.RecmaintenanceTime.value, vbShortTime)
   rs("endmaintenanceDate").value = endmaintenanceDate.value
   rs("RecmaintenanceDate").value = RecmaintenanceDate.value
   rs("reciverid").value = val(Me.reciverid.BoundText)
   rs("reciverRemarks").value = IIf(reciverRemarks.Text = "", Null, (reciverRemarks.Text))
   rs("reqmainid").value = IIf(TxtOrder.Text = "", Null, val(TxtOrder.Text))
   rs("DcbBranchFrom").value = val(DcbBranchFrom.BoundText)
   rs("LeaderID").value = val(DcbLeaderID.BoundText)
   rs("CurrKM").value = val(TxtCurrKM.Text)
   rs("LastKM").value = val(TxtLastKM.Text)
   
   If ChLeaderType(1).value = True Then
   rs("LeaderType").value = 1
   Else
   rs("LeaderType").value = 0
   End If
   rs("LeaderName").value = IIf(TxtLeaderName.Text = "", Null, (TxtLeaderName.Text))
   rs("DrievID").value = val(DcbDrievID.BoundText)
   If ChDrievType(1).value = True Then
   rs("DrievType").value = 1
   Else
   rs("DrievType").value = 0
   End If
   rs("StutsMaint").value = val(DcbStutsMaint.ListIndex)
   rs("EnterDate").value = EnterDate.value
   rs("DrievName").value = IIf(TxtDrievName.Text = "", Null, (TxtDrievName.Text))
   rs("EquepmentName").value = IIf(TxtEquepmentName.Text = "", Null, (TxtEquepmentName.Text))
   rs("TotalMaint").value = val(lbl(56).Caption)
   rs("TotalSpare").value = val(lbl(53).Caption)
   rs("TotalSand").value = val(lbl(58).Caption)
   rs("OperatorN").value = IIf(TxtOperatorN.Text = "", Null, (TxtOperatorN.Text))
   rs("BoardNO").value = IIf(TxtBoardNO.Text = "", Null, (TxtBoardNO.Text))
   rs("MaintPlan").value = IIf(MaintPlan.Text = "", Null, (MaintPlan.Text))
   rs("MaintenanceTypesLineNo").value = val(DCMaintenanceTypes.BoundText)
   
   rs!BaisedOn = IIf(BaisedOn(0), 0, 1)
   '******************************************************
   
If separatedreport.value = vbChecked Then
rs("separatedreport").value = 1
Else
rs("separatedreport").value = 0
End If


If separatedreport1.value = vbChecked Then
rs("separatedreport1").value = 1
Else
rs("separatedreport1").value = 0
End If

 
 



   rs("report1des").value = IIf(report1des.Text = "", Null, (report1des.Text))
      rs("report1des1").value = IIf(report1des1.Text = "", Null, (report1des1.Text))
         rs("carendperiod").value = IIf(carendperiod.Text = "", Null, (carendperiod.Text))
            rs("carendperiod1").value = IIf(carendperiod1.Text = "", Null, (carendperiod1.Text))
               rs("mangercomment").value = IIf(mangercomment.Text = "", Null, (mangercomment.Text))
                  rs("alarms").value = IIf(alarms.Text = "", Null, (alarms.Text))
                     rs("alarmsPeriod").value = IIf(alarmsPeriod.Text = "", Null, (alarmsPeriod.Text))
                     
On Error GoTo ErrTrap
     rs.update
 
 '**********************************************************
 
 
 Dim RsDetails1 As ADODB.Recordset

        'Dim temp As Integer
        'temp = -1
      Set RsDetails1 = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.tblordermaintenancetypes Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
With gridMaintenance
       For I = .FixedRows To .Rows - 1
        If val(.TextMatrix(I, .ColIndex("id"))) <> 0 Then
          RsDetails1.AddNew
                RsDetails1("ORderID").value = val(XPTxtID.Text)
                RsDetails1("qty").value = val(.TextMatrix(I, .ColIndex("qty")))
                RsDetails1("GroupID").value = val(.TextMatrix(I, .ColIndex("GroupID")))
                RsDetails1("maintenanceid").value = val(.TextMatrix(I, .ColIndex("id")))
                RsDetails1("Remarks").value = (.TextMatrix(I, .ColIndex("Remarks")))
                RsDetails1("DeptID").value = val(.TextMatrix(I, .ColIndex("DeptID")))
                RsDetails1("SuperID").value = val(.TextMatrix(I, .ColIndex("SuperID")))
                RsDetails1("EmpID").value = val(.TextMatrix(I, .ColIndex("EmpID")))
                RsDetails1("LocaMaint").value = val(.TextMatrix(I, .ColIndex("LocaMaint")))
                RsDetails1("company").value = (.TextMatrix(I, .ColIndex("company")))
                RsDetails1("CusID").value = val(.TextMatrix(I, .ColIndex("CusID")))
                RsDetails1("Price").value = val(.TextMatrix(I, .ColIndex("Price")))
                RsDetails1("Total").value = val(.TextMatrix(I, .ColIndex("Total")))
                RsDetails1("Head_Details").value = IIf(val(.TextMatrix(I, .ColIndex("Head_Details"))) = 0, Null, val(.TextMatrix(I, .ColIndex("Head_Details"))) - 1)
                RsDetails1("TypeTrans").value = 0
               RsDetails1.update
          End If
        Next I
End With
''/////////////////////////////////
      Set RsDetails1 = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.tblordermaintenancetypes Where (1 = -1)"
       RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
With Fgpart
       For I = .FixedRows To .Rows - 1
        If (.TextMatrix(I, .ColIndex("PartName"))) <> "" Then
          RsDetails1.AddNew
                RsDetails1("ORderID").value = val(XPTxtID.Text)
                RsDetails1("qty").value = val(.TextMatrix(I, .ColIndex("qty")))
                RsDetails1("PartName").value = (.TextMatrix(I, .ColIndex("PartName")))
                RsDetails1("Remarks").value = (.TextMatrix(I, .ColIndex("Remarks")))
                RsDetails1("Price").value = val(.TextMatrix(I, .ColIndex("Price")))
                RsDetails1("Total").value = val(.TextMatrix(I, .ColIndex("Total")))
                RsDetails1("BillNo").value = (.TextMatrix(I, .ColIndex("BillNo")))
                RsDetails1("company").value = (.TextMatrix(I, .ColIndex("company")))
                RsDetails1("CusMobile").value = (.TextMatrix(I, .ColIndex("CusMobile")))
                RsDetails1("TypeTrans").value = 1
               RsDetails1.update
          End If
        Next I
End With
    Set RsDetails1 = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.tblordermaintenancetypes Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
With vchrgrid
For I = 1 To .Rows - 1
If val(.TextMatrix(I, .ColIndex("ID"))) <> 0 Then
RsDetails1.AddNew
RsDetails1("ORderID").value = val(XPTxtID.Text)
RsDetails1("Transaction_IDDet").value = val(.TextMatrix(I, .ColIndex("ID")))
RsDetails1("Transaction_ID").value = val(.TextMatrix(I, .ColIndex("Transaction_ID")))
RsDetails1("TypeTrans").value = 2
RsDetails1.update
If val(.TextMatrix(I, .ColIndex("OperPrice"))) <> 0 Then
StrSQL = " update  Transaction_Details  set OperPrice =" & val(.TextMatrix(I, .ColIndex("OperPrice"))) & " where id =" & val(.TextMatrix(I, .ColIndex("ID"))) & ""

Cn.Execute StrSQL
End If
End If
Next I
End With
'*************************************************************************************************************
'#############################################################################################################################################
        Dim rs_det As ADODB.Recordset
        Set rs_det = New ADODB.Recordset
    
        If TxtModFlg.Text = "E" Then
             Cn.Execute "delete from tblordermaintenancetypes where ORderID = " & val(Me.XPTxtID.Text) & " and tblordermaintenancetypes.TypeTrans = 2"
        End If
    
        StrSQL = "SELECT  *  From tblordermaintenancetypes"
    
        rs_det.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        With VSFlexGrid13
            For I = 1 To .Rows - 1
                rs_det.AddNew
                rs_det("TypeTrans").value = 2
                rs_det("ORderID").value = IIf(Me.XPTxtID.Text = 0, Null, Me.XPTxtID.Text)
                rs_det("PartID").value = IIf(.TextMatrix(I, .ColIndex("PartID")) = "", Null, .TextMatrix(I, .ColIndex("PartID")))
                rs_det.update
            Next
            
            If .Rows - 1 = 0 Then
            
                  rs_det.AddNew
                rs_det("TypeTrans").value = 2
                rs_det("ORderID").value = IIf(Me.XPTxtID.Text = 0, Null, Me.XPTxtID.Text)
                rs_det("PartID").value = val(DcbEquepment.BoundText)
                rs_det.update
            End If
        End With
        If val(TxtCurrKM.Text) <> 0 Then
         Cn.Execute "Update  TblCarsData set LastKMCounter=" & val(TxtCurrKM.Text) & " where fixedAssetid=" & val(DcbEquepment.BoundText) & ""
         Cn.Execute "Update  TblRequerMainten set ManualKM=" & val(TxtCurrKM.Text) & " where id=" & val(TxtOrder.Text) & ""
        
        End If
        
        
 
        
     If ended.value = vbChecked Then
        Cn.Execute "Update  TblRequerMainten set StatusMaint=2" & " where id=" & val(TxtOrder.Text)
     Else
        If val(TxtOrder) <> 0 Then
            If DcbStutsMaint.ListIndex = 0 Then
                Cn.Execute "Update  TblRequerMainten set StatusMaint=3" & " where id=" & val(TxtOrder.Text)
            End If
        End If
    End If
    Cn.Execute "Update  TblRequerMainten set LastKM=" & val(TxtCurrKM) & " where id=" & val(TxtOrder.Text)
    
    
        changePlanstatuse
'#############################################################################################################################################
        Cn.CommitTrans
        BeginTrans = False
    
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.Text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
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
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    ElseIf Err.Number = -2147217864 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "Â–« «·”‰œ  „ «· ⁄œÌ· ⁄·ÌÂ „‰ ﬁ»· „” Œœ„ ¬Œ— »—Ã«¡ «⁄«œ…  Õ„Ì· «·”‰œ »€·ﬁ «·‘«‘… Ê› ÕÂ« „—… «Œ—Ï ﬁ»· «·Õ›Ÿ" & CHR(13)
        
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Unload Me
        Me.show
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

    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                    StrSQL = "Delete From tblordermaintenancetypes Where ORderID=" & val(Me.XPTxtID.Text)
             Cn.Execute StrSQL, , adExecuteNoRecords
              Cn.Execute "delete from tblordermaintenancetypes where ORderID = " & val(Me.XPTxtID.Text) & " and tblordermaintenancetypes.TypeTrans = 2"
 
         If val(TxtLastKM.Text) <> 0 Then
         Cn.Execute "Update  TblCarsData set LastKMCounter=" & val(TxtLastKM.Text) & " where fixedAssetid=" & val(DcbEquepment.BoundText) & ""
         Cn.Execute "Update  TblRequerMainten set ManualKM=" & val(TxtLastKM.Text) & " where id=" & val(TxtOrder.Text) & ""
         
        End If
 
                rs.delete
              '  StrSQL = "Delete From TblOrderMaint Where ID=" & val(Me.XPTxtID.text)
              '  Cn.Execute StrSQL, , adExecuteNoRecords
                
           
 
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

'    End If
    
    

'End Function



'Function fillapprovData()
'Dim Num As Integer
' Dim RsDetails As New ADODB.Recordset
' Dim StrSQL As String
 
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

' If Not (RsDetails.EOF Or RsDetails.BOF) Then
'        GRID2.Rows = RsDetails.RecordCount + 1
'
'
'        For Num = 1 To RsDetails.RecordCount
'
'       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
'                 If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
'                GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
'                Else
'                 GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
'                 End If
'
'        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
'           If SystemOptions.UserInterface = ArabicInterface Then
'            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
 '         Else
''             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
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

'End Function

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
sql = sql & " where BoardNO='" & BoardNO & "'"
End If
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Me.TxtLastKM.Text = IIf(IsNull(Rs3("LastKMCounter").value), "", Rs3("LastKMCounter").value)
If Typ <> 1 Then
Me.TxtOperatorN.Text = IIf(IsNull(Rs3("OperatorN").value), "", Rs3("OperatorN").value)
End If
If Typ <> 2 Then
TxtBoardNO.Text = IIf(IsNull(Rs3("BoardNO").value), "", Rs3("BoardNO").value)
End If
If Typ <> 0 Then
DcbEquepment.BoundText = IIf(IsNull(Rs3("FixedassetId").value), 0, Rs3("FixedassetId").value)
End If
DcbLeaderID.BoundText = IIf(IsNull(Rs3("Emp_id").value), 0, Rs3("Emp_id").value)
Else
If Typ <> 1 Then
TxtOperatorN.Text = ""
End If
If Typ <> 2 Then
TxtBoardNO.Text = ""
End If
If Typ <> 0 Then
DcbEquepment.BoundText = 0
End If
End If
End If
End Sub
Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.Text = ""
If Len(txtNum4.Text) > 0 Then
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
Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.Text = ""
If Len(txtNum3.Text) > 0 Then
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
Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.Text = ""
If Len(txtNum2.Text) > 0 Then
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
Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.Text = ""
If Len(txtNum1.Text) > 0 Then
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
Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.Text = ""
If Len(txtLetter4.Text) > 0 Then
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
Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.Text = ""
If Len(txtLetter3.Text) > 0 Then
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
Private Sub txtLetter2_KeyPress(KeyAscii As Integer)
txtLetter2.Text = ""
If Len(txtLetter2.Text) > 0 Then
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
Private Sub txtLetter1_KeyPress(KeyAscii As Integer)
txtLetter1.Text = ""
If Len(txtLetter1.Text) > 0 Then
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

Private Sub txtLetter1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub Cal_Board()
    TxtBoardNO.Text = txtLetter1.Text & " " & txtLetter2.Text & " " & txtLetter3.Text & " " & txtLetter4.Text & " " & txtNum1.Text & " " & txtNum2.Text & " " & txtNum3.Text & " " & txtNum4.Text
    RetriveCarsInfo , , TxtBoardNO.Text, 2
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
    
    BaisedOn_Click 0
    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, "  «„— ‘€· ’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«„— ‘€· ’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«„— ‘€· ’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«„— ‘€· ’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«„— ‘€· ’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«„— ‘€· ’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hwnd, "«„— ‘€· ’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«„— ‘€· ’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«„— ‘€· ’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«„— ‘€· ’Ì«‰…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«„— ‘€· ’Ì«‰…", 1, 15204351, -2147483630
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
Private Sub Retrive_CarParts()
    Dim I As Integer
    Dim rs_CarParts As ADODB.Recordset
    Set rs_CarParts = New ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = " SELECT     dbo.TblCarsDataDet.ID AS PID, dbo.TblCarsDataDet.PartID, dbo.FixedAssets.Name, dbo.FixedAssets.code, dbo.FixedAssets.namee"
    StrSQL = StrSQL & "  FROM         dbo.TblCarsDataDet LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.FixedAssets ON dbo.TblCarsDataDet.PartID = dbo.FixedAssets.id"
    StrSQL = StrSQL & " Where TblCarsDataDet.EqupID = " & val(Me.DcbEquepment.BoundText) & " "
    
    rs_CarParts.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    VSFlexGrid13.Rows = 1
    If rs_CarParts.RecordCount > 0 Then
        rs_CarParts.MoveFirst
        With VSFlexGrid13
            .Rows = rs_CarParts.RecordCount + 1
            For I = 1 To .Rows - 1
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs_CarParts("PID").value), 0, rs_CarParts("PID").value)
                .TextMatrix(I, .ColIndex("PartID")) = IIf(IsNull(rs_CarParts("PartID").value), 0, rs_CarParts("PartID").value)
                .TextMatrix(I, .ColIndex("Code")) = IIf(IsNull(rs_CarParts("code").value), "", rs_CarParts("code").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs_CarParts("Name").value), "", rs_CarParts("Name").value)
                Else
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs_CarParts("namee").value), "", rs_CarParts("namee").value)
                End If
                rs_CarParts.MoveNext
            Next
         End With
    End If
End Sub
Private Sub showAll_Click()
If Me.TxtModFlg.Text <> "R" Then
Retrive_CarParts
End If
End Sub
Private Sub RemoveMyGridRow()
    With Me.VSFlexGrid13
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub
Private Sub clearMyGrid()
    VSFlexGrid13.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid13.Rows = 1
End Sub
Private Sub removeRow_Click()
    RemoveMyGridRow
End Sub
Private Sub clearGridBtn_Click()
    clearMyGrid
End Sub
Public Sub RetriveMaintPlan(PlanID As String)
   
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    
    On Error GoTo ErrTrap
    DCMaintenanceTypes.Enabled = True
    loadDcbMaint 0, val(MaintPlan.Text)
    
    StrSQL = "SELECT TblCarsData.id,TblCarMaintenancePlan.REMARKS, TblCarsData.Branch_NO, TblCarsData.code, TblCarsData.Fullcode, TblCarsData.prifix, TblCarsData.CarsTypeId, TblCarsData.LicenseNO, TblCarsData.Name, TblCarsData.BoardNO, TblCarsData.Model, "
    StrSQL = StrSQL & " TblCarsData.PurchaseDate, TblCarsData.LastKMCounter, TblCarsData.InsuranceCompanyId, TblCarsData.LicenseExpireDate, TblCarsData.Emp_id, TblCarsData.InsuranceExpireDate, TblCarsData.TestExpireDate,"
    StrSQL = StrSQL & " TblCarsData.Notes, TblCarsData.LicenseExpireDateH, TblCarsData.InsuranceExpireDateH, TblCarsData.TestExpireDateH, TblCarsData.fixedAssetid, TblCarsData.VehicleLong, TblCarsData.EquQty, TblCarsData.Capacity,"
    StrSQL = StrSQL & " TblCarsData.ContractID, TblCarsData.EndContractDate, TblCarsData.SetCount, TblCarsData.Rate, TblCarsData.EndContractDateH, TblCarsData.EndAllocationDate, TblCarsData.Rep, TblCarsData.MaxCap, TblCarsData.OperatorN,"
    StrSQL = StrSQL & " TblCarsData.EqupName, TblCarsData.TypeCar, TblCarsData.Gearno, TblCarsData.Gearno1, TblCarsData.Machineno, TblCarsData.Machineno1, TblCarsData.VType, TblCarsData.VColor, TblCarsData.VModel, TblCarsData.Chesis,"
    StrSQL = StrSQL & " TblCarsData.LocationID, TblCarsData.Total, TblCarsData.LetterCount, TblCarsData.LetterPrice, TblCarsData.FormOrignal, TblCarsData.authorizeLicense, TblCarsData.authorizeExamination, TblCarsData.cleaner,"
    StrSQL = StrSQL & " TblCarsData.sideMirror, TblCarsData.driverMirror, TblCarsData.InnerLights, TblCarsData.Pedals, TblCarsData.SunScreens, TblCarsData.Recorder, TblCarsData.Anntena, TblCarsData.Battery, TblCarsData.SpareTyre,"
    StrSQL = StrSQL & " TblCarsData.Crane, TblCarsData.CoverKey, TblCarsData.Guarantee, TblCarsData.Stickers, TblCarsData.EmpType, TblCarsData.LeaderName, TblCarsData.CounryID, TblCarsData.CityID, TblCarsData.OwnerName,"
    StrSQL = StrSQL & " TblCarsData.TrackingNo, TblCarsData.Insurance, TblCarsData.Authorization2, TblCarsData.AuthorType, TblCarsData.AuthorDate, TblCarsData.Licenses, TblCarsData.Exam, TblCarsData.KeyReserve, TblCarsData.Receipt,"
    StrSQL = StrSQL & " TblCarsData.Triangle, TblCarsData.TrackingDevice, TblCarsData.FireExtingui, TblCarsData.SubsBattery, TblCarsData.BagAmbulance, TblCarsData.Natinality, TblCarsData.Job, TblCarsData.Department,"
    StrSQL = StrSQL & " TblCarsData.DriLicenseNo, TblCarsData.DriLicenseDate, TblCarsData.DriLicense, TblCarsData.Keys, TblCarsData.InsuranceNO, TblCarsData.FlagFixedasset, TblCarsData.StutsID, TblCarsData.OwnerName2,"
    StrSQL = StrSQL & " TblEmployee.emp_name , TblEmployee.Emp_Namee"
    StrSQL = StrSQL & " FROM TblCarsData LEFT OUTER JOIN"
    StrSQL = StrSQL & " TblEmployee ON TblCarsData.Emp_id = TblEmployee.Emp_ID RIGHT OUTER JOIN"
    StrSQL = StrSQL & " TblCarMaintenancePlan ON TblCarsData.id = TblCarMaintenancePlan.CarId"
    StrSQL = StrSQL & " where TblCarMaintenancePlan.Planid = " & PlanID & " "
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        DcbEquepment.BoundText = IIf(IsNull(rs("fixedAssetid").value), "", rs("fixedAssetid").value)
        'DcbLeaderID.BoundText = IIf(IsNull(rs("Emp_Id").value), "", rs("Emp_Id").value)
        'TxtBoardNO.Text = IIf(IsNull(rs("BoarderNO").value), "", rs("BoarderNO").value)
        'TxtOperatorN.Text = IIf(IsNull(rs("OperatorN").value), "", rs("OperatorN").value)
        txtDes.Text = IIf(IsNull(rs("REMARKS").value), "", rs("REMARKS").value)
     Else
        MaintPlan = ""
        
        txtDes.Text = ""
        DcbEquepment.BoundText = 0
        MsgBox "—ﬁ„ «·Œÿ… €Ì— „”Ã· "
        'DcbLeaderID.BoundText = 0
        'TxtBoardNO.Text = ""
        'TxtOperatorN.Text = ""
    End If
    
    
    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If
   
    StrSQL = "Select Top 1 id From   TblCarMaintenancePlanDetails where AlarmINDate < = " & SQLDate(Me.EnterDate.value, True) & " And PlanId = " & val(MaintPlan)
    If rs.State = 1 Then rs.Close
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        DCMaintenanceTypes.BoundText = val(rs!ID & "")
    End If
    
    Exit Sub
ErrTrap:

End Sub

Sub loadDcbMaint(Optional ID As Double, Optional ByVal mmID As Long = 0)
DCMaintenanceTypes.BoundText = ""
Dim My_SQL As String
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "SELECT     TblCarMaintenancePlanDetails.id, name, namee"
    Else
    My_SQL = "SELECT     TblCarMaintenancePlanDetails.id, namee"
    End If
My_SQL = My_SQL & " From dbo.TblMaintenanceType"
My_SQL = My_SQL & " Inner join TblCarMaintenancePlanDetails On TblCarMaintenancePlanDetails.MaintenanceID =TblMaintenanceType.Id "
My_SQL = My_SQL & " where MainType<>1 and (done=0 Or IsNull(OrderMaintinID,0) = " & val(XPTxtID) & ")"
If ID <> 0 Then
My_SQL = My_SQL & " and FollowID=" & ID & " "
End If
If mmID <> 0 Then
    My_SQL = My_SQL & " and TblCarMaintenancePlanDetails.Planid=" & mmID & " "
End If
 
    fill_combo DCMaintenanceTypes, My_SQL
End Sub


Sub changePlanstatuse()
    Dim StrSQL As String
    If val(MaintPlan) <> 0 Then
        StrSQL = "update TblCarMaintenancePlanDetails set done=0,OrderMaintinID = 0 where  PlanId = " & val(MaintPlan) & " And OrderMaintinID = " & val(XPTxtID)
    'AlarmINDate < = " & SQLDate(Me.EnterDate.value, True) & " "
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If
'    StrSQL = "update TblCarMaintenancePlanDetails set done=1,OrderMaintinID = " & val(XPTxtID) & " where Id = " & val(DCMaintenanceTypes.BoundText) & " and PlanId = " & val(MaintPlan)
'    'AlarmINDate < = " & SQLDate(Me.EnterDate.value, True) & " "
'    Cn.Execute StrSQL, , adExecuteNoRecords
End Sub

