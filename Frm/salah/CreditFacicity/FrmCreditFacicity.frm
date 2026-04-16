VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmCreditFacicity 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·»  ”ÂÌ·«  ≈∆ „«‰Ì…"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13455
   Icon            =   "FrmCreditFacicity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   13455
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   13800
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   13800
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
      TabIndex        =   26
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   735
      Width           =   1335
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   13440
      _cx             =   23707
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
      Caption         =   "ÿ·»  ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«» ⁄„Ì·"
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
         ButtonImage     =   "FrmCreditFacicity.frx":038A
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
         ButtonImage     =   "FrmCreditFacicity.frx":0724
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
         ButtonImage     =   "FrmCreditFacicity.frx":0ABE
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
         ButtonImage     =   "FrmCreditFacicity.frx":0E58
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
         Top             =   0
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   8580
      TabIndex        =   6
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
      Left            =   2430
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8220
      Width           =   10185
      _cx             =   17965
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
         Left            =   9270
         TabIndex        =   8
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
         Left            =   8415
         TabIndex        =   9
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
         Left            =   7560
         TabIndex        =   10
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
         Left            =   6720
         TabIndex        =   11
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
         Left            =   5865
         TabIndex        =   12
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
         Left            =   240
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
         Left            =   1095
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
         Left            =   5040
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
         Left            =   4200
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   8
         Left            =   2040
         TabIndex        =   117
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄Â  ’œÌﬁ «·€—›… «· Ã«—Ì…"
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
      Left            =   8700
      TabIndex        =   15
      Top             =   7800
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
      Left            =   13800
      TabIndex        =   16
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
      TabIndex        =   27
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
      Bindings        =   "FrmCreditFacicity.frx":11F2
      Height          =   315
      Left            =   4680
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
      Height          =   6615
      Left            =   120
      TabIndex        =   37
      Top             =   1080
      Width           =   13320
      _cx             =   23495
      _cy             =   11668
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
      Caption         =   " ”ÂÌ·«  ≈∆ „«‰Ì…|Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "FrmCreditFacicity.frx":1207
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6150
         Left            =   13965
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   45
         Width           =   13230
         _cx             =   23336
         _cy             =   10848
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
            FormatString    =   $"FrmCreditFacicity.frx":15A1
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
            TabIndex        =   49
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
            Visible         =   0   'False
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6150
         Index           =   15
         Left            =   45
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   45
         Width           =   13230
         _cx             =   23336
         _cy             =   10848
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
         _GridInfo       =   $"FrmCreditFacicity.frx":16ED
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6120
            Index           =   16
            Left            =   15
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   15
            Width           =   13200
            _cx             =   23283
            _cy             =   10795
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
            Begin XtremeSuiteControls.GroupBox lblinformationsh 
               Height          =   2550
               Left            =   0
               TabIndex        =   87
               Top             =   2565
               Width           =   6255
               _Version        =   786432
               _ExtentX        =   11033
               _ExtentY        =   4498
               _StockProps     =   79
               Caption         =   "„⁄·Ê„«  ⁄Ì‰«  «·⁄—÷"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
               Begin VB.ComboBox ComStopMD 
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   2160
                  Width           =   1575
               End
               Begin VB.TextBox TxtStopAcc 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   2160
                  Width           =   915
               End
               Begin XtremeSuiteControls.GroupBox lbltypesshow 
                  Height          =   1095
                  Left            =   0
                  TabIndex        =   96
                  Top             =   960
                  Width           =   6255
                  _Version        =   786432
                  _ExtentX        =   11033
                  _ExtentY        =   1931
                  _StockProps     =   79
                  Caption         =   "‰Ê⁄ ⁄Ì‰«  «·⁄—÷"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
                  Begin VB.TextBox TxtShowTy3 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   100
                     Top             =   600
                     Width           =   1965
                  End
                  Begin VB.TextBox TxtShowTy4 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   99
                     Top             =   600
                     Width           =   1965
                  End
                  Begin VB.TextBox TxtShowTy2 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   98
                     Top             =   240
                     Width           =   1965
                  End
                  Begin VB.TextBox TxtShowTy1 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   97
                     Top             =   240
                     Width           =   1965
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·‰Ê⁄ «·—«»⁄"
                     Height          =   375
                     Index           =   31
                     Left            =   2160
                     TabIndex        =   104
                     Top             =   600
                     Width           =   825
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·‰Ê⁄ «·À«·À"
                     Height          =   375
                     Index           =   29
                     Left            =   5160
                     TabIndex        =   103
                     Top             =   600
                     Width           =   825
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·‰Ê⁄ «·À«‰Ì"
                     Height          =   375
                     Index           =   28
                     Left            =   2160
                     TabIndex        =   102
                     Top             =   240
                     Width           =   825
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·‰Ê⁄ «·«Ê·"
                     Height          =   375
                     Index           =   26
                     Left            =   5160
                     TabIndex        =   101
                     Top             =   240
                     Width           =   825
                  End
               End
               Begin VB.TextBox TxtShowAmount 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   600
                  Width           =   1965
               End
               Begin VB.TextBox TxtShowNO 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   600
                  Width           =   1965
               End
               Begin VB.TextBox TxtWordAmount 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   240
                  Width           =   1965
               End
               Begin VB.TextBox TxtAmount 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3120
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   240
                  Width           =   1965
               End
               Begin VB.Label lblstopacc 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ì „ Êﬁ› «·Õ”«» ›Ì Õ«·… «· Êﬁ› ⁄‰ «·”Õ» ·„œ…"
                  Height          =   330
                  Left            =   2790
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   2160
                  Width           =   3375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·⁄Ì‰« "
                  Height          =   375
                  Index           =   24
                  Left            =   2040
                  TabIndex        =   94
                  Top             =   600
                  Width           =   1065
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·⁄Ì‰«  "
                  Height          =   375
                  Index           =   23
                  Left            =   5040
                  TabIndex        =   92
                  Top             =   600
                  Width           =   1140
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„… »«·Õ—Ê›"
                  Height          =   375
                  Index           =   22
                  Left            =   2040
                  TabIndex        =   90
                  Top             =   240
                  Width           =   1065
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·”ÕÊ»« "
                  Height          =   375
                  Index           =   21
                  Left            =   5160
                  TabIndex        =   88
                  Top             =   240
                  Width           =   1020
               End
            End
            Begin XtremeSuiteControls.GroupBox lblBankinformation 
               Height          =   1065
               Left            =   0
               TabIndex        =   78
               Top             =   5115
               Width           =   6315
               _Version        =   786432
               _ExtentX        =   11139
               _ExtentY        =   1879
               _StockProps     =   79
               Caption         =   "„⁄·Ê„«  »‰ﬂÌ…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
               Begin VB.TextBox TxtAccOficer 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   600
                  Width           =   1965
               End
               Begin VB.TextBox TxtAccNo 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   600
                  Width           =   1965
               End
               Begin VB.TextBox TxtBankBranch 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   240
                  Width           =   1965
               End
               Begin VB.TextBox TxtBankname 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3240
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   240
                  Width           =   1965
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„ÊŸ› «·Õ”«»« "
                  Height          =   375
                  Index           =   20
                  Left            =   2040
                  TabIndex        =   85
                  Top             =   600
                  Width           =   1140
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ”«»"
                  Height          =   375
                  Index           =   19
                  Left            =   5040
                  TabIndex        =   83
                  Top             =   600
                  Width           =   1020
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   375
                  Index           =   18
                  Left            =   2400
                  TabIndex        =   81
                  Top             =   240
                  Width           =   780
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·»‰ﬂ"
                  Height          =   375
                  Index           =   15
                  Left            =   5040
                  TabIndex        =   79
                  Top             =   240
                  Width           =   1020
               End
            End
            Begin VB.TextBox TxtCCNO 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   2085
               Width           =   2700
            End
            Begin VB.ComboBox ComMD 
               Height          =   315
               Left            =   3975
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   2085
               Width           =   1815
            End
            Begin VB.TextBox TxtLong 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5910
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   2085
               Width           =   1155
            End
            Begin XtremeSuiteControls.GroupBox LblFg 
               Height          =   1875
               Left            =   6240
               TabIndex        =   70
               Top             =   2550
               Width           =   6975
               _Version        =   786432
               _ExtentX        =   12303
               _ExtentY        =   3307
               _StockProps     =   79
               Caption         =   "«·«‘Œ«’ «·„›Ê÷Ê‰ »«· ÊﬁÌ⁄ ⁄·Ï «Ê«„— «·‘—«¡"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
               Begin VSFlex8Ctl.VSFlexGrid fg 
                  Height          =   1215
                  Left            =   120
                  TabIndex        =   71
                  Top             =   240
                  Width           =   6720
                  _cx             =   11853
                  _cy             =   2143
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
                  FormatString    =   $"FrmCreditFacicity.frx":1721
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
                  Index           =   21
                  Left            =   6120
                  TabIndex        =   118
                  Top             =   1440
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
                  ButtonImage     =   "FrmCreditFacicity.frx":1813
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin VB.TextBox TxtAcredit 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   8190
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   2055
               Width           =   3495
            End
            Begin VB.TextBox TxtZipCOd 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   1545
               Width           =   2700
            End
            Begin VB.TextBox TxtPOBox 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   3990
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   1545
               Width           =   3075
            End
            Begin VB.TextBox TxtCRSource 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   8190
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   1545
               Width           =   1320
            End
            Begin VB.TextBox TxtCRNo 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   1545
               Width           =   2085
            End
            Begin VB.TextBox Txtphone 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   1110
               Width           =   2700
            End
            Begin VB.TextBox TxtTypeBusnis 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   645
               Width           =   2700
            End
            Begin VB.TextBox TxtFax 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   3990
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   1110
               Width           =   3075
            End
            Begin VB.TextBox TxtAddress 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   195
               Width           =   2700
            End
            Begin VB.TextBox TxtCity 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   3990
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   645
               Width           =   3075
            End
            Begin VB.TextBox TxtNameOwner 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   3990
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   195
               Width           =   3075
            End
            Begin VB.TextBox TxtEmail 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   8190
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   1110
               Width           =   3495
            End
            Begin VB.TextBox TXtStreet 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   8190
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   645
               Width           =   3495
            End
            Begin VB.TextBox TxtNameApplicant 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   8190
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   195
               Width           =   3495
            End
            Begin XtremeSuiteControls.GroupBox lblfg2 
               Height          =   1635
               Left            =   6225
               TabIndex        =   72
               Top             =   4410
               Width           =   6960
               _Version        =   786432
               _ExtentX        =   12277
               _ExtentY        =   2884
               _StockProps     =   79
               Caption         =   "«·«‘Œ«’ «·„›Ê÷Ê‰ »«” ·«„ «·»÷«⁄…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
               BorderStyle     =   1
               Begin VSFlex8Ctl.VSFlexGrid FG2 
                  Height          =   1095
                  Left            =   120
                  TabIndex        =   73
                  Top             =   240
                  Width           =   6720
                  _cx             =   11853
                  _cy             =   1931
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
                  FormatString    =   $"FrmCreditFacicity.frx":1DAD
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
                  Height          =   270
                  Index           =   10
                  Left            =   6120
                  TabIndex        =   119
                  Top             =   1320
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
                  ButtonImage     =   "FrmCreditFacicity.frx":1E9F
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„œÌ‰…"
               Height          =   210
               Index           =   3
               Left            =   7200
               TabIndex        =   116
               Top             =   780
               Width           =   855
            End
            Begin VB.Label lblowner 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·„«·ﬂ"
               Height          =   420
               Left            =   7140
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   300
               Width           =   960
            End
            Begin VB.Label lblfax 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "›«ﬂ” "
               Height          =   420
               Left            =   7215
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   1200
               Width           =   825
            End
            Begin VB.Label lblbox 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "’.»"
               Height          =   420
               Left            =   7275
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   1665
               Width           =   705
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ „ﬁœ„ «·ÿ·»"
               Height          =   255
               Index           =   2
               Left            =   11820
               TabIndex        =   112
               Top             =   285
               Width           =   1320
            End
            Begin VB.Label lbstreet 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·‘«—⁄"
               Height          =   240
               Left            =   12060
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   735
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»—Ìœ «·ﬂ —Ê‰Ì"
               Height          =   255
               Index           =   5
               Left            =   12060
               TabIndex        =   110
               Top             =   1185
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·”Ã· Ê„’œ—Â"
               Height          =   225
               Index           =   14
               Left            =   11700
               TabIndex        =   109
               Top             =   1635
               Width           =   1440
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»·€ «·„ÿ·Ê»"
               Height          =   390
               Index           =   17
               Left            =   12060
               TabIndex        =   108
               Top             =   2085
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·€—›… «· Ã«—Ì…"
               Height          =   480
               Index           =   11
               Left            =   2775
               TabIndex        =   77
               Top             =   2115
               Width           =   975
            End
            Begin VB.Label Lbllong 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„œ… «· ”ÂÌ·"
               Height          =   420
               Left            =   7170
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   2040
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—„“ «·»—ÌœÌ"
               Height          =   345
               Index           =   16
               Left            =   2775
               TabIndex        =   66
               Top             =   1710
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·⁄„·"
               Height          =   345
               Index           =   13
               Left            =   2700
               TabIndex        =   62
               Top             =   720
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄‰Ê«‰ «·„«·ﬂ"
               Height          =   255
               Index           =   10
               Left            =   2700
               TabIndex        =   55
               Top             =   285
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ·›Ê‰"
               Height          =   270
               Index           =   9
               Left            =   2790
               TabIndex        =   53
               Top             =   1230
               Width           =   945
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6120
            Index           =   9
            Left            =   15
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   15
            Width           =   13200
            _cx             =   23283
            _cy             =   10795
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
               Height          =   4590
               Left            =   3435
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   1275
               Width           =   720
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   3255
               Left            =   4335
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   1590
               Width           =   1140
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3255
               Index           =   67
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1590
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   3060
               Index           =   68
               Left            =   4155
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   2040
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
               Height          =   3720
               Index           =   69
               Left            =   3105
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1590
               Width           =   330
            End
         End
      End
   End
   Begin ImpulseButton.ISButton Accredit 
      Height          =   375
      Left            =   1080
      TabIndex        =   120
      Top             =   8280
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo DcbDetpartment 
      Height          =   315
      Left            =   240
      TabIndex        =   121
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«œ«—… «·ÿ«·»…"
      Height          =   375
      Index           =   37
      Left            =   3315
      TabIndex        =   122
      Top             =   720
      Width           =   1365
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
   Begin VB.Label lblBr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Left            =   7560
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
      Left            =   12810
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   285
      Index           =   4
      Left            =   12390
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   735
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   11445
      TabIndex        =   22
      Top             =   7875
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2550
      TabIndex        =   21
      Top             =   7950
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   810
      TabIndex        =   20
      Top             =   7950
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   210
      TabIndex        =   19
      Top             =   7980
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1860
      TabIndex        =   18
      Top             =   7980
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   13350
      TabIndex        =   17
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmCreditFacicity"
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
    Dim Sql As String
    Dim BeginTrans As Boolean
   
    SendTopost Me.Name, "TblCreditFacicity", "Id", val(DcbDetpartment.BoundText), val(dcBranch.BoundText), val(XPTxtID.Text), XPTxtID.Text
    
   rs.Resync
    
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

 
  Retrive (val(XPTxtID.Text))

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
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
              Fg2.Clear flexClearScrollable, flexClearEverything
            Fg2.Rows = 2
            Fg2.Enabled = True
        
            
              GRID2.Clear flexClearScrollable, flexClearEverything
    GRID2.Rows = 1
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
Fg.Rows = Fg.Rows + 1
            Fg.Enabled = True
            Fg2.Rows = Fg.Rows + 1
            Fg2.Enabled = True
            TxtModFlg.Text = "E"
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
        Load FrmCreditSearch
         FrmCreditSearch.show

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.Text, , 200

        Case 8
           If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                print_report val(Me.XPTxtID.Text), 1
        
        
            End If
            
            
                 Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                print_report val(Me.XPTxtID.Text)
        
        
            End If
            Case 21
            RemoveGridRow
            Case 10
            RemoveGridRow1
        
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String, Optional Index As Integer = 0)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



MySQL = "SELECT     dbo.TblCreditFacicity.ID, dbo.TblCreditFacicity.RecordDate, dbo.TblCreditFacicity.Posted, dbo.TblCreditFacicity.UserID, dbo.TblCreditFacicity.BranchID,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCreditFacicity.NameApplicant, dbo.TblCreditFacicity.NameOwner,"
MySQL = MySQL & "                        dbo.TblCreditFacicity.Street, dbo.TblCreditFacicity.City, dbo.TblCreditFacicity.Email, dbo.TblCreditFacicity.Fax, dbo.TblCreditFacicity.Phone,"
MySQL = MySQL & "                        dbo.TblCreditFacicity.CRNo, dbo.TblCreditFacicity.CRSource, dbo.TblCreditFacicity.POBox, dbo.TblCreditFacicity.ZipCode, dbo.TblCreditFacicity.Address,"
MySQL = MySQL & "                        dbo.TblCreditFacicity.TypeBusines, dbo.TblCreditFacicity.longT, dbo.TblCreditFacicity.Acredit, dbo.TblCreditFacicityDetails1.Code, dbo.TblCreditFacicityDetails1.name,"
MySQL = MySQL & "                        dbo.TblCreditFacicityDetails1.job, dbo.TblCreditFacicityDetails1.iqamano, dbo.TblCreditFacicityDetails1.nationality, dbo.TblCreditFacicityDetails1.Type,"
MySQL = MySQL & "                        dbo.TblCreditFacicity.DMY, dbo.TblCreditFacicity.CCNO, dbo.TblCreditFacicity.Amount, dbo.TblCreditFacicity.WordAmount, dbo.TblCreditFacicity.ShowNo,"
 MySQL = MySQL & "                       dbo.TblCreditFacicity.Showtype1, dbo.TblCreditFacicity.Showtype2, dbo.TblCreditFacicity.Showtype3, dbo.TblCreditFacicity.Showtype4,"
MySQL = MySQL & "                        dbo.TblCreditFacicity.StopAccount, dbo.TblCreditFacicity.StopDMY, dbo.TblCreditFacicity.BanckName, dbo.TblCreditFacicity.BanckBranch, dbo.TblCreditFacicity.AccNo,"
MySQL = MySQL & "                        dbo.TblCreditFacicity.AccOficer , dbo.TblCreditFacicity.ShowAmount"
MySQL = MySQL & "   FROM         dbo.TblCreditFacicity  LEFT OUTER join "
 MySQL = MySQL & "                       dbo.TblCreditFacicityDetails1 ON dbo.TblCreditFacicity.ID = dbo.TblCreditFacicityDetails1.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblCreditFacicity.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & "   Where (dbo.TblCreditFacicity.id =" & val(XPTxtID.Text) & ")"
 If Index = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAcreditFacicity.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAcreditFacicity.rpt"
        End If
End If
 If Index = 1 Then
   If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAcreditFacicity CC.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAcreditFacicity CC.rpt"
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
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
  '      xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
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

Private Sub DcbCarType_Click(Area As Integer)
 
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub





Private Sub Fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim StrAccountCode1 As String

    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
      Dim Rs1 As New ADODB.Recordset
    Dim StrSQL1 As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim StrComboList As String
            
    
    With Fg

        Select Case .ColKey(Col)
 
            Case "name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Code"), False, True)
                .TextMatrix(Row, .ColIndex("Code")) = StrAccountCode
                

               '     StrSQL = " SELECT * FROM  TblEmployee where Emp_Code=" & val(StrAccountCode)
              
             
            
               ' rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
               '    StrSQL1 = "Select JobTypeName From TblEmpJobsTypes where JobTypeID=" & val(StrAccountCode)
               '      rs1.Open StrSQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText
               ' If rs.RecordCount > 0 Then
               '     .TextMatrix(Row, .ColIndex("nationality")) = IIf(IsNull(rs("Nationality").value), "", rs("Nationality").value)
               '      .TextMatrix(Row, .ColIndex("iqamano")) = IIf(IsNull(rs("NumEkama").value), "", rs("NumEkama").value)
               '      .TextMatrix(Row, .ColIndex("job")) = IIf(IsNull(rs1("JobTypeName").value), "", rs1("JobTypeName").value)
               '
               ' Else
               '     .TextMatrix(Row, .ColIndex("value")) = ""
               ' End If
 

 
 '   Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    
                     
 '               If Rs3.RecordCount > 0 Then
 '                   .TextMatrix(Row, .ColIndex("JobName")) = IIf(IsNull(Rs3("JobTypeName").value), Null, Rs3("JobTypeName").value)
 '                    .TextMatrix(Row, .ColIndex("ProjectName")) = IIf(IsNull(Rs3("GroupName").value), "", Rs3("GroupName").value)
 '
 '               End If
     
     
                
           '  MsgBox StrAccountCode
                   End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub
Private Sub ReLineGrid()
    Dim I As Integer
    Dim IntCounter  As Integer
     
    IntCounter = 0

    With Fg

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("serial")) = IntCounter
       
            End If
 
        Next I
 
    End With
        IntCounter = 0

    With Fg2

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("serial")) = IntCounter
       
            End If
 
        Next I
 
    End With

End Sub
Private Sub Fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
' With fg

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
      '  Select Case .ColKey(Col)
      '
      '      Case "Code"
      '         Cancel = True
      '    Case "job"
      '        fg.ComboList = ""
      '          Cancel = True
      '          Case "nationality"
      '        fg.ComboList = ""
      '          Cancel = True
      '           Case "iqamano"
      '         fg.ComboList = ""
      '          Cancel = True
     '   End Select

'    End With

    
End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    With Fg

        Select Case .ColKey(Col)

           ' Case "name"
           '     StrSQL = "select * from TblEmployee"
           '     rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    StrComboList = fg.BuildComboList(rs, "Emp_Name", "Emp_Code")
''                Else
'                    StrComboList = fg.BuildComboList(rs, "Emp_Namee", "Emp_Code")
'                End If
'
'                If StrComboList <> "" Then
'                    StrComboList = "|" & StrComboList
'                End If
'                 .ComboList = StrComboList
 
        End Select

    End With

End Sub



 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 3
        FrmEmployeeSearch.show
  
    End If

End Sub

Private Sub DcboEmpName_Click(Area As Integer)
 End Sub

Private Sub FG2_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim StrAccountCode1 As String
Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
     Dim Rs1 As New ADODB.Recordset
    Dim StrSQL1 As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim StrComboList As String
            
    
    With Fg2

       ' Select Case .ColKey(Col)
 
       '     Case "name"
       '         '  .TextMatrix(Row, .ColIndex("userid")) = user_id
       '
       '         StrAccountCode = .ComboData
       '         LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Code"), False, True)
       '         .TextMatrix(Row, .ColIndex("Code")) = StrAccountCode
       '
'
'                    StrSQL = " SELECT * FROM  TblEmployee where Emp_Code=" & val(StrAccountCode)
'
             
'
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                   StrSQL1 = "Select JobTypeName From TblEmpJobsTypes where JobTypeID=" & val(StrAccountCode)
'                     rs1.Open StrSQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'                If rs.RecordCount > 0 Then
'                    .TextMatrix(Row, .ColIndex("nationality")) = IIf(IsNull(rs("Nationality").value), "", rs("Nationality").value)
'                     .TextMatrix(Row, .ColIndex("iqamano")) = IIf(IsNull(rs("NumEkama").value), "", rs("NumEkama").value)
'                     .TextMatrix(Row, .ColIndex("job")) = IIf(IsNull(rs1("JobTypeName").value), "", rs1("JobTypeName").value)
'
'                Else
'                    .TextMatrix(Row, .ColIndex("value")) = ""
'                End If
 

 

'                   End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub


Private Sub FG2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'With FG2
'
'        '   If Row > .FixedRows Then
'        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
'        '           Cancel = True
'        '       End If
'        '   End If
'        Select Case .ColKey(Col)
'
'            Case "Code"
'               Cancel = True
'          Case "job"
'              FG2.ComboList = ""
'                Cancel = True
'                Case "nationality"
'              FG2.ComboList = ""
'                Cancel = True
'                 Case "iqamano"
'               FG2.ComboList = ""
'                Cancel = True
'        End Select
'
'    End With

End Sub

Private Sub FG2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
' Dim rs As New ADODB.Recordset
'    Dim StrSQL  As String
'    Dim StrAccountType As String
'    Dim StrComboList As String
'    Dim Msg As String
'
'    With FG2
'
'        Select Case .ColKey(Col)
'
'            Case "name"
'                StrSQL = "select * from TblEmployee"
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

'                If SystemOptions.UserInterface = ArabicInterface Then
'                    StrComboList = FG2.BuildComboList(rs, "Emp_Name", "Emp_Code")
'                Else
'                    StrComboList = FG2.BuildComboList(rs, "Emp_Namee", "Emp_Code")
'                End If
'
'                If StrComboList <> "" Then
'                    StrComboList = "|" & StrComboList
'                End If
'                 .ComboList = StrComboList
 
'        End Select

'    End With
End Sub

Private Sub TxtAmount_Change()
Me.TxtWordAmount.Text = WriteNo(Me.TxtAmount.Text, 1)
End Sub

Private Sub TxtNameApplicant_Change()
If Me.TxtModFlg <> "R" Then
TxtNameOwner.Text = TxtNameApplicant.Text
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
    Dcombos.GetEmpDepartments Me.DcbDetpartment
   'Dcombos.GetEmployees Me.DcboEmpName
   Dcombos.GetBranches Me.dcBranch
 If SystemOptions.UserInterface = EnglishInterface Then
        ComMD.AddItem "Day"
        ComMD.AddItem "Month"
        ComMD.AddItem "Year"
        ComStopMD.AddItem "Day"
        ComStopMD.AddItem "Month"
        ComStopMD.AddItem "Year"
        Else
        ComMD.AddItem "ÌÊ„"
        ComMD.AddItem "‘Â—"
        ComMD.AddItem "”‰Â"
        ComStopMD.AddItem "ÌÊ„"
        ComStopMD.AddItem "‘Â—"
        ComStopMD.AddItem "”‰Â"
    End If
  
    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
   ' YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblCreditFacicity     Order By ID"
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
  '  Label1.Visible = False
Cmd(10).Caption = "Delete"
Cmd(21).Caption = "Delete"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Prient"
    Cmd(6).Caption = "Exit"
    Cmd(8).Caption = "Prient CC"
    CmdHelp.Caption = "Help"
    LblFg.Caption = "Authorized Person"
    lblfg2.Caption = "Authorized Person"
Cmd(9).Caption = "Prient"
    Me.Caption = "APPLICATION FOR CREDIT FACICITY"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lblBr.Caption = "Branch"
    lbl(2).Caption = "Name Applicant"
    lblowner.Caption = "Nmae Owner"
    lbl(10).Caption = "Address"
   lbstreet.Caption = "Street"
    lbl(3).Caption = " City"
    lbl(13).Caption = "Type of Business"
    lbl(5).Caption = " E-Mail"
    lblfax.Caption = "Fax"
    lbl(9).Caption = " Telephone"
    lbl(14).Caption = "C.R.NO"
    lblbox.Caption = "B.O.Box"
    lbl(17).Caption = "Acredit"
    lbl(16).Caption = "Zipe Code"
    lbllong.Caption = "Period"
    XPTab301.Caption = "Acredit Facicity"
    lbl(11).Caption = "C.C NO"
    lbl(21).Caption = "The value of Withdrawals"
    lbl(22).Caption = "Amount Words"
    lbl(23).Caption = "Samples NO"
    lbl(24).Caption = "Samples Amount"
    lbl(26).Caption = "Type1"
    lbl(28).Caption = "Type2"
    lbl(29).Caption = "Type3"
    lbl(31).Caption = "Type4"
    XPTab301.Caption = "Credit Facicity"
    lbl(15).Caption = "Bank Name"
    lbl(18).Caption = "Barnch"
    lbl(19).Caption = "Acc No"
    lbl(20).Caption = "Acc Officer"
    lblstopacc.Caption = "Is to stop the account to stop the clouds for"
    Me.lblBankinformation.Caption = "Bank Information"
    Me.lblinformationsh.Caption = "Samples Information"
    Me.lbltypesshow.Caption = "Samples Type of Show"
   ' lbl(3).Caption = "Employee"
   ' lbl(2).Caption = "value"
   ' lbl(0).Caption = "Box"
   ' Fra(0).Caption = "payments Method"
   ' lbl(9).Caption = "Count"
   ' lbl(10).Caption = "Start"
   ' lbl(11).Caption = "Month"
   ' lbl(12).Caption = "Year"
   ' Cmd(8).Caption = "Calc Dates"
   ' ChkSaleryDis.Caption = "Auto Discount"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"

    With Me.Fg
        .TextMatrix(0, .ColIndex("name")) = "Name"
        .TextMatrix(0, .ColIndex("serial")) = "Serial"
        .TextMatrix(0, .ColIndex("job")) = "Title"
         .TextMatrix(0, .ColIndex("Code")) = "Code"
    .TextMatrix(0, .ColIndex("nationality")) = "Nationality"
.TextMatrix(0, .ColIndex("iqamano")) = "IqamaNo"
    End With
    With Me.Fg2
        .TextMatrix(0, .ColIndex("name")) = "Name"
        .TextMatrix(0, .ColIndex("serial")) = "Serial"
        .TextMatrix(0, .ColIndex("job")) = "Title"
         .TextMatrix(0, .ColIndex("Code")) = "Code"
          .TextMatrix(0, .ColIndex("nationality")) = "Nationality"
.TextMatrix(0, .ColIndex("iqamano")) = "IqamaNo"
    End With

End Sub

'Private Sub YearMonth()

  '  Dim i As Integer
  '  Dim IntDefIndex As Integer
'
'    CmbMonth.Clear
'
'    For i = 1 To 12
'        CmbMonth.AddItem MonthName(i)
'    Next
'
'    CmbMonth.ListIndex = Month(Date) - 1
'    CboYear.Clear
'
 '   For i = 2010 To 2050
''        CboYear.AddItem i
'
'        If i = year(Date) Then
'            IntDefIndex = CboYear.NewIndex
'        End If

'    Next

'    CboYear.ListIndex = IntDefIndex
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
           ' TxtAdvanceValue.Locked = False
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
           ' TxtAdvanceValue.Locked = False
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
    Dim I As Integer
    Dim StrSQL As String
Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
            Fg2.Clear flexClearScrollable, flexClearEverything
            Fg2.Rows = 2
            Fg2.Enabled = True
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
Me.DcbDetpartment.BoundText = IIf(IsNull(rs("DeptID").value), "", rs("DeptID").value)
   XPTxtID.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
   XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
   dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
  ' Me.DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
   Me.TxtNameApplicant.Text = IIf(IsNull(rs("NameApplicant").value), "", rs("NameApplicant").value)
   Me.TxtNameOwner.Text = IIf(IsNull(rs("NameOwner").value), "", rs("NameOwner").value)
   Me.TxtStreet.Text = IIf(IsNull(rs("Street").value), "", rs("Street").value)
    Me.TxtCity.Text = IIf(IsNull(rs("City").value), "", rs("City").value)
    Me.TxtEmail.Text = IIf(IsNull(rs("Email").value), "", rs("Email").value)
    Me.TxtFax.Text = IIf(IsNull(rs("Fax").value), "", rs("Fax").value)
    Me.TxtPhone.Text = IIf(IsNull(rs("Phone").value), "", rs("Phone").value)
    Me.TxtCRNo.Text = IIf(IsNull(rs("CRNo").value), "", rs("CRNo").value)
    Me.TxtCRSource.Text = IIf(IsNull(rs("CRSource").value), "", rs("CRSource").value)
    Me.TxtPOBox.Text = IIf(IsNull(rs("POBox").value), "", rs("POBox").value)
    Me.TxtZipCOd.Text = IIf(IsNull(rs("ZipCode").value), "", rs("ZipCode").value)
    Me.TxtAddress.Text = IIf(IsNull(rs("Address").value), "", rs("Address").value)
    Me.TxtTypeBusnis.Text = IIf(IsNull(rs("TypeBusines").value), "", rs("TypeBusines").value)
    Me.TxtLong.Text = IIf(IsNull(rs("longT").value), "", rs("longT").value)
   Me.TxtAcredit.Text = val(IIf(IsNull(rs("Acredit").value), 0, rs("Acredit").value))
   Me.ComMD.Text = IIf(IsNull(rs("DMY").value), "", rs("DMY").value)
   Me.TxtCCNo.Text = IIf(IsNull(rs("CCNO").value), "", rs("CCNO").value)
   Me.TxtAmount.Text = val(IIf(IsNull(rs("Amount").value), 0, rs("Amount").value))
   Me.TxtWordAmount.Text = IIf(IsNull(rs("WordAmount").value), "", rs("WordAmount").value)
   Me.TxtShowNO.Text = IIf(IsNull(rs("ShowNo").value), "", rs("ShowNo").value)
   Me.TxtShowTy1.Text = IIf(IsNull(rs("Showtype1").value), "", rs("Showtype1").value)
   Me.TxtShowTy2.Text = IIf(IsNull(rs("Showtype2").value), "", rs("Showtype2").value)
   Me.TxtShowTy3.Text = IIf(IsNull(rs("Showtype3").value), "", rs("Showtype3").value)
   Me.TxtShowTy4.Text = IIf(IsNull(rs("Showtype4").value), "", rs("Showtype4").value)
   Me.TxtStopAcc.Text = IIf(IsNull(rs("StopAccount").value), "", rs("StopAccount").value)
   Me.ComStopMD.Text = IIf(IsNull(rs("StopDMY").value), "", rs("StopDMY").value)
   Me.txtBankName.Text = IIf(IsNull(rs("BanckName").value), "", rs("BanckName").value)
   Me.TxtBankBranch.Text = IIf(IsNull(rs("BanckBranch").value), "", rs("BanckBranch").value)
   Me.TxtAccNo.Text = IIf(IsNull(rs("AccNo").value), "", rs("AccNo").value)
   Me.TxtAccOficer.Text = IIf(IsNull(rs("AccOficer").value), "", rs("AccOficer").value)
   Me.TxtShowAmount.Text = val(IIf(IsNull(rs("ShowAmount").value), 0, rs("ShowAmount").value))
   ' TxtFromName.text = IIf(IsNull(rs("FromName").value), "", rs("FromName").value)
   ' TxtPersonalDept.text = IIf(IsNull(rs("PersonalDept").value), "", rs("PersonalDept").value)
      '  DcboEmpDepartments.BoundText = IIf(IsNull(rs("DeparmentID").value), "", rs("DeparmentID").value)

    'DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)

    'DcboJobsType.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)

   'lbl(23).Caption = IIf(IsNull(rs("basicSalary").value), "", rs("basicSalary").value)
 
   ' lbl(22).Caption = IIf(IsNull(rs("EmpDue").value), "", rs("EmpDue").value)
   'lbl(20).Caption = IIf(IsNull(rs("Contractvalid").value), "", rs("Contractvalid").value)
   'lbl(21).Caption = IIf(IsNull(rs("oldAdvance").value), "", rs("oldAdvance").value)
 
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


fillapprovData

   
   
    Set RsDetails = New ADODB.Recordset
    StrSQL = "Select * From  TblCreditFacicityDetails1 Where (ID=" & val(XPTxtID.Text) & ") And (Type = 0)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.Rows = Fg.FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        Fg.Rows = Fg.FixedRows + RsDetails.RecordCount

        For I = Me.Fg.FixedRows To Fg.Rows - 1
          '  fg.TextMatrix(i, fg.ColIndex("Code")) = RsDetails("Code").value
            Fg.TextMatrix(I, Fg.ColIndex("name")) = RsDetails("name").value
             Fg.TextMatrix(I, Fg.ColIndex("job")) = RsDetails("job").value

               Fg.TextMatrix(I, Fg.ColIndex("nationality")) = RsDetails("nationality").value
               Fg.TextMatrix(I, Fg.ColIndex("id")) = RsDetails("ID").value
               Fg.TextMatrix(I, Fg.ColIndex("iqamano")) = RsDetails("iqamano").value
            RsDetails.MoveNext
        Next I

    End If

    RsDetails.Close
    Set RsDetails = Nothing
    
    
        Set RsDetails1 = New ADODB.Recordset
  StrSQL = "SELECT *From dbo.TblCreditFacicityDetails1 WHERE     (ID=" & val(XPTxtID.Text) & ") And (Type = 1)"
    RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Fg2.Clear flexClearScrollable, flexClearEverything
    Fg2.Rows = Fg2.FixedRows

    If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
        RsDetails1.MoveFirst
        Fg2.Rows = Fg2.FixedRows + RsDetails1.RecordCount

        For I = Me.Fg2.FixedRows To Fg2.Rows - 1
          '  FG2.TextMatrix(i, FG2.ColIndex("Code")) = RsDetails1("Code").value
            Fg2.TextMatrix(I, Fg2.ColIndex("name")) = RsDetails1("name").value
             Fg2.TextMatrix(I, Fg2.ColIndex("job")) = RsDetails1("job").value
                       Fg2.TextMatrix(I, Fg2.ColIndex("nationality")) = RsDetails1("nationality").value
               Fg2.TextMatrix(I, Fg2.ColIndex("id")) = RsDetails1("ID").value
               Fg2.TextMatrix(I, Fg2.ColIndex("iqamano")) = RsDetails1("iqamano").value
            RsDetails1.MoveNext
        Next I

    End If

    RsDetails1.Close
    Set RsDetails1 = Nothing
    
    
    ReLineGrid
    
    
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
    Dim I As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap

 
   If Me.TxtNameApplicant.Text = "" Then
            Msg = "ÌÃ» «œŒ«· «”„ „ﬁœ„ «·ÿ·»..!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtNameApplicant.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
  If Me.TxtNameOwner.Text = "" Then
            Msg = "ÌÃ» «œŒ«· «”„  «·„«·ﬂ..!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtNameOwner.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

    
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then

            XPTxtID.Text = CStr(new_id("TblCreditFacicity", "ID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.Text = "E" Then
            StrSQL = "Delete From TblCreditFacicityDetails1 Where ID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If
        rs("ID").value = val(XPTxtID.Text)
        rs("DeptID").value = IIf(Me.DcbDetpartment.BoundText = "", Null, Me.DcbDetpartment.BoundText)
        rs("RecordDate").value = XPDtbTrans.value
        rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
    '    rs("EmpID").value = val(IIf(Me.DcboEmpName.BoundText = "", Null, Me.DcboEmpName.BoundText))
        rs("NameApplicant").value = IIf(TxtNameApplicant.Text = "", Null, TxtNameApplicant.Text)
        rs("NameOwner").value = IIf(TxtNameOwner.Text = "", Null, TxtNameOwner.Text)
       rs("Street").value = IIf(TxtStreet.Text = "", Null, TxtStreet.Text)
        rs("City").value = IIf(TxtCity.Text = "", Null, TxtCity.Text)
        rs("Email").value = IIf(TxtEmail.Text = "", Null, TxtEmail.Text)
        rs("Fax").value = IIf(TxtFax.Text = "", Null, TxtFax.Text)
      rs("Phone").value = IIf(TxtPhone.Text = "", Null, TxtPhone.Text)
      rs("CRNo").value = IIf(TxtCRNo.Text = "", Null, TxtCRNo.Text)
      rs("CRSource").value = IIf(TxtCRSource.Text = "", Null, TxtCRSource.Text)
      rs("POBox").value = IIf(TxtPOBox.Text = "", Null, TxtPOBox.Text)
      rs("ZipCode").value = IIf(TxtZipCOd.Text = "", Null, TxtZipCOd.Text)
      rs("Address").value = IIf(TxtAddress.Text = "", Null, TxtAddress.Text)
      rs("TypeBusines").value = IIf(TxtTypeBusnis.Text = "", Null, TxtTypeBusnis.Text)
      rs("longT").value = IIf(TxtLong.Text = "", Null, TxtLong.Text)
      rs("Acredit").value = val(IIf(TxtAcredit.Text = "", 0, TxtAcredit.Text))
      rs("DMY").value = IIf(ComMD.Text = "", Null, ComMD.Text)
      rs("CCNO").value = IIf(TxtCCNo.Text = "", Null, TxtCCNo.Text)
     rs("Amount").value = val(IIf(Me.TxtAmount.Text = "", 0, TxtAmount.Text))
      rs("WordAmount").value = IIf(Me.TxtWordAmount.Text = "", Null, TxtWordAmount.Text)
        rs("ShowNo").value = IIf(Me.TxtShowNO.Text = "", Null, TxtShowNO.Text)
      rs("Showtype1").value = IIf(Me.TxtShowTy1.Text = "", Null, TxtShowTy1.Text)
         rs("Showtype2").value = IIf(TxtShowTy2.Text = "", Null, TxtShowTy2.Text)
      rs("Showtype3").value = IIf(TxtShowTy3.Text = "", Null, TxtShowTy3.Text)
        rs("Showtype4").value = IIf(TxtShowTy4.Text = "", Null, TxtShowTy4.Text)
      rs("StopAccount").value = IIf(Me.TxtStopAcc.Text = "", Null, TxtStopAcc.Text)
         rs("StopDMY").value = IIf(Me.ComStopMD.Text = "", Null, ComStopMD.Text)
      rs("BanckName").value = IIf(Me.txtBankName.Text = "", Null, txtBankName.Text)
         rs("BanckBranch").value = IIf(Me.TxtBankBranch.Text = "", Null, TxtBankBranch.Text)
      rs("AccNo").value = IIf(Me.TxtAccNo.Text = "", Null, TxtAccNo.Text)
       rs("AccOficer").value = IIf(Me.TxtAccOficer.Text = "", Null, TxtAccOficer.Text)
      rs("ShowAmount").value = val(IIf(Me.TxtShowAmount.Text = "", 0, TxtShowAmount.Text))
 
        rs("UserID").value = Me.DCboUserName.BoundText
  
        rs.update
        
        Set RsDetails = New ADODB.Recordset
        RsDetails.Open "TblCreditFacicityDetails1", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

       For I = Me.Fg.FixedRows To Fg.Rows - 1
       If Fg.TextMatrix(I, Fg.ColIndex("name")) <> "" Then
            RsDetails.AddNew
           RsDetails("ID").value = val(XPTxtID.Text)
            RsDetails("type").value = 0
            RsDetails("name").value = Fg.TextMatrix(I, Fg.ColIndex("name"))
            RsDetails("job").value = Fg.TextMatrix(I, Fg.ColIndex("job"))
     ' RsDetails("Code").value = fg.TextMatrix(i, fg.ColIndex("Code"))
            RsDetails("nationality").value = Fg.TextMatrix(I, Fg.ColIndex("nationality"))
            RsDetails("iqamano").value = Fg.TextMatrix(I, Fg.ColIndex("iqamano"))
            RsDetails.update
            End If
        Next I
        
        
        Set RsDetails1 = New ADODB.Recordset
        RsDetails1.Open "TblCreditFacicityDetails1", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

       For I = Me.Fg2.FixedRows To Fg2.Rows - 1
       If Fg2.TextMatrix(I, Fg2.ColIndex("name")) <> "" Then
            RsDetails1.AddNew
           RsDetails1("ID").value = val(XPTxtID.Text)
             RsDetails1("Type").value = 1
             RsDetails1("name").value = Fg2.TextMatrix(I, Fg2.ColIndex("name"))
            RsDetails1("job").value = Fg2.TextMatrix(I, Fg2.ColIndex("job"))
          '  RsDetails1("Code").value = FG2.TextMatrix(i, FG2.ColIndex("Code"))
            RsDetails1("nationality").value = Fg2.TextMatrix(I, Fg2.ColIndex("nationality"))
            RsDetails1("iqamano").value = Fg2.TextMatrix(I, Fg2.ColIndex("iqamano"))
            RsDetails1.update
            End If
        Next I
    
   
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
        RsDetails.Close
        Set RsDetails = Nothing
         RsDetails1.Close
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
'End If

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
Dim strsql2 As String
    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
        StrSQL1 = "Delete From TblCreditFacicityDetails1 Where ID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL1, , adExecuteNoRecords
        
                If rs.RecordCount < 1 Then
                 
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 2
            Fg.Enabled = True
               Fg2.Clear flexClearScrollable, flexClearEverything
            Fg2.Rows = 2
            Fg2.Enabled = True
            End If
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
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

Function fillapprovDataxx()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
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
        .Create Me.hWnd, " ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«» ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«»", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«»", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«»", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«»", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«»", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, " ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«»", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«»", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«»", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«»", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ”ÂÌ·«  ≈∆ „«‰Ì…/› Õ Õ”«»", 1, 15204351, -2147483630
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

Private Sub TxtAdvanceValue_KeyPress(KeyAscii As Integer)
    
End Sub

Private Function CheckDate() As Boolean
     
End Function

Private Function CheckPartCal() As Boolean
   
End Function

Private Sub CalCulateParts()
    

End Sub

Private Sub RemoveGridRow()

    With Me.Fg

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub RemoveGridRow1()

    With Me.Fg2

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub
