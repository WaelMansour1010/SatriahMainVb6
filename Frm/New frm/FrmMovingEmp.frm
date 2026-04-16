VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmMovingEmp 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ’—ÌÕ Œ—ÊÃ „ƒﬁ - ≈Ã«“… ⁄«—÷…    "
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13005
   Icon            =   "FrmMovingEmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   13005
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   199
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtid 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
      Height          =   1575
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   600
      Width           =   2175
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   255
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Œ—ÊÃ „ƒﬁ "
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   40
         Top             =   600
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "«” ∆–«‰"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "«Ã«“… ⁄«—÷…"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   173
         Top             =   1200
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "«Ã«“… »œÊ‰ —« »"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.TextBox TxtSearchCodeB 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10320
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   1530
      Width           =   1095
   End
   Begin VB.TextBox TxtSearchCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   1185
      Width           =   1095
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   13440
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
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
      Left            =   -120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   13125
      _cx             =   23151
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
      Caption         =   " ’—ÌÕ Œ—ÊÃ „ƒﬁ - ≈Ã«“… ⁄«—÷…    "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   2
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
         Left            =   1305
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
         ButtonImage     =   "FrmMovingEmp.frx":038A
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
         Left            =   240
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
         ButtonImage     =   "FrmMovingEmp.frx":0724
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
         Left            =   1830
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
         ButtonImage     =   "FrmMovingEmp.frx":0ABE
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
         Left            =   765
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
         ButtonImage     =   "FrmMovingEmp.frx":0E58
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
         Picture         =   "FrmMovingEmp.frx":11F2
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
         TabIndex        =   31
         Top             =   0
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   6960
      TabIndex        =   7
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   207093761
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DcboEmpName 
      Height          =   315
      Left            =   6960
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
      Left            =   2670
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6900
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
         Left            =   3825
         TabIndex        =   14
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
         Left            =   855
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
         Left            =   2760
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
         Left            =   1920
         TabIndex        =   34
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
      Left            =   8940
      TabIndex        =   17
      Top             =   6480
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
      Bindings        =   "FrmMovingEmp.frx":4E5A
      Height          =   315
      Left            =   2520
      TabIndex        =   29
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSDataListLib.DataCombo DcboBossName 
      Height          =   315
      Left            =   6960
      TabIndex        =   37
      Top             =   1530
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker ReDta 
      Height          =   315
      Left            =   0
      TabIndex        =   42
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   207093761
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   4095
      Left            =   0
      TabIndex        =   43
      Top             =   2370
      Width           =   12960
      _cx             =   22860
      _cy             =   7223
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
      Caption         =   "«·»Ì«‰« |Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "FrmMovingEmp.frx":4E6F
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   3630
         Left            =   13605
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   45
         Width           =   12870
         _cx             =   22701
         _cy             =   6403
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
            Height          =   3270
            Left            =   0
            TabIndex        =   176
            Tag             =   "1"
            Top             =   0
            Width           =   12870
            _cx             =   22701
            _cy             =   5768
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
            FormatString    =   $"FrmMovingEmp.frx":5209
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
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   177
            Top             =   3360
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3630
         Left            =   45
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   45
         Width           =   12870
         _cx             =   22701
         _cy             =   6403
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3510
            Index           =   0
            Left            =   0
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   0
            Width           =   12870
            _cx             =   22701
            _cy             =   6191
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
            Begin VB.TextBox TxtSalary 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   2280
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   197
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox TxtNoVaction 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   155
               Top             =   0
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox TxtAbceDay 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   9960
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   154
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox TxtMaxDay 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   153
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox TxtNoDay 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   4080
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   152
               Top             =   840
               Width           =   975
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «‰ﬁ·"
               Height          =   1545
               Left            =   13815
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   0
               Width           =   6105
               Begin MSDataListLib.DataCombo DataCombo3 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   138
                  Top             =   240
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DTPicker2 
                  Height          =   315
                  Left            =   6480
                  TabIndex        =   139
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   207093761
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DataCombo4 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   140
                  Top             =   600
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DataCombo5 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   141
                  Top             =   240
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DTPicker3 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   142
                  Top             =   600
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   207093761
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DataCombo6 
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   143
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
                  Caption         =   "„‰ ﬁ”„"
                  Height          =   285
                  Index           =   42
                  Left            =   5280
                  TabIndex        =   151
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   43
                  Left            =   6240
                  TabIndex        =   150
                  Top             =   480
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï ﬁ”„"
                  Height          =   285
                  Index           =   44
                  Left            =   2280
                  TabIndex        =   149
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
                  Height          =   285
                  Index           =   45
                  Left            =   6360
                  TabIndex        =   148
                  Top             =   360
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—« » «·«”«”Ì"
                  Height          =   285
                  Index           =   46
                  Left            =   6600
                  TabIndex        =   147
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»ÊŸÌ›…"
                  Height          =   285
                  Index           =   47
                  Left            =   5160
                  TabIndex        =   146
                  Top             =   600
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰ ÌÊ„"
                  Height          =   285
                  Index           =   48
                  Left            =   2280
                  TabIndex        =   145
                  Top             =   600
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»ÊŸÌ›…"
                  Height          =   285
                  Index           =   49
                  Left            =   5160
                  TabIndex        =   144
                  Top             =   960
                  Width           =   645
               End
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  „«·Ì…"
               Height          =   1005
               Left            =   14760
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   0
               Width           =   6015
               Begin MSDataListLib.DataCombo DataCombo7 
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   128
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
                  Caption         =   "«·„— »…"
                  Height          =   285
                  Index           =   50
                  Left            =   4800
                  TabIndex        =   136
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   51
                  Left            =   3240
                  TabIndex        =   135
                  Top             =   720
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   52
                  Left            =   960
                  TabIndex        =   134
                  Top             =   360
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   53
                  Left            =   960
                  TabIndex        =   133
                  Top             =   720
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   285
                  Index           =   54
                  Left            =   -240
                  TabIndex        =   132
                  Top             =   720
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”·› ·„  ”œœ"
                  Height          =   285
                  Index           =   55
                  Left            =   1800
                  TabIndex        =   131
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·⁄ﬁœ «·„ »ﬁÌ…"
                  Height          =   285
                  Index           =   56
                  Left            =   1560
                  TabIndex        =   130
                  Top             =   720
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·„” Õﬁ«  ··„ÊŸ›"
                  Height          =   285
                  Index           =   57
                  Left            =   3960
                  TabIndex        =   129
                  Top             =   720
                  Width           =   1965
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
               Height          =   3765
               Index           =   1
               Left            =   14145
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   360
               Width           =   6135
               Begin VB.ComboBox Combo1 
                  Height          =   315
                  Left            =   4110
                  Style           =   2  'Dropdown List
                  TabIndex        =   119
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Œ’„ „‰ «·„— »  ·ﬁ«∆Ì«"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3960
                  TabIndex        =   118
                  Top             =   2160
                  Value           =   1  'Checked
                  Visible         =   0   'False
                  Width           =   1965
               End
               Begin VB.ComboBox Combo2 
                  Height          =   315
                  Left            =   4110
                  Style           =   2  'Dropdown List
                  TabIndex        =   117
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.TextBox Text1 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   4110
                  MaxLength       =   2
                  TabIndex        =   116
                  Top             =   240
                  Width           =   825
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   435
                  Index           =   10
                  Left            =   4080
                  TabIndex        =   120
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
                  ButtonImage     =   "FrmMovingEmp.frx":534C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
                  Height          =   2325
                  Left            =   90
                  TabIndex        =   121
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
                  FormatString    =   $"FrmMovingEmp.frx":56E6
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
                  Caption         =   "”‰…"
                  Height          =   315
                  Index           =   58
                  Left            =   5250
                  TabIndex        =   126
                  Top             =   1320
                  Width           =   405
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   315
                  Index           =   59
                  Left            =   5250
                  TabIndex        =   125
                  Top             =   990
                  Width           =   405
               End
               Begin VB.Label Label2 
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
                  TabIndex        =   124
                  Top             =   2280
                  Visible         =   0   'False
                  Width           =   2595
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «Ê· œ›⁄…"
                  Height          =   285
                  Index           =   60
                  Left            =   4380
                  TabIndex        =   123
                  Top             =   690
                  Width           =   1665
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·œ›⁄« "
                  Height          =   285
                  Index           =   61
                  Left            =   4830
                  TabIndex        =   122
                  Top             =   300
                  Width           =   975
               End
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   13410
               MaxLength       =   10
               TabIndex        =   114
               Top             =   2100
               Width           =   1425
            End
            Begin VB.TextBox TxtRemark2 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   113
               Top             =   840
               Width           =   3135
            End
            Begin VB.TextBox TxtBalanceDay 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   4080
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   112
               Top             =   375
               Width           =   975
            End
            Begin MSComCtl2.DTPicker ToDate 
               Height          =   375
               Left            =   7260
               TabIndex        =   156
               Top             =   840
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _Version        =   393216
               Format          =   204406785
               CurrentDate     =   38784
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   270
               Left            =   240
               TabIndex        =   157
               Top             =   3675
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   476
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
            Begin MSComCtl2.DTPicker FromDate 
               Height          =   375
               Left            =   9480
               TabIndex        =   158
               Top             =   840
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _Version        =   393216
               Format          =   204406785
               CurrentDate     =   38784
            End
            Begin XtremeSuiteControls.RadioButton TypeDisc 
               Height          =   285
               Index           =   0
               Left            =   8160
               TabIndex        =   159
               Top             =   1320
               Width           =   1695
               _Version        =   786432
               _ExtentX        =   2990
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Œ’„ „‰ «·—« »"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton TypeDisc 
               Height          =   285
               Index           =   1
               Left            =   6120
               TabIndex        =   160
               Top             =   1335
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Œ’„ „‰ —’Ìœ «·«Ã«“« "
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton TypeDisc 
               Height          =   285
               Index           =   2
               Left            =   5280
               TabIndex        =   161
               Top             =   1335
               Visible         =   0   'False
               Width           =   855
               _Version        =   786432
               _ExtentX        =   1508
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "·«ÌÊÃœ"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbVacation 
               Height          =   315
               Left            =   7260
               TabIndex        =   162
               Top             =   375
               Width           =   3675
               _ExtentX        =   6482
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTypeVaction 
               Height          =   285
               Index           =   0
               Left            =   8880
               TabIndex        =   178
               Top             =   360
               Width           =   3495
               _Version        =   786432
               _ExtentX        =   6165
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Œ’„ „‰ —’Ìœ «·«Ã«“…"
               ForeColor       =   8388608
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTypeVaction 
               Height          =   285
               Index           =   1
               Left            =   4680
               TabIndex        =   179
               Top             =   360
               Width           =   3135
               _Version        =   786432
               _ExtentX        =   5530
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Œ’„ „‰ «Ì«„ «·⁄„·"
               ForeColor       =   8388608
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   555
               Index           =   11
               Left            =   120
               TabIndex        =   180
               TabStop         =   0   'False
               Top             =   2880
               Width           =   9615
               _cx             =   16960
               _cy             =   979
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
               Begin VB.CommandButton Command9 
                  Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   185
                  Top             =   135
                  Width           =   1710
               End
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   1950
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   184
                  Top             =   135
                  Width           =   2295
               End
               Begin VB.TextBox TxtNoteID 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   8895
                  RightToLeft     =   -1  'True
                  TabIndex        =   183
                  Top             =   -15
                  Visible         =   0   'False
                  Width           =   2070
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "≈‰‘«¡ ﬁÌœ «·«” Õﬁ«ﬁ"
                  Height          =   345
                  Left            =   7605
                  RightToLeft     =   -1  'True
                  TabIndex        =   182
                  Top             =   135
                  Width           =   1710
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Õ–› ﬁÌœ «·«” Õﬁ«ﬁ"
                  Height          =   345
                  Left            =   5325
                  RightToLeft     =   -1  'True
                  TabIndex        =   181
                  Top             =   135
                  Width           =   1710
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   390
                  Index           =   35
                  Left            =   3855
                  RightToLeft     =   -1  'True
                  TabIndex        =   186
                  Top             =   135
                  Width           =   1125
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   1035
               Left            =   0
               TabIndex        =   187
               TabStop         =   0   'False
               Top             =   1800
               Width           =   12840
               _cx             =   22648
               _cy             =   1826
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
               BackColor       =   12648447
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
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   855
                  Left            =   0
                  TabIndex        =   188
                  Top             =   30
                  Width           =   12810
                  _cx             =   22595
                  _cy             =   1508
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
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   65
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmMovingEmp.frx":5771
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   690
                  Index           =   3
                  Left            =   7320
                  TabIndex        =   189
                  TabStop         =   0   'False
                  Top             =   1155
                  Width           =   1605
                  _cx             =   2831
                  _cy             =   1217
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
                  Caption         =   "≈Œ Ì«— «· «—ÌŒ"
                  Align           =   0
                  AutoSizeChildren=   7
                  BorderWidth     =   6
                  ChildSpacing    =   4
                  Splitter        =   0   'False
                  FloodDirection  =   0
                  FloodPercent    =   0
                  CaptionPos      =   6
                  WordWrap        =   -1  'True
                  MaxChildSize    =   0
                  MinChildSize    =   0
                  TagWidth        =   0
                  TagPosition     =   0
                  Style           =   1
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
                  Begin VB.ComboBox Combo4 
                     Height          =   315
                     Left            =   90
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   191
                     Top             =   420
                     Width           =   1755
                  End
                  Begin VB.ComboBox Combo3 
                     Height          =   315
                     Left            =   90
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   190
                     Top             =   180
                     Width           =   1755
                  End
                  Begin ImpulseButton.ISButton CmdOk 
                     Height          =   240
                     Left            =   90
                     TabIndex        =   192
                     Top             =   660
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   423
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "⁄—÷  "
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
                     ButtonImage     =   "FrmMovingEmp.frx":5F68
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Â—"
                     Height          =   15
                     Index           =   73
                     Left            =   90
                     RightToLeft     =   -1  'True
                     TabIndex        =   194
                     Top             =   1425
                     Width           =   1755
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”‰…"
                     Height          =   15
                     Index           =   30
                     Left            =   90
                     RightToLeft     =   -1  'True
                     TabIndex        =   193
                     Top             =   1395
                     Width           =   1755
                  End
               End
               Begin MSDataListLib.DataCombo Dcemp 
                  Height          =   315
                  Left            =   900
                  TabIndex        =   195
                  Top             =   1155
                  Width           =   2295
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„ÊŸ› „Õœœ"
                  DataField       =   "Õœœ"
                  Height          =   195
                  Index           =   74
                  Left            =   3045
                  RightToLeft     =   -1  'True
                  TabIndex        =   196
                  Top             =   1170
                  Width           =   1065
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—« »"
               Height          =   330
               Index           =   75
               Left            =   3000
               TabIndex        =   198
               Top             =   1320
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «Ì«„ «·€Ì«»"
               Height          =   330
               Index           =   72
               Left            =   10680
               TabIndex        =   172
               Top             =   1350
               Width           =   1920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ﬁ’Ï „œ…"
               Height          =   330
               Index           =   71
               Left            =   2400
               TabIndex        =   171
               Top             =   405
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «Ì«„ «·«Ã«“…"
               Height          =   330
               Index           =   70
               Left            =   5040
               TabIndex        =   170
               Top             =   885
               Width           =   1920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï"
               Height          =   210
               Index           =   69
               Left            =   8280
               TabIndex        =   169
               Top             =   885
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÌÊ„"
               ForeColor       =   &H00C00000&
               Height          =   330
               Index           =   67
               Left            =   3480
               TabIndex        =   168
               Top             =   405
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÊÌŒ’„ „‰ «·”·› „»·€« Êﬁœ—…"
               Height          =   330
               Index           =   64
               Left            =   12765
               TabIndex        =   167
               Top             =   1425
               Width           =   2280
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   330
               Index           =   65
               Left            =   3000
               TabIndex        =   166
               Top             =   945
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—’Ìœ «·«Ã«“«  «·⁄«—÷…"
               Height          =   330
               Index           =   66
               Left            =   5040
               TabIndex        =   165
               Top             =   405
               Width           =   1920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã«“… „‰"
               Height          =   330
               Index           =   68
               Left            =   10680
               TabIndex        =   164
               Top             =   885
               Width           =   1920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã«“… "
               Height          =   330
               Index           =   63
               Left            =   10680
               TabIndex        =   163
               Top             =   360
               Width           =   1920
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3630
            Index           =   16
            Left            =   0
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   0
            Width           =   12870
            _cx             =   22701
            _cy             =   6403
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
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «‰ﬁ·"
               Height          =   1545
               Left            =   13815
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   0
               Width           =   6105
               Begin MSDataListLib.DataCombo DcboEmpDepartments 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   76
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
                  TabIndex        =   77
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   204472321
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcboJobsType 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   78
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
                  TabIndex        =   79
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
                  TabIndex        =   80
                  Top             =   600
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   204472321
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DataCombo2 
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   81
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
                  Caption         =   "„‰ ﬁ”„"
                  Height          =   285
                  Index           =   24
                  Left            =   5280
                  TabIndex        =   89
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
                  TabIndex        =   88
                  Top             =   480
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï ﬁ”„"
                  Height          =   285
                  Index           =   15
                  Left            =   2280
                  TabIndex        =   87
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
                  Height          =   285
                  Index           =   13
                  Left            =   6360
                  TabIndex        =   86
                  Top             =   360
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—« » «·«”«”Ì"
                  Height          =   285
                  Index           =   5
                  Left            =   6600
                  TabIndex        =   85
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»ÊŸÌ›…"
                  Height          =   285
                  Index           =   36
                  Left            =   5160
                  TabIndex        =   84
                  Top             =   600
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰ ÌÊ„"
                  Height          =   285
                  Index           =   37
                  Left            =   2280
                  TabIndex        =   83
                  Top             =   600
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»ÊŸÌ›…"
                  Height          =   285
                  Index           =   38
                  Left            =   5160
                  TabIndex        =   82
                  Top             =   960
                  Width           =   645
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  „«·Ì…"
               Height          =   1005
               Left            =   14760
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   0
               Width           =   6015
               Begin MSDataListLib.DataCombo DcboSpecifications 
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   66
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
                  Caption         =   "«·„— »…"
                  Height          =   285
                  Index           =   14
                  Left            =   4800
                  TabIndex        =   74
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   22
                  Left            =   3240
                  TabIndex        =   73
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
                  TabIndex        =   72
                  Top             =   360
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   20
                  Left            =   960
                  TabIndex        =   71
                  Top             =   720
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   285
                  Index           =   16
                  Left            =   -240
                  TabIndex        =   70
                  Top             =   720
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”·› ·„  ”œœ"
                  Height          =   285
                  Index           =   19
                  Left            =   1800
                  TabIndex        =   69
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·⁄ﬁœ «·„ »ﬁÌ…"
                  Height          =   285
                  Index           =   18
                  Left            =   1560
                  TabIndex        =   68
                  Top             =   720
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·„” Õﬁ«  ··„ÊŸ›"
                  Height          =   285
                  Index           =   17
                  Left            =   3960
                  TabIndex        =   67
                  Top             =   720
                  Width           =   1965
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
               Height          =   3765
               Index           =   0
               Left            =   14145
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   360
               Width           =   6135
               Begin VB.ComboBox CboYear 
                  Height          =   315
                  Left            =   4110
                  Style           =   2  'Dropdown List
                  TabIndex        =   57
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.CheckBox ChkSaleryDis 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Œ’„ „‰ «·„— »  ·ﬁ«∆Ì«"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3960
                  TabIndex        =   56
                  Top             =   2160
                  Value           =   1  'Checked
                  Visible         =   0   'False
                  Width           =   1965
               End
               Begin VB.ComboBox CmbMonth 
                  Height          =   315
                  Left            =   4110
                  Style           =   2  'Dropdown List
                  TabIndex        =   55
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.TextBox TxtPaymentCounts 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   4110
                  MaxLength       =   2
                  TabIndex        =   54
                  Top             =   240
                  Width           =   825
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   435
                  Index           =   8
                  Left            =   4080
                  TabIndex        =   58
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
                  ButtonImage     =   "FrmMovingEmp.frx":6302
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   2325
                  Left            =   90
                  TabIndex        =   59
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
                  FormatString    =   $"FrmMovingEmp.frx":669C
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
                  Caption         =   "”‰…"
                  Height          =   315
                  Index           =   12
                  Left            =   5250
                  TabIndex        =   64
                  Top             =   1320
                  Width           =   405
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   315
                  Index           =   11
                  Left            =   5250
                  TabIndex        =   63
                  Top             =   990
                  Width           =   405
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
                  Index           =   0
                  Left            =   60
                  TabIndex        =   62
                  Top             =   2280
                  Visible         =   0   'False
                  Width           =   2595
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «Ê· œ›⁄…"
                  Height          =   285
                  Index           =   10
                  Left            =   4380
                  TabIndex        =   61
                  Top             =   690
                  Width           =   1665
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·œ›⁄« "
                  Height          =   285
                  Index           =   9
                  Left            =   4830
                  TabIndex        =   60
                  Top             =   300
                  Width           =   975
               End
            End
            Begin VB.TextBox TxtDiscount 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   13410
               MaxLength       =   10
               TabIndex        =   52
               Top             =   2100
               Width           =   1425
            End
            Begin VB.TextBox txtRemark 
               Alignment       =   1  'Right Justify
               Height          =   960
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   51
               Top             =   1305
               Width           =   10635
            End
            Begin VB.TextBox bossNotes 
               Alignment       =   1  'Right Justify
               Height          =   720
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   50
               Top             =   2400
               Width           =   10635
            End
            Begin VB.TextBox TxtInterval 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   10
               TabIndex        =   49
               Top             =   735
               Width           =   1395
            End
            Begin MSComCtl2.DTPicker TxtExpectedouttime 
               Height          =   315
               Left            =   9300
               TabIndex        =   90
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   204406787
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSComCtl2.DTPicker txtExpectedIntime 
               Height          =   375
               Left            =   9300
               TabIndex        =   91
               Top             =   735
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   204406787
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSComCtl2.DTPicker txtActualouttime 
               Height          =   315
               Left            =   5880
               TabIndex        =   92
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   204406787
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSComCtl2.DTPicker txtActualIntime 
               Height          =   375
               Left            =   5880
               TabIndex        =   93
               Top             =   735
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   204406787
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSDataListLib.DataCombo DcOutType 
               Height          =   315
               Left            =   960
               TabIndex        =   94
               Top             =   240
               Width           =   2715
               _ExtentX        =   4789
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2520
               Index           =   62
               Left            =   2370
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   1155
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÊÌŒ’„ „‰ «·”·› „»·€« Êﬁœ—…"
               Height          =   330
               Index           =   26
               Left            =   12765
               TabIndex        =   104
               Top             =   1425
               Width           =   2280
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   330
               Index           =   28
               Left            =   10800
               TabIndex        =   103
               Top             =   1665
               Width           =   1920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·Œ—ÊÃ «·„ Êﬁ⁄"
               Height          =   330
               Index           =   31
               Left            =   10800
               TabIndex        =   102
               Top             =   285
               Width           =   1920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·⁄Êœ… «·„ Êﬁ⁄"
               Height          =   330
               Index           =   32
               Left            =   10800
               TabIndex        =   101
               Top             =   705
               Width           =   1920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·Œ—ÊÃ «·›⁄·Ì"
               Height          =   210
               Index           =   34
               Left            =   7560
               TabIndex        =   100
               Top             =   240
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·⁄Êœ… «·›⁄·Ì"
               Height          =   255
               Index           =   35
               Left            =   7560
               TabIndex        =   99
               Top             =   720
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„—∆Ì«  «·—∆Ì” «·„»«‘— "
               Height          =   330
               Index           =   40
               Left            =   10800
               TabIndex        =   98
               Top             =   2640
               Width           =   1920
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·«–‰"
               Height          =   285
               Index           =   33
               Left            =   3480
               TabIndex        =   97
               Top             =   255
               Width           =   1365
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”«⁄Â"
               Height          =   285
               Index           =   29
               Left            =   840
               TabIndex        =   96
               Top             =   720
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„œ…"
               Height          =   285
               Index           =   2
               Left            =   3750
               TabIndex        =   95
               Top             =   735
               Width           =   1005
            End
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   8880
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label Label111000 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   46
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
            TabIndex        =   45
            Top             =   4560
            Width           =   3375
         End
      End
   End
   Begin MSDataListLib.DataCombo DcbMang 
      Height          =   315
      Left            =   2520
      TabIndex        =   107
      Top             =   1200
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker TempDate 
      Height          =   315
      Left            =   4440
      TabIndex        =   110
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   213385217
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Accredit 
      Height          =   390
      Left            =   960
      TabIndex        =   174
      Top             =   6960
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   688
      ButtonPositionImage=   1
      Caption         =   "«—”«· ··«⁄ „«œ"
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorButton     =   -2147483635
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin MSDataListLib.DataCombo DCGroupID 
      Height          =   315
      Left            =   6930
      TabIndex        =   200
      Top             =   1920
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Êﬁ⁄"
      Height          =   285
      Index           =   46
      Left            =   11700
      TabIndex        =   201
      Top             =   1950
      Width           =   915
   End
   Begin VB.Label LBLWhereSTR 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   109
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·«œ«—…"
      Height          =   255
      Index           =   25
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   106
      Top             =   1215
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·—∆Ì” «·„»«‘—"
      Height          =   285
      Index           =   41
      Left            =   11520
      TabIndex        =   35
      Top             =   1530
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Index           =   39
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   735
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «· ’—ÌÕ"
      Height          =   285
      Index           =   4
      Left            =   11880
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
      Left            =   11880
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
      Left            =   8310
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
      Left            =   11685
      TabIndex        =   24
      Top             =   6435
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2790
      TabIndex        =   23
      Top             =   6630
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   1050
      TabIndex        =   22
      Top             =   6630
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   450
      TabIndex        =   21
      Top             =   6660
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   2100
      TabIndex        =   20
      Top             =   6660
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
Attribute VB_Name = "FrmMovingEmp"
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
Dim BossId As String
Dim FixedOrChanged(40) As Integer
Dim AddOrDiscount(40) As Integer
Dim ViewComp(40) As Boolean
Dim showMofradAll(40) As Boolean
Dim culc30orRminder(40) As Integer
Dim Account_code(40) As String
Dim Account_code1(40) As String
Dim ZmamAccount(40) As String
Dim AdvPaymentdAccount(40) As String
Dim componentname(40) As String
Function GetHobStatus() As Integer
Dim sql As String
GetHobStatus = 0
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     id, Vacation"
sql = sql & " From dbo.jopstatus"
sql = sql & " Where (Vacation = 1)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetHobStatus = IIf(IsNull(Rs3("id").value), 0, Rs3("id").value)
Else
GetHobStatus = 0
End If
End Function
Function CheckEmbractin(Optional ID As Double) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from  TblEmbarkation where NoVationUnPaed=" & ID & " and TypeVacation=1 "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckEmbractin = True
Else
CheckEmbractin = False
End If
End Function
Private Sub Accredit_Click()
    Dim BeginTrans As Boolean

 If val(XPTxtID.text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
      
      
    Cn.BeginTrans
    BeginTrans = True

 
    SendTopost Me.Name, "TblEmpPassOver", "AdvanceID", val(DcbMang.BoundText), val(dcBranch.BoundText), val(XPTxtID.text), XPTxtID
    
   rs.Resync
   
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To Approval "
End If

    Cn.CommitTrans
    BeginTrans = False
 
    Retrive (val(Me.XPTxtID.text))
End Sub



Private Sub Cmd_Click(index As Integer)

    ' On Error GoTo ErrTrap
    Select Case index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
             TypeDisc(1).value = True
            Me.DCboUserName.BoundText = user_id
            TxtPaymentCounts.text = 1
dcBranch.BoundText = Current_branch
    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable
   End With
            'XPDtbTrans.SetFocus
                  Grid2.Clear flexClearScrollable, flexClearEverything
            Grid2.rows = 1
            Accredit.Caption = ""
            Accredit.Enabled = False
               ' If SystemOptions.UserInterface = ArabicInterface Then
                  '                                  Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                 '                                 Else
                         '                           Accredit.Caption = " send to Approval   "
                        '                       End If
                                               
             
    Grid2.Clear flexClearScrollable, flexClearEverything
    Grid2.rows = 1
    
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
        If TxtNoteSerial.text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· .Ì—ÃÏ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ"
        Else
        MsgBox "Can Not edit .Delete JE"
        End If
        Exit Sub
        End If
If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not edit.This process associated with approvals"
         End If
         Exit Sub
       End If
   If Rd(3).value = True Then
   If CheckEmbractin(val(XPTxtID.text)) = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·Õ—ﬂ… „— »ÿ… »«·„»«‘—«  "
   Else
   MsgBox "Can not edit employee work directly"
   End If
   Exit Sub
   End If
   End If
     
            TxtModFlg.text = "E"
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
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText
 If Rd(3).value = True Then
If GetHobStatus() = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  ÕœÌœ «·«Ã«“… „‰ ‘«‘… Õ«·«  «·⁄„·"
Else
MsgBox "Please Select Vacaion From Screen Job Situations"
End If
Exit Sub
End If
End If
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
    If TxtNoteSerial.text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «·Õ–› .Ì—ÃÏ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ"
        Else
        MsgBox "Can Not delete .Delete JE"
        End If
        Exit Sub
        End If
        
      If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «·Õ–›.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not delete.This process associated with approvals"
         End If
         Exit Sub
       End If
       
     If Rd(3).value = True Then
   If CheckEmbractin(val(XPTxtID.text)) = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… „— »ÿ… »«·„»«‘—«  "
   Else
   MsgBox "Can not delete employee work directly"
   End If
   Exit Sub
   End If
   End If
  
            Del_Trans

        Case 5
            FrmEmpAdvanceSearch.mIndex = 1
             Load FrmEmpAdvanceSearch
              FrmEmpAdvanceSearch.show

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200


            
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
On Error Resume Next
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  If Me.Rd(2).value = True Then
 MySQL = "  SELECT     dbo.TblEmpPassOver.AdvanceID, dbo.TblEmpPassOver.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
 MySQL = MySQL & "                      dbo.TblEmpPassOver.Emp_id, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
 MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmpPassOver.AdvanceDate, dbo.TblEmpPassOver.MangID,"
 MySQL = MySQL & "                     dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpPassOver.NoVaction, dbo.TblEmpPassOver.ToDate,"
 MySQL = MySQL & "                     dbo.TblEmpPassOver.FromDate, dbo.TblEmpPassOver.TypeDisc, dbo.TblEmpPassOver.Remark2, dbo.TblEmpPassOver.AbceDay, dbo.TblEmpPassOver.NoDay,"
 MySQL = MySQL & "                     dbo.TblEmpPassOver.MaxDay, dbo.TblEmpPassOver.BalanceDay, dbo.TblEmpPassOver.TypeTrans, dbo.TblEmpPassOver.VacationID, dbo.tblVacancy.Vac_Name,"
 MySQL = MySQL & "                     dbo.tblVacancy.NameE, dbo.TblEmpPassOver.BossId, TblEmployee_1.Emp_Name AS MangerEmp_Name, TblEmployee_1.Fullcode AS MangerFullcode,"
 MySQL = MySQL & "                     TblEmployee_1.Emp_Namee AS MangerEmp_NameE"
 MySQL = MySQL & " FROM         dbo.TblEmpPassOver LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee TblEmployee_1 ON dbo.TblEmpPassOver.BossId = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.tblVacancy ON dbo.TblEmpPassOver.VacationID = dbo.tblVacancy.Vac_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmpDepartments ON dbo.TblEmpPassOver.MangID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee ON dbo.TblEmpPassOver.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblEmpPassOver.Branch_NO = dbo.TblBranchesData.branch_id"
  MySQL = MySQL & " Where (dbo.TblEmpPassOver.TypeTrans = 2) And (dbo.TblEmpPassOver.advanceID = " & val(XPTxtID.text) & ")"
  Else
'MySQL = " SELECT     dbo.TblEmployee.Emp_Name, dbo.TblEmpPassOver.[interval], dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpPassOver.AdvanceDate, "
'MySQL = MySQL & "                       dbo.TblEmpPassOver.bossNotes , dbo.TblEmpPassOver.Remark, dbo.TblEmpPassOver.TypeTrans"
'MySQL = MySQL & " FROM         dbo.TblEmpPassOver INNER JOIN"
'MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TblEmpPassOver.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID RIGHT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblEmpPassOver.Emp_id = dbo.TblEmployee.Emp_ID"
'MySQL = MySQL & " where (dbo.TblEmpPassOver.TypeTrans <> 2 or dbo.TblEmpPassOver.TypeTrans is null) and TblEmpPassOver.AdvanceID = " & val(XPTxtID.Text)
MySQL = "SELECT     dbo.TblEmployee.Emp_Name, dbo.TblEmpPassOver.[interval], dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpPassOver.AdvanceDate, "
MySQL = MySQL & "                       dbo.TblEmpPassOver.BossNotes, dbo.TblEmpPassOver.Remark, dbo.TblEmpPassOver.TypeTrans, dbo.TblEmpPassOver.Expectedouttime,"
MySQL = MySQL & "                       dbo.TblEmpPassOver.ExpectedIntime, dbo.TblEmpPassOver.Actualouttime, dbo.TblEmpPassOver.ActualIntime, dbo.TblEmpPassOver.PostedDate,"
MySQL = MySQL & "                       dbo.TblEmpPassOver.NoteSerial"
MySQL = MySQL & " FROM         dbo.TblEmpPassOver INNER JOIN"
MySQL = MySQL & "                       dbo.TblEmpJobsTypes ON dbo.TblEmpPassOver.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee ON dbo.TblEmpPassOver.Emp_id = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & " where (dbo.TblEmpPassOver.TypeTrans <> 2 or dbo.TblEmpPassOver.TypeTrans is null) and TblEmpPassOver.AdvanceID = " & val(XPTxtID.text)
 
 End If
 If Me.Rd(2).value = True Then
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCasualVacation.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCasualVacationE.rpt"
        End If
 Else
 
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "PermissionRepAba.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "PermissionRepAba.rpt"
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
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(DcboBossName.text), "  ", DcboBossName.text)
        xReport.ParameterFields(3).AddCurrentValue IIf(IsNull(DcOutType.text), "  ", DcOutType.text)
        
    
        StrReportTitle = "" '& StrAccountName
 
    Else
    End If


    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    'xReport.ReportAuthor = App.title
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

Private Sub Command2_Click()
Dim Msg As String
Dim StrSQL As String
Dim X As Integer
         
          If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
    Else
        Msg = "Confirm Delete  "
    End If
        X = MsgBox(Msg, vbCritical + vbYesNo)

      If X = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "update   TblEmpPassOver set NoteSerial=null,NoteID=null Where AdvanceID=" & val(Me.XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        rs.Resync
        Retrive (val(Me.XPTxtID.text))
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
        Else
            Msg = " This voucher deleted  "
        End If
        MsgBox Msg
       End If
End Sub

Private Sub Command5_Click()
Dim StrSQL As String

              
          If ChekClodePeriod(XPDtbTrans.value) = True Then
                                 If SystemOptions.UserInterface = ArabicInterface Then
                                  MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «· ”ÊÌ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                                 Else
                                 MsgBox "Please Change Date Becouse This is Period is Closed"
                                End If
              Exit Sub
         End If

                
If TxtNoteSerial.text = "" Then
 CheckAccounts
createVoucher
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ «·ﬁÌœ"
        Else
            MsgBox "Done"
        End If
End If
End Sub
Function CheckAccounts() As Boolean
CheckAccounts = True
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
    Dim ColumnName As String
    Dim showinMosirVac(40) As Boolean
    Dim i As Integer
    sql = "select * from mofrad order by id  "
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
 
        For i = 1 To rs.RecordCount
            FixedOrChanged(i) = IIf(IsNull(rs("FixedOrChanged").value), 0, rs("FixedOrChanged").value)
            AddOrDiscount(i) = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
            ViewComp(i) = IIf(IsNull(rs("ViewComp").value), False, rs("ViewComp").value)
            Account_code(i) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            Account_code1(i) = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
            showMofradAll(i) = IIf(IsNull(rs("showMofradAll").value), False, rs("showMofradAll").value)
            culc30orRminder(i) = IIf(IsNull(rs("culc30orRminder").value), 0, rs("culc30orRminder").value)
            showinMosirVac(i) = IIf(IsNull(rs("showinMosirVac").value), False, rs("showinMosirVac").value)
      '      If Account_Code(i) = "" Then
      ''      MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & ViewComp(i), vbCritical
       '     getTitlesName = False
       '     Exit Function
       '     End If
            
            
            ZmamAccount(i) = IIf(IsNull(rs("ZmamAccount").value), 0, rs("ZmamAccount").value)
            AdvPaymentdAccount(i) = IIf(IsNull(rs("AdvPaymentdAccount").value), 0, rs("AdvPaymentdAccount").value)
            
            
    
              'AdvPaymentdAccount
            If SystemOptions.UserInterface = ArabicInterface Then
                componentname(i) = IIf(IsNull(rs("name").value), "", rs("name").value)
            Else
                componentname(i) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
            End If
             
              
            If ViewComp(i) = True And Account_code(i) = "" And (ZmamAccount(i) <> "True" And AdvPaymentdAccount(i) <> "True") Then
            MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & componentname(i), vbCritical
            CheckAccounts = False
          
           ' Unload Me
              Exit Function
            End If
          
             
              
         If SystemOptions.ProjectEmployeeGV = True And SystemOptions.ProjectDiscountPolicy = 1 Then 'xxx
                  If ViewComp(i) = True And AddOrDiscount(i) = -1 And Account_code1(i) = "" And (ZmamAccount(i) <> "True" And AdvPaymentdAccount(i) <> "True") Then
                MsgBox " ·„ Ì „ —»ÿ Õ”«» «·«Ì—«œ«  «· Ì  ⁄·Ì «·Œ’„ «·Œ«’ » " & componentname(i), vbCritical
        '        CheckAccounts = False
                
                '  Unload Me
                    Exit Function
                  End If
              
             End If
             
             
            rs.MoveNext
             
        Next i
  
    End If
 
    rs.Close
End Function
Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.text, , 200
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub
Private Sub DcboBossName_Change()
DcboBossName_Click (0)
End Sub



Private Sub DcbVacation_Change()
DcbVacation_Click (0)
End Sub

Private Sub DcbVacation_Click(Area As Integer)
ChDate
End Sub

Private Sub FromDate_Change()
ChDate
ShowJL
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub
Sub ClCulateAbcens()
Dim i As Integer
Dim Diff As Integer
Dim dY As Double
Dim Dy2 As Double
Dim Dy3 As Double
Dim Dy4 As Double
Dim TemDay As Double
Dim NoDay As Double
Dy4 = val(TxtNoDay.text) - val(TxtAbceDay.text)
Diff = DateDiff("m", FromDate.value, ToDate.value)
For i = 0 To Diff
If i = 0 Then
dY = day(FromDate.value)
TempDate = MonthLastDay(FromDate.value)
Dy2 = day(TempDate.value)
Dy3 = Dy2 - dY + 1
If Dy3 > Dy4 Then
NoDay = Dy3 - Dy4
Dy4 = 0
Else
NoDay = 0
Dy4 = Dy4 - Dy3
End If
 
 If TypeDisc(1).value Then
    SaveDataAbcens val(DcboEmpName.BoundText), val(dcBranch.BoundText), GetMofrad(Me.DcbMang.BoundText, 1), NoDay, FromDate.value
Else
    SaveDataAbcens val(DcboEmpName.BoundText), val(dcBranch.BoundText), GetMofrad(Me.DcbMang.BoundText), NoDay, FromDate.value
End If
ElseIf i = Diff Then
Dy2 = day(ToDate.value) - 1
Dy3 = Dy2
If Dy3 > Dy4 Then
NoDay = Dy3 - Dy4
Dy4 = 0
Else
NoDay = 0
Dy4 = Dy4 - Dy3
End If

 
 If TypeDisc(1).value Then
    SaveDataAbcens val(DcboEmpName.BoundText), val(dcBranch.BoundText), GetMofrad(Me.DcbMang.BoundText, 1), NoDay, ToDate.value
    
Else
    SaveDataAbcens val(DcboEmpName.BoundText), val(dcBranch.BoundText), GetMofrad(Me.DcbMang.BoundText, 0), NoDay, ToDate.value
End If


 
Else
   Dim str As String
    str = "01/" & Month(FromDate.value) & "/" & year(FromDate.value)
    TempDate.value = MonthLastDay(CDate(str))
TempDate.value = DateAdd("m", i, TempDate.value)
Dy2 = day(TempDate.value)
Dy3 = Dy2
If Dy3 > Dy4 Then
NoDay = Dy3 - Dy4
Dy4 = 0
Else
NoDay = 0
Dy4 = Dy4 - Dy3
End If


 
 If TypeDisc(1).value Then
    SaveDataAbcens val(DcboEmpName.BoundText), val(dcBranch.BoundText), GetMofrad(Me.DcbMang.BoundText, 0), NoDay, TempDate.value
    
Else
    SaveDataAbcens val(DcboEmpName.BoundText), val(dcBranch.BoundText), GetMofrad(Me.DcbMang.BoundText), NoDay, TempDate.value
End If



 
End If
Next i
End Sub
Function check_employee_accounts() As Boolean
    Dim Employee_account As String
    Dim error_string As String
    error_string = ""
    check_employee_accounts = True
    Dim i As Integer

    With Grid

        For i = .FixedRows To .rows - 2
                   If val(.TextMatrix(i, .ColIndex("BranchId"))) = 0 Then
                   error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡    ÕœÌœ «·›—⁄ «· «»⁄ ·Â"
        
                check_employee_accounts = False
                   End If
                   
                   
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code")

            If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«» –„ …"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«» –„ … ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
  Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
                    If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«» «·«ÃÊ— «·„” Õﬁ…"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«» «·«ÃÊ— «·„” Õﬁ… ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
            
  Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3")
                    If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«»   «·„œ›Ê⁄«  «·„ﬁœ„…"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«»    «·„œ›Ê⁄«  «·„ﬁœ„… ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
            '     If Val(.TextMatrix(i, .ColIndex("Emp_Salary"))) = 0 Then
            '     error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & " ·„ Ì „  ÕœÌœ —« » «”«”Ì ·Â  " & vbCrLf
            '
            '    check_employee_accounts = False
            '
            '     End If
            If error_string <> "" Then
           CreatLog_File_for_error (error_string)
       End If
        Next i

    End With

    Dim X As Integer
    Dim StrLogFileName As String

    If error_string <> "" Then
        X = MsgBox("Â·  —Ìœ › Õ «·„·› ··„—«Ã⁄Â", vbCritical + vbYesNo, "ÌÊÃœ Œÿ√ ›Ì Õ”«»«  «·„ÊŸ›Ì‰  Ì„ﬂ‰ „—«Ã⁄ … ›Ì „·› «·«Œÿ«¡")

        If X = vbYes Then
            StrLogFileName = App.path & "\employee_account_error.txt"
            ShellExecute 0&, vbNullString, StrLogFileName, vbNullString, vbNullString, vbNormalFocus
        End If
    End If

End Function
Public Sub CreatLog_File_for_error(str As String)
    Dim StrLogFileName As String
    Dim IntFreeFile As Integer
    Dim ss As String

    StrLogFileName = App.path & "\employee_account_error.txt"

    If Dir(StrLogFileName) <> "" Then
        Kill StrLogFileName
    End If

    ss = "»Ì«‰ »«”„«¡ «·„ÊŸ›Ì‰ «·–Ì‰ ·œÌÂ„ „‘«ﬂ·  "
    ss = ss & vbCrLf & "Byte Informations Systems "
    ss = ss & vbCrLf & "BYTE "
    ss = ss & vbCrLf & "Create Date:- " & Now
    ss = ss & vbCrLf & str & vbCrLf
    IntFreeFile = FreeFile

    Open StrLogFileName For Output As #IntFreeFile
    Print #IntFreeFile, ss
    Close #IntFreeFile
End Sub

Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "   «·«Ã«“«  «·⁄«—÷…" & XPTxtID.text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
tablename = "TblEmpPassOver"
Filedname = "AdvanceID"
NoteSerial1 = val(XPTxtID)
Notevalue = 0
notytype = 9091
Notevalue = val(txtSalary.text)
BranchID = val(dcBranch.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
        
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TXTNoteID.text = NoteID
                                                     TxtNoteSerial.text = NoteSerial

CREATE_VOUCHER_GE val(TXTNoteID.text), BranchID, user_id, NoteDate
rs.Resync adAffectCurrent
     End If
End Function

Private Sub Rd_Click(index As Integer)
ELe(0).Visible = False
ELe(16).Visible = False
If index = 2 Or index = 3 Then
Hid
ShowJL
ELe(0).Visible = True
Else
ELe(16).Visible = True
End If
End Sub
Sub Hid()
If Rd(3).value = True Then
lbl(63).Visible = False
DcbVacation.Visible = False
TxtAbceDay.Visible = False
lbl(72).Visible = False
TypeDisc(0).Visible = False
TypeDisc(1).Visible = False
TypeDisc(2).Visible = False
TxtBalanceDay.Visible = False
lbl(67).Visible = False
lbl(71).Visible = False
lbl(66).Visible = False
TxtMaxDay.Visible = False
RdTypeVaction(0).Visible = True
RdTypeVaction(1).Visible = True
Else
RdTypeVaction(0).Visible = False
RdTypeVaction(1).Visible = False
lbl(66).Visible = True
lbl(63).Visible = True
DcbVacation.Visible = True
TxtAbceDay.Visible = True
lbl(72).Visible = True
TypeDisc(0).Visible = True
TypeDisc(1).Visible = True
TypeDisc(2).Visible = True
TxtBalanceDay.Visible = True
lbl(67).Visible = True
lbl(71).Visible = True
TxtMaxDay.Visible = True
End If
End Sub
Sub ChDate()
If Me.TxtModFlg.text <> "R" Then
If val(DcboEmpName.BoundText) <> 0 Then
TxtNoDay.text = DateDiff("d", FromDate.value, ToDate.value) + 1
If DcbVacation.BoundText = "" Then Exit Sub
RetrivVaction val(DcbVacation.BoundText)
TxtBalanceDay.text = val(TxtBalanceDay.text) - SumNoVaction()

TxtAbceDay.text = 0
If val(TxtNoDay.text) > val(TxtMaxDay.text) Then
TxtAbceDay.text = val(TxtNoDay.text) - val(TxtMaxDay.text)
End If
If (val(TxtMaxDay.text) > val(TxtBalanceDay.text)) And (val(TxtNoDay.text) > val(TxtBalanceDay.text)) Then
TxtAbceDay.text = val(TxtNoDay.text) - val(TxtBalanceDay.text)
End If
TxtNoVaction.text = val(TxtNoDay.text) - val(TxtAbceDay.text)
TxtNoVaction.text = Abs(TxtNoVaction.text)
ShowJL
Else
ShowJL
Exit Sub
End If
End If
End Sub

Private Sub ToDate_Change()
ChDate
ShowJL
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    Dim BossId As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
        
        GetEmployeeIDFromCode TxtSearchCodeB.text, BossId
        DcboBossName.BoundText = BossId
    End If

End Sub

 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 11
        Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If

End Sub
Private Sub DcboBossName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboBossName.BoundText) = 0 Then Exit Sub

    Dim BossCode  As String
    
    GetEmployeeIDFromCode , , DcboBossName.BoundText, BossCode
    TxtSearchCodeB.text = BossCode
    
   If Me.TxtModFlg = "R" Then Exit Sub
   End Sub

Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
    
    Dim Myrs As New ADODB.Recordset
    
    Dim myStrSQL As String
    
    myStrSQL = "select mangerid,GroupID  from TblEmployee where Emp_Id = " & val(DcboEmpName.BoundText)
    
    Myrs.Open myStrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    
    DcboBossName.BoundText = IIf(IsNull(Myrs("mangerid").value), "", Myrs("mangerid").value)
    DCGroupID.BoundText = IIf(IsNull(Myrs("GroupID").value), "", Myrs("GroupID").value)
    
   If Me.TxtModFlg = "R" Then Exit Sub
   
   
    Dim StrSQL As String

 
        
        
        Dim IssueDate As Date
        Dim DepID As Double
        Dim specid As Double
        Dim JobTypeID As Double
        Dim gradeID As Double
        Dim Account_code2 As String
           Dim Account_code  As String
        Dim Balance As String
        Dim endContractPerMonth As Double
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth
        
          WriteCustomerBalPublic Account_code2, Balance
          
          lbl(22).Caption = val(Balance)

          WriteCustomerBalPublic Account_code, Balance
          
  lbl(21).Caption = val(Balance)
  lbl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
        DBIssueDate.value = IssueDate
        DcbMang.BoundText = DepID
        DcboEmpDepartments.BoundText = DepID
        DcboSpecifications.BoundText = gradeID
        DcboJobsType.BoundText = JobTypeID
        lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
   ShowJL
    'End If

End Sub
Sub ShowJL()
     If SystemOptions.CreateJLVactionAratha = True And Rd(2).value = True Then
        ShowComponent
        C1Elastic3.Visible = True
        ELe(11).Visible = True
        Else
        ELe(11).Visible = False
        C1Elastic3.Visible = False
        End If
End Sub

Private Sub XPDtbTrans_Change()

    If Trim(txtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = txtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    txtNoteSerial1.text = ""

End Sub
Function SumNoVaction() As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     SUM(NoVaction) AS SumNoVaction, Emp_id, YEAR(AdvanceDate) AS YerID"
sql = sql & " From dbo.TblEmpPassOver"
sql = sql & " Where (Emp_id = " & val(DcboEmpName.BoundText) & ")"
sql = sql & " And (year(AdvanceDate) = " & year(FromDate.value) & ")"
sql = sql & " And (year(AdvanceDate) =" & year(ToDate.value) & ")"
sql = sql & " GROUP BY Emp_id, YEAR(AdvanceDate)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
SumNoVaction = IIf(IsNull(Rs3("SumNoVaction").value), 0, Rs3("SumNoVaction").value)
Else
SumNoVaction = 0
End If
End Function
Private Sub Dcbranch_Click(Area As Integer)
 
    TxtNoteSerial.text = ""
    txtNoteSerial1.text = ""
End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        .RowHeightMin = 300
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

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
    YearMonth
    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetOutType Me.DcOutType
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmployees Me.DcboBossName
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetEmpDepartments Me.DcbMang
    Dcombos.GetEmpDepartments Me.DcboEmpDepartments
    Dcombos.GetEmpJobsTypes Me.DcboJobsType
    Dcombos.GetEmpGrades Me.DcboSpecifications
    Dcombos.GetEmpVactionAreth Me.DcbVacation
        
     Dcombos.GetEmpLocations Me.DCGroupID
     
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEmpPassOver     Order By AdvanceID"
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
Function ChecPeriodSalary(Optional TemDate As Date, Optional ByRef MonthID As Integer, Optional ByRef YearID As Integer) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from TblDurations2Salary where  FromDate<=" & SQLDate(TemDate, True) & ""
sql = sql & " and ToDate>=" & SQLDate(TemDate, True) & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
ChecPeriodSalary = True
 MonthID = IIf(IsNull(Rs3("MonthID").value), 0, Rs3("MonthID").value)
YearID = IIf(IsNull(Rs3("YearID").value), 0, Rs3("YearID").value)
Else
ChecPeriodSalary = False
End If
End Function
Private Sub SaveDataAbcens(Optional Emp_id As Integer, Optional BranchID As Integer, Optional MofrdID As Integer, Optional NoofDays As Double, Optional RecDate As Date)
    Dim Msg As String
    Dim BasicSalary As Double
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim HourRate As Double
    Dim value As Double
    Dim Rs3 As ADODB.Recordset
    Dim MonthID As Integer
    Dim YearID As Integer
    Dim i As Integer
     Dim NoOfHours As Double
    Dim Equation As Double
 If NoofDays = 0 Then Exit Sub
 
            StrSQL = "Delete From TblChangedComponentRegister Where CasualVID=" & val(Me.XPTxtID.text) & " "
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblChangedComponentRegisterDetails Where CasualVID=" & val(Me.XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
 For i = 1 To NoofDays
    '-------------------------------------------------------------------------------------------
    If i = 1 Then
    ReDta.value = RecDate
    Else
    ReDta.value = DateAdd("d", 1, ReDta.value)
    End If
        If ChecPeriodSalary(ReDta.value, MonthID, YearID) = True Then
        CboYear.text = YearID
    CmbMonth.ListIndex = MonthID - 1
    Else
        CboYear.text = year(ReDta.value)
    CmbMonth.ListIndex = Month(ReDta.value) - 1
    End If
    Dim EmployeeSalary As Double
    Dim NoDayMonth As Integer


                LBLWhereSTR.Caption = GetSpecificComponentIncalculations(MofrdID, Equation)
                    EmployeeSalary = GetEmployeeSalaryAccordingToComponent(Emp_id, LBLWhereSTR)
                    '«Ì«„
                    If SystemOptions.MonthIs30days = True Then
                    
                        HourRate = (EmployeeSalary / 30)
                    Else
                        
                       ' HourRate = (EmployeeSalary * 12 / 365)
                       
                       HourRate = (EmployeeSalary / GetDaysInMonth(val(CboYear.text), val(Month(ReDta.value))))
                       
                    End If
                    value = Round(HourRate * 1, SystemOptions.EmpComponentDigts)
                     BasicSalary = EmployeeSalary
                    NoOfHours = 0

    Set Rs3 = New ADODB.Recordset
    StrSQL = "select * from TblChangedComponentRegister where 1=-1"
    Rs3.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        Me.txtID.text = CStr(new_id("TblChangedComponentRegister", "ChangedComponentid", "", True))
        Rs3.AddNew
        Rs3("ChangedComponentid").value = val(Me.txtID.text)
    Rs3("RecordDate").value = ReDta.value
    Rs3("year").value = year(ReDta.value)   'val(CboYear.ListIndex)
    Rs3("month").value = CmbMonth.ListIndex
    Rs3("Actualyear").value = val(CboYear.text)
    Rs3("Actualmonth").value = val(CmbMonth.ListIndex) + 1
    Rs3("ComponentID").value = MofrdID
    Rs3("BranchId").value = BranchID
    Rs3("Finger").value = 1
    Rs3("CasualVID").value = val(XPTxtID.text)
    Rs3.update
 Dim xx As Integer
    Set RsDev = New ADODB.Recordset
    StrSQL = " SELECT     * FROM         dbo.TblChangedComponentRegisterDetails WHERE     (Emp_ID = - 1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If Emp_id <> 0 Then
                RsDev.AddNew
                RsDev("CasualVID").value = val(XPTxtID.text)
                RsDev("ChangedComponentid").value = val(txtID.text)
                RsDev("Emp_ID").value = Emp_id
                RsDev("NoofDays").value = 1
                RsDev("HourRate").value = HourRate
             '   RsDev("NoOfHour").value = NoOfHours
                RsDev("Salary").value = BasicSalary
                RsDev("Value").value = value
                RsDev.update
            End If
       Next i
End Sub

Function GetDaysInMonth(ByVal year As Integer, ByVal Month As Integer) As Integer
    Dim daysInMonth As Integer
   ' daysInMonth = DateTime.daysInMonth(year, month)
    
     Select Case Month
        Case 1, 3, 5, 7, 8, 10, 12
            daysInMonth = 31
        Case 4, 6, 9, 11
            daysInMonth = 30
        Case 2
            If ((year Mod 4 = 0) And (year Mod 100 <> 0)) Or (year Mod 400 = 0) Then
                daysInMonth = 29
            Else
                daysInMonth = 28
            End If
    End Select
    
    GetDaysInMonth = daysInMonth
End Function

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    'Label1.Visible = False
    RdTypeVaction(0).Caption = "Discount from Vacation"
    RdTypeVaction(0).Caption = "Discount from Vorking Days"
    
    XPLbl(46).Caption = "Location"
Cmd(9).Caption = "Print"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Accredit.Caption = "Send To Approv."
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
lbl(25).Caption = "Management"
    Me.Caption = "Temparary vacation-authorized exit"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(3).Caption = "Employee"
    lbl(33).Caption = "Type Permit"
    lbl(39).Caption = "Branch"
    lbl(2).Caption = "Long"
    lbl(29).Caption = "Hour"
    lbl(40).Caption = "Remarks"
    Rd(0).RightToLeft = False
    Rd(1).RightToLeft = False
    Rd(2).RightToLeft = False
    Rd(3).RightToLeft = False
    TypeDisc(0).RightToLeft = False
    TypeDisc(1).RightToLeft = False
    TypeDisc(0).Caption = "Discunt from Salary"
    TypeDisc(1).Caption = "Discount From Vacation"
    
    Rd(0).Caption = "Temporary exit"
    Rd(1).Caption = "Permission"
    lbl(72).Caption = "No Absences Day"
    Rd(2).Caption = "Casual Vacation"
    Rd(3).Caption = "Unpaid Vacation"
    lbl(65).Caption = "Remarks"
    lbl(66).Caption = "Balance Holiday"
    lbl(67).Caption = "Day"
    lbl(71).Caption = "Max. Period"
    lbl(68).Caption = "From"
    lbl(69).Caption = "To"
    lbl(70).Caption = "No Days Holiday"
lbl(31).Caption = "Time Out Expected"
lbl(32).Caption = "Time Return Expected   "
lbl(34).Caption = "Actual Time Out"
lbl(35).Caption = "Back Actual Time'"
lbl(28).Caption = "Task"
XPTab301.Caption = "Approve|Data"

lbl(63).Caption = "Vacation"
lbl(41).Caption = "Manager"
Frame3.Caption = "Transaction Type"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"

    With Me.Grid2
    .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
      .TextMatrix(0, .ColIndex("EmpName")) = "EmpName"
    End With

End Sub

Private Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2006 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next
    CboYear.ListIndex = IntDefIndex
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
'Accredit.Enabled = False
    Select Case Me.TxtModFlg.text
  

        Case "R"
         ' Accredit.Enabled = True
              If SystemOptions.UserInterface = ArabicInterface Then
                                                  Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                 Else
                                                  Accredit.Caption = " send to Approval   "
                                             End If
                  Me.Caption = " ’—ÌÕ Œ—ÊÃ „ƒﬁ "
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
            TxtInterval.locked = True
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
            TxtInterval.locked = False
            Me.DcboBox.locked = False
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
            TxtInterval.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtPaymentCounts_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtPaymentCounts.text, 1)
End Sub

Private Sub TxtPaymentCounts_LostFocus()

    If val(TxtPaymentCounts.text) > 84 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "«·œ›«⁄  «ﬂ»— „‰ «·Õœ ", vbOKOnly, App.Title
    Else
    MsgBox "Payments greater than the limit ", vbOKOnly, App.Title
    End If
        Exit Sub
    End If
 
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
Sub RetrivVaction(Optional ID As Double)
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "SELECT     NoDay, MaxDay, Vac_ID"
sql = sql & " From dbo.tblVacancy"
sql = sql & " WHERE     (VacAretha = 1) AND (Vac_ID = N'" & ID & "')"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
TxtMaxDay.text = IIf(IsNull(Rs3("MaxDay").value), 0, Rs3("MaxDay").value)
TxtBalanceDay.text = IIf(IsNull(Rs3("NoDay").value), 0, Rs3("NoDay").value)
Else
TxtBalanceDay.text = 0
TxtMaxDay.text = 0
End If
End Sub
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

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
            rs.Find "AdvanceID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
''////////
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
   txtSalary = IIf(IsNull(rs("salary").value), "", (rs("salary").value))
   TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
   Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), 0, (rs("NoteID").value))
Me.DcbMang.BoundText = IIf(IsNull(rs("MangID").value), 0, (rs("MangID").value))
TxtNoVaction.text = IIf(IsNull(rs("NoVaction").value), 0, (rs("NoVaction").value))
TxtBalanceDay.text = IIf(IsNull(rs("BalanceDay").value), "", (rs("BalanceDay").value))
DcbVacation.BoundText = IIf(IsNull(rs("VacationID").value), "", (rs("VacationID").value))
TxtMaxDay.text = IIf(IsNull(rs("MaxDay").value), "", (rs("MaxDay").value))
TxtNoDay.text = IIf(IsNull(rs("NoDay").value), "", (rs("NoDay").value))
TxtAbceDay.text = IIf(IsNull(rs("AbceDay").value), "", (rs("AbceDay").value))
TxtRemark2.text = IIf(IsNull(rs("Remark2").value), "", (rs("Remark2").value))
FromDate.value = IIf(IsNull(rs("FromDate").value), Date, (rs("FromDate").value))
ToDate.value = IIf(IsNull(rs("ToDate").value), Date, (rs("ToDate").value))
Me.DCGroupID.BoundText = IIf(IsNull(rs("GroupID").value), "", rs("GroupID").value)

If Not IsNull(rs("RdTypeVaction").value) Then
If (rs("RdTypeVaction").value) = 0 Then
RdTypeVaction(0).value = True
ElseIf (rs("RdTypeVaction").value) = 1 Then
RdTypeVaction(1).value = True
End If
Else
RdTypeVaction(0).value = True
End If

If Not IsNull(rs("TypeTrans").value) Then
If (rs("TypeTrans").value) = 0 Then
Rd(0).value = True
ElseIf (rs("TypeTrans").value) = 1 Then
Rd(1).value = True
ElseIf (rs("TypeTrans").value) = 2 Then
Rd(2).value = True
ElseIf (rs("TypeTrans").value) = 3 Then
Rd(3).value = True
End If
Else
Rd(0).value = True
End If
If Not IsNull(rs("TypeDisc").value) Then
If (rs("TypeDisc").value) = 0 Then
TypeDisc(0).value = True
ElseIf (rs("TypeDisc").value) = 1 Then
TypeDisc(1).value = True
ElseIf (rs("TypeDisc").value) = 2 Then
TypeDisc(2).value = True
End If
Else
TypeDisc(2).value = True
End If



'////
    XPTxtID.text = IIf(IsNull(rs("AdvanceID").value), "", (rs("AdvanceID").value))
    XPDtbTrans.value = IIf(IsNull(rs("AdvanceDate").value), Date, rs("AdvanceDate").value)
    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
  
    DcOutType.BoundText = IIf(IsNull(rs("OutTypeID").value), "", rs("OutTypeID").value)
  DcboEmpDepartments.BoundText = IIf(IsNull(rs("DeparmentID").value), "", rs("DeparmentID").value)
  DcboJobsType.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
  txtRemark.text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
  bossNotes.text = IIf(IsNull(rs("BossNotes").value), "", rs("BossNotes").value)
  Me.DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
  Me.DcboBossName.BoundText = IIf(IsNull(rs("BossId").value), "", rs("BossId").value)
    TxtInterval.text = IIf(IsNull(rs("Interval").value), "", rs("Interval").value)
  '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
  
   TxtExpectedouttime.value = IIf(IsNull(rs("Expectedouttime").value), "", rs("Expectedouttime").value)
 txtExpectedIntime.value = IIf(IsNull(rs("ExpectedIntime").value), "", rs("ExpectedIntime").value)
 txtActualouttime.value = IIf(IsNull(rs("Actualouttime").value), "", rs("Actualouttime").value)
 txtActualIntime.value = IIf(IsNull(rs("ActualIntime").value), "", rs("ActualIntime").value)

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
    ShowJL
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Public Function MonthLastDay(ByVal dCurrDate As Date) As Date
    Dim dFirstDayNextMonth As Date
  
    MonthLastDay = Empty
    dCurrDate = Format(dCurrDate, "DD/MM/YYYY")
  
    dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
    MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
  
    Exit Function
 
End Function

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If Me.DcboEmpName.text = "" Or val(DcboEmpName.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
            Else
            Msg = "Please Select Employee"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboEmpName.SetFocus
           Sendkeys "{F4}"
            Exit Sub
        End If
        
    If Me.DcbMang.text = "" Or val(DcbMang.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «·«œ«—… „‰ „·› «·„ÊŸ›Ì‰..!! "
            Else
            Msg = "Please Select Management"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcbMang.SetFocus
            Sendkeys "{F4}"
            Exit Sub
     End If
        If Rd(2).value = True And TypeDisc(0).value = True Then
     If GetMofrad((Me.DcbMang.BoundText)) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "Ì—ÃÏ «Œ Ì«— „›—œ «·€Ì«» „‰ ‘«‘… «·«œ«—« "
     Else
     MsgBox "Please Select Component From Screen of Management"
     End If
     Exit Sub
     End If
    End If
       ' If Me.DcboBossName.BoundText = "" Then
       '     Msg = "ÌÃ»  ÕœÌœ «”„ «·„”ƒÊ·..!! "
       '     MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       '     DcboBossName.SetFocus
       '     'SendKeys "{F4}"
       '     Exit Sub
       ' End If

 If Rd(2).value = True Then
 If val(DcbVacation.BoundText) = 0 Or DcbVacation.text = "" Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "Ì—ÃÏ «Œ Ì«— «·«Ã«“…"
     Else
     MsgBox "Please Select Vacation"
   End If
   Exit Sub
 End If
 End If
 
'
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("TblEmpPassOver", "AdvanceID", "", True))
       '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
       '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
  StrSQL = "Delete From tblVacationData Where CasualVID='" & val(XPTxtID.text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
               StrSQL = "Delete From TblChangedComponentRegister Where CasualVID=" & val(Me.XPTxtID.text) & " "
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblChangedComponentRegisterDetails Where CasualVID=" & val(Me.XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblInforVacatiom Where prkid=" & val(Me.XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        

        End If

        rs("branch_no").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
 
'     rs("AdvanceID").value = val(XPTxtID.Text)
If TypeDisc(0).value = True Then
rs("TypeDisc").value = 0
ElseIf TypeDisc(1).value = True Then
rs("TypeDisc").value = 1
ElseIf TypeDisc(2).value = True Then
rs("TypeDisc").value = 2
End If
If Rd(0).value = True Then
rs("TypeTrans").value = 0
ElseIf Rd(1).value = True Then
rs("TypeTrans").value = 1
ElseIf Rd(2).value = True Then
rs("TypeTrans").value = 2
ElseIf Rd(3).value = True Then
rs("TypeTrans").value = 3
End If
If RdTypeVaction(1).value = True Then
rs("RdTypeVaction").value = 1
Else
rs("RdTypeVaction").value = 0
End If

        rs("Salary").value = val(txtSalary.text)
        rs("MangID").value = val(DcbMang.BoundText)
        rs("NoVaction").value = val(TxtNoVaction.text)
        rs("VacationID").value = val(DcbVacation.BoundText)
        rs("BalanceDay").value = val(TxtBalanceDay.text)
        rs("Remark2").value = (TxtRemark2.text)
        rs("AbceDay").value = val(TxtAbceDay.text)
        rs("NoDay").value = val(TxtNoDay.text)
        rs("MaxDay").value = val(TxtMaxDay.text)
        rs("ToDate").value = ToDate.value
        rs("FromDate").value = FromDate.value
        rs("AdvanceDate").value = XPDtbTrans.value
        rs("Emp_ID").value = val(Me.DcboEmpName.BoundText)
        rs("BossId").value = val(Me.DcboBossName.BoundText)
        rs("DeparmentID").value = val(Me.DcboEmpDepartments.BoundText)
        rs("JobTypeID").value = val(Me.DcboJobsType.BoundText)
        rs("Remark").value = IIf(txtRemark.text = "", Null, (txtRemark.text))
        rs("BossNotes").value = IIf(bossNotes.text = "", Null, (bossNotes.text))
        rs("interval").value = IIf(TxtInterval.text = "", Null, val(TxtInterval.text))
        rs("UserID").value = Me.DCboUserName.BoundText
        rs("OutTypeID").value = val(Me.DcOutType.BoundText)
        rs("Expectedouttime").value = TxtExpectedouttime.value
        rs("ExpectedIntime").value = txtExpectedIntime.value
        rs("Actualouttime").value = txtActualouttime.value
        rs("ActualIntime").value = txtActualIntime.value
        If val(Me.DCGroupID.BoundText) = 0 Then
            rs("GroupID").value = Null
        Else
            rs("GroupID").value = val(Me.DCGroupID.BoundText)
        End If
      '  DB_CreateField "TblEmpPassOver", "GroupID", adInteger, adColNullable, , , "  ", False, True
         rs.update
   If val(TxtNoDay.text) <> 0 Then
 If Rd(2).value = True And TypeDisc(1).value = True Then
 SaveVacation val(DcboEmpName.BoundText), val(TxtNoDay.text)
    End If
   'If Rd(2).value = True And TypeDisc(0).value = True Then
   If Rd(2).value = True Then
   Dim Diff As Integer
   Diff = DateDiff("m", FromDate.value, ToDate.value)
SaveInformationVacation 0, val(DcboEmpName.BoundText), val(TxtAbceDay.text)
   If Diff = 0 Then
    
    If TypeDisc(1).value Then
        SaveDataAbcens val(DcboEmpName.BoundText), val(dcBranch.BoundText), GetMofrad(Me.DcbMang.BoundText, 1), val(TxtAbceDay.text), FromDate.value
    Else
    
        SaveDataAbcens val(DcboEmpName.BoundText), val(dcBranch.BoundText), GetMofrad(Me.DcbMang.BoundText), val(TxtAbceDay.text), FromDate.value
    End If
   Else
   ClCulateAbcens

    End If
    End If
  End If
        Cn.CommitTrans
        BeginTrans = False
    
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
             Else
             Msg = "Thi is record alredy saved" & CHR(13)
             Msg = Msg & "you want enter another record"
             End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
        End Select

        TxtModFlg.text = "R"
        Retrive
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
       Else
        Msg = "Can not save this data" & CHR(13)
        Msg = Msg + "It has been insert incorrect data " & CHR(13)
        Msg = Msg + "Make sure of the accuracy of the data and try again"
       End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
   Else
   Msg = "Sorry ...error douring save"
   End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Sub SaveInformationVacation(Optional TypeVacation As Integer = 0, Optional EmpID As Integer = 0, Optional NoDay As Double = 0)
Dim sql As String
Dim str As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If SystemOptions.UserInterface = ArabicInterface Then
str = " «Ã«“… ⁄—÷…"
Else
str = "Casual Vacation"
End If
sql = "select * from TblInforVacatiom where (1=-1)"
    Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      Rs7.AddNew
      Rs7("PrkID").value = val(XPTxtID.text)
      Rs7("EmpID").value = EmpID
      Rs7("NoDay").value = (NoDay)
      Rs7("RecordDate").value = XPDtbTrans.value
      Rs7("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
      Rs7("TypeVacation").value = TypeVacation
      Rs7("Remarks").value = str
      Rs7.update
End Sub
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "AdvanceID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

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

    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
     Else
     Msg = "Confirm Delete "
     End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
    Deletepost Me.Name, "TblEmpPassOver", "AdvanceID", val(DcbMang.BoundText), val(dcBranch.BoundText), val(XPTxtID.text), XPTxtID
    StrSQL = "Update TblEmployee Set  jopstatusid=1,workstate=1 Where Emp_ID=" & GetEmIDUnpaidVacation(val(XPTxtID.text)) & ""
                                      Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From tblVacationData Where CasualVID='" & val(XPTxtID.text) & "'"
          Cn.Execute StrSQL, , adExecuteNoRecords
           StrSQL = "Delete From TblChangedComponentRegister Where CasualVID=" & val(Me.XPTxtID.text) & " "
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblChangedComponentRegisterDetails Where CasualVID=" & val(Me.XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblInforVacatiom Where CasualVID=" & val(Me.XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
            If Not rs.RecordCount < 1 Then
                rs.delete
          
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This process is not available as there are no records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry...error douring delete"
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub



Function FillApprovedTable()
Exit Function
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
StrSQL = StrSQL + " FROM         dbo.ApprovalData left  JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
   Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label3.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label3.Caption = "Approved"
                                 End If
                            Label3.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label3.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label3.Caption = "Currently required Approve"
                            End If
                 Label3.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 Grid2.rows = 1
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
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
                 If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
                Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
                Else
                 Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
                 End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
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
 Grid2.rows = 1
    End If
RsDetails.Close

End Function

Sub SaveVacation(Optional EmpID As Double = 0, Optional NoDay As Double = 0)
Dim sql As String
Dim str As String
If Abs(NoDay) = 0 Then Exit Sub
If SystemOptions.UserInterface = ArabicInterface Then
str = " «Ã«“… ⁄—÷…"
Else
str = "Casual Vacation"
End If
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "select * from tblVacationData where (1=-1)"
    Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      Rs7.AddNew
      Rs7("CasualVID").value = val(XPTxtID.text)
      Rs7("EmpID").value = EmpID
      Rs7("Value").value = (NoDay * -1)
     Rs7("Remark").value = str
      Rs7.update
End Sub


Function GetMofrad(Optional DeparmentID As Integer, Optional ByVal mType As Integer = 0) As Integer
Dim sql As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
sql = "Select * from TblEmpDepartments where DeparmentID=" & DeparmentID & ""
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
If mType = 0 Then
    GetMofrad = IIf(IsNull(Rs6("AbscenID").value), 0, Rs6("AbscenID").value)
Else
    GetMofrad = IIf(IsNull(Rs6("MokafahVacID").value), 0, Rs6("MokafahVacID").value)
End If

Else
GetMofrad = 0
End If
End Function


Function GetComponentValuePerBranch(BramchId As Integer, componentname As String) As Double
    Dim SUM As Double
    SUM = 0
    Dim i As Integer

    With Grid

        For i = .FixedRows To .rows - 2
    
            If val(.TextMatrix(i, .ColIndex(componentname))) > 0 And val(.TextMatrix(i, .ColIndex("BranchId"))) = BramchId Then
                SUM = SUM + val(.TextMatrix(i, .ColIndex(componentname)))
            End If

        Next i

    End With

    GetComponentValuePerBranch = SUM
End Function
Function getTitlesName() As Boolean
Grid.ColHidden(Grid.ColIndex("TotalAdvance")) = False
getTitlesName = True
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
    Dim ColumnName As String
    Dim i As Integer
    sql = "select * from mofrad order by id  "
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
 
        For i = 1 To rs.RecordCount
            FixedOrChanged(i) = IIf(IsNull(rs("FixedOrChanged").value), 0, rs("FixedOrChanged").value)
            AddOrDiscount(i) = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
            ViewComp(i) = IIf(IsNull(rs("ViewComp").value), False, rs("ViewComp").value)
            Account_code(i) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
             Account_code1(i) = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
             Account_code1(i) = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
            showMofradAll(i) = IIf(IsNull(rs("showMofradAll").value), False, rs("showMofradAll").value)
            culc30orRminder(i) = IIf(IsNull(rs("culc30orRminder").value), 0, rs("culc30orRminder").value)
      '      If Account_Code(i) = "" Then
      ''      MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & ViewComp(i), vbCritical
       '     getTitlesName = False
       '     Exit Function
       '     End If
            
            
            ZmamAccount(i) = IIf(IsNull(rs("ZmamAccount").value), 0, rs("ZmamAccount").value)
            AdvPaymentdAccount(i) = IIf(IsNull(rs("AdvPaymentdAccount").value), 0, rs("AdvPaymentdAccount").value)
            
            

            
            
              'AdvPaymentdAccount
            If SystemOptions.UserInterface = ArabicInterface Then
                componentname(i) = IIf(IsNull(rs("name").value), "", rs("name").value)
            Else
                componentname(i) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
            End If
             
             
         '   If ViewComp(i) = True And Account_Code(i) = "" And (ZmamAccount(i) <> "True" And AdvPaymentdAccount(i) <> "True") Then
         '   MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & componentname(i), vbCritical
         '   getTitlesName = False
          
           ' Unload Me
         '     Exit Function
         '   End If
              
              
            With Me.Grid
             
                ColumnName = "Comp" & i

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(0, .ColIndex(ColumnName)) = IIf(IsNull(rs("name").value), "", rs("name").value)
                Else
                    .TextMatrix(0, .ColIndex(ColumnName)) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                End If
                     
                If ViewComp(i) = True Then
                    .ColHidden(.ColIndex(ColumnName)) = False
                Else
                    .ColHidden(.ColIndex(ColumnName)) = True
                End If
                     
            End With
             
 
             
            rs.MoveNext
             
        Next i
  
    End If
 
    rs.Close
End Function
Private Sub ShowComponent()
    On Error Resume Next

If DcboEmpName.BoundText = "" Then Exit Sub
'firstrun = False
    If getTitlesName = True Then
   End If
    DoEvents
    FillGridWithData
 
    Dim i As Integer
        With Grid
For i = 1 To 40

                 If val((.TextMatrix(.rows - 1, .ColIndex("Comp" & i & "")))) = 0 Then
                   .ColHidden(.ColIndex("Comp" & i)) = True
                End If


                If val((.TextMatrix(.rows - 1, .ColIndex("sgn")))) = 0 Then
                  .ColHidden(.ColIndex("sgn")) = True
                End If
               If val((.TextMatrix(.rows - 1, .ColIndex("TotalAdvance")))) = 0 Then
                  .ColHidden(.ColIndex("TotalAdvance")) = True
                End If
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("TotalDiscount")))) = 0 Then
                  .ColHidden(.ColIndex("TotalDiscount")) = True
                  Else
                '  TxtDecrease.Text = val((.TextMatrix(.Rows - 1, .ColIndex("TotalDiscount"))))
                End If
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("Mokafea")))) = 0 Then
                  .ColHidden(.ColIndex("Mokafea")) = True
                End If
Next i
End With
End Sub
Function GetValueAllwIntro(Optional MothID As Integer, Optional YerID As Integer, Optional EmpID As Double, Optional MofrdID As Integer) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     MordValue / ISNULL(TypeMofrd, 1) AS Valu"
sql = sql & " From dbo.TblComponentYearDet"
sql = sql & " WHERE       (EmpID = " & EmpID & ") AND (MofrdID = " & MofrdID & ") and "
sql = sql & "               ((month(RecDate1) =" & MothID & " and Year(RecDate1) =" & YerID & ") or    ((month(RecDate2) =" & MothID & " and Year(RecDate2) =" & YerID & ")))"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetValueAllwIntro = IIf(IsNull(Rs3("Valu").value), 0, Rs3("Valu").value)
Else
GetValueAllwIntro = 0
End If
End Function
Public Sub FillGridWithData()

    Dim i As Integer
    Dim j As Integer
    Dim countFlag As Integer
    Dim AllwIntro As Double
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String
    Dim ColumnName As String
    Dim TotalAddtion As Double
    Dim TotalDiscount As Double
    Dim DaysInMonth22 As Double
    Dim CountDays22 As Double
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

countFlag = 1
 

    IntYear = year(XPDtbTrans.value)
    IntMonth = Month(XPDtbTrans.value)

      Grid.Clear flexClearScrollable, flexClearEverything
              Grid.rows = 1
              
        Dim ID As String
 
    My_SQL = " Select Emp_Namee, lastHolidaydate,BignDateWork,  fullcode,groupid,  BranchId,Emp_ID,Emp_Code,Emp_Name,DepartmentID,project_id ,cost_center_id,IsNUll(Emp_Salary,0)as Emp_Salary,IsNUll(Emp_Salary_sakn,0)as Emp_Salary_sakn,IsNUll(Emp_Salary_bus,0)as Emp_Salary_bus,IsNUll(Emp_Salary_food,0)as Emp_Salary_food,IsNUll(Emp_Salary_others,0)as Emp_Salary_others,IsNUll(Emp_Salary_mob,0)as Emp_Salary_mob,IsNUll(Emp_Salary_mang,0)as Emp_Salary_mang,  IsNUll( TotalDiscount,0)as TotalDiscount,IsNUll(TotalMokafea, 0) As TotalMokafea,(IsNUll(Emp_Salary,0)+IsNUll( TotalMokafea,0))-(IsNUll(TotalDiscount,0)) as EmpTotalNet ,JobTypeName, JobTypeNamee,branch_name,branch_namee,projectFullcode,Project_name,Project_nameE" & CHR(13)
  My_SQL = My_SQL + "  From (" & CHR(13)

  My_SQL = My_SQL + "  SELECT     TOP 100 PERCENT  dbo.TblEmployee.Emp_Namee , dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.BranchId, dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.cost_center_id, SUM(QryAllDiscountWithMkafea.TotalDiscount) AS TotalDiscount, SUM(QryAllDiscountWithMkafea.Mokafea) AS TotalMokafea," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee," & CHR(13)
  My_SQL = My_SQL + "                       dbo.projects.Fullcode AS projectFullcode, dbo.projects.Project_name, dbo.projects.Project_nameE" & CHR(13)
  My_SQL = My_SQL + " FROM         dbo.TblEmpJobsTypes INNER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.QryAllDiscountWithMkafea(" & IntMonth & ", " & IntYear & ") QryAllDiscountWithMkafea ON dbo.TblEmployee.Emp_ID = QryAllDiscountWithMkafea.Emp_ID" & CHR(13)

 
        My_SQL = My_SQL + " and dbo.TblEmployee.BignDateWork<" & SQLDate(XPDtbTrans.value, True)
                If DcboEmpName.text <> "" Then
            My_SQL = My_SQL + " Where  dbo.TblEmployee.Emp_id=" & val(DcboEmpName.BoundText) ' & "'"
        End If

 'DcboEmpName
 My_SQL = My_SQL + "  GROUP BY dbo.TblEmployee.Emp_Namee , dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID, dbo.TblEmployee.BranchId, " & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.cost_center_id, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.projects.Fullcode, dbo.projects.Project_name," & CHR(13)
My_SQL = My_SQL + "                      dbo.Projects.Project_nameE" & CHR(13)
My_SQL = My_SQL + " ORDER BY dbo.TblEmployee.Fullcode" & CHR(13)

My_SQL = My_SQL + "  )XTable"


    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst
Dim CountDays As Double
 
Dim MonthDayNo  As Double

MonthDayNo = daysInMonth(XPDtbTrans.value)

If MonthDayNo = 28 Then
MonthDayNo = 30
ElseIf MonthDayNo = 31 Then
MonthDayNo = 30
End If

            For i = 1 To .rows - 1
         countFlag = 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
            .TextMatrix(i, .ColIndex("BignDateWork")) = IIf(IsNull(rs.Fields("BignDateWork").value), "", rs.Fields("BignDateWork").value)
            .TextMatrix(i, .ColIndex("lastHolidaydate")) = IIf(IsNull(rs.Fields("lastHolidaydate").value), "", rs.Fields("lastHolidaydate").value)

           
           CountDays = day(XPDtbTrans.value)
           
           If MonthDayNo <= CountDays Then
CountDays = 30
 
End If

MonthDayNo = 30
CountDays = val(TxtAbceDay.text)
   CountDays22 = day(XPDtbTrans.value)
   'Abs(DateDiff("D", MonthLastDay(DateSta.value), DateSta.value))
         '  CountDays22 = CountDays22 + 1
           DaysInMonth22 = daysInMonth(XPDtbTrans.value)
           
            .TextMatrix(i, .ColIndex("CountDays")) = CountDays
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), 1, rs.Fields("BranchId").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("fullcode").value), "", rs.Fields("fullcode").value)
                .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs.Fields("cost_center_id").value), "", rs.Fields("cost_center_id").value)
     
                
                      If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs.Fields("JobTypeName").value), "", rs.Fields("JobTypeName").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
           Else
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs.Fields("JobTypeNamee").value), "", rs.Fields("JobTypeNamee").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value)
           End If
                TotalAddtion = 0
                TotalDiscount = 0

                For j = 1 To 40
                    ColumnName = "Comp" & j

                    If ViewComp(j) = True Then
                    AllwIntro = GetValueAllwIntro(Month(XPDtbTrans.value), year(XPDtbTrans.value), val(DcboEmpName.BoundText), j)
                    If AllwIntro <= 0 Then
                        If FixedOrChanged(j) = 0 Then
                            .TextMatrix(i, .ColIndex(ColumnName)) = GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(i, .ColIndex("Emp_ID"))), CStr(j), , XPDtbTrans.value)
                                           
                                           If countFlag = 1 Then
                                           If showMofradAll(j) = False Then
                                            If culc30orRminder(j) = 0 Then
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * CountDays, 2)
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / DaysInMonth22 * CountDays22, 2)
                                          End If
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))), 2)
                                          End If
                                           End If
                                           
                        Else
                            .TextMatrix(i, .ColIndex(ColumnName)) = GetEmployeeChangedSalary(val(.TextMatrix(i, .ColIndex("Emp_ID"))), j, val(CboYear.text), CmbMonth.ListIndex + 1)
                           ' .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))), 2)
                          
                        End If
                       Else
                       .TextMatrix(i, .ColIndex(ColumnName)) = AllwIntro
                       End If
                    End If
    
                Next j
    
                 '         .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Round(rs.Fields("TotalDiscount").value, Decimal_Places))
             
                '.TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Round(rs.Fields("TotalMokafea").value, Decimal_Places))
              
                rs.MoveNext
            
            Next

            rs.Close
        End If

    '  GetAdvanceValues IntMonth, IntYear
        ' GetWorkHours
        CalculateNets
        .rows = .rows + 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        Else
            .TextMatrix(.rows - 1, .ColIndex("Ser")) = "Total"
        End If

        .IsSubtotal(.rows - 1) = True
        Dim SngTotal As Single

        For j = 1 To 40
            ColumnName = "Comp" & j
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex(ColumnName), .rows - 1, .ColIndex(ColumnName))
            .TextMatrix(.rows - 1, .ColIndex(ColumnName)) = SngTotal
     
        Next j
 
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .rows, .ColIndex("total1"))
        .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
        
        If val(TxtAbceDay.text) <> 0 Then
        txtSalary = SngTotal
       Else
       txtSalary = 0
       End If
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .rows, .ColIndex("total2"))
        .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows, .ColIndex("EmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
'

        .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
        .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
        .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
        .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With
 

'rs.Close
Set rs = Nothing

'    Coloring
ErrTrap:

End Sub
Private Sub CalculateNets()
    Dim i As Integer
    Dim SngHourPrice As Single
    Dim SngOverTimePrice As Single

    Dim NetTotal As Single
    Dim SngTemp As Single
    Dim TotalAddtion As Double
    Dim TotalDiscount As Double
    Dim ColumnName As String
    Dim SngTotal As Double
    Dim j As Integer
    'On Error GoTo ErrTrap
    On Error Resume Next

    With Me.Grid

        If .FixedRows = .rows Then Exit Sub

        For i = .FixedRows To .rows - 1

            TotalAddtion = 0
            TotalDiscount = 0

            For j = 1 To 40
                ColumnName = "Comp" & j

                If AddOrDiscount(j) = 0 Then
                    TotalAddtion = TotalAddtion + val(.TextMatrix(i, .ColIndex(ColumnName)))
                Else
                    TotalDiscount = TotalDiscount + val(.TextMatrix(i, .ColIndex(ColumnName)))
                End If

            Next j
        
            .TextMatrix(i, .ColIndex("total1")) = val(.TextMatrix(i, .ColIndex("Mokafea"))) + TotalAddtion
            .TextMatrix(i, .ColIndex("total2")) = val(.TextMatrix(i, .ColIndex("TotalAdvance"))) + val(.TextMatrix(i, .ColIndex("TotalDiscount"))) + TotalDiscount
            .TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2")))

            If i Mod 2 = 0 Then
                .cell(flexcpBackColor, i, 1, i, 41) = &HE0E0E0
     
            End If
        
        Next i
    
    End With

    Exit Sub
ErrTrap:
    'Resume
End Sub
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords



    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
        Msg = "«Ã«“… ⁄«—÷… —ﬁ„ —ﬁ„" & XPTxtID & " ··„ÊŸ› " & DcboEmpName.text
    If check_employee_accounts = False Then
        Exit Function
    End If

 
        
BasicSalaryAccount = ""
 notes_id = general_noteid
                  
    For j = 1 To 40
        ColumnName = "Comp" & j

        If ViewComp(j) = True Then
                                  
            If CheckAccountToJE(Account_code(j)) = False Then
                Account_code(j) = SalaryAccount
            End If
        End If
    
    Next j
        
 
  
     my_branch = val(dcBranch.BoundText)
 
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                 
    Dim CValue As Double
    Dim Branch As Integer
    Dim projectId As Integer
    
       BranchID = val(dcBranch.BoundText)
    If val(txtSalary.text) > 0 Then
                          Employee_account = get_EMPLOYEE_Account(val(DcboEmpName.BoundText), "Account_Code2")
                            If ModAccounts.AddNewDev(LngDevID, line_no, Employee_account, val(txtSalary.text), 0, Msg & "    Õ”«»  „”Õﬁ«  «·«Ã«“…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(my_branch)) = False Then
                                GoTo ErrTrap
                            End If
                              line_no = line_no + 1
     End If
    
    With Grid
BranchID = .TextMatrix(1, .ColIndex("BranchId"))
End With
    With Grid

        For j = 1 To 40

'
            ColumnName = "Comp" & j

            If ViewComp(j) = True And AddOrDiscount(j) = 0 Then '«·ŸÂÊ— Ê«÷«›… Ê·Ì” –„„ Ê·Ì” „ﬁœ„
                       If BasicSalaryAccount = "" Then
                                                                        BasicSalaryAccount = Account_code(j)
                                                 End If
                                                 
                If ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True Then

                        CValue = GetComponentValuePerBranch(BranchID, ColumnName)
                               If val(TxtAbceDay.text) = 0 Then
                               CValue = 0
                        
                                                 
                               End If
                               
                        If CValue > 0 Then
                                            
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 1, Msg & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & TxtAbceDay & " ÌÊ„  " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If

                            line_no = line_no + 1
                        End If
                             
                End If
                             
            End If
    
        Next j
       
                                      


        For i = .FixedRows To .rows - 2
        
                 If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) < 0 And val(val(TxtAbceDay.text)) <> 0 Then         '«·«ÃÊ— «·„” Õﬁ… „œÌ‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Abs(.TextMatrix(i, .ColIndex("EmpTotalNet"))), 1, Msg & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & TxtAbceDay & " ÌÊ„  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
            
'*******************************„œ›Ê⁄«  „ﬁ
            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True And AdvPaymentdAccount(j) = True Then
                     
                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3") 'œ›⁄«  „ﬁœ„…
                    StrAccountCode = Employee_account
                                 If AddOrDiscount(j) = 0 Then
                                                    If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                                            GoTo ErrTrap
                                                        End If
                                
                                                        line_no = line_no + 1
                                                    End If
                        
                        Else
                        
                                
                        
                        
                        End If
                        
                 
                End If

            Next j
                 

            
'*******************************„œ›Ê⁄«  „ﬁ
 
        Next i

    End With
 SystemOptions.ProjectEmployeeGV = False

  If SystemOptions.ProjectEmployeeGV = True Then
'rs.Close
    Dim sql As String
    
    Dim Balance As Double
Dim mofradAccount As String
Dim mofradAccount1 As String
Dim Emp_id As Double
Dim Salary_account As String
 Dim Project_name As String
 Dim mofradname As String
  Dim AddOrDiscount1 As Integer
        sql = "SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.mofrad.Account_Code AS mofradAccount,  dbo.mofrad.Account_Code1 AS mofradAccount1, dbo.TblChangedComponentRegisterDetails.projectid,"
sql = sql & " dbo.Projects.Salary_account , dbo.Projects.Project_name, dbo.MOFRAD.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount"
sql = sql & " FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & "                       dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 0) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(" & SQLDate(NoteDate, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR(" & SQLDate(NoteDate, True) & "))"
sql = sql & " GROUP BY dbo.mofrad.Account_Code,dbo.mofrad.Account_Code1, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount"
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText 'stop

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("mofradAccount").value), "", rs("mofradAccount").value)
     mofradAccount1 = IIf(IsNull(rs("mofradAccount1").value), "", rs("mofradAccount1").value)
     
    'mofradAccount1
     
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), 0, rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & "", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
               
                             
                    Else ' Œ’„

    
                             If SystemOptions.ProjectDiscountPolicy = 1 Then
                             
                                        If mofradAccount1 <> "" Then
                                        Salary_account = mofradAccount1
                                        End If
                            
                             
                             End If
                             
                                line_no = line_no + 1
                         If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
                                
            line_no = line_no + 1

                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
       
'«·„‘«—Ì⁄ Ê·ﬂ‰ –„„
 Dim empAccount_Codezmam As String
 Dim emp_Name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(  " & SQLDate(NoteDate, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(NoteDate, True) & " ))"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code"
 
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText '0000000

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                       If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close

    
    
   ' Õ„Ì· «·„’—Ê›«  ⁄·Ï «·„‘«—Ì⁄
    
       sql = "SELECT      SUM(ROUND(dbo.EmpSalaryComponent.[Value] * dbo.opr_employee_details.[interval] / 30, 2)) AS Total, dbo.mofrad.Account_Code, "
sql = sql & " dbo.mofrad.AddOrDiscount, dbo.EmpSalaryComponent.EntIncresDataM, dbo.projects.Salary_account, 2006 + dbo.opr_Employee.Years AS [year],"
sql = sql & " dbo.opr_Employee.Months, SUM(dbo.opr_employee_details.[interval]) AS Intervals, dbo.opr_employee_details.ProjectID, dbo.mofrdat.mofrad_name,"
sql = sql & " dbo.Projects.Project_name , dbo.TblEmployee.BranchId"
sql = sql & " FROM         dbo.opr_employee_details INNER JOIN"
sql = sql & " dbo.projects ON dbo.opr_employee_details.ProjectID = dbo.projects.id INNER JOIN"
sql = sql & " dbo.opr_Employee ON dbo.opr_employee_details.pk_id = dbo.opr_Employee.id INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.opr_employee_details.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.EmpSalaryComponent ON dbo.opr_employee_details.Emp_id = dbo.EmpSalaryComponent.emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad INNER JOIN"
sql = sql & " dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type ON dbo.EmpSalaryComponent.AccountCode = dbo.mofrdat.mofrad_code"
sql = sql & " GROUP BY dbo.mofrad.Account_Code, dbo.EmpSalaryComponent.EntIncresDataM, dbo.projects.Salary_account, 2006 + dbo.opr_Employee.Years, dbo.opr_Employee.Months,"
sql = sql & " dbo.MOFRAD.AddOrDiscount , dbo.opr_employee_details.ProjectID, dbo.mofrdat.mofrad_name, dbo.Projects.Project_name, dbo.TblEmployee.BranchId"
sql = sql & " HAVING      (dbo.EmpSalaryComponent.EntIncresDataM IS NULL  OR"
sql = sql & "  dbo.EmpSalaryComponent.EntIncresDataM >= " & SQLDate(NoteDate, True) & " )"

sql = sql & "   AND (dbo.opr_Employee.Months = " & CmbMonth.ListIndex & ") AND (2006 + dbo.opr_Employee.Years = " & val(CboYear.text) & ")"


sql = sql & " ORDER BY dbo.opr_employee_details.ProjectID"

 
   
  
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
             projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
    
    
    
    
    
    
    
'«·„‘«—Ì⁄ Ê·ﬂ‰ œ›⁄«  „ﬁœ„…
 'Dim empAccount_Codezmam As String
 'Dim emp_name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code3"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.AdvPaymentdAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(   " & SQLDate(NoteDate, True) & "  )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(NoteDate, True) & "  ))"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code3"
 
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code3").value), "", rs("Account_Code3").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
        
                        line_no = line_no + 1
                        
                
                             
                    Else ' Œ’„
                    
                        If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1

                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    

End If
       
       sql = " "

    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
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

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, " ’—ÌÕ Œ—ÊÃ „ƒﬁ /«” ∆–«‰/«Ã«“… ⁄«—÷…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’—ÌÕ Œ—ÊÃ „ƒﬁ /«” ∆–«‰/«Ã«“… ⁄«—÷…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’—ÌÕ Œ—ÊÃ „ƒﬁ /«” ∆–«‰/«Ã«“… ⁄«—÷…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’—ÌÕ Œ—ÊÃ „ƒﬁ /«” ∆–«‰/«Ã«“… ⁄«—÷…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’—ÌÕ Œ—ÊÃ „ƒﬁ /«” ∆–«‰/«Ã«“… ⁄«—÷…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’—ÌÕ Œ—ÊÃ „ƒﬁ /«” ∆–«‰/«Ã«“… ⁄«—÷…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, " ’—ÌÕ Œ—ÊÃ „ƒﬁ /«” ∆–«‰/«Ã«“… ⁄«—÷…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’—ÌÕ Œ—ÊÃ „ƒﬁ /«” ∆–«‰/«Ã«“… ⁄«—÷…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’—ÌÕ Œ—ÊÃ „ƒﬁ /«” ∆–«‰/«Ã«“… ⁄«—÷…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’—ÌÕ Œ—ÊÃ „ƒﬁ /«” ∆–«‰/«Ã«“… ⁄«—÷…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ’—ÌÕ Œ—ÊÃ „ƒﬁ /«” ∆–«‰/«Ã«“… ⁄«—÷…", 1, 15204351, -2147483630
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

Private Sub TxtInterval_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtInterval.text, 0)
End Sub

