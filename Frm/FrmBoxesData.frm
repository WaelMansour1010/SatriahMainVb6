VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBoxesData 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»Ì«‰«  «·Œ“‰   Ê «·⁄Âœ"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "FrmBoxesData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPriod 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   51
      Top             =   3360
      Width           =   885
   End
   Begin VB.ComboBox DcbType 
      Height          =   315
      ItemData        =   "FrmBoxesData.frx":038A
      Left            =   4440
      List            =   "FrmBoxesData.frx":038C
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   5160
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·⁄Âœ…"
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   2640
      Width           =   5655
      Begin VB.TextBox txtboxValue 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   200
         Width           =   1185
      End
      Begin VB.OptionButton BTOpt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄Âœ… „” œÌ„…"
         Height          =   195
         Index           =   1
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton BTOpt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄Âœ… „ƒﬁ …"
         Height          =   195
         Index           =   0
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«ﬁ’Ì ﬁÌ„…"
         Height          =   315
         Index           =   10
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·—’Ìœ «·√›  «ÕÏ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1305
      Index           =   1
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   5040
      Width           =   3105
      Begin VB.TextBox txtopening_balance_voucher_id 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œÌ‰"
         Height          =   255
         Index           =   0
         Left            =   2190
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   210
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "œ«∆‰"
         Height          =   255
         Index           =   1
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   210
         Width           =   765
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "€Ì— „Õœœ"
         Height          =   255
         Index           =   2
         Left            =   330
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   210
         Width           =   1005
      End
      Begin VB.TextBox TxtOpenBalance 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   510
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker Dtp 
         Height          =   330
         Left            =   360
         TabIndex        =   39
         Top             =   870
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   12648447
         CustomFormat    =   "yyyy/M/d"
         Format          =   93519875
         CurrentDate     =   38718
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· ”ÃÌ·"
         Height          =   285
         Index           =   9
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·—’Ìœ "
         Height          =   255
         Index           =   8
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   540
         Width           =   1275
      End
   End
   Begin VB.CheckBox chkChequeBox 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " › Õ Õ«›Ÿ… ‘Ìﬂ« "
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox XPTxtBoxNamee 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   180
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1920
      Width           =   4065
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄ÂœÂ"
      Height          =   195
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2400
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Œ“Ì‰…"
      Height          =   195
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2400
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox XPMTxtRemark 
      Alignment       =   1  'Right Justify
      Height          =   435
      Left            =   150
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4140
      Width           =   4065
   End
   Begin VB.TextBox XPTxtBoxName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   180
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1485
      Width           =   4065
   End
   Begin VB.TextBox XPTxtBoxID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   750
      Width           =   1185
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6315
      _cx             =   11139
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
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "«·Œ“‰  Ê«·⁄Âœ"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   1
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
      CaptionStyle    =   1
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   -180
         Visible         =   0   'False
         Width           =   855
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   0
         Left            =   1155
         TabIndex        =   2
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
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
         ButtonImage     =   "FrmBoxesData.frx":038E
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
         Height          =   345
         Index           =   2
         Left            =   90
         TabIndex        =   3
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
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
         ButtonImage     =   "FrmBoxesData.frx":0728
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
         Height          =   345
         Index           =   1
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
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
         ButtonImage     =   "FrmBoxesData.frx":0AC2
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
         Height          =   345
         Index           =   3
         Left            =   615
         TabIndex        =   5
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
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
         ButtonImage     =   "FrmBoxesData.frx":0E5C
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
         Left            =   2160
         Picture         =   "FrmBoxesData.frx":11F6
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   9
      Top             =   6930
      Width           =   585
      _ExtentX        =   1032
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
      Left            =   4890
      TabIndex        =   10
      Top             =   6930
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   4155
      TabIndex        =   11
      Top             =   6930
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   3405
      TabIndex        =   12
      Top             =   6930
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   2745
      TabIndex        =   13
      Top             =   6930
      Width           =   585
      _ExtentX        =   1032
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
      Left            =   840
      TabIndex        =   14
      Top             =   6930
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   2070
      TabIndex        =   15
      Top             =   6450
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   2160
      TabIndex        =   23
      Top             =   6930
      Width           =   585
      _ExtentX        =   1032
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
   Begin MSDataListLib.DataCombo DcEmp 
      Height          =   315
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   180
      TabIndex        =   29
      Top             =   1080
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   0
      TabIndex        =   46
      Top             =   6930
      Width           =   795
      _ExtentX        =   1402
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
      Index           =   7
      Left            =   1560
      TabIndex        =   49
      Top             =   6930
      Width           =   585
      _ExtentX        =   1032
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
   Begin MSDataListLib.DataCombo DboParentAccount 
      Height          =   315
      Left            =   120
      TabIndex        =   54
      Top             =   4680
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Õ”«» «·—∆Ì”Ì"
      Height          =   315
      Index           =   33
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   4680
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÌÊ„"
      Height          =   315
      Index           =   11
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«ﬁ’Ï „œ… ·· ’›ÌÂ"
      Height          =   285
      Index           =   21
      Left            =   2160
      TabIndex        =   52
      Top             =   3360
      Width           =   1605
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ «‰Ã·Ì“Ì"
      Height          =   315
      Index           =   7
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4845
      TabIndex        =   30
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ÊŸ›"
      Height          =   315
      Index           =   6
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·‰Ê⁄"
      Height          =   315
      Index           =   5
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   6480
      Width           =   825
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   6480
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ "
      Height          =   285
      Index           =   0
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   765
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   315
      Index           =   1
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4260
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   2
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6480
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ ⁄—»Ì"
      Height          =   315
      Index           =   3
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1485
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   4
      Left            =   780
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   6480
      Width           =   1155
   End
End
Attribute VB_Name = "FrmBoxesData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Sub FillParntAcount()
Dim Account_Code_dynamic As String
        If Option1.value = True Then
            Account_Code_dynamic = get_account_code_branch(6, my_branch)
         End If
            If Option2.value = True Then
            Account_Code_dynamic = get_account_code_branch(35, my_branch)
         End If
         Me.DboParentAccount.BoundText = Account_Code_dynamic
End Sub
Private Sub Cmd_Click(Index As Integer)
'    On Error GoTo ErrTrap
    Dim FirstPeriodDateInthisYear As Date
    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

    Me.Dtp.value = FirstPeriodDateInthisYear

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            
            '        XPTxtBoxName.SetFocus
            Option1.Enabled = True
            Option2.Enabled = True
            Option1.value = True
 
            Me.DcBranch.BoundText = Current_branch
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

            Me.Dtp.value = FirstPeriodDateInthisYear
            OptType(2).value = True
        BTOpt(0).Enabled = True
      FillParntAcount
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            Option1.Enabled = False
            Option2.Enabled = False

        Case 2

            If SystemOptions.ChequeBox = True And Option1.value = True Then
                chkChequeBox.value = vbChecked
            Else
                chkChequeBox.value = vbUnchecked
            End If
''///////////////
        Dim Account_Code_dynamic As String
        If Option1.value = True Then
            Account_Code_dynamic = get_account_code_branch(6, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
               Exit Sub
                
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·’‰«œÌﬁ   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
       Exit Sub
                End If
                
            End If
         End If
            If Option2.value = True Then
            Account_Code_dynamic = get_account_code_branch(35, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
               Exit Sub
                
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·⁄Âœ   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        Exit Sub
                End If
               
            End If
         End If
        '    DboParentAccount.BoundText = Account_Code_dynamic
         ''''''////////////
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Company

        Case 5
                
                FrmExpensesSearch.RetrunType = 21
                FrmExpensesSearch.Indx = 2
                FrmExpensesSearch.Caption = Me.Caption
                FrmExpensesSearch.show
                
        Case 6
            Unload Me
         Case 7
         print_report2
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAttach_Click()
            On Error Resume Next
ShowAttachments XPTxtBoxID, "0701201405"
 

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches DcBranch
    End If

End Sub

Private Sub DCEmP_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetEmployees Me.DCEmp
    End If


   If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 23
        Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If
    
    
    
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Option1_Click()
If Option1.value = True Then
    Frame1.Visible = False
Else
    Frame1.Visible = True
End If
FillParntAcount
End Sub

Private Sub Option2_Click()
If Option2.value = True Then
    Frame1.Visible = True
Else
    Frame1.Visible = False
End If
FillParntAcount
End Sub

Private Sub OptType_Click(Index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.Text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.Text)
End Sub

Private Sub txtboxValue_KeyPress(KeyAscii As Integer)
   KeyAscii = KeyAscii_Num(KeyAscii, Me.txtboxValue.Text, 0)
End Sub

Private Sub TxtOpenBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOpenBalance.Text, 0)
End Sub

Private Sub Form_Activate()
    XPTxtBoxID.SetFocus
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

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos
    
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DCEmp
    If SystemOptions.UserInterface = EnglishInterface Then
     DcbType.AddItem "Day"
     DcbType.AddItem "Month"
     DcbType.AddItem "Year"
     Else
     DcbType.AddItem "ÌÊ„"
     DcbType.AddItem "‘Â—"
     DcbType.AddItem "”‰Â"
     End If

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  »Ì«‰«  «·Œ“‰   Ê «·⁄Âœ  "
    LogTexte = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
    'My_SQL = "  select branch_id,branch_name from TblBranchesData  order by branch_id  "
    'fill_combo DcBranch, My_SQL
  
    Dcombos.GetBranches DcBranch
  Dcombos.GetAccountingCodes Me.DboParentAccount, False, True
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.DcBranch.Enabled = True
    End If

    If SystemOptions.ChequeBox = False Then
        Me.chkChequeBox.Visible = False
  
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
'    rs.Open "tblBoxesData", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
   Dim StrSQL As String
     If SystemOptions.usertype <> UserAdminAll Then
      
StrSQL = "SELECT  *  From tblBoxesData    where BranchId=" & Current_branch
  Else
 StrSQL = "SELECT  *  From tblBoxesData"
    End If
    
    
  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
        
        
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

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

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
  lbl(33).Caption = "Parent Acc"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(5).Caption = "Type"
    lbl(11).Caption = "Days"
    Option1.Caption = "Box"
    Option2.Caption = "Cash On Hand"
    chkChequeBox.Caption = "Have Cheque Box"
lbl(21).Caption = "maximum liquidation"
Frame1.Caption = "Type"
BTOpt(0).Caption = "Temp"
BTOpt(1).Caption = "Perm."
lbl(10).Caption = "Max. Val."
CmdAttach.Caption = "Attachments"
Cmd(7).Caption = "Print"
    Me.Fra(1).Caption = "Open Balance State"
    OptType(0).Caption = "Debit"
    OptType(1).Caption = "Credit"
    OptType(2).Caption = "Un Sign"
    lbl(8).Caption = "Balance Value"
    lbl(9).Caption = "Rec Date"

    Me.Caption = "Cash On Hand"
    EleHeader.Caption = Me.Caption
    lbl(6).Caption = "Employee"
    lbl(0).Caption = "Box Code"
    Label3.Caption = "Branch"
    lbl(3).Caption = " Name Ar"
    lbl(7).Caption = " Name En"

    lbl(1).Caption = "Remarks"
    lbl(2).Caption = "Current Record"
    lbl(4).Caption = "NO. Recordes"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    'Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «·Œ“‰   Ê «·⁄Âœ  "
    LogTexte = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·Œ“‰   Ê«·⁄Âœ"
            Else
                Me.Caption = "Boxes Data"
            End If
       DboParentAccount.Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = True
            Me.XPMTxtRemark.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

        Case "N"
            DboParentAccount.Enabled = True
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·Œ“‰   Ê«·⁄Âœ( ÃœÌœ )"
            Else
                Me.Caption = "Boxes Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·Œ“‰   Ê«·⁄Âœ( ÃœÌœ )"
            Else
                Me.Caption = "Boxes Data(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            '       Me.XPBtnMove(0).Enabled = False
            '       Me.XPBtnMove(1).Enabled = False
            '       Me.XPBtnMove(2).Enabled = False
            '       Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = False
            Me.XPMTxtRemark.locked = False

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·Œ“‰   Ê«·⁄Âœ(  ⁄œÌ· )"
            Else
                Me.Caption = "Boxes Data(Edit)"
            End If
           DboParentAccount.Enabled = False
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = False
            Me.XPMTxtRemark.locked = False
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = "SELECT     dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.Comments, dbo.TblBoxesData.Type, dbo.TblBoxesData.Account_Code, "
MySQL = MySQL & "                      dbo.TblBoxesData.empid, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblBoxesData.BranchId, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblBoxesData.BoxNameE, dbo.TblBoxesData.Account_Code1, dbo.TblBoxesData.OpenBalanceDate,"
MySQL = MySQL & "                      dbo.TblBoxesData.OpenBalanceType, dbo.TblBoxesData.OpenBalance, dbo.TblBoxesData.boxValue, dbo.TblBoxesData.Account_Code2, dbo.TblBoxesData.BTtype,"
MySQL = MySQL & "                      dbo.TblBoxesData.Driverid , dbo.TblBoxesData.opening_balance_voucher_id, dbo.TblBoxesData.ChequeBox, dbo.TblBoxesData.ParentAccount"
MySQL = MySQL & " FROM         dbo.TblBoxesData LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblBoxesData.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblBoxesData.empid = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & " Where (dbo.TblBoxesData.BoxID =" & val(XPTxtBoxID.Text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBoxesData.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBoxesData.rpt"
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
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function
Public Sub Retrive(Optional Lngid As Long = 0)
Dim Account_Code_dynamic As String
    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    If rs("type").value = 0 Then
        Option1.value = True
        Frame1.Visible = False
    Else

        Option2.value = True
        Frame1.Visible = True
    End If
        Dim i As Integer
    If Lngid <> 0 Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount

            If rs("BoxID").value = Lngid Then
                GoTo ll
            End If

            rs.MoveNext
        Next i

        Exit Sub
    End If
ll:
  If Option1.value = True Then
            Account_Code_dynamic = get_account_code_branch(6, my_branch)
         End If
            If Option2.value = True Then
            Account_Code_dynamic = get_account_code_branch(35, my_branch)
         End If
         
            DboParentAccount.BoundText = Account_Code_dynamic
   Me.DboParentAccount.BoundText = IIf(IsNull(rs("parent_account").value), Account_Code_dynamic, rs("parent_account").value)
    DcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)

    XPTxtBoxID.Text = IIf(IsNull(rs("BoxID").value), "", val(rs("BoxID").value))

    XPTxtBoxName.Text = IIf(IsNull(rs("BoxName").value), "", Trim(rs("BoxName").value))
    XPTxtBoxNamee.Text = IIf(IsNull(rs("BoxNameE").value), "", Trim(rs("BoxNameE").value))

    XPMTxtRemark.Text = IIf(IsNull(rs("Comments").value), "", Trim(rs("Comments").value))
    DCEmp.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
''//
 DcbType.ListIndex = IIf(IsNull(rs("PriodDMY").value), -1, rs("PriodDMY").value)
 Me.TxtPriod.Text = IIf(IsNull(rs("Priod")), "", Trim(rs("Priod")))
 ''//
Me.txtboxValue.Text = IIf(IsNull(rs("boxValue")), "", Trim(rs("boxValue")))


 
 If Not IsNull(rs("BTtype").value) Then
            If rs("BTtype").value = 0 Then
                BTOpt(0).value = True
            Else
        
                BTOpt(1).value = True
            End If
 Else
         BTOpt(0).value = True
 End If
 
    
    
 
    If rs("ChequeBox").value = False Then
        Me.chkChequeBox.value = vbUnchecked
    ElseIf rs("ChequeBox").value = True Then
        Me.chkChequeBox.value = vbChecked
    End If
    
    Dim FirstPeriodDateInthisYear As Date
    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

    txtopening_balance_voucher_id.Text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)

    If Not (IsNull(rs("OpenBalanceDate").value)) Then
        Me.Dtp.value = rs("OpenBalanceDate").value
        ' Me.Dtp.Enabled = True
    Else
        
        Me.Dtp.value = FirstPeriodDateInthisYear
        '  Me.Dtp.Enabled = False
    End If
    
    If Not IsNull(rs("OpenBalanceType").value) Then
        Me.TxtOpenBalance.Text = IIf(IsNull(rs("OpenBalance")), "", Trim(rs("OpenBalance")))

        If rs("OpenBalanceType").value = 0 Then
            OptType(0).value = True
            OptType_Click 0
        ElseIf rs("OpenBalanceType").value = 1 Then
            OptType(1).value = True
            OptType_Click 1
        End If
        
    Else
        Me.TxtOpenBalance.Text = 0
        Me.OptType(2).value = True
        OptType_Click 2
    End If
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

'Private Sub TxtTypeValu_KeyPress(KeyAscii As Integer)
'KeyAscii = KeyAscii_Num(KeyAscii, TxtTypeValu.text, 1)
'End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
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

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«   «·Œ“‰ Ê «·⁄Âœ " & CHR(13) & " ﬂÊœ «·Œ“Ì‰…/«·⁄ÂœÂ  " & XPTxtBoxID.Text & CHR(13) & " «·›—⁄ " & DcBranch.Text & CHR(13) & "«·«”„ ⁄—»Ì  " & XPTxtBoxName & CHR(13) & " «·‰Ê⁄     "

    If Option1.value = True Then
        LogTextA = LogTextA & " Œ“Ì‰…  "
    ElseIf Option2.value = True Then
        LogTextA = LogTextA & "  ⁄Âœ…  "
  
    End If

    LogTextA = LogTextA & CHR(13) & "› Õ Õ«›Ÿ… «·‘Ìﬂ«  "

    If chkChequeBox.value = vbChecked Then
        LogTextA = LogTextA & "‰⁄„ "
    Else
        LogTextA = LogTextA & "·« "
    End If

    LogTextA = LogTextA & CHR(13) & "«”„ «·„ÊŸ›   " & DCEmp.Text
                    
    LogTextA = LogTextA & CHR(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "

    If OptType(0).value = True Then
        LogTextA = LogTextA & "„œÌ‰"
    ElseIf OptType(1).value = True Then
        LogTextA = LogTextA & "œ«∆‰"
    ElseIf OptType(2).value = True Then
        LogTextA = LogTextA & "€Ì— „Õœœ"
    End If

    LogTextA = LogTextA & CHR(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ  " & TxtOpenBalance
    LogTextA = LogTextA & CHR(13) & "„·«ÕŸ«    " & XPMTxtRemark
      
    'sssssssssssssssssss
    LogTexte = "  Save Screen " & " Boxes Data " & CHR(13) & " Code " & XPTxtBoxID.Text & CHR(13) & " Branch " & DcBranch.Text & CHR(13) & "«Name  " & XPTxtBoxName & CHR(13) & " Type     "

    If Option1.value = True Then
        LogTexte = LogTexte & " Box  "
    ElseIf Option2.value = True Then
        LogTexte = LogTexte & "  Era  "
  
    End If

    LogTexte = LogTexte & CHR(13) & "Open CHeque Box "

    If chkChequeBox.value = vbChecked Then
        LogTexte = LogTexte & "Yes "
    Else
        LogTexte = LogTexte & "No "
    End If

    LogTexte = LogTexte & CHR(13) & " Employee Name" & DCEmp.Text
                    
    LogTexte = LogTexte & CHR(13) & "Opening Balance Type"

    If OptType(0).value = True Then
        LogTexte = LogTexte & "Debit"
    ElseIf OptType(1).value = True Then
        LogTexte = LogTexte & "Credit"
    ElseIf OptType(2).value = True Then
        LogTexte = LogTexte & "Na"
    End If

    LogTexte = LogTexte & CHR(13) & " Opening Balance  Value " & TxtOpenBalance
    LogTexte = LogTexte & CHR(13) & " Remarks   " & XPMTxtRemark
      
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", ""
    End If

End Function
 
Private Sub SaveData()
    Dim Account_Code_dynamic As String
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
    
      '  If Trim(DcBranch.BoundText) = "" Then
      '      If SystemOptions.UserInterface = EnglishInterface Then
      '          Msg = "Specify Departement"
      '      Else
      '          Msg = "Õœœ «·›—⁄ «Ê·« "
      '      End If

      '      MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      '      DcBranch.SetFocus
      '      SendKeys "{F4}"
      '      Screen.MousePointer = vbDefault
      '      Exit Sub
      '  End If
    
        If XPTxtBoxName.Text = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·Œ“‰… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtBoxName.SetFocus
            Exit Sub
        End If

        Select Case Me.TxtModFlg.Text

            Case "N"
                StrSQL = "select * From  tblBoxesData where BoxName='" & Trim(XPTxtBoxName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    Msg = "Â‰«ﬂ Œ“‰… „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·Œ“‰…"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtBoxName.SetFocus
                    Exit Sub
                End If

            Case "E"
                StrSQL = "select * From  tblBoxesData where BoxName='" & Trim(XPTxtBoxName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("BoxID").value <> val(XPTxtBoxID.Text) Then
                        Msg = "Â‰«ﬂ Œ“‰…  „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·Œ“‰…"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        XPTxtBoxName.SetFocus
                        Exit Sub
                    End If
                End If

        End Select

        Select Case Me.TxtModFlg.Text

            Case "N"

                

              '  If Option1.value = True Then
              '      Account_Code_dynamic = get_account_code_branch(6, my_branch)
              '
              '      If Account_Code_dynamic = "NO branch" Then
              '          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
              '          GoTo ErrTrap
              '      Else
'
'                        If Account_Code_dynamic = "NO account" Then
'                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··’‰«œÌﬁ   ›Ì «·›—⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                            GoTo ErrTrap
'
'                        End If
'                    End If
'
'                Else
'
'                    Account_Code_dynamic = get_account_code_branch(35, my_branch)
'
'                    If Account_Code_dynamic = "NO branch" Then
'                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'                        GoTo ErrTrap
'                    Else
'
'                        If Account_Code_dynamic = "NO account" Then
'                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··⁄Âœ   ›Ì «·›—⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                            GoTo ErrTrap
'
'                        End If
'                    End If
'
'                End If
        
                rs.AddNew
                        XPTxtBoxID.Text = CStr(new_id("tblBoxesData", "BoxID", "", True))
        End Select



        Cn.BeginTrans
        BeginTrans = True
        rs("BranchId").value = IIf(Me.DcBranch.BoundText = "", 0, val(DcBranch.BoundText))
        Account_Code_dynamic = Me.DboParentAccount.BoundText
        rs("parent_account").value = IIf(Me.DboParentAccount.BoundText = "", Null, (Me.DboParentAccount.BoundText))
        rs("BoxID").value = val(XPTxtBoxID.Text)
        rs("BoxName").value = Trim(XPTxtBoxName.Text)
        rs("BoxNamee").value = Trim(XPTxtBoxNamee.Text)
        rs("OpenBalanceDate").value = Me.Dtp.value
        ''///
       rs("Priod").value = val(Me.TxtPriod.Text)
       rs("PriodDMY").value = IIf(DcbType.ListIndex = -1, Null, val(DcbType.ListIndex))
       ''//
       rs("boxValue").value = val(Me.txtboxValue.Text)
        rs("EmpID").value = IIf(DCEmp.BoundText = "", Null, val(DCEmp.BoundText))
    
        rs("Comments").value = IIf(XPMTxtRemark.Text = "", Null, Trim(XPMTxtRemark.Text))

        If Option1.value = True Then
            rs("type").value = 0
        Else
            rs("type").value = 1
        End If


        If BTOpt(0).value = True Then
            rs("BTtype").value = 0
        Else
            rs("BTtype").value = 1
        End If
        
        
        Dim ParentAccount As String

        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            If Me.TxtModFlg.Text = "N" Then

                If SystemOptions.ChequeBox = False And SystemOptions.BoxLossandIncreae = False Then
        
                    rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtBoxName.Text), True, False, XPTxtBoxNamee.Text)
                Else
                
                    If chkChequeBox.value = vbChecked Then
                        ParentAccount = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtBoxName.Text, False, False, XPTxtBoxNamee.Text)
                        rs("ParentAccount").value = ParentAccount
                        rs("ChequeBox").value = 1
                        rs("Account_Code").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtBoxName.Text), True, False, XPTxtBoxNamee.Text)
                        rs("Account_Code1").value = ModAccounts.AddNewAccount(ParentAccount, " Õ«›Ÿ… ‘Ìﬂ«   " & Trim$(Me.XPTxtBoxName.Text), True, False, XPTxtBoxNamee.Text & "  Cheque Box")
                    
                                If SystemOptions.BoxLossandIncreae = True Then
                                  rs("Account_Code2").value = ModAccounts.AddNewAccount(ParentAccount, " ⁄Ã“ Ê“Ì«œ… «·‰ﬁœÌ…-    " & Trim$(Me.XPTxtBoxName.Text), True, False, XPTxtBoxNamee.Text & "  Loss And Increase")

                                End If
                    
                    
                    Else
                    
                          If SystemOptions.BoxLossandIncreae = True Then
                              ParentAccount = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtBoxName.Text, False, False, XPTxtBoxNamee.Text)
                             rs("ParentAccount").value = ParentAccount
                          rs("Account_Code").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtBoxName.Text), True, False, XPTxtBoxNamee.Text)

                       rs("Account_Code2").value = ModAccounts.AddNewAccount(ParentAccount, " ⁄Ã“ Ê“Ì«œ… «·‰ﬁœÌ…-    " & Trim$(Me.XPTxtBoxName.Text), True, False, XPTxtBoxNamee.Text & "  Loss And Increase")

                         Else
                          
                        'rs("Account_Code").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtBoxName.text), True, False, XPTxtBoxNamee.text)
                        rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtBoxName.Text), True, False, XPTxtBoxNamee.Text)
                        
                        rs("ParentAccount").value = Null
                        rs("ChequeBox").value = 0
                        
                          End If
                                

                    End If
             
                End If
         
                'Rs("Account_Code").value = ModAccounts.AddNewAccount("a1a2a1", Trim$(Me.XPTxtBoxName.text), True, False)
            Else
        
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
          
                If rs("ChequeBox").value = 0 Then
                    If Not IsNull(rs("Account_Code").value) Then
                        ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtBoxName.Text, XPTxtBoxNamee.Text, , , , , , , , , , , , , , , , , True
                    End If
            
                Else
          
                    If Not IsNull(rs("ParentAccount").value) Then
                        ModAccounts.EditAccount rs("ParentAccount").value, Me.XPTxtBoxName.Text, Trim(XPTxtBoxNamee.Text), , , , , , , , , , , , , , , , , False
                    End If
            
                    If Not IsNull(rs("Account_Code").value) Then
                        ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtBoxName.Text, XPTxtBoxNamee.Text, , , , , , , , , , , , , , , , , True
                    End If
            
                    If Not IsNull(rs("Account_Code1").value) Then
                        ModAccounts.EditAccount rs("Account_Code1").value, "  Õ«›Ÿ… ‘Ìﬂ«   " & Me.XPTxtBoxName.Text, XPTxtBoxNamee.Text & "  Cheque Box ", , , , , , , , , , , , , , , , , True
                    End If
          
                    If Not IsNull(rs("Account_Code2").value) Then
                        ModAccounts.EditAccount rs("Account_Code2").value, "  ⁄Ã“ Ê“Ì«œ… «·‰ﬁœÌ…- " & Me.XPTxtBoxName.Text, XPTxtBoxNamee.Text & "  Loss and Increase   ", , , , , , , , , , , , , , , , , True
                    End If
                    
                End If
        
            End If
        End If
    
        If val(TxtOpenBalance.Text) = 0 Then
            txtopening_balance_voucher_id = 0
        End If
       
        
        If Me.OptType(2).value = True Then
            rs("OpenBalance").value = 0
            rs("OpenBalanceType").value = Null
        ElseIf Me.OptType(0).value = True Then
            rs("OpenBalance").value = val(Me.TxtOpenBalance.Text)
            rs("OpenBalanceType").value = 0
        ElseIf Me.OptType(1).value = True Then
            rs("OpenBalance").value = val(Me.TxtOpenBalance.Text)
            rs("OpenBalanceType").value = 1
        End If

If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
       
         '   If val(Me.txtopening_balance_voucher_id.text) = 0 Then
                txtopening_balance_voucher_id.Text = get_opening_balance_voucher_id
            
         '   End If '
        End If '

        rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.Text)

        rs.update
    
        Dim StrDes As String

        If SystemOptions.UserInterface = ArabicInterface Then
            StrDes = "«·—’Ìœ «·≈›  «ÕÏ ·‹ "
        Else
            StrDes = " Opening Balance For: "
        End If
        
        If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
                Dim LngDevID As Long
                Dim LngOpenID As Long
                Dim Account_Code_dynamic1 As String
             
                'LngOpenID = ModAccounts.AddNewOpenBalance(Val(Me.XPTxtCusID.text), Me.Dtp.value)
                ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
            
                If Me.OptType(0).value = True Then
                   
                    Account_Code_dynamic1 = get_account_code_branch(57, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
            
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code").value, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtBoxName.Text) & "  " & Trim$(Me.XPTxtBoxNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtBoxName.Text) & "  " & Trim$(Me.XPTxtBoxNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
                    '  If ModAccounts.AddNewDev(LngDevID, 2, "a2a1a1", _
                       Val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    '         GoTo ErrTrap
                    ' End If
                ElseIf Me.OptType(1).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(57, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtBoxName.Text) & "  " & Trim$(Me.XPTxtBoxNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    ' If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
                      Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    
                    '       GoTo ErrTrap
                    'End If
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code").value, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtBoxName.Text) & "  " & Trim$(Me.XPTxtBoxNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(DcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If

                '   update_account_opening_balance rs("Account_Code").value
                'update_account_opening_balance Account_Code_dynamic1
                 
            End If
        End If

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·Œ“‰…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
            rs.find "BoxID='" & val(XPTxtBoxID.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode2 As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If XPTxtBoxID.Text <> "" Then

        If Not IsNull(rs("Driverid").value) Then
            Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·Œ“‰…" & CHR(13)
            Msg = Msg + "·«‰Â« „‰‘∆… «·Ì« „‰ „·› «·”«∆ﬁÌ‰"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    
        StrAccountCode = rs("Account_Code").value
        StrAccountCode2 = IIf(IsNull(rs("Account_Code2").value), "", rs("Account_Code2").value)
        ParentAccount = IIf(IsNull(rs("ParentAccount").value), "", rs("ParentAccount").value) '  rs("ParentAccount").value
        
   
        StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where Account_Code='" & StrAccountCode & "'"

        If Not IsNull(rs("Account_Code1").value) Then
            StrSQL = StrSQL & " or   Account_Code='" & rs("Account_Code1").value & "'"
        End If
    
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·Œ“‰…" & CHR(13)
            Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «·Œ“‰…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

        Msg = "”Ì „ Õ–› »Ì«‰«  «·Œ“‰… —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBoxID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            DeleteOpeningBalance
    
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            
                If chkChequeBox.value = vbChecked Then
                    StrAccountCode1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
                    ParentAccount = IIf(IsNull(rs("ParentAccount").value), "", rs("ParentAccount").value) '  rs("ParentAccount").value
             If SystemOptions.BoxLossandIncreae = True And Option2.value = True Then
                                                            
                                                            If ModAccounts.DeleteAccount(StrAccountCode2, True) = True And ModAccounts.DeleteAccount(ParentAccount, True) = True Then
                                                            
                                                            End If
                                                
                                                
                                                End If
                                                
                                    If ModAccounts.DeleteAccount(StrAccountCode, True) = True And ModAccounts.DeleteAccount(StrAccountCode1, True) = True And ModAccounts.DeleteAccount(StrAccountCode2, True) = True And ModAccounts.DeleteAccount(ParentAccount, True) = True Then
                                        rs.delete
                                        Msg = " „  ⁄„·Ì… «·Õ–›."
                                        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            
                                    Else
                                        GoTo ErrTrap
                                    End If

                Else




                                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                                
                                                If SystemOptions.BoxLossandIncreae = True And Option2.value = True Then
                                                            ParentAccount = IIf(IsNull(rs("ParentAccount").value), "", rs("ParentAccount").value) '  rs("ParentAccount").value
                                                            If ModAccounts.DeleteAccount(StrAccountCode2, True) = True And ModAccounts.DeleteAccount(ParentAccount, True) = True Then
                                                            
                                                            End If
                                                
                                                
                                                End If
                                
                                    CuurentLogdata ("D")
                                    rs.delete
                                Else
                                    Exit Sub
                                End If
                                
 
  
  
                End If
            
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
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·Œ“‰… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub

Function DeleteOpeningBalance()
    Cmd_Click (1)
    OptType(2).value = True
    TxtOpenBalance.Text = 0
    Cmd_Click (2)

End Function

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  Œ“‰… ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·Œ“‰… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtBoxName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtBoxNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
