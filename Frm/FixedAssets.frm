VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FixedAssets 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·«’Ê· «·À«» …"
   ClientHeight    =   9480
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   11445
   Icon            =   "FixedAssets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   11445
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   165
      Top             =   2280
      Width           =   1245
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2340
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   163
      Top             =   2280
      Width           =   1245
   End
   Begin VB.CheckBox chkIsContainer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«ÊÌ« "
      Height          =   195
      Left            =   4260
      RightToLeft     =   -1  'True
      TabIndex        =   160
      Top             =   2280
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.OptionButton RdMove 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " €Ì— „‰ﬁÊ· «·„œ…"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   1
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   159
      Top             =   2280
      Width           =   1455
   End
   Begin VB.OptionButton RdMove 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰ﬁÊ· «·„œ…"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   0
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   158
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox TxtYearNotMove 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   5400
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox TxtYearMove 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   8580
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   2310
      Width           =   1575
   End
   Begin VB.TextBox TotalQest 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   345
      Left            =   5400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   152
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox TxtQstRemNo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   150
      Top             =   4080
      Width           =   3525
   End
   Begin VB.TextBox TxtQstAdNo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   5400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   148
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox TxtBookValue 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   8550
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   146
      Top             =   4440
      Width           =   1605
   End
   Begin VB.TextBox TxtDisposedValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   8550
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   144
      Top             =   4080
      Width           =   1605
   End
   Begin VB.TextBox TxtAddValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   8550
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   142
      Top             =   3720
      Width           =   1605
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   " ÕœÌÀ"
      Height          =   255
      Index           =   0
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   137
      ToolTipText     =   "Ì„ﬂ‰  ÕœÌÀ  «—ÌŒ «·«” ·«„ Ê «—ÌŒ »œ«Ì… «·«Â·«ﬂ »‘—ÿ ⁄œ„ ÊÃÊœ «ﬁ”«ÿ ”«»ﬁ… ⁄·Ï «·«’·"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CheckBox ISEQUP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„—ﬂ»… / „⁄œ…"
      Height          =   315
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   134
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3030
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   130
      Top             =   1560
      Width           =   1245
   End
   Begin VB.TextBox TxtNameE 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   4755
   End
   Begin VB.TextBox TxtMinusValue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   109
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   375
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtopening_balance_voucher_id 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtNoteID 
      Height          =   285
      Left            =   12480
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Height          =   435
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   675
      Width           =   2175
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÃœÌœ"
         Height          =   195
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   120
         Width           =   915
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "«›  «ÕÌ"
         Height          =   195
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox txtPurchaseBillId 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2640
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   840
      Width           =   1245
   End
   Begin VB.TextBox TxtKhordaPrice 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3000
      Width           =   3525
   End
   Begin VB.TextBox TxtCurrentValue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   8550
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   3360
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Caption         =   "»Ì«‰«  «·«Â·«ﬂ"
      Height          =   2415
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   10080
      Width           =   8535
   End
   Begin VB.TextBox txtinstallDo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   3720
      Width           =   3525
   End
   Begin VB.TextBox txtinstallmentresult 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   4440
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.ComboBox cStatus 
      Height          =   315
      ItemData        =   "FixedAssets.frx":000C
      Left            =   240
      List            =   "FixedAssets.frx":001C
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   1560
      Width           =   1845
   End
   Begin VB.ComboBox CBoDepreciation_Type_id 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FixedAssets.frx":005C
      Left            =   8580
      List            =   "FixedAssets.frx":0066
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox TxtnoOfInst 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   5400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtinstallValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   3360
      Width           =   3525
   End
   Begin VB.TextBox TxtAccDepreciation 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   5400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox XPTxtID 
      Height          =   285
      Left            =   6960
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   9960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtPurchasePrice 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   8550
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3000
      Width           =   1605
   End
   Begin VB.TextBox TxtName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1215
      Width           =   4755
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9030
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   840
      Width           =   1125
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   10320
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   675
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   11475
      _cx             =   20241
      _cy             =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   20.25
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
      Caption         =   "«·«’Ê· «·À«» …"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
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
      Begin VB.CommandButton cmdReSave 
         Caption         =   "÷»ÿ «·Õ—ﬂ« "
         Height          =   330
         Left            =   3450
         TabIndex        =   162
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2580
         PasswordChar    =   "*"
         TabIndex        =   161
         Top             =   165
         Width           =   750
      End
      Begin VB.TextBox Txtfullcode 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   5160
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   141
         Top             =   240
         Width           =   1605
      End
      Begin VB.TextBox BiLLID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   128
         Text            =   "Text1"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1155
         TabIndex        =   15
         Top             =   120
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
         ButtonImage     =   "FixedAssets.frx":0089
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
         Left            =   90
         TabIndex        =   16
         Top             =   120
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
         ButtonImage     =   "FixedAssets.frx":0423
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
         Left            =   1680
         TabIndex        =   17
         Top             =   120
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
         ButtonImage     =   "FixedAssets.frx":07BD
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
         Left            =   615
         TabIndex        =   18
         Top             =   120
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
         ButtonImage     =   "FixedAssets.frx":0B57
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   8670
      TabIndex        =   19
      Top             =   8355
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   7800
      TabIndex        =   20
      Top             =   8355
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   6915
      TabIndex        =   21
      Top             =   8355
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   6045
      TabIndex        =   22
      Top             =   8355
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   3960
      TabIndex        =   23
      Top             =   8355
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   930
      TabIndex        =   24
      Top             =   8355
      Width           =   795
      _ExtentX        =   1402
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
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   345
      Left            =   9960
      TabIndex        =   35
      Top             =   10560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   274464769
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   120
      TabIndex        =   36
      Top             =   9120
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker dpStartDepreciationDate 
      Height          =   345
      Left            =   5400
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      CalendarBackColor=   12640511
      CalendarForeColor=   255
      CalendarTitleBackColor=   -2147483635
      CalendarTitleForeColor=   255
      Format          =   274464769
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   240
      TabIndex        =   46
      Top             =   1200
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DPReceiveDate 
      Height          =   345
      Left            =   5400
      TabIndex        =   48
      Top             =   1920
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      _Version        =   393216
      Format          =   274464769
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   51
      Top             =   8355
      Width           =   1035
      _ExtentX        =   1826
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
      Left            =   7800
      TabIndex        =   52
      Top             =   8760
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«Ìﬁ«› «·«Â·«ﬂ"
      BackColor       =   14871017
      Enabled         =   0   'False
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
      Index           =   8
      Left            =   6240
      TabIndex        =   53
      Top             =   8760
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "≈⁄«œ…  ‘€Ì· «·«Â·«ﬂ"
      BackColor       =   14871017
      Enabled         =   0   'False
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
      Index           =   9
      Left            =   4680
      TabIndex        =   54
      Top             =   8760
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«· Œ·’ „‰ «·«’·"
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
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   7560
      TabIndex        =   55
      Top             =   840
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   315
      Index           =   10
      Left            =   1560
      TabIndex        =   66
      Top             =   840
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      ButtonPositionImage=   1
      Caption         =   "⁄—÷ «·›« Ê—…"
      BackColor       =   14871017
      ForeColor       =   16711680
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
      ColorToggledText=   16711680
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSComCtl2.DTPicker DpPurchaseDate 
      Height          =   345
      Left            =   5400
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      Format          =   274464769
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker DPLastDepreciationDate 
      Height          =   345
      Left            =   2400
      TabIndex        =   68
      Top             =   2640
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      Format          =   274464769
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   11
      Left            =   0
      TabIndex        =   72
      Top             =   10635
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
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
   Begin MSDataListLib.DataCombo DCGroup 
      Height          =   315
      Left            =   7680
      TabIndex        =   4
      Top             =   1920
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcEmployee 
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   12
      Left            =   1920
      TabIndex        =   76
      Top             =   8355
      Width           =   795
      _ExtentX        =   1402
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   3480
      Left            =   0
      TabIndex        =   78
      Top             =   4800
      Width           =   11490
      _cx             =   20267
      _cy             =   6138
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
      ForeColor       =   0
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "»Ì«‰«  «”«”Ì…|»Ì«‰«   Õ·Ì·Ì…|»Ì«‰«  «·«÷«›…"
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
      Picture(0)      =   "FixedAssets.frx":0EF1
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì‰«  «·«÷«›…"
         Height          =   3015
         Left            =   12435
         RightToLeft     =   -1  'True
         TabIndex        =   138
         Top             =   45
         Width           =   11400
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   2595
            Left            =   120
            TabIndex        =   139
            Top             =   240
            Width           =   11205
            _cx             =   19764
            _cy             =   4577
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FixedAssets.frx":128B
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
            Begin VB.ComboBox CboType 
               BackColor       =   &H00C0E0FF&
               Height          =   315
               ItemData        =   "FixedAssets.frx":136F
               Left            =   0
               List            =   "FixedAssets.frx":1371
               RightToLeft     =   -1  'True
               TabIndex        =   140
               Text            =   "CboType"
               Top             =   600
               Visible         =   0   'False
               Width           =   3855
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3015
         Left            =   12135
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   45
         Width           =   11400
         _cx             =   20108
         _cy             =   5318
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
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «Œ—Ï  ›’Ì·Ì…"
            Height          =   3255
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   120
            Width           =   6000
            Begin VB.TextBox TxtOprNo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3480
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   960
               Width           =   1245
            End
            Begin VB.TextBox txtBoardNo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3480
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   240
               Width           =   1245
            End
            Begin VB.TextBox txtSerialNo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1200
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   240
               Width           =   1245
            End
            Begin VB.TextBox txtChaseeNo 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3480
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   600
               Width           =   1245
            End
            Begin VB.TextBox txtModel 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1200
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   600
               Width           =   1245
            End
            Begin MSDataListLib.DataCombo dcContryid 
               Height          =   315
               Left            =   720
               TabIndex        =   116
               Top             =   1440
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcVendorid 
               Height          =   315
               Left            =   720
               TabIndex        =   117
               Top             =   1800
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker dpEndLicense 
               Height          =   345
               Left            =   3360
               TabIndex        =   118
               Top             =   2280
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   609
               _Version        =   393216
               Format          =   274530305
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker dpEndTest 
               Height          =   345
               Left            =   720
               TabIndex        =   119
               Top             =   2280
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   609
               _Version        =   393216
               Format          =   274530305
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—ﬁ„ «· ‘€Ì·Ì"
               Height          =   315
               Index           =   14
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   960
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "  «‰ Â«¡ «·«” „«—…"
               Height          =   315
               Index           =   10
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   2280
               Width           =   1275
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "  «‰ Â«¡ «·›Õ’"
               Height          =   315
               Index           =   8
               Left            =   4800
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   2280
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Ê—œ"
               Height          =   315
               Index           =   7
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   1920
               Width           =   1035
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»·œ «·„‰‘√"
               Height          =   315
               Index           =   6
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   1560
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «··ÊÕ…"
               Height          =   315
               Index           =   1
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   240
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·«” „«—…"
               Height          =   315
               Index           =   2
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   240
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ÊœÌ·"
               Height          =   315
               Index           =   3
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   600
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·‘«”Ì…"
               Height          =   315
               Index           =   4
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   600
               Width           =   915
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   3015
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   45
         Width           =   11400
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            Height          =   2895
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   120
            Width           =   4695
            Begin VB.TextBox txtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   2400
               Visible         =   0   'False
               Width           =   3045
            End
            Begin VB.TextBox TxtSalePrice 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   1680
               Width           =   3045
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   2040
               Width           =   3045
            End
            Begin VB.TextBox TxtNotes 
               Alignment       =   1  'Right Justify
               Height          =   675
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   100
               Top             =   840
               Width           =   4485
            End
            Begin MSDataListLib.DataCombo DcCostCenter 
               Height          =   315
               Left            =   120
               TabIndex        =   132
               Top             =   240
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„—ﬂ“ «· ﬂ·›…"
               Height          =   195
               Index           =   13
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   240
               Width           =   1125
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„ﬂ”» «Ê «·Œ”«—…"
               Height          =   375
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   2040
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   195
               Index           =   0
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   2400
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "”⁄— «·»Ì⁄"
               Height          =   255
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ê’› «·«’·"
               Height          =   195
               Index           =   124
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   600
               Width           =   1125
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  „Ã„Ê⁄Â «·«’·"
            Enabled         =   0   'False
            Height          =   2895
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   120
            Width           =   6495
            Begin VB.TextBox TXtPercentage1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2760
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   360
               Width           =   1245
            End
            Begin VB.TextBox txtPercentage2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2760
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   720
               Width           =   1245
            End
            Begin VB.TextBox TXT40 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   2520
               Width           =   3885
            End
            Begin VB.TextBox TXT31 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   2160
               Width           =   3885
            End
            Begin VB.TextBox TXT25 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   1800
               Width           =   3885
            End
            Begin VB.TextBox TXT26 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   1440
               Width           =   3885
            End
            Begin VB.TextBox TXT24 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   1080
               Width           =   3885
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·Â «Â·«ﬂ"
               Enabled         =   0   'False
               Height          =   225
               Index           =   0
               Left            =   2370
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   120
               Width           =   1815
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·Ì” ·Â «Â·«ﬂ"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   120
               Value           =   -1  'True
               Width           =   2295
            End
            Begin VB.TextBox TxtAge 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   360
               Width           =   525
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»… «·«Â·«ﬂ ⁄‰œ «·«Ìﬁ«›"
               Height          =   255
               Index           =   110
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   720
               Width           =   1995
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»… «·«Â·«ﬂ"
               Height          =   255
               Index           =   109
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   360
               Width           =   1995
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ”«»   Œ”«∆— »Ì⁄"
               Height          =   255
               Index           =   115
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   2520
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ”«»   «—»«Õ »Ì⁄"
               Height          =   255
               Index           =   114
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   2160
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ”«»    „’—Ê›«  «·«Â·«ﬂ"
               Height          =   255
               Index           =   113
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   1800
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ”«» „Ã„⁄ «·«Â·«ﬂ"
               Height          =   255
               Index           =   112
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   1440
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ”«» «·«’·  »«·„Ì“«‰Ì…"
               Height          =   255
               Index           =   111
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   1080
               Width           =   1995
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„— «·«› —«÷Ì ··«’· »«·‘Â—"
               Height          =   255
               Index           =   9
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   360
               Width           =   2115
            End
         End
      End
   End
   Begin VB.TextBox txtNoteSerial1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtNoteID1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   13
      Left            =   2880
      TabIndex        =   108
      Top             =   8355
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "‰”Œ… „„«À·Â"
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
      Caption         =   "«·«Ã„«·Ì"
      Height          =   315
      Index           =   22
      Left            =   1530
      RightToLeft     =   -1  'True
      TabIndex        =   166
      Top             =   2280
      Width           =   585
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”⁄—"
      Height          =   315
      Index           =   21
      Left            =   3630
      RightToLeft     =   -1  'True
      TabIndex        =   164
      Top             =   2280
      Width           =   585
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Ã„Ê⁄Â"
      Height          =   315
      Index           =   103
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   155
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»⁄ÂœÂ"
      Height          =   315
      Index           =   104
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   154
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈Ã„«·Ì «·«ﬁ”«ÿ"
      Height          =   255
      Index           =   20
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   153
      Top             =   4440
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„ »ﬁÌ «·«ﬁ”«ÿ"
      Height          =   255
      Index           =   19
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   151
      Top             =   4080
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ «·«ﬁ”«ÿ «·„÷«›…"
      Height          =   255
      Index           =   18
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   149
      Top             =   4080
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌ„… «·œ› —Ì… «·Õ«·Ì…"
      Height          =   315
      Index           =   17
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   147
      Top             =   4440
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«” »⁄«œ« "
      Height          =   315
      Index           =   16
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   145
      Top             =   4080
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«÷«›« "
      Height          =   315
      Index           =   15
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   143
      Top             =   3720
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·⁄œœ"
      Height          =   315
      Index           =   12
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   131
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ «‰Ã·Ì“Ì"
      Height          =   315
      Index           =   11
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   129
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ  «Œ— «Â·«ﬂ"
      Height          =   375
      Index           =   120
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·‘—«¡"
      Height          =   255
      Index           =   128
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   3000
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ›« Ê—… «·‘—«¡"
      Height          =   255
      Index           =   116
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… «·«’· ﬂŒ—œ…"
      Height          =   375
      Index           =   121
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… ‘—« «·«’·"
      Height          =   315
      Index           =   106
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   3000
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«ﬁ”«ÿ «·„‰›–…"
      Height          =   255
      Index           =   130
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„ »ﬁÏ"
      Height          =   255
      Index           =   123
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   4440
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ… «·«Â·«ﬂ"
      Height          =   315
      Index           =   105
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   2640
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·«” ·«„"
      Height          =   375
      Index           =   119
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   315
      Index           =   117
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   1200
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ «ﬁ”«ÿ   «·«Â·«ﬂ"
      Height          =   255
      Index           =   108
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   3720
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… ﬁ”ÿ «·«Â·«ﬂ"
      Height          =   255
      Index           =   122
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã„⁄ «·«Â·«ﬂ"
      Height          =   255
      Index           =   129
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   3360
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ »œ«Ì… «·«Â·«ﬂ*"
      Height          =   255
      Index           =   127
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   2640
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·«’·"
      Height          =   255
      Index           =   118
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   315
      Index           =   5
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   9120
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   375
      Left            =   8280
      TabIndex        =   33
      Top             =   10080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LngDevID 
      Height          =   375
      Left            =   6960
      TabIndex        =   32
      Top             =   10080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌ„… «·œ› —Ì…"
      Height          =   315
      Index           =   107
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   3360
      Width           =   1395
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   2010
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   8880
      Width           =   465
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   8880
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   126
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   8880
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   125
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   8880
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ ⁄—»Ì"
      Height          =   315
      Index           =   102
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1215
      Width           =   1395
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·«’·"
      Height          =   315
      Index           =   101
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   840
      Width           =   1395
   End
End
Attribute VB_Name = "FixedAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RSAss As New ADODB.Recordset
Dim FirstPeriodDateInthisYear  As Date
Dim TTP As clstooltip
Dim ScreenNameArabic As String
Dim ScreenNameEnglish As String
Dim mIsFinishSave As Boolean
Dim IsSaveWithOutMsg As Boolean
Dim mIsStart As Boolean

Private Sub Cmd_Click(index As Integer)
    Dim msgstr  As String
'    On Error GoTo ErrTrap
    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
    XPDtbTrans.value = FirstPeriodDateInthisYear

    Select Case index

        Case 0
             If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            
            TxtModFlg.text = "N"
            clear_all Me
 
            Me.DCboUserName.BoundText = user_id
            Option1.value = True
     
            CBoDepreciation_Type_id.ListIndex = 0
            cStatus.ListIndex = 0
             txtopening_balance_voucher_id.text = 0 ' get_opening_balance_voucher_id
            Frame3.Enabled = True
            Me.dcBranch.BoundText = branch_id
'Option1.SetFocus
        Case 13
            TxtModFlg.text = "N"
            txtID.text = ""
 
            Me.DCboUserName.BoundText = user_id
            Option1.value = True
            CBoDepreciation_Type_id.ListIndex = 0
            cStatus.ListIndex = 0
           txtopening_balance_voucher_id.text = 0 '  get_opening_balance_voucher_id
            Frame3.Enabled = True
            Me.dcBranch.BoundText = branch_id
     
        Case 1
   If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            
            If Option1.value = True Then
                If IsSaveWithOutMsg Then Exit Sub
                If checkEneringPurchaseInvoices = False Then
                    'Exit Sub
                End If
            End If

            Dim noOfInstallments  As Integer
            noOfInstallments = CheCkInstallmentCount(val(Me.XPTxtID.text))

            If noOfInstallments > 0 Then
            
            
                If IsSaveWithOutMsg Then Exit Sub
                If SystemOptions.UserInterface = ArabicInterface Then
                    msgstr = " ·« Ì„ﬂ‰ «· ⁄œÌ·  „  ‰›Ì– «ﬁ”«ÿ ⁄·Ï «·«’·  " & CHR(13)
                    msgstr = msgstr & TxtName.text & CHR(13)
                    msgstr = msgstr & "⁄œœ «·«ﬁ”«ÿ «·„‰›–… Õ Ï «·«‰ " & noOfInstallments
                    MsgBox msgstr, vbCritical
                Else
                    msgstr = " Can't Edit Fixed Asset   " & CHR(13)
                    msgstr = msgstr & TxtName.text & CHR(13)
                    msgstr = msgstr & "No Of Executed Installments " & noOfInstallments
                    MsgBox msgstr, vbCritical
                End If

                Exit Sub
            End If

            TxtModFlg.text = "E"
            If val(DCboUserName.BoundText) = 0 Then
             Me.DCboUserName.BoundText = user_id
            End If
            If cStatus.ListIndex = 0 Then 'Ã«— «·«Â·«ﬂ
       
                Cmd(7).Enabled = True ' «Ìﬁ«› «·«Â·«ﬂ
                Cmd(9).Enabled = True ' «· Œ·’ „‰ «·«’·
            ElseIf cStatus.ListIndex = 1 Then '
       
                Cmd(8).Enabled = True '«⁄«œ…  ‘€Ì· «·«Â·«ﬂ
                Cmd(9).Enabled = True ' «· Œ·’ „‰ «·«’·
            End If
         
         '   Me.Dcbranch.BoundText = my_branch
        Frame3.Enabled = False
            CuurentLogdata

        Case 2
    
            'If DCPreFix.text = "" Then
            'MsgBox "Õœœ «·Ã“¡ «·À«Ì "
            'DCPreFix.SetFocus
            'SendKeys "{F4}"

            'Exit Sub
            'End If
           
            If val(TxtAccDepreciation.text) > val(TxtPurchasePrice.text) Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    msgstr = " ·« Ì„ﬂ‰  «‰ ÌﬂÊ‰ „Ã„⁄ «·«Â·«ﬂ «ﬂ»— „‰  ﬁÌ„… ‘—«¡ «·«’·  " & CHR(13)
'
'                    MsgBox msgstr, vbCritical
'                Else
'                    msgstr = " Error : TxtPurchasePrice < TxtAccDepreciation   " & CHR(13)
'
'                    MsgBox msgstr, vbCritical
'                End If

                Exit Sub
            End If

            Dim currentcode As String

            If txtID.text = "" Then
                currentcode = get_coding(branch_id, "FixedAssets", 1, Me.DCPreFix.text)

                If currentcode = "miniError" Then
                    MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
                    Exit Sub
            
                ElseIf currentcode = "Manual" Then
                    MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·ÕﬁÊ·"
                Else
                    txtID = currentcode
                End If
            End If

            SaveData

        Case 3
            Call Undo

        Case 4
             If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
            Del_AssetType

        Case 5
         '   VIEW_ATTACH
                On Error Resume Next
                      If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments DCPreFix.text & txtID.text, "0701201403"
 


        Case 6
            Unload Me

        Case 7 ' «Ìﬁ«› «·«Â·«ﬂ
            cStatus.ListIndex = 1
            Cmd(7).Enabled = False

        Case 8 ' «⁄«œ…  ‘€Ì· «·«Â·«ﬂ
            cStatus.ListIndex = 0
            Cmd(8).Enabled = False

        Case 9 ' «· Œ·’ „‰ «·«’·
    
            cStatus.ListIndex = 3
            Cmd(9).Enabled = False
    
        Case 10
            FrmExpenses4.show

            If val(Me.BiLLID.text) = 0 Then
                FrmExpenses4.Retrive -1
            Else
                FrmExpenses4.Retrive val(Me.BiLLID.text)
            End If
    
        Case 11
            ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 12
           If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
            
            FixedAssetsSearch.RetrunType = 0
            FixedAssetsSearch.show vbModal
       
    End Select

    Exit Sub
ErrTrap:
End Sub

Function VIEW_ATTACH()
    'On Error Resume Next
 
    'If TxtEmp_Code.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „ÊŸ› «Ê·«": Exit Sub

    imaged.show
    imaged.Label9.Caption = "„—›ﬁ«  «·«’· —ﬁ„"
    imaged.Caption = "„—›ﬁ«  «·«’·  "
    imaged.txtopeation_type = "„—›ﬁ«  «·«’·"
    imaged.SUBJECT_NO = 1 'TxtEmp_Code.text
    imaged.Label6.Caption = "ﬂÊœ «·«’·"
    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  «·«’·' and subject_no='" & "1" & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Function

Private Sub DcboBox_Change()

End Sub

Private Sub cmdReSave_Click()
 

   Dim s As String
     
    Dim i As Double
     XPBtnMove_Click (2)
    DoEvents
    
    
For i = 1 To rs.RecordCount
  IsSaveWithOutMsg = True
 Cmd_Click (1)
 DoEvents
'  NewGrid.DtpBillDate_Change
  DoEvents
  DoEvents
  
  IsSaveWithOutMsg = True
  DoEvents
             Cmd_Click (2)
         DoEvents
 
        
            XPBtnMove_Click (0)
       '   Cmd_Click (0)
    
Next i
                 
 
  
    IsSaveWithOutMsg = False
    MsgBox " „ «·Õ›Ÿ"


End Sub

Private Sub cmdUpdate_Click(index As Integer)
Dim msgstr As String

  '          Dim noOfInstallments  As Integer
  '          noOfInstallments = CheCkInstallmentCount(val(Me.XPTxtID.text))
'
'            If noOfInstallments > 0 Then
'                If SystemOptions.UserInterface = ArabicInterface Then
''                    msgstr = " ·« Ì„ﬂ‰ «· ⁄œÌ·  „  ‰›Ì– «ﬁ”«ÿ ⁄·Ï «·«’·  " & Chr(13)
 '                   msgstr = msgstr & TxtName.text & Chr(13)
 '                   msgstr = msgstr & "⁄œœ «·«ﬁ”«ÿ «·„‰›–… Õ Ï «·«‰ " & noOfInstallments
 '                   MsgBox msgstr, vbCritical
 '               Else
 '                   msgstr = " Can't Edit Fixed Asset   " & Chr(13)
'                    msgstr = msgstr & TxtName.text & Chr(13)
'                    msgstr = msgstr & "No Of Executed Installments " & noOfInstallments
'                    MsgBox msgstr, vbCritical
'                End If
'
'                Exit Sub
'            End If
'
 Dim StrSQL As String
  StrSQL = "    update  dbo.FixedAssets set Status_id=" & cStatus.ListIndex & "  , Name= '" & TxtName & "',Namee='" & TxtName & "', ReceiveDate=" & SQLDate(DPReceiveDate.value, True) & " , StartDepreciationDate=" & SQLDate(dpStartDepreciationDate.value, True) & "Where id=" & val(Me.XPTxtID.text)
  Cn.Execute StrSQL
  rs.Resync
  MsgBox " „ «· ÕœÌÀ"
  
End Sub

Private Sub cStatus_Click()

    If cStatus.Enabled = False Then Exit Sub
    If Me.TxtModFlg.text = "R" Then Exit Sub
    TxtSalePrice.Enabled = False
    Dim Msg As String
    Dim msge As String

    If cStatus.ListIndex > -1 Then
        If cStatus.ListIndex = 0 = True Then   'ÃœÌœ  Ê «›  «ÕÌ ÊÃ«—Ì «·«Â·«ﬂ
    
        ElseIf cStatus.ListIndex = 1 Then  'ÃœÌœ Ê «›  «ÕÌ  Ê„ Êﬁ›
            Msg = "·« Ì„ﬂ‰  ⁄œÌ· «·Õ«·… «·Ï „ Êﬁ› ·«‰Â« Õ—ﬂ… «·Ì…"
            cStatus.ListIndex = -1
        ElseIf cStatus.ListIndex = 2 And Option1.value = True Then 'ÃœÌœ  „ «·«Â·«ﬂ
            Msg = "·« Ì„ﬂ‰  «œŒ«· «’· ÃœÌœ  „ «Â·«ﬂÂ „„ﬂ‰ –·ﬂ »√Œ Ì«— «›  «ÕÌ"
            cStatus.ListIndex = -1
        ElseIf cStatus.ListIndex = 3 And Option1.value = True Then 'ÃœÌœ  „ «· Œ·’
            Msg = "·« Ì„ﬂ‰  «œŒ«· «’· ÃœÌœ  „ «· Œ·’ „‰… „„ﬂ‰ –·ﬂ »√Œ Ì«— «›  «ÕÌ    "
            cStatus.ListIndex = -1
        ElseIf cStatus.ListIndex = 2 And Option2.value = True Then '«›  «ÕÌ  „ «·«Â·«ﬂ
        ElseIf cStatus.ListIndex = 3 And Option2.value = True Then '«›  «ÕÌ  „ «· Œ·’
            TxtSalePrice.Enabled = True
        End If
    End If

    If Msg <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox Msg, vbCritical
        Else

        End If
    End If

End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 12
    End If
End Sub

Private Sub DCEmployee_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 2511
        FrmEmployeeSearch.show
  
    End If
End Sub

Private Sub DCGroup_Click(Area As Integer)
'1-  ⁄‰œ «Œ Ì«— «·„Ã„Ê⁄Â »ÌÃÌ» Õ”«»« Â« Ê»ÌÃÌ» «·‰”»

 
 
 
    If val(Me.dcBranch.BoundText) = 0 And Me.TxtModFlg <> "R" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical
        Else
            MsgBox "Select Branch Firstly    ", vbCritical
        End If

        dcBranch.SetFocus
        Sendkeys "{F4}"
    End If
 
    On Error Resume Next
'«·”ÿ— œÂ »ÌÃÌ» ﬂÊœ «·„Ã„Ê⁄Â
    If val(DCGroup.BoundText) = 0 Then Exit Sub
    Me.DCPreFix.text = GetPrefix(val(DCGroup.BoundText), "FixedAssetsGroup")

    If Len(Me.DCPreFix.text) > 1 Then
'        Me.DCPreFix.text = Mid(Me.DCPreFix.text, 2, Len(Me.DCPreFix.text))
    End If
 
    Dim AccountName As String
    Dim Percentage1 As Integer
    Dim Percentage2 As Integer
    Dim DepType As Integer
    Dim Account_code As String
    Dim Account_code1 As String
    Dim Account_code2 As String
    Dim Account_code3 As String
    Dim Account_code4 As String
'Â‰« »ÌÃÌ» Õ”«»«  «·„Ã„Ê⁄Â Ê‰”» «·«Â·«ﬂ Ê»ÌÕ”» ⁄„—  «·«’· »«·‘Â—
    GetFixedAssetsGroupAccount val(DCGroup.BoundText), , val(Me.dcBranch.BoundText), , , Percentage1, Percentage2, DepType, Account_code, Account_code1, Account_code2, Account_code3, Account_code4
 'Â‰« »ÌÃÌ» Õ”«»«  «·„Ã„Ê⁄Â
    TXT24.text = Get_Account_name(, Account_code)
    TXT26.text = Get_Account_name(, Account_code2)
    TXT25.text = Get_Account_name(, Account_code1)
    TXT31.text = Get_Account_name(, Account_code3)
    TXT40.text = Get_Account_name(, Account_code4)
    TXtPercentage1.text = Percentage1
    txtPercentage2.text = Percentage2
  
  TxtAge = Round(100 / val(TXtPercentage1) * 12, 0)
  
    ' GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 26, Val(Me.dcBranch.BoundText), , AccountName  '„’—Ê›«  «·«Â·«ﬂ
    '   TXT26.text = AccountName
    ' GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 25, Val(Me.dcBranch.BoundText), , AccountName '„Ã„⁄ «·«Â·«ﬂ
    '   TXT25.text = AccountName
    '  GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 31, Val(Me.dcBranch.BoundText), , AccountName '«—»«Õ »Ì⁄
    '    TXT31.text = AccountName
    '  GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 40, Val(Me.dcBranch.BoundText), , AccountName  'Œ”«∆— »Ì⁄
    '    TXT40.text = AccountName
    ' GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 40, Val(Me.dcBranch.BoundText), , , Percentage1   '  ‰”»… «·«Â·«ﬂ
    ' TXtPercentage1.text = Percentage1
    ' GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 40, Val(Me.dcBranch.BoundText), , , , Percentage2 '  ‰”»… «·«Â·«ﬂ ⁄‰œ «·«Ìﬁ«›
    ' txtPercentage2.text = Percentage2
    ' GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 40, Val(Me.dcBranch.BoundText), , , , , DepType '       ·Â «Â·«ﬂ 1 Ê·Ì” ·Â 0
 If TXT26.text = "" Then
 TxtAccDepreciation.text = 0
 TxtAccDepreciation.Enabled = False
 
 End If
 
    If DepType = 1 Then ' Â· «·«’· ·Â «Â·«ﬂ
        opt(0).value = True ' ·Â «Â·«ﬂ
        CBoDepreciation_Type_id.Enabled = True
    Else
        opt(1).value = True ' ·Ì” ·Â «Â·«ﬂ
        CBoDepreciation_Type_id.Enabled = False
    End If
cStatus.Enabled = True
End Sub

Private Sub Form_Activate()
    'XPTxtID.SetFocus
End Sub

Function GetAddValue(Optional FixedID As Integer = 0) As Double
If FixedID <> 0 Then
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = sql & " SELECT     SUM(AddValue) AS SmAddValue, FixedID"
sql = sql & " From dbo.TblAdditionsAssest"
sql = sql & "  Where (FixedID = " & FixedID & ")"
sql = sql & " GROUP BY FixedID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GetAddValue = IIf(IsNull(Rs8("SmAddValue").value), 0, Rs8("SmAddValue").value)
Else
GetAddValue = 0
End If
End If
End Function
Function GetQstAddNo(Optional FixedID As Integer = 0) As Double
If FixedID <> 0 Then
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = sql & " SELECT     SUM(QstIncNo) AS SmQst, FixedID"
sql = sql & " From dbo.TblAdditionsAssest"
sql = sql & "  Where (FixedID = " & FixedID & ")"
sql = sql & " GROUP BY FixedID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GetQstAddNo = IIf(IsNull(Rs8("SmQst").value), 0, Rs8("SmQst").value)
Else
GetQstAddNo = 0
End If
End If
End Function

 Sub FullGridData(Optional ID As Double = 0)
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 1
sql = "SELECT     Distrbute, QstNo, QstValue, AddValue, SatrtDate, DateAdd, ID, FixedID, TypeSand"
sql = sql & " From dbo.TblAdditionsAssest"
sql = sql & " Where (FixedID = " & ID & ") "
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("DateAdd")) = IIf(IsNull(Rs1("DateAdd").value), "", Rs1("DateAdd").value)
                   .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs1("ID").value), 0, Rs1("ID").value)
                   .TextMatrix(i, .ColIndex("AddValue")) = IIf(IsNull(Rs1("AddValue").value), 0, Rs1("AddValue").value)
              If Not (IsNull(Rs1("TypeSand").value)) Then
                CboType.ListIndex = IIf(IsNull(Rs1("TypeSand").value), -1, Rs1("TypeSand").value)
              .TextMatrix(i, .ColIndex("TypeSand")) = CboType.text
              End If
              Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ «·«’· " & DCPreFix & txtID.text & CHR(13) & " «”„  «·«’·   " & TxtName & CHR(13) & "   «·„Ã„Ê⁄Â   " & DCGroup & CHR(13) & "      «·›—⁄   " & dcBranch & CHR(13) & " Õ«·… «·«’· " & cStatus

    If Option1.value = True Then
        LogTextA = LogTextA & CHR(13) & "      ÃœÌœ     "
    ElseIf Option2.value = True Then
        LogTextA = LogTextA & CHR(13) & "   «›  «ÕÌ  "
                
    End If
                    
     LogTextA = LogTextA & CHR(13) & "   ÿ—Ìﬁ… «·«Â·«ﬂ   " & CBoDepreciation_Type_id & CHR(13) & "    «—ÌŒ »œ«Ì…  «·«Â·«ﬂ   " & dpStartDepreciationDate & CHR(13) & "    «—ÌŒ «Œ—  «·«Â·«ﬂ   " & DPLastDepreciationDate & CHR(13) & "        ﬁÌ„… ‘—«¡ «·«’·    " & TxtPurchasePrice & CHR(13) & "         «—ÌŒ ‘—«¡ «·«’·   " & DpPurchaseDate & CHR(13) & "        ﬁÌ„… «·«’· ﬂŒ—œ…   " & TxtKhordaPrice & CHR(13) & " «·ﬁÌ„… «·œ› —Ì…  " & txtCurrentValue & CHR(13) & "  „Ã„⁄ «·«Â·«ﬂ   " & TxtAccDepreciation & CHR(13) & "     ﬁÌ„… ﬁ”ÿ  «·«Â·«ﬂ   " & txtinstallValue & CHR(13) & "          «ﬁ”«ÿ «·«Â·«ﬂ  «·„‰›–…  " & txtinstallDo & CHR(13) & "        ⁄œœ «ﬁ”«ÿ «·«Â·«ﬂ  «·„ »ﬁÌ… " & txtinstallmentresult & CHR(13) & " «·⁄„— «·«› —«÷Ì ··«’· »«·‘Â—" & TxtAge & CHR(13) & "  Ê’› «·«’· " & TxtNotes
       LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "F.A. Code   " & DCPreFix & txtID.text & CHR(13) & " F.A. Name " & TxtName & CHR(13) & "   Group   " & DCGroup & CHR(13) & "      Branch   " & dcBranch & CHR(13) & "    Status " & cStatus

    If Option1.value = True Then
        LogTexte = LogTextA & CHR(13) & "      New     "
    ElseIf Option2.value = True Then
        LogTexte = LogTextA & CHR(13) & "   Opening  "
                
    End If
                    
     LogTexte = LogTexte & CHR(13) & " Depreciation Type  " & CBoDepreciation_Type_id & CHR(13) & "  Start Depreciation Date    " & dpStartDepreciationDate & CHR(13) & "  LastDepreciationDate   " & DPLastDepreciationDate & CHR(13) & "   PurchasePrice    " & TxtPurchasePrice & CHR(13) & " PurchaseDate" & DpPurchaseDate & CHR(13) & "       Khorda Price  " & TxtKhordaPrice & CHR(13) & " CurrentValue" & txtCurrentValue & CHR(13) & "  Acc. Depreciation   " & TxtAccDepreciation & CHR(13) & "    intinstallment Value   " & txtinstallValue & CHR(13) & "  installment Done " & txtinstallDo & CHR(13) & "      Remain installment " & txtinstallmentresult & CHR(13) & " Age Range By Month " & TxtAge & CHR(13) & " Remarks " & TxtNotes
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , val(TxtNoteSerial)
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , val(TxtNoteSerial)
    End If
    
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

Private Sub Form_Load()
    'On Error GoTo ErrTrap
    Dim Dcombos As New ClsDataCombos
Dim StrSQL As String
    'Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetFixedAssetsGroup DCGroup

    Dcombos.GetPrefix Me.DCPreFix, 1, 0
    Dcombos.GetCountriesNames Me.dcContryid
    Dcombos.GetCustomersSuppliers 3, Me.dcVendorid, True
  
  StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL


    Dim My_SQL As String

    ScreenNameArabic = " «·«’Ê· «·À«» …"
    ScreenNameEnglish = " Fixed Asset Data"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    'My_SQL = "  select branch_id,branch_name from branches   "
    'fill_combo DcBranch, My_SQL
If SystemOptions.UserInterface = ArabicInterface Then
    With Me.CboType
        .Clear
        .AddItem "««÷«›… ﬁÌ„Â ·√’·"
        .AddItem "œ„Ã «’·"
        .AddItem "≈” »⁄«œ «’·"
        End With
Else
  With Me.CboType
        .Clear
        .AddItem "Assets Additions"
        .AddItem "Assets Merge"
        .AddItem "Assets Exclusion"
    End With
End If
    Dcombos.GetBranches dcBranch
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select Emp_ID,Emp_name  from TblEmployee order by Emp_name   "
 Else
 My_SQL = "  select Emp_ID,Emp_namee  from TblEmployee order by Emp_namee   "
 End If
    fill_combo DCEmployee, My_SQL

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Dcombos.GetAccountingCodes Me.DcboCreditSide

    Resize_Form Me

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    AddTip
    Set rs = New ADODB.Recordset
   ' rs.Open "FixedAssets", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 
    StrSQL = "select * from  FixedAssets where 1=1"
      StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
      
        If SystemOptions.usertype <> UserAdminAll Then
         '   StrSQL = StrSQL & " and (  Branch_NO=0 or   Branch_NO=" & Current_branch & ")"
            
        End If
         StrSQL = StrSQL & " and (FlgCarNotFixed is null)"
         rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        'StrSQL = StrSQL & " order by fullcode "
        
        
        



    Me.TxtModFlg.text = "R"
    Retrive

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub
Sub SaveAssest(Optional FexdID As Double = 0)
Dim sql As String
Dim StrSQL As String
Dim Msg As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
If Me.TxtModFlg.text = "E" Then
sql = "delete  from TblAssestes where CarsDataID=" & val(XPTxtID.text) & "and FlgCarNotFixed=3 "
Cn.Execute sql
End If
sql = "Select * from TblAssestes where 1=-1"
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
Rs5.AddNew
Rs5("CarsDataID").value = val(XPTxtID.text)
Rs5("FlgCarNotFixed").value = 3
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "„‰ „·› «·«’Ê·"
Else
Msg = "From Fixed Assest File"
End If
Rs5("AsFixedID").value = FexdID
Rs5("AsDes").value = Msg
If SystemOptions.UserInterface = ArabicInterface Then
Rs5("AsName").value = TxtName.text
Else
Rs5("AsName").value = TxtNameE.text
End If
Rs5("AsCode").value = val(txtID.text)
Rs5.update

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

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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

Function EnaableFields()
     Dim i As Integer

    For i = 101 To 130
        lbl(i).ForeColor = &H0&
    Next i
 
    CBoDepreciation_Type_id.Enabled = False
    'dpStartDepreciationDate.Enabled = False
    DPLastDepreciationDate.Enabled = False
    TxtPurchasePrice.Enabled = False
    DpPurchaseDate.Enabled = False
    TxtKhordaPrice.Enabled = False
    txtCurrentValue.Enabled = False
    TxtAccDepreciation.Enabled = False
    txtinstallValue.Enabled = False
    TxtnoOfInst.Enabled = False
    txtinstallDo.Enabled = False
    txtinstallmentresult.Enabled = False
    txtPurchaseBillId.Enabled = False

    If cStatus.ListIndex = 2 Then ' „ «·«Â·«ﬂ

    End If

    If Option1.value = True And opt(0).value = True Then  'ÃœÌœ Ê·Â «Â·«ﬂ
        CBoDepreciation_Type_id.Enabled = True  '‰Ê⁄ «·«Â·«ﬂ
        dpStartDepreciationDate.Enabled = True ' «—ÌŒ »œ«Ì… «·«Â·«ﬂ
        TxtKhordaPrice.Enabled = True '”⁄—«·«’· Œ—œ…
        lbl(105).ForeColor = vbBlue
        lbl(127).ForeColor = vbBlue
        lbl(121).ForeColor = vbBlue
 
    ElseIf Option1.value = True And opt(1).value = True Then  'ÃœÌœ Ê·Ì” ·Â «Â·«ﬂ

    ElseIf Option2.value = True And opt(0).value = True Then  '«ﬁ  «ÕÌ  Ê ·Â «Â·«ﬂ
        CBoDepreciation_Type_id.Enabled = True  '‰Ê⁄ «·«Â·«ﬂ
        dpStartDepreciationDate.Enabled = True ' «—ÌŒ »œ«Ì… «·«Â·«ﬂ
        DPLastDepreciationDate.Enabled = True ' «—ÌŒ «Œ— «Â·«ﬂ
        TxtPurchasePrice.Enabled = True '”⁄— «·‘—«¡
        DpPurchaseDate.Enabled = True ' «—ÌŒ «·‘—«¡
        txtPurchaseBillId.Enabled = True '—ﬁ„ ›« Ê—… «·‘—«¡
        TxtKhordaPrice.Enabled = True '”⁄—«·«’· Œ—œ…
        'TxtCurrentValue.Enabled = True '«·ﬁÌ„… «·œ› —Ì… ··«’·
        TxtAccDepreciation.Enabled = True '„Ã„⁄ «·«Â·«ﬂ
        'txtinstallDo.Enabled = True '    ⁄œœ «·«ﬁ”«ÿ «·„‰›–…

        lbl(105).ForeColor = vbBlue
        lbl(127).ForeColor = vbBlue
        lbl(120).ForeColor = vbBlue
        lbl(106).ForeColor = vbBlue
        lbl(128).ForeColor = vbBlue
        lbl(121).ForeColor = vbBlue
        lbl(116).ForeColor = vbBlue
        ' lbl(107).ForeColor = vbblue
        lbl(129).ForeColor = vbBlue
        ' lbl(130).ForeColor = vbblue
    ElseIf Option2.value = True And opt(1).value = True Then  '«ﬁ  «ÕÌ  Ê  ·Ì” ·Â «Â·«ﬂ
        TxtPurchasePrice.Enabled = True '”⁄— «·‘—«¡
        DpPurchaseDate.Enabled = True ' «—ÌŒ «·‘—«¡
        txtPurchaseBillId.Enabled = True '—ﬁ„ ›« Ê—… «·‘—«¡
 
        lbl(106).ForeColor = vbBlue
        lbl(128).ForeColor = vbBlue
        lbl(116).ForeColor = vbBlue

    End If

End Function

Function CheckEnteredData() As Boolean
    CheckEnteredData = False

    If Option1.value = True And opt(0).value = True Then  'ÃœÌœ Ê·Â «Â·«ﬂ
        If val(Me.CBoDepreciation_Type_id.ListIndex) = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ ' ÿ—Ìﬁ… «·Â·«ﬂ", vbCritical
            Else
                MsgBox "Select Depreciation Method Firstly    ", vbCritical
            End If

            CBoDepreciation_Type_id.SetFocus
            Sendkeys "{F4}"
            Exit Function
        End If
 
        If val(TxtKhordaPrice.text) = 0 Then
            '      If SystemOptions.UserInterface = ArabicInterface Then
            '     MsgBox "Õœœ ”⁄— «·«’· ﬂŒ—œ… «Ê·«      ", vbCritical
            '     Else
            '     MsgBox "Enter Khorda Price   Firstly    ", vbCritical
            '     End If
            '
            '    TxtKhordaPrice.SetFocus
            '    Exit Function
        End If
   
    ElseIf Option1.value = True And opt(1).value = True Then  'ÃœÌœ Ê·Ì” ·Â «Â·«ﬂ

    ElseIf Option2.value = True And opt(0).value = True Then  '«ﬁ  «ÕÌ  Ê ·Â «Â·«ﬂ
    
        If val(Me.CBoDepreciation_Type_id.ListIndex) = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ ' ÿ—Ìﬁ… «·Â·«ﬂ", vbCritical
            Else
                MsgBox "Select Depreciation Method Firstly    ", vbCritical
            End If

            CBoDepreciation_Type_id.SetFocus
            Sendkeys "{F4}"
            Exit Function
        End If

        If val(TxtPurchasePrice.text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ ”⁄— ‘—«¡ «·«’·      «Ê·«", vbCritical
            Else
                MsgBox "Enter Purcahse Price  Firstly    ", vbCritical
            End If
    
            TxtPurchasePrice.SetFocus
            Exit Function
        End If

        If txtPurchaseBillId.text = "" Then
            '       If SystemOptions.UserInterface = ArabicInterface Then
            '      MsgBox "Õœœ —ﬁ„ ›« Êƒ… ‘—«¡ «·«’·    «Ê·«", vbCritical
            '      Else
            '      MsgBox "Enter Purcahse Price  Bill NO. Firstly    ", vbCritical
            '      End If
    
            '  txtPurchaseBillId.SetFocus
            '  Exit Function
            txtPurchaseBillId.text = 0
        End If
    
        If val(TxtKhordaPrice.text) = 0 Then
            '          If SystemOptions.UserInterface = ArabicInterface Then
            ''         MsgBox "Õœœ ”⁄— «·«’· ﬂŒ—œ… «Ê·«      ", vbCritical
            '        Else
            '        MsgBox "Enter Khorda Price   Firstly    ", vbCritical
            '        End If
            '
            '    TxtKhordaPrice.SetFocus
            '    Exit Function
        End If
    
        '       If Val(TxtCurrentValue.text) = 0 Then
        '             If SystemOptions.UserInterface = ArabicInterface Then
        '            MsgBox "Õœœ      ﬁÌ„… «·«’· «·œ› —Ì…  «Ê·«", vbCritical
        '            Else
        '            MsgBox "Enter Current Price   Firstly    ", vbCritical
        '            End If
        '
        '        TxtCurrentValue.SetFocus
        '        Exit Function
        '    End If
   
        If val(TxtAccDepreciation.text) = 0 Then
        '    If SystemOptions.UserInterface = ArabicInterface Then
        '        MsgBox "Õœœ      ﬁÌ„… „Ã„⁄ «·«Â·«ﬂ  «·«’·    «Ê·«", vbCritical
        '    Else
        '        MsgBox "Enter Acc. DepreciationFirstly    ", vbCritical
        '    End If
    '
    '        TxtAccDepreciation.SetFocus
    '        Exit Function
        End If
    
        '       If Val(txtinstallDo.text) = 0 Then
        '             If SystemOptions.UserInterface = ArabicInterface Then
        '            MsgBox "Õœœ ⁄œœ «·«ﬁ”«ÿ «·„‰›–…  Õ Ï «·«‰   «Ê·«", vbCritical
        '            Else
        '            MsgBox "Enter NO Of Executed Installments   ", vbCritical
        '            End If
        '
        '        txtinstallDo.SetFocus
        '        Exit Function
        '    End If
 
    ElseIf Option2.value = True And opt(1).value = True Then  '«ﬁ  «ÕÌ  Ê  ·Ì” ·Â «Â·«ﬂ

        If val(TxtPurchasePrice.text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ ”⁄— ‘—«¡ «·«’·      «Ê·«", vbCritical
            Else
                MsgBox "Enter Purcahse Price  Firstly    ", vbCritical
            End If
    
            TxtPurchasePrice.SetFocus
            Exit Function
        End If

        '    If txtPurchaseBillId.text = "" Then
        '             If SystemOptions.UserInterface = ArabicInterface Then
        '            MsgBox "Õœœ —ﬁ„ ›« Êƒ… ‘—«¡ «·«’·    «Ê·«", vbCritical
        '            Else
        '            MsgBox "Enter Purcahse Price  Bill NO. Firstly    ", vbCritical
        '            End If
        '
        '        txtPurchaseBillId.SetFocus
        '        Exit Function
        '    End If

    End If

    CheckEnteredData = True
End Function



Private Sub GridInstallments_CellButtonClick(ByVal row As Long, ByVal Col As Long)
With Me.GridInstallments
Select Case .ColKey(Col)
Case "Show"
  If checkApility("FrmExpenses40A") = False Then
                Exit Sub
            End If
 Unload FrmExpenses40A
Load FrmExpenses40A
FrmExpenses40A.show
   FrmExpenses40A.Retrive (.TextMatrix(row, .ColIndex("ID")))

End Select
End With
End Sub

Private Sub GridInstallments_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.GridInstallments
Select Case .ColKey(Col)
 Case "Show"
            .ColComboList(.ColIndex("Show")) = "..."
     End Select
    End With
End Sub

Private Sub Opt_Click(index As Integer)
    EnaableFields
End Sub

Private Sub Option1_Click()
    EnaableFields
End Sub

Private Sub Option2_Click()
    EnaableFields
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TxtAddValue_Change()
TxtBookValue.text = val(txtCurrentValue.text) + val(TxtAddValue.text) - val(TxtDisposedValue.text)
End Sub

Private Sub TxtCurrentValue_Change()
TxtBookValue.text = val(txtCurrentValue.text) + val(TxtAddValue.text) - val(TxtDisposedValue.text)
End Sub

Private Sub TxtDisposedValue_Change()
TxtBookValue.text = val(txtCurrentValue.text) + val(TxtAddValue.text) - val(TxtDisposedValue.text)
End Sub

Private Sub txtinstallmentresult_Change()
TxtQstRemNo.text = val(TxtQstAdNo.text) + val(txtinstallmentresult.text)
End Sub

Private Sub TxtKhordaPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtKhordaPrice.text, 0)
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, txtID.text, 0)
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '   Me.Caption = "«·«’Ê· «·À«» …"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            cStatus.Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            Frame3.Enabled = False
        
            If rs.RecordCount < 1 Then
                '      Me.XPBtnMove(0).Enabled = False
                '      Me.XPBtnMove(1).Enabled = False
                '      Me.XPBtnMove(2).Enabled = False
                '      Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

        Case "N"
            '   Me.Caption = "√‰Ê«⁄ «·„’—Ê›« ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            cStatus.Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
         
        Case "E"
            Frame3.Enabled = False
            '   Me.Caption = "√‰Ê«⁄ «·„’—Ê›« (  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            cStatus.Enabled = False
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
         
    End Select

    Exit Sub
ErrTrap:
End Sub

Function getCurrentData(FixedassetId As Integer, Optional ByRef currentvalue As Double, Optional ByRef AccDepreciation As Double)

    currentvalue = 789
    AccDepreciation = 1011
End Function

Public Sub Retrive(Optional Lngid As Long = 0)
    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If Not (rs.EOF Or rs.BOF) Then
        If Lngid <> 0 Then
            rs.Find "id=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If

    End If
    DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)

    Me.BiLLID.text = IIf(IsNull(rs("BiLLID").value), 0, (rs("BiLLID").value))
        If IsNull(rs("ISEQUP").value) Then
        ISEQUP.value = vbUnchecked
   
    Else

        If rs("ISEQUP").value = True Then
            ISEQUP.value = vbChecked
        Else
            ISEQUP.value = vbUnchecked
        End If
    End If
    
    Me.XPTxtID.text = IIf(val(rs("id").value) = 0, 0, val(rs("id").value))
    Me.txtID.text = IIf(IsNull(rs("code").value), "", rs("code").value)
    Me.TxtName.text = IIf(IsNull(rs("Name").value), "", rs("Name").value)
    Me.TxtNameE.text = IIf(IsNull(rs("Namee").value), "", rs("Namee").value)
    TxtQuantity.text = IIf(IsNull(rs("Quantity").value), "", rs("Quantity").value)
    Me.TxtFullcode.text = IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)
    
      If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    Else
        Me.DcCostCenter.BoundText = ""
    End If
    
    If rs("IsContainer").value = vbTrue Then
        Me.chkIsContainer.value = vbChecked
    Else
        Me.chkIsContainer.value = vbUnchecked
    End If
    
    Me.TxtNotes.text = IIf(IsNull(rs("Notes").value), "", rs("Notes").value)
    dcVendorid.BoundText = IIf(IsNull(rs("Vendorid").value), "", (rs("Vendorid").value))
    dcContryid.BoundText = IIf(IsNull(rs("Contryid").value), "", (rs("Contryid").value))
    Me.txtBoardNo.text = IIf(IsNull(rs("BoardNo").value), "", rs("BoardNo").value)
    Me.txtSerialNo.text = IIf(IsNull(rs("SerialNo").value), "", rs("SerialNo").value)

    Me.TxtMinusValue.text = IIf(IsNull(rs("MinusValue").value), 0, rs("MinusValue").value)
'''///////

Me.TxtAddValue.text = GetAddValue(val(Me.XPTxtID.text))

Me.TxtQstAdNo.text = GetQstAddNo(val(Me.XPTxtID.text))

''////////
    Me.TxtModel.text = IIf(IsNull(rs("Model").value), "", rs("Model").value)
    Me.txtChaseeNo.text = IIf(IsNull(rs("ChaseeNo").value), "", rs("ChaseeNo").value)
    Me.TxtOprNo.text = IIf(IsNull(rs("OprNo").value), "", rs("OprNo").value)
    
    dpEndLicense.value = IIf(IsNull(rs("EndLicense").value), Date, rs("EndLicense").value)
    dpEndTest.value = IIf(IsNull(rs("EndTest").value), Date, rs("EndTest").value)
          
    DCGroup.BoundText = IIf(IsNull((rs("group_id").value)), 0, (rs("group_id").value))
    dcBranch.BoundText = IIf(IsNull(rs("Branch_NO").value), 0, (rs("Branch_NO").value))
    DCEmployee.BoundText = IIf(IsNull(rs("Emp_id").value), 0, (rs("Emp_id").value))
    DPReceiveDate.value = IIf(IsNull(rs("ReceiveDate").value), Date, rs("ReceiveDate").value)
    txtCurrentValue.text = IIf(IsNull((rs("CurrentValue").value)), 0, rs("CurrentValue").value)
    TxtAccDepreciation.text = IIf(IsNull((rs("AccDepreciation").value)), 0, rs("AccDepreciation").value)
    cStatus.ListIndex = IIf(IsNull((rs("Status_id").value)), -1, rs("Status_id").value)
    CBoDepreciation_Type_id.ListIndex = IIf(IsNull((rs("Depreciation_Type_id").value)), -1, rs("Depreciation_Type_id").value)
    TxtAge.text = IIf(IsNull(rs("DefaultAge").value), 0, (rs("DefaultAge").value))
    dpStartDepreciationDate.value = IIf(IsNull(rs("StartDepreciationDate").value), Date, rs("StartDepreciationDate").value)
    Me.DPLastDepreciationDate.value = IIf(IsNull(rs("LastDepreciationDate").value), Date, rs("LastDepreciationDate").value)
    txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)
           
    TxtnoOfInst.text = IIf(IsNull((rs("noOfInstallments").value)), 0, rs("noOfInstallments").value)
    txtinstallDo.text = IIf(IsNull((rs("EXEInstallments").value)), 0, rs("EXEInstallments").value)
    txtinstallmentresult.text = IIf(IsNull((rs("RemainInstallments").value)), 0, rs("RemainInstallments").value)
    txtinstallValue.text = IIf(IsNull((rs("Installmentvalue").value)), 0, rs("Installmentvalue").value)
           
    TxtPrice.text = IIf(IsNull((rs("Price").value)), 0, rs("Price").value)
           
    txtTotal = val(TxtPrice) * val(TxtQuantity)
    ' Dim PurchasePrice As Double
    ' Dim PurchaseDate As Date
    ' Dim PurchaseBillId As String
   '  getPurchaseInformations val(Me.XPTxtID), PurchaseDate, purchaseprice, PurchaseBillId
    ' DpPurchaseDate.value = PurchasePrice
    'TxtPurchasePrice.text = PurchaseDate
    'txtPurchaseBillId.text = PurchaseBillId
    
    DpPurchaseDate.value = IIf(IsNull((rs("PurchaseDate").value)), Date, rs("PurchaseDate").value)
    TxtPurchasePrice.text = IIf(IsNull((rs("purchaseprice").value)), 0, rs("purchaseprice").value)
    txtPurchaseBillId.text = IIf(IsNull(rs("PurchaseBillId").value), "", rs("PurchaseBillId").value)
    TxtKhordaPrice.text = IIf(IsNull((rs("KhordaPrice").value)), 0, rs("KhordaPrice").value)
If Not (IsNull(rs("New_or_opening").value)) Then
    If rs("New_or_opening").value = 0 Then
        Option1.value = True
    Else
        Option2.value = True
    End If
  Else
  Option1.value = True
End If
    DCPreFix.text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
    Dim YearMove As Double
    Dim YearNotMove As Double
   GetAssestMoveYearly DpPurchaseDate.value, YearMove, YearNotMove
   TxtYearMove.text = YearMove
   TxtYearNotMove.text = YearNotMove
   RetriveMoveassest
    '   XPDtbTrans.value = rs("ReceiveDate")
    DCGroup_Click (0)
    FullGridData val(Me.XPTxtID.text)
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    
     mIsFinishSave = True
    Exit Sub
ErrTrap:
End Sub
Sub RetriveMoveassest()
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " select  * from FixedAssetsGroup where GroupID=" & val(DCGroup.BoundText) & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
    If Not IsNull(rs2("AsstMove").value) Then
    If (rs2("AsstMove").value) = 1 Then
    RdMove(1).value = True
    Else
    RdMove(0).value = True
    End If
    Else
    RdMove(0).value = True
    End If
End If
End Sub

Function getPurchaseInformations(FixedassetId As Integer, Optional ByRef PurchaseDate As Date, Optional ByRef purchaseprice As Double, Optional ByRef PurchaseBillId As String)
    PurchaseDate = "05-03-2000"
    purchaseprice = 7800
    PurchaseBillId = "123"
End Function

Private Sub TxtName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNameE_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub TxtnoOfInst_Change()
TotalQest.text = val(Me.TxtnoOfInst.text) + val(TxtQstAdNo.text)
End Sub

Private Sub txtPassword_Change()
If Trim(txtPassword) = "Alex2025" Then
    cmdReSave.Visible = True
Else
    cmdReSave.Visible = False
End If

End Sub

Private Sub TxtPercentage1_Change()

    If val(TXtPercentage1) = 0 Then
        Me.cStatus.ListIndex = -1
        Me.cStatus.Enabled = False
    Else
 '   if me.tx
        Me.cStatus.Enabled = True
    End If

End Sub

Private Sub txtPrice_Change()
txtTotal = val(TxtPrice) * val(TxtQuantity)
If val(txtTotal) <> 0 Then
    TxtPurchasePrice = txtTotal
End If
End Sub

Private Sub TxtQstAdNo_Change()
TxtQstRemNo.text = val(TxtQstAdNo.text) + val(txtinstallmentresult.text)
TotalQest.text = val(Me.TxtnoOfInst.text) + val(TxtQstAdNo.text)
End Sub

Private Sub txtQuantity_Change()
txtTotal = val(TxtPrice) * val(TxtQuantity)
End Sub

Private Sub XPBtnMove_Click(index As Integer)
    'On Error GoTo ErrTrap
 
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

Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim RsDev As New ADODB.Recordset
    Dim RsNot As New ADODB.Recordset

    Dim BeginTrans As Boolean
 '  On Error GoTo ErrTrap
    If IsSaveWithOutMsg Then GoTo GoIsSaveWithOutMsg
    If Me.TxtModFlg.text <> "R" Then
 
        If TxtName.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ «”„ «·«’·  «Ê·«", vbCritical
            Else
                MsgBox "Select Name Firstly    ", vbCritical
            End If
    
            TxtName.SetFocus
            Exit Sub
        End If
    
    
        If TxtNameE.text = "" Then
    TxtNameE.text = TxtName.text
    End If
        If val(Me.dcBranch.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical
            Else
                MsgBox "Select Branch Firstly    ", vbCritical
            End If

            dcBranch.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
 
        If val(Me.DCGroup.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ „Ã„Ê⁄Â «·«’·  «Ê·«", vbCritical
            Else
                MsgBox "Select Group Firstly    ", vbCritical
            End If

            DCGroup.SetFocus
             Sendkeys "{F4}"
            Exit Sub
        End If
 
        If Me.cStatus.ListIndex = -1 And opt(0).value = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ  Õ«·Â  «·«’·     ", vbCritical
            Else
                MsgBox "Select Status     ", vbCritical
            End If

'            cStatus.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
    
  '      If val(Me.DcEmployee.BoundText) = 0 Then
  '          If SystemOptions.UserInterface = ArabicInterface Then
  '              MsgBox "Õœœ   «·«’· »⁄ÂœÂ  «Ê·«", vbCritical
  '          Else
  '              MsgBox "Select Holder Name   ", vbCritical
  '          End If
  '
  '          DcEmployee.SetFocus
  '          SendKeys "{F4}"
  '          Exit Sub
  '      End If
   
        If opt(0).value = True Then
            If Me.CBoDepreciation_Type_id.ListIndex = -1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Õœœ     'ÿ—Ìﬁ… «·«Â·«ﬂ     ", vbCritical
                Else
                    MsgBox "Specify DDD type       ", vbCritical
                End If

                CBoDepreciation_Type_id.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        End If

        '
        If CheckEnteredData = False Then
            Exit Sub
        End If
GoIsSaveWithOutMsg:
        Dim noOfInstallments As Integer
        Dim Age As Integer
        Dim currentvalue As Double
        Dim installValue As Double
        Dim RemainInstallments As Double
        Dim EXEInstallments As Double
'Â‰« »Ì»œ√ ÌÕ”» «·ﬁÌ„Â «·œ› —Ì… ··«’· Ê„Ã„⁄ «·«Â·«ﬂ Ê⁄œœ «·«ﬁ”««ÿ «·„‰›–Â ÊﬁÌ„Â «·ﬁ”ÿ Ê„ »ﬁÌ  ﬂ«„ ﬁ”ÿ Ê⁄œœ «·«ﬁ”«ÿ «·«Ã„«·ÌÂ
        GetAndCalculateAll val(Me.XPTxtID), val(Me.TXtPercentage1.text), noOfInstallments, Age, val(Me.TxtPurchasePrice.text), val(Me.TxtKhordaPrice.text), val(Me.TxtAccDepreciation.text), currentvalue, installValue, EXEInstallments, RemainInstallments

        Me.TxtnoOfInst.text = noOfInstallments
        Me.TxtAge.text = Age

        If Option1.value = True Then
            Me.txtCurrentValue.text = 0
            Me.txtinstallValue.text = 0
            Me.txtinstallmentresult.text = 0
            Me.txtinstallDo.text = 0
            TxtAccDepreciation.text = 0

        Else
            Me.txtCurrentValue.text = currentvalue '«·ﬁÌ„Â «·œ› —Ì…
            Me.txtinstallValue.text = installValue 'ﬁÌ„Â «·ﬁ”ÿ
            Me.txtinstallmentresult.text = RemainInstallments '«·«ﬁ”«ÿ «·„ »ﬁÌ…
            Me.txtinstallDo.text = val(EXEInstallments) '«·«ﬁ”«ÿ «·„‰›–…
        End If
    
        If opt(1).value = True Then
            Me.txtCurrentValue.text = TxtPurchasePrice.text
            Me.txtinstallValue.text = 0
            Me.txtinstallmentresult.text = 0
            Me.txtinstallDo.text = 0
            TxtAccDepreciation.text = 0
            TxtKhordaPrice.text = 0
        End If

        Select Case Me.TxtModFlg.text

            Case "N"

                If TxtNoteSerial.text = "" And Option1.value = False Then
        '            If Notes_coding(val(Me.dcBranch.BoundText), XPDtbTrans.value) = "error" Then
        '                If SystemOptions.UserInterface = ArabicInterface Then
        '                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        '                Else
        '                    MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number ": Exit Sub
        '                End If
'
'                    ElseIf Notes_coding(val(Me.dcBranch.BoundText), XPDtbTrans.value) = "" Then
'
'                        If SystemOptions.UserInterface = ArabicInterface Then
'                            MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
'                        Else
'                            MsgBox "You must Define JE Coding ": Exit Sub
'                        End If
'
'                    Else
'                        txtNoteSerial.text = Notes_coding(val(Me.dcBranch.BoundText), XPDtbTrans.value)
'                        txtNoteID = CStr(new_id("Notes", "NoteID", "", True))
                 '   End If
                End If

                StrSQL = "select * From  FixedAssets where Name='" & Trim(TxtName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    Msg = "Â‰«ﬂ «’· „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «’· «·„Õœœ"
             '       MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
             '       TxtName.SetFocus
             '       Exit Sub
                End If

            Case "E"
             StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End Select

        If Me.TxtModFlg.text = "N" Then
   
        End If

        Cn.BeginTrans
        BeginTrans = True

        Select Case Me.TxtModFlg.text

            Case "N"
                XPTxtID.text = CStr(new_id("FixedAssets", "id", "", True))
            
                rs.AddNew
            
            Case "E"
        
        End Select

     
        rs("id").value = val(Me.XPTxtID.text)
        
               If ISEQUP.value = vbChecked Then
            rs("ISEQUP").value = 1
        Else
            rs("ISEQUP").value = 0
        End If



        rs("code").value = txtID.text
        rs("UserID").value = val(DCboUserName.BoundText)
        rs("Name").value = IIf(Trim(TxtName.text) = "", Null, TxtName.text)
        rs("Namee").value = IIf(Trim(TxtNameE.text) = "", Null, TxtNameE.text)
        
        rs("Notes").value = IIf(Trim(TxtNotes.text) = "", Null, TxtNotes.text)
            
        rs("BiLLID").value = IIf((BiLLID.text) = "", Null, val(BiLLID.text))
        rs("Quantity").value = IIf((TxtQuantity.text) = "", Null, val(TxtQuantity.text))
                    
          
           
        rs("group_id").value = IIf(val(DCGroup.BoundText) = 0, Null, DCGroup.BoundText)
        rs("Branch_NO").value = IIf(val(dcBranch.BoundText) = 0, Null, dcBranch.BoundText)
      rs("Emp_id").value = IIf(val(DCEmployee.BoundText) = 0, Null, DCEmployee.BoundText)
        rs("ReceiveDate").value = DPReceiveDate.value

        If Option1.value = True Then 'ÃœÌœ Ê   ·Â «Â·«ﬂ
            rs("NoteID").value = Null
            rs("NoteSerial").value = Null
        Else
            rs("NoteID").value = IIf(Trim(Me.TXTNoteID.text) = "", Null, TXTNoteID.text)
            rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, TxtNoteSerial.text)
        End If
           
        If Me.chkIsContainer.value = vbChecked Then
            rs.Fields("IsContainer").value = 1
        Else
            rs.Fields("IsContainer").value = 0
        End If
           
        If Option2.value = True And Me.cStatus.ListIndex = 3 Then '   «›  «ÕÌ Ê „ «· Œ·’ „‰…
            '  rs("NoteID1").value = IIf(Trim(Me.txtNoteID1.text) = "", Null, txtNoteID1.text)
            '  rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, TxtNoteSerial1.text)
            '   rs("SalePrice").value = Val(Me.TxtSalePrice)
            '    rs("variance").value = Val(Me.TxtSalePrice)
            
        Else
            '             rs("NoteID1").value = Null
            '            rs("NoteSerial1").value = Null
            
        End If
        '''//////////
   
        rs("CurrentValue").value = IIf(val(txtCurrentValue.text) = 0, 0, val(txtCurrentValue.text))
        rs("AccDepreciation").value = IIf(val(TxtAccDepreciation.text) = 0, 0, val(TxtAccDepreciation.text))
        'rs("Status_id").value = GetStatus_id
        rs("Status_id").value = cStatus.ListIndex
        rs("Depreciation_Type_id").value = CBoDepreciation_Type_id.ListIndex
        'rs("DefaultAge").value = GetDefaultAge
        rs("DefaultAge").value = val(TxtAge.text)
           
        rs("StartDepreciationDate").value = dpStartDepreciationDate.value
        'Dim LastDepreciationDate As Date
        'If Option1.value = True Then
        'LastDepreciationDate = getLastDepreciationDate
        ' Me.DPLastDepreciationDate.value = LastDepreciationDate
        'End If
        rs("LastDepreciationDate").value = Me.DPLastDepreciationDate.value
        
        rs("Price").value = val(TxtPrice)
        
        
        ' Dim NoOfInstallments As Integer
        ' Dim EXEInstallments As Integer
        ' Dim RemainInstallments As Integer
        ' Dim InstallmentValue As Double
        'GetInstallmentsInformations Me.XPTxtID, NoOfInstallments _
        ', EXEInstallments, RemainInstallments, , , Val(TXtPercentage1.text), InstallmentValue
         
        '    rs("NoOfInstallments").value = NoOfInstallments
         
        '     rs("RemainInstallments").value = RemainInstallments
        '     rs("InstallmentValue").value = InstallmentValue
        'If Option1.value = True Then
        '
        '     rs("EXEInstallments").value = EXEInstallments
        '
        '   Else
         
        '     rs("EXEInstallments").value = Val(Me.txtinstallDo)
        'End If
            
        rs("NoOfInstallments").value = val(TxtnoOfInst.text)
        rs("RemainInstallments").value = val(txtinstallmentresult.text)
        rs("InstallmentValue").value = val(txtinstallValue.text)
        rs("EXEInstallments").value = val(Me.txtinstallDo.text)

        '   Dim PurchasePrice As Double
        '    Dim PurchaseDate As Data
        '    Dim PurchaseBillId As String
        '
        '    getPurchaseInformations Val(Me.XPTxtID), PurchaseDate, PurchasePrice, PurchaseBillId
        '     If Option1.value = True Then
        '     rs("PurchasePrice").value = PurchasePrice
        '     rs("PurchaseDate").value = PurchaseDate
        ''     rs("PurchaseBillId") = PurchaseBillId
        '     Else
        'rs("PurchasePrice").value = Val(TxtPurchasePrice.text)
        'rs("PurchaseDate").value = DpPurchaseDate.value
        'rs("PurchaseBillId") = txtPurchaseBillId.text
        'End If
        rs("PurchasePrice").value = val(TxtPurchasePrice.text)
        rs("PurchaseDate").value = DpPurchaseDate.value
        rs("PurchaseBillId") = txtPurchaseBillId.text
        rs("KhordaPrice").value = IIf(val(TxtKhordaPrice.text) = 0, 0, val(TxtKhordaPrice.text))
               
        If Option1.value = True Then
            rs("New_or_opening").value = 0
        Else
            rs("New_or_opening").value = 1
        End If

        rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.text) & IIf(Trim(txtID.text) = "", Null, txtID.text)
        rs("prifix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)

        If Me.opt(0).value = True Then
            rs("HaveDepreciation").value = 1
        Else
        
            rs("HaveDepreciation").value = 0
        End If
           
        rs("Vendorid").value = IIf(val(Me.dcVendorid.BoundText) = 0, Null, val(Me.dcVendorid.BoundText))
        rs("Contryid").value = IIf(val(Me.dcContryid.BoundText) = 0, Null, val(Me.dcContryid.BoundText))
        rs("BoardNo").value = IIf((txtBoardNo.text) = "", Null, (txtBoardNo.text))
        rs("Model").value = IIf((TxtModel.text) = "", Null, (TxtModel.text))
        rs("SerialNo").value = IIf((txtSerialNo.text) = "", Null, (txtSerialNo.text))
           rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
  
        rs("ChaseeNo").value = IIf((txtChaseeNo.text) = "", Null, (txtChaseeNo.text))
        rs("OprNo").value = IIf((TxtOprNo.text) = "", Null, (TxtOprNo.text))
        
        
        rs("EndLicense").value = dpEndLicense.value
        rs("EndTest").value = dpEndTest.value
          txtopening_balance_voucher_id = 0
         If Option1.value = True Then
            txtopening_balance_voucher_id = 0
        End If
         
        If val(txtopening_balance_voucher_id) = 0 And Option2.value = True Then
            txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
        End If

        rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
        
        rs.update
    End If
 
    '**************************************************************************

    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
    
    If Option2.value = True And opt(0).value = True And cStatus.ListIndex = 0 Then '   Ê«·Õ«·… Ã«—Ì  «·«Â·«ﬂ '«›  «ÕÌ Ê·Â «Â·«ﬂ
    ' »Ì÷Ì›  ›«’Ì· «ﬁ”«ÿ «·«Â·«ﬂ ⁄‘«‰  ŸÂ— »‘«‘Â «·«ﬁ”«ÿ
        updateFixedAsseTInstallmentInformations val(Me.XPTxtID.text), , , , Me.XPDtbTrans.value, , , , False, True

    End If

    If Option1.value = True Then
    Else

        If CreateJL = False Then '«‰‘«¡ «·ﬁÌÊœ «·«›  «ÕÌ…
            GoTo ErrTrap
        End If
   
    End If

    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    CuurentLogdata
    '  SaveAssest   œÂ » Õ›Ÿ ‘«‘Â «·«’Ê· ›Ì „·› «·⁄Âœ ⁄‘«‰ Ìﬁœ— Ì” Œœ„Â« ‘∆Ê‰ «·„ÊŸ›Ì‰
SaveAssest val(XPTxtID.text)

If IsSaveWithOutMsg Then Exit Sub
    Select Case Me.TxtModFlg.text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·‰Ê⁄" & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If
            
        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End Select

    TxtModFlg.text = "R"
 
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
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

Function CreateJL() As Boolean
    CreateJL = False
    Dim LngDevID As Long
    Dim DepitAccount As String
    Dim CreditAccount1 As String
    Dim CreditAccount2 As String
    Dim Msg As String
    'GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 24, Val(Me.DcBranch.BoundText), DepitAccount    'Õ”«» «·«’·
    'GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 26, Val(Me.DcBranch.BoundText), CreditAccount1    '„Ã„⁄ «·«Â·«ﬂ

    Dim Account_code As String
    Dim Account_code2 As String

    GetFixedAssetsGroupAccount val(DCGroup.BoundText), , val(Me.dcBranch.BoundText), , , , , , Account_code, , Account_code2
    DepitAccount = Account_code
    CreditAccount1 = Account_code2

    CreditAccount2 = get_account_code_branch(41, val(Me.dcBranch.BoundText))

    If CreditAccount2 = "NO branch" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÌÊÃœ Õ”«»«  ·Â–« «·›—⁄"
        Else
            Msg = "No Accounts For This Branch"
        End If

        MsgBox Msg, vbCritical
        CreateJL = False
        Exit Function

    ElseIf CreditAccount2 = "NO account" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Õ”«» Ê”Ìÿ «›  «ÕÌ ··«’Ê· €Ì— „Õœœ ›Ï «·›—⁄"
        Else
            Msg = "Fixed Asset Opening Balance Account Not Defined In this Branch"
        End If

        MsgBox Msg, vbCritical
        CreateJL = False
        Exit Function
    End If

    Dim sql As String

    'sql = "Delete   from notes where NoteID=" & Val(TxtNoteID.text)
    'Cn.Execute sql
    '«‰‘«¡ «·ﬁÌÊœ
    If Option1.value = True Or (TXT24.text) = "" Then    'ÃœÌœ
        CreateJL = True
        Exit Function
    Else
        '   Dim RsNotes As ADODB.Recordset
        '   Dim RsDev As ADODB.Recordset
        '   Dim NoteID As String
        '   Set RsNotes = New ADODB.Recordset
        Dim StrSQL As String
   
        '   RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
        '        Set RsDev = New ADODB.Recordset
        '        RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        If Me.TxtModFlg.text = "N" Then
            '                RsNotes.AddNew
            '    my_branch = Val(Me.dcBranch.BoundText)
            '                RsNotes("NoteID").value = CStr(TXTNoteID.text)
            '                RsNotes("Note_Value").value = Val(TxtPurchasePrice.text)
            '               RsNotes("branch_no").value = Val(Me.dcBranch.BoundText)
            '                RsNotes("Remark").value = ""
            '                RsNotes("NoteType").value = 90
            '                RsNotes("NoteDate").value = XPDtbTrans.value
            '                RsNotes("UserID").value = user_id
            '                RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '„”·”· «·ﬁÌœ
            '                RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
            '                RsNotes("sanad_year").value = year(Date)
            '                RsNotes("sanad_month").value = Month(Date)
            ''                RsNotes("note_value_by_characters").value = WriteNo(Format(Val(TxtPurchasePrice.text), "0.00"), 0, True, ".")
            '                RsNotes.update
        Else
       '     Cn.Execute "Delete DOUBLE_ENTREY_VOUCHERS  Where Notes_ID=" & val(txtNoteID.text)
            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
             
        End If

        Dim des As String
        Dim LngOpenID  As Long
        LngOpenID = 1

        If Option2.value = True And opt(1).value = True And Me.cStatus.ListIndex = -1 Then '«›  «ÕÌ Ê ·Ì” ·Â «Â·«ﬂ
            ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
  
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
            
            If SystemOptions.UserInterface <> ArabicInterface Then
                des = "Fixed Asset Opening Balance For Asset  " & Me.TxtName & "  And have No Depreciation '"
            Else
                des = "»‰«¡ ⁄·Ï —’Ìœ «›  «ÕÌ ··«’· " & Me.TxtName & "  Ê·Ì” ·Â «Â·«ﬂ -ﬁÌ„… «·«’·'"
            End If
            
            If ModAccounts.AddNewDev(LngDevID, 0, DepitAccount, val(Me.TxtPurchasePrice.text), 0, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , , , , True, val(Me.txtopening_balance_voucher_id.text), val(Me.XPTxtID.text), val(Me.DCGroup.BoundText), val(Me.dcBranch.BoundText), val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
 
            If ModAccounts.AddNewDev(LngDevID, 1, CreditAccount2, val(Me.TxtPurchasePrice.text), 1, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , , , , True, val(Me.txtopening_balance_voucher_id.text), val(Me.XPTxtID.text), val(Me.DCGroup.BoundText), val(Me.dcBranch.BoundText), val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
 
            '„œÌ‰
            '     If ModAccounts.AddNewDev(LngDevID, 0, _
                  DepitAccount, Val(TxtPurchasePrice.text), 0, _
                  des, Val(Me.TxtNoteID), , , _
                  SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
                  , , , , , Val(Me.XPTxtID.text), Val(Me.DCGroup.BoundText), Val(Me.DcBranch.BoundText)) = False Then
            '         GoTo ErrTrap
                    
            '    End If
            '            œ«∆‰ 1
            '  If ModAccounts.AddNewDev(LngDevID, 1, _
               CreditAccount2, Val(TxtPurchasePrice.text), 1, _
               des, Val(Me.TXTNoteID), , , _
               SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
               , , , , , Val(Me.XPTxtID.text), Val(Me.DCGroup.BoundText), Val(Me.dcBranch.BoundText)) = False Then
            '         GoTo ErrTrap
                    
            '    End If
            
        ElseIf Option2.value = True And opt(0).value = True And Me.cStatus.ListIndex = 0 Then '  Ê«·Õ«·… Ã«—Ì «·«Â·«ﬂ' '«›  «ÕÌ Ê   ·Â «Â·«ﬂ
    
            ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)

            If SystemOptions.UserInterface <> ArabicInterface Then
                des = "Fixed Asset Opening Balance For Asset  " & Me.TxtName & "  And have Depreciation '"
            Else
                des = "»‰«¡ ⁄·Ï —’Ìœ «›  «ÕÌ ··«’· " & Me.TxtName & "    ·Â «Â·«ﬂ '"
            End If
            
            If ModAccounts.AddNewDev(LngDevID, 1, DepitAccount, val(Me.TxtPurchasePrice.text), 0, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , , , , True, val(Me.txtopening_balance_voucher_id.text), val(Me.XPTxtID.text), val(Me.DCGroup.BoundText), val(Me.dcBranch.BoundText), val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
                    
            '„œÌ‰
            '            If ModAccounts.AddNewDev(LngDevID, 1, _
                         DepitAccount, Val(TxtPurchasePrice.text), 0, _
                         des, Val(Me.TxtNoteID), , , _
                         SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
                         , , , , , Val(Me.XPTxtID.text), Val(Me.DCGroup.BoundText), Val(Me.DcBranch.BoundText)) = False Then
            '       GoTo ErrTrap
                    
            '            End If
            '             „Ã„⁄ «·«Â·«ﬂ œ«∆‰ 1
            If val(TxtAccDepreciation.text) > 0 Then
                If ModAccounts.AddNewDev(LngDevID, 2, CreditAccount1, val(Me.TxtAccDepreciation.text), 1, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , , , , True, val(Me.txtopening_balance_voucher_id.text), val(Me.XPTxtID.text), val(Me.DCGroup.BoundText), val(Me.dcBranch.BoundText), val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
                    
                '          If ModAccounts.AddNewDev(LngDevID, 2, _
                           CreditAccount1, Val(TxtAccDepreciation.text), 1, _
                           des, Val(Me.TxtNoteID), , , _
                           SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
                           , , , , , Val(Me.XPTxtID.text), Val(Me.DCGroup.BoundText), Val(Me.DcBranch.BoundText)) = False Then
                '    GoTo ErrTrap
                    
                '            End If
            End If

            '            Ê”Ìÿ «›  «ÕÌ 2
            If val(Me.TxtPurchasePrice.text) - val(Me.TxtAccDepreciation.text) > 0 Then
                If ModAccounts.AddNewDev(LngDevID, 3, CreditAccount2, val(Me.TxtPurchasePrice.text) - val(Me.TxtAccDepreciation.text), 1, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , , , , True, val(Me.txtopening_balance_voucher_id.text), val(Me.XPTxtID.text), val(Me.DCGroup.BoundText), val(Me.dcBranch.BoundText), val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
            End If


'ﬁÌ„… «·«’· ﬂŒ—œ…
      If val(TxtKhordaPrice) > 0 Then
   '             If ModAccounts.AddNewDev(LngDevID, 3, DepitAccount, val(Me.TxtKhordaPrice.text), 1, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , , , , True, val(Me.txtopening_balance_voucher_id.text), val(Me.XPTxtID.text), val(Me.DCGroup.BoundText), val(Me.DcBranch.BoundText), val(Me.DcBranch.BoundText)) = False Then
   '                 GoTo ErrTrap
   '             End If
            End If
            

            '    If ModAccounts.AddNewDev(LngDevID, 3, _
                 CreditAccount2, Val(TxtPurchasePrice.text) - Val(TxtAccDepreciation.text), 1, _
                 des, Val(Me.TxtNoteID), , , _
                 SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
                 , , , , , Val(Me.XPTxtID.text), Val(Me.DCGroup.BoundText), Val(Me.DcBranch.BoundText)) = False Then
            '   GoTo ErrTrap
                    
            '      End If

        ElseIf Option2.value = True And opt(0).value = True And Me.cStatus.ListIndex = 2 Then  '«›  «ÕÌ Ê·Â «Â·«ﬂ  Ê  „   «·«Â·«ﬂ
            '     LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
       
            If SystemOptions.UserInterface <> ArabicInterface Then
                des = "Fixed Asset Fully Depreciation , Name IS  " & Me.TxtName
            Else
                des = "«’· «›  «ÕÌ Ê „ «Â·«ﬂ «·«’· " & Me.TxtName
            End If
            
            If ModAccounts.AddNewDev(LngDevID, 0, DepitAccount, val(Me.TxtKhordaPrice.text), 0, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , , , , True, val(Me.txtopening_balance_voucher_id.text), val(Me.XPTxtID.text), val(Me.DCGroup.BoundText), val(Me.dcBranch.BoundText), val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

            '„œÌ‰
            '            If ModAccounts.AddNewDev(LngDevID, 0, _
                         DepitAccount, Val(TxtKhordaPrice.text), 0, _
                         des, Val(Me.TxtNoteID), , , _
                         SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
                         , , , , , Val(Me.XPTxtID.text), Val(Me.DCGroup.BoundText), Val(Me.dcBranch.BoundText)) = False Then
            '                 GoTo ErrTrap
                    
            '            End If
            If ModAccounts.AddNewDev(LngDevID, 3, CreditAccount2, val(Me.TxtKhordaPrice.text), 1, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , , , , True, val(Me.txtopening_balance_voucher_id.text), val(Me.XPTxtID.text), val(Me.DCGroup.BoundText), val(Me.dcBranch.BoundText), val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
      
            '            œ«∆‰ 1
            '          If ModAccounts.AddNewDev(LngDevID, 1, _
                       CreditAccount2, Val(TxtKhordaPrice.text), 1, _
                       des, Val(Me.TxtNoteID), , , _
                       SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
                       , , , , , Val(Me.XPTxtID.text), Val(Me.DCGroup.BoundText), Val(Me.dcBranch.BoundText)) = False Then
            '                 GoTo ErrTrap
                    
            '            End If

        End If
    End If

    CreateJL = True
    Exit Function
ErrTrap:
    CreateJL = False
End Function

Function GetStatus_id() As Integer
    GetStatus_id = 1
End Function

Function GetDefaultAge() As Integer
    GetDefaultAge = 10
End Function

Function getLastDepreciationDate() As Date
    getLastDepreciationDate = "05-05-2012"
End Function

Function GetInstallmentsInformations(FixedassetId As Integer, Optional ByRef noOfInstallments As Integer, Optional ByRef EXEInstallments As Integer, Optional ByRef RemainInstallments As Integer, Optional ByRef purchaseprice As Double, Optional ByRef KhordaPrice As Double, Optional Depreciation_Percentage As Double, Optional ByRef Installmentvalue As Double)
    noOfInstallments = 10
    EXEInstallments = 4
    RemainInstallments = 6
    Installmentvalue = 700
End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "ID=" & val(XPTxtID.text) & "", , adSearchForward, adBookmarkFirst

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

Function checkEneringPurchaseInvoices() As Boolean
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim msgstr As String
    sql = "select * from DOUBLE_ENTREY_VOUCHERS where  FixedAssetId=" & val(Me.XPTxtID.text)
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            msgstr = " ·« Ì„ﬂ‰ «· ⁄œÌ·  „  ‰›Ì– ›« Ê—… ‘—«¡  ⁄·Ï «·«’·  " & CHR(13)
            msgstr = msgstr & TxtName.text & CHR(13)
            msgstr = msgstr & "«–Â» «·Ï ›Ê« Ì— ‘—«¡ «·«’· ·Õ–› «·›« Ê—… «Ê·« "
            MsgBox msgstr, vbCritical
        Else
            msgstr = " Can't Modify Fixed Asset   " & CHR(13)
            msgstr = msgstr & TxtName.text & CHR(13)
            msgstr = msgstr & " It Have Purchase Invoices "
            MsgBox msgstr, vbCritical
        End If

        Exit Function
        checkEneringPurchaseInvoices = False
    Else
        checkEneringPurchaseInvoices = True
    End If

End Function

Private Sub Del_AssetType()
    Dim msgstr  As String
    Dim noOfInstallments As Integer
    noOfInstallments = CheCkInstallmentCount(val(Me.XPTxtID.text))

    If noOfInstallments > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            msgstr = " ·« Ì„ﬂ‰ «·Õ–›  „  ‰›Ì– «ﬁ”«ÿ ⁄·Ï «·«’·  " & CHR(13)
            msgstr = msgstr & TxtName.text & CHR(13)
            msgstr = msgstr & "⁄œœ «·«ﬁ”«ÿ «·„‰›–… Õ Ï «·«‰ " & noOfInstallments
            MsgBox msgstr, vbCritical
        Else
            msgstr = " Can't Delete Fixed Asset   " & CHR(13)
            msgstr = msgstr & TxtName.text & CHR(13)
            msgstr = msgstr & "No Of Executed Installments " & noOfInstallments
            MsgBox msgstr, vbCritical
        End If

        Exit Sub
    End If

    Dim sql As String
 
    Dim rs2 As New ADODB.Recordset
   sql = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetId"
   sql = sql & "    FROM         dbo.notes_all LEFT OUTER JOIN"
   sql = sql & "                    dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.notes_all.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.notes_all"
   sql = sql & " Where (dbo.notes_all.notetype = 80) And (dbo.notes_all.bill_Type = 2) And (dbo.DOUBLE_ENTREY_VOUCHERS.FixedassetId = " & XPTxtID.text & ")"

    'Sql = "select * from DOUBLE_ENTREY_VOUCHERS where  FixedAssetId=" & val(Me.XPTxtID.Text)
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

     If rs2.RecordCount > 0 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            msgstr = " ·« Ì„ﬂ‰ «·Õ–›  „  ‰›Ì– ›« Ê—… ‘—«¡  ⁄·Ï «·«’·  " & CHR(13)
            msgstr = msgstr & TxtName.text & CHR(13)
            msgstr = msgstr & "«–Â» «·Ï ›Ê« Ì— ‘—«¡ «·«’· ·Õ–› «·›« Ê—… «Ê·« "
            MsgBox msgstr, vbCritical
        Else
            msgstr = " Can't Delete Fixed Asset   " & CHR(13)
            msgstr = msgstr & TxtName.text & CHR(13)
            msgstr = msgstr & " It Have Purchase Invoices "
            MsgBox msgstr, vbCritical
        End If

        Exit Sub
 
    End If

    Dim Msg As String
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (Me.XPTxtID.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                sql = "Delete From TblAssestes Where CarsDataID =" & val(Me.XPTxtID.text) & "  And FlgCarNotFixed = 3"
              Cn.Execute sql, , adExecuteNoRecords
              
                sql = "Delete   from notes where NoteID=" & val(TXTNoteID.text)
                Cn.Execute sql
        
                sql = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
                Cn.Execute sql, , adExecuteNoRecords
                sql = "delete  FixedAssetInstallmentsDetails where FixedAssetID=" & val(Me.XPTxtID.text)
                Cn.Execute sql, , adExecuteNoRecords
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
             
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
 
Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ‰Ê⁄ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·‰Ê⁄ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

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
    'Lbl(27).Caption = "Purcahse Value"
    'Lbl(28).Caption = "Damage Value"
    'Lbl(26).Caption = "Done"
    'Lbl(4).Caption = "Not Done"
    Frame1.Caption = "Group Data"
    RdMove(0).Caption = "Movable"
    RdMove(1).Caption = "Not Movable"
    RdMove(0).RightToLeft = False
    RdMove(1).RightToLeft = False
    'Lbl(24).Caption = "Depreciation %"
    'Lbl(25).Caption = "Stop Depreciation %"

    'Lbl(22).Caption = "F.A Sale Profit"
    'Lbl(23).Caption = "F.A Sale Losses"
'****************************
lbl(20).Caption = "Total Installm."
lbl(15).Caption = "Added Value"
lbl(16).Caption = "Disposed"

lbl(14).Caption = "Opr No"
lbl(12).Caption = "Qty"
lbl(13).Caption = "CC"
ISEQUP.Caption = "IS Equp."
With cStatus
.Clear
.AddItem "Under Depreciation"
.AddItem "Stoped"
.AddItem "Been selling"
.AddItem "Been scrapping"

End With
Cmd(13).Caption = "Copy"
XPTab301.TabCaption(0) = "Basic Data"
XPTab301.TabCaption(1) = "Specific Data"
XPTab301.TabCaption(2) = "Added Data"
Frame4.Caption = XPTab301.TabCaption(1)
lbl(1).Caption = "Board No."
lbl(4).Caption = "Chassis No.."
lbl(2).Caption = "License No.."
lbl(3).Caption = "Model"
lbl(6).Caption = "Origin"
lbl(7).Caption = "Supplier"
lbl(8).Caption = "Examination End"
lbl(10).Caption = "License End"

'****************************

    Me.Caption = "Fixed Assets"
    Me.Ele.Caption = Me.Caption
    Me.lbl(101).Caption = "Code"
    Me.lbl(102).Caption = "Arabic Name"
Me.lbl(11).Caption = "English  Name"
    Option2.Caption = "Opening"
    Option1.Caption = "New"
    lbl(116).Caption = "Purchas Vchr"

    Me.lbl(117).Caption = "Branch"
    Me.lbl(103).Caption = "Group"
    Me.lbl(118).Caption = "Status"

    Me.lbl(104).Caption = "Employee"
    Me.lbl(119).Caption = "Received D"
    Me.lbl(105).Caption = "Depreciation Type"
    Me.lbl(127).Caption = "Start Deprec"
    Me.lbl(120).Caption = "Last Deprec"

    Me.lbl(106).Caption = "Purchase Value"
    Me.lbl(107).Caption = "Current Value"
    Me.lbl(17).Caption = "Current Value"
    
    Me.lbl(108).Caption = "No of installm."
    Me.lbl(130).Caption = "Exec. Inst."
    Me.lbl(123).Caption = "Remains installm."
     Me.lbl(19).Caption = "Remains installm."
     
    Me.lbl(128).Caption = "Purchase Date"
    Me.lbl(121).Caption = "Damage Price"
    Me.lbl(129).Caption = "Acc. deprec "
    Me.lbl(122).Caption = "Installent Value "
    Me.lbl(18).Caption = "Installent Added "
    
    Frame1.Caption = "Groups Data"
    Me.lbl(109).Caption = "Depreciation %"
    Me.lbl(110).Caption = "Stop %"

    opt(0).Caption = "Have Dep"
    opt(1).Caption = "Have't Dep"
 
    Me.lbl(111).Caption = "Asset Account"
    Me.lbl(112).Caption = "Accumulated depreciation Acc."
    Me.lbl(113).Caption = "depreciation Acc. Expenses"
    Me.lbl(114).Caption = "Sale Profit Acc"
    Me.lbl(115).Caption = "Sale Loss Acc"
 
    Label1.Caption = "Sale Price"
    Label3.Caption = "Loss Or Profit"
    lbl(0).Caption = "Ge No"

    Me.lbl(5).Caption = "By"
    Me.lbl(9).Caption = "Default Age"
    Cmd(12).Caption = "Search"
 
    Me.lbl(124).Caption = "Remark"

    Cmd(8).Caption = "Depreciation Restart"
    Cmd(9).Caption = "Asset Disposal"
    Cmd(5).Caption = "Asset Image"
    Cmd(10).Caption = " Show bill"

    Me.lbl(125).Caption = "Current Record:"
    Me.lbl(126).Caption = "Records NO:"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(6).Caption = "Exit"
    Cmd(7).Caption = "Stop Dep"
    Frame7.Caption = "Added Data"
With GridInstallments
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("ID")) = "Trans.No"
.TextMatrix(0, .ColIndex("AddValue")) = "Added Value"
.TextMatrix(0, .ColIndex("DateAdd")) = "Added Date"
.TextMatrix(0, .ColIndex("TypeSand")) = "Type"
.TextMatrix(0, .ColIndex("Show")) = "Show"
End With
    With CBoDepreciation_Type_id
        .Clear
        .AddItem "fixed "
        .AddItem "Decreasing"
    End With

End Sub

 
