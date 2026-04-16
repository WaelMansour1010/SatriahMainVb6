VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmShareholders 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·„”«Â„Ì‰"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13980
   HelpContextID   =   50
   Icon            =   "FrmShareholders.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   13980
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox DcbType 
      Height          =   315
      ItemData        =   "FrmShareholders.frx":038A
      Left            =   2880
      List            =   "FrmShareholders.frx":038C
      RightToLeft     =   -1  'True
      TabIndex        =   232
      Top             =   600
      Width           =   1515
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   795
      Left            =   60
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   213
      Top             =   2520
      Width           =   8505
   End
   Begin VB.TextBox TxtIBAN 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   211
      Top             =   1800
      Width           =   3435
   End
   Begin VB.ComboBox CboSaleType 
      Height          =   315
      Left            =   5520
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   200
      Top             =   10200
      Width           =   3015
   End
   Begin VB.TextBox TxtFullcode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   14400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   191
      Top             =   960
      Width           =   1125
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   2880
      TabIndex        =   175
      Top             =   10440
      Width           =   7935
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   5
         Left            =   1440
         Picture         =   "FrmShareholders.frx":038E
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   4
         Left            =   2880
         Picture         =   "FrmShareholders.frx":07D9
         Style           =   1  'Graphical
         TabIndex        =   186
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   3
         Left            =   3600
         Picture         =   "FrmShareholders.frx":0D31
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   2
         Left            =   5040
         Picture         =   "FrmShareholders.frx":11EA
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   1
         Left            =   3240
         Picture         =   "FrmShareholders.frx":16BA
         Style           =   1  'Graphical
         TabIndex        =   183
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmShareholders.frx":1B5B
         Height          =   555
         Index           =   0
         Left            =   7200
         Picture         =   "FrmShareholders.frx":8E8D
         Style           =   1  'Graphical
         TabIndex        =   182
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmShareholders.frx":9434
         Height          =   555
         Index           =   6
         Left            =   5760
         Picture         =   "FrmShareholders.frx":10766
         Style           =   1  'Graphical
         TabIndex        =   181
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmShareholders.frx":10C07
         Height          =   555
         Index           =   7
         Left            =   4320
         Picture         =   "FrmShareholders.frx":17F39
         Style           =   1  'Graphical
         TabIndex        =   180
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   8
         Left            =   2160
         Picture         =   "FrmShareholders.frx":187C9
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmShareholders.frx":18CAE
         Height          =   555
         Index           =   9
         Left            =   720
         Picture         =   "FrmShareholders.frx":1FFE0
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmShareholders.frx":20500
         Height          =   555
         Index           =   10
         Left            =   6480
         Picture         =   "FrmShareholders.frx":20AE7
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmShareholders.frx":210CE
         Height          =   555
         Index           =   11
         Left            =   0
         Picture         =   "FrmShareholders.frx":28400
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   5520
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox XPMTxtRemarks2 
         Alignment       =   1  'Right Justify
         Height          =   795
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   480
         Width           =   5145
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”»» «·«Ìﬁ«›"
         Height          =   285
         Index           =   32
         Left            =   1950
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Œ’Ê„«  Œ«’… ··„”«Â„ ›Ï ›Ê« Ì— «·‘—«¡"
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
      Height          =   1125
      Index           =   6
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   119
      Top             =   9120
      Width           =   3045
      Begin VB.ComboBox CboDiscountTypePur 
         Height          =   315
         Left            =   270
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtDiscountValuePur 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         RightToLeft     =   -1  'True
         TabIndex        =   120
         Top             =   600
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·Œ’„"
         Height          =   285
         Index           =   29
         Left            =   1770
         RightToLeft     =   -1  'True
         TabIndex        =   124
         Top             =   270
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·Œ’„"
         Height          =   285
         Index           =   28
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   123
         Top             =   660
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   122
         Top             =   690
         Width           =   195
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Œ’Ê„«  Œ«’… ··„”«Â„ ›Ï ›Ê« Ì— «·»Ì⁄"
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
      Height          =   1125
      Index           =   4
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   113
      Top             =   8880
      Width           =   2925
      Begin VB.TextBox TxtDiscountValue 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         RightToLeft     =   -1  'True
         TabIndex        =   115
         Top             =   600
         Width           =   1425
      End
      Begin VB.ComboBox CboDiscountType 
         Height          =   315
         Left            =   270
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   118
         Top             =   690
         Width           =   195
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·Œ’„"
         Height          =   285
         Index           =   20
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   117
         Top             =   660
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·Œ’„"
         Height          =   285
         Index           =   19
         Left            =   1770
         RightToLeft     =   -1  'True
         TabIndex        =   116
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.TextBox txtCustGID 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2880
      MaxLength       =   50
      TabIndex        =   57
      Top             =   960
      Width           =   1485
   End
   Begin VB.CheckBox chkCustomerandVendor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄„Ì· Ê„Ê—œ"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   9840
      Width           =   1815
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10830
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   600
      Width           =   1605
   End
   Begin VB.TextBox txtopening_balance_voucher_id 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox XPTxtCusNamee 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5520
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1020
      Width           =   3045
   End
   Begin VB.TextBox c2 
      Height          =   345
      Left            =   600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox c1 
      Height          =   345
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   1
      Top             =   -120
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„«  ⁄‰ «·„”«Â„"
      Height          =   555
      Left            =   3750
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1155
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·›Ê« Ì—"
         Height          =   285
         Index           =   18
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1770
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÕÃ„ «·„»Ì⁄«  «· Ã«—Ï"
         Height          =   285
         Index           =   17
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1470
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÕÃ„ «·„»Ì⁄«  «·ﬁÿ«⁄Ï"
         Height          =   285
         Index           =   16
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1170
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÕÃ„ «·„»Ì⁄«  «·√Ã·…"
         Height          =   285
         Index           =   15
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   870
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÕÃ„ «·„»Ì⁄«  «·‰ﬁœÌ…"
         Height          =   285
         Index           =   14
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÕÃ„ „»Ì⁄« Â"
         Height          =   285
         Index           =   13
         Left            =   2190
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·≈ ’«·"
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
      Height          =   2025
      Index           =   3
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1410
      Width           =   4365
      Begin VB.TextBox txtJob 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   197
         Top             =   600
         Width           =   2805
      End
      Begin VB.TextBox TxtResponsibleContact 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   210
         Width           =   2805
      End
      Begin VB.TextBox TxtFaxNumber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1710
         Width           =   2805
      End
      Begin VB.TextBox XPTxtPhone 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1020
         Width           =   2805
      End
      Begin VB.TextBox XPTxtmobile 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1365
         Width           =   2805
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›…"
         Height          =   315
         Index           =   69
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   198
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„”∆Ê· «·≈ ’«·"
         Height          =   315
         Index           =   23
         Left            =   2730
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·›«ﬂ”"
         Height          =   315
         Index           =   4
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1590
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÃÊ«·"
         Height          =   255
         Index           =   2
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1275
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Â« ›"
         Height          =   315
         Index           =   3
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   930
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      Height          =   345
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   90
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox XPTxtCusName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   9600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1020
      Width           =   2805
   End
   Begin VB.TextBox XPTxtCusID 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   6030
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   1455
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   705
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   -120
      Width           =   13995
      _cx             =   24686
      _cy             =   1244
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
      Caption         =   "»Ì«‰«  «·„”«Â„Ì‰  "
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
         Height          =   345
         Index           =   0
         Left            =   1185
         TabIndex        =   11
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
         ButtonImage     =   "FrmShareholders.frx":28F94
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
         Left            =   120
         TabIndex        =   13
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
         ButtonImage     =   "FrmShareholders.frx":2932E
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
         Left            =   1710
         TabIndex        =   10
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
         ButtonImage     =   "FrmShareholders.frx":296C8
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
         Left            =   645
         TabIndex        =   12
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
         ButtonImage     =   "FrmShareholders.frx":29A62
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin VB.Image Img 
         Height          =   480
         Left            =   7080
         Picture         =   "FrmShareholders.frx":29DFC
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   11355
      TabIndex        =   20
      Top             =   7710
      Width           =   705
      _ExtentX        =   1244
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   10635
      TabIndex        =   21
      Top             =   7710
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   9915
      TabIndex        =   22
      Top             =   7710
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   3
      Left            =   9195
      TabIndex        =   23
      Top             =   7710
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   4
      Left            =   8355
      TabIndex        =   24
      Top             =   7710
      Width           =   825
      _ExtentX        =   1455
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   5
      Left            =   7530
      TabIndex        =   25
      Top             =   7710
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   6
      Left            =   4920
      TabIndex        =   26
      Top             =   7710
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   6690
      TabIndex        =   27
      Top             =   7710
      Width           =   795
      _ExtentX        =   1402
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   375
      Left            =   5670
      TabIndex        =   28
      Top             =   7710
      Width           =   975
      _ExtentX        =   1720
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
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   285
      Index           =   8
      Left            =   -660
      TabIndex        =   29
      Top             =   900
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ›Ÿ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DrawFocusRectangle=   0   'False
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   9600
      TabIndex        =   52
      Top             =   600
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   5520
      TabIndex        =   54
      Top             =   600
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   58
      Top             =   3360
      Width           =   13905
      _cx             =   24527
      _cy             =   7223
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   128
      FrontTabColor   =   14871017
      BackTabColor    =   8454143
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "»Ì«‰«  «”«”Ì…|»Ì«‰«  „ Œ’’Â"
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3675
         Left            =   14550
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   45
         Width           =   13815
         _cx             =   24368
         _cy             =   6482
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
         GridRows        =   10
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
            Caption         =   "»Ì«‰«  Œ«’… »«·œÂ»"
            ForeColor       =   &H000000FF&
            Height          =   3015
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   0
            Width           =   5175
            Begin VB.TextBox TxtbalancedC 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   360
               Locked          =   -1  'True
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   170
               Top             =   2520
               Width           =   1425
            End
            Begin VB.TextBox Txtbalanced 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   2280
               Locked          =   -1  'True
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   169
               Top             =   2520
               Width           =   1425
            End
            Begin VB.TextBox txtTotalc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               Height          =   315
               Left            =   360
               Locked          =   -1  'True
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   2160
               Width           =   1425
            End
            Begin VB.TextBox txtTotald 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               Height          =   315
               Left            =   2280
               Locked          =   -1  'True
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   2160
               Width           =   1425
            End
            Begin VB.TextBox TxtShowQty1c 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   360
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice1C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   165
               Top             =   720
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice2C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   1080
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries1C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   163
               Top             =   1440
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries2C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   1800
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   159
               Top             =   1800
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   156
               Top             =   1440
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   154
               Top             =   1080
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   720
               Width           =   1425
            End
            Begin VB.TextBox TxtShowQty1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   360
               Width           =   1425
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·—’Ìœ"
               Height          =   255
               Index           =   59
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   172
               Top             =   2520
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì"
               Height          =   255
               Index           =   58
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "œ«∆‰"
               Height          =   255
               Index           =   57
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„œÌ‰"
               Height          =   255
               Index           =   56
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«ÃÊ— ‰—ﬂÌ»"
               Height          =   255
               Index           =   55
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«ÃÊ— ’Ì«€Â"
               Height          =   255
               Index           =   54
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   157
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·«·„«”"
               Height          =   255
               Index           =   53
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·œÂ»"
               Height          =   255
               Index           =   52
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Ê“‰ «·œÂ»"
               Height          =   255
               Index           =   51
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.CommandButton CDMOldContract 
            Caption         =   "⁄ﬁÊœ ”«»ﬁ…"
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·⁄„·"
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
            Height          =   3885
            Index           =   7
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   240
            Width           =   6135
            Begin VB.ComboBox CboSex 
               Height          =   315
               ItemData        =   "FrmShareholders.frx":2AAC6
               Left            =   3000
               List            =   "FrmShareholders.frx":2AAC8
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   106
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TxtMobile2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3000
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   3480
               Width           =   1485
            End
            Begin VB.TextBox TxtMobile1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   3120
               Width           =   1485
            End
            Begin VB.TextBox TxtHomeTel 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3000
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   3120
               Width           =   1485
            End
            Begin VB.TextBox TxtJobTelConvert 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   2760
               Width           =   1485
            End
            Begin VB.TextBox TxtJobTel 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3000
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   2760
               Width           =   1485
            End
            Begin VB.TextBox TxtJobAddress 
               Alignment       =   1  'Right Justify
               Height          =   555
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   92
               Top             =   2160
               Width           =   4425
            End
            Begin VB.TextBox TxtSalary 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2400
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   1680
               Width           =   2085
            End
            Begin VB.TextBox TXTJobTitle 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   1320
               Width           =   4365
            End
            Begin VB.TextBox TxtCompany 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   960
               Width           =   4365
            End
            Begin MSDataListLib.DataCombo DcboCountryID2 
               Height          =   315
               Left            =   1440
               TabIndex        =   107
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ã‰”"
               Height          =   315
               Index           =   46
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ã‰”Ì…"
               Height          =   315
               Index           =   45
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ ÃÊ«· «Œ— "
               Height          =   315
               Index           =   44
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   3480
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÃÊ«·"
               Height          =   315
               Index           =   43
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   3120
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Â« › «·„‰“·"
               Height          =   315
               Index           =   42
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   3120
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕÊÌ·Â"
               Height          =   315
               Index           =   41
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   2880
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Â« › «·⁄„·"
               Height          =   315
               Index           =   40
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   2760
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄‰Ê«‰ «·⁄„·"
               Height          =   315
               Index           =   39
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   2280
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ﬁœ«— «·—« »"
               Height          =   315
               Index           =   38
               Left            =   4620
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   1710
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„”„Ï «·ÊŸÌ›Ì"
               Height          =   315
               Index           =   37
               Left            =   4620
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   1350
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÃÂ… «·⁄„·"
               Height          =   315
               Index           =   36
               Left            =   4500
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   870
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   255
               Index           =   35
               Left            =   1290
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   240
               Width           =   825
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3675
         Index           =   2
         Left            =   45
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   45
         Width           =   13815
         _cx             =   24368
         _cy             =   6482
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
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄‰Ê«‰"
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
            Height          =   1905
            Index           =   5
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   222
            Top             =   1560
            Width           =   4065
            Begin VB.TextBox TxtAddress 
               Alignment       =   1  'Right Justify
               Height          =   705
               Left            =   150
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   223
               Top             =   1140
               Width           =   2625
            End
            Begin MSDataListLib.DataCombo DcboCountryID 
               Height          =   315
               Left            =   150
               TabIndex        =   224
               Top             =   150
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboGovernmentID 
               Height          =   315
               Left            =   150
               TabIndex        =   225
               Top             =   480
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCityID 
               Height          =   315
               Left            =   150
               TabIndex        =   226
               Top             =   810
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄‰Ê«‰ »«· ›’Ì·"
               Height          =   585
               Index           =   26
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   230
               Top             =   1140
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÕÌ"
               Height          =   225
               Index           =   25
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   229
               Top             =   840
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Õ«›Ÿ…"
               Height          =   225
               Index           =   24
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   228
               Top             =   510
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·œÊ·…"
               Height          =   225
               Index           =   22
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   227
               Top             =   210
               Width           =   765
            End
         End
         Begin VB.TextBox TxtE_mail 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   480
            Width           =   2715
         End
         Begin VB.TextBox TxtBox 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   840
            Width           =   2715
         End
         Begin VB.TextBox TxtZib 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   215
            Top             =   1200
            Width           =   2715
         End
         Begin VB.TextBox TxtEntry 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9480
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   120
            Width           =   2715
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3315
            Index           =   1
            Left            =   120
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   0
            Width           =   9225
            _cx             =   16272
            _cy             =   5847
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
            Begin VB.CheckBox creditlocked 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·€«¡ «· ⁄«„· «·«Ã·"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   8040
               RightToLeft     =   -1  'True
               TabIndex        =   204
               Top             =   7680
               Width           =   1665
            End
            Begin VB.CheckBox locked 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ìﬁ«› «· ⁄«„·"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   7305
               RightToLeft     =   -1  'True
               TabIndex        =   202
               Top             =   1800
               Width           =   1320
            End
            Begin VB.TextBox TxtMap 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   600
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   193
               Top             =   6720
               Width           =   4575
            End
            Begin VB.ComboBox DcbDigCustomer 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   9930
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   174
               Top             =   6840
               Width           =   2715
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "—’Ìœ „œ›Ê⁄«  „ﬁœ„…   "
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
               Index           =   9
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   360
               Width           =   2685
               Begin VB.TextBox TxtOpenBalance2 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   510
                  Width           =   1365
               End
               Begin VB.OptionButton OptType2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   1005
               End
               Begin VB.OptionButton OptType2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   1080
                  RightToLeft     =   -1  'True
                  TabIndex        =   143
                  Top             =   210
                  Width           =   765
               End
               Begin VB.OptionButton OptType2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   1950
                  RightToLeft     =   -1  'True
                  TabIndex        =   142
                  Top             =   210
                  Width           =   765
               End
               Begin MSComCtl2.DTPicker Dtp2 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   146
                  Top             =   870
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   75235331
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   255
                  Index           =   50
                  Left            =   1260
                  RightToLeft     =   -1  'True
                  TabIndex        =   148
                  Top             =   540
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   285
                  Index           =   49
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   930
                  Width           =   1215
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "—’Ìœ ‘Ìﬂ«   Õ  «· Õ’Ì· "
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
               Index           =   8
               Left            =   2940
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   360
               Width           =   2715
               Begin VB.OptionButton OptType1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   1950
                  RightToLeft     =   -1  'True
                  TabIndex        =   137
                  Top             =   210
                  Width           =   765
               End
               Begin VB.OptionButton OptType1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   1080
                  RightToLeft     =   -1  'True
                  TabIndex        =   136
                  Top             =   210
                  Width           =   765
               End
               Begin VB.OptionButton OptType1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   1005
               End
               Begin VB.TextBox TxtOpenBalance1 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   134
                  Top             =   510
                  Width           =   1365
               End
               Begin MSComCtl2.DTPicker Dtp1 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   138
                  Top             =   870
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   75235331
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   285
                  Index           =   48
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   930
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   255
                  Index           =   47
                  Left            =   1260
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   540
                  Width           =   1275
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·—’Ìœ «·√›  «ÕÏ «·Ã«—Ì"
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
               Left            =   5910
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   360
               Width           =   2685
               Begin VB.TextBox TxtOpenBalance 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   510
                  Width           =   1365
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   128
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   1005
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   1080
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   210
                  Width           =   765
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   1950
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   210
                  Width           =   765
               End
               Begin MSComCtl2.DTPicker Dtp 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   130
                  Top             =   870
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   75235331
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   255
                  Index           =   5
                  Left            =   1260
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   540
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   285
                  Index           =   6
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   930
                  Width           =   1215
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  „Õ«”»Ì…"
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
               Height          =   1005
               Index           =   0
               Left            =   4020
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   5640
               Width           =   5805
               Begin VB.TextBox TxtCreditlimitCredit 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2670
                  MaxLength       =   8
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   540
                  Width           =   1395
               End
               Begin VB.TextBox TxtCreditLimit 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2670
                  MaxLength       =   8
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   180
                  Width           =   1395
               End
               Begin VB.TextBox TxtDepitInterval 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   240
                  Width           =   495
               End
               Begin VB.TextBox TxtCreditInterval 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   600
                  Width           =   495
               End
               Begin VB.ComboBox dcDepitIntervalID 
                  Height          =   315
                  ItemData        =   "FrmShareholders.frx":2AACA
                  Left            =   120
                  List            =   "FrmShareholders.frx":2AACC
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   240
                  Width           =   975
               End
               Begin VB.ComboBox dcCreditIntervalID 
                  Height          =   315
                  ItemData        =   "FrmShareholders.frx":2AACE
                  Left            =   120
                  List            =   "FrmShareholders.frx":2AAD0
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœ «·√∆ „«‰(œ«∆‰)"
                  Height          =   285
                  Index           =   11
                  Left            =   4230
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   540
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœ «·√∆ „«‰(„œÌ‰)"
                  Height          =   285
                  Index           =   7
                  Left            =   4230
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   180
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÂ «·«∆ „«‰"
                  Height          =   285
                  Index           =   30
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   240
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÂ «·«∆ „«‰"
                  Height          =   285
                  Index           =   31
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   600
                  Width           =   885
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "—’Ìœ «·„”«Â„ «·Õ«·Ï"
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
               Height          =   645
               Index           =   2
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   2160
               Width           =   3315
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   435
                  Index           =   9
                  Left            =   120
                  TabIndex        =   66
                  Top             =   150
                  Width           =   1185
                  _ExtentX        =   2090
                  _ExtentY        =   767
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷  ﬁ—Ì— ﬂ‘› Õ”«»"
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
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Height          =   255
                  Index           =   9
                  Left            =   1290
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   240
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   255
                  Index           =   8
                  Left            =   1410
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   240
                  Width           =   1785
               End
            End
            Begin VB.TextBox txtidffff 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   -3840
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   9360
               Width           =   2130
            End
            Begin MSDataListLib.DataCombo DboParentAccount 
               Height          =   315
               Left            =   240
               TabIndex        =   80
               Top             =   1800
               Width           =   4605
               _ExtentX        =   8123
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ALLButtonS.ALLButton ALLButton1 
               Height          =   255
               Left            =   6120
               TabIndex        =   199
               Top             =   1800
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   "«·”»»"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmShareholders.frx":2AAD2
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Œ—«∆ÿ ÃÊÃ·"
               Height          =   345
               Index           =   67
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   196
               Top             =   6720
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   285
               Index           =   66
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   195
               Top             =   8400
               Width           =   585
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   285
               Index           =   65
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   194
               Top             =   21600
               Width           =   585
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·‰ÃÊ„"
               Height          =   285
               Index           =   62
               Left            =   12750
               RightToLeft     =   -1  'True
               TabIndex        =   173
               Top             =   6840
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ”«» «·—∆Ì”Ì"
               Height          =   315
               Index           =   33
               Left            =   4710
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   1800
               Width           =   1365
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   13485
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   960
               Width           =   810
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»—Ìœ «·≈·ﬂ —Ê‰Ï"
            Height          =   285
            Index           =   12
            Left            =   12330
            RightToLeft     =   -1  'True
            TabIndex        =   221
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’‰œÊﬁ »—Ìœ"
            Height          =   285
            Index           =   60
            Left            =   12300
            RightToLeft     =   -1  'True
            TabIndex        =   220
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—„“ «·»—ÌœÌ"
            Height          =   285
            Index           =   61
            Left            =   12300
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "œ«Œ·Ì"
            Height          =   315
            Index           =   68
            Left            =   12330
            RightToLeft     =   -1  'True
            TabIndex        =   218
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   315
            Index           =   34
            Left            =   8400
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   6330
            Width           =   1125
         End
      End
   End
   Begin Dynamic_Byte.NourHijriCal Txt_DateExpLincH 
      Height          =   345
      Left            =   0
      TabIndex        =   82
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
   End
   Begin MSDataListLib.DataCombo DcCustomerType 
      Height          =   315
      Left            =   5520
      TabIndex        =   109
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   12480
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCEmP 
      Height          =   315
      Left            =   0
      TabIndex        =   111
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   9240
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   10
      Left            =   2640
      TabIndex        =   188
      Top             =   7710
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿÌ«⁄Â"
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin MSComCtl2.DTPicker DtRecord 
      Height          =   330
      Left            =   360
      TabIndex        =   189
      Top             =   600
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      CalendarBackColor=   12648447
      CustomFormat    =   "yyyy/M/d"
      Format          =   75235331
      CurrentDate     =   38718
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   11
      Left            =   4080
      TabIndex        =   203
      Top             =   7710
      Width           =   765
      _ExtentX        =   1349
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
   Begin MSDataListLib.DataCombo DcbTypeInvestor 
      Height          =   315
      Left            =   5490
      TabIndex        =   205
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   1440
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbGroupInvestor 
      Height          =   315
      Left            =   120
      TabIndex        =   207
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   1440
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbBank 
      Height          =   315
      Left            =   5490
      TabIndex        =   209
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   1800
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„”«Â„"
      Height          =   285
      Index           =   70
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   233
      Top             =   630
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   285
      Index           =   1
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   231
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ Õ”«» «·»‰ﬂ"
      Height          =   285
      Index           =   6
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   212
      Top             =   1800
      Width           =   2010
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·»‰ﬂ"
      Height          =   285
      Index           =   5
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   210
      Top             =   1800
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã„Ê⁄… «·„” À„—"
      Height          =   285
      Index           =   4
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   208
      Top             =   1440
      Width           =   1890
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„” À„—"
      Height          =   285
      Index           =   3
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   206
      Top             =   1560
      Width           =   1890
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "”Ì«”… «·»Ì⁄ "
      Height          =   405
      Index           =   10
      Left            =   7860
      RightToLeft     =   -1  'True
      TabIndex        =   201
      Top             =   10230
      Width           =   1665
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   285
      Index           =   64
      Left            =   12870
      RightToLeft     =   -1  'True
      TabIndex        =   192
      Top             =   6600
      Width           =   585
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   285
      Index           =   63
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   190
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„‰œÊ»"
      Height          =   285
      Index           =   1
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   112
      Top             =   9240
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„”«Â„"
      Height          =   285
      Index           =   2
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   110
      Top             =   12480
      Width           =   1890
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·«‰ Â«¡"
      Height          =   345
      Index           =   12
      Left            =   1830
      TabIndex        =   83
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”Ã·"
      Height          =   345
      Index           =   5
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   1080
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8505
      TabIndex        =   55
      Top             =   600
      Width           =   690
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ «‰Ã·Ì“Ì"
      Height          =   255
      Index           =   4
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·⁄„Ì· «·‰Â«∆Ì"
      Height          =   315
      Index           =   3
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   360
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„Ì·"
      Height          =   315
      Index           =   2
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   420
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   750
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   7740
      Width           =   165
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ ⁄—»Ì"
      Height          =   375
      Index           =   0
      Left            =   12930
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1005
      Width           =   915
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ "
      Height          =   315
      Index           =   1
      Left            =   12930
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   660
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   0
      Left            =   1470
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   7740
      Width           =   1215
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      Height          =   285
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   7740
      Width           =   465
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      Height          =   285
      Left            =   210
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   7740
      Width           =   615
   End
End
Attribute VB_Name = "FrmShareholders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim CusReport As ClsCustemerReport
Dim Dcombo As ClsDataCombos
Dim cSearch(2) As clsDCboSearch
Dim FirstPeriodDateInthisYear  As Date
Public calledFromForm As Boolean
Private m_DealingForm As GridTransType
      Dim Account_Code_dynamic As String
Private m_DcboCustomers As DataCombo
Public Property Get DcboCustomers() As DataCombo
    Set DcboCustomers = m_DcboCustomers
End Property

Public Property Set DcboCustomers(ByVal vNewValue As DataCombo)
    Set m_DcboCustomers = vNewValue
End Property
Public Property Get DealingForm() As GridTransType
    DealingForm = m_DealingForm
End Property

Public Property Let DealingForm(ByVal vNewValue As GridTransType)
    'If vNewValue = OpeningBalance Or vNewValue = PurchaseTransaction Or vNewValue = InvoiceTransaction Then
    m_DealingForm = vNewValue
    'End If
End Property

 Function ChekShare(Optional SharID As Double = 0) As Boolean
If SharID <> 0 Then
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = "Select * from TblIPOSharer where SharID=" & SharID & ""
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
ChekShare = True
Else
ChekShare = False
End If
End If
End Function

'Private Sub PassData()
'    Dim StrSQL As String
'    On Error GoTo ErrTrap
'    StrSQL = "SELECT * From TblCustemers"
' If Me.calledFromForm = False Then Exit Sub
'    Select Case Me.DealingForm
'
'
'        Case InvoiceTransaction
'            fill_combo Me.DcboCustomers, StrSQL
'            Me.DcboCustomers.BoundText = val(XPTxtCusID.text)
'
'        Case PriceList
'    ' StrSQL = "SELECT * From TblCustemers where Type=2"
'      '   fill_combo FrmMainPriceList.DBCboSupplierName, StrSQL
'      '       FrmMainPriceList.DBCboSupplierName.BoundText = val(XPTxtCusID.text)
'
'            '⁄—÷ «·√”⁄«—
'        Case ShowPrice
'        '   fill_combo FrmShowPrice.DBCboClientName, StrSQL
'        '     FrmShowPrice.DBCboClientName.BoundText = val(XPTxtCusID.text)
'    End Select
' Me.calledFromForm = False
'Unload Me
'    Exit Sub
'ErrTrap:
'End Sub

Private Sub ALLButton1_Click()
    Frame2.Visible = True
End Sub

Private Sub CboDiscountType_Change()
    Me.lbl(21).Visible = (Me.CboDiscountType.ListIndex = 2)

    If CboDiscountType.ListIndex = 0 Then
        lbl(20).Visible = False
        TxtDiscountValue.Visible = False
        lbl(21).Visible = False
    Else
        lbl(20).Visible = True
        TxtDiscountValue.Visible = True
        lbl(21).Visible = True
    End If

End Sub

Private Sub CboDiscountType_Click()
    CboDiscountType_Change
End Sub

Private Sub CboDiscountTypePur_Change()
    Me.lbl(27).Visible = (Me.CboDiscountTypePur.ListIndex = 2)

    If CboDiscountTypePur.ListIndex = 0 Then
        lbl(28).Visible = False
        TxtDiscountValuePur.Visible = False
        lbl(27).Visible = False
    Else
        lbl(28).Visible = True
        TxtDiscountValuePur.Visible = True
        lbl(27).Visible = True
    End If

End Sub

Private Sub CboDiscountTypePur_Click()
    CboDiscountTypePur_Change
End Sub

Function DeleteOpeningBalance()
    Cmd_Click (1)
    OptType(2).value = True
    TxtOpenBalance.Text = 0
    Cmd_Click (2)

End Function

Private Sub CDMOldContract_Click()
FrmOldContract.show
End Sub

Private Sub Cmd_Click(Index As Integer)

    On Error GoTo ErrTrap

    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

    Me.Dtp.value = FirstPeriodDateInthisYear
    Me.Dtp1.value = FirstPeriodDateInthisYear
       Me.Dtp2.value = FirstPeriodDateInthisYear
       
    Dim Msg As String

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
    DcbType.ListIndex = 0
            Txt_DateExpLincH.value = ToHijriDate(Date)
                
      
                If DcbType.ListIndex = 0 Then
               
            Account_Code_dynamic = get_account_code_branch(108, 0)
     Else
     Account_Code_dynamic = get_account_code_branch(110, 0)
     End If
     
          
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
                
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·„”«Â„Ì‰   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
       
                End If
            End If
        
            DboParentAccount.BoundText = Account_Code_dynamic
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

            Me.Dtp.value = FirstPeriodDateInthisYear
            Me.Dcbranch.BoundText = Current_branch
            OptType(2).value = True
    OptType1(2).value = True
        OptType2(2).value = True

                                               
   Dim EmpID As Integer
 
   If SystemOptions.usertype <> UserAdminAll Then
  
  GetUserData user_id, , , , , , EmpID
        Me.DCEmP.BoundText = EmpID
 
  End If
  
  lbl(8).Caption = 0
        Case 1
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

            Me.Dtp.value = FirstPeriodDateInthisYear
 
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            If XPTxtCusID.Text = 2 Then
          '      Msg = "·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰«  Â–« «·”Ã·"
          '      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          '      Exit Sub
            End If

            TxtModFlg.Text = "E"

        Case 2

            Dim currentcode As String

            If txtid.Text = "" Then
                currentcode = get_coding(Current_branch, "TblCustemers", 9, Me.DCPreFix.Text)

                If currentcode = "miniError" Then
                    MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
                    Exit Sub
            
                ElseIf currentcode = "Manual" Then
                    MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·ÕﬁÊ·"
                    Exit Sub
                Else
                    txtid = currentcode
                End If
            End If

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            If XPTxtCusID.Text = 2 Then
                Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·”Ã·"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If
If ChekShare(val(Me.XPTxtCusID.Text)) = False Then
            Del_Member
    Else
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ Õ–› Â–« «·„”«Â„ ·«— »«ÿÂ »⁄„·Ì… "
    Else
   
    MsgBox "The Shareholder can not be deleted because it is linked up process"
    End If
    Exit Sub
    End If
    
        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
        FrmCustemerSearch.SearchType = 20
            FrmCustemerSearch.show vbModal
        
        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            
print_report2
            '        Text1.text = 2
          '  PrintReport

        Case 9

            ' If DoPremis(Do_Print, "ReportCustomers", True) = False Then
            'Exit Sub
            'End If
            '     ShowCusBalance
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
        
            Dim FirstPeriod As Date
            getFirstPeriodDateInthisYear FirstPeriod
            ShowReport IIf(IsNull(rs("Account_Code")), "", rs("Account_Code")), XPTxtCusName.Text, FirstPeriod, Date
        
        
        
        Case 10
            If val(Me.XPTxtCusID.Text) <> 0 Then
                print_report val(Me.XPTxtCusID.Text)
        '" & val(XPTxtCusID.text) & ")"
 
            End If
        
               Case 11
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments DCPreFix.Text & txtid.Text, "0701201401"
 
 
 
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub DboParentAccount_KeyUp(KeyCode As Integer, _
                                   Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 2116
    End If

End Sub

Private Sub DcboCityID_Change()
    LoadDataCombos False, False, True
End Sub

Private Sub DcboCityID_Click(Area As Integer)
    DcboCityID_Change
End Sub

Private Sub DcboCountryID_Change()
    LoadDataCombos True, False, False
End Sub

Private Sub DcboCountryID_Click(Area As Integer)

    If val(Me.DcboCountryID.BoundText) <> 0 Then
        DcboCountryID_Change
    End If

End Sub

Private Sub DcboGovernmentID_Change()
    LoadDataCombos False, True, False
End Sub

Private Sub DcboGovernmentID_Click(Area As Integer)
    DcboGovernmentID_Change
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        LoadDataCombos
    End If

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
MySQL = " SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.Cus_Phone, "
MySQL = MySQL & "                      dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Remark, dbo.TblCustemers.Type, dbo.TblCustemers.OpenBalance, dbo.TblCustemers.OpenBalanceType,"
MySQL = MySQL & "                      dbo.TblCustemers.OpenBalanceDate, dbo.TblCustemers.CreditLimit, dbo.TblCustemers.Account_Code_As_Client, dbo.TblCustemers.Account_Code_As_Supplier,"
MySQL = MySQL & "                      dbo.TblCustemers.CreditlimitCredit, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail, dbo.TblCustemers.SaleType, dbo.TblCustemers.Account_Code,"
MySQL = MySQL & "                      dbo.TblCustemers.Trans_Discount, dbo.TblCustemers.Trans_DiscountType, dbo.TblCustemers.CountryID, dbo.TblCountriesData.CountryName,"
MySQL = MySQL & "                      dbo.TblCustemers.CityID, dbo.TblCountriesGovernmentsCities.CityName, dbo.TblCustemers.GovernmentID, dbo.TblCountriesGovernments.GovernmentName,"
MySQL = MySQL & "                      dbo.TblCustemers.Address, dbo.TblCustemers.Trans_DiscountPur, dbo.TblCustemers.Trans_DiscountTypePur, dbo.TblCustemers.CountEmp, dbo.TblCustemers.ToTal,"
MySQL = MySQL & "                       dbo.TblCustemers.c1, dbo.TblCustemers.c2, dbo.TblCustemers.Remark2, dbo.TblCustemers.locked, dbo.TblCustemers.parent_account,"
MySQL = MySQL & "                      dbo.TblCustemers.opening_balance_voucher_id, dbo.TblCustemers.DepitInterval, dbo.TblCustemers.CreditInterval, dbo.TblCustemers.DepitIntervalID,"
MySQL = MySQL & "                      dbo.TblCustemers.CreditIntervalID, dbo.TblCustemers.EmpId, dbo.TblCustemers.prifix, dbo.TblCustemers.code, dbo.TblCustemers.Fullcode,"
MySQL = MySQL & "                      dbo.TblCustemers.CustomerandVendor, dbo.TblCustemers.CustomerTypeID, dbo.TblCustemers.BranchId, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblCustemers.CustGID, dbo.TblCustemers.ExpireDateH, dbo.TblCustemers.Company, dbo.TblCustemers.JobTitle,"
MySQL = MySQL & "                      dbo.TblCustemers.Salary, dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.HomeTel,"
MySQL = MySQL & "                      dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2, dbo.TblCustemers.CountryID2, dbo.TblCustemers.Sex, dbo.TblCustemers.Account_Code1,"
MySQL = MySQL & "                      dbo.TblCustemers.Account_Code2, dbo.TblCustemers.ParentAccount, dbo.TblCustemers.OpenBalanceType1, dbo.TblCustemers.OpenBalance1,"
MySQL = MySQL & "                      dbo.TblCustemers.OpenBalanceType2, dbo.TblCustemers.OpenBalance2, dbo.TblCustemers.ShowQty1, dbo.TblCustemers.showPrice1,"
MySQL = MySQL & "                      dbo.TblCustemers.showPrice2, dbo.TblCustemers.Salaries1, dbo.TblCustemers.Salaries2, dbo.TblCustemers.ShowQty1c, dbo.TblCustemers.showPrice1c,"
MySQL = MySQL & "                      dbo.TblCustemers.showPrice2c, dbo.TblCustemers.Salaries1c, dbo.TblCustemers.Salaries2c, dbo.TblCustemers.Totald, dbo.TblCustemers.Totalc,"
MySQL = MySQL & "                      dbo.TblCustemers.RecordDate, dbo.TblCustemers.balanced, dbo.TblCustemers.balancec, dbo.TblCustemers.TypeCustomer, dbo.TblCustemers.BoxMil,"
MySQL = MySQL & "                      dbo.TblCustemers.ZipCode , dbo.ACCOUNTS.account_serial"
MySQL = MySQL & " FROM         dbo.TblCustemers LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblCustemers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernments ON dbo.TblCustemers.GovernmentID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernmentsCities ON dbo.TblCustemers.CityID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesData ON dbo.TblCustemers.CountryID = dbo.TblCountriesData.CountryID"
MySQL = MySQL & " Where (dbo.TblCustemers.CusID =" & val(XPTxtCusID.Text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repCustomer.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repCustomer.rpt"
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

Private Sub DcbType_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                If DcbType.ListIndex = 0 Then
               
            Account_Code_dynamic = get_account_code_branch(108, 0)
     Else
     Account_Code_dynamic = get_account_code_branch(110, 0)
     End If
     DboParentAccount.BoundText = Account_Code_dynamic
     
End If
End Sub

Private Sub DcbType_Click()
DcbType_Change
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

 ' If mdifrmmain.CarMaintenance.Visible = True Then
 ' Me.Height = 10560
 '
 'Else
 ' Me.Height = 9270
 '
 ' End If
    
    
    
 If 1 = 0 Then
'Me.Height = 10560
 Frame3.Visible = True
 Else
 Frame3.Visible = False
'Me.Height = 9270
End If

Dim StrSQL As String

    On Error GoTo ErrTrap

    'Resize_Form Me
    AddTip
    Dim Msg As String
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    LoadDataCombos

    With Me.CboSaleType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "ﬁÿ«⁄Ì"
            .AddItem " Ã«—Ï"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .AddItem "Retail"
            .AddItem "WholeSale"
        End If

    End With
  If SystemOptions.UserInterface = ArabicInterface Then
     With DcbType
       .Clear
     .AddItem "«—«÷Ì"
       .AddItem "⁄ﬁ«—"
       
    End With
 Else
    With DcbType
      .Clear
     .AddItem "Land"
      .AddItem "Estate"
       
   End With
End If

    With CboDiscountType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "·«ÌÊÃœ Œ’„"
            .AddItem "Œ’„ »ﬁÌ„…"
            .AddItem "Œ’„ »‰”»…"
        Else
            .AddItem "No"
            .AddItem "Value"
            .AddItem "percentage"
        End If

    End With
Me.DcbDigCustomer.AddItem "1"
Me.DcbDigCustomer.AddItem "2"
Me.DcbDigCustomer.AddItem "3"
Me.DcbDigCustomer.AddItem "4"
Me.DcbDigCustomer.AddItem "5"
Me.DcbDigCustomer.AddItem "6"
Me.DcbDigCustomer.AddItem "7"
    With CboSex
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "–ﬂ—"
            .AddItem "«‰ÀÏ"
        Else
            .AddItem "Male"
            .AddItem "Female"
    
        End If

    End With

    With CboDiscountTypePur
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "·«ÌÊÃœ Œ’„"
            .AddItem "Œ’„ »ﬁÌ„…"
            .AddItem "Œ’„ »‰”»…"
        Else
            .AddItem "No"
            .AddItem "Value"
            .AddItem "percentage"
        End If

    End With

    With Me.dcCreditIntervalID
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "ÌÊ„"
            .AddItem "‘Â—"
            .AddItem "”‰…"
        Else
            .AddItem "day"
            .AddItem "month"
            .AddItem "year"
        End If

    End With

    With Me.dcDepitIntervalID
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "ÌÊ„"
            .AddItem "‘Â—"
            .AddItem "”‰…"
        Else
            .AddItem "day"
            .AddItem "month"
            .AddItem "year"
        End If

    End With

    Dim Dcombos As New ClsDataCombos
    Dcombos.GetAccountingCodes Me.DboParentAccount, False, True
    Dcombos.GetCodeing Me.DCPreFix, 9
    Dcombos.GetBanks Me.DcbBank
    Dcombos.GetInvStoreGroup Me.DcbGroupInvestor
    Dcombos.GetInvStoreType Me.DcbTypeInvestor
    Dcombos.GetSalesRepData Me.DCEmP
    Me.Dtp.value = Date
    DtRecord.value = Date
    StrSQL = "select * From TblCustemers where (type=20 and Flg=1)"
        If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
            
            

  
        End If
        

        
            If SystemOptions.usertype <> UserAdminAll Then
 
          If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  Empid = " & user_id
      End If
  
        Me.Dcbranch.Enabled = True
       ' DCEmP.Enabled = False
     
    End If


        
        
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " »Ì«‰«  «·„”«Â„Ì‰  "
    LogTextE = " Open Window " & "  Customers Data "
   ' AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ErrTrap
    LogTextA = "  «·Œ—ÊÃ „‰  " & " »Ì«‰«  «·„”«Â„Ì‰  "
    LogTextE = " Exit   Window " & "  Shareholders Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "O", "", ""

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
    Set CusReport = Nothing
    Set Dcombo = Nothing

    Exit Sub
ErrTrap:
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«   «·„”«Â„Ì‰ " _
       & CHR(13) & " ﬂÊœ   " & DCPreFix & txtid.Text _
       & CHR(13) & "«·«”„ ⁄—»Ì  " & XPTxtCusName _
       & CHR(13) & "   „”∆Ê· «·« ’«·   " & TxtResponsibleContact _
       & CHR(13) & " —ﬁ„ «·Â« ›     " & XPTxtPhone _
       & CHR(13) & " —ﬁ„ «·ÃÊ«·     " & XPTxtmobile _
       & CHR(13) & " —ﬁ„ «·›«ﬂ”     " & TxtFaxNumber _
       & CHR(13) & "  «·»—Ìœ «·«·ﬂ —Ê‰Ì       " & TxtE_mail _
       & CHR(13) & " «·œÊ·Â   " & DcboCountryID.Text _
       & CHR(13) & " «·„Õ«›Ÿ…   " & DcboGovernmentID.Text _
       & CHR(13) & "  «·„œÌ‰…  " & DcboCityID.Text _
       & CHR(13) & "  «·⁄‰Ê«‰ »«· ›’Ì· " & TxtAddress _
       & CHR(13) & " „·«ÕŸ«   " & XPMTxtRemarks _
       & CHR(13) & " ‰Ê⁄ «·Œ’„ ··„»Ì⁄«    " & CboDiscountType.Text _
       & CHR(13) & "   ﬁÌ„Â «·Œ’„  " & TxtDiscountValue _
       & CHR(13) & " ‰Ê⁄ «·Œ’„ ··„‘ —Ì«    " & CboDiscountTypePur.Text _
       & CHR(13) & "   ﬁÌ„Â «·Œ’„  " & TxtDiscountValuePur _
       & CHR(13) & "  ‰Ê⁄ «·„”«Â„  " & DcCustomerType.Text _
       & CHR(13) & " «·„‰œÊ»   " & DCEmP.Text _
       & CHR(13) & " Õœ «·«∆ „«‰ „œÌ‰  " & TxtCreditLimit _
       & CHR(13) & " „œ… «·«∆ „«‰     " & TxtDepitInterval.Text & "   " & dcDepitIntervalID.Text _
       & CHR(13) & " Õœ «·«∆ „«‰ œ«∆‰   " & TxtCreditlimitCredit _
       & CHR(13) & " „œ… «·«∆ „«‰      " & TxtCreditInterval.Text & "   " & dcCreditIntervalID.Text _
                    
       LogTextA = LogTextA & CHR(13) & "„”«Â„ „Ê—œ ø       "

    If chkCustomerandVendor.value = vbChecked Then
        LogTextA = LogTextA & "‰⁄„"
    Else
        LogTextA = LogTextA & "·«"
    End If

    LogTextA = LogTextA & CHR(13) & "«Ìﬁ«› «· ⁄«„·   ø     "

    If locked.value = vbChecked Then
        LogTextA = LogTextA & "‰⁄„"
        LogTextA = LogTextA & CHR(13) & "  ”»» «·«Ìﬁ«›   "
        LogTextA = LogTextA & CHR(13) & XPMTxtRemarks2
    Else
        LogTextA = LogTextA & "·«"
    End If


    LogTextA = LogTextA & CHR(13) & "«Ìﬁ«› «· ⁄«„·  «·«Ã·   ø     "

    If creditlocked.value = vbChecked Then
        LogTextA = LogTextA & "‰⁄„"
       
    Else
        LogTextA = LogTextA & "·«"
    End If
    
    
    
    LogTextA = LogTextA & CHR(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "

    If OptType(0).value = True Then
        LogTextA = LogTextA & "„œÌ‰"
    ElseIf OptType(1).value = True Then
        LogTextA = LogTextA & "   œ«∆‰"
    ElseIf OptType(2).value = True Then
        LogTextA = LogTextA & "€Ì— „Õœœ"
    End If

    LogTextA = LogTextA & CHR(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ     " & TxtOpenBalance
    LogTextA = LogTextA & CHR(13) & "«·Õ”«» «·—∆Ì”Ì    " & DboParentAccount

    LogTextE = "  Õ›Ÿ ‘«‘… " & " Customers Data  " _
       & CHR(13) & "  Code  " & DCPreFix & txtid.Text _
       & CHR(13) & "Name " & XPTxtCusNamee _
       & CHR(13) & " Contact Person" & TxtResponsibleContact _
       & CHR(13) & " Tel " & XPTxtPhone _
       & CHR(13) & "Mob " & XPTxtmobile _
       & CHR(13) & " Fax  " & TxtFaxNumber _
       & CHR(13) & "  Email   " & TxtE_mail _
       & CHR(13) & " Contry   " & DcboCountryID.Text _
       & CHR(13) & " City   " & DcboGovernmentID.Text _
       & CHR(13) & "  Town  " & DcboCityID.Text _
       & CHR(13) & " Address " & TxtAddress _
       & CHR(13) & " Remarks  " & XPMTxtRemarks _
       & CHR(13) & " Sales Discount  type  " & CboDiscountType.Text _
       & CHR(13) & " Discount Value " & TxtDiscountValue _
       & CHR(13) & " Purchase Discount type " & CboDiscountTypePur.Text _
       & CHR(13) & "  Discount Value" & TxtDiscountValuePur _
       & CHR(13) & "  Cust. Type " & DcCustomerType.Text _
       & CHR(13) & " Sales Person   " & DCEmP.Text _
       & CHR(13) & "The limit for debit  " & TxtCreditLimit _
       & CHR(13) & " Period     " & TxtDepitInterval.Text & "   " & dcDepitIntervalID.Text _
       & CHR(13) & "The limit for Credit   " & TxtCreditlimitCredit _
       & CHR(13) & " Period " & TxtCreditInterval.Text & "   " & dcCreditIntervalID.Text _
                    
       LogTextE = LogTextE & CHR(13) & "Customer & Supplier ?  "

    If chkCustomerandVendor.value = vbChecked Then
        LogTextE = LogTextE & " Yes "
    Else
        LogTextE = LogTextE & " No "
    End If

    LogTextE = LogTextE & CHR(13) & "Locked"

    If locked.value = vbChecked Then
        LogTextE = LogTextE & "Yes "
        LogTextE = LogTextE & CHR(13) & "  Reasons  "
        LogTextE = LogTextE & CHR(13) & XPMTxtRemarks2
    Else
        LogTextE = LogTextE & "No "
    End If

    LogTextE = LogTextE & CHR(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "

    If OptType(0).value = True Then
        LogTextE = LogTextE & "„œÌ‰"
    ElseIf OptType(1).value = True Then
        LogTextE = LogTextE & "œ«∆‰"
    ElseIf OptType(2).value = True Then
        LogTextE = LogTextE & "€Ì— „Õœœ"
    End If

    LogTextE = LogTextE & CHR(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ  " & TxtOpenBalance
    LogTextE = LogTextE & CHR(13) & "  Parent Acc. " & DboParentAccount
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg, "", ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D", "", ""
    End If

End Function

Private Sub Label2_Click()
    Frame2.Visible = False
End Sub

Private Sub menue_Click(Index As Integer)
showsforms Index
End Sub

Private Sub OptType_Click(Index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.Text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.Text)
End Sub

Private Sub TxtCreditLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCreditLimit.Text, 1)
End Sub

Private Sub TxtCreditlimitCredit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCreditlimitCredit.Text, 1)
End Sub

Private Sub txtCustGID_Change()
    Dim Custcode As String
    Dim CustName As String
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        If Len(txtCustGID.Text) >= 10 Then
            If CheckCustomerID(txtCustGID, Custcode, CustName) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Â–« «·„”«Â„ „”Ã· „‰ ﬁ»·   "
                    Msg = Msg & CHR(13) & " ﬂÊœ : " & Custcode
                    Msg = Msg & CHR(13) & " «”„ «·„”«Â„ : " & CustName
                Else
                    Msg = "This Customer Already Exist"
                    Msg = Msg & CHR(13) & " Customer Code  " & Custcode
                    Msg = Msg & CHR(13) & "Customer Name  " & CustName
                                                                 
                End If

                MsgBox Msg, vbCritical
                txtCustGID.Text = ""
                                        
            End If
        End If
    End If

End Sub

Private Sub txtCustGID_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.txtCustGID.Text, 0)
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            DboParentAccount.Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„”«Â„Ì‰"
            Else
                Me.Caption = "Shareholders Data"
            End If
DcbType.Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            '    TxtCustGID.locked = True
            XPTxtCusID.locked = True
            XPTxtCusName.locked = True
            XPTxtPhone.locked = True
            XPTxtmobile.locked = True
            XPMTxtRemarks.locked = True
        
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            End If

            Fra(0).Enabled = False
            'Me.Dtp.Enabled = True
            Me.CboSaleType.Enabled = False
            Me.CboDiscountType.Enabled = False
            Me.TxtDiscountValue.Enabled = False

        Case "N"
DcbType.Enabled = True
            txtCustGID.locked = False
            DboParentAccount.Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„”«Â„Ì‰( ÃœÌœ )"
            Else
                Me.Caption = "Customers Data(Enter New Customer)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
        
            XPTxtCusID.locked = True
            XPTxtCusName.locked = False
            XPTxtPhone.locked = False
            XPTxtmobile.locked = False
            XPMTxtRemarks.locked = False
            Fra(0).Enabled = True
            '     Me.Dtp.value = Date
            '     Me.Dtp.Enabled = True
            Me.CboSaleType.Enabled = True
            Me.CboDiscountType.Enabled = True
            Me.TxtDiscountValue.Enabled = True

        Case "E"
            '  TxtCustGID.locked = True
    DcbType.Enabled = False
            DboParentAccount.Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„”«Â„Ì‰(  ⁄œÌ· )"
            Else
                Me.Caption = "Customers Data(Edit Current Customer)"
            End If

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            XPTxtCusID.locked = True
            XPTxtCusName.locked = False
            XPTxtPhone.locked = False
            XPTxtmobile.locked = False
            XPMTxtRemarks.locked = False
            Fra(0).Enabled = True
            '     Me.Dtp.value = Date
            '     Me.Dtp.Enabled = True
            Me.CboSaleType.Enabled = True
            Me.CboDiscountType.Enabled = True
            Me.TxtDiscountValue.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtOpenBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOpenBalance.Text, 0)
End Sub

Private Sub TxtSalaries1_Change()
ClcAll
End Sub

Private Sub TxtSalaries1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtSalaries1.Text, 0)
TxtSalaries1C = 0
End Sub

Private Sub TxtSalaries1C_Change()
ClcAll
End Sub

Private Sub TxtSalaries1C_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtSalaries1C.Text, 0)
TxtSalaries1 = 0
End Sub

Private Sub TxtSalaries2_Change()
ClcAll
End Sub

Private Sub TxtSalaries2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtSalaries2.Text, 0)
TxtSalaries2C = 0
End Sub

Private Sub TxtSalaries2C_Change()
ClcAll
End Sub

Private Sub TxtSalaries2C_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtSalaries2C.Text, 0)
TxtSalaries2 = 0
End Sub

Private Sub TxtSalary_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSalary.Text, 0)
End Sub

Private Sub TxtshowPrice1_Change()
ClcAll
End Sub

Private Sub TxtshowPrice1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtshowPrice1.Text, 0)
TxtshowPrice1C = 0
End Sub

Private Sub TxtshowPrice1C_Change()
ClcAll
End Sub

Private Sub TxtshowPrice1C_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtshowPrice1C.Text, 0)
 TxtshowPrice1 = 0
 
End Sub

Private Sub TxtshowPrice2_Change()
ClcAll
End Sub

Private Sub TxtshowPrice2_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtshowPrice2.Text, 0)
 TxtshowPrice2C = 0
End Sub

Private Sub TxtshowPrice2C_Change()
ClcAll
End Sub

Private Sub TxtshowPrice2C_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtshowPrice2C.Text, 0)
TxtshowPrice2 = 0
End Sub

Private Sub TxtShowQty1_Change()
ClcAll
End Sub

Private Sub TxtShowQty1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtShowQty1.Text, 0)
TxtShowQty1c = 0
End Sub

Private Sub TxtShowQty1c_Change()
ClcAll
End Sub

Private Sub TxtShowQty1c_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtShowQty1c.Text, 0)
 TxtShowQty1 = 0
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
    Dim SngCusBegainAccount As Single

    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If Not (rs.EOF Or rs.BOF) Then
        If Lngid <> 0 Then
            rs.find "CusID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    
        If rs("CustomerandVendor").value = True Then
            chkCustomerandVendor.value = vbChecked

        Else
            chkCustomerandVendor.value = vbUnchecked
        End If
        DcbType.ListIndex = IIf(IsNull(rs("Typ").value), -1, rs("Typ").value)
        Dcbranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
        Me.DtRecord.value = IIf(IsNull(rs("RecordDate")), Date, rs("RecordDate"))
        DCPreFix.Text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
        Me.txtid.Text = IIf(IsNull(rs("code").value), "", rs("code").value)
        txtopening_balance_voucher_id.Text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)
        XPTxtCusID.Text = IIf(IsNull(rs("CusID")), "", val(rs("CusID")))
        XPTxtCusName.Text = IIf(IsNull(rs("CusName")), "", Trim(rs("CusName")))
        XPTxtCusNamee.Text = IIf(IsNull(rs("CusNamee")), "", Trim(rs("CusNamee")))
        c1.Text = IIf(IsNull(rs("c1")), "", Trim(rs("c1")))
        c2.Text = IIf(IsNull(rs("c2")), "", Trim(rs("c2")))
        ''''''''/////////
         Me.DcbBank.BoundText = IIf(IsNull(rs("BankID").value), "", rs("BankID").value)
         Me.DcbTypeInvestor.BoundText = IIf(IsNull(rs("TypeInvestor").value), "", rs("TypeInvestor").value)
         Me.DcbGroupInvestor.BoundText = IIf(IsNull(rs("GroupInvestor").value), "", rs("GroupInvestor").value)
         TxtIBAN.Text = IIf(IsNull(rs("IBAN").value), "", rs("IBAN").value)
        ''/////////////salah
    Me.TxtMap.Text = IIf(IsNull(rs("Map").value), "", rs("Map").value)
    Me.txtJob.Text = IIf(IsNull(rs("JobName").value), "", rs("JobName").value)
    Me.TxtEntry.Text = IIf(IsNull(rs("Entry").value), "", rs("Entry").value)
    '''///////////////////
        Me.TxtResponsibleContact.Text = IIf(IsNull(rs("ResponsibleContact").value), "", rs("ResponsibleContact").value)
        XPTxtPhone.Text = IIf(IsNull(rs("Cus_Phone")), "", Trim(rs("Cus_Phone")))
        txtCustGID.Text = IIf(IsNull(rs("CustGID")), "", (rs("CustGID")))
    ''///
    Me.txtBox.Text = IIf(IsNull(rs("Boxmil")), "", Trim(rs("Boxmil")))
        Me.TxtZib.Text = IIf(IsNull(rs("ZipCode")), "", (rs("ZipCode")))
        DcbDigCustomer.ListIndex = IIf(IsNull(rs("TypeCustomer")), -1, (rs("TypeCustomer")))
        
        ''//
        XPTxtmobile.Text = IIf(IsNull(rs("Cus_mobile")), "", Trim(rs("Cus_mobile")))
        XPMTxtRemarks.Text = IIf(IsNull(rs("Remark")), "", Trim(rs("Remark")))
        XPMTxtRemarks2.Text = IIf(IsNull(rs("Remark2")), "", Trim(rs("Remark2")))
        Me.DboParentAccount.BoundText = IIf(IsNull(rs("parent_account")), "", rs("parent_account"))

        '    Me.CboSex = IIf(IsNull(rs("Sex")), "", rs("Sex"))
        If Not (IsNull(rs("Sex"))) Then
            If rs("Sex") = "Male" Or rs("Sex") = "–ﬂ—" Then
                Me.CboSex.ListIndex = 0
            Else
                Me.CboSex.ListIndex = 1
            End If
     
        Else
            Me.CboSex.ListIndex = 0
        End If
         
        locked.value = IIf(rs("locked") = True, 1, 0)
        creditlocked.value = IIf(rs("creditlocked") = 1, 1, 0)
        
        'creditlocked
    
        TxtDepitInterval.Text = IIf(IsNull(rs("DepitInterval")), 0, rs("DepitInterval"))
        TxtCreditInterval.Text = IIf(IsNull(rs("CreditInterval")), 0, rs("CreditInterval"))
    
        dcDepitIntervalID.ListIndex = IIf(IsNull(rs("DepitIntervalID")), -1, rs("DepitIntervalID"))
        dcCreditIntervalID.ListIndex = IIf(IsNull(rs("CreditIntervalID")), -1, rs("CreditIntervalID"))
     
        TxtCreditLimit.Text = IIf(IsNull(rs("CreditLimit").value), "0", rs("CreditLimit").value)
'gooooooooooooold
TxtShowQty1.Text = IIf(IsNull(rs("ShowQty1").value), "0", rs("ShowQty1").value)
TxtshowPrice1.Text = IIf(IsNull(rs("showPrice1").value), "0", rs("showPrice1").value)
TxtshowPrice2.Text = IIf(IsNull(rs("showPrice2").value), "0", rs("showPrice2").value)
TxtSalaries1.Text = IIf(IsNull(rs("Salaries1").value), "0", rs("Salaries1").value)
TxtSalaries2.Text = IIf(IsNull(rs("Salaries2").value), "0", rs("Salaries2").value)
TxtShowQty1c.Text = IIf(IsNull(rs("ShowQty1c").value), "0", rs("ShowQty1c").value)
TxtshowPrice1C.Text = IIf(IsNull(rs("showPrice1C").value), "0", rs("showPrice1C").value)
TxtshowPrice2C.Text = IIf(IsNull(rs("showPrice2C").value), "0", rs("showPrice2C").value)
TxtSalaries1C.Text = IIf(IsNull(rs("Salaries1C").value), "0", rs("Salaries1C").value)
 TxtSalaries2C.Text = IIf(IsNull(rs("Salaries2C").value), "0", rs("Salaries2C").value)

'txtTotald.text = IIf(IsNull(rs("Totald").value), "0", rs("Totald").value)
'txtTotalc.text = IIf(IsNull(rs("Totalc").value), "0", rs("Totalc").value)
'Txtbalanced.text = IIf(IsNull(rs("balanced").value), "0", rs("balanced").value)
'TxtbalancedC.text = IIf(IsNull(rs("balancec").value), "0", rs("balancec").value)
'gooooo
 
        
        
        
   
        
        If Not (IsNull(rs("OpenBalanceDate").value)) Then
            Me.Dtp.value = rs("OpenBalanceDate").value
            '       Me.Dtp.Enabled = True
        Else
        
            Me.Dtp.value = Date
            Me.Dtp.Enabled = False
        End If
    
    
    
    
            If Not (IsNull(rs("OpenBalanceDate").value)) Then
            Me.Dtp1.value = rs("OpenBalanceDate").value
            '       Me.Dtp.Enabled = True
        Else
        
            Me.Dtp1.value = Date
            Me.Dtp1.Enabled = False
        End If
        
        
        
                If Not (IsNull(rs("OpenBalanceDate").value)) Then
            Me.Dtp2.value = rs("OpenBalanceDate").value
            '       Me.Dtp.Enabled = True
        Else
        
            Me.Dtp2.value = Date
            Me.Dtp2.Enabled = False
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
       
       
       
       
       
       
        If Not IsNull(rs("OpenBalanceType1").value) Then
        Me.TxtOpenBalance1.Text = IIf(IsNull(rs("OpenBalance1")), "", Trim(rs("OpenBalance1")))

        If rs("OpenBalanceType1").value = 0 Then
            OptType1(0).value = True
            OptType1_Click 0
        ElseIf rs("OpenBalanceType1").value = 1 Then
            OptType1(1).value = True
            OptType1_Click 1
        End If
    
    Else
        Me.TxtOpenBalance1.Text = 0
        Me.OptType1(2).value = True
        OptType1_Click 2
    End If

    If Not IsNull(rs("OpenBalanceType2").value) Then
        Me.TxtOpenBalance2.Text = IIf(IsNull(rs("OpenBalance2")), "", Trim(rs("OpenBalance2")))

        If rs("OpenBalanceType2").value = 0 Then
            OptType2(0).value = True
            OptType2_Click 0
        ElseIf rs("OpenBalanceType2").value = 1 Then
            OptType2(1).value = True
            OptType2_Click 1
        End If
    
    Else
        Me.TxtOpenBalance2.Text = 0
        Me.OptType2(2).value = True
        OptType2_Click 2
    End If


        Me.TxtCreditlimitCredit.Text = IIf(IsNull(rs("CreditlimitCredit").value), "0", rs("CreditlimitCredit").value)
        Me.TxtFaxNumber.Text = IIf(IsNull(rs("FaxNumber").value), "", rs("FaxNumber").value)
        Me.TxtE_mail.Text = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
          SngCusBegainAccount = GetCustomerAccount(val(XPTxtCusID.Text), True)
        Dim balanceString As String
WriteCustomerBalPublic IIf(IsNull(rs("Account_Code")), "", rs("Account_Code")), , balanceString
        lbl(8).Caption = balanceString
    
        '    If SngCusBegainAccount < 0 Then
        '        Me.lbl(8).Caption = Abs(SngCusBegainAccount)
        '        If SystemOptions.UserInterface = ArabicInterface Then
        '        Me.lbl(9).Caption = "„œÌ‰"
        '        Else
        '        Me.lbl(9).Caption = "Depit"
        '        End If
        
        '    ElseIf SngCusBegainAccount > 0 Then
        '        Me.lbl(8).Caption = Abs(SngCusBegainAccount)
        '        If SystemOptions.UserInterface = ArabicInterface Then
        '        Me.lbl(9).Caption = "œ«∆‰"
        '
        '        Else
        '        Me.lbl(9).Caption = "Credit"
        '        End If
        '    Else
        '        Me.lbl(8).Caption = 0
        '        If SystemOptions.UserInterface = ArabicInterface Then
        '        Me.lbl(9).Caption = " "
        '        Else
        '        Me.lbl(9).Caption = ""
        '        End If
        '    End If
    
        If IsNull(rs("SaleType").value) Then
            Me.CboSaleType.ListIndex = -1
        ElseIf rs("SaleType").value = 0 Then
            Me.CboSaleType.ListIndex = 0
        ElseIf rs("SaleType").value = 1 Then
            Me.CboSaleType.ListIndex = 1
        End If

        If IsNull(rs("Trans_DiscountType").value) Then
            Me.CboDiscountType.ListIndex = 0
            Me.TxtDiscountValue.Text = 0
        ElseIf rs("Trans_DiscountType").value = 0 Then
            Me.CboDiscountType.ListIndex = 0
            Me.TxtDiscountValue.Text = 0
        ElseIf rs("Trans_DiscountType").value = 1 Then
            Me.CboDiscountType.ListIndex = 1
            Me.TxtDiscountValue.Text = IIf(IsNull(rs("Trans_Discount").value), "", rs("Trans_Discount").value)
        ElseIf rs("Trans_DiscountType").value = 2 Then
            Me.CboDiscountType.ListIndex = 2
            Me.TxtDiscountValue.Text = IIf(IsNull(rs("Trans_Discount").value), "", rs("Trans_Discount").value)
        End If
    
        If IsNull(rs("Trans_DiscountTypePur").value) Then
            Me.CboDiscountTypePur.ListIndex = 0
            Me.TxtDiscountValuePur.Text = 0
        ElseIf rs("Trans_DiscountTypePur").value = 0 Then
            Me.CboDiscountTypePur.ListIndex = 0
            Me.TxtDiscountValuePur.Text = 0
        ElseIf rs("Trans_DiscountTypePur").value = 1 Then
            Me.CboDiscountTypePur.ListIndex = 1
            Me.TxtDiscountValuePur.Text = IIf(IsNull(rs("Trans_DiscountPur").value), "", rs("Trans_DiscountPur").value)
        ElseIf rs("Trans_DiscountTypePur").value = 2 Then
            Me.CboDiscountTypePur.ListIndex = 2
            Me.TxtDiscountValuePur.Text = IIf(IsNull(rs("Trans_DiscountPur").value), "", rs("Trans_DiscountPur").value)
        End If
    
        Me.DcboCountryID.BoundText = IIf(IsNull(rs("CountryID")), "", rs("CountryID"))
        Me.DcboCountryID2.BoundText = IIf(IsNull(rs("CountryID2")), "", rs("CountryID2"))
    
        Me.DcboGovernmentID.BoundText = IIf(IsNull(rs("GovernmentID")), "", rs("GovernmentID"))
        Me.DcboCityID.BoundText = IIf(IsNull(rs("CityID")), "", rs("CityID"))
        Me.DCEmP.BoundText = IIf(IsNull(rs("EmpId")), "", rs("EmpId"))
        Me.DcCustomerType.BoundText = IIf(IsNull(rs("CustomerTypeID")), "", rs("CustomerTypeID"))
      
        Me.TxtAddress.Text = IIf(IsNull(rs("Address")), "", Trim(rs("Address")))
        '19082013
        Txt_DateExpLincH.value = IIf(IsNull(rs("ExpireDateH").value), ToHijriDate(Date), rs("ExpireDateH").value)

        Me.TxtCompany.Text = IIf(IsNull(rs("Company")), "", Trim(rs("Company")))
        Me.TxtJobTitle.Text = IIf(IsNull(rs("JobTitle")), "", Trim(rs("JobTitle")))
        Me.TxtSalary.Text = IIf(IsNull(rs("Salary")), 0, Trim(rs("Salary")))
        Me.TxtJobAddress.Text = IIf(IsNull(rs("JobAddress")), "", Trim(rs("JobAddress")))
        Me.TxtJobTel.Text = IIf(IsNull(rs("JobTel")), "", Trim(rs("JobTel")))
        Me.TxtJobTelConvert.Text = IIf(IsNull(rs("JobTelConvert")), "", Trim(rs("JobTelConvert")))
        Me.TxtHomeTel.Text = IIf(IsNull(rs("HomeTel")), "", Trim(rs("HomeTel")))
        Me.TxtMobile1.Text = IIf(IsNull(rs("Mobile1")), "", Trim(rs("Mobile1")))
        Me.TXtMobile2.Text = IIf(IsNull(rs("Mobile2")), "", Trim(rs("Mobile2")))
    
    End If

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Member()
    Dim Msg As String
    Dim IntRes As Integer
    Dim BegainTrans As Boolean
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If XPTxtCusID.Text <> "" Then

        Msg = "”Ì „ Õ–› »Ì«‰«  «·„”«Â„   " & CHR(13)
        Msg = Msg + (XPTxtCusName.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

        If IntRes = vbYes Then
            If Not rs.RecordCount < 1 Then
                DeleteOpeningBalance
                Cn.BeginTrans
                BegainTrans = True
          
                ' StrSQL = "DELETE From NOTES Where NoteType=101 AND CusID=" & Val(Me.XPTxtCusID.text)
                ' Cn.Execute StrSQL, , adExecuteNoRecords
            
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
             
                '   update_account_opening_balance get_account_code_branch(19, my_branch)
               
                Dim StrAccountCode As String
                Dim StrAccountCode1 As String
                Dim StrAccountCode2 As String
                Dim ParentAccount As String
                
StrAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
If SystemOptions.CustomerhavethreeAccounts1 = True Then
                'StrAccountCode1 = rs("Account_Code1").value
                'StrAccountCode2 = rs("Account_Code2").value
                StrAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                        StrAccountCode1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
                    StrAccountCode2 = IIf(IsNull(rs("Account_Code2").value), "", rs("Account_Code2").value)
                    
                    
 End If
 
                StrSQL = "Delete From Accounts Where Account_Code='" & StrAccountCode & "'"
                     
If SystemOptions.CustomerhavethreeAccounts1 = True Then
                    If Not IsNull(rs("Account_Code1").value) Then
                   StrSQL = StrSQL & " or   Account_Code='" & rs("Account_Code1").value & "'"
                   End If
        
        
             If Not IsNull(rs("Account_Code2").value) Then
            StrSQL = StrSQL & " or   Account_Code='" & rs("Account_Code2").value & "'"
          End If
        
   End If
                Cn.Execute StrSQL, , adExecuteNoRecords
                CuurentLogdata ("D")

                      If SystemOptions.CustomerhavethreeAccounts1 = True Then
                    StrAccountCode1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
                    StrAccountCode2 = IIf(IsNull(rs("Account_Code2").value), "", rs("Account_Code2").value)
                    
                    ParentAccount = IIf(IsNull(rs("ParentAccount").value), "", rs("ParentAccount").value)

                                    If ModAccounts.DeleteAccount(StrAccountCode, True) = True And ModAccounts.DeleteAccount(StrAccountCode1, True) = True And ModAccounts.DeleteAccount(StrAccountCode2, True) = True And ModAccounts.DeleteAccount(ParentAccount, True) = True Then
                                       CuurentLogdata ("D")
                                        rs.delete
                                  '      Msg = " „  ⁄„·Ì… «·Õ–›."
                                  '      MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            
                                    Else
                                        GoTo ErrTrap
                                    End If

                Else

                                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                                    CuurentLogdata ("D")
                                    rs.delete
                                Else
                                    Exit Sub
                                End If
                End If
                

                Msg = " „  ⁄„·Ì… «·Õ–›."
                MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            
                Cn.CommitTrans
                BegainTrans = False
                XPBtnMove_Click 2

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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„”«Â„ "
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate

    If BegainTrans = True Then
        Cn.RollbackTrans
        BegainTrans = False
    End If

    'End If
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "CusID='" & val(XPTxtCusID.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub OptType1_Click(Index As Integer)
    Me.TxtOpenBalance1.Enabled = Not OptType1(2).value
    Me.TxtOpenBalance1.Text = IIf(OptType1(2).value = True, 0, Me.TxtOpenBalance1.Text)

End Sub

Private Sub OptType2_Click(Index As Integer)
    Me.TxtOpenBalance2.Enabled = Not OptType2(2).value
    Me.TxtOpenBalance2.Text = IIf(OptType2(2).value = True, 0, Me.TxtOpenBalance2.Text)

End Sub

Function ClcAll()
Dim different As Double
txtTotald = val(TxtshowPrice1) + val(TxtshowPrice2) + val(TxtSalaries1) + val(TxtSalaries2)
txtTotalc = val(TxtshowPrice1C) + val(TxtshowPrice2C) + val(TxtSalaries1C) + val(TxtSalaries2C)
different = txtTotald - txtTotalc
If different > 0 Then
Txtbalanced = different
TxtbalancedC = 0
Else
TxtbalancedC = different * -1
Txtbalanced = 0
End If

   If val(Txtbalanced.Text) > 0 Then
       OptType(0).value = True
       TxtOpenBalance.Text = val(Txtbalanced.Text)
       
    ElseIf val(TxtbalancedC.Text) > 0 Then
       OptType(1).value = True
       TxtOpenBalance.Text = val(TxtbalancedC.Text)
           
           
       Else
       OptType(2).value = True
         TxtOpenBalance.Text = 0
       End If
       
End Function
Private Sub SaveData()
    
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim IntRes As Integer
    Dim BeginTrans As Boolean
    Dim RsNotes As ADODB.Recordset
    Dim LngOpenID As Long

    On Error GoTo ErrTrap

    If Trim(Dcbranch.BoundText) = "" Then

    End If

    If Me.TxtModFlg.Text <> "R" Then
        If XPTxtCusName.Text = "" Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "ÌÃ» «œŒ«· «”„ «·„”«Â„"
                Else
                Msg = "Enter Customer name"
                End If
       
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtCusName.SetFocus
            Exit Sub
        End If

        If Me.OptType(2).value = False Then
                    If val(Me.TxtOpenBalance.Text) = 0 Then
                        
                        
                     If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·—’Ìœ «·Ã«—Ì...!!!"
                Else
                Msg = "Enter  Opening  Balance"
                End If

                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
                                If TxtOpenBalance.Enabled = True Then
                                    TxtOpenBalance.SetFocus
                                End If
        
                        Exit Sub
                    End If
        End If




        If Me.OptType1(2).value = False Then
                    If val(Me.TxtOpenBalance1.Text) = 0 Then
                        
                        
                                     If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·—’Ìœ ··‘Ìﬂ«  „ƒÃ·…··„”«Â„...!!!"
                Else
                Msg = "Enter  Opening  Balance for Checks"
                End If
                
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
                                If TxtOpenBalance1.Enabled = True Then
                                    TxtOpenBalance1.SetFocus
                                End If
        
                        Exit Sub
                    End If
        End If
        
        
        
                If Me.OptType2(2).value = False Then
                    If val(Me.TxtOpenBalance2.Text) = 0 Then
                        
                        
                        
                         If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·—’Ìœ ··œ›⁄«  «·„ﬁœ„… ··„”«Â„...!!!"
                Else
                Msg = "Enter  Opening  Balance for Adv Payments"
                End If
                
                
                        
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
                                If TxtOpenBalance2.Enabled = True Then
                                    TxtOpenBalance2.SetFocus
                                End If
        
                        Exit Sub
                    End If
        End If
        
        
        
        
        If val(Me.TxtCreditLimit.Text) > 0 Then

            'Â‰«ﬂ Õœ ≈∆ „«‰ ( „œÌ‰)ÊÌÃ» «· «ﬂœ „‰ «·—’Ìœ «·√›  «ÕÏ «·„œŒ·
            If Me.OptType(0).value = True Then

                '«·—’Ìœ «·√›  «ÕÏ „œÌ‰
                If val(Me.TxtOpenBalance.Text) > val(Me.TxtCreditLimit.Text) Then
                    
                                  If SystemOptions.UserInterface = ArabicInterface Then
                   
                    Msg = "≈‰ »Â ....!!!"
                    Msg = Msg & CHR(13) & "·ﬁœ Ê÷⁄  Õœ ≈∆ „«‰ („œÌ‰ ) ··„”«Â„ " & val(Me.TxtCreditLimit.Text)
                    Msg = Msg & CHR(13) & "·ﬂ‰ﬂ Ê÷⁄  «·—’Ìœ «·≈›  «ÕÏ ··„”«Â„ „œÌ‰ »‹  " & val(Me.TxtOpenBalance.Text)
                    Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·»Ì«‰«  «· Ï «œŒ· Â«...øøø"
               
                Else
                  
                    Msg = "Hint  ....!!!"
                    Msg = Msg & CHR(13) & "Credit  Is  " & val(Me.TxtCreditLimit.Text)
                    Msg = Msg & CHR(13) & "Depit opening balance is   " & val(Me.TxtOpenBalance.Text)
                    Msg = Msg & CHR(13) & "???????"
               
                End If
                    
                     
                    
                    IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRtlReading + vbMsgBoxRight, App.title)

                    If IntRes = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        If val(Me.TxtCreditlimitCredit.Text) > 0 Then

            'Â‰«ﬂ Õœ ≈∆ „«‰ ( „œÌ‰)ÊÌÃ» «· «ﬂœ „‰ «·—’Ìœ «·√›  «ÕÏ «·„œŒ·
            If Me.OptType(1).value = True Then

                '«·—’Ìœ «·√›  «ÕÏ œ«∆‰
                If val(Me.TxtOpenBalance.Text) > val(Me.TxtCreditlimitCredit.Text) Then
                    
                                If SystemOptions.UserInterface = ArabicInterface Then
                   
                   Msg = "≈‰ »Â ....!!!"
                    Msg = Msg & CHR(13) & "·ﬁœ Ê÷⁄  Õœ ≈∆ „«‰ (œ«∆‰ ) ··„”«Â„ " & val(Me.TxtCreditlimitCredit.Text)
                    Msg = Msg & CHR(13) & "·ﬂ‰ﬂ Ê÷⁄  «·—’Ìœ «·≈›  «ÕÏ ··„”«Â„ œ«∆‰ »‹  " & val(Me.TxtOpenBalance.Text)
                    Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·»Ì«‰«  «· Ï «œŒ· Â«...øøø"
                    IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRtlReading + vbMsgBoxRight, App.title)
 
                Else
                  
                    Msg = "Hint  ....!!!"
                    Msg = Msg & CHR(13) & "Credit  Is  " & val(Me.TxtCreditLimit.Text)
                    Msg = Msg & CHR(13) & "Credit opening balance is   " & val(Me.TxtOpenBalance.Text)
                    Msg = Msg & CHR(13) & "???????"
               
                End If
                
               
                    If IntRes = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        If Me.CboDiscountType.ListIndex = -1 Or Me.CboDiscountType.ListIndex = 0 Then
            Me.TxtDiscountValue.Text = 0
        ElseIf Me.CboDiscountType.ListIndex = 1 Then

            If val(Me.TxtDiscountValue.Text) = 0 Then
             If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·Œ’„ «·Œ«’… »«·„”«Â„...!!!"
Else
Msg = "Enter Discount value "
End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtDiscountValue.SetFocus
                Exit Sub
            End If

        ElseIf Me.CboDiscountType.ListIndex = 2 Then

            If val(Me.TxtDiscountValue.Text) = 0 Then
                
                
                             If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "ÌÃ» ﬂ «»… ‰”»… «·Œ’„ «·Œ«’… »«·„”«Â„...!!!"
Else
Msg = "Enter Discount %  "
End If


                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtDiscountValue.SetFocus
                Exit Sub
            ElseIf val(Me.TxtDiscountValue.Text) > 100 Then
            
                                         If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ‰”»… «·Œ’„ «ﬂ»— „‰ 100 ...!!!"
Else
Msg = "  Discount % cant > 100  "
End If


                
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtDiscountValue.SetFocus
                Exit Sub
            End If
        End If
    
        If Me.CboDiscountTypePur.ListIndex = -1 Or Me.CboDiscountTypePur.ListIndex = 0 Then
            Me.TxtDiscountValuePur.Text = 0
        ElseIf Me.CboDiscountTypePur.ListIndex = 1 Then

            If val(Me.TxtDiscountValuePur.Text) = 0 Then
                                                     If SystemOptions.UserInterface = ArabicInterface Then

               Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·Œ’„ «·Œ«’… »«·„”«Â„ ›Ï ›Ê« Ì— «·‘—«¡...!!!"
Else
Msg = "  Enter Discount   value For purchase invoices  "
End If

                
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtDiscountValuePur.SetFocus
                Exit Sub
            End If

        ElseIf Me.CboDiscountTypePur.ListIndex = 2 Then

            If val(Me.TxtDiscountValuePur.Text) = 0 Then
                
                
          If SystemOptions.UserInterface = ArabicInterface Then

               Msg = "ÌÃ» ﬂ «»… ‰”»… «·Œ’„ «·Œ«’… «·„”«Â„ ›Ï ›Ê« Ì— «·‘—«¡..!!!"
Else
Msg = "  Enter Discount   %  For purchase invoices  "
End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtDiscountValuePur.SetFocus
                Exit Sub
            ElseIf val(Me.TxtDiscountValuePur.Text) > 100 Then
                 
                
      If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ‰”»… «·Œ’„ «ﬂ»— „‰ 100 ...!!!"
Else
Msg = "  Discount % cant > 100  "
End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtDiscountValuePur.SetFocus
                Exit Sub
            End If
        End If
    
    
    
        Select Case TxtModFlg.Text

            Case "N"
                XPTxtCusID.Text = CStr(new_id("TblCustemers", "CusID", "", True))
            
                StrSQL = "Select * From TblCustemers where Type=20 And CusName='" & Trim(XPTxtCusName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                      If SystemOptions.UserInterface = ArabicInterface Then

                 Msg = "ÌÊÃœ „”«Â„ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                    Else
                     Msg = "this Customer Already Exist" & CHR(13)
                     
                    End If

                   
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtCusName.SetFocus
                    Exit Sub
                End If

            RsTemp.Close
            Set RsTemp = Nothing
          StrSQL = "Select * From TblCustemers where   Type=20 And fullcode='" & Trim(DCPreFix.Text & txtid.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                      If SystemOptions.UserInterface = ArabicInterface Then

                 Msg = "ÌÊÃœ „”«Â„  „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ " & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                    Else
                     Msg = "this Customer Already Exist" & CHR(13)
                     
                    End If

                   
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtCusName.SetFocus
                    Exit Sub
                End If
                
                
                

            Case "E"
                StrSQL = "select * From TblCustemers where Type=20 And CusName='" & Trim(XPTxtCusName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText


                If RsTemp.RecordCount > 0 Then
                                If RsTemp("CusID").value <> val(XPTxtCusID.Text) Then
                                     
                                                      If SystemOptions.UserInterface = ArabicInterface Then
                                
                                                 Msg = "ÌÊÃœ „”«Â„ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                                                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                                                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                                                    Else
                                                     Msg = "this Customer Already Exist" & CHR(13)
                                                     
                                                    End If
            
                                     MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                    XPTxtCusName.SetFocus
                                    Exit Sub
                                End If
                End If

     RsTemp.Close
            Set RsTemp = Nothing
          StrSQL = "Select * From TblCustemers where Type=20 And fullcode='" & Trim(DCPreFix.Text & txtid.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If RsTemp.RecordCount > 0 Then
                                If RsTemp("CusID").value <> val(XPTxtCusID.Text) Then
                                     
                                                      If SystemOptions.UserInterface = ArabicInterface Then
                                
                                                 Msg = "ÌÊÃœ „”«Â„ „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ " & CHR(13)
                                                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                                                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                                                    Else
                                                     Msg = "this Customer Already Exist" & CHR(13)
                                                     
                                                    End If
            
                                     MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                    XPTxtCusName.SetFocus
                                    Exit Sub
                                End If
                End If
                

        End Select

        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then
            Dim Account_Code_dynamic As String

            Account_Code_dynamic = Me.DboParentAccount.BoundText
            rs.AddNew

            rs("CusID").value = val(XPTxtCusID.Text)
        ElseIf Me.TxtModFlg.Text = "E" Then
            Account_Code_dynamic = Me.DboParentAccount.BoundText
            '  StrSQL = "DELETE From NOTES Where NoteType=101 AND CusID=" & Val(Me.XPTxtCusID.text)
            '   Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
        End If
        
    
            If val(TxtOpenBalance.Text) <> 0 Or val(TxtOpenBalance1.Text) <> 0 Or val(TxtOpenBalance2.Text) <> 0 Then
                txtopening_balance_voucher_id.Text = get_opening_balance_voucher_id
                rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.Text)
            Else
                rs("opening_balance_voucher_id").value = Null
            End If
            
       
       
           
       '     If val(TxtOpenBalance.text) <> 0 Or val(TxtOpenBalance1.text) <> 0 Or val(TxtOpenBalance2.text) <> 0 Then
       '         txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
       '         rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
       '     Else
       '         rs("opening_balance_voucher_id").value = Null
       '     End If
       
       
 
         
        rs("code").value = txtid.Text
        rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.Text) & IIf(Trim(txtid.Text) = "", Null, txtid.Text)
        rs("prifix").value = IIf(DCPreFix.Text = "", Null, DCPreFix.Text)
 Me.TxtFullcode = IIf(DCPreFix.BoundText = "", Null, DCPreFix.Text) & IIf(Trim(txtid.Text) = "", Null, txtid.Text)
     
        If chkCustomerandVendor.value = vbChecked Then
            rs("CustomerandVendor").value = 1

        Else
            rs("CustomerandVendor").value = 0
        End If
        rs("Typ").value = val(Me.DcbType.ListIndex)
        rs("BranchId").value = IIf(Me.Dcbranch.BoundText = "", 0, val(Dcbranch.BoundText))
        rs("RecordDate").value = Me.DtRecord.value
        rs("CusName").value = Trim(XPTxtCusName.Text)
        If Trim(XPTxtCusNamee.Text) = "" Then XPTxtCusNamee.Text = Trim(XPTxtCusName.Text)
        rs("CusNamee").value = IIf(Trim(XPTxtCusNamee.Text) = "", Trim(XPTxtCusName.Text), Trim(XPTxtCusNamee.Text))
        rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.Text)
    
        rs("c1").value = Trim(c1.Text)
        rs("c2").value = Trim(c2.Text)
    
        rs("CustGID").value = IIf(txtCustGID.Text = "", Null, val(txtCustGID.Text))
       
        rs("Cus_Phone").value = IIf(XPTxtPhone.Text = "", "", Trim(XPTxtPhone.Text))
        rs("Cus_mobile").value = IIf(XPTxtmobile.Text = "", "", Trim(XPTxtmobile.Text))
        rs("Remark").value = IIf(XPMTxtRemarks.Text = "", "", Trim(XPMTxtRemarks.Text))
        rs("Remark2").value = IIf(XPMTxtRemarks2.Text = "", "", Trim(XPMTxtRemarks2.Text))
        rs("parent_account").value = IIf(Me.DboParentAccount.BoundText = "", Null, (Me.DboParentAccount.BoundText))
        rs("EmpId").value = IIf(Me.DCEmP.BoundText = "", Null, (Me.DCEmP.BoundText))
    
        If locked.value = vbChecked Then
            rs("locked").value = 1
        Else
            rs("locked").value = 0
        End If
'creditlocked
    If creditlocked.value = vbChecked Then
            rs("creditlocked").value = 1
        Else
            rs("creditlocked").value = 0
        End If
        
        rs("CreditLimit").value = val(Me.TxtCreditLimit.Text)
        rs("Type").value = 20
        '''''
        rs("Flg").value = 1
        rs("BankID").value = val(Me.DcbBank.BoundText)
        rs("TypeInvestor").value = val(Me.DcbTypeInvestor.BoundText)
        rs("GroupInvestor").value = val(Me.DcbGroupInvestor.BoundText)
        rs("IBAN").value = Me.TxtIBAN.Text
        '''''
        
        rs("DepitInterval").value = val(TxtDepitInterval.Text)
        rs("CreditInterval").value = val(TxtCreditInterval.Text)
        
        rs("DepitIntervalID").value = val(dcDepitIntervalID.ListIndex)
        rs("CreditIntervalID").value = val(dcCreditIntervalID.ListIndex)
    'goooooooooooold
    
       rs("ShowQty1").value = val(Me.TxtShowQty1.Text)
       rs("showPrice1").value = val(Me.TxtshowPrice1.Text)
       rs("showPrice2").value = val(Me.TxtshowPrice2.Text)
        rs("Salaries1").value = val(Me.TxtSalaries1.Text)
        rs("Salaries2").value = val(Me.TxtSalaries2.Text)
        
       rs("ShowQty1c").value = val(Me.TxtShowQty1c.Text)
       rs("showPrice1c").value = val(Me.TxtshowPrice1C.Text)
       rs("showPrice2c").value = val(Me.TxtshowPrice2C.Text)
        rs("Salaries1c").value = val(Me.TxtSalaries1C.Text)
        rs("Salaries2c").value = val(Me.TxtSalaries2C.Text)
        
        
        rs("Totald").value = val(Me.txtTotald.Text)
        rs("Totalc").value = val(Me.txtTotalc.Text)
       rs("balanced").value = val(Me.Txtbalanced.Text)
        rs("balancec").value = val(Me.TxtbalancedC.Text)
        
    
       
        
       
       'goooooooooooold
       
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



        If Me.OptType1(2).value = True Then
            rs("OpenBalance1").value = 0
            rs("OpenBalanceType1").value = Null
        ElseIf Me.OptType1(0).value = True Then
            rs("OpenBalance1").value = val(Me.TxtOpenBalance1.Text)
            rs("OpenBalanceType1").value = 0
        ElseIf Me.OptType1(1).value = True Then
            rs("OpenBalance1").value = val(Me.TxtOpenBalance1.Text)
            rs("OpenBalanceType1").value = 1
        End If
        
        
        If Me.OptType2(2).value = True Then
            rs("OpenBalance2").value = 0
            rs("OpenBalanceType2").value = Null
        ElseIf Me.OptType2(0).value = True Then
            rs("OpenBalance2").value = val(Me.TxtOpenBalance2.Text)
            rs("OpenBalanceType2").value = 0
        ElseIf Me.OptType2(1).value = True Then
            rs("OpenBalance2").value = val(Me.TxtOpenBalance2.Text)
            rs("OpenBalanceType2").value = 1
        End If
        
        
        
        rs("OpenBalanceDate").value = Me.Dtp.value
    
        
        rs("CreditlimitCredit").value = val(Me.TxtCreditlimitCredit.Text)
        rs("FaxNumber").value = IIf(Trim$(Me.TxtFaxNumber.Text) = "", Null, Trim$(Me.TxtFaxNumber.Text))
        rs("E_mail").value = IIf(Trim$(Me.TxtE_mail.Text) = "", Null, Trim$(Me.TxtE_mail.Text))

        If Me.CboSaleType.ListIndex = 0 Or Me.CboSaleType.ListIndex = -1 Then
            rs("SaleType").value = 0
        Else
            rs("SaleType").value = 1
        End If

        If Me.CboDiscountType.ListIndex = -1 Or Me.CboDiscountType.ListIndex = 0 Then
            rs("Trans_DiscountType").value = 0
            rs("Trans_Discount").value = 0
        ElseIf Me.CboDiscountType.ListIndex = 1 Then
            rs("Trans_DiscountType").value = 1
            rs("Trans_Discount").value = val(Me.TxtDiscountValue.Text)
        ElseIf Me.CboDiscountType.ListIndex = 2 Then
            rs("Trans_DiscountType").value = 2
            rs("Trans_Discount").value = val(Me.TxtDiscountValue.Text)
        End If
    
        If Me.CboDiscountTypePur.ListIndex = -1 Or Me.CboDiscountTypePur.ListIndex = 0 Then
            rs("Trans_DiscountTypePur").value = 0
            rs("Trans_DiscountPur").value = 0
        ElseIf Me.CboDiscountTypePur.ListIndex = 1 Then
            rs("Trans_DiscountTypePur").value = 1
            rs("Trans_DiscountPur").value = val(Me.TxtDiscountValuePur.Text)
        ElseIf Me.CboDiscountTypePur.ListIndex = 2 Then
            rs("Trans_DiscountTypePur").value = 2
            rs("Trans_DiscountPur").value = val(Me.TxtDiscountValuePur.Text)
        End If
    
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            Dim ParentAccount As String
            
            If Me.TxtModFlg.Text = "N" Then
        
       '         rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtCusName.text), True, False, Trim$(Me.XPTxtCusNamee.text))
                
          If SystemOptions.CustomerhavethreeAccounts1 = False Then
        
                                   rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtCusName.Text), True, False, Trim$(Me.XPTxtCusNamee.Text))
          Else
                
                                        If SystemOptions.CustomerhavethreeAccounts1 = True Then
                                            ParentAccount = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtCusName.Text, False, False, XPTxtCusNamee.Text)
                                            rs("ParentAccount").value = ParentAccount
                                         
                                            rs("Account_Code").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.Text), True, False, XPTxtCusNamee.Text)
                                            rs("Account_Code1").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.Text) & "   ‘Ìﬂ«   „ƒÃ·…", True, False, XPTxtCusNamee.Text & "  Under Collection Cheque  ")
                                            rs("Account_Code2").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.Text) & "   œ›⁄«  „ﬁœ„…   ", True, False, XPTxtCusNamee.Text & " Advanced Payments")

                                        Else
                                            rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtCusName.Text), True, False, XPTxtCusNamee.Text)
                                            rs("ParentAccount").value = Null
                                            
                                        End If
             
        End If
                
                
                
                
                
                'Rs("Account_Code").value = ModAccounts.AddNewAccount("a1a2a3", Trim$(Me.XPTxtCusName.text), True, False)
            Else

                 '       If Not IsNull(rs("Account_Code").value) Then
                 '           ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtCusName.text, Me.XPTxtCusNamee.text, , , , , , , , , , , , , , , , , True
                 '       End If
                        
                        
                         If SystemOptions.CustomerhavethreeAccounts1 = False Then
                    If Not IsNull(rs("Account_Code").value) Then
                        ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtCusName.Text, XPTxtCusNamee.Text, , , , , , , , , , , , , , , , , True
                    End If
            
                Else
          
                    If Not IsNull(rs("ParentAccount").value) And Not (rs("ParentAccount").value) = "" Then
                        ModAccounts.EditAccount rs("ParentAccount").value, Me.XPTxtCusName.Text, Trim(XPTxtCusNamee.Text), , , , , , , , , , , , , , , , , False
                        Else
                           ' rs("ParentAccount").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtCusName.text), True, False, XPTxtCusNamee.text)
                               '  rs("ParentAccount").value = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtCusName.text, False, False, XPTxtCusNamee.text)
                                     
                                     ParentAccount = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtCusName.Text, False, False, XPTxtCusNamee.Text)
                                            rs("ParentAccount").value = ParentAccount
                                            
                                     '       rs("ParentAccount").value = ParentAccount

                    End If
            
                    If Not IsNull(rs("Account_Code").value) And Not (rs("Account_Code").value) = "" Then
                        ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtCusName.Text, XPTxtCusNamee.Text, , , , , , , , , , , , , , , , , True
                      Else
                          rs("Account_Code").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.Text), True, False, XPTxtCusNamee.Text)
                             
                    End If
            
                    If Not IsNull(rs("Account_Code1").value) And Not (rs("Account_Code1").value) = "" Then
                        ModAccounts.EditAccount rs("Account_Code1").value, Me.XPTxtCusName.Text & "    ‘Ìﬂ«   „ƒÃ·… ", XPTxtCusNamee.Text & "  Cheque Box ", , , , , , , , , , , , , , , , , True
                        Else
                                               rs("Account_Code1").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.Text) & "   ‘Ìﬂ«   „ƒÃ·…", True, False, XPTxtCusNamee.Text & "  Under Collection Cheque  ")
                                         

                    End If
          
                    If Not IsNull(rs("Account_Code2").value) And Not (rs("Account_Code2").value) = "" Then
                        ModAccounts.EditAccount rs("Account_Code2").value, Me.XPTxtCusName.Text & "   œ›⁄«  „ﬁœ„…   ", XPTxtCusNamee.Text & "  Cheque Box ", , , , , , , , , , , , , , , , , True
                        Else
                                               rs("Account_Code2").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.Text) & "   œ›⁄«  „ﬁœ„…   ", True, False, XPTxtCusNamee.Text & " Advanced Payment  ")

                    End If
                    
                End If
        
        
                        
                        
                        
                        
                        
                        
                        
            End If
            
            
        End If

        rs("CustomerTypeID").value = IIf(val(Me.DcCustomerType.BoundText) = 0, Null, val(Me.DcCustomerType.BoundText))
        rs("CountryID").value = IIf(val(Me.DcboCountryID.BoundText) = 0, Null, val(Me.DcboCountryID.BoundText))
        rs("CountryID2").value = IIf(val(Me.DcboCountryID2.BoundText) = 0, Null, val(Me.DcboCountryID2.BoundText))
         rs("Boxmil").value = txtBox.Text
      rs("ZipCode").value = Me.TxtZib.Text
        rs("TypeCustomer").value = val(DcbDigCustomer.ListIndex)
       rs("Map").value = Trim$(Me.TxtMap.Text)
       rs("Entry").value = Trim$(Me.TxtEntry.Text)
       rs("JobName").value = Trim$(Me.txtJob.Text)
       
        rs("GovernmentID").value = IIf(val(Me.DcboGovernmentID.BoundText) = 0, Null, val(Me.DcboGovernmentID.BoundText))
        rs("CityID").value = IIf(val(Me.DcboCityID.BoundText) = 0, Null, val(Me.DcboCityID.BoundText))
        rs("ResponsibleContact").value = Trim$(Me.TxtResponsibleContact.Text)
        rs("Address").value = Trim$(Me.TxtAddress.Text)
        rs("Sex").value = Trim$(Me.CboSex.Text)
        '19 08 2013
        rs("ExpireDateH").value = Txt_DateExpLincH.value
        rs("Company").value = Trim(TxtCompany.Text)
        rs("JobTitle").value = Trim(TxtJobTitle.Text)
        rs("Salary").value = val(TxtSalary.Text)
        rs("JobAddress").value = Trim(TxtJobAddress.Text)
        rs("JobTel").value = Trim(TxtJobTel.Text)
        rs("JobTelConvert").value = Trim(TxtJobTelConvert.Text)
        rs("HomeTel").value = Trim(TxtHomeTel.Text)
        rs("Mobile1").value = Trim(TxtMobile1.Text)
        rs("Mobile2").value = Trim(TXtMobile2.Text)
      
        rs.update

        Dim StrDes As String

        If SystemOptions.UserInterface = ArabicInterface Then
            StrDes = "«·—’Ìœ «·≈›  «ÕÏ ·‹ "
        Else
            StrDes = " Opening Balance For: "
        End If
        
               Dim LngDevID As Long
                Dim Account_Code_dynamic1 As String
         
        If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
       
                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
            
                If Me.OptType(0).value = True Then
                   
                    Account_Code_dynamic1 = get_account_code_branch(109, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
            
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code").value, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText), , val(TxtShowQty1) - val(TxtShowQty1c), val(TxtshowPrice1) - val(TxtshowPrice1C), val(TxtshowPrice2) - val(TxtshowPrice2C), val(TxtSalaries1) - val(TxtSalaries1C), val(TxtSalaries2) - val(TxtSalaries2C)) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText), , val(TxtShowQty1c) - val(TxtShowQty1), val(TxtshowPrice1C) - val(TxtshowPrice1), val(TxtshowPrice2C) - val(TxtshowPrice2), val(TxtSalaries1C) - val(TxtSalaries1), val(TxtSalaries2C) - val(TxtSalaries2)) = False Then
                        GoTo ErrTrap
                    End If
                    
                    '  If ModAccounts.AddNewDev(LngDevID, 2, "a2a1a1", _
                       Val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    '         GoTo ErrTrap
                    ' End If
                ElseIf Me.OptType(1).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(59, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText), , val(TxtShowQty1c) - val(TxtShowQty1), val(TxtshowPrice1C) - val(TxtshowPrice1), val(TxtshowPrice2C) - val(TxtshowPrice2), val(TxtSalaries1C) - val(TxtSalaries1), val(TxtSalaries2C) - val(TxtSalaries2)) = False Then
                        GoTo ErrTrap
                    End If

                    ' If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
                      Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    
                    '       GoTo ErrTrap
                    'End If
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code").value, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText), , val(TxtShowQty1) - val(TxtShowQty1c), val(TxtshowPrice1) - val(TxtshowPrice1C), val(TxtshowPrice2) - val(TxtshowPrice2C), val(TxtSalaries1) - val(TxtSalaries1C), val(TxtSalaries2) - val(TxtSalaries2C)) = False Then
                        GoTo ErrTrap
                    End If
                End If

                '   update_account_opening_balance rs("Account_Code").value
                'update_account_opening_balance Account_Code_dynamic1
                 
            End If
        End If




If SystemOptions.CustomerhavethreeAccounts1 = True Then
' 2
     If Me.OptType1(0).value = True Or Me.OptType1(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
 
                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
            
                If Me.OptType1(0).value = True Then
                   
                    Account_Code_dynamic1 = get_account_code_branch(59, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
            
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code1").value, Round(Me.TxtOpenBalance1.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance1.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
                    '  If ModAccounts.AddNewDev(LngDevID, 2, "a2a1a1", _
                       Val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    '         GoTo ErrTrap
                    ' End If
                ElseIf Me.OptType1(1).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(59, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance1.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    ' If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
                      Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    
                    '       GoTo ErrTrap
                    'End If
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code1").value, Round(Me.TxtOpenBalance1.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If

                '   update_account_opening_balance rs("Account_Code").value
                'update_account_opening_balance Account_Code_dynamic1
                 
            End If
        End If
'3
     If Me.OptType2(0).value = True Or Me.OptType2(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
 
                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
            
                If Me.OptType2(0).value = True Then
                   
                    Account_Code_dynamic1 = get_account_code_branch(59, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
            
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code2").value, Round(Me.TxtOpenBalance2.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance2.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
                    '  If ModAccounts.AddNewDev(LngDevID, 2, "a2a1a1", _
                       Val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    '         GoTo ErrTrap
                    ' End If
                ElseIf Me.OptType2(1).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(59, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance2.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    ' If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
                      Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    
                    '       GoTo ErrTrap
                    'End If
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code2").value, Round(Me.TxtOpenBalance2.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If

                '   update_account_opening_balance rs("Account_Code").value
                'update_account_opening_balance Account_Code_dynamic1
                 
            End If
        End If

End If

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        'update_account_opening_balance Me.DcboDebitSide.BoundText
        ' update_account_opening_balance Me.DcboCreditSide.BoundText
        CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„”«Â„ " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Done do you want new customer"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "    Saved  ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.Text = "R"
        
       ' PassData
       ' If FrmCustemers.Height = 10560 Then
       '       FrmCarAuthontication.TxtClientCode.text = Me.TxtFullcode
       '      FrmCarAuthontication.retInfoCustomer
       '      Unload Me
       ' End If
        
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
            Msg = "Error  In Entry Data"
            End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
       If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Error During Saving"
    End If
    Msg = Msg & CHR(13) & Err.description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
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

    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If
If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  „”«Â„ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  „”«Â„" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  „”«Â„ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  „”«Â„" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ „”«Â„" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„”«Â„Ì‰", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, BolRtl
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New..." & Wrap & "Add New  Data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print the current  data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit the current  data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the current editing or Save the new  data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo..." & Wrap & "Undo in the adding new record" & Wrap & "OR undo editing current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete...." & Wrap & "Delete the current  data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search" & Wrap & "Search for a ..." & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next..." & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last..." & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Shareholders Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "Show Help File", BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

'Private Sub PrintReport()
'    On Error GoTo ErrTrap
'
'    If XPTxtCusID.text <> "" Then
'        Set CusReport = New ClsCustemerReport
'        CusReport.CustemerData XPTxtCusID.text, 1
'    End If
'
'    Exit Sub
'ErrTrap:
'End Sub
Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.Cus_Phone, "
 MySQL = MySQL & "                     dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Remark, dbo.TblCustemers.Type, dbo.TblCustemers.OpenBalance, dbo.TblCustemers.OpenBalanceType,"
MySQL = MySQL & "                      dbo.TblCustemers.OpenBalanceDate, dbo.TblCustemers.CreditLimit, dbo.TblCustemers.Account_Code_As_Client, dbo.TblCustemers.Account_Code_As_Supplier,"
MySQL = MySQL & "                      dbo.TblCustemers.CreditlimitCredit, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail, dbo.TblCustemers.SaleType, dbo.TblCustemers.Account_Code,"
MySQL = MySQL & "                      dbo.TblCustemers.Trans_Discount, dbo.TblCustemers.Trans_DiscountType, dbo.TblCustemers.CountryID, dbo.TblCountriesData.CountryName,"
MySQL = MySQL & "                      dbo.TblCustemers.CityID, dbo.TblCountriesGovernmentsCities.CityName, dbo.TblCustemers.GovernmentID, dbo.TblCountriesGovernments.GovernmentName,"
MySQL = MySQL & "                      dbo.TblCustemers.Address, dbo.TblCustemers.Trans_DiscountPur, dbo.TblCustemers.Trans_DiscountTypePur, dbo.TblCustemers.CountEmp, dbo.TblCustemers.ToTal,"
MySQL = MySQL & "                       dbo.TblCustemers.c1, dbo.TblCustemers.c2, dbo.TblCustemers.Remark2, dbo.TblCustemers.locked, dbo.TblCustemers.parent_account,"
MySQL = MySQL & "                      dbo.TblCustemers.opening_balance_voucher_id, dbo.TblCustemers.DepitInterval, dbo.TblCustemers.CreditInterval, dbo.TblCustemers.DepitIntervalID,"
MySQL = MySQL & "                      dbo.TblCustemers.CreditIntervalID, dbo.TblCustemers.EmpId, dbo.TblCustemers.prifix, dbo.TblCustemers.code, dbo.TblCustemers.Fullcode,"
MySQL = MySQL & "                      dbo.TblCustemers.CustomerandVendor, dbo.TblCustemers.CustomerTypeID, dbo.TblCustemers.BranchId, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblCustemers.CustGID, dbo.TblCustemers.ExpireDateH, dbo.TblCustemers.Company, dbo.TblCustemers.JobTitle,"
MySQL = MySQL & "                      dbo.TblCustemers.Salary, dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.HomeTel,"
 MySQL = MySQL & "                     dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2, dbo.TblCustemers.CountryID2, dbo.TblCustemers.Sex, dbo.TblCustemers.Account_Code1,"
MySQL = MySQL & "                      dbo.TblCustemers.Account_Code2, dbo.TblCustemers.ParentAccount, dbo.TblCustemers.OpenBalanceType1, dbo.TblCustemers.OpenBalance1,"
MySQL = MySQL & "                      dbo.TblCustemers.OpenBalanceType2, dbo.TblCustemers.OpenBalance2, dbo.TblCustemers.ShowQty1, dbo.TblCustemers.showPrice1,"
MySQL = MySQL & "                      dbo.TblCustemers.showPrice2, dbo.TblCustemers.Salaries1, dbo.TblCustemers.Salaries2, dbo.TblCustemers.ShowQty1c, dbo.TblCustemers.showPrice1c,"
MySQL = MySQL & "                      dbo.TblCustemers.showPrice2c, dbo.TblCustemers.Salaries1c, dbo.TblCustemers.Salaries2c, dbo.TblCustemers.Totald, dbo.TblCustemers.Totalc,"
MySQL = MySQL & "                     dbo.TblCustemers.RecordDate , dbo.TblCustemers.balanced , dbo.TblCustemers.balancec, dbo.TblCustemers.TypeCustomer, dbo.TblCustemers.BoxMil, dbo.TblCustemers.ZipCode"
MySQL = MySQL & " FROM         dbo.TblCustemers LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblCustemers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernments ON dbo.TblCustemers.GovernmentID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernmentsCities ON dbo.TblCustemers.CityID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
MySQL = MySQL & "                    dbo.TblCountriesData ON dbo.TblCustemers.CountryID = dbo.TblCountriesData.CountryID"
MySQL = MySQL & " Where (dbo.TblCustemers.CusID =" & val(XPTxtCusID.Text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CartCustomer.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CartCustomer.rpt"
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

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    chkCustomerandVendor.Caption = "Customer / Supplier"
    Label1(2).Caption = "Type"
    Label3.Caption = "Branch"
Cmd(11).Caption = "Attachments."
     
    XPLbl(5).Caption = "G.ID"
    lbl(23).Caption = "Contact person"
    lbl(63).Caption = "RegDate"
    lbl(19).Caption = " type"
    lbl(20).Caption = "Value"
    Label1(1).Caption = "S. Preson"
    lbl(29).Caption = " type"
    lbl(28).Caption = "Value"
    lbl(22).Caption = "State"
    lbl(24).Caption = "Province"
    lbl(25).Caption = " City "
    lbl(26).Caption = "Address"
    lbl(68).Caption = "Entry"
    lbl(67).Caption = "Map"
    Cmd(10).Caption = "Print Card"
    Fra(5).Caption = "Work Address"
    Fra(4).Caption = "Discounts sales invoices"
    Fra(6).Caption = "Discounts purchase invoices"
    lbl(33).Caption = "Parent Acc"
    lbl(69).Caption = "Job"
lbl(60).Caption = "Box Mail"
lbl(61).Caption = "Zip Code"
lbl(62).Caption = "Stars No"
    Me.Caption = "Customers Data"
    EleHeader.Caption = Me.Caption
    XPLbl(1).Caption = "Cust. Code"
    XPLbl(0).Caption = "Cust. Name"
    XPLbl(4).Caption = "Eng. Name"

    lbl(3).Caption = "Phone"
    lbl(2).Caption = "Mobile"
    lbl(1).Caption = "Remarks"
    lbl(0).Caption = "Current Record"
    lbl(4).Caption = "Fax NO."
    lbl(7).Caption = "Credit Limit(Debit)"
    lbl(10).Caption = "Customer Sale Type"
    lbl(11).Caption = "Credit Limit(Credit)"
    lbl(12).Caption = "E Mail"
    Me.Fra(0).Caption = "Open Balance"
    Me.Fra(1).Caption = "Open Balance State"
    Me.Fra(8).Caption = "Checks Under Collected"
    Me.Fra(9).Caption = "Advanced Payments"
    
    OptType(0).Caption = "Debit"
    OptType(1).Caption = "Credit"
    OptType(2).Caption = "Un Sign"
    Me.Fra(3).Caption = "Contact Info."

    lbl(5).Caption = "Balance Value"
    lbl(6).Caption = "Record Date"
'**************************************
    OptType1(0).Caption = "Debit"
    OptType1(1).Caption = "Credit"
    OptType1(2).Caption = "Un Sign"
    lbl(47).Caption = "Balance Value"
    lbl(48).Caption = "Record Date"
    
    
    OptType2(0).Caption = "Debit"
    OptType2(1).Caption = "Credit"
    OptType2(2).Caption = "Un Sign"
    lbl(50).Caption = "Balance Value"
    lbl(49).Caption = "Record Date"
    
     Frame3.Caption = "Gold Data"
     CDMOldContract.Caption = "Old Contract"
     lbl(56).Caption = "Depit"
     lbl(57).Caption = "Credit"
     lbl(51).Caption = "G Weight"
     lbl(52).Caption = "G value"
     
     lbl(53).Caption = "D Value"
     lbl(54).Caption = "Form Value"
     
     lbl(55).Caption = "Inst Value"
     lbl(58).Caption = "Total"
     lbl(59).Caption = "Balance"
     
    '**************************************
    
    Me.Fra(2).Caption = "Current Balance State"
    Me.Cmd(9).Caption = "Customer Balance Report"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"

    With Me.CboSaleType
        .Clear
  
        .AddItem "Retail"
        .AddItem "WholeSale"
 
    End With

    With CboDiscountType
        .Clear
        .AddItem "No"
        .AddItem "Value"
        .AddItem "percentage"
    End With

    With CboDiscountTypePur
        .Clear
        .AddItem "no"
        .AddItem "Value"
        .AddItem "percentage"
    End With

    XPLbl(2).Caption = "Client NO."
    XPLbl(3).Caption = "End User"

    locked.Caption = "locked"
    ALLButton1.Caption = "Reason"
    lbl(32).Caption = "reason"
    lbl(30).Caption = "period"
    lbl(31).Caption = "period"
    XPLbl(12).Caption = "Expire date"
    Me.C1Tab1.TabCaption(0) = "Data"
    Me.C1Tab1.TabCaption(1) = "Specific Data"
    Fra(7).Caption = "Jbb Data"
    lbl(45).Caption = "Nationality"
    lbl(46).Caption = "Sex"
    lbl(36).Caption = "Compqny"
    lbl(37).Caption = "Jbb Title"
    lbl(38).Caption = "Salary"
    lbl(39).Caption = "Jbb Address"
    lbl(40).Caption = "Jbb Tel"
    lbl(41).Caption = "Convert"
    lbl(42).Caption = "Home Tel"
    lbl(43).Caption = "Mob#"
    lbl(44).Caption = "Mob2#"

    With CboSex
        .Clear
        .AddItem "Male"
        .AddItem "Female"
    End With

End Sub

Private Sub ShowCusBalance()
    Dim LngCusID As Long

    LngCusID = val(XPTxtCusID.Text)
    OpenScreen PopUpShowCustomerBalanceScreen, LngCusID, 0
End Sub

Private Sub LoadDataCombos(Optional BolExceptCountries As Boolean = False, _
                           Optional BolExceptGovern As Boolean = False, _
                           Optional BolExceptCities As Boolean = False)
    Set Dcombo = New ClsDataCombos
    Dcombo.GetCountriesNames Me.DcboCountryID2
  
    If BolExceptCountries = False Then
        Dcombo.GetCountriesNames Me.DcboCountryID
        Set cSearch(0) = New clsDCboSearch
        Set cSearch(0).Client = Me.DcboCountryID
    End If

    If BolExceptGovern = False Then
        Dcombo.getCountriesGovernments Me.DcboGovernmentID, val(Me.DcboCountryID.BoundText)
        Set cSearch(1) = New clsDCboSearch
        Set cSearch(1).Client = Me.DcboGovernmentID
    End If

    If BolExceptCities = False Then
        Dcombo.GetCountriesGovernCities Me.DcboCityID, val(Me.DcboCountryID.BoundText), val(Me.DcboGovernmentID.BoundText)
        Set cSearch(2) = New clsDCboSearch
        Set cSearch(2).Client = Me.DcboCityID
    End If

    Dcombo.GetCustomerType Me.DcCustomerType

    Dcombo.GetBranches Dcbranch
  
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.Dcbranch.Enabled = True
    End If

End Sub

Private Sub XPTxtCusName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtCusNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

