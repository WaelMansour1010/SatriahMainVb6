VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmCustemers 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·⁄„·«¡"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13980
   HelpContextID   =   50
   Icon            =   "FrmMembers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   13980
   Begin VB.TextBox TxtVATNO 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5520
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1800
      Width           =   3075
   End
   Begin VB.CommandButton CDMOldContract 
      Caption         =   "›Ê« Ì— Ê⁄ﬁÊœ ”«»ﬁ…"
      Height          =   375
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   252
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄„·«¡ «·ÕÃ Ê«·⁄„—…"
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
      Index           =   10
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   247
      Top             =   2280
      Width           =   3405
      Begin VB.CheckBox TypeOmrh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄„—…"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   251
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox TypeHaj 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÕÃ"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   250
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton HajEnter_Out 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Œ«—ÃÌ"
         Height          =   255
         Index           =   1
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   249
         Top             =   720
         Width           =   1605
      End
      Begin VB.OptionButton HajEnter_Out 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "œ«Œ·Ì"
         Height          =   255
         Index           =   0
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   248
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.ComboBox CboSaleType 
      Height          =   315
      Left            =   2370
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   219
      Top             =   1800
      Width           =   1995
   End
   Begin VB.TextBox TxtFullcode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   14400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   208
      Top             =   960
      Width           =   1125
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Index           =   0
      Left            =   2880
      TabIndex        =   192
      Top             =   8760
      Width           =   7935
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   5
         Left            =   1440
         Picture         =   "FrmMembers.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   204
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   4
         Left            =   2880
         Picture         =   "FrmMembers.frx":07D5
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   3
         Left            =   3600
         Picture         =   "FrmMembers.frx":0D2D
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   2
         Left            =   5040
         Picture         =   "FrmMembers.frx":11E6
         Style           =   1  'Graphical
         TabIndex        =   201
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   1
         Left            =   3240
         Picture         =   "FrmMembers.frx":16B6
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmMembers.frx":1B57
         Height          =   555
         Index           =   0
         Left            =   7200
         Picture         =   "FrmMembers.frx":8E89
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmMembers.frx":9430
         Height          =   555
         Index           =   6
         Left            =   5760
         Picture         =   "FrmMembers.frx":10762
         Style           =   1  'Graphical
         TabIndex        =   198
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmMembers.frx":10C03
         Height          =   555
         Index           =   7
         Left            =   4320
         Picture         =   "FrmMembers.frx":17F35
         Style           =   1  'Graphical
         TabIndex        =   197
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   8
         Left            =   2160
         Picture         =   "FrmMembers.frx":187C5
         Style           =   1  'Graphical
         TabIndex        =   196
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmMembers.frx":18CAA
         Height          =   555
         Index           =   9
         Left            =   720
         Picture         =   "FrmMembers.frx":1FFDC
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmMembers.frx":204FC
         Height          =   555
         Index           =   10
         Left            =   6480
         Picture         =   "FrmMembers.frx":20AE3
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmMembers.frx":210CA
         Height          =   555
         Index           =   11
         Left            =   0
         Picture         =   "FrmMembers.frx":283FC
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Œ’Ê„«  Œ«’… ··⁄„Ì· ›Ï ›Ê« Ì— «·‘—«¡"
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   125
      Top             =   2280
      Width           =   2925
      Begin VB.ComboBox CboDiscountTypePur 
         Height          =   315
         Left            =   270
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   127
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtDiscountValuePur 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         RightToLeft     =   -1  'True
         TabIndex        =   126
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
         Top             =   690
         Width           =   195
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Œ’Ê„«  Œ«’… ··⁄„Ì· ›Ï ›Ê« Ì— «·»Ì⁄"
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
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   119
      Top             =   2280
      Width           =   2925
      Begin VB.TextBox TxtDiscountValue 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         RightToLeft     =   -1  'True
         TabIndex        =   121
         Top             =   600
         Width           =   1425
      End
      Begin VB.ComboBox CboDiscountType 
         Height          =   315
         Left            =   270
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   120
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.TextBox txtCustGID 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2880
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1020
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
      TabIndex        =   58
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   11310
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   600
      Width           =   1125
   End
   Begin VB.TextBox txtopening_balance_voucher_id 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox XPTxtCusNamee 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5520
      MaxLength       =   50
      TabIndex        =   10
      Top             =   1020
      Width           =   3045
   End
   Begin VB.TextBox c2 
      Height          =   345
      Left            =   9600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox c1 
      Height          =   345
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„«  ⁄‰ «·⁄„Ì·"
      Height          =   555
      Left            =   3750
      RightToLeft     =   -1  'True
      TabIndex        =   40
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
      Height          =   2145
      Index           =   3
      Left            =   9630
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   1260
      Width           =   4365
      Begin VB.TextBox txtJob 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   216
         Top             =   600
         Width           =   2805
      End
      Begin VB.TextBox TxtResponsibleContact 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   210
         Width           =   2805
      End
      Begin VB.TextBox TxtFaxNumber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1710
         Width           =   2805
      End
      Begin VB.TextBox XPTxtPhone 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1020
         Width           =   2805
      End
      Begin VB.TextBox XPTxtmobile 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         MaxLength       =   14
         RightToLeft     =   -1  'True
         TabIndex        =   12
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
         TabIndex        =   217
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
         TabIndex        =   47
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
         TabIndex        =   39
         Top             =   1710
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·„Ê»«Ì·"
         Height          =   255
         Index           =   2
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1395
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Â« › «·«—÷Ì"
         Height          =   315
         Index           =   3
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1050
         Width           =   1335
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      Height          =   345
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   19
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
      TabIndex        =   0
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
      TabIndex        =   7
      Top             =   45
      Visible         =   0   'False
      Width           =   1455
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   705
      Left            =   0
      TabIndex        =   14
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
      Caption         =   "»Ì«‰«  «·⁄„·«¡  "
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
         TabIndex        =   16
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmMembers.frx":28F90
         ColorButton     =   16777215
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
         TabIndex        =   18
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmMembers.frx":2932A
         ColorButton     =   16777215
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
         TabIndex        =   15
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmMembers.frx":296C4
         ColorButton     =   16777215
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
         TabIndex        =   17
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmMembers.frx":29A5E
         ColorButton     =   16777215
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
         Picture         =   "FrmMembers.frx":29DF8
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   12915
      TabIndex        =   25
      Top             =   8310
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
      Left            =   12195
      TabIndex        =   26
      Top             =   8310
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
      Left            =   11475
      TabIndex        =   27
      Top             =   8310
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
      Left            =   10755
      TabIndex        =   28
      Top             =   8310
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
      Left            =   9915
      TabIndex        =   29
      Top             =   8310
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
      Left            =   9090
      TabIndex        =   30
      Top             =   8310
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
      Left            =   2760
      TabIndex        =   31
      Top             =   8310
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
      Left            =   8250
      TabIndex        =   32
      Top             =   8310
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
      Left            =   7230
      TabIndex        =   33
      Top             =   8310
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
      TabIndex        =   34
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
      Left            =   10080
      TabIndex        =   57
      Top             =   600
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   5520
      TabIndex        =   59
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
      Height          =   4935
      Left            =   120
      TabIndex        =   62
      Top             =   3360
      Width           =   13905
      _cx             =   24527
      _cy             =   8705
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
      Caption         =   "»Ì«‰«  «”«”Ì…|»Ì«‰«  „ Œ’’Â|ÃÂ«  «· ⁄«„·|”Ì«—« | ›«’Ì· Œ«’…"
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   4515
         Left            =   15150
         TabIndex        =   281
         TabStop         =   0   'False
         Top             =   45
         Width           =   13815
         _cx             =   24368
         _cy             =   7964
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   1935
            Index           =   1
            Left            =   0
            TabIndex        =   282
            TabStop         =   0   'False
            Top             =   120
            Width           =   13815
            _cx             =   24368
            _cy             =   3413
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
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00E2E9E9&
               Height          =   1275
               Left            =   0
               ScaleHeight     =   1215
               ScaleWidth      =   4575
               TabIndex        =   295
               Top             =   270
               Width           =   4635
               Begin VB.Frame Frame4 
                  BackColor       =   &H00E2E9E9&
                  Height          =   615
                  Index           =   1
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   305
                  Top             =   -30
                  Width           =   4485
                  Begin VB.TextBox txtBoardNo 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   3060
                     RightToLeft     =   -1  'True
                     TabIndex        =   344
                     Top             =   270
                     Width           =   1395
                  End
                  Begin VB.TextBox txtNum4 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   0
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   313
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox txtLetter4 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1560
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   312
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox txtNum3 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   360
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   311
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox txtNum2 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   720
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   310
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox txtNum1 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1080
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   309
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox txtLetter3 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   308
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox txtLetter2 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   2280
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   307
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox txtLetter1 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   2640
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   306
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "En"
                     Height          =   285
                     Index           =   79
                     Left            =   3630
                     RightToLeft     =   -1  'True
                     TabIndex        =   346
                     Top             =   60
                     Width           =   825
                  End
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H00E2E9E9&
                  Height          =   645
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   296
                  Top             =   570
                  Width           =   4485
                  Begin VB.TextBox txtnBoardNo 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   3060
                     RightToLeft     =   -1  'True
                     TabIndex        =   345
                     Top             =   270
                     Width           =   1395
                  End
                  Begin VB.TextBox ntxtLetter1 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   2640
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   304
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox ntxtLetter2 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   2280
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   303
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox ntxtLetter3 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1920
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   302
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox ntxtNum1 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1080
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   301
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox ntxtNum2 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   720
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   300
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox ntxtNum3 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   360
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   299
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox ntxtLetter4 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1560
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   298
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.TextBox ntxtNum4 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   0
                     MaxLength       =   1
                     RightToLeft     =   -1  'True
                     TabIndex        =   297
                     Top             =   270
                     Width           =   405
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄"
                     Height          =   285
                     Index           =   80
                     Left            =   3600
                     RightToLeft     =   -1  'True
                     TabIndex        =   347
                     Top             =   60
                     Width           =   795
                  End
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„À«· "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   315
                  Index           =   7
                  Left            =   1290
                  RightToLeft     =   -1  'True
                  TabIndex        =   315
                  Top             =   -240
                  Width           =   465
               End
               Begin VB.Label XPLbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ » Ã  1 2 3"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Index           =   10
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   314
                  Top             =   -240
                  Width           =   1185
               End
            End
            Begin VB.TextBox txtCount 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4710
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   293
               Top             =   120
               Width           =   1128
            End
            Begin VB.ComboBox cbModel 
               Height          =   315
               ItemData        =   "FrmMembers.frx":2AAC2
               Left            =   10800
               List            =   "FrmMembers.frx":2AAC4
               RightToLeft     =   -1  'True
               TabIndex        =   292
               Top             =   840
               Width           =   1455
            End
            Begin VB.TextBox txtChassis 
               Height          =   315
               Left            =   9480
               MaxLength       =   50
               TabIndex        =   291
               Top             =   120
               Width           =   2775
            End
            Begin VB.ComboBox DcCity 
               Height          =   315
               ItemData        =   "FrmMembers.frx":2AAC6
               Left            =   6900
               List            =   "FrmMembers.frx":2AAC8
               RightToLeft     =   -1  'True
               TabIndex        =   290
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox txtDriverName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   8520
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   289
               Top             =   480
               Width           =   4155
            End
            Begin VB.TextBox txtDriverTel 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4710
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   288
               Top             =   480
               Width           =   1128
            End
            Begin VB.TextBox accessoryTxt 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6900
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   287
               Top             =   1560
               Width           =   5565
            End
            Begin VB.TextBox TxtPrice 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4710
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   286
               Top             =   1200
               Width           =   1128
            End
            Begin VB.TextBox TxtPartPrice 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4710
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   285
               Top             =   1560
               Width           =   1128
            End
            Begin VB.ComboBox DcbTypTrans 
               Height          =   315
               ItemData        =   "FrmMembers.frx":2AACA
               Left            =   6900
               List            =   "FrmMembers.frx":2AACC
               RightToLeft     =   -1  'True
               TabIndex        =   284
               Top             =   1200
               Width           =   5565
            End
            Begin VB.ComboBox DcbTypTrans1 
               Height          =   315
               ItemData        =   "FrmMembers.frx":2AACE
               Left            =   0
               List            =   "FrmMembers.frx":2AAD0
               RightToLeft     =   -1  'True
               TabIndex        =   283
               Top             =   480
               Visible         =   0   'False
               Width           =   3252
            End
            Begin MSDataListLib.DataCombo dcBrand 
               Height          =   315
               Left            =   6900
               TabIndex        =   316
               Top             =   120
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   14
               Left            =   2880
               TabIndex        =   317
               Top             =   1530
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈÷«›…"
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
               ButtonImage     =   "FrmMembers.frx":2AAD2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   13
               Left            =   1560
               TabIndex        =   318
               Top             =   1530
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› "
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
               ButtonImage     =   "FrmMembers.frx":2AE6C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   15
               Left            =   120
               TabIndex        =   319
               Top             =   1530
               Width           =   1200
               _ExtentX        =   2117
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
               ButtonImage     =   "FrmMembers.frx":4A056
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo DcbCarModel 
               Height          =   315
               Left            =   8520
               TabIndex        =   333
               Top             =   840
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbColor 
               Height          =   315
               Left            =   6900
               TabIndex        =   334
               Top             =   840
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «··ÊÕ…"
               Height          =   315
               Index           =   24
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   335
               Top             =   -30
               Width           =   1065
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«··Ê‰"
               Height          =   315
               Index           =   23
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   332
               Top             =   840
               Width           =   1185
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÿ—«“"
               Height          =   315
               Index           =   22
               Left            =   9480
               RightToLeft     =   -1  'True
               TabIndex        =   331
               Top             =   840
               Width           =   1185
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ·Ì›Ê‰ «·”«∆ﬁ"
               Height          =   315
               Index           =   21
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   279
               Top             =   480
               Width           =   945
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„⁄œ· «·«—ﬂ«»"
               Height          =   315
               Index           =   20
               Left            =   5940
               RightToLeft     =   -1  'True
               TabIndex        =   280
               Top             =   840
               Width           =   885
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·„ﬁ«⁄œ"
               Height          =   315
               Index           =   19
               Left            =   5700
               RightToLeft     =   -1  'True
               TabIndex        =   330
               Top             =   120
               Width           =   1065
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ÊœÌ·"
               Height          =   312
               Index           =   13
               Left            =   12477
               RightToLeft     =   -1  'True
               TabIndex        =   329
               Top             =   840
               Width           =   1188
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·‘«”Ì…"
               Height          =   315
               Index           =   9
               Left            =   12480
               RightToLeft     =   -1  'True
               TabIndex        =   328
               Top             =   120
               Width           =   1185
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„«—ﬂ…"
               Height          =   225
               Index           =   77
               Left            =   8550
               RightToLeft     =   -1  'True
               TabIndex        =   327
               Top             =   120
               Width           =   765
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·”«∆ﬁ"
               Height          =   312
               Index           =   11
               Left            =   12477
               RightToLeft     =   -1  'True
               TabIndex        =   326
               Top             =   480
               Width           =   1188
            End
            Begin VB.Label txtRate 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               Caption         =   "1.3"
               Height          =   315
               Left            =   4710
               RightToLeft     =   -1  'True
               TabIndex        =   324
               Top             =   840
               Width           =   1125
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„·Õﬁ"
               Height          =   315
               Index           =   15
               Left            =   12720
               RightToLeft     =   -1  'True
               TabIndex        =   323
               Top             =   1560
               Width           =   945
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Height          =   315
               Index           =   16
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   322
               Top             =   1200
               Width           =   945
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„… ··„·Õﬁ"
               Height          =   315
               Index           =   17
               Left            =   5640
               RightToLeft     =   -1  'True
               TabIndex        =   321
               Top             =   1560
               Width           =   1185
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
               Height          =   315
               Index           =   18
               Left            =   12720
               RightToLeft     =   -1  'True
               TabIndex        =   320
               Top             =   1200
               Width           =   945
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· »⁄Ì…"
               Height          =   315
               Index           =   8
               Left            =   12720
               RightToLeft     =   -1  'True
               TabIndex        =   325
               Top             =   1200
               Visible         =   0   'False
               Width           =   945
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   2475
            Left            =   0
            TabIndex        =   294
            Top             =   2040
            Width           =   13785
            _cx             =   24315
            _cy             =   4366
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16776960
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
            Rows            =   50
            Cols            =   25
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmMembers.frx":69240
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   4515
         Index           =   0
         Left            =   14850
         TabIndex        =   226
         TabStop         =   0   'False
         Top             =   45
         Width           =   13815
         _cx             =   24368
         _cy             =   7964
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
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   2025
            Left            =   120
            TabIndex        =   227
            Top             =   240
            Width           =   13515
            Begin VB.ComboBox STDUDENTStatusID 
               Height          =   315
               ItemData        =   "FrmMembers.frx":695E3
               Left            =   3540
               List            =   "FrmMembers.frx":695E5
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   276
               Top             =   1680
               Visible         =   0   'False
               Width           =   1995
            End
            Begin VB.TextBox TxtIQAMA 
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
               Left            =   7200
               TabIndex        =   268
               Top             =   1320
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.TextBox TxtPassport 
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
               Left            =   10440
               TabIndex        =   267
               Top             =   1320
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.ComboBox DcbLevel 
               Height          =   315
               ItemData        =   "FrmMembers.frx":695E7
               Left            =   3540
               List            =   "FrmMembers.frx":695E9
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   263
               Top             =   960
               Visible         =   0   'False
               Width           =   1995
            End
            Begin VB.ComboBox DcbFM 
               Height          =   315
               ItemData        =   "FrmMembers.frx":695EB
               Left            =   120
               List            =   "FrmMembers.frx":695ED
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   261
               Top             =   600
               Visible         =   0   'False
               Width           =   1995
            End
            Begin VB.ComboBox DcbCurrClass 
               Height          =   315
               ItemData        =   "FrmMembers.frx":695EF
               Left            =   3540
               List            =   "FrmMembers.frx":695F1
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   259
               Top             =   600
               Visible         =   0   'False
               Width           =   1995
            End
            Begin VB.ComboBox DcbFirstClass 
               Height          =   315
               ItemData        =   "FrmMembers.frx":695F3
               Left            =   7740
               List            =   "FrmMembers.frx":695F5
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   257
               Top             =   600
               Visible         =   0   'False
               Width           =   1635
            End
            Begin VB.TextBox TxtMangerName 
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
               Left            =   7200
               TabIndex        =   238
               Top             =   600
               Width           =   4695
            End
            Begin VB.TextBox TxtName 
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
               Left            =   7200
               TabIndex        =   231
               Top             =   120
               Width           =   4695
            End
            Begin VB.TextBox TxtNameE 
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
               Left            =   120
               TabIndex        =   230
               Top             =   165
               Width           =   5775
            End
            Begin VB.ComboBox Combo2 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmMembers.frx":695F7
               Left            =   2280
               List            =   "FrmMembers.frx":69607
               Style           =   2  'Dropdown List
               TabIndex        =   229
               Top             =   3030
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox txtid1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   0
               MaxLength       =   20
               TabIndex        =   228
               Top             =   0
               Visible         =   0   'False
               Width           =   945
            End
            Begin ImpulseButton.ISButton ISButton2 
               Height          =   315
               Left            =   120
               TabIndex        =   232
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   960
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               Caption         =   "«÷«›…"
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
               ButtonImage     =   "FrmMembers.frx":69620
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin MSDataListLib.DataCombo DcbClass 
               Height          =   315
               Left            =   7200
               TabIndex        =   266
               Top             =   960
               Visible         =   0   'False
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCNAtionalID 
               Height          =   315
               Left            =   3540
               TabIndex        =   272
               Top             =   1320
               Visible         =   0   'False
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DBRegisterDate 
               Height          =   330
               Left            =   10440
               TabIndex        =   273
               Top             =   1680
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   582
               _Version        =   393216
               CalendarBackColor=   12648447
               CustomFormat    =   "yyyy/M/d"
               Format          =   129630211
               CurrentDate     =   38718
            End
            Begin MSComCtl2.DTPicker DBENDDATE 
               Height          =   345
               Left            =   7200
               TabIndex        =   278
               Top             =   1680
               Visible         =   0   'False
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   129630209
               CurrentDate     =   38784
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ«·Â"
               Height          =   285
               Index           =   14
               Left            =   5640
               TabIndex        =   277
               Top             =   1680
               Visible         =   0   'False
               Width           =   1410
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·„€«œ—…"
               Height          =   285
               Index           =   76
               Left            =   9000
               RightToLeft     =   -1  'True
               TabIndex        =   275
               Top             =   1680
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «· ”ÃÌ·"
               Height          =   285
               Index           =   75
               Left            =   12000
               RightToLeft     =   -1  'True
               TabIndex        =   274
               Top             =   1680
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ã‰”Ì…"
               Height          =   285
               Index           =   13
               Left            =   5640
               TabIndex        =   271
               Top             =   1320
               Visible         =   0   'False
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·«ﬁ«„…"
               Height          =   285
               Index           =   12
               Left            =   9120
               TabIndex        =   270
               Top             =   1320
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÃÊ«“"
               Height          =   285
               Index           =   11
               Left            =   12000
               TabIndex        =   269
               Top             =   1320
               Visible         =   0   'False
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„—Õ·…"
               Height          =   285
               Index           =   10
               Left            =   5640
               TabIndex        =   265
               Top             =   960
               Visible         =   0   'False
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›’·"
               Height          =   285
               Index           =   9
               Left            =   12000
               TabIndex        =   264
               Top             =   960
               Visible         =   0   'False
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "M/F"
               Height          =   285
               Index           =   8
               Left            =   2160
               TabIndex        =   262
               Top             =   600
               Visible         =   0   'False
               Width           =   930
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›’· «·œ—«”Ì «·Õ«·Ì"
               Height          =   285
               Index           =   7
               Left            =   5580
               TabIndex        =   260
               Top             =   600
               Visible         =   0   'False
               Width           =   1530
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ê· ›’· œ—«”Ì"
               Height          =   285
               Index           =   6
               Left            =   5640
               TabIndex        =   258
               Top             =   600
               Visible         =   0   'False
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ ⁄—»Ì"
               Height          =   285
               Index           =   5
               Left            =   11925
               TabIndex        =   235
               Top             =   150
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   4
               Left            =   5925
               TabIndex        =   234
               Top             =   150
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·‘Œ’ «·„”∆Ê·"
               Height          =   285
               Index           =   3
               Left            =   12000
               TabIndex        =   233
               Top             =   600
               Width           =   1410
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   1755
            Left            =   120
            TabIndex        =   236
            Top             =   2280
            Width           =   13515
            _cx             =   23839
            _cy             =   3096
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
            Rows            =   50
            Cols            =   19
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmMembers.frx":6FE82
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   225
            Index           =   12
            Left            =   11640
            TabIndex        =   237
            Top             =   4080
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   397
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
            ButtonImage     =   "FrmMembers.frx":70178
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   4515
         Left            =   14550
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   45
         Width           =   13815
         _cx             =   24368
         _cy             =   7964
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
         Begin VB.CheckBox chkIsNotCommission 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·« Ì „ «Õ ”«» ⁄„Ê·Â ⁄·ÌÂ "
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   357
            Top             =   420
            Width           =   2115
         End
         Begin VB.Frame Frame3 
            Caption         =   "»Ì«‰«  Œ«’… »«·œÂ»"
            ForeColor       =   &H000000FF&
            Height          =   3015
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   162
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
               TabIndex        =   183
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
               TabIndex        =   182
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
               TabIndex        =   181
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
               TabIndex        =   180
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
               TabIndex        =   179
               Top             =   360
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice1C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Top             =   720
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice2C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   1080
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries1C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   1440
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries2C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   175
               Top             =   1800
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   172
               Top             =   1800
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   169
               Top             =   1440
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   1080
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   165
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
               TabIndex        =   163
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
               TabIndex        =   185
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
               TabIndex        =   184
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
               TabIndex        =   174
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
               TabIndex        =   173
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
               TabIndex        =   171
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
               TabIndex        =   170
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
               TabIndex        =   168
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
               TabIndex        =   166
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
               TabIndex        =   164
               Top             =   360
               Width           =   1215
            End
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
            TabIndex        =   91
            Top             =   240
            Width           =   6135
            Begin VB.ComboBox CboSex 
               Height          =   315
               ItemData        =   "FrmMembers.frx":70712
               Left            =   3000
               List            =   "FrmMembers.frx":70714
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   113
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TxtMobile2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3000
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   3480
               Width           =   1485
            End
            Begin VB.TextBox TxtMobile1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   3120
               Width           =   1485
            End
            Begin VB.TextBox TxtHomeTel 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3000
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   3120
               Width           =   1485
            End
            Begin VB.TextBox TxtJobTelConvert 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   2760
               Width           =   1485
            End
            Begin VB.TextBox TxtJobTel 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3000
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   102
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
               TabIndex        =   99
               Top             =   2160
               Width           =   4425
            End
            Begin VB.TextBox TxtSalary 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2400
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   1680
               Width           =   2085
            End
            Begin VB.TextBox TXTJobTitle 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   1320
               Width           =   4365
            End
            Begin VB.TextBox TxtCompany 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   960
               Width           =   4365
            End
            Begin MSDataListLib.DataCombo DcboCountryID2 
               Height          =   315
               Left            =   1440
               TabIndex        =   114
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
               TabIndex        =   112
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
               TabIndex        =   111
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
               TabIndex        =   109
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
               TabIndex        =   107
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
               TabIndex        =   105
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
               TabIndex        =   103
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
               TabIndex        =   101
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
               TabIndex        =   100
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
               TabIndex        =   98
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
               TabIndex        =   96
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
               TabIndex        =   94
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
               TabIndex        =   92
               Top             =   240
               Width           =   825
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4515
         Index           =   2
         Left            =   45
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   45
         Width           =   13815
         _cx             =   24368
         _cy             =   7964
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5955
            Index           =   1
            Left            =   -600
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   -120
            Width           =   15105
            _cx             =   26644
            _cy             =   10504
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
            Begin VB.TextBox txtBranchName 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   630
               RightToLeft     =   -1  'True
               TabIndex        =   342
               Top             =   900
               Width           =   2565
            End
            Begin VB.ComboBox CboPayMentType 
               Height          =   315
               Left            =   11310
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   337
               Top             =   120
               Width           =   1560
            End
            Begin VB.TextBox TxtBankIBAN 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   720
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   245
               Top             =   3120
               Width           =   1788
            End
            Begin VB.TextBox txtBankAccount 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3627
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   243
               Top             =   3120
               Width           =   1788
            End
            Begin VB.TextBox TxtBankCode 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   720
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   241
               Top             =   2760
               Width           =   1788
            End
            Begin VB.TextBox txtBankName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3627
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   239
               Top             =   2760
               Width           =   1788
            End
            Begin VB.TextBox TxtBankAddress 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   720
               MaxLength       =   50
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   224
               Top             =   3480
               Width           =   4695
            End
            Begin VB.CheckBox creditlocked 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·€«¡ «· ⁄«„· «·«Ã·"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   8160
               RightToLeft     =   -1  'True
               TabIndex        =   223
               Top             =   3960
               Width           =   1695
            End
            Begin VB.CheckBox locked 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ìﬁ«› «· ⁄«„·"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   8520
               RightToLeft     =   -1  'True
               TabIndex        =   221
               Top             =   3720
               Width           =   1335
            End
            Begin VB.TextBox TxtEntry 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10080
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   214
               Top             =   480
               Width           =   2775
            End
            Begin VB.TextBox TxtMap 
               Alignment       =   1  'Right Justify
               Height          =   435
               Left            =   720
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   210
               Top             =   4200
               Width           =   4695
            End
            Begin VB.ComboBox DcbDigCustomer 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   10080
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   191
               Top             =   1920
               Width           =   2775
            End
            Begin VB.TextBox TxtZib 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Top             =   1560
               Width           =   2775
            End
            Begin VB.TextBox TxtBox 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   186
               Top             =   1200
               Width           =   2775
            End
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
               Height          =   1875
               Index           =   5
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   2310
               Width           =   4125
               Begin VB.TextBox TxtAddress 
                  Alignment       =   1  'Right Justify
                  Height          =   705
                  Left            =   150
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   156
                  Top             =   1140
                  Width           =   2625
               End
               Begin MSDataListLib.DataCombo DcboCountryID 
                  Height          =   315
                  Left            =   150
                  TabIndex        =   157
                  Top             =   120
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
                  TabIndex        =   1
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
                  TabIndex        =   2
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
                  Caption         =   "«·œÊ·…"
                  Height          =   225
                  Index           =   22
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   161
                  Top             =   210
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
                  TabIndex        =   160
                  Top             =   510
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
                  TabIndex        =   159
                  Top             =   840
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄‰Ê«‰ »«· ›’Ì·"
                  Height          =   585
                  Index           =   26
                  Left            =   3030
                  RightToLeft     =   -1  'True
                  TabIndex        =   158
                  Top             =   1140
                  Width           =   765
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "—’Ìœ „œ›Ê⁄«  „ﬁœ„…  ··⁄„Ì·"
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
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   147
               Top             =   1320
               Width           =   2745
               Begin VB.TextBox TxtOpenBalance2 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   151
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
                  TabIndex        =   150
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
                  TabIndex        =   149
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
                  TabIndex        =   148
                  Top             =   210
                  Width           =   765
               End
               Begin MSComCtl2.DTPicker Dtp2 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   152
                  Top             =   870
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   155320323
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
                  TabIndex        =   154
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
                  TabIndex        =   153
                  Top             =   930
                  Width           =   1215
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "—’Ìœ ‘Ìﬂ«   Õ  «· Õ’Ì· ··⁄„Ì·"
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
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   1320
               Width           =   2745
               Begin VB.OptionButton OptType1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   1950
                  RightToLeft     =   -1  'True
                  TabIndex        =   143
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
                  TabIndex        =   142
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
                  TabIndex        =   141
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   1005
               End
               Begin VB.TextBox TxtOpenBalance1 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   510
                  Width           =   1365
               End
               Begin MSComCtl2.DTPicker Dtp1 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   144
                  Top             =   870
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   155320323
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
                  TabIndex        =   146
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
                  TabIndex        =   145
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
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   1320
               Width           =   2745
               Begin VB.TextBox TxtOpenBalance 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
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
                  TabIndex        =   134
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
                  TabIndex        =   133
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
                  TabIndex        =   132
                  Top             =   240
                  Width           =   765
               End
               Begin MSComCtl2.DTPicker Dtp 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   136
                  Top             =   870
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   155320323
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
                  TabIndex        =   138
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
                  TabIndex        =   137
                  Top             =   930
                  Width           =   1215
               End
            End
            Begin VB.TextBox TxtE_mail 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   840
               Width           =   2775
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
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   240
               Width           =   5895
               Begin VB.TextBox TxtCreditlimitCredit 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2670
                  MaxLength       =   8
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   540
                  Width           =   1395
               End
               Begin VB.TextBox TxtCreditLimit 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2670
                  MaxLength       =   8
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   180
                  Width           =   1395
               End
               Begin VB.TextBox TxtDepitInterval 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   240
                  Width           =   495
               End
               Begin VB.TextBox TxtCreditInterval 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   600
                  Width           =   495
               End
               Begin VB.ComboBox dcDepitIntervalID 
                  Height          =   315
                  ItemData        =   "FrmMembers.frx":70716
                  Left            =   120
                  List            =   "FrmMembers.frx":70718
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   240
                  Width           =   975
               End
               Begin VB.ComboBox dcCreditIntervalID 
                  Height          =   315
                  ItemData        =   "FrmMembers.frx":7071A
                  Left            =   120
                  List            =   "FrmMembers.frx":7071C
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
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
                  TabIndex        =   85
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
                  TabIndex        =   84
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
                  TabIndex        =   83
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
                  TabIndex        =   82
                  Top             =   600
                  Width           =   885
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "—’Ìœ «·⁄„Ì· «·Õ«·Ï"
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
               Left            =   6600
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   3000
               Width           =   3375
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   435
                  Index           =   9
                  Left            =   120
                  TabIndex        =   72
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
                  TabIndex        =   74
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
                  TabIndex        =   73
                  Top             =   240
                  Width           =   1785
               End
            End
            Begin VB.TextBox XPMTxtRemarks 
               Alignment       =   1  'Right Justify
               Height          =   435
               Left            =   6720
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   69
               Top             =   4200
               Width           =   5985
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
               Left            =   -3900
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   9360
               Width           =   2160
            End
            Begin MSDataListLib.DataCombo DboParentAccount 
               Height          =   315
               Left            =   720
               TabIndex        =   86
               Top             =   3840
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ALLButtonS.ALLButton ALLButton1 
               Height          =   255
               Left            =   6600
               TabIndex        =   218
               Top             =   3720
               Width           =   1335
               _ExtentX        =   2355
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
               MICON           =   "FrmMembers.frx":7071E
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ALLButtonS.ALLButton ALLButton2 
               Height          =   255
               Left            =   6720
               TabIndex        =   256
               Top             =   2640
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   "⁄—÷ «·—’Ìœ"
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
               MICON           =   "FrmMembers.frx":7073A
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSDataListLib.DataCombo dcGroups 
               Height          =   315
               Left            =   630
               TabIndex        =   338
               Top             =   300
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcClass 
               Height          =   315
               Left            =   630
               TabIndex        =   341
               Top             =   600
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·›—⁄"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   2
               Left            =   3360
               TabIndex        =   343
               Top             =   930
               Width           =   690
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«· ’‰Ì›"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   1
               Left            =   3330
               TabIndex        =   340
               Top             =   600
               Width           =   690
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„Ã„Ê⁄…"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   0
               Left            =   3330
               TabIndex        =   339
               Top             =   300
               Width           =   690
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·œ›⁄"
               Height          =   315
               Index           =   78
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   336
               Top             =   150
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·«Ì»«‰"
               Height          =   285
               Index           =   72
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   246
               Top             =   3150
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ”«» «·»‰ﬂ"
               Height          =   315
               Index           =   73
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   244
               Top             =   3120
               Width           =   1365
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—„“ «·»‰ﬂ"
               Height          =   315
               Index           =   70
               Left            =   2190
               RightToLeft     =   -1  'True
               TabIndex        =   242
               Top             =   2790
               Width           =   1365
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·»‰ﬂ"
               Height          =   315
               Index           =   71
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   240
               Top             =   2790
               Width           =   1365
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄‰Ê«‰ «·»‰ﬂ"
               Height          =   315
               Index           =   74
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   225
               Top             =   3510
               Width           =   1365
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ«Œ·Ì"
               Height          =   315
               Index           =   68
               Left            =   12990
               RightToLeft     =   -1  'True
               TabIndex        =   215
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Œ—«∆ÿ ÃÊÃ·"
               Height          =   315
               Index           =   67
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   213
               Top             =   4200
               Width           =   1365
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   285
               Index           =   66
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   212
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
               TabIndex        =   211
               Top             =   21600
               Width           =   585
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·‰ÃÊ„"
               Height          =   285
               Index           =   62
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   190
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—„“ «·»—ÌœÌ"
               Height          =   285
               Index           =   61
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   189
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’‰œÊﬁ »—Ìœ"
               Height          =   285
               Index           =   60
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   188
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·»—Ìœ «·≈·ﬂ —Ê‰Ï"
               Height          =   285
               Index           =   12
               Left            =   12990
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ”«» «·—∆Ì”Ì"
               Height          =   315
               Index           =   33
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   3840
               Width           =   1365
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   285
               Index           =   1
               Left            =   13470
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   4320
               Width           =   585
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
               Left            =   13695
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   960
               Width           =   840
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   315
            Index           =   34
            Left            =   8400
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   90
            Width           =   1125
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   4515
         Left            =   15450
         TabIndex        =   348
         TabStop         =   0   'False
         Top             =   45
         Width           =   13815
         _cx             =   24368
         _cy             =   7964
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
         Begin VSFlex8UCtl.VSFlexGrid fg 
            Height          =   1545
            Left            =   60
            TabIndex        =   349
            Top             =   780
            Width           =   13620
            _cx             =   24024
            _cy             =   2725
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
            SelectionMode   =   0
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
            FormatString    =   $"FrmMembers.frx":70756
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
            WallPaperAlignment=   0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSDataListLib.DataCombo cmbSizesNames 
            Height          =   315
            Left            =   10260
            TabIndex        =   350
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   180
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker txtDateSize 
            Height          =   330
            Left            =   5910
            TabIndex        =   352
            Top             =   150
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CustomFormat    =   "yyyy/M/d"
            Format          =   155189251
            CurrentDate     =   38718
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   285
            Index           =   3
            Left            =   1890
            TabIndex        =   354
            Top             =   180
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   503
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
            ButtonImage     =   "FrmMembers.frx":70845
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   285
            Index           =   3
            Left            =   270
            TabIndex        =   355
            Top             =   180
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   503
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
            ButtonImage     =   "FrmMembers.frx":70DDF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   360
            Left            =   3615
            TabIndex        =   356
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   150
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   635
            Caption         =   "«÷«›…  ”ÿ—"
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
            ButtonImage     =   "FrmMembers.frx":71379
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin VSFlex8UCtl.VSFlexGrid grdHistEye 
            Height          =   1695
            Left            =   60
            TabIndex        =   358
            Top             =   2610
            Width           =   13620
            _cx             =   24024
            _cy             =   2990
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
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmMembers.frx":77BDB
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
            WallPaperAlignment=   0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ "
            Height          =   285
            Index           =   82
            Left            =   6900
            RightToLeft     =   -1  'True
            TabIndex        =   353
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„ﬁ«”"
            Height          =   315
            Index           =   81
            Left            =   12360
            RightToLeft     =   -1  'True
            TabIndex        =   351
            Top             =   210
            Width           =   1215
         End
      End
   End
   Begin Dynamic_Byte.NourHijriCal Txt_DateExpLincH 
      Height          =   345
      Left            =   0
      TabIndex        =   89
      Top             =   960
      Width           =   1815
      _extentx        =   3201
      _extenty        =   609
   End
   Begin MSDataListLib.DataCombo DcCustomerType 
      Height          =   315
      Left            =   5520
      TabIndex        =   115
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   1440
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCEmP 
      Height          =   315
      Left            =   90
      TabIndex        =   117
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   1440
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
      Left            =   3480
      TabIndex        =   205
      Top             =   8310
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… ﬂ—  ⁄„Ì·"
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
      TabIndex        =   206
      Top             =   600
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      CalendarBackColor=   12648447
      CustomFormat    =   "yyyy/M/d"
      Format          =   155189251
      CurrentDate     =   38718
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   11
      Left            =   4800
      TabIndex        =   222
      Top             =   8310
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
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   5760
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox XPMTxtRemarks2 
         Alignment       =   1  'Right Justify
         Height          =   795
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
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
         TabIndex        =   55
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
         TabIndex        =   54
         Top             =   240
         Width           =   825
      End
   End
   Begin MSDataListLib.DataCombo DcbCurrency 
      Height          =   315
      Left            =   90
      TabIndex        =   253
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label XPLbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «· ”ÃÌ· VAT"
      Height          =   465
      Index           =   6
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   255
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·⁄„·…"
      Height          =   255
      Index           =   14
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   254
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "”Ì«”… «·»Ì⁄ "
      Height          =   405
      Index           =   10
      Left            =   4305
      RightToLeft     =   -1  'True
      TabIndex        =   220
      Top             =   1830
      Width           =   1170
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   285
      Index           =   64
      Left            =   12870
      RightToLeft     =   -1  'True
      TabIndex        =   209
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
      TabIndex        =   207
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„‰œÊ»"
      Height          =   285
      Index           =   1
      Left            =   4305
      RightToLeft     =   -1  'True
      TabIndex        =   118
      Top             =   1440
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·⁄„Ì·"
      Height          =   285
      Index           =   2
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   116
      Top             =   1440
      Width           =   1890
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·«‰ Â«¡"
      Height          =   345
      Index           =   12
      Left            =   1830
      TabIndex        =   90
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”Ã· «· Ã«—Ì"
      Height          =   345
      Index           =   5
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8505
      TabIndex        =   60
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
      TabIndex        =   51
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
      TabIndex        =   50
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
      TabIndex        =   49
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
      TabIndex        =   35
      Top             =   8340
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
      TabIndex        =   23
      Top             =   1005
      Width           =   915
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„Ì·"
      Height          =   315
      Index           =   1
      Left            =   12930
      RightToLeft     =   -1  'True
      TabIndex        =   24
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
      TabIndex        =   22
      Top             =   8340
      Width           =   1215
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      Height          =   285
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   8340
      Width           =   465
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      Height          =   285
      Left            =   210
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   8340
      Width           =   615
   End
End
Attribute VB_Name = "FrmCustemers"
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
Dim mAllowEditCreditLimit As Boolean, mAllowEditCreditBalance As Boolean
Private m_DcboCustomers As DataCombo
Dim mCustId As Long
Public FormNamee As String

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

 

Private Sub PassData()
    Dim StrSQL As String
    On Error GoTo ErrTrap
    StrSQL = "SELECT * From TblCustemers"
    If Me.calledFromForm = False Then Exit Sub
    Select Case Me.DealingForm

 
        Case InvoiceTransaction
            fill_combo Me.DcboCustomers, StrSQL
            Me.DcboCustomers.BoundText = val(XPTxtCusID.Text)
 
        Case PriceList
     StrSQL = "SELECT * From TblCustemers where Type=2"
         'fill_combo FrmMainPriceList.DBCboSupplierName, StrSQL
         '  FrmMainPriceList.DBCboSupplierName.BoundText = val(XPTxtCusID.Text)

            '⁄—÷ «·√”⁄«—
        Case ShowPrice
'wael
          'fill_combo FrmShowPrice.DBCboClientName, StrSQL
    '       FrmShowPrice.DBCboClientName.BoundText = val(XPTxtCusID.Text)
    End Select
 Me.calledFromForm = False
Unload Me
    Exit Sub
ErrTrap:
End Sub

Private Sub ALLButton1_Click()
    Frame2.Visible = True
End Sub

Private Sub ALLButton2_Click()
   Dim balanceString As String
        WriteCustomerBalPublic IIf(IsNull(rs("Account_Code")), "", rs("Account_Code")), , balanceString
        lbl(8).Caption = balanceString
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
   ' Cmd_Click (1)
    OptType(2).value = True
    TxtOpenBalance.Text = 0
   ' Cmd_Click (2)

End Function

Private Sub CDMOldContract_Click()
Unload FrmOldContract
FrmOldContract.ScrenFlg = 0
FrmOldContract.show


End Sub

Public Sub Cmd_Click(Index As Integer)

'    On Error GoTo ErrTrap
Dim IntRes As Integer
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
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            fg.Rows = 1
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            Txt_DateExpLincH.value = ToHijriDate(Date)
            If SystemOptions.CanCustomerandVendor Then
                chkCustomerandVendor.Enabled = False
            End If
            TxtCreditLimit.Text = val(SystemOptions.LimitDefaultCredit)
            TxtDepitInterval.Text = val(SystemOptions.LimitDefaultCreditDays)
            dcDepitIntervalID.ListIndex = 0
            If Not SystemOptions.AllowEditCreditLimit Then
                TxtCreditLimit.Enabled = False
                TxtDepitInterval.Enabled = False
                dcDepitIntervalID.Enabled = False
            End If

            
            Dim Account_Code_dynamic As String
            Account_Code_dynamic = get_account_code_branch(8, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                MsgBox "No Branch was created", vbCritical
                End If
                GoTo ErrTrap
            Else
                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» ··⁄„·«¡ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "No Account was defined for clients in this branch for the operation", vbCritical
                    End If
                End If
            End If
        
            DboParentAccount.BoundText = Account_Code_dynamic
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

            Me.Dtp.value = FirstPeriodDateInthisYear
            Me.dcBranch.BoundText = Current_branch
            OptType(2).value = True
            OptType1(2).value = True
            OptType2(2).value = True

                                               
   Dim EmpID As Integer
 
   If SystemOptions.usertype <> UserAdminAll Then
  
  GetUserData user_id, , , , , , EmpID
        Me.DCEmP.BoundText = EmpID
 
  End If
  
  lbl(8).Caption = 0
  DcbCurrency.BoundText = MainCurrency()
 CboPayMentType.ListIndex = 1
 
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
VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
            TxtModFlg.Text = "E"

        Case 2

    If Me.DboParentAccount.BoundText = "" Then
    If SystemOptions.UserInterface = EnglishInterface Then
             Msg = "Specify Parent Account"
       Else
           Msg = " Õœœ «·Õ”«» «·—∆Ì”Ì   «Ê·« "
     End If
 If val(XPTxtCusID.Text) <> 2 Then
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        DboParentAccount.SetFocus
         SendKeys "{F4}"
       Screen.MousePointer = vbDefault
       
      Exit Sub
  End If
End If

            Dim currentcode As String

            If txtid.Text = "" Then
                currentcode = get_coding(Current_branch, "TblCustemers", 4, Me.DCPreFix.Text)

                If currentcode = "miniError" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
                    Else
                        MsgBox "The Number of digits for the code is too small please change the coding policy or connect your administrator"
                    End If
                    Exit Sub
            
                ElseIf currentcode = "Manual" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·ÕﬁÊ·"
                    Else
                        MsgBox "Please enter the code manually"
                    End If
                    Exit Sub
                Else
                    txtid = currentcode
                End If
            End If
If SystemOptions.CustMobNoMandatory Then
    If Trim(XPTxtmobile) = "" Or Len(Trim(XPTxtmobile)) < 10 Or mId(Trim(XPTxtmobile), 1, 1) <> "0" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ ÃÊ«· ’ÕÌÕ Ì»œ√ » ’›—  Ê„ﬂÊ‰ „‰ 10 Œ«‰«  "
        Else
            MsgBox "Please Enter Mob No."
        End If
        Exit Sub
    Else
    
    
    

    If SystemOptions.CheckMobileFormatCorrect = True Then
   ' txtCodeSend = "+966"
   Dim isFound As Boolean
                If Not FindString(XPTxtmobile, "+966", 1) Then
                                    If Not FindString(XPTxtmobile, "00966", 1) Then
                                        isFound = False
                                    Else
                                        isFound = True
                                    End If
                                If Not isFound Then
                                    isFound = False
                                    XPTxtmobile = "00966" & mId(XPTxtmobile, 2, Len(XPTxtmobile))
                                End If
                End If
    
                If Len(XPTxtmobile) < 14 Then
                    MsgBox "«œŒ· —ﬁ„ ÃÊ«· ’ÕÌÕ"
                    'txtCodeSend = "+966"
                     Exit Sub
                Else
                  '  txtCodeSend = ""
                    'txtPhoneCust = "+966" & val(txtPhoneCust)
                End If
    End If
    
    
    End If
End If


If SystemOptions.CustomerRecordNoIsnotManda = False Then
If txtCustGID.Text = "" And SystemOptions.CustMobNoMandatory = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ «·”Ã·"
Else
MsgBox "Please Enter Record No."
End If
txtCustGID.SetFocus
Exit Sub
End If
End If
If TxtVATNO.Text = "" And SystemOptions.CustomerRecordNoIsnotManda = False Then
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "·„ Ì „ «œŒ«· —ﬁ„ «·›« .Â·  —Ìœ «·Õ›Ÿ"
Else
Msg = "Confirm Save Without VAT No."
End If
IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)
If IntRes = vbNo Then
TxtVATNO.SetFocus
Exit Sub
End If
End If
If XPTxtmobile.Text = "" And SystemOptions.CustomerRecordNoIsnotManda = False Then
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "·„ Ì „ «œŒ«· —ﬁ„ «·ÃÊ«·.Â·  —Ìœ «·Õ›Ÿ"
Else
Msg = "Confirm Save Without Mobile No."
End If
IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)
If IntRes = vbNo Then
XPTxtmobile.SetFocus
Exit Sub
End If
End If

 
If DcboCountryID.Text = "" And SystemOptions.CustomerRecordNoIsnotManda = False Then
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "·„ Ì „   «Œ Ì«— «·œÊ·….Â·  —Ìœ «·Õ›Ÿ"
Else
Msg = "Confirm Save Without Country "
End If
IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)
If IntRes = vbNo Then
'DcboCountryID.SetFocus
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

            If XPTxtCusID.Text = 2 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·”Ã·"
                Else
                    Msg = "This record can't be deleted"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            Del_Member

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
         FrmCustemerSearch.SearchType = 0
       FrmCustemerSearch.RetrunType = 0
            FrmCustemerSearch.show vbModal
      
        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            
print_report2
            '        txtBranchName.text = 2
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
ShowAttachments DCPreFix.Text & txtid.Text, "0701201401"
 
Case 12
If Me.TxtModFlg.Text <> "R" Then
RemoveGridRow
End If
 Case 13
            
              If VSFlexGrid1.Row < VSFlexGrid1.FixedRows Then Exit Sub
                Dim StrSQL  As String
                
                If ISCarAllowDelete(val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("ID")))) Then
                          VSFlexGrid1.RemoveItem (VSFlexGrid1.Row)
                Else
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox ("€Ì— „”„ÊÕ Õ–› «·„⁄œÂ/«·”Ì«—… " & VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("BoardNo")) & " · ﬂ«„· «·»Ì«‰«  ")
                    Else
                        MsgBox ("vehicle can't be deleted  " & VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("BoardNo")) & "for data integration ")
                    End If
                End If
                
            Case 15
                   DelAll
Case 14
 addrow
 
 
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub DelAll()
 If VSFlexGrid1.Rows <= VSFlexGrid1.FixedRows Then Exit Sub
 Dim i  As Integer, m As Integer, StrSQL As String
 m = VSFlexGrid1.Rows - VSFlexGrid1.FixedRows
 i = VSFlexGrid1.Rows - 1
 Do While VSFlexGrid1.Rows > VSFlexGrid1.FixedRows
      If ISCarAllowDelete(val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("ID")))) Then
                  StrSQL = " delete from TblVendorCars where ID =    " & val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("ID")))
                  Cn.Execute StrSQL, , adExecuteNoRecords
                  VSFlexGrid1.RemoveItem (i)
                  i = i - 1
        Else
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox ("€Ì— „”„ÊÕ Õ–› «·„⁄œÂ/«·”Ì«—… " & VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("BoardNo")) & " · ﬂ«„· «·»Ì«‰«  ")
                Else
                    MsgBox ("Vehicle can't be deleted  " & VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("BoardNo")) & "for data integration ")
                End If
                
                  Exit Sub
        End If
 Loop
End Sub
Private Sub addrow()
If dcBrand.BoundText = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«Œ — «·„«—ﬂ… «Ê·«")
    Else
        MsgBox ("Select Brand")
    End If
    dcBrand.SetFocus
    SendKeys ("{F4}")
    Exit Sub
End If

Dim board As String
Dim lettter As String
Dim Num As String
Dim nboard As String
Dim nlettter As String
Dim nNum As String


lettter = txtLetter1.Text & " " & txtLetter2.Text & " " & txtLetter3.Text & " " & txtLetter4.Text
Num = txtNum1.Text & " " & txtNum2.Text & " " & txtNum3.Text & " " & txtNum4.Text

nlettter = ntxtLetter1.Text & " " & ntxtLetter2.Text & " " & ntxtLetter3.Text & " " & ntxtLetter4.Text
nNum = ntxtNum1.Text & " " & ntxtNum2.Text & " " & ntxtNum3.Text & " " & ntxtNum4.Text

board = lettter & " " & Num

nboard = nlettter & " " & nNum

If Len(lettter) + Len(Num) < 2 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("·«»œ „‰ « »«⁄  ›Ê—„«  «·«œŒ«· ")
    Else
        MsgBox ("Should follow the input formate ")
    End If
    Exit Sub
End If
    If Trim(nboard) <> Trim(txtnBoardNo) Then
        nboard = txtnBoardNo
    End If
    
    If Trim(board) <> Trim(txtBoardNo) Then
        board = txtBoardNo
    End If
    If Trim(nboard) = "" Then Exit Sub
    Dim s As Integer
    
    For s = 1 To VSFlexGrid1.Rows - 1
            If (Replace(VSFlexGrid1.TextMatrix(s, VSFlexGrid1.ColIndex("BoardNo")), " ", "") = Replace(nboard, " ", "")) Then
                If SystemOptions.AllowRepeateCar And Trim(VSFlexGrid1.TextMatrix(s, VSFlexGrid1.ColIndex("ChasisNo"))) = Trim(txtChassis.Text) Then
                      If SystemOptions.UserInterface = ArabicInterface Then
                             MsgBox (" „ «÷«›… «·„⁄œÂ/«·”Ì«—… Â–Â „‰ ﬁ»·  ")
                         Else
                             MsgBox ("This vehicle has been add befor")
                         End If
                    
                     If VSFlexGrid1.Rows > 1 Then
                         'VSFlexGrid1.RemoveItem s
                     End If
                     Exit Sub
                Else
                    If Not SystemOptions.AllowRepeateCar Then
                         If SystemOptions.UserInterface = ArabicInterface Then
                             MsgBox (" „ «÷«›… «·„⁄œÂ/«·”Ì«—… Â–Â „‰ ﬁ»·  ")
                         Else
                             MsgBox ("This vehicle has been add befor")
                         End If
                    
                     If VSFlexGrid1.Rows > 1 Then
                         'VSFlexGrid1.RemoveItem s
                     End If
                     Exit Sub
                    End If
                 End If
                
            End If
    Next

Dim i As Integer


Dim j As Integer, ad As Boolean

ad = False

For j = 1 To VSFlexGrid1.Rows - 1
        If VSFlexGrid1.TextMatrix(j, VSFlexGrid1.ColIndex("BrandID")) = "" Then
                i = j
                ad = True
                Exit For
        End If
Next

If ad = False Then
VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
i = VSFlexGrid1.Rows
i = i - 1
End If


With VSFlexGrid1
    .TextMatrix(i, .ColIndex("Serial")) = i - 1
    .TextMatrix(i, .ColIndex("BoardNo")) = nboard
    .TextMatrix(i, .ColIndex("nBoardNo")) = board
    .TextMatrix(i, .ColIndex("ChasisNo")) = txtChassis.Text
    .TextMatrix(i, .ColIndex("BrandID")) = dcBrand.BoundText
    .TextMatrix(i, .ColIndex("Brand")) = dcBrand.Text
    .TextMatrix(i, .ColIndex("ModelID")) = cbModel.ListIndex
    .TextMatrix(i, .ColIndex("Model")) = cbModel.Text
    .TextMatrix(i, .ColIndex("Count")) = txtCount.Text
    .TextMatrix(i, .ColIndex("Rate")) = txtRate.Caption
    .TextMatrix(i, .ColIndex("CityID")) = DcCity.ListIndex
    .TextMatrix(i, .ColIndex("City")) = DcCity.Text
    .TextMatrix(i, .ColIndex("DriverName")) = txtDriverName.Text
    .TextMatrix(i, .ColIndex("DriverTel")) = txtDriverTel.Text
    .TextMatrix(i, .ColIndex("accessory")) = accessoryTxt.Text
    .TextMatrix(i, .ColIndex("Price")) = val(TxtPrice.Text)
    .TextMatrix(i, .ColIndex("PartPrice")) = val(TxtPartPrice.Text)
    .TextMatrix(i, .ColIndex("TypeTransID")) = val(DcbTypTrans.ListIndex) + 1
    .TextMatrix(i, .ColIndex("TypeTrans")) = DcbTypTrans.Text
    .TextMatrix(i, .ColIndex("ModelName")) = DcbCarModel.Text
    .TextMatrix(i, .ColIndex("CarModelID")) = val(DcbCarModel.BoundText)
    .TextMatrix(i, .ColIndex("ColorName")) = DcbColor.Text
    .TextMatrix(i, .ColIndex("ColorID")) = val(DcbColor.BoundText)
End With
'
VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
txtChassis.Text = ""
dcBrand.BoundText = ""
cbModel.ListIndex = -1
txtCount.Text = ""
txtDriverName.Text = ""
DcCity.ListIndex = -1
 
txtLetter1.Text = ""
txtLetter2.Text = ""
txtLetter3.Text = ""
txtLetter4.Text = ""

txtNum1.Text = ""
txtNum2.Text = ""
txtNum3.Text = ""
txtNum4.Text = ""


ntxtLetter1.Text = ""
ntxtLetter2.Text = ""
ntxtLetter3.Text = ""
ntxtLetter4.Text = ""

ntxtNum1.Text = ""
ntxtNum2.Text = ""
ntxtNum3.Text = ""
ntxtNum4.Text = ""
txtnBoardNo = ""
txtBoardNo = ""
DcbCarModel.BoundText = 0
DcbColor.BoundText = 0
End Sub
Private Sub RemoveGridRow()
Dim Msg As String
    With Me.Grid
        If .Row <= 0 Then Exit Sub
                If CheckDelLocations(val(.TextMatrix(.Row, .ColIndex("ID")))) = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
                  Else
                  Msg = "Can't Delete...!!!"
                  End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
        Cn.Execute "Delete TblCustomersLocations  where id =" & val(.TextMatrix(.Row, .ColIndex("ID"))) & "  "
        .RemoveItem .Row
    End With
   ' ReLineGrid
End Sub

Private Sub Cmd_DeleteAll_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then

    
        fg.Rows = 1
        fg.Rows = 2
   

End If
End Sub

Private Sub Cmd_DeleteRow_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
      With Me.fg
'MsgBox .Row
        If .Row <= 0 Then
                .Rows = 2
        Exit Sub
        Else
        .RemoveItem .Row
        End If
    End With



End If
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub DboParentAccount_KeyUp(KeyCode As Integer, _
                                   Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 77
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no record to show"
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

Private Sub ISButton3_Click()

  Dim k As Long, LngNewRow As Long
  If Trim(fg.TextMatrix(fg.Rows - 1, fg.ColIndex("SizesNamesName"))) = "" Then
        fg.Rows = fg.Rows - 1
    End If
    If fg.Rows = 1 Then fg.Rows = 2 Else fg.Rows = fg.Rows + 1
    
    
    k = fg.Rows
   
    If fg.Rows <= 1 Then
        fg.Rows = fg.Rows + 1
    End If
    LngNewRow = fg.Rows - 1
     'mNewId = LngNewRow
     

    fg.TextMatrix(LngNewRow, fg.ColIndex("SizesNamesName")) = cmbSizesNames.Text
    fg.TextMatrix(LngNewRow, fg.ColIndex("SizesNamesID")) = cmbSizesNames.BoundText
    fg.TextMatrix(LngNewRow, fg.ColIndex("DateSize")) = txtDateSize.value
End Sub

Private Sub ntxtLetter1_Change()
FilltxtBord
End Sub
Private Sub FilltxtBord()
Dim board As String
Dim lettter As String
Dim Num As String
Dim nboard As String
Dim nlettter As String
Dim nNum As String


lettter = txtLetter1.Text & " " & txtLetter2.Text & " " & txtLetter3.Text & " " & txtLetter4.Text
Num = txtNum1.Text & " " & txtNum2.Text & " " & txtNum3.Text & " " & txtNum4.Text

nlettter = ntxtLetter1.Text & " " & ntxtLetter2.Text & " " & ntxtLetter3.Text & " " & ntxtLetter4.Text
nNum = ntxtNum1.Text & " " & ntxtNum2.Text & " " & ntxtNum3.Text & " " & ntxtNum4.Text

board = lettter & " " & Num

nboard = nlettter & " " & nNum

txtBoardNo = board
txtnBoardNo = nboard
End Sub

Private Sub ntxtLetter2_Change()
FilltxtBord
End Sub

Private Sub ntxtLetter3_Change()
FilltxtBord
End Sub

Private Sub ntxtLetter4_Change()
FilltxtBord
End Sub

Private Sub ntxtNum1_Change()
FilltxtBord
End Sub

Private Sub ntxtNum2_Change()
FilltxtBord
End Sub

Private Sub ntxtNum3_Change()
FilltxtBord
End Sub

Private Sub ntxtNum4_Change()
FilltxtBord
End Sub

Private Sub txtChassis_Change()
'txtChassis.Text = UCase(txtChassis.Text)
End Sub

Private Sub txtChassis_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtLetter1_Change()
txtLetter1.Text = UCase(txtLetter1.Text)
FilltxtBord
End Sub

Private Sub txtLetter2_Change()
txtLetter2.Text = UCase(txtLetter2.Text)
FilltxtBord
End Sub

Private Sub txtLetter3_Change()
txtLetter3.Text = UCase(txtLetter3.Text)
FilltxtBord
End Sub


Private Sub txtLetter4_Change()
txtLetter4.Text = UCase(txtLetter4.Text)
FilltxtBord
End Sub


 


Private Sub txtLetter1_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
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
End Sub
Private Sub ntxtLetter1_KeyPress(KeyAscii As Integer)
ntxtLetter1.Text = ""
If Len(ntxtLetter1.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        ntxtLetter2.SetFocus
End Select
End Sub
Private Sub ntxtLetter2_KeyPress(KeyAscii As Integer)
ntxtLetter2.Text = ""
If Len(ntxtLetter2.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        ntxtLetter3.SetFocus
End Select
End Sub

Private Sub ntxtLetter3_KeyPress(KeyAscii As Integer)
ntxtLetter3.Text = ""
If Len(ntxtLetter3.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        ntxtLetter4.SetFocus
End Select
End Sub

Private Sub ntxtLetter4_KeyPress(KeyAscii As Integer)
ntxtLetter4.Text = ""
If Len(ntxtLetter4.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        ntxtNum1.SetFocus
End Select
End Sub

Private Sub txtNum1_Change()
FilltxtBord
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
End Sub

Private Sub txtNum3_Change()
FilltxtBord
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
End Sub

Private Sub txtNum4_Change()
FilltxtBord
End Sub

Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.Text = ""
If Len(txtNum4.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
End Sub

Private Sub ntxtNum1_KeyPress(KeyAscii As Integer)
ntxtNum1.Text = ""
If Len(ntxtNum1.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        ntxtNum2.SetFocus
End If
End Sub

Private Sub ntxtNum2_KeyPress(KeyAscii As Integer)
ntxtNum2.Text = ""
If Len(ntxtNum2.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        ntxtNum3.SetFocus
End If
End Sub

Private Sub ntxtNum3_KeyPress(KeyAscii As Integer)
ntxtNum3.Text = ""
If Len(ntxtNum3.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        ntxtNum4.SetFocus

End If
End Sub

Private Sub ntxtNum4_KeyPress(KeyAscii As Integer)
ntxtNum4.Text = ""
If Len(ntxtNum4.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
End Sub
Private Sub Form_Load()
    
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If mdifrmmain.CarMaintenance.Visible = True Then
        Me.Height = 10560
    Else
        Me.Height = 9270
    End If

    If mdifrmmain.hajMnu.Visible = True Then
        Fra(10).Visible = True
    Else
        Fra(10).Visible = False
    End If
        
    If 1 = 0 Then
        'Me.Height = 10560
        Frame3.Visible = True
    Else
        Frame3.Visible = False
        'Me.Height = 9270
    End If
    txtDateSize.value = Date
    Dim StrSQL As String
    If SystemOptions.CanCustomerandVendor Then
        chkCustomerandVendor.Enabled = False
    End If
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    s = "SELECT isNull(AllowEditCreditLimit,0) AllowEditCreditLimit ,isNull(AllowEditCreditBalance,0) AllowEditCreditBalance  From TblUsers WHERE TblUsers.UserID= " & user_id & ""
    Set rsDummy = Nothing
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rsDummy.EOF Then
     mAllowEditCreditLimit = CBool(rsDummy!AllowEditCreditLimit)
     mAllowEditCreditBalance = CBool(rsDummy!AllowEditCreditBalance)
    End If
 
'    grdHistEye.MergeCells = flexMergeFixedOnly
'    grdHistEye.MergeRow(0) = True
'    grdHistEye.MergeRow(1) = True
    'On Error GoTo ErrTrap

    'Resize_Form Me
      If SystemOptions.UserInterface = ArabicInterface Then
        VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("TypeTransID")) = "#1; ﬂ«„·|#2; »œÊ‰ „·Õﬁ"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("TypeTransID")) = "#1;Complete |#2;Without Part "
    End If
 With DcbTypTrans
.Clear
If SystemOptions.UserInterface = ArabicInterface Then
.AddItem "ﬂ«„·"
.AddItem "»œÊ‰ „·Õﬁ"
Else
.AddItem "Complete"
.AddItem "Without Part "
End If
End With
    Dim str As String
If SystemOptions.UserInterface = ArabicInterface Then
str = " Select id , name  from TBLCarTypes "
Else
str = " Select id , namee  from TBLCarTypes "
End If
fill_combo dcBrand, str
Dim k As Integer
With Me.cbModel
.Clear
'1995
For k = 1900 To 2100
        .AddItem k
Next
End With


        With CboPayMentType
             .Clear
             .AddItem "‰ﬁœ«"
             .AddItem "¬Ã·"
         End With

     If SystemOptions.AllowScInterface = True Then

        FrmCustemers.Caption = "√Ê·Ì«¡ «·«„Ê—"
        EleHeader.Caption = FrmCustemers.Caption
        Label1(6).Visible = True
        Label1(7).Visible = True
        Label1(8).Visible = True
        Label1(9).Visible = True
        Label1(10).Visible = True
        DcbFM.Visible = True
        DcbClass.Visible = True
        DcbLevel.Visible = True
        'DcbFirstClass.Visible = True
        DcbCurrClass.Visible = True
        Label1(11).Visible = True
        Label1(12).Visible = True
        TxtPassport.Visible = True
        TxtIQAMA.Visible = True
        lbl(75).Visible = True
        lbl(76).Visible = True
  
        DBRegisterDate.Visible = True
        DBENDDATE.Visible = True
        Label1(13).Visible = True
        Label1(14).Visible = True
        DCNAtionalID.Visible = True
        STDUDENTStatusID.Visible = True
  
        Me.C1Tab1.TabCaption(2) = "«·ÿ·«»"
        With Grid
           '.ColHidden(.ColIndex("FirstClass")) = False
            .ColHidden(.ColIndex("CurrClass")) = False
            .ColHidden(.ColIndex("Class")) = False
            .ColHidden(.ColIndex("MF")) = False
            .ColHidden(.ColIndex("Level")) = False
            .ColHidden(.ColIndex("TxtIQAMA")) = False
            .ColHidden(.ColIndex("TxtPassport")) = False
            .ColHidden(.ColIndex("DBRegisterDate")) = False
            .ColHidden(.ColIndex("DBENDDATE")) = False
            '.ColHidden(.ColIndex("DCNAtionalID")) = False
            .ColHidden(.ColIndex("DCNAtionaNAme")) = False
            .ColHidden(.ColIndex("STDUDENTStatusID")) = False
            .ColHidden(.ColIndex("CurrClass")) = False
        End With

 
    End If
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
        
    
    AddTip
    Dim Msg As String
    StrSQL = " select id,code from currency"
    fill_combo Me.DcbCurrency, StrSQL
    
    
    StrSQL = "SELECT * From TblSizesNames "
    fill_combo cmbSizesNames, StrSQL

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
    
    With Grid
        .ColComboList(.ColIndex("FirstClass")) = "#1;KG1 |#2;KG2 |#3;KG3 |#4;Grade1 |#5;Grade2 |#6;Grade3 |#7;Grade4 |#8;Grade5 |#9;Grade6 |#10;Grade7 |#11;Grade8 |#12;Grade9 |#13;Grade10 |#14;Grade11 |#15;Grade12"
        .ColComboList(.ColIndex("CurrClass")) = "#1;KG1 |#2;KG2 |#3;KG3 |#4;Grade1 |#5;Grade2 |#6;Grade3 |#7;Grade4 |#8;Grade5 |#9;Grade6 |#10;Grade7 |#11;Grade8 |#12;Grade9 |#13;Grade10 |#14;Grade11 |#15;Grade12"
        .ColComboList(.ColIndex("Level")) = "#1;—Ê÷… |#2;«» œ«∆Ì |#3;„ Ê”ÿ |#4;À«‰ÊÌ"
        .ColComboList(.ColIndex("MF")) = "#1;M |#2;F "
        .ColComboList(.ColIndex("STDUDENTStatusID")) = "#1;„” „— |#2;ÃœÌœ "
    End With

    Dim My_SQL As String
    Dim Dcombos As New ClsDataCombos
    
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from Nationality  "
    Else
        My_SQL = "  select  id,namee  from Nationality  "
    End If

    fill_combo DCNAtionalID, My_SQL


    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  GroupID,GroupName  from GroupsCustomers  "
    Else
        My_SQL = "  select  GroupID,GroupNamee  from GroupsCustomers  "
    End If

    fill_combo dcGroups, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  ID,Name  from ClassCustomers  "
    Else
        My_SQL = "  select  ID,Namee  from ClassCustomers  "
    End If

    fill_combo dcClass, My_SQL


       
       
       
    DBRegisterDate.value = Date
    DBENDDATE.value = Date
    
    If SystemOptions.UserInterface = ArabicInterface Then
        With STDUDENTStatusID
            .Clear
            .AddItem "„” „—(ﬁœÌ„)"
            .AddItem "ÃœÌœ"
        End With

        With DcbLevel
            .Clear
            .AddItem "—Ê÷…"
            .AddItem "«» œ«∆Ì"
            .AddItem "„ Ê”ÿ"
            .AddItem "À«‰ÊÌ"
        End With
    
        With DcbFM
            .Clear
            .AddItem "M"
            .AddItem "F"
        End With

        With DcbFirstClass
            .Clear
            .AddItem "KG1"
            .AddItem "KG2"
            .AddItem "KG3"
            .AddItem "Grade1"
            .AddItem "Grade2"
            .AddItem "Grade3"
            .AddItem "Grade4"
            .AddItem "Grade5"
            .AddItem "Grade6"
            .AddItem "Grade7"
            .AddItem "Grade8"
            .AddItem "Grade9"
            .AddItem "Grade10"
            .AddItem "Grade11"
            .AddItem "Grade12"
        End With

        With DcbCurrClass
            .Clear
            .AddItem "KG1"
            .AddItem "KG2"
            .AddItem "KG3"
            .AddItem "Grade1"
            .AddItem "Grade2"
            .AddItem "Grade3"
            .AddItem "Grade4"
            .AddItem "Grade5"
            .AddItem "Grade6"
            .AddItem "Grade7"
            .AddItem "Grade8"
            .AddItem "Grade9"
            .AddItem "Grade10"
            .AddItem "Grade11"
            .AddItem "Grade12"
        End With
    Else
            With STDUDENTStatusID
            .Clear
            .AddItem "Continuous (old)"
            .AddItem "New"
        End With

        With DcbLevel
            .Clear
            .AddItem "Kindergarten"
            .AddItem "Elementary"
            .AddItem "Intermediate"
            .AddItem "High School"
        End With
    
        With DcbFM
            .Clear
            .AddItem "M"
            .AddItem "F"
        End With

        With DcbFirstClass
            .Clear
            .AddItem "KG1"
            .AddItem "KG2"
            .AddItem "KG3"
            .AddItem "Grade1"
            .AddItem "Grade2"
            .AddItem "Grade3"
            .AddItem "Grade4"
            .AddItem "Grade5"
            .AddItem "Grade6"
            .AddItem "Grade7"
            .AddItem "Grade8"
            .AddItem "Grade9"
            .AddItem "Grade10"
            .AddItem "Grade11"
            .AddItem "Grade12"
        End With

        With DcbCurrClass
            .Clear
            .AddItem "KG1"
            .AddItem "KG2"
            .AddItem "KG3"
            .AddItem "Grade1"
            .AddItem "Grade2"
            .AddItem "Grade3"
            .AddItem "Grade4"
            .AddItem "Grade5"
            .AddItem "Grade6"
            .AddItem "Grade7"
            .AddItem "Grade8"
            .AddItem "Grade9"
            .AddItem "Grade10"
            .AddItem "Grade11"
            .AddItem "Grade12"
        End With
    End If
    
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
             
             'Dim str As String, k As Integer
           Dim i As Integer
              i = 0
              '"#1;KG1 |#2;KG2
              str = ""
              VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("Model")) = ""
             For k = 1900 To 2050
                    
                    'str = str & "#" & i & "|" & k
                    VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("Model")) = VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("Model")) & "#" & i & ";" & k & IIf(k = 2050, "", "|")
                    str = str & "#" & i & ";" & k & "|"
                    i = i + 1
             Next
           '  VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("Model")) = str
            '.ComboList = str
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

      Dcombos.GetAccountingCodes Me.DboParentAccount, False, True, , , 1
    Dcombos.GetCodeing Me.DCPreFix, 4
    'Dcombos.GetEmployees Me.DCEmp
    Dcombos.GetSalesRepData Me.DCEmP
    Dcombos.GetTblCarModels Me.DcbCarModel
    Dcombos.GetTblColor Me.DcbColor
    Dcombos.GetClass Me.DcbClass
    Me.Dtp.value = Date
    DtRecord.value = Date
    StrSQL = "select * From TblCustemers where type=1"
    StrSQL = StrSQL & "  AND   (BranchId=0 or BranchId is null or     BranchId in(" & Current_branchSql & "))"
     
    If SystemOptions.usertype <> UserAdminAll Then
        'StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
    End If
    
    If SystemOptions.usertype <> UserAdminAll Then
        If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  Empid = " & user_id
        End If
        Me.dcBranch.Enabled = True
       'DCEmP.Enabled = False
    End If

    Set rs = New ADODB.Recordset
    StrSQL = StrSQL & "Order By Fullcode"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " »Ì«‰«  «·⁄„·«¡  "
    LogTexte = " Open Window " & "  Customers Data "
   'AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""

   
    
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ErrTrap
    LogTextA = "  «·Œ—ÊÃ „‰  " & " »Ì«‰«  «·⁄„·«¡  "
    LogTexte = " Exit   Window " & "  Customers Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""
mCustId = val(XPTxtCusID)
Dim Dcombos As New ClsDataCombos
If FormNamee = "FrmDefinCompItem" Then
    
    Dcombos.GetCustomersSuppliers 1, FrmDefinCompItem.DBCboClientName, True
    
    FrmDefinCompItem.DBCboClientName.BoundText = mCustId
    FrmDefinCompItem.CboPayMentType.ListIndex = CboPayMentType.ListIndex
    
End If

If FormNamee = "FrmItemsClass" Then
   FrmItemsClass.TxtClientCode = txtid
   FrmItemsClass.TxtCliientName = XPTxtCusName
   FrmItemsClass.TxtCusID = XPTxtCusID
   FrmItemsClass.retInfoCustomer
End If

If FormNamee = "frmsalebill" Then
    
    Dcombos.GetCustomersSuppliers 1, frmsalebill.DBCboClientName, True
    frmsalebill.DBCboClientName.BoundText = mCustId
  '   frmsalebill.CboPayMentType.ListIndex = CboPayMentType.ListIndex
    
End If

If FormNamee = "frmsalebill4" Then
    
    Dcombos.GetCustomersSuppliers 1, frmsalebill4.DBCboClientName, True
    frmsalebill4.DBCboClientName.BoundText = mCustId
  '   frmsalebill.CboPayMentType.ListIndex = CboPayMentType.ListIndex
    
End If

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
    LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«   «·⁄„·«¡ " _
       & CHR(13) & " ﬂÊœ «·⁄„Ì·  " & DCPreFix & txtid.Text _
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
       & CHR(13) & "  ‰Ê⁄ «·⁄„Ì·  " & DcCustomerType.Text _
       & CHR(13) & " «·„‰œÊ»   " & DCEmP.Text _
       & CHR(13) & " Õœ «·«∆ „«‰ „œÌ‰  " & TxtCreditLimit _
       & CHR(13) & " „œ… «·«∆ „«‰     " & TxtDepitInterval.Text & "   " & dcDepitIntervalID.Text _
       & CHR(13) & " Õœ «·«∆ „«‰ œ«∆‰   " & TxtCreditlimitCredit _
       & CHR(13) & " „œ… «·«∆ „«‰      " & TxtCreditInterval.Text & "   " & dcCreditIntervalID.Text _
                    
       LogTextA = LogTextA & CHR(13) & "⁄„Ì· „Ê—œ ø       "

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

    LogTexte = "  Õ›Ÿ ‘«‘… " & " Customers Data  " _
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
                    
       LogTexte = LogTexte & CHR(13) & "Customer & Supplier ?  "

    If chkCustomerandVendor.value = vbChecked Then
        LogTexte = LogTexte & " Yes "
    Else
        LogTexte = LogTexte & " No "
    End If

    LogTexte = LogTexte & CHR(13) & "Locked"

    If locked.value = vbChecked Then
        LogTexte = LogTexte & "Yes "
        LogTexte = LogTexte & CHR(13) & "  Reasons  "
        LogTexte = LogTexte & CHR(13) & XPMTxtRemarks2
    Else
        LogTexte = LogTexte & "No "
    End If

    LogTexte = LogTexte & CHR(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "

    If OptType(0).value = True Then
        LogTexte = LogTexte & "„œÌ‰"
    ElseIf OptType(1).value = True Then
        LogTexte = LogTexte & "œ«∆‰"
    ElseIf OptType(2).value = True Then
        LogTexte = LogTexte & "€Ì— „Õœœ"
    End If

    LogTexte = LogTexte & CHR(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ  " & TxtOpenBalance
    LogTexte = LogTexte & CHR(13) & "  Parent Acc. " & DboParentAccount
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", ""
    End If

End Function



Sub filgrid()
Dim i As Integer
Dim k As Integer
If val(txtid1.Text) = 0 Then
With Grid
k = .Rows
.Rows = .Rows + 1
For i = k To .Rows - 1
.TextMatrix(i, .ColIndex("MangerName")) = Me.TxtMangerName
 
.TextMatrix(i, .ColIndex("ClassID")) = val(DcbClass.BoundText)
.TextMatrix(i, .ColIndex("Class")) = DcbClass.Text
.TextMatrix(i, .ColIndex("NameE")) = Me.TxtNameE.Text
.TextMatrix(i, .ColIndex("Name")) = Me.TxtName.Text
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("CurrClass")) = val(Me.DcbCurrClass.ListIndex) + 1
.TextMatrix(i, .ColIndex("FirstClass")) = val(Me.DcbFirstClass.ListIndex) + 1
.TextMatrix(i, .ColIndex("MF")) = val(DcbFM.ListIndex) + 1
.TextMatrix(i, .ColIndex("Level")) = val(Me.DcbLevel.ListIndex) + 1

.TextMatrix(i, .ColIndex("STDUDENTStatusID")) = val(Me.STDUDENTStatusID.ListIndex) + 1

.TextMatrix(i, .ColIndex("DCNAtionalID")) = val(DCNAtionalID.BoundText)
.TextMatrix(i, .ColIndex("DCNAtionaNAme")) = (DCNAtionalID.Text)

.TextMatrix(i, .ColIndex("DBRegisterDate")) = Me.DBRegisterDate.value
If Not IsNull(DBENDDATE.value) Then
.TextMatrix(i, .ColIndex("DBENDDATE")) = Me.DBENDDATE.value
Else
.TextMatrix(i, .ColIndex("DBENDDATE")) = ""
End If
.TextMatrix(i, .ColIndex("TxtIQAMA")) = Me.TxtIQAMA.Text
.TextMatrix(i, .ColIndex("TxtPassport")) = Me.TxtPassport.Text

Next i
End With
Else
With Grid
.TextMatrix(val(txtid1.Text), .ColIndex("MangerName")) = Me.TxtMangerName.Text
.TextMatrix(val(txtid1.Text), .ColIndex("NameE")) = Me.TxtNameE.Text
.TextMatrix(val(txtid1.Text), .ColIndex("Name")) = Me.TxtName.Text
.TextMatrix(val(txtid1.Text), .ColIndex("CurrClass")) = val(Me.DcbCurrClass.ListIndex) + 1
.TextMatrix(val(txtid1.Text), .ColIndex("FirstClass")) = val(Me.DcbFirstClass.ListIndex) + 1
.TextMatrix(val(txtid1.Text), .ColIndex("MF")) = val(DcbFM.ListIndex) + 1
.TextMatrix(val(txtid1.Text), .ColIndex("Level")) = val(Me.DcbLevel.ListIndex) + 1
.TextMatrix(val(txtid1.Text), .ColIndex("ClassID")) = val(DcbClass.BoundText)
.TextMatrix(val(txtid1.Text), .ColIndex("Class")) = DcbClass.Text


.TextMatrix(val(txtid1.Text), .ColIndex("STDUDENTStatusID")) = val(Me.STDUDENTStatusID.ListIndex) + 1

.TextMatrix(val(txtid1.Text), .ColIndex("DCNAtionalID")) = val(DCNAtionalID.BoundText)
.TextMatrix(val(txtid1.Text), .ColIndex("DCNAtionaNAme")) = (DCNAtionalID.Text)
.TextMatrix(val(txtid1.Text), .ColIndex("DBRegisterDate")) = Me.DBRegisterDate.value
If Not IsNull(DBENDDATE.value) Then
.TextMatrix(val(txtid1.Text), .ColIndex("DBENDDATE")) = Me.DBENDDATE.value
Else
.TextMatrix(val(txtid1.Text), .ColIndex("DBENDDATE")) = ""
End If
.TextMatrix(val(txtid1.Text), .ColIndex("TxtIQAMA")) = Me.TxtIQAMA.Text
.TextMatrix(val(txtid1.Text), .ColIndex("TxtPassport")) = Me.TxtPassport.Text


End With
End If
End Sub


Sub FullGrid()
Dim i As Integer
Dim sql As String
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
'sql = "SELECT     dbo.TblCustomersLocations.ID, dbo.TblCustomersLocations.Name, dbo.TblCustomersLocations.NameE, dbo.TblCustomersLocations.MangerName, "
'sql = sql & "                      dbo.TblCustomersLocations.FirstClass, dbo.TblCustomersLocations.CurrClass, dbo.TblClass.Name AS ClassName, dbo.TblClass.NameE AS ClassNameW,"
'sql = sql & "                      dbo.TblCustomersLocations.CusID , dbo.TblCustomersLocations.MF, dbo.TblCustomersLocations.[Level], dbo.TblCustomersLocations.ClassId"
'sql = sql & "  FROM         dbo.TblCustomersLocations LEFT OUTER JOIN"
'sql = sql & "                       dbo.TblClass ON dbo.TblCustomersLocations.ClassID = dbo.TblClass.ID"
sql = "SELECT     dbo.TblCustomersLocations.ID, dbo.TblCustomersLocations.Name, dbo.TblCustomersLocations.NameE, dbo.TblCustomersLocations.MangerName, "
sql = sql & "                                          dbo.TblCustomersLocations.FirstClass, dbo.TblCustomersLocations.CurrClass, dbo.TblClass.Name AS ClassName, dbo.TblClass.NameE AS ClassNameW,"
sql = sql & "                                          dbo.TblCustomersLocations.CusId, dbo.TblCustomersLocations.MF, dbo.TblCustomersLocations.[Level], dbo.TblCustomersLocations.ClassID,"
sql = sql & "                                          dbo.TblCustomersLocations.DBRegisterDate, dbo.TblCustomersLocations.DBENDDATE, dbo.TblCustomersLocations.TxtIQAMA,"
sql = sql & "                                          dbo.TblCustomersLocations.TxtPassport, dbo.TblCustomersLocations.STDUDENTStatusID, dbo.TblCustomersLocations.DCNAtionalID,"
sql = sql & "                                          dbo.Nationality.name AS NationalityA, dbo.Nationality.namee AS NationalityE"
sql = sql & "                    FROM         dbo.TblCustomersLocations LEFT OUTER JOIN"
sql = sql & "                                          dbo.Nationality ON dbo.TblCustomersLocations.DCNAtionalID = dbo.Nationality.id LEFT OUTER JOIN"
sql = sql & "                                          dbo.TblClass ON dbo.TblCustomersLocations.ClassID = dbo.TblClass.ID"
sql = sql & "   Where (dbo.TblCustomersLocations.CusID = " & val(Me.XPTxtCusID.Text) & ")"

 
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With Grid
.Rows = Rs3.RecordCount + 1
Rs3.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
.TextMatrix(i, .ColIndex("FirstClass")) = IIf(IsNull(Rs3("FirstClass").value), "", Rs3("FirstClass").value)
.TextMatrix(i, .ColIndex("CurrClass")) = IIf(IsNull(Rs3("CurrClass").value), "", Rs3("CurrClass").value)
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
.TextMatrix(i, .ColIndex("NameE")) = IIf(IsNull(Rs3("NameE").value), "", Rs3("NameE").value)
.TextMatrix(i, .ColIndex("MangerName")) = IIf(IsNull(Rs3("MangerName").value), "", Rs3("MangerName").value)
.TextMatrix(i, .ColIndex("MF")) = IIf(IsNull(Rs3("MF").value), "", Rs3("MF").value)
.TextMatrix(i, .ColIndex("Level")) = IIf(IsNull(Rs3("Level").value), "", Rs3("Level").value)
.TextMatrix(i, .ColIndex("ClassID")) = IIf(IsNull(Rs3("ClassID").value), 0, Rs3("ClassID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Class")) = IIf(IsNull(Rs3("ClassName").value), "", Rs3("ClassName").value)
Else
.TextMatrix(i, .ColIndex("Class")) = IIf(IsNull(Rs3("ClassNameW").value), "", Rs3("ClassNameW").value)
End If


.TextMatrix(i, .ColIndex("DCNAtionalID")) = IIf(IsNull(Rs3("DCNAtionalID").value), 0, Rs3("DCNAtionalID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("DCNAtionaNAme")) = IIf(IsNull(Rs3("NationalityA").value), "", Rs3("NationalityA").value)
Else
.TextMatrix(i, .ColIndex("DCNAtionaNAme")) = IIf(IsNull(Rs3("NationalityE").value), "", Rs3("NationalityE").value)
End If

.TextMatrix(i, .ColIndex("TxtIQAMA")) = IIf(IsNull(Rs3("TxtIQAMA").value), "", Rs3("TxtIQAMA").value)
.TextMatrix(i, .ColIndex("TxtPassport")) = IIf(IsNull(Rs3("TxtPassport").value), "", Rs3("TxtPassport").value)

.TextMatrix(i, .ColIndex("DBRegisterDate")) = IIf(IsNull(Rs3("DBRegisterDate").value), "", Rs3("DBRegisterDate").value)
.TextMatrix(i, .ColIndex("DBENDDATE")) = IIf(IsNull(Rs3("DBENDDATE").value), "", Rs3("DBENDDATE").value)
.TextMatrix(i, .ColIndex("STDUDENTStatusID")) = IIf(IsNull(Rs3("STDUDENTStatusID").value), 0, Rs3("STDUDENTStatusID").value)


Rs3.MoveNext
Next i

End With
End If
       Dim j As Integer
       
       Dim d As Integer
        Dim str As String
      sql = " SELECT     dbo.TblCusCar.DriverName, dbo.TblCusCar.DriverTel, dbo.TblCusCar.EndAllocationDate, dbo.TblCusCar.ID, dbo.TblCusCar.Serial, dbo.TblCusCar.BoardNo,"
      sql = sql & "           dbo.TblCusCar.nBoardNo, dbo.TblCusCar.ChasisNo, dbo.TblCusCar.BrandID, dbo.TblCusCar.ModelID, dbo.TblCusCar.CityID, dbo.TblCusCar.Rate,"
      sql = sql & "                dbo.TblCusCar.CustomerID, dbo.TBLCarTypes.name, dbo.TblCusCar.Accessory, dbo.TblCusCar.Price, dbo.TblCusCar.PartPrice, dbo.TblCusCar.TypeTransID,"
      sql = sql & "                dbo.TblCusCar.[Count], dbo.TBLCarTypes.namee, dbo.TblCusCar.ColorID, dbo.TblColor.name AS ColorName, dbo.TblColor.namee AS ColorNameE,"
      sql = sql & "                dbo.TblCusCar.CarModelID , dbo.TblCarModels.Model, dbo.TblCarModels.ModelE"
      sql = sql & "      FROM         dbo.TblCusCar LEFT OUTER JOIN"
      sql = sql & "                dbo.TblCarModels ON dbo.TblCusCar.CarModelID = dbo.TblCarModels.Id LEFT OUTER JOIN"
      sql = sql & "                dbo.TblColor ON dbo.TblCusCar.ColorID = dbo.TblColor.Id LEFT OUTER JOIN"
      sql = sql & "                dbo.TBLCarTypes ON dbo.TblCusCar.BrandID = dbo.TBLCarTypes.id"
      sql = sql & "  where customerID = " & val(XPTxtCusID.Text)
      Dim RsCusCar As ADODB.Recordset
       Set RsCusCar = New ADODB.Recordset
       RsCusCar.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  '     RsCusCar.MoveFirst
       
       
       VSFlexGrid1.Rows = 1
       With VSFlexGrid1
       .Rows = RsCusCar.RecordCount + 1
       For j = 1 To RsCusCar.RecordCount
              .TextMatrix(j, .ColIndex("PartPrice")) = IIf(IsNull(RsCusCar("PartPrice").value), "", RsCusCar("PartPrice").value)
              .TextMatrix(j, .ColIndex("Price")) = IIf(IsNull(RsCusCar("Price").value), "", RsCusCar("Price").value)
              .TextMatrix(j, .ColIndex("serial")) = IIf(IsNull(RsCusCar("serial").value), "", RsCusCar("serial").value)
              .TextMatrix(j, .ColIndex("ID")) = IIf(IsNull(RsCusCar("ID").value), "", RsCusCar("ID").value)
              .TextMatrix(j, .ColIndex("BoardNo")) = IIf(IsNull(RsCusCar("BoardNo").value), "", RsCusCar("BoardNo").value)
              .TextMatrix(j, .ColIndex("nBoardNo")) = IIf(IsNull(RsCusCar("nBoardNo").value), "", RsCusCar("nBoardNo").value)
              .TextMatrix(j, .ColIndex("ChasisNo")) = IIf(IsNull(RsCusCar("ChasisNo").value), "", RsCusCar("ChasisNo").value)
              .TextMatrix(j, .ColIndex("BrandID")) = IIf(IsNull(RsCusCar("BrandID").value), "", RsCusCar("BrandID").value)
              .TextMatrix(j, .ColIndex("ModelID")) = IIf(IsNull(RsCusCar("ModelID").value), "", RsCusCar("ModelID").value)
              .TextMatrix(j, .ColIndex("Count")) = IIf(IsNull(RsCusCar("Count").value), "", RsCusCar("Count").value)
              .TextMatrix(j, .ColIndex("CityID")) = IIf(IsNull(RsCusCar("CityID").value), "", RsCusCar("CityID").value)
              .TextMatrix(j, .ColIndex("Rate")) = IIf(IsNull(RsCusCar("Rate").value), "", RsCusCar("Rate").value)
              .TextMatrix(j, .ColIndex("Model")) = IIf(IsNull(RsCusCar("ModelID").value), "", val(RsCusCar("ModelID").value) + 1900)
              .TextMatrix(j, .ColIndex("Brand")) = IIf(IsNull(RsCusCar("name").value), "", RsCusCar("name").value)
              .TextMatrix(j, .ColIndex("DriverName")) = IIf(IsNull(RsCusCar("drivername").value), "", RsCusCar("drivername").value)
              .TextMatrix(j, .ColIndex("DriverTel")) = IIf(IsNull(RsCusCar("DriverTel").value), "", RsCusCar("DriverTel").value)
              .TextMatrix(j, .ColIndex("EndDate")) = IIf(IsNull(RsCusCar("EndAllocationDate").value), "", RsCusCar("EndAllocationDate").value)
              .TextMatrix(j, .ColIndex("accessory")) = IIf(IsNull(RsCusCar("accessory").value), "", RsCusCar("accessory").value)
              .TextMatrix(j, .ColIndex("TypeTransID")) = IIf(IsNull(RsCusCar("TypeTransID").value), "", RsCusCar("TypeTransID").value)
              .TextMatrix(j, .ColIndex("ColorID")) = IIf(IsNull(RsCusCar("ColorID").value), 0, RsCusCar("ColorID").value)
              .TextMatrix(j, .ColIndex("CarModelID")) = IIf(IsNull(RsCusCar("CarModelID").value), 0, RsCusCar("CarModelID").value)
            If SystemOptions.UserInterface = ArabicInterface Then
              .TextMatrix(j, .ColIndex("ModelName")) = IIf(IsNull(RsCusCar("Model").value), "", RsCusCar("Model").value)
              .TextMatrix(j, .ColIndex("ColorName")) = IIf(IsNull(RsCusCar("ColorName").value), "", RsCusCar("ColorName").value)
            Else
              .TextMatrix(j, .ColIndex("ModelName")) = IIf(IsNull(RsCusCar("ModelE").value), "", RsCusCar("ModelE").value)
              .TextMatrix(j, .ColIndex("ColorName")) = IIf(IsNull(RsCusCar("ColorNameE").value), "", RsCusCar("ColorNameE").value)
            End If
          .TextMatrix(j, .ColIndex("City")) = str
          RsCusCar.MoveNext
       Next
       End With
       
       
          sql = "SELECT td.RSPH,td.*,t.Transaction_Date FROM Transaction_Details AS td"
    sql = sql & " INNER JOIN Transactions AS t ON t.Transaction_ID = td.Transaction_ID"
    sql = sql & " Where t.CusID = " & val(txtid)
    sql = sql & " AND ISNULL(td.RSPH,'') <> '' OR ISNULL(td.LSPH,'') <> ''"
    
    
    
sql = "SELECT ROrL = CASE ISNULL(td.RSPH, '')"
sql = sql & "                    WHEN '' THEN 'Left'"
sql = sql & "                    Else            'RIGHT'"
sql = sql & "               END,"
sql = sql & "        IsAdd     = CASE ISNULL("
sql = sql & "                         CASE ISNULL(td.RADD, '')"
sql = sql & "                              WHEN '' THEN LADD"
sql = sql & "                              Else RADD"
sql = sql & "                         END,"
sql = sql & "                         ''"
sql = sql & "                     )"
sql = sql & "                     WHEN '' THEN 0"
sql = sql & "                     ELSE 1"
sql = sql & "                END,"
sql = sql & "        SPH       = CASE ISNULL(td.RSPH, '')"
sql = sql & "                   WHEN '' THEN LSPH"
sql = sql & "                   Else RSPH"
sql = sql & " END,"
sql = sql & "        AXIS      = CASE ISNULL(td.RAXIS, '')"
sql = sql & "                    WHEN '' THEN LAXIS"
sql = sql & "                    Else RAxis"
sql = sql & "               END,"
sql = sql & "        CYL       = CASE ISNULL(td.RCYL, '')"
sql = sql & "                   WHEN '' THEN LCYL"
sql = sql & "                   Else RCYL"
sql = sql & " END,"
sql = sql & "        [ADD]     = CASE ISNULL(td.RADD, '')"
sql = sql & "                     WHEN '' THEN LADD"
sql = sql & "                     Else RADD"
sql = sql & "                END,"
sql = sql & "        CL        = CASE ISNULL(td.RCL, '')"
sql = sql & "                  WHEN '' THEN LCL"
sql = sql & "                  Else td.RCL"
sql = sql & "             END,"
sql = sql & "        td.Remarks,"
sql = sql & "        td.price,"
sql = sql & "        t.Transaction_Date"
sql = sql & " FROM   Transaction_Details      AS td"
sql = sql & "        INNER JOIN Transactions  AS t"
sql = sql & "             ON  t.Transaction_ID = td.Transaction_ID"
sql = sql & "          Where t.CusID = " & val(txtid)
sql = sql & "             AND ISNULL(td.RSPH, '') <> '' OR ISNULL(td.LSPH, '') <> ''"

    loadgrid sql, grdHistEye, True, False
    
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
   Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    With Grid
        Select Case .ColKey(Col)
           Case "Class"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ClassID"), False, True)
                .TextMatrix(Row, .ColIndex("ClassID")) = StrAccountCode
        End Select
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Me.TxtModFlg.Text = "R" Then
Cancel = True
End If
End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FillTextFromGrid
ErrTrap:
End Sub
Sub FillTextFromGrid()
Dim i As Integer
With Me.Grid
If .Row > 0 Then
txtid1.Text = .Row
TxtName.Text = .TextMatrix(.Row, .ColIndex("Name"))
TxtNameE.Text = .TextMatrix(.Row, .ColIndex("NameE"))
Me.TxtMangerName.Text = .TextMatrix(.Row, .ColIndex("MangerName"))
Me.DcbLevel.ListIndex = val(.TextMatrix(.Row, .ColIndex("Level"))) - 1
Me.DcbFM.ListIndex = val(.TextMatrix(.Row, .ColIndex("MF"))) - 1
Me.DcbClass.BoundText = val(.TextMatrix(.Row, .ColIndex("ClassID")))
Me.DcbCurrClass.ListIndex = val(.TextMatrix(.Row, .ColIndex("CurrClass"))) - 1

Me.TxtPassport.Text = .TextMatrix(.Row, .ColIndex("TxtPassport"))
Me.TxtIQAMA.Text = .TextMatrix(.Row, .ColIndex("TxtIQAMA"))

Me.DBRegisterDate.value = .TextMatrix(.Row, .ColIndex("DBRegisterDate"))


Me.DBENDDATE.value = IIf(.TextMatrix(.Row, .ColIndex("DBENDDATE")) = "", Null, .TextMatrix(.Row, .ColIndex("DBENDDATE")))

Me.DCNAtionalID.BoundText = val(.TextMatrix(.Row, .ColIndex("DCNAtionalID")))
Me.STDUDENTStatusID.ListIndex = val(.TextMatrix(.Row, .ColIndex("STDUDENTStatusID"))) - 1



'Me.DcbDeptManger.BoundText = val(.TextMatrix(.Row, .ColIndex("MangerID")))
End If
End With

End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String
    Dim Msg As String
    With Grid

        Select Case .ColKey(Col)
          Case "Class"
                 StrSQL = " SELECT     ID, Name, NameE"
                 StrSQL = StrSQL & "             FROM         dbo.TblClass"
                 Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Name", "ID")
                Else
                StrComboList = .BuildComboList(rs, "NameE", "ID")
                End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ComboList = StrComboList
        End Select

    End With
End Sub

Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
If SystemOptions.UserInterface = ArabicInterface Then
If Me.TxtName.Text = "" Then
MsgBox "«œŒ· «·«”„"
Me.TxtName.SetFocus
Exit Sub
End If
Else
If Me.TxtNameE.Text = "" Then
MsgBox "Please Eneter Name"
Me.TxtNameE.SetFocus
Exit Sub
End If
End If
filgrid
Me.TxtNameE.Text = ""
Me.TxtName.Text = ""
Me.TxtMangerName.Text = ""
txtid1.Text = 0
DcbCurrClass.ListIndex = -1
Me.DcbFirstClass.ListIndex = -1
End If


End Sub

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
                    Msg = "Â–« «·⁄„Ì· „”Ã· „‰ ﬁ»·   "
                    Msg = Msg & CHR(13) & " ﬂÊœ «·⁄„Ì·: " & Custcode
                    Msg = Msg & CHR(13) & " «”„ «·⁄„Ì· : " & CustName
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
                Me.Caption = "»Ì«‰«  «·⁄„·«¡"
            Else
                Me.Caption = "Customers Data"
            End If

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
            txtCustGID.locked = False
            DboParentAccount.Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·⁄„·«¡( ÃœÌœ )"
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
    
            DboParentAccount.Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·⁄„·«¡(  ⁄œÌ· )"
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

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
   With VSFlexGrid1
     
         Select Case .ColKey(Col)
         
                    Case "Brand"
                        
                        .TextMatrix(.Row, .ColIndex("BrandID")) = .ComboData
                        VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
                        
                     Case "Model"
                     Dim k As Integer
                     k = val(.TextMatrix(.Row, .ColIndex("Model")))
                     k = k - 1900
                    .TextMatrix(.Row, .ColIndex("ModelID")) = k
                    
          End Select

    End With
End Sub
 Public Function ISCarAllowDelete(CarID As Integer) As Boolean
Dim str As String, allow As Boolean
Dim RsTemp As ADODB.Recordset
allow = True
str = " Select * from TblCarBillMentains where CarID =   " & CarID
Set RsTemp = New ADODB.Recordset
RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsTemp.RecordCount > 0 Then
        ISCarAllowDelete = False
        Exit Function
End If

str = " Select * from TblCardAuthorizationReform  where  CarID  =   " & CarID
Set RsTemp = New ADODB.Recordset
RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsTemp.RecordCount > 0 Then
        ISCarAllowDelete = False
        Exit Function
End If
 
 ISCarAllowDelete = True
 
 End Function
Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Me.TxtModFlg.Text = "R" Then
Cancel = True
Exit Sub
End If
With VSFlexGrid1
     Dim i As Integer
     i = IIf(.TextMatrix(Row, .ColIndex("id")) = "", 0, val(.TextMatrix(Row, .ColIndex("id"))))
     If i > 0 Then
            If ISCarAllowDelete(i) = False Then
                    Cancel = True
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox ("€Ì— „”„ÊÕ «· ⁄œÌ· ⁄·Ï  «·„⁄œÂ/«·”Ì«—… " & VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("BoardNo")) & " · ﬂ«„· «·»Ì«‰«  ")
                    Else
                        MsgBox ("editing isn't allowed for this vehicle" & Grid.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("BoardNo")) & "for data integration")
                    End If
                    Exit Sub
            End If
     End If
     Select Case .ColKey(Col)

    Case "BoardNo"
            Cancel = True

End Select

End With
End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    
     With VSFlexGrid1
     Select Case .ColKey(Col)
    
      Case "BoardNo"
         .ComboList = ""
      Case "ChasisNo"
            .ComboList = ""
      Case "Brand"
                    
          StrSQL = "  Select id , name  from TBLCarTypes ORDER BY ID "
          RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          StrComboList = VSFlexGrid1.BuildComboList(RsTemp, "Name", "ID")
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList
          
          
      Case "Model"
        '    StrSQL = "  Select id , name  from TBLCarTypes ORDER BY ID "
        '    rstemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
           ' StrComboList = Grid.BuildComboList(rstemp, "Name", "ID")
           ' If StrComboList <> "" Then
           '        StrComboList = "|" & StrComboList
           ' End If
            Dim str As String, k As Integer
           
             
             For k = 1900 To 2050
                    str = str & "|" & k
             Next
            .ComboList = str
        Case "Count"
                .ComboList = ""
                
        Case "Rate"
        .ComboList = ""
        
        Case "DriverName"
        .ComboList = ""
        
        Case "DriverTel"
        .ComboList = ""
        
        Case Is = "EndDate"
        .ComboList = ""

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

Public Sub Retrive(Optional Lngid As Long = 0, Optional ByVal mFormName As String = "")
    Dim SngCusBegainAccount As Single

    On Error GoTo ErrTrap
'FormNamee = mFormName
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    txtDateSize.value = Date
    If Not (rs.EOF Or rs.BOF) Then
        If Lngid <> 0 Then
            rs.Find "CusID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    
        If rs("CustomerandVendor").value = True Then
            chkCustomerandVendor.value = vbChecked

        Else
            chkCustomerandVendor.value = vbUnchecked
        End If
        
        fg.Rows = 1
        TxtVATNO.Text = IIf(IsNull(rs("VATNO").value), "", rs("VATNO").value)
        dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
      Me.DtRecord.value = IIf(IsNull(rs("RecordDate")), Date, rs("RecordDate"))
        DCPreFix.Text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
        Me.txtid.Text = IIf(IsNull(rs("code").value), "", rs("code").value)
        txtopening_balance_voucher_id.Text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)
        XPTxtCusID.Text = IIf(IsNull(rs("CusID")), "", val(rs("CusID")))
        XPTxtCusName.Text = IIf(IsNull(rs("CusName")), "", Trim(rs("CusName")))
        XPTxtCusNamee.Text = IIf(IsNull(rs("CusNamee")), "", Trim(rs("CusNamee")))
        c1.Text = IIf(IsNull(rs("c1")), "", Trim(rs("c1")))
        c2.Text = IIf(IsNull(rs("c2")), "", Trim(rs("c2")))
        CboPayMentType.ListIndex = IIf(IsNull(rs("CPaymentType").value), 0, rs("CPaymentType").value)
        dcGroups.BoundText = IIf(IsNull(rs("ClassCustomersId").value), "", rs("ClassCustomersId").value)
        dcClass.BoundText = IIf(IsNull(rs("GroupsCustomersId").value), "", rs("GroupsCustomersId").value)
        txtBranchName.Text = IIf(IsNull(rs("BranchName")), "", Trim(rs("BranchName")))
        
        ''/////////////salah
    Me.TxtMap.Text = IIf(IsNull(rs("Map").value), "", rs("Map").value)
    Me.txtJob.Text = IIf(IsNull(rs("JobName").value), "", rs("JobName").value)
    Me.TxtEntry.Text = IIf(IsNull(rs("Entry").value), "", rs("Entry").value)
    '''///////////////////
     TxtBankCode.Text = IIf(IsNull(rs("BankCode").value), "", rs("BankCode").value)
     TxtBankIBAN.Text = IIf(IsNull(rs("BankIBAN").value), "", rs("BankIBAN").value)
     TxtBankAddress.Text = IIf(IsNull(rs("BankAddress").value), "", rs("BankAddress").value)
     txtBankAccount.Text = IIf(IsNull(rs("BankAccount").value), "", rs("BankAccount").value)
     txtBankName.Text = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
     If Not IsNull(rs("TypeOmrh").value) Then
     If rs("TypeOmrh").value = 1 Then
     TypeOmrh.value = vbChecked
     Else
     TypeOmrh.value = vbUnchecked
     End If
     Else
     TypeOmrh.value = vbUnchecked
     End If
     If Not IsNull(rs("TypeHaj").value) Then
     If rs("TypeHaj").value = 1 Then
     TypeHaj.value = vbChecked
     Else
     TypeHaj.value = vbUnchecked
     End If
     Else
     TypeHaj.value = vbUnchecked
     End If

     If Not IsNull(rs("IsNotCommission").value) Then
        If rs("IsNotCommission").value = 1 Then
           chkIsNotCommission.value = vbChecked
        Else
           chkIsNotCommission.value = vbUnchecked
        End If
     Else
        chkIsNotCommission.value = vbUnchecked
     End If


     DcbCurrency.BoundText = IIf(IsNull(rs("CurrncyID").value), "", rs("CurrncyID").value)
     If Not IsNull(rs("HajEnter_Out").value) Then
     If rs("HajEnter_Out").value = 1 Then
     HajEnter_Out(1).value = True
     ElseIf rs("HajEnter_Out").value = 0 Then
     HajEnter_Out(0).value = True
     End If
     Else
     HajEnter_Out(0).value = True
     End If
    '''////
        Me.TxtResponsibleContact.Text = IIf(IsNull(rs("ResponsibleContact").value), "", rs("ResponsibleContact").value)
        XPTxtPhone.Text = IIf(IsNull(rs("Cus_Phone")), "", Trim(rs("Cus_Phone")))
        txtCustGID.Text = IIf(IsNull(rs("CustGID")), "", (rs("CustGID")))
    ''///
    Me.TxtBox.Text = IIf(IsNull(rs("Boxmil")), "", Trim(rs("Boxmil")))
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
         ' SngCusBegainAccount = GetCustomerAccount(val(XPTxtCusID.Text), True)
      '  Dim balanceString As String
'WriteCustomerBalPublic IIf(IsNull(rs("Account_Code")), "", rs("Account_Code")), , balanceString
       lbl(8).Caption = ""
    
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
        Me.TXTJobTitle.Text = IIf(IsNull(rs("JobTitle")), "", Trim(rs("JobTitle")))
        Me.TxtSalary.Text = IIf(IsNull(rs("Salary")), 0, Trim(rs("Salary")))
        Me.TxtJobAddress.Text = IIf(IsNull(rs("JobAddress")), "", Trim(rs("JobAddress")))
        Me.TxtJobTel.Text = IIf(IsNull(rs("JobTel")), "", Trim(rs("JobTel")))
        Me.TxtJobTelConvert.Text = IIf(IsNull(rs("JobTelConvert")), "", Trim(rs("JobTelConvert")))
        Me.TxtHomeTel.Text = IIf(IsNull(rs("HomeTel")), "", Trim(rs("HomeTel")))
        Me.TxtMobile1.Text = IIf(IsNull(rs("Mobile1")), "", Trim(rs("Mobile1")))
        Me.TxtMobile2.Text = IIf(IsNull(rs("Mobile2")), "", Trim(rs("Mobile2")))
    
    End If
    
        Dim s As String
        Dim rsDummy As New ADODB.Recordset
        s = "SELECT isNull(AllowEditCreditLimit,0) AllowEditCreditLimit ,isNull(AllowEditCreditBalance,0) AllowEditCreditBalance  From TblUsers WHERE TblUsers.UserID= " & user_id & ""
        Set rsDummy = Nothing
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not rsDummy.EOF Then
            mAllowEditCreditLimit = CBool((rsDummy!AllowEditCreditLimit))
            mAllowEditCreditBalance = CBool((rsDummy!AllowEditCreditBalance))
        End If
            
        Fra(0).Enabled = mAllowEditCreditLimit
        TxtCreditLimit.locked = Not mAllowEditCreditLimit
        TxtCreditlimitCredit.locked = Not mAllowEditCreditLimit
        TxtDepitInterval.locked = Not mAllowEditCreditLimit
        TxtCreditInterval.locked = Not mAllowEditCreditLimit
        dcDepitIntervalID.locked = Not mAllowEditCreditLimit
        dcCreditIntervalID.locked = Not mAllowEditCreditLimit
        Fra(1).Enabled = mAllowEditCreditBalance
        Fra(8).Enabled = mAllowEditCreditBalance
        Fra(9).Enabled = mAllowEditCreditBalance
FullGrid

    
    s = "Select TblCustomerSizes.*,TblSizesNames.Name as SizesNamesName from TblCustomerSizes Inner join TblSizesNames On TblSizesNames.ID =TblCustomerSizes.SizesNamesID Where CusId = " & val(Me.XPTxtCusID.Text)
    loadgrid s, fg, True, True
    
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„Ì·   " & CHR(13)
            Msg = Msg + (XPTxtCusName.Text) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Data will be deleted" & CHR(13)
            Msg = Msg + (XPTxtCusName.Text) & CHR(13)
            Msg = Msg + "Do you want to continue"
        End If
        IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

        If IntRes = vbYes Then
            If Not rs.RecordCount < 1 Then
            ' If VSFlexGrid1.Rows <= VSFlexGrid1.FixedRows Then Exit Sub
 Dim i  As Integer, m As Integer
 m = VSFlexGrid1.Rows - VSFlexGrid1.FixedRows
 i = VSFlexGrid1.Rows - 1
 Do While VSFlexGrid1.Rows > VSFlexGrid1.FixedRows
      If ISCarAllowDelete(val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("ID")))) Then
                  StrSQL = " delete from TblVendorCars where ID =    " & val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("ID")))
                  Cn.Execute StrSQL, , adExecuteNoRecords
                  VSFlexGrid1.RemoveItem (i)
                  i = i - 1
        Else
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox ("€Ì— „”„ÊÕ Õ–› «·„⁄œÂ/«·”Ì«—… " & VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("BoardNo")) & " · ﬂ«„· «·»Ì«‰«  ")
                Else
                    MsgBox ("Vehicle can't be deleted  " & VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("BoardNo")) & "for data integration ")
                End If
                
                  Exit Sub
        End If
 Loop
   StrSQL = " delete from TblCusCar where CustomerID =    " & val(XPTxtCusID.Text)
                  Cn.Execute StrSQL, , adExecuteNoRecords
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
If SystemOptions.CustomerhavethreeAccounts = True Then
                'StrAccountCode1 = rs("Account_Code1").value
                'StrAccountCode2 = rs("Account_Code2").value
                StrAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                      StrAccountCode1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
                    StrAccountCode2 = IIf(IsNull(rs("Account_Code2").value), "", rs("Account_Code2").value)
                    
                    
 End If
 
                StrSQL = "Delete From Accounts Where Account_Code='" & StrAccountCode & "'"
                     
If SystemOptions.CustomerhavethreeAccounts = True Then
                    If Not IsNull(rs("Account_Code1").value) Then
                   StrSQL = StrSQL & " or   Account_Code='" & rs("Account_Code1").value & "'"
                   End If
        
        
             If Not IsNull(rs("Account_Code2").value) Then
            StrSQL = StrSQL & " or   Account_Code='" & rs("Account_Code2").value & "'"
          End If
        
   End If
                Cn.Execute StrSQL, , adExecuteNoRecords
                CuurentLogdata ("D")

                      If SystemOptions.CustomerhavethreeAccounts = True Then
                    StrAccountCode1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
                    StrAccountCode2 = IIf(IsNull(rs("Account_Code2").value), "", rs("Account_Code2").value)
                    
                    ParentAccount = IIf(IsNull(rs("ParentAccount").value), "", rs("ParentAccount").value)

                                    If ModAccounts.DeleteAccount(StrAccountCode, True) = True And ModAccounts.DeleteAccount(StrAccountCode1, True) = True And ModAccounts.DeleteAccount(StrAccountCode2, True) = True Then
                                    If StrAccountCode <> "" And StrAccountCode1 <> "" And StrAccountCode2 <> "" Then
                                    If ModAccounts.DeleteAccount(ParentAccount, True) = True Then
                                    End If
                                    End If
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
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " „  ⁄„·Ì… «·Õ–›."
                Else
                    Msg = "Record deleted successfully"
                End If
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "This operation is not available due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·⁄„Ì· "
    Else
        Msg = "sorry, this record cannot be deleted due to data integration"
    End If
    
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
            rs.Find "CusID='" & val(XPTxtCusID.Text) & "'", , adSearchForward, adBookmarkFirst

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

   ' On Error GoTo ErrTrap


    If Trim(dcBranch.BoundText) = "" Then
 '       If SystemOptions.UserInterface = EnglishInterface Then
 '           Msg = "Specify Departement"
 '       Else
 '           Msg = " Õœœ ›—⁄ «Ê·« "
 '       End If
'
'        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        dcBranch.SetFocus
'        SendKeys "{F4}"
'        Screen.MousePointer = vbDefault
'        Exit Sub
    End If

    If Me.TxtModFlg.Text <> "R" Then
        If XPTxtCusName.Text = "" Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "ÌÃ» «œŒ«· «”„ «·⁄„Ì·"
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

 If val(XPTxtCusID.Text) <> 2 Then
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
                                If TxtOpenBalance.Enabled = True Then
                                    TxtOpenBalance.SetFocus
                                End If
        
                        Exit Sub
                        
End If
                        
                    End If
        End If




        If Me.OptType1(2).value = False Then
                    If val(Me.TxtOpenBalance1.Text) = 0 Then
                        
                        
                                     If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·—’Ìœ ··‘Ìﬂ«   Õ  «· Õ’Ì· ··⁄„Ì·...!!!"
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
                    Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·—’Ìœ ··œ›⁄«  «·„ﬁœ„… ··⁄„Ì·...!!!"
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
                    Msg = Msg & CHR(13) & "·ﬁœ Ê÷⁄  Õœ ≈∆ „«‰ („œÌ‰ ) ··⁄„Ì· " & val(Me.TxtCreditLimit.Text)
                    Msg = Msg & CHR(13) & "·ﬂ‰ﬂ Ê÷⁄  «·—’Ìœ «·≈›  «ÕÏ ··⁄„Ì· „œÌ‰ »‹  " & val(Me.TxtOpenBalance.Text)
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
                    Msg = Msg & CHR(13) & "·ﬁœ Ê÷⁄  Õœ ≈∆ „«‰ (œ«∆‰ ) ··⁄„Ì· " & val(Me.TxtCreditlimitCredit.Text)
                    Msg = Msg & CHR(13) & "·ﬂ‰ﬂ Ê÷⁄  «·—’Ìœ «·≈›  «ÕÏ ··⁄„Ì· œ«∆‰ »‹  " & val(Me.TxtOpenBalance.Text)
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

                Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·Œ’„ «·Œ«’… »«·⁄„Ì·...!!!"
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

                Msg = "ÌÃ» ﬂ «»… ‰”»… «·Œ’„ «·Œ«’… »«·⁄„Ì·...!!!"
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

               Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·Œ’„ «·Œ«’… »«·⁄„Ì· ›Ï ›Ê« Ì— «·‘—«¡...!!!"
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

               Msg = "ÌÃ» ﬂ «»… ‰”»… «·Œ’„ «·Œ«’… »«·⁄„Ì· ›Ï ›Ê« Ì— «·‘—«¡..!!!"
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
            
                StrSQL = "Select * From TblCustemers where Type=1 And CusName='" & Trim(XPTxtCusName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                      If SystemOptions.UserInterface = ArabicInterface Then

                 Msg = "ÌÊÃœ ⁄„Ì· „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
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
            ''//////////////////
              If SystemOptions.CustomerRecordNoIsnotManda = False Then
            Set RsTemp = Nothing
                StrSQL = "Select * From TblCustemers where Type=1 And CustGID=" & val(txtCustGID.Text) & " and CusID <>" & val(XPTxtCusID.Text) & ""
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                      If SystemOptions.UserInterface = ArabicInterface Then

                 Msg = "ÌÊÃœ ⁄„Ì· „”Ã· „”»ﬁ«  »‰›” —ﬁ„ «·”Ã·" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                    Else
                     Msg = "this Customer Already Exist" & CHR(13)
                     
                    End If
                    If val(XPTxtCusID.Text) <> 2 Then
                            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            txtCustGID.SetFocus
                            Exit Sub
                    End If
                End If

            RsTemp.Close
            End If
            '////////////////////////
            
            Set RsTemp = Nothing
          StrSQL = "Select * From TblCustemers where   Type=1 And fullcode='" & Trim(DCPreFix.Text & txtid.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                      If SystemOptions.UserInterface = ArabicInterface Then

                 Msg = "ÌÊÃœ ⁄„Ì· „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ " & CHR(13)
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
                StrSQL = "select * From TblCustemers where Type=1 And CusName='" & Trim(XPTxtCusName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText


                If RsTemp.RecordCount > 0 Then
                                If RsTemp("CusID").value <> val(XPTxtCusID.Text) Then
                                     
                                                      If SystemOptions.UserInterface = ArabicInterface Then
                                
                                                 Msg = "ÌÊÃœ ⁄„Ì· „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
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
     
        ''//////////////////
            If SystemOptions.CustomerRecordNoIsnotManda = False Then
            Set RsTemp = Nothing
                StrSQL = "Select * From TblCustemers where Type=1 And CustGID=" & val(txtCustGID.Text) & " and CusID <>" & val(XPTxtCusID.Text) & ""
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                      If SystemOptions.UserInterface = ArabicInterface Then

                 Msg = "ÌÊÃœ ⁄„Ì· „”Ã· „”»ﬁ«  »‰›” —ﬁ„ «·”Ã·" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                    Else
                     Msg = "this Customer Already Exist" & CHR(13)
                     
                    End If

                   If val(XPTxtCusID.Text) <> 2 Then
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    txtCustGID.SetFocus
                    Exit Sub
                    
                   End If
                End If

            RsTemp.Close
           End If
            '////////////////////////
            Set RsTemp = Nothing
          StrSQL = "Select * From TblCustemers where Type=1 And fullcode='" & Trim(DCPreFix.Text & txtid.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If RsTemp.RecordCount > 0 Then
                                If RsTemp("CusID").value <> val(XPTxtCusID.Text) Then
                                     
                                                      If SystemOptions.UserInterface = ArabicInterface Then
                                
                                                 Msg = "ÌÊÃœ ⁄„Ì· „”Ã· „”»ﬁ« »Â–« «·ﬂÊœ " & CHR(13)
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
            Cn.Execute "Delete from TblCustomersLocations  where CusID =" & val(XPTxtCusID.Text)
            
            
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
             
     If HajEnter_Out(0).value = True Then
       rs("HajEnter_Out").value = 0
     ElseIf HajEnter_Out(1).value = True Then
       rs("HajEnter_Out").value = 1
     End If
     If TypeHaj.value = vbChecked Then
     rs("TypeHaj").value = 1
     Else
     rs("TypeHaj").value = Null
     End If
     
     If chkIsNotCommission.value = vbChecked Then
        rs("IsNotCommission").value = True
     Else
        rs("IsNotCommission").value = Null
     End If
    
     
     If TypeOmrh.value = vbChecked Then
     rs("TypeOmrh").value = 1
     Else
     rs("TypeOmrh").value = Null
     End If
        rs("VATNO").value = TxtVATNO.Text
        rs("CurrncyID").value = IIf(Me.DcbCurrency.BoundText = "", 0, val(DcbCurrency.BoundText))
        rs("BankCode").value = Trim(TxtBankCode.Text)
        rs("BankIBAN").value = Trim(TxtBankIBAN.Text)
        rs("BankAddress").value = Trim(TxtBankAddress.Text)
        rs("BankAccount").value = IIf(txtBankAccount.Text = "", "", Trim(txtBankAccount.Text))
        rs("BankName").value = IIf(txtBankName.Text = "", "", Trim(txtBankName.Text))
        rs("code").value = txtid.Text
        rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.Text) & IIf(Trim(txtid.Text) = "", Null, txtid.Text)
        rs("prifix").value = IIf(DCPreFix.Text = "", Null, DCPreFix.Text)
 Me.TxtFullcode = IIf(DCPreFix.BoundText = "", Null, DCPreFix.Text) & IIf(Trim(txtid.Text) = "", Null, txtid.Text)
     
        If chkCustomerandVendor.value = vbChecked Then
            rs("CustomerandVendor").value = 1

        Else
            rs("CustomerandVendor").value = 0
        End If
'
        If CboPayMentType.ListIndex = -1 Then
            rs("CPaymentType").value = 0
        Else
            rs("CPaymentType").value = val(CboPayMentType.ListIndex)
        End If
        rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
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
    
        rs("ClassCustomersId").value = IIf(Me.dcClass.BoundText = "", Null, (Me.dcClass.BoundText))
        rs("GroupsCustomersId").value = IIf(Me.dcGroups.BoundText = "", Null, (Me.dcGroups.BoundText))
        rs("BranchName").value = IIf(txtBranchName.Text = "", "", Trim(txtBranchName.Text))
        
        
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
        rs("Type").value = 1
        
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
                
          If SystemOptions.CustomerhavethreeAccounts = False Then
        
                                   rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtCusName.Text), True, False, Trim$(Me.XPTxtCusNamee.Text))

          Else
                
                                        If SystemOptions.CustomerhavethreeAccounts = True Then
                                            ParentAccount = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtCusName.Text, False, False, XPTxtCusNamee.Text)
                                            rs("ParentAccount").value = ParentAccount
                                         
                                            rs("Account_Code").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.Text), True, False, XPTxtCusNamee.Text)
                                            rs("Account_Code1").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.Text) & "   ‘Ìﬂ«    Õ  «· Õ’Ì· ", True, False, XPTxtCusNamee.Text & "  Under Collection Cheque  ")
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
                        
                        
                         If SystemOptions.CustomerhavethreeAccounts = False Then
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
                        ModAccounts.EditAccount rs("Account_Code1").value, Me.XPTxtCusName.Text & "    ‘Ìﬂ«    Õ  «· Õ’Ì·  ", XPTxtCusNamee.Text & "  Cheque Box ", , , , , , , , , , , , , , , , , True
                        Else
                                               rs("Account_Code1").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.Text) & "   ‘Ìﬂ«    Õ  «· Õ’Ì· ", True, False, XPTxtCusNamee.Text & "  Under Collection Cheque  ")
                                         

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
         rs("Boxmil").value = TxtBox.Text
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
        rs("JobTitle").value = Trim(TXTJobTitle.Text)
        rs("Salary").value = val(TxtSalary.Text)
        rs("JobAddress").value = Trim(TxtJobAddress.Text)
        rs("JobTel").value = Trim(TxtJobTel.Text)
        rs("JobTelConvert").value = Trim(TxtJobTelConvert.Text)
        rs("HomeTel").value = Trim(TxtHomeTel.Text)
        rs("Mobile1").value = Trim(TxtMobile1.Text)
        rs("Mobile2").value = Trim(TxtMobile2.Text)
      
        rs.update
'////////////saveLocations
   Dim StrRecID As Double
   Dim sql As String
   Dim i As Double
    Dim Rs4 As New ADODB.Recordset
    sql = "SELECT  *  from TblCustomersLocations Where (1 = -1)"
    Rs4.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.Grid
    For i = 1 To .Rows - 1
    If .TextMatrix(i, .ColIndex("Name")) <> "" Or .TextMatrix(i, .ColIndex("NameE")) <> "" Then
    If val(.TextMatrix(i, .ColIndex("ID"))) = 0 Then
    StrRecID = new_id("TblCustomersLocations", "ID", "")
    
    Else
    StrRecID = val(.TextMatrix(i, .ColIndex("ID")))
    End If
    Rs4.AddNew
    Rs4("ID").value = StrRecID
    Rs4("Cusid").value = val(XPTxtCusID.Text)
    Rs4("FirstClass").value = IIf(.TextMatrix(i, .ColIndex("FirstClass")) = "", Null, val(.TextMatrix(i, .ColIndex("FirstClass"))))
    Rs4("CurrClass").value = IIf(.TextMatrix(i, .ColIndex("CurrClass")) = "", Null, val(.TextMatrix(i, .ColIndex("CurrClass"))))
    Rs4("Name").value = IIf(.TextMatrix(i, .ColIndex("Name")) = "", Null, .TextMatrix(i, .ColIndex("Name")))
    Rs4("NameE").value = IIf(.TextMatrix(i, .ColIndex("NameE")) = "", Null, .TextMatrix(i, .ColIndex("NameE")))
    Rs4("MangerName").value = IIf(.TextMatrix(i, .ColIndex("MangerName")) = "", Null, .TextMatrix(i, .ColIndex("MangerName")))
    Rs4("ClassID").value = IIf(.TextMatrix(i, .ColIndex("ClassID")) = "", Null, val(.TextMatrix(i, .ColIndex("ClassID"))))
    Rs4("MF").value = IIf(.TextMatrix(i, .ColIndex("MF")) = "", Null, val(.TextMatrix(i, .ColIndex("MF"))))
    Rs4("Level").value = IIf(.TextMatrix(i, .ColIndex("Level")) = "", Null, val(.TextMatrix(i, .ColIndex("Level"))))
    
    Rs4("TxtIQAMA").value = IIf(.TextMatrix(i, .ColIndex("TxtIQAMA")) = "", Null, .TextMatrix(i, .ColIndex("TxtIQAMA")))
    Rs4("TxtPassport").value = IIf(.TextMatrix(i, .ColIndex("TxtPassport")) = "", Null, .TextMatrix(i, .ColIndex("TxtPassport")))
    
    Rs4("DBRegisterDate").value = IIf(.TextMatrix(i, .ColIndex("DBRegisterDate")) = "", Null, .TextMatrix(i, .ColIndex("DBRegisterDate")))
    Rs4("DBENDDATE").value = IIf(.TextMatrix(i, .ColIndex("DBENDDATE")) = "", Null, .TextMatrix(i, .ColIndex("DBENDDATE")))
    
    Rs4("DCNAtionalID").value = IIf(.TextMatrix(i, .ColIndex("DCNAtionalID")) = "", Null, val(.TextMatrix(i, .ColIndex("DCNAtionalID"))))
    Rs4("STDUDENTStatusID").value = IIf(.TextMatrix(i, .ColIndex("STDUDENTStatusID")) = "", Null, val(.TextMatrix(i, .ColIndex("STDUDENTStatusID"))))
    
    Rs4.update
    End If
    Next i
    End With
'////////////saveLocations


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
                   
                    Account_Code_dynamic1 = get_account_code_branch(59, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Else
                            MsgBox "No Branch was created", vbCritical
                        End If
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            Else
                                MsgBox "No opennig plance account was specified in this branch for this operation", vbCritical
                            End If
                            GoTo ErrTrap
                        End If
                    End If
            
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code").value, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText), , val(TxtShowQty1) - val(TxtShowQty1c), val(TxtshowPrice1) - val(TxtshowPrice1C), val(TxtshowPrice2) - val(TxtshowPrice2C), val(TxtSalaries1) - val(TxtSalaries1C), val(TxtSalaries2) - val(TxtSalaries2C)) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText), , val(TxtShowQty1c) - val(TxtShowQty1), val(TxtshowPrice1C) - val(TxtshowPrice1), val(TxtshowPrice2C) - val(TxtshowPrice2), val(TxtSalaries1C) - val(TxtSalaries1), val(TxtSalaries2C) - val(TxtSalaries2)) = False Then
                        GoTo ErrTrap
                    End If
                    
                    '  If ModAccounts.AddNewDev(LngDevID, 2, "a2a1a1", _
                       Val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    '         GoTo ErrTrap
                    ' End If
                ElseIf Me.OptType(1).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(59, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Else
                            MsgBox "No Branch was created", vbCritical
                        End If
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            Else
                                MsgBox "No opennig plance account was specified in this branch for this operation", vbCritical
                            End If
                            GoTo ErrTrap
                        End If
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText), , val(TxtShowQty1c) - val(TxtShowQty1), val(TxtshowPrice1C) - val(TxtshowPrice1), val(TxtshowPrice2C) - val(TxtshowPrice2), val(TxtSalaries1C) - val(TxtSalaries1), val(TxtSalaries2C) - val(TxtSalaries2)) = False Then
                        GoTo ErrTrap
                    End If

                    ' If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
                      Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    
                    '       GoTo ErrTrap
                    'End If
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code").value, Round(Me.TxtOpenBalance.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText), , val(TxtShowQty1) - val(TxtShowQty1c), val(TxtshowPrice1) - val(TxtshowPrice1C), val(TxtshowPrice2) - val(TxtshowPrice2C), val(TxtSalaries1) - val(TxtSalaries1C), val(TxtSalaries2) - val(TxtSalaries2C)) = False Then
                        GoTo ErrTrap
                    End If
                End If

                '   update_account_opening_balance rs("Account_Code").value
                'update_account_opening_balance Account_Code_dynamic1
                 
            End If
        End If




If SystemOptions.CustomerhavethreeAccounts = True Then
' 2
     If Me.OptType1(0).value = True Or Me.OptType1(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
 
                LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
            
                If Me.OptType1(0).value = True Then
                   
                    Account_Code_dynamic1 = get_account_code_branch(59, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Else
                            MsgBox "No Branch was created", vbCritical
                        End If
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            Else
                                MsgBox "No opennig plance account was specified in this branch for this operation", vbCritical
                            End If
                            GoTo ErrTrap
                        End If
                    End If
            
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code1").value, Round(Me.TxtOpenBalance1.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance1.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
                    '  If ModAccounts.AddNewDev(LngDevID, 2, "a2a1a1", _
                       Val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    '         GoTo ErrTrap
                    ' End If
                ElseIf Me.OptType1(1).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(59, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Else
                            MsgBox "No Branch was created", vbCritical
                        End If
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            Else
                                MsgBox "No opennig plance account was specified in this branch for this operation", vbCritical
                            End If
                            GoTo ErrTrap
                        End If
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance1.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    ' If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
                      Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    
                    '       GoTo ErrTrap
                    'End If
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code1").value, Round(Me.TxtOpenBalance1.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText)) = False Then
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
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Else
                            MsgBox "No Branch was created", vbCritical
                        End If
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            Else
                                MsgBox "No opennig plance account was specified in this branch for this operation", vbCritical
                            End If
                            GoTo ErrTrap
                        End If
                    End If
            
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code2").value, Round(Me.TxtOpenBalance2.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance2.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
                    '  If ModAccounts.AddNewDev(LngDevID, 2, "a2a1a1", _
                       Val(Me.TxtOpenBalance.text), 1, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    '         GoTo ErrTrap
                    ' End If
                ElseIf Me.OptType2(1).value = True Then
                    Account_Code_dynamic1 = get_account_code_branch(59, my_branch)
                
                    If Account_Code_dynamic1 = "NO branch" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Else
                            MsgBox "No Branch was created", vbCritical
                        End If
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            Else
                                MsgBox "No opennig plance account was specified in this branch for this operation", vbCritical
                            End If
                            GoTo ErrTrap
                        End If
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance2.Text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    ' If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
                      Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    
                    '       GoTo ErrTrap
                    'End If
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code2").value, Round(Me.TxtOpenBalance2.Text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.Text) & "  " & Trim$(Me.XPTxtCusNamee.Text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If

                '   update_account_opening_balance rs("Account_Code").value
                'update_account_opening_balance Account_Code_dynamic1
                 
            End If
        End If

End If
      '////////////////////////////////////////////////////////////
        Dim RsCusCar As ADODB.Recordset
        Set RsCusCar = New ADODB.Recordset
        StrSQL = "select * from TblCusCar  order by id "
        RsCusCar.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        With VSFlexGrid1
     
       ' FgInstallments.Rows = val(TxtPaymentCount.text) + 1
        Dim AllID As String
    '    RsCusCar.MoveFirst
       Dim j As Integer
        For j = VSFlexGrid1.FixedRows To VSFlexGrid1.Rows - 1
           If .TextMatrix(j, .ColIndex("BoardNo")) <> "" Then
                     If Me.TxtModFlg.Text = "N" Then
                        FormNamee = "FrmCarAuthontication"
                    End If
                    If .TextMatrix(j, .ColIndex("ID")) = "" Then
                            RsCusCar.AddNew
                            RsCusCar("ID") = CStr(new_id("TblCusCar", "ID", "", True))
                    Else
                            RsCusCar.Find " ID ='" & val(.TextMatrix(j, .ColIndex("ID"))) & "'", , adSearchForward, adBookmarkFirst
                            
                            If RsCusCar.EOF Or RsCusCar.BOF Then
                                    Exit Sub
                            End If
                    End If
                     RsCusCar("PartPrice").value = IIf(.TextMatrix(j, .ColIndex("PartPrice")) = "", Null, val(.TextMatrix(j, .ColIndex("PartPrice"))))
                     RsCusCar("Price").value = IIf(.TextMatrix(j, .ColIndex("Price")) = "", Null, val(.TextMatrix(j, .ColIndex("Price"))))
                     RsCusCar("Serial").value = IIf(.TextMatrix(j, .ColIndex("serial")) = "", Null, .TextMatrix(j, .ColIndex("serial")))
                     RsCusCar("BoardNo").value = IIf(.TextMatrix(j, .ColIndex("BoardNo")) = "", "", .TextMatrix(j, .ColIndex("BoardNo")))
                     RsCusCar("nBoardNo").value = IIf(.TextMatrix(j, .ColIndex("nBoardNo")) = "", "", .TextMatrix(j, .ColIndex("nBoardNo")))
                     RsCusCar("ChasisNo").value = IIf(.TextMatrix(j, .ColIndex("ChasisNo")) = "", Null, .TextMatrix(j, .ColIndex("ChasisNo")))
                     RsCusCar("BrandID").value = IIf(.TextMatrix(j, .ColIndex("BrandID")) = "", Null, .TextMatrix(j, .ColIndex("BrandID")))
                     RsCusCar("ModelID").value = IIf(.TextMatrix(j, .ColIndex("ModelID")) = "", Null, .TextMatrix(j, .ColIndex("ModelID")))
                     RsCusCar("Count").value = IIf(.TextMatrix(j, .ColIndex("Count")) = "", 0, .TextMatrix(j, .ColIndex("Count")))
                     RsCusCar("CityID").value = IIf(.TextMatrix(j, .ColIndex("CityID")) = "", Null, .TextMatrix(j, .ColIndex("CityID")))
                     RsCusCar("Rate").value = IIf(.TextMatrix(j, .ColIndex("Rate")) = "", 0, .TextMatrix(j, .ColIndex("Rate")))
                     RsCusCar("CustomerID").value = val(XPTxtCusID.Text)
                     RsCusCar("DriverName").value = IIf(.TextMatrix(j, .ColIndex("DriverName")) = "", Null, .TextMatrix(j, .ColIndex("DriverName")))
                     RsCusCar("DriverTel").value = IIf(.TextMatrix(j, .ColIndex("DriverTel")) = "", Null, .TextMatrix(j, .ColIndex("DriverTel")))
                     RsCusCar("accessory").value = IIf(.TextMatrix(j, .ColIndex("accessory")) = "", Null, .TextMatrix(j, .ColIndex("accessory")))
                     RsCusCar("TypeTransID").value = IIf(.TextMatrix(j, .ColIndex("TypeTransID")) = "", Null, val(.TextMatrix(j, .ColIndex("TypeTransID"))))
                     RsCusCar("CarModelID").value = IIf(.TextMatrix(j, .ColIndex("CarModelID")) = "", Null, val(.TextMatrix(j, .ColIndex("CarModelID"))))
                     RsCusCar("ColorID").value = IIf(.TextMatrix(j, .ColIndex("ColorID")) = "", Null, val(.TextMatrix(j, .ColIndex("ColorID"))))
                     RsCusCar.update
                    
                If j = VSFlexGrid1.FixedRows Then
                    AllID = RsCusCar("ID").value
                Else
                    AllID = AllID & "  ,  " & CStr(RsCusCar("ID").value)
                End If
                    
            End If
           Next
        End With
        
        
         'Dim strSQL As String
         If AllID <> "" Then
                StrSQL = "delete from TblCusCar  where customerid = " & val(XPTxtCusID.Text) & " and  id not in  ( " & AllID & "  ) "
                 Cn.Execute StrSQL, , adExecuteNoRecords
         End If
        Cn.CommitTrans
        
        StrSQL = "delete from TblCustomerSizes  where cusId = " & val(XPTxtCusID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        Dim s As String
        s = "Select * from TblCustomerSizes Where CusId = " & val(Me.XPTxtCusID.Text)
       
        saveGrid s, fg, "SizesNamesID", "", "CusID", val(Me.XPTxtCusID.Text)
        
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        'update_account_opening_balance Me.DcboDebitSide.BoundText
        ' update_account_opening_balance Me.DcboCreditSide.BoundText
        CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·⁄„Ì· " & CHR(13)
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
        
        PassData
        If FrmCustemers.Height = 10560 Then
            
            
            
            If FormNamee = "FrmCarAuthontication" Then
                FrmCarAuthontication.Cmd_Click (0)
                  FrmCarAuthontication.mFromCustomerForm = True
                      
                  FrmCarAuthontication.TxtClientCode.Text = Me.TxtFullcode
                  FrmCarAuthontication.retInfoCustomer Me.TxtFullcode
                  Unload Me
            
            ElseIf FormNamee = "FrmItemsClass" Then
                FrmItemsClass.TxtClientCode = txtid
                FrmItemsClass.TxtCliientName = XPTxtCusName
                FrmItemsClass.TxtCusID = XPTxtCusID
                FrmItemsClass.retInfoCustomer
                Unload Me
                End If
        End If
        
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
Private Sub DcbCarType_Change()
Dim Dcombos As ClsDataCombos
      Set Dcombos = New ClsDataCombos
    
      If val(Me.dcBrand.BoundText) <> 0 Then
      
   Dcombos.GetTblCarModels Me.DcbCarModel, , val(Me.dcBrand.BoundText)
   End If
End Sub

Private Sub dcBrand_Click(Area As Integer)
Dim Dcombos As ClsDataCombos
      Set Dcombos = New ClsDataCombos
    
      If val(Me.dcBrand.BoundText) <> 0 Then
   Dcombos.GetTblCarModels Me.DcbCarModel, , val(Me.dcBrand.BoundText)
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
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    ÿ·» ’Ì«‰…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(0), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    ’—› ﬁÿ⁄ €Ì«—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(2), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…     ›« Ê—… ’Ì«‰…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(3), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   ‰»ÌÂ«  «·⁄„·«¡     ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(4), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   ﬁ«—Ì— «·’Ì«‰Â    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(5), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   ‘«‘… ›Õ’ ﬂ„»ÌÊ —  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(6), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   «·‰»ÌÂ«  «·„› ÊÕ…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(7), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    «· ‰»ÌÂ«  «·⁄«„…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(8), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    «·⁄„Ê·«  «·„” Õﬁ…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(9), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    «·⁄„·«¡   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(10), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
           With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…     ﬁ«—Ì— «·⁄„Ê·«    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(10), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·⁄„Ì·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„Ì· «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„Ì·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„Ì·" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, BolRtl
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New..." & Wrap & "Add New Customer Data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print the current customer data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit the current customer data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the current editing or Save the new customer data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo..." & Wrap & "Undo in the adding new record" & Wrap & "OR undo editing current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete...." & Wrap & "Delete the current customer data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search" & Wrap & "Search for a customer..." & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next..." & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last..." & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
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
mCustId = val(txtid)
If FormNamee = "FrmDefinCompItem" Then
  '  FrmDefinCompItem.mCustId = mCustId
End If
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
    XPLbl(6).Caption = "VAT No."
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    XPLbl(22).Caption = "Type"
XPLbl(23).Caption = "Color  "
    XPLbl(24).Caption = "Plate No"
    
    
    creditlocked.Caption = "Cancel Debt Deal"
    chkCustomerandVendor.Caption = "Customer / Supplier"
    Label1(2).Caption = "Type"
    lbl(71).Caption = "Banck Name"
    lbl(70).Caption = "Banck Account"
    XPLbl(14).Caption = "Currency"
    Label3.Caption = "Branch"
    Cmd(11).Caption = "Attachments."
    lbl(72).Caption = "IBAN"
    lbl(73).Caption = "Bank Code"
    lbl(74).Caption = "Bank Address"
     
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

    OptType1(0).Caption = "Debit"
    OptType1(1).Caption = "Credit"
    OptType1(2).Caption = "Un Sign"
    lbl(47).Caption = "Balance Value"
    lbl(48).Caption = "Record Date"
           With CboPayMentType
            .Clear
            'AddItem "Cash"
            
            .AddItem "Cash"
            .AddItem "Credit"
        End With
    
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
     
    ALLButton2.Caption = "Show Balance"
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
        .AddItem "Percentage"
    End With

    With CboDiscountTypePur
        .Clear
        .AddItem "No"
        .AddItem "Value"
        .AddItem "Percentage"
    End With
   With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("BoardNo")) = "Board No"
        .TextMatrix(0, .ColIndex("nBoardNo")) = "Board No E"
        .TextMatrix(0, .ColIndex("ChasisNo")) = "Chasis No"
        .TextMatrix(0, .ColIndex("Brand")) = "Brand"
        .TextMatrix(0, .ColIndex("Model")) = "Model"
        .TextMatrix(0, .ColIndex("count")) = "Set Count"
        .TextMatrix(0, .ColIndex("Rate")) = "Pass. Rate"
        .TextMatrix(0, .ColIndex("City")) = "For"
        .TextMatrix(0, .ColIndex("DriverName")) = "Driver Name"
        .TextMatrix(0, .ColIndex("DriverTel")) = "Driver Phone"
        .TextMatrix(0, .ColIndex("EndDate")) = "License End Date"
        .TextMatrix(0, .ColIndex("accessory")) = "Accessory"
        .TextMatrix(0, .ColIndex("Price")) = "Value"
        .TextMatrix(0, .ColIndex("TypeTransID")) = "Type"
        
    End With
        XPLbl(9).Caption = "Chasis No"
        XPLbl(11).Caption = "Driver Name"
        XPLbl(13).Caption = "Model"
        lbl(77).Caption = "Brand"
        XPLbl(8).Caption = "For"
        XPLbl(19).Caption = "Set Count"
        XPLbl(21).Caption = "Telephone"
        XPLbl(20).Caption = "Pass. Rate"
        XPLbl(15).Caption = "Accessory"
        XPLbl(17).Caption = "Accessory Value"
        XPLbl(16).Caption = "Value"
        XPLbl(18).Caption = "Type"
    Frame6.Caption = "Board No."
   ' XPLbl(3).Caption = "Exp."
   
    
     
    
    
    Cmd(14).Caption = "Add"
    Cmd(13).Caption = "Delete"
    Cmd(15).Caption = "Delete All"
    XPLbl(2).Caption = "Client NO."
    XPLbl(3).Caption = "End User"

lbl(78).Caption = "Payment"
Label4(0).Caption = "Group"
Label4(1).Caption = "Groups2"
Label4(2).Caption = "Branch"
    locked.Caption = "locked"
    ALLButton1.Caption = "Reason"
    lbl(32).Caption = "reason"
    lbl(30).Caption = "period"
    lbl(31).Caption = "period"
    XPLbl(12).Caption = "Expire date"
    Me.C1Tab1.TabCaption(0) = "Data"
    Me.C1Tab1.TabCaption(1) = "Specific Data"
    
    Me.C1Tab1.TabCaption(4) = "Others"
    
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
    
    HajEnter_Out(0).Caption = "Domestic"
    HajEnter_Out(1).Caption = "International"
    
    TypeHaj.Caption = "Hajj"
    TypeOmrh.Caption = "Omrah"
    
    If SystemOptions.AllowScInterface = False Then
        FrmCustemers.Caption = "Customers Data"
        EleHeader.Caption = FrmCustemers.Caption
        
        Label1(5).Caption = "Arabic Name"
        Label1(4).Caption = "English Name"
        Label1(3).Caption = "Responsible Person"
        With Grid
            .TextMatrix(0, .ColIndex("Ser")) = "Ser"
            .TextMatrix(0, .ColIndex("Name")) = "Arabic Name"
            .TextMatrix(0, .ColIndex("NameE")) = "English Name"
            .TextMatrix(0, .ColIndex("MangerName")) = "Responsible Person"
        End With
        Me.C1Tab1.TabCaption(2) = "Represented Entities"
    ElseIf SystemOptions.AllowScInterface = True Then
        FrmCustemers.Caption = "Parents "
        EleHeader.Caption = FrmCustemers.Caption
        
        Label1(5).Caption = "Arabic Name"
        Label1(4).Caption = "English Name"
        Label1(3).Caption = "Responsible Person"
        Label1(7).Caption = "Current Semester"
        Label1(9).Caption = "Class"
        Label1(10).Caption = "Grade"
        Label1(11).Caption = "Passport No."
        Label1(12).Caption = "Iqama No."
        Label1(13).Caption = "Nationality"
        lbl(75).Caption = "Registration Date"
        lbl(76).Caption = "Check-out Date"
        Label1(14).Caption = "Status"
        With Grid
            .TextMatrix(0, .ColIndex("Ser")) = "Ser"
            .TextMatrix(0, .ColIndex("Name")) = "Arabic Name"
            .TextMatrix(0, .ColIndex("NameE")) = "English Name"
            .TextMatrix(0, .ColIndex("STDUDENTStatusID")) = "Status"
            .TextMatrix(0, .ColIndex("DBRegisterDate")) = "Registration Date"
            .TextMatrix(0, .ColIndex("DBENDDATE")) = "Check-out Date"
            .TextMatrix(0, .ColIndex("DCNAtionaNAme")) = "Nationality"
            .TextMatrix(0, .ColIndex("TxtPassport")) = "Passport No."
            .TextMatrix(0, .ColIndex("TxtIQAMA")) = "Iqama No."
            .TextMatrix(0, .ColIndex("MangerName")) = "Responsible Person"
            .TextMatrix(0, .ColIndex("CurrClass")) = "Current Semester"
            .TextMatrix(0, .ColIndex("Level")) = "Level"
            .TextMatrix(0, .ColIndex("Class")) = "Class"
            
        End With
        Me.C1Tab1.TabCaption(2) = "Students"
    End If
    ISButton2.Caption = "Add"
    Cmd(12).Caption = "Delete Line"
    Me.C1Tab1.TabCaption(3) = "Data Of Cars"

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

    Dcombo.GetBranches dcBranch
  
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
    End If

End Sub

Private Sub XPTxtCusName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtCusNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Function FindString(Control As Control, FindStr As String, Optional StartPos As Integer = 1) As Boolean
Dim a As Integer
    a = InStr(StartPos, LCase$(Control.Text), LCase$(FindStr))
    If a = 0 Then
        FindString = False
    Else
        FindString = True
        Control.SetFocus
        Control.SelStart = a - 1
        Control.SelLength = Len(FindStr)
    End If
End Function

Private Sub XPTxtmobile_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtmobile.Text, 1)
End Sub
