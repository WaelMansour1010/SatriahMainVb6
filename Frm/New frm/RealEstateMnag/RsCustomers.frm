VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form RsCustomers 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·„” √Ã—Ì‰"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13980
   HelpContextID   =   50
   Icon            =   "RsCustomers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   13980
   Begin VB.TextBox TxtVATNO 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   234
      Top             =   3000
      Width           =   3075
   End
   Begin VB.CheckBox chkSendMessage 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«„ﬂ«‰Ì… «—”«· «·—”«∆·"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   227
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   1920
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
   Begin VB.TextBox TxtMap 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3900
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   223
      Top             =   2280
      Width           =   4665
   End
   Begin VB.CheckBox locked 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Ìﬁ«› «· ⁄«„·"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   221
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox TxtRecordNo 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5490
      MaxLength       =   50
      TabIndex        =   219
      Top             =   1800
      Width           =   3075
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‘—ﬂ« "
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   218
      Top             =   600
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "√›—«œ"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   217
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox CboSaleType 
      Height          =   315
      Left            =   18120
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   215
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox TxtFullcode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   14400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   207
      Top             =   960
      Width           =   1125
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   2880
      TabIndex        =   191
      Top             =   9240
      Width           =   7935
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   5
         Left            =   1440
         Picture         =   "RsCustomers.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   4
         Left            =   2880
         Picture         =   "RsCustomers.frx":07D5
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   3
         Left            =   3600
         Picture         =   "RsCustomers.frx":0D2D
         Style           =   1  'Graphical
         TabIndex        =   201
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   2
         Left            =   5040
         Picture         =   "RsCustomers.frx":11E6
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   1
         Left            =   3240
         Picture         =   "RsCustomers.frx":16B6
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "RsCustomers.frx":1B57
         Height          =   555
         Index           =   0
         Left            =   7200
         Picture         =   "RsCustomers.frx":8E89
         Style           =   1  'Graphical
         TabIndex        =   198
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "RsCustomers.frx":9430
         Height          =   555
         Index           =   6
         Left            =   5760
         Picture         =   "RsCustomers.frx":10762
         Style           =   1  'Graphical
         TabIndex        =   197
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "RsCustomers.frx":10C03
         Height          =   555
         Index           =   7
         Left            =   4320
         Picture         =   "RsCustomers.frx":17F35
         Style           =   1  'Graphical
         TabIndex        =   196
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   8
         Left            =   2160
         Picture         =   "RsCustomers.frx":187C5
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "RsCustomers.frx":18CAA
         Height          =   555
         Index           =   9
         Left            =   720
         Picture         =   "RsCustomers.frx":1FFDC
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "RsCustomers.frx":204FC
         Height          =   555
         Index           =   10
         Left            =   6480
         Picture         =   "RsCustomers.frx":20AE3
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "RsCustomers.frx":210CA
         Height          =   555
         Index           =   11
         Left            =   0
         Picture         =   "RsCustomers.frx":283FC
         Style           =   1  'Graphical
         TabIndex        =   192
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
      Left            =   16320
      RightToLeft     =   -1  'True
      TabIndex        =   122
      Top             =   2160
      Width           =   2925
      Begin VB.ComboBox CboDiscountTypePur 
         Height          =   315
         Left            =   270
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   124
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtDiscountValuePur 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         RightToLeft     =   -1  'True
         TabIndex        =   123
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
         TabIndex        =   127
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
         TabIndex        =   126
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
         TabIndex        =   125
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
      Left            =   15360
      RightToLeft     =   -1  'True
      TabIndex        =   116
      Top             =   2160
      Width           =   3645
      Begin VB.TextBox TxtDiscountValue 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   300
         RightToLeft     =   -1  'True
         TabIndex        =   118
         Top             =   600
         Width           =   1425
      End
      Begin VB.ComboBox CboDiscountType 
         Height          =   315
         Left            =   270
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   117
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.TextBox txtCustGID 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2790
      MaxLength       =   50
      TabIndex        =   57
      Top             =   930
      Width           =   1365
   End
   Begin VB.CheckBox chkCustomerandVendor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄„Ì· Ê„Ê—œ"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   15480
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   11070
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   600
      Width           =   1365
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
      Height          =   285
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   1
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
      Height          =   2145
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
         TabIndex        =   213
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
         TabIndex        =   214
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
      BackColor       =   &H0080FFFF&
      Height          =   345
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
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
      Caption         =   "»Ì«‰«  «·„” √Ã—Ì‰  "
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
         ButtonImage     =   "RsCustomers.frx":28F90
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
         ButtonImage     =   "RsCustomers.frx":2932A
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
         ButtonImage     =   "RsCustomers.frx":296C4
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
         ButtonImage     =   "RsCustomers.frx":29A5E
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
         Left            =   5400
         Picture         =   "RsCustomers.frx":29DF8
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Image Img 
         Height          =   480
         Left            =   7080
         Picture         =   "RsCustomers.frx":2DA60
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
      Left            =   10635
      TabIndex        =   21
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
      Left            =   9915
      TabIndex        =   22
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
      Left            =   9195
      TabIndex        =   23
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
      Left            =   8355
      TabIndex        =   24
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
      Left            =   7530
      TabIndex        =   25
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
      Left            =   3360
      TabIndex        =   26
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
      Left            =   6690
      TabIndex        =   27
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
      Left            =   5670
      TabIndex        =   28
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
      Width           =   1395
      _ExtentX        =   2461
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
      Height          =   4935
      Left            =   120
      TabIndex        =   58
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
      Caption         =   "»Ì«‰«  «”«”Ì…|»Ì«‰«  «·⁄„·"
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
         Height          =   4515
         Left            =   14550
         TabIndex        =   59
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
         Begin VB.Frame Frame3 
            Caption         =   "»Ì«‰«  Œ«’… »«·œÂ»"
            ForeColor       =   &H000000FF&
            Height          =   3015
            Left            =   17160
            RightToLeft     =   -1  'True
            TabIndex        =   161
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
               TabIndex        =   182
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
               TabIndex        =   181
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
               TabIndex        =   180
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
               TabIndex        =   179
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
               TabIndex        =   178
               Top             =   360
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice1C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   720
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice2C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   1080
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries1C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   175
               Top             =   1440
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries2C 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   174
               Top             =   1800
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   1800
               Width           =   1425
            End
            Begin VB.TextBox TxtSalaries1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   1440
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   1080
               Width           =   1425
            End
            Begin VB.TextBox TxtshowPrice1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   164
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
               TabIndex        =   162
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
               TabIndex        =   184
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
               TabIndex        =   183
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
               TabIndex        =   173
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
               TabIndex        =   172
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
               TabIndex        =   170
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
               TabIndex        =   169
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
               TabIndex        =   167
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
               TabIndex        =   165
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
               TabIndex        =   163
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.CommandButton CDMOldContract 
            Caption         =   "⁄ﬁÊœ ”«»ﬁ…"
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   111
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
            TabIndex        =   87
            Top             =   240
            Width           =   6135
            Begin VB.ComboBox CboSex 
               Height          =   315
               ItemData        =   "RsCustomers.frx":2E72A
               Left            =   3000
               List            =   "RsCustomers.frx":2E72C
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   109
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TxtMobile2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3000
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   3480
               Width           =   1485
            End
            Begin VB.TextBox TxtMobile1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   3120
               Width           =   1485
            End
            Begin VB.TextBox TxtHomeTel 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3000
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   3120
               Width           =   1485
            End
            Begin VB.TextBox TxtJobTelConvert 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   2760
               Width           =   1485
            End
            Begin VB.TextBox TxtJobTel 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3000
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   98
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
               TabIndex        =   95
               Top             =   2160
               Width           =   4425
            End
            Begin VB.TextBox TxtSalary 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2400
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   1680
               Width           =   2085
            End
            Begin VB.TextBox TXTJobTitle 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   1320
               Width           =   4365
            End
            Begin VB.TextBox TxtCompany 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   960
               Width           =   4365
            End
            Begin MSDataListLib.DataCombo DcboCountryID2 
               Height          =   315
               Left            =   1440
               TabIndex        =   110
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
               TabIndex        =   108
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
               TabIndex        =   107
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
               TabIndex        =   105
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
               TabIndex        =   103
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
               TabIndex        =   101
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
               TabIndex        =   99
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
               TabIndex        =   97
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
               TabIndex        =   96
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
               TabIndex        =   94
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
               TabIndex        =   92
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
               TabIndex        =   90
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
               TabIndex        =   88
               Top             =   240
               Width           =   825
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4515
         Index           =   2
         Left            =   45
         TabIndex        =   60
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
            Left            =   -480
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   -30
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
            Begin VB.TextBox txtNoOFDigitUser 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   0
               Left            =   600
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   248
               Top             =   1650
               Width           =   1755
            End
            Begin VB.TextBox txtNoOFDigitUser 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   8
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   241
               Top             =   870
               Width           =   1005
            End
            Begin VB.TextBox txtNoOFDigitUser 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   3
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   240
               Top             =   1140
               Width           =   1005
            End
            Begin VB.TextBox txtNoOFDigitUser 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   5
               Left            =   2520
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   239
               Top             =   1260
               Width           =   1005
            End
            Begin VB.TextBox txtNoOFDigitUser 
               Alignment       =   1  'Right Justify
               BackColor       =   &H008080FF&
               Height          =   315
               Index           =   10
               Left            =   2520
               MaxLength       =   2
               RightToLeft     =   -1  'True
               TabIndex        =   238
               Top             =   900
               Width           =   1005
            End
            Begin VB.TextBox txtNoOFDigitUser 
               Alignment       =   1  'Right Justify
               BackColor       =   &H008080FF&
               Height          =   315
               Index           =   2
               Left            =   2490
               RightToLeft     =   -1  'True
               TabIndex        =   237
               Top             =   210
               Width           =   1005
            End
            Begin VB.TextBox txtNoOFDigitUser 
               Alignment       =   1  'Right Justify
               BackColor       =   &H008080FF&
               Height          =   315
               Index           =   4
               Left            =   600
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   236
               Tag             =   "4 digit at least"
               Top             =   180
               Width           =   1005
            End
            Begin VB.TextBox TxtEntry 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10080
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   211
               Top             =   480
               Width           =   2775
            End
            Begin VB.ComboBox DcbDigCustomer 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   10080
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   190
               Top             =   1920
               Width           =   2775
            End
            Begin VB.TextBox TxtZib 
               Alignment       =   1  'Right Justify
               BackColor       =   &H008080FF&
               Height          =   315
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   186
               Top             =   1560
               Width           =   2775
            End
            Begin VB.TextBox TxtBox 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   185
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
               Height          =   1905
               Index           =   5
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   2160
               Width           =   4125
               Begin VB.TextBox TxtAddress 
                  Alignment       =   1  'Right Justify
                  Height          =   705
                  Left            =   150
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   153
                  Top             =   1140
                  Width           =   2625
               End
               Begin MSDataListLib.DataCombo DcboCountryID 
                  Height          =   315
                  Left            =   150
                  TabIndex        =   154
                  Top             =   150
                  Width           =   2625
                  _ExtentX        =   4630
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   8421631
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboGovernmentID 
                  Height          =   315
                  Left            =   150
                  TabIndex        =   155
                  Top             =   480
                  Width           =   2625
                  _ExtentX        =   4630
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   8421631
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboCityID 
                  Height          =   315
                  Left            =   180
                  TabIndex        =   156
                  Top             =   840
                  Width           =   2625
                  _ExtentX        =   4630
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   8421631
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
                  TabIndex        =   160
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
                  TabIndex        =   159
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
                  TabIndex        =   158
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
                  TabIndex        =   157
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
               Left            =   7080
               RightToLeft     =   -1  'True
               TabIndex        =   144
               Top             =   1560
               Width           =   2865
               Begin VB.TextBox TxtOpenBalance2 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   148
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
                  TabIndex        =   147
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
                  TabIndex        =   146
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
                  TabIndex        =   145
                  Top             =   210
                  Width           =   765
               End
               Begin MSComCtl2.DTPicker Dtp2 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   149
                  Top             =   870
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   236257283
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
                  TabIndex        =   151
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
                  TabIndex        =   150
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
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   240
               Width           =   2745
               Begin VB.OptionButton OptType1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   1950
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
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
                  TabIndex        =   139
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
                  TabIndex        =   138
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   1005
               End
               Begin VB.TextBox TxtOpenBalance1 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   137
                  Top             =   510
                  Width           =   1365
               End
               Begin MSComCtl2.DTPicker Dtp1 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   141
                  Top             =   870
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   236257283
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
                  TabIndex        =   143
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
                  TabIndex        =   142
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
               TabIndex        =   128
               Top             =   240
               Width           =   2865
               Begin VB.TextBox TxtOpenBalance 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
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
                  TabIndex        =   131
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
                  TabIndex        =   130
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
                  TabIndex        =   129
                  Top             =   210
                  Width           =   765
               End
               Begin MSComCtl2.DTPicker Dtp 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   133
                  Top             =   870
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   582
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   236257283
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
                  TabIndex        =   135
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
                  TabIndex        =   134
                  Top             =   930
                  Width           =   1215
               End
            End
            Begin VB.TextBox TxtE_mail 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   84
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
               Left            =   16560
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   240
               Width           =   5895
               Begin VB.TextBox TxtCreditlimitCredit 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2670
                  MaxLength       =   8
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   540
                  Width           =   1395
               End
               Begin VB.TextBox TxtCreditLimit 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2670
                  MaxLength       =   8
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   180
                  Width           =   1395
               End
               Begin VB.TextBox TxtDepitInterval 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   240
                  Width           =   495
               End
               Begin VB.TextBox TxtCreditInterval 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   600
                  Width           =   495
               End
               Begin VB.ComboBox dcDepitIntervalID 
                  Height          =   315
                  ItemData        =   "RsCustomers.frx":2E72E
                  Left            =   120
                  List            =   "RsCustomers.frx":2E730
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   240
                  Width           =   975
               End
               Begin VB.ComboBox dcCreditIntervalID 
                  Height          =   315
                  ItemData        =   "RsCustomers.frx":2E732
                  Left            =   120
                  List            =   "RsCustomers.frx":2E734
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
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
                  TabIndex        =   81
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
                  TabIndex        =   80
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
                  TabIndex        =   79
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
                  TabIndex        =   78
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
               TabIndex        =   67
               Top             =   3360
               Width           =   3375
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   435
                  Index           =   9
                  Left            =   60
                  TabIndex        =   68
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
                  TabIndex        =   70
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
                  TabIndex        =   69
                  Top             =   240
                  Width           =   1785
               End
            End
            Begin VB.TextBox XPMTxtRemarks 
               Alignment       =   1  'Right Justify
               Height          =   675
               Left            =   2520
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   65
               Top             =   4200
               Width           =   10185
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
               TabIndex        =   62
               Top             =   9360
               Width           =   2160
            End
            Begin MSDataListLib.DataCombo DboParentAccount 
               Height          =   315
               Left            =   600
               TabIndex        =   82
               Top             =   3600
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin DBPIXLib.DBPix20 DBPix201 
               Height          =   1695
               Left            =   4860
               TabIndex        =   231
               Top             =   1680
               Visible         =   0   'False
               Width           =   1575
               _Version        =   131072
               _ExtentX        =   2778
               _ExtentY        =   2990
               _StockProps     =   1
               BackColor       =   16777152
               _Image          =   "RsCustomers.frx":2E736
               ImageResampleWidth=   100
               ImageResampleHeight=   100
               ImageResampleMode=   1
               ImageSaveFormat =   0
               JPEGQuality     =   75
               JPEGEncoding    =   0
               JPEGColorMode   =   0
               JPEGNoRecompress=   -1  'True
               JPEGRotateWarning=   0
               PNGColorDepth   =   0
               PNGCompression  =   0
               PNGFilter       =   0
               PNGInterlace    =   1
               ImageDitherMethod=   3
               ImagePaletteMethod=   4
               ImagePreviewMode=   0   'False
               ImageKeepMetaData=   -1  'True
               UseAmbientBackcolor=   -1  'True
               ViewAsyncDecoding=   -1  'True
               ViewEnableMouseZoom=   -1  'True
               ViewInitialZoom =   0
               ViewHAlign      =   1
               ViewVAlign      =   1
               ViewMenuMode    =   0
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   495
               Left            =   720
               TabIndex        =   232
               Top             =   2520
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   873
               ButtonPositionImage=   1
               Caption         =   "«œ—«Ã «·»’„Â"
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
            Begin ImpulseButton.ISButton ISButton2 
               Height          =   495
               Left            =   720
               TabIndex        =   233
               Top             =   3000
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   873
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»’„Â"
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ 700"
               Height          =   375
               Index           =   85
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   249
               Top             =   1710
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„œÌ‰… «·›—⁄Ì…"
               Height          =   375
               Index           =   89
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   247
               Top             =   900
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·‘«—⁄2"
               Height          =   375
               Index           =   88
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   246
               Top             =   1230
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·„Œÿÿ"
               Height          =   375
               Index           =   87
               Left            =   3480
               RightToLeft     =   -1  'True
               TabIndex        =   245
               Top             =   1260
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·œÊ·…*"
               Height          =   255
               Index           =   86
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   244
               Top             =   900
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·‘«—⁄*"
               Height          =   375
               Index           =   90
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   243
               Top             =   180
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·„»‰Ï*"
               Height          =   255
               Index           =   91
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   242
               Top             =   210
               Width           =   1005
            End
            Begin VB.Shape Shape1 
               Height          =   1965
               Left            =   4770
               Top             =   1620
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ«Œ·Ì"
               Height          =   315
               Index           =   68
               Left            =   12990
               RightToLeft     =   -1  'True
               TabIndex        =   212
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   285
               Index           =   66
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   210
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
               TabIndex        =   209
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
               TabIndex        =   189
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
               TabIndex        =   188
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
               TabIndex        =   187
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
               TabIndex        =   85
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
               TabIndex        =   83
               Top             =   3600
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
               TabIndex        =   66
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
               TabIndex        =   63
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
            TabIndex        =   64
            Top             =   90
            Width           =   1125
         End
      End
   End
   Begin MSDataListLib.DataCombo DcCustomerType 
      Height          =   315
      Left            =   5490
      TabIndex        =   112
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
      Left            =   0
      TabIndex        =   114
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
      Left            =   6120
      TabIndex        =   204
      Top             =   8790
      Width           =   1545
      _ExtentX        =   2725
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
      TabIndex        =   205
      Top             =   600
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      CalendarBackColor=   12648447
      CustomFormat    =   "yyyy/M/d"
      Format          =   234094595
      CurrentDate     =   38718
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   255
      Left            =   1200
      TabIndex        =   222
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
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
      MICON           =   "RsCustomers.frx":2E74E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   11
      Left            =   4200
      TabIndex        =   225
      Top             =   8310
      Width           =   1245
      _ExtentX        =   2196
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
   Begin Dynamic_Byte.NourHijriCal Txt_DateExpLincH 
      Height          =   225
      Left            =   360
      TabIndex        =   226
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   397
   End
   Begin MSComCtl2.DTPicker BrithDate 
      Height          =   345
      Left            =   5490
      TabIndex        =   228
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   10383715
      Format          =   234160131
      CurrentDate     =   41640
   End
   Begin Dynamic_Byte.NourHijriCal BrithDateH 
      Height          =   345
      Left            =   7110
      TabIndex        =   229
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
   End
   Begin VB.Label XPLbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «· ”ÃÌ· VAT"
      Height          =   465
      Index           =   7
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   235
      Top             =   3000
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «·„Ì·«œ"
      Height          =   285
      Index           =   22
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   230
      Top             =   2640
      Width           =   930
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Œ—«∆ÿ ÃÊÃ·"
      Height          =   345
      Index           =   67
      Left            =   8370
      RightToLeft     =   -1  'True
      TabIndex        =   224
      Top             =   2280
      Width           =   1125
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”Ã·"
      Height          =   345
      Index           =   6
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   220
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "”Ì«”… «·»Ì⁄ "
      Height          =   405
      Index           =   10
      Left            =   20460
      RightToLeft     =   -1  'True
      TabIndex        =   216
      Top             =   1830
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
      TabIndex        =   208
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
      TabIndex        =   206
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
      TabIndex        =   115
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·⁄„Ì·"
      Height          =   285
      Index           =   2
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   113
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
      TabIndex        =   86
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÂÊÌ…/«·«ﬁ«„…"
      Height          =   345
      Index           =   5
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   960
      Width           =   1275
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
      TabIndex        =   18
      Top             =   1005
      Width           =   915
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·⁄„Ì·"
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
      Left            =   1590
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   8340
      Width           =   1455
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      Height          =   285
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   8340
      Width           =   615
   End
End
Attribute VB_Name = "RsCustomers"
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
Dim Account_Code_dynamic166 As String
Private m_DcboCustomers As DataCombo
Public index As Integer
Dim FTempLen As Integer
Dim FRegTemplate As String
Dim FRegTemp As Variant
Dim FingerCount As Long
Dim fpcHandle As Long
Dim FFingerNames() As String
Dim FMatchType As Integer

Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long


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
            Me.DcboCustomers.BoundText = val(XPTxtCusID.text)
 
        Case PriceList
     StrSQL = "SELECT * From TblCustemers where Type=2"
          'fill_combo FrmMainPriceList.DBCboSupplierName, StrSQL
          '  FrmMainPriceList.DBCboSupplierName.BoundText = val(XPTxtCusID.Text)

            '⁄—÷ «·√”⁄«—
        Case ShowPrice
           fill_combo FrmShowPrice.DBCboClientName, StrSQL
            FrmShowPrice.DBCboClientName.BoundText = val(XPTxtCusID.text)
    End Select
 Me.calledFromForm = False
Unload Me
    Exit Sub
ErrTrap:
End Sub

Private Sub ALLButton1_Click()
    Frame2.Visible = True
End Sub

Private Sub BrithDate_Change()
If Me.TxtModFlg.text <> "R" Then
BrithDateH.value = ToHijriDate(BrithDate.value)
End If
End Sub

Private Sub BrithDateH_LostFocus()
If Me.TxtModFlg.text <> "R" Then
      VBA.Calendar = vbCalGreg
    BrithDate.value = ToGregorianDate(BrithDateH.value)
 End If
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
    TxtOpenBalance.text = 0
    Cmd_Click (2)

End Function

Private Sub CDMOldContract_Click()
FrmOldContract.show
End Sub

Private Sub Cmd_Click(index As Integer)

    On Error GoTo ErrTrap

    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

    Me.Dtp.value = FirstPeriodDateInthisYear
 
    Dim Msg As String

    Select Case index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
    
            Txt_DateExpLincH.value = ToHijriDate(Date)
                
            Dim Account_Code_dynamic As String
            Account_Code_dynamic = get_account_code_branch(48, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» ··„·«ﬂ   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
       
                End If
            End If
        
            DboParentAccount.BoundText = Account_Code_dynamic
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

            Me.Dtp.value = FirstPeriodDateInthisYear
       '     Me.Dcbranch.BoundText = Current_branch
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

            If XPTxtCusID.text = 2 Then
                Msg = "·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰«  Â–« «·”Ã·"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            TxtModFlg.text = "E"

        Case 2

            Dim currentcode As String
        CREATEADDRESS
        
        If checkEeinvoice = False Then Exit Sub
            If txtid.text = "" Then
                currentcode = get_coding(branch_id, "TblCustemers", 15, Me.DCPreFix.text)

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



If SystemOptions.CreateInsuranceAccountForCustomers = True Then
    Account_Code_dynamic166 = get_account_code_branch(166, my_branch)
                
                    If Account_Code_dynamic166 = "NO branch" Then
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Exit Sub
                    Else

                                    If Account_Code_dynamic166 = "NO account" Then
                                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   √„Ì‰«  ··⁄Ì—  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                       Exit Sub
                                    End If
                    End If
                    
End If

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            If XPTxtCusID.text = 2 Then
                Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·”Ã·"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Del_Member

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
       FrmCustemerSearch.SearchType = 552
              FrmCustemerSearch.show vbModal
        
        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            '        Text1.text = 2
            print_report

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
            ShowReport IIf(IsNull(rs("Account_Code")), "", rs("Account_Code")), XPTxtCusName.text, FirstPeriod, Date
        Case 10
            If val(Me.XPTxtCusID.text) <> 0 Then
                print_report val(Me.XPTxtCusID.text)
        '" & val(XPTxtCusID.text) & ")"
        
            End If
        
        Case 11
                    On Error Resume Next
ShowAttachments DCPreFix.text & txtid.text, "270120152"
 
 
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
        Account_search.case_id = 21220142
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

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

  If 1 = 0 Then
  Me.Height = 9900
 Frame3.Visible = True
 Else
 Frame3.Visible = False
 Me.Height = 9270
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
    Dcombos.GetCodeing Me.DCPreFix, 4
    ' Dcombos.GetEmployees Me.DCEmp
    Dcombos.GetSalesRepData Me.DCEmP
    Me.Dtp.value = Date
    DtRecord.value = Date
    StrSQL = "select * From TblCustemers where type=56"
        If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
        End If
   If SystemOptions.usertype <> UserAdminAll Then
 
          If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  Empid = " & user_id
             End If
  
        Me.dcBranch.Enabled = True
       ' DCEmP.Enabled = False
     
    End If


        
        
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " »Ì«‰«  «·„” √Ã—Ì‰  "
    LogTexte = " Open Window " & "  Customers Data "
   ' AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ErrTrap
    LogTextA = "  «·Œ—ÊÃ „‰  " & " »Ì«‰«  «·„” √Ã—Ì‰  "
    LogTexte = " Exit   Window " & "  Customers Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

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
    LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«   «·„” √Ã—Ì‰ " _
       & CHR(13) & " ﬂÊœ «·⁄„Ì·  " & DCPreFix & txtid.text _
       & CHR(13) & "«·«”„ ⁄—»Ì  " & XPTxtCusName _
       & CHR(13) & "   „”∆Ê· «·« ’«·   " & TxtResponsibleContact _
       & CHR(13) & " —ﬁ„ «·Â« ›     " & XPTxtPhone _
       & CHR(13) & " —ﬁ„ «·ÃÊ«·     " & XPTxtmobile _
       & CHR(13) & " —ﬁ„ «·›«ﬂ”     " & TxtFaxNumber _
       & CHR(13) & "  «·»—Ìœ «·«·ﬂ —Ê‰Ì       " & TxtE_mail _
       & CHR(13) & " «·œÊ·Â   " & DcboCountryID.text _
       & CHR(13) & " «·„Õ«›Ÿ…   " & DcboGovernmentID.text _
       & CHR(13) & "  «·„œÌ‰…  " & DcboCityID.text _
       & CHR(13) & "  «·⁄‰Ê«‰ »«· ›’Ì· " & TxtAddress _
       & CHR(13) & " „·«ÕŸ«   " & XPMTxtRemarks _
       & CHR(13) & " ‰Ê⁄ «·Œ’„ ··„»Ì⁄«    " & CboDiscountType.text _
       & CHR(13) & "   ﬁÌ„Â «·Œ’„  " & TxtDiscountValue _
       & CHR(13) & " ‰Ê⁄ «·Œ’„ ··„‘ —Ì«    " & CboDiscountTypePur.text _
       & CHR(13) & "   ﬁÌ„Â «·Œ’„  " & TxtDiscountValuePur _
       & CHR(13) & "  ‰Ê⁄ «·⁄„Ì·  " & DcCustomerType.text _
       & CHR(13) & " «·„‰œÊ»   " & DCEmP.text _
       & CHR(13) & " Õœ «·«∆ „«‰ „œÌ‰  " & TxtCreditLimit _
       & CHR(13) & " „œ… «·«∆ „«‰     " & TxtDepitInterval.text & "   " & dcDepitIntervalID.text _
       & CHR(13) & " Õœ «·«∆ „«‰ œ«∆‰   " & TxtCreditlimitCredit _
       & CHR(13) & " „œ… «·«∆ „«‰      " & TxtCreditInterval.text & "   " & dcCreditIntervalID.text _
                    
       LogTextA = LogTextA & CHR(13) & "«·„” «Ã—  ø       "

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
       & CHR(13) & "  Code  " & DCPreFix & txtid.text _
       & CHR(13) & "Name " & XPTxtCusNamee _
       & CHR(13) & " Contact Person" & TxtResponsibleContact _
       & CHR(13) & " Tel " & XPTxtPhone _
       & CHR(13) & "Mob " & XPTxtmobile _
       & CHR(13) & " Fax  " & TxtFaxNumber _
       & CHR(13) & "  Email   " & TxtE_mail _
       & CHR(13) & " Contry   " & DcboCountryID.text _
       & CHR(13) & " City   " & DcboGovernmentID.text _
       & CHR(13) & "  Town  " & DcboCityID.text _
       & CHR(13) & " Address " & TxtAddress _
       & CHR(13) & " Remarks  " & XPMTxtRemarks _
       & CHR(13) & " Sales Discount  type  " & CboDiscountType.text _
       & CHR(13) & " Discount Value " & TxtDiscountValue _
       & CHR(13) & " Purchase Discount type " & CboDiscountTypePur.text _
       & CHR(13) & "  Discount Value" & TxtDiscountValuePur _
       & CHR(13) & "  Cust. Type " & DcCustomerType.text _
       & CHR(13) & " Sales Person   " & DCEmP.text _
       & CHR(13) & "The limit for debit  " & TxtCreditLimit _
       & CHR(13) & " Period     " & TxtDepitInterval.text & "   " & dcDepitIntervalID.text _
       & CHR(13) & "The limit for Credit   " & TxtCreditlimitCredit _
       & CHR(13) & " Period " & TxtCreditInterval.text & "   " & dcCreditIntervalID.text _
                    
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

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()

' „ «Œ›«¡ Â–Â «·ﬂÊœ ÊÂ–Â «·‘«‘… „‰ «Ã· «·Ê«ÃÂ… «·ÃœÌœ…
' FPFRM.show
' FPFRM.FramerETERIVE.Visible = False
'FPFRM.txtID = val(Me.XPTxtCusID.text)





 'FrmToolsSerials.show
 ' FrmToolsSerials.LblPath = Me.Name
 ' FrmToolsSerials.LblID = val(Me.XPTxtCusID)
 ' FrmToolsSerials.cmdInit_Click
'' FrmToolsSerials.Width = 3405
End Sub

Public Sub ISButton2_Click()
'Dim FULL_path As String
'FULL_path = App.path & "\Images\FP"
'    If XPTxtCusID.Text <> "" Then
'        DBPix201.ImageClear
'
'        If Dir(FULL_path & "\" & XPTxtCusID.Text & ".JPG") <> "" Then
'            DBPix201.ImageLoadFile (FULL_path & "\" & Me.Name & "\" & XPTxtCusID.Text & ".JPG")
'        End If
'
'
'
''    End If

' „ «Œ›«¡ Â–Â «·ﬂÊœ ÊÂ–Â «·‘«‘… „‰ «Ã· «·Ê«ÃÂ… «·ÃœÌœ…
' FPFRM.show
' FPFRM.FramerETERIVE.Visible = True
'FPFRM.txtID = val(Me.XPTxtCusID.text)
'FPFRM.retimage val(FPFRM.txtID)


End Sub

 
 
  
Private Sub Label2_Click()
    Frame2.Visible = False
End Sub

Private Sub menue_Click(index As Integer)
showsforms index
End Sub

Private Sub OptType_Click(index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.text)
End Sub

Private Sub TxtCreditLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCreditLimit.text, 1)
End Sub

Private Sub TxtCreditlimitCredit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCreditlimitCredit.text, 1)
End Sub

Private Sub txtCustGID_Change()
    Dim Custcode As String
    Dim CustName As String
    Dim Msg As String
Dim reson As String
Dim locked As Boolean
    If Me.TxtModFlg.text = "N" Then
        If Len(txtCustGID.text) >= 10 Then
            If CheckCustomerID(txtCustGID, Custcode, CustName, locked, reson) = True Then
           If locked = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Â–« «·„” √Ã— „”Ã· „‰ ﬁ»·   "
                    Msg = Msg & CHR(13) & " ﬂÊœ «·„” √Ã—: " & Custcode
                    Msg = Msg & CHR(13) & " «”„ «·„” √Ã— : " & CustName
                Else
                    Msg = "This Customer Already Exist"
                    Msg = Msg & CHR(13) & " Customer Code  " & Custcode
                    Msg = Msg & CHR(13) & "Customer Name  " & CustName
                                                                 
                End If
            Else
                  If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Â–« «·„” √Ã—   ›Ì «·ﬁ«∆„Â «·”Êœ«¡   "
                    Msg = Msg & CHR(13) & " ﬂÊœ «·„” √Ã—: " & Custcode
                    Msg = Msg & CHR(13) & " «”„ «·„” √Ã— : " & CustName
                    Msg = Msg & CHR(13) & " «·”»»  : " & reson
                Else
                    Msg = "This Customer  Black list"
                    Msg = Msg & CHR(13) & " Customer Code  " & Custcode
                    Msg = Msg & CHR(13) & "Customer Name  " & CustName
                    Msg = Msg & CHR(13) & " Reason  : " & reson
                                                                 
                End If
            End If

                MsgBox Msg, vbCritical
                txtCustGID.text = ""
                                        
            End If
        End If
    End If

End Sub

Private Sub txtCustGID_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.txtCustGID.text, 0)
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            DboParentAccount.Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„” √Ã—Ì‰"
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
                Me.Caption = "»Ì«‰«  «·„” √Ã—Ì‰( ÃœÌœ )"
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
                Me.Caption = "»Ì«‰«  «·„” √Ã—Ì‰(  ⁄œÌ· )"
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
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOpenBalance.text, 0)
End Sub

Private Sub TxtSalaries1_Change()
ClcAll
End Sub

Private Sub TxtSalaries1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtSalaries1.text, 0)
TxtSalaries1C = 0
End Sub

Private Sub TxtSalaries1C_Change()
ClcAll
End Sub

Private Sub TxtSalaries1C_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtSalaries1C.text, 0)
TxtSalaries1 = 0
End Sub

Private Sub TxtSalaries2_Change()
ClcAll
End Sub

Private Sub TxtSalaries2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtSalaries2.text, 0)
TxtSalaries2C = 0
End Sub

Private Sub TxtSalaries2C_Change()
ClcAll
End Sub

Private Sub TxtSalaries2C_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtSalaries2C.text, 0)
TxtSalaries2 = 0
End Sub

Private Sub TxtSalary_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSalary.text, 0)
End Sub

Private Sub TxtshowPrice1_Change()
ClcAll
End Sub

Private Sub TxtshowPrice1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtshowPrice1.text, 0)
TxtshowPrice1C = 0
End Sub

Private Sub TxtshowPrice1C_Change()
ClcAll
End Sub

Private Sub TxtshowPrice1C_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtshowPrice1C.text, 0)
 TxtshowPrice1 = 0
 
End Sub

Private Sub TxtshowPrice2_Change()
ClcAll
End Sub

Private Sub TxtshowPrice2_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtshowPrice2.text, 0)
 TxtshowPrice2C = 0
End Sub

Private Sub TxtshowPrice2C_Change()
ClcAll
End Sub

Private Sub TxtshowPrice2C_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtshowPrice2C.text, 0)
TxtshowPrice2 = 0
End Sub

Private Sub TxtShowQty1_Change()
ClcAll
End Sub

Private Sub TxtShowQty1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtShowQty1.text, 0)
TxtShowQty1c = 0
End Sub

Private Sub TxtShowQty1c_Change()
ClcAll
End Sub

Private Sub TxtShowQty1c_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, TxtShowQty1c.text, 0)
 TxtShowQty1 = 0
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
        TxtVATNO.text = IIf(IsNull(rs("VATNO").value), "", rs("VATNO").value)
        
        dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
Me.DtRecord.value = IIf(IsNull(rs("RecordDate")), Date, rs("RecordDate"))
        DCPreFix.text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
        Me.txtid.text = IIf(IsNull(rs("code").value), "", rs("code").value)
        txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)
        XPTxtCusID.text = IIf(IsNull(rs("CusID")), "", val(rs("CusID")))
        XPTxtCusName.text = IIf(IsNull(rs("CusName")), "", Trim(rs("CusName")))
        XPTxtCusNamee.text = IIf(IsNull(rs("CusNamee")), "", Trim(rs("CusNamee")))
        c1.text = IIf(IsNull(rs("c1")), "", Trim(rs("c1")))
        c2.text = IIf(IsNull(rs("c2")), "", Trim(rs("c2")))
        ''/////////////salah
    Me.TxtMap.text = IIf(IsNull(rs("Map").value), "", rs("Map").value)
    Me.txtJob.text = IIf(IsNull(rs("JobName").value), "", rs("JobName").value)
    Me.TxtEntry.text = IIf(IsNull(rs("Entry").value), "", rs("Entry").value)
    '''///////////////////
        Me.TxtResponsibleContact.text = IIf(IsNull(rs("ResponsibleContact").value), "", rs("ResponsibleContact").value)
        XPTxtPhone.text = IIf(IsNull(rs("Cus_Phone")), "", Trim(rs("Cus_Phone")))
        txtCustGID.text = IIf(IsNull(rs("CustGID")), "", (rs("CustGID")))
         BrithDate.value = IIf(IsNull(rs("BrithDate")), Date, Trim(rs("BrithDate")))
         BrithDateH.value = IIf(IsNull(rs("BrithDateH")), ToHijriDate(BrithDate.value), Trim(rs("BrithDateH")))
         TxtRecordNo.text = IIf(IsNull(rs("RecordNo")), "", rs("RecordNo"))
    ''///
    Me.TxtBox.text = IIf(IsNull(rs("Boxmil")), "", Trim(rs("Boxmil")))
        Me.TxtZib.text = IIf(IsNull(rs("ZipCode")), "", (rs("ZipCode")))
        DcbDigCustomer.ListIndex = IIf(IsNull(rs("TypeCustomer")), -1, (rs("TypeCustomer")))
            If IsNull(rs("PassWord").value) Then
    chkSendMessage.value = vbUnchecked
    Else
    
    chkSendMessage.value = IIf(rs("SendMessage") = 1, 1, 0)
    
    End If
        txtNoOFDigitUser(2).text = IIf(IsNull(rs("StreetName").value), "", rs("StreetName").value)
txtNoOFDigitUser(4).text = IIf(IsNull(rs("BuildingNumber").value), "", rs("BuildingNumber").value)
'txtNoOFDigitUser(9).Text = IIf(IsNull(rs("CitySubdivisionName").value), "", rs("CitySubdivisionName").value)
'txtNoOFDigitUser(6).Text = IIf(IsNull(rs("CityName").value), "", rs("CityName").value)
'txtNoOFDigitUser(7).Text = IIf(IsNull(rs("PostalZone").value), "", rs("PostalZone").value)
txtNoOFDigitUser(10).text = IIf(IsNull(rs("IdentificationCode").value), "", rs("IdentificationCode").value)
txtNoOFDigitUser(5).text = IIf(IsNull(rs("PlotIdentification").value), "", rs("PlotIdentification").value)
txtNoOFDigitUser(3).text = IIf(IsNull(rs("AdditionalStreetName").value), "", rs("AdditionalStreetName").value)
txtNoOFDigitUser(8).text = IIf(IsNull(rs("CountrySubentity").value), "", rs("CountrySubentity").value)

txtNoOFDigitUser(0).text = IIf(IsNull(rs("Id700").value), "", rs("Id700").value)
 
 
    '''///////////////////
'     TxtBankCode.text = IIf(IsNull(rs("BankCode").value), "", rs("BankCode").value)
'     TxtBankIBAN.text = IIf(IsNull(rs("BankIBAN").value), "", rs("BankIBAN").value)

    
        ''//
        XPTxtmobile.text = IIf(IsNull(rs("Cus_mobile")), "", Trim(rs("Cus_mobile")))
        XPMTxtRemarks.text = IIf(IsNull(rs("Remark")), "", Trim(rs("Remark")))
        XPMTxtRemarks2.text = IIf(IsNull(rs("Remark2")), "", Trim(rs("Remark2")))
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
    
        TxtDepitInterval.text = IIf(IsNull(rs("DepitInterval")), 0, rs("DepitInterval"))
        TxtCreditInterval.text = IIf(IsNull(rs("CreditInterval")), 0, rs("CreditInterval"))
    
        dcDepitIntervalID.ListIndex = IIf(IsNull(rs("DepitIntervalID")), -1, rs("DepitIntervalID"))
        dcCreditIntervalID.ListIndex = IIf(IsNull(rs("CreditIntervalID")), -1, rs("CreditIntervalID"))
     
        TxtCreditLimit.text = IIf(IsNull(rs("CreditLimit").value), "0", rs("CreditLimit").value)
'gooooooooooooold
TxtShowQty1.text = IIf(IsNull(rs("ShowQty1").value), "0", rs("ShowQty1").value)
TxtshowPrice1.text = IIf(IsNull(rs("showPrice1").value), "0", rs("showPrice1").value)
TxtshowPrice2.text = IIf(IsNull(rs("showPrice2").value), "0", rs("showPrice2").value)
TxtSalaries1.text = IIf(IsNull(rs("Salaries1").value), "0", rs("Salaries1").value)
TxtSalaries2.text = IIf(IsNull(rs("Salaries2").value), "0", rs("Salaries2").value)
TxtShowQty1c.text = IIf(IsNull(rs("ShowQty1c").value), "0", rs("ShowQty1c").value)
TxtshowPrice1C.text = IIf(IsNull(rs("showPrice1C").value), "0", rs("showPrice1C").value)
TxtshowPrice2C.text = IIf(IsNull(rs("showPrice2C").value), "0", rs("showPrice2C").value)
TxtSalaries1C.text = IIf(IsNull(rs("Salaries1C").value), "0", rs("Salaries1C").value)
 TxtSalaries2C.text = IIf(IsNull(rs("Salaries2C").value), "0", rs("Salaries2C").value)

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
            Me.TxtOpenBalance.text = IIf(IsNull(rs("OpenBalance")), "", Trim(rs("OpenBalance")))

            If rs("OpenBalanceType").value = 0 Then
                OptType(0).value = True
                OptType_Click 0
            ElseIf rs("OpenBalanceType").value = 1 Then
                OptType(1).value = True
                OptType_Click 1
            End If
        
        Else
            Me.TxtOpenBalance.text = 0
            Me.OptType(2).value = True
            OptType_Click 2
        End If
       
       
       
       
       
       
        If Not IsNull(rs("OpenBalanceType1").value) Then
        Me.TxtOpenBalance1.text = IIf(IsNull(rs("OpenBalance1")), "", Trim(rs("OpenBalance1")))

                    If rs("OpenBalanceType1").value = 0 Then
                        OptType1(0).value = True
                        OptType1_Click 0
                    ElseIf rs("OpenBalanceType1").value = 1 Then
                        OptType1(1).value = True
                        OptType1_Click 1
                    End If
    
    Else
        Me.TxtOpenBalance1.text = 0
        Me.OptType1(2).value = True
        OptType1_Click 2
    End If

    If Not IsNull(rs("OpenBalanceType2").value) Then
        Me.TxtOpenBalance2.text = IIf(IsNull(rs("OpenBalance2")), "", Trim(rs("OpenBalance2")))

        If rs("OpenBalanceType2").value = 0 Then
            OptType2(0).value = True
            OptType2_Click 0
        ElseIf rs("OpenBalanceType2").value = 1 Then
            OptType2(1).value = True
            OptType2_Click 1
        End If
    
    Else
        Me.TxtOpenBalance2.text = 0
        Me.OptType2(2).value = True
        OptType2_Click 2
    End If


        Me.TxtCreditlimitCredit.text = IIf(IsNull(rs("CreditlimitCredit").value), "0", rs("CreditlimitCredit").value)
        Me.TxtFaxNumber.text = IIf(IsNull(rs("FaxNumber").value), "", rs("FaxNumber").value)
        Me.TxtE_mail.text = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
          SngCusBegainAccount = GetCustomerAccount(val(XPTxtCusID.text), True)
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
            Me.TxtDiscountValue.text = 0
        ElseIf rs("Trans_DiscountType").value = 0 Then
            Me.CboDiscountType.ListIndex = 0
            Me.TxtDiscountValue.text = 0
        ElseIf rs("Trans_DiscountType").value = 1 Then
            Me.CboDiscountType.ListIndex = 1
            Me.TxtDiscountValue.text = IIf(IsNull(rs("Trans_Discount").value), "", rs("Trans_Discount").value)
        ElseIf rs("Trans_DiscountType").value = 2 Then
            Me.CboDiscountType.ListIndex = 2
            Me.TxtDiscountValue.text = IIf(IsNull(rs("Trans_Discount").value), "", rs("Trans_Discount").value)
        End If
    
        If IsNull(rs("Trans_DiscountTypePur").value) Then
            Me.CboDiscountTypePur.ListIndex = 0
            Me.TxtDiscountValuePur.text = 0
        ElseIf rs("Trans_DiscountTypePur").value = 0 Then
            Me.CboDiscountTypePur.ListIndex = 0
            Me.TxtDiscountValuePur.text = 0
        ElseIf rs("Trans_DiscountTypePur").value = 1 Then
            Me.CboDiscountTypePur.ListIndex = 1
            Me.TxtDiscountValuePur.text = IIf(IsNull(rs("Trans_DiscountPur").value), "", rs("Trans_DiscountPur").value)
        ElseIf rs("Trans_DiscountTypePur").value = 2 Then
            Me.CboDiscountTypePur.ListIndex = 2
            Me.TxtDiscountValuePur.text = IIf(IsNull(rs("Trans_DiscountPur").value), "", rs("Trans_DiscountPur").value)
        End If
    
        Me.DcboCountryID.BoundText = IIf(IsNull(rs("CountryID")), "", rs("CountryID"))
        Me.DcboCountryID2.BoundText = IIf(IsNull(rs("CountryID2")), "", rs("CountryID2"))
    
        Me.DcboGovernmentID.BoundText = IIf(IsNull(rs("GovernmentID")), "", rs("GovernmentID"))
        Me.DcboCityID.BoundText = IIf(IsNull(rs("CityID")), "", rs("CityID"))
        Me.DCEmP.BoundText = IIf(IsNull(rs("EmpId")), "", rs("EmpId"))
        Me.DcCustomerType.BoundText = IIf(IsNull(rs("CustomerTypeID")), "", rs("CustomerTypeID"))
      
        Me.TxtAddress.text = IIf(IsNull(rs("Address")), "", Trim(rs("Address")))
        '19082013
        Txt_DateExpLincH.value = IIf(IsNull(rs("ExpireDateH").value), ToHijriDate(Date), rs("ExpireDateH").value)

        Me.TxtCompany.text = IIf(IsNull(rs("Company")), "", Trim(rs("Company")))
        Me.TXTJobTitle.text = IIf(IsNull(rs("JobTitle")), "", Trim(rs("JobTitle")))
        Me.TxtSalary.text = IIf(IsNull(rs("Salary")), 0, Trim(rs("Salary")))
        Me.TxtJobAddress.text = IIf(IsNull(rs("JobAddress")), "", Trim(rs("JobAddress")))
        Me.TxtJobTel.text = IIf(IsNull(rs("JobTel")), "", Trim(rs("JobTel")))
        Me.TxtJobTelConvert.text = IIf(IsNull(rs("JobTelConvert")), "", Trim(rs("JobTelConvert")))
        Me.TxtHomeTel.text = IIf(IsNull(rs("HomeTel")), "", Trim(rs("HomeTel")))
        Me.TxtMobile1.text = IIf(IsNull(rs("Mobile1")), "", Trim(rs("Mobile1")))
        Me.TxtMobile2.text = IIf(IsNull(rs("Mobile2")), "", Trim(rs("Mobile2")))
    
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

    If XPTxtCusID.text <> "" Then

        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„Ì·   " & CHR(13)
        Msg = Msg + (XPTxtCusName.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

        If IntRes = vbYes Then
            If Not rs.RecordCount < 1 Then
                DeleteOpeningBalance
                Cn.BeginTrans
                BegainTrans = True
          
                ' StrSQL = "DELETE From NOTES Where NoteType=101 AND CusID=" & Val(Me.XPTxtCusID.text)
                ' Cn.Execute StrSQL, , adExecuteNoRecords
            
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
             
                '   update_account_opening_balance get_account_code_branch(19, my_branch)
               
                Dim StrAccountCode As String
                Dim StrAccountCode1 As String
                Dim StrAccountCode2 As String
                Dim InsuranceAccount As String
                Dim ParentAccount As String
                
StrAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
If SystemOptions.CustomerhavethreeAccounts = True Then
                'StrAccountCode1 = rs("Account_Code1").value
                'StrAccountCode2 = rs("Account_Code2").value
                StrAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                        StrAccountCode1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
                    StrAccountCode2 = IIf(IsNull(rs("Account_Code2").value), "", rs("Account_Code2").value)
                    InsuranceAccount = IIf(IsNull(rs("InsuranceAccount").value), "", rs("InsuranceAccount").value)
                    
 End If
 
                StrSQL = "Delete From Accounts Where Account_Code='" & StrAccountCode & "'"
                     
                     
                     If SystemOptions.CreateInsuranceAccountForCustomers = True Then
                                         If Not IsNull(rs("InsuranceAccount").value) Then
                   StrSQL = StrSQL & " or   Account_Code='" & rs("InsuranceAccount").value & "'"
                   End If
                     End If
                     
                     
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
                MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·⁄„Ì· "
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate

    If BegainTrans = True Then
        Cn.RollbackTrans
        BegainTrans = False
    End If

    'End If
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "CusID='" & val(XPTxtCusID.text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub OptType1_Click(index As Integer)
    Me.TxtOpenBalance1.Enabled = Not OptType1(2).value
    Me.TxtOpenBalance1.text = IIf(OptType1(2).value = True, 0, Me.TxtOpenBalance1.text)

End Sub

Private Sub OptType2_Click(index As Integer)
    Me.TxtOpenBalance2.Enabled = Not OptType2(2).value
    Me.TxtOpenBalance2.text = IIf(OptType2(2).value = True, 0, Me.TxtOpenBalance2.text)

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

   If val(Txtbalanced.text) > 0 Then
       OptType(0).value = True
       TxtOpenBalance.text = val(Txtbalanced.text)
       
    ElseIf val(TxtbalancedC.text) > 0 Then
       OptType(1).value = True
       TxtOpenBalance.text = val(TxtbalancedC.text)
           
           
       Else
       OptType(2).value = True
         TxtOpenBalance.text = 0
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
 

    If Me.TxtModFlg.text <> "R" Then
        If XPTxtCusName.text = "" Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "ÌÃ» «œŒ«· «”„ «·„” √Ã—"
                Else
                Msg = "Enter Renter name"
                End If
       
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtCusName.SetFocus
            Exit Sub
        End If

        If Me.OptType(2).value = False Then
                    If val(Me.TxtOpenBalance.text) = 0 Then
                        
                        
                     If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·—’Ìœ «·Ã«—Ì...!!!"
                Else
                Msg = "Enter  Opening  Balance"
                End If

                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        
                                If TxtOpenBalance.Enabled = True Then
                                    TxtOpenBalance.SetFocus
                                End If
        
                        Exit Sub
                    End If
        End If




        If Me.OptType1(2).value = False Then
                    If val(Me.TxtOpenBalance1.text) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·—’Ìœ ··‘Ìﬂ«   Õ  «· Õ’Ì· «·„” √Ã—...!!!"
                Else
                Msg = "Enter  Opening  Balance for Checks"
                End If
                
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        
                                If TxtOpenBalance1.Enabled = True Then
                                    TxtOpenBalance1.SetFocus
                                End If
        
                        Exit Sub
                    End If
        End If
        
        
        
                If Me.OptType2(2).value = False Then
                    If val(Me.TxtOpenBalance2.text) = 0 Then
                        
                        
                        
                         If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·—’Ìœ ··œ›⁄«  «·„ﬁœ„… «·„” √Ã—Ì‰...!!!"
                Else
                Msg = "Enter  Opening  Balance for Adv Payments"
                End If
                
                
                        
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        
                                If TxtOpenBalance2.Enabled = True Then
                                    TxtOpenBalance2.SetFocus
                                End If
        
                        Exit Sub
                    End If
        End If
        
        
        
        
        If val(Me.TxtCreditLimit.text) > 0 Then

            'Â‰«ﬂ Õœ ≈∆ „«‰ ( „œÌ‰)ÊÌÃ» «· «ﬂœ „‰ «·—’Ìœ «·√›  «ÕÏ «·„œŒ·
            If Me.OptType(0).value = True Then

                '«·—’Ìœ «·√›  «ÕÏ „œÌ‰
                If val(Me.TxtOpenBalance.text) > val(Me.TxtCreditLimit.text) Then
                    
                                  If SystemOptions.UserInterface = ArabicInterface Then
                   
                    Msg = "≈‰ »Â ....!!!"
                    Msg = Msg & CHR(13) & "·ﬁœ Ê÷⁄  Õœ ≈∆ „«‰ („œÌ‰ ) «·„” √Ã— " & val(Me.TxtCreditLimit.text)
                    Msg = Msg & CHR(13) & "·ﬂ‰ﬂ Ê÷⁄  «·—’Ìœ «·≈›  «ÕÏ «·„” √Ã— „œÌ‰ »‹  " & val(Me.TxtOpenBalance.text)
                    Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·»Ì«‰«  «· Ï «œŒ· Â«...øøø"
               
                Else
                  
                    Msg = "Hint  ....!!!"
                    Msg = Msg & CHR(13) & "Credit  Is  " & val(Me.TxtCreditLimit.text)
                    Msg = Msg & CHR(13) & "Depit opening balance is   " & val(Me.TxtOpenBalance.text)
                    Msg = Msg & CHR(13) & "???????"
               
                End If
                    
                     
                    
                    IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title)

                    If IntRes = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        If val(Me.TxtCreditlimitCredit.text) > 0 Then

            'Â‰«ﬂ Õœ ≈∆ „«‰ ( „œÌ‰)ÊÌÃ» «· «ﬂœ „‰ «·—’Ìœ «·√›  «ÕÏ «·„œŒ·
            If Me.OptType(1).value = True Then

                '«·—’Ìœ «·√›  «ÕÏ œ«∆‰
                If val(Me.TxtOpenBalance.text) > val(Me.TxtCreditlimitCredit.text) Then
                    
                                If SystemOptions.UserInterface = ArabicInterface Then
                   
                   Msg = "≈‰ »Â ....!!!"
                    Msg = Msg & CHR(13) & "·ﬁœ Ê÷⁄  Õœ ≈∆ „«‰ (œ«∆‰ ) «·„” √Ã— " & val(Me.TxtCreditlimitCredit.text)
                    Msg = Msg & CHR(13) & "·ﬂ‰ﬂ Ê÷⁄  «·—’Ìœ «·≈›  «ÕÏ «·„” √Ã— œ«∆‰ »‹  " & val(Me.TxtOpenBalance.text)
                    Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·»Ì«‰«  «· Ï «œŒ· Â«...øøø"
                    IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title)
 
                Else
                  
                    Msg = "Hint  ....!!!"
                    Msg = Msg & CHR(13) & "Credit  Is  " & val(Me.TxtCreditLimit.text)
                    Msg = Msg & CHR(13) & "Credit opening balance is   " & val(Me.TxtOpenBalance.text)
                    Msg = Msg & CHR(13) & "???????"
               
                End If
                
               
                    If IntRes = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        If Me.CboDiscountType.ListIndex = -1 Or Me.CboDiscountType.ListIndex = 0 Then
            Me.TxtDiscountValue.text = 0
        ElseIf Me.CboDiscountType.ListIndex = 1 Then

            If val(Me.TxtDiscountValue.text) = 0 Then
             If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·Œ’„ «·Œ«’… «·„” √Ã—...!!!"
Else
Msg = "Enter Discount value "
End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValue.SetFocus
                Exit Sub
            End If

        ElseIf Me.CboDiscountType.ListIndex = 2 Then

            If val(Me.TxtDiscountValue.text) = 0 Then
                
                
                             If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "ÌÃ» ﬂ «»… ‰”»… «·Œ’„ «·Œ«’… «·„” √Ã—...!!!"
Else
Msg = "Enter Discount %  "
End If


                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValue.SetFocus
                Exit Sub
            ElseIf val(Me.TxtDiscountValue.text) > 100 Then
            
                                         If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ‰”»… «·Œ’„ «ﬂ»— „‰ 100 ...!!!"
Else
Msg = "  Discount % cant > 100  "
End If


                
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValue.SetFocus
                Exit Sub
            End If
        End If
    
        If Me.CboDiscountTypePur.ListIndex = -1 Or Me.CboDiscountTypePur.ListIndex = 0 Then
            Me.TxtDiscountValuePur.text = 0
        ElseIf Me.CboDiscountTypePur.ListIndex = 1 Then

            If val(Me.TxtDiscountValuePur.text) = 0 Then
                                                     If SystemOptions.UserInterface = ArabicInterface Then

               Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·Œ’„ «·Œ«’… «·„” √Ã— ›Ï ›Ê« Ì— «·‘—«¡...!!!"
Else
Msg = "  Enter Discount   value For purchase invoices  "
End If

                
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValuePur.SetFocus
                Exit Sub
            End If

        ElseIf Me.CboDiscountTypePur.ListIndex = 2 Then

            If val(Me.TxtDiscountValuePur.text) = 0 Then
                
                
          If SystemOptions.UserInterface = ArabicInterface Then

               Msg = "ÌÃ» ﬂ «»… ‰”»… «·Œ’„ «·Œ«’… «·„” √Ã— ›Ï ›Ê« Ì— «·‘—«¡..!!!"
Else
Msg = "  Enter Discount   %  For purchase invoices  "
End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValuePur.SetFocus
                Exit Sub
            ElseIf val(Me.TxtDiscountValuePur.text) > 100 Then
                 
                
      If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ‰”»… «·Œ’„ «ﬂ»— „‰ 100 ...!!!"
Else
Msg = "  Discount % cant > 100  "
End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtDiscountValuePur.SetFocus
                Exit Sub
            End If
        End If
    
    
    
        Select Case TxtModFlg.text

            Case "N"
                XPTxtCusID.text = CStr(new_id("TblCustemers", "CusID", "", True))
            
                StrSQL = "Select * From TblCustemers where Type=56 And CusName='" & Trim(XPTxtCusName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                   If SystemOptions.UserInterface = ArabicInterface Then

                    Msg = "ÌÊÃœ „” √Ã— „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                   Else
                    Msg = "this Customer Already Exist" & CHR(13)
                   End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTxtCusName.SetFocus
                    Exit Sub
                End If
                Set RsTemp = New ADODB.Recordset
                   StrSQL = "Select * From TblCustemers where Type=56 And CustGID='" & Trim(txtCustGID.text) & "'"
                   RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                   If SystemOptions.UserInterface = ArabicInterface Then

                    Msg = "ÌÊÃœ „” √Ã— „”Ã· „”»ﬁ« »Â–« »‰›” —ﬁ„ «·ÂÊÌ…/«·«ﬁ«„…" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                   Else
                    Msg = "this Customer Already Exist" & CHR(13)
                   End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTxtCusName.SetFocus
                    Exit Sub
                End If

            Case "E"
                StrSQL = "select * From TblCustemers where Type=56 And CusName='" & Trim(XPTxtCusName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText


                If RsTemp.RecordCount > 0 Then
                                If RsTemp("CusID").value <> val(XPTxtCusID.text) Then
                                     
                                                      If SystemOptions.UserInterface = ArabicInterface Then
                                
                                                 Msg = "ÌÊÃœ „” √Ã— „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                                                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                                                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                                                    Else
                                                     Msg = "this Customer Already Exist" & CHR(13)
                                                     
                                                    End If
            
                                     MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                    XPTxtCusName.SetFocus
                                    Exit Sub
                                End If
                End If
                Set RsTemp = New ADODB.Recordset
                  StrSQL = "Select * From TblCustemers where Type=56 And CustGID='" & Trim(txtCustGID.text) & "'"
                   RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                If RsTemp("CusID").value <> val(XPTxtCusID.text) Then
                   If SystemOptions.UserInterface = ArabicInterface Then

                    Msg = "ÌÊÃœ „” √Ã— „”Ã· „”»ﬁ« »Â–« »‰›” —ﬁ„ «·ÂÊÌ…/«·«ﬁ«„…" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & CHR(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
                   Else
                    Msg = "this Customer Already Exist" & CHR(13)
                   End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTxtCusName.SetFocus
                    Exit Sub
                End If
            End If
            
        End Select

        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then
            Dim Account_Code_dynamic As String

            Account_Code_dynamic = Me.DboParentAccount.BoundText
            rs.AddNew

            rs("CusID").value = val(XPTxtCusID.text)
           
        ElseIf Me.TxtModFlg.text = "E" Then
           Account_Code_dynamic = Me.DboParentAccount.BoundText
            '  StrSQL = "DELETE From NOTES Where NoteType=101 AND CusID=" & Val(Me.XPTxtCusID.text)
            '   Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
        End If
        
    
            If val(TxtOpenBalance.text) <> 0 Or val(TxtOpenBalance1.text) <> 0 Or val(TxtOpenBalance2.text) <> 0 Then
                txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
                rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
            Else
                rs("opening_balance_voucher_id").value = Null
            End If
            
       
       
           
       '     If val(TxtOpenBalance.text) <> 0 Or val(TxtOpenBalance1.text) <> 0 Or val(TxtOpenBalance2.text) <> 0 Then
       '         txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
       '         rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
       '     Else
       '         rs("opening_balance_voucher_id").value = Null
       '     End If
       
       
 
    If chkSendMessage.value = vbChecked Then
              rs("SendMessage").value = 1
            Else
                rs("SendMessage").value = 0
            End If

         
        rs("code").value = txtid.text
        rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.text) & IIf(Trim(txtid.text) = "", Null, txtid.text)
        rs("prifix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)
 Me.TxtFullcode = IIf(DCPreFix.BoundText = "", Null, DCPreFix.text) & IIf(Trim(txtid.text) = "", Null, txtid.text)
     
        If chkCustomerandVendor.value = vbChecked Then
            rs("CustomerandVendor").value = 1

        Else
            rs("CustomerandVendor").value = 0
        End If
'
        rs("VATNO").value = TxtVATNO.text
        rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
        rs("RecordDate").value = Me.DtRecord.value
        rs("CusName").value = Trim(XPTxtCusName.text)
        If Trim(XPTxtCusNamee.text) = "" Then XPTxtCusNamee.text = Trim(XPTxtCusName.text)
        rs("CusNamee").value = IIf(Trim(XPTxtCusNamee.text) = "", Trim(XPTxtCusName.text), Trim(XPTxtCusNamee.text))
        rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
    
        rs("c1").value = Trim(c1.text)
        rs("c2").value = Trim(c2.text)
    
        rs("CustGID").value = IIf(txtCustGID.text = "", Null, val(txtCustGID.text))
       
        rs("Cus_Phone").value = IIf(XPTxtPhone.text = "", "", Trim(XPTxtPhone.text))
        rs("Cus_mobile").value = IIf(XPTxtmobile.text = "", "", Trim(XPTxtmobile.text))
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        rs("Remark2").value = IIf(XPMTxtRemarks2.text = "", "", Trim(XPMTxtRemarks2.text))
        rs("parent_account").value = IIf(Me.DboParentAccount.BoundText = "", Null, (Me.DboParentAccount.BoundText))
        rs("EmpId").value = IIf(Me.DCEmP.BoundText = "", Null, (Me.DCEmP.BoundText))
            rs("StreetName").value = txtNoOFDigitUser(2).text
        rs("BuildingNumber").value = txtNoOFDigitUser(4).text
         rs("CitySubdivisionName").value = DcboCityID.text
          rs("CityName").value = DcboGovernmentID.text
           rs("PostalZone").value = TxtZib.text
            rs("IdentificationCode").value = txtNoOFDigitUser(10).text
             rs("PlotIdentification").value = txtNoOFDigitUser(5).text
              rs("AdditionalStreetName").value = txtNoOFDigitUser(3).text
              rs("CountrySubentity").value = txtNoOFDigitUser(8).text
        If locked.value = vbChecked Then
            rs("locked").value = 1
        Else
            rs("locked").value = 0
        End If

        rs("CreditLimit").value = val(Me.TxtCreditLimit.text)
        rs("Type").value = 56
        
        rs("DepitInterval").value = val(TxtDepitInterval.text)
        rs("CreditInterval").value = val(TxtCreditInterval.text)
        
        rs("DepitIntervalID").value = val(dcDepitIntervalID.ListIndex)
        rs("CreditIntervalID").value = val(dcCreditIntervalID.ListIndex)
    'goooooooooooold
    
       rs("ShowQty1").value = val(Me.TxtShowQty1.text)
       rs("showPrice1").value = val(Me.TxtshowPrice1.text)
       rs("showPrice2").value = val(Me.TxtshowPrice2.text)
        rs("Salaries1").value = val(Me.TxtSalaries1.text)
        rs("Salaries2").value = val(Me.TxtSalaries2.text)
        
       rs("ShowQty1c").value = val(Me.TxtShowQty1c.text)
       rs("showPrice1c").value = val(Me.TxtshowPrice1C.text)
       rs("showPrice2c").value = val(Me.TxtshowPrice2C.text)
        rs("Salaries1c").value = val(Me.TxtSalaries1C.text)
        rs("Salaries2c").value = val(Me.TxtSalaries2C.text)
        rs("Id700").value = txtNoOFDigitUser(0).text
        
        rs("Totald").value = val(Me.txtTotald.text)
        rs("Totalc").value = val(Me.txtTotalc.text)
        rs("balanced").value = val(Me.Txtbalanced.text)
        rs("balancec").value = val(Me.TxtbalancedC.text)
        rs("RecordNo").value = TxtRecordNo.text
    
       
        
       
       'goooooooooooold
       rs("BrithDate").value = BrithDate.value
       rs("BrithDateH").value = BrithDateH.value
        If Me.OptType(2).value = True Then
            rs("OpenBalance").value = 0
            rs("OpenBalanceType").value = Null
        ElseIf Me.OptType(0).value = True Then
            rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
            rs("OpenBalanceType").value = 0
        ElseIf Me.OptType(1).value = True Then
            rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
            rs("OpenBalanceType").value = 1
        End If



        If Me.OptType1(2).value = True Then
            rs("OpenBalance1").value = 0
            rs("OpenBalanceType1").value = Null
        ElseIf Me.OptType1(0).value = True Then
            rs("OpenBalance1").value = val(Me.TxtOpenBalance1.text)
            rs("OpenBalanceType1").value = 0
        ElseIf Me.OptType1(1).value = True Then
            rs("OpenBalance1").value = val(Me.TxtOpenBalance1.text)
            rs("OpenBalanceType1").value = 1
        End If
        
        
        If Me.OptType2(2).value = True Then
            rs("OpenBalance2").value = 0
            rs("OpenBalanceType2").value = Null
        ElseIf Me.OptType2(0).value = True Then
            rs("OpenBalance2").value = val(Me.TxtOpenBalance2.text)
            rs("OpenBalanceType2").value = 0
        ElseIf Me.OptType2(1).value = True Then
            rs("OpenBalance2").value = val(Me.TxtOpenBalance2.text)
            rs("OpenBalanceType2").value = 1
        End If
        
        
        
        rs("OpenBalanceDate").value = Me.Dtp.value
    
        
        rs("CreditlimitCredit").value = val(Me.TxtCreditlimitCredit.text)
        rs("FaxNumber").value = IIf(Trim$(Me.TxtFaxNumber.text) = "", Null, Trim$(Me.TxtFaxNumber.text))
        rs("E_mail").value = IIf(Trim$(Me.TxtE_mail.text) = "", Null, Trim$(Me.TxtE_mail.text))

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
            rs("Trans_Discount").value = val(Me.TxtDiscountValue.text)
        ElseIf Me.CboDiscountType.ListIndex = 2 Then
            rs("Trans_DiscountType").value = 2
            rs("Trans_Discount").value = val(Me.TxtDiscountValue.text)
        End If
    
        If Me.CboDiscountTypePur.ListIndex = -1 Or Me.CboDiscountTypePur.ListIndex = 0 Then
            rs("Trans_DiscountTypePur").value = 0
            rs("Trans_DiscountPur").value = 0
        ElseIf Me.CboDiscountTypePur.ListIndex = 1 Then
            rs("Trans_DiscountTypePur").value = 1
            rs("Trans_DiscountPur").value = val(Me.TxtDiscountValuePur.text)
        ElseIf Me.CboDiscountTypePur.ListIndex = 2 Then
            rs("Trans_DiscountTypePur").value = 2
            rs("Trans_DiscountPur").value = val(Me.TxtDiscountValuePur.text)
        End If
    
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            Dim ParentAccount As String
            
            If Me.TxtModFlg.text = "N" Then
        
       '         rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtCusName.text), True, False, Trim$(Me.XPTxtCusNamee.text))
                
          If SystemOptions.CustomerhavethreeAccounts = False Then
        
                rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtCusName.text), True, False, Trim$(Me.XPTxtCusNamee.text))

                          If SystemOptions.CreateInsuranceAccountForCustomers = True Then
                
                                                   rs("InsuranceAccount").value = ModAccounts.AddNewAccount(Account_Code_dynamic166, Trim$(Me.XPTxtCusName.text & "  -  √„Ì‰«  „” —œ… "), True, False, Trim$(Me.XPTxtCusNamee.text) & "    √„Ì‰«  ··€Ì— ")
                    Else
                    
                     rs("InsuranceAccount").value = Null
                        End If

          Else
                
                                        If SystemOptions.CustomerhavethreeAccounts = True Then
                                            ParentAccount = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtCusName.text, False, False, XPTxtCusNamee.text)
                                            rs("ParentAccount").value = ParentAccount
                                         
                                            rs("Account_Code").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.text), True, False, XPTxtCusNamee.text)
                                            rs("Account_Code1").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.text) & "   ‘Ìﬂ«    Õ  «· Õ’Ì· ", True, False, XPTxtCusNamee.text & "  Under Collection Cheque  ")
                                            rs("Account_Code2").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.text) & "   œ›⁄«  „ﬁœ„…   ", True, False, XPTxtCusNamee.text & " Advanced Payments")

                                                     If SystemOptions.CreateInsuranceAccountForCustomers = True Then
                    
                                                       rs("InsuranceAccount").value = ModAccounts.AddNewAccount(Account_Code_dynamic166, Trim$(Me.XPTxtCusName.text) & "  -  √„Ì‰«  „” —œ… ", True, False, Trim$(Me.XPTxtCusNamee.text) & "  -  √„Ì‰«  „” —œ… ")
                                                       Else
                                                         rs("InsuranceAccount").value = Null
                    
                                                       End If
                            

                                        Else
                                            rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtCusName.text), True, False, XPTxtCusNamee.text)
                                            rs("ParentAccount").value = Null
                                            
                                        End If
             
        End If
                
                
                
       ' If SystemOptions.CreateInsuranceAccountForCustomers = True Then
                                       
       '   rs("InsuranceAccount").value = ModAccounts.AddNewAccount(Account_Code_dynamic166, Trim$(Me.XPTxtCusName.Text) & "    √„Ì‰«  ··€Ì— ", True, False, Trim$(Me.XPTxtCusNamee.Text) & "   insurance   ")
            
       '     End If
                     
                     
                
                'Rs("Account_Code").value = ModAccounts.AddNewAccount("a1a2a3", Trim$(Me.XPTxtCusName.text), True, False)
            Else

                 '       If Not IsNull(rs("Account_Code").value) Then
                 '           ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtCusName.text, Me.XPTxtCusNamee.text, , , , , , , , , , , , , , , , , True
                 '       End If
                        
                              If SystemOptions.CreateInsuranceAccountForCustomers = True Then
                                               If Not IsNull(rs("InsuranceAccount").value) And Not (rs("InsuranceAccount").value) = "" Then 'edit
                                               ModAccounts.EditAccount rs("InsuranceAccount").value, Me.XPTxtCusName.text & "    √„Ì‰«  „” —œ… ", XPTxtCusNamee.text & "Insurance ", , , , , , , , , , , , , , , , , True
                                               Else
                                               rs("InsuranceAccount").value = ModAccounts.AddNewAccount(Account_Code_dynamic166, Trim$(Me.XPTxtCusName.text) & "    √„Ì‰«  „” —œ… ", True, False, Trim$(Me.XPTxtCusNamee.text) & "   insurance   ")
                                               End If
                                            
                                                              
                           Else
                                 rs("InsuranceAccount").value = Null
                    
                     End If
                                                       
                         If SystemOptions.CustomerhavethreeAccounts = False Then
                '    If Not IsNull(rs("Account_Code").value) And (rs("Account_Code").value) = "" Then
                                     If Not IsNull(rs("Account_Code").value) And Not (rs("Account_Code").value) = "" Then
                                            ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtCusName.text, XPTxtCusNamee.text, , , , , , , , , , , , , , , , , True
                                        Else
                                              rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.XPTxtCusName.text), True, False, Trim$(Me.XPTxtCusNamee.text))

                                       '      ModAccounts.AddNewAccount rs("Account_Code").value, Me.XPTxtCusName.text, XPTxtCusNamee.text, , , , , , , , , , , , , , , , , True
                                        End If
                                
                                                
           
            
                Else
          
                    If Not IsNull(rs("ParentAccount").value) And Not (rs("ParentAccount").value) = "" Then
                        ModAccounts.EditAccount rs("ParentAccount").value, Me.XPTxtCusName.text, Trim(XPTxtCusNamee.text), , , , , , , , , , , , , , , , , False
                        Else
                          
                                     ParentAccount = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtCusName.text, False, False, XPTxtCusNamee.text)
                                            rs("ParentAccount").value = ParentAccount
                                            
                                   
                    End If
            
                    If Not IsNull(rs("Account_Code").value) And Not (rs("Account_Code").value) = "" Then
                        ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtCusName.text, XPTxtCusNamee.text, , , , , , , , , , , , , , , , , True
                      Else
                          rs("Account_Code").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.text), True, False, XPTxtCusNamee.text)
                             
                    End If
            
                    If Not IsNull(rs("Account_Code1").value) And Not (rs("Account_Code1").value) = "" Then
                        ModAccounts.EditAccount rs("Account_Code1").value, Me.XPTxtCusName.text & "    ‘Ìﬂ«    Õ  «· Õ’Ì·  ", XPTxtCusNamee.text & "  Cheque Box ", , , , , , , , , , , , , , , , , True
                        Else
                                               rs("Account_Code1").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.text) & "   ‘Ìﬂ«    Õ  «· Õ’Ì· ", True, False, XPTxtCusNamee.text & "  Under Collection Cheque  ")
                                         

                    End If
          
                    If Not IsNull(rs("Account_Code2").value) And Not (rs("Account_Code2").value) = "" Then
                        ModAccounts.EditAccount rs("Account_Code2").value, Me.XPTxtCusName.text & "   œ›⁄«  „ﬁœ„…   ", XPTxtCusNamee.text & "  Cheque Box ", , , , , , , , , , , , , , , , , True
                        Else
                                               rs("Account_Code2").value = ModAccounts.AddNewAccount(ParentAccount, Trim$(Me.XPTxtCusName.text) & "   œ›⁄«  „ﬁœ„…   ", True, False, XPTxtCusNamee.text & " Advanced Payment  ")

                    End If
                    
                End If
        
        
                        
                        
                        
                        
                        
                        
                        
            End If
            
            
        End If

        rs("CustomerTypeID").value = IIf(val(Me.DcCustomerType.BoundText) = 0, Null, val(Me.DcCustomerType.BoundText))
        rs("CountryID").value = IIf(val(Me.DcboCountryID.BoundText) = 0, Null, val(Me.DcboCountryID.BoundText))
        rs("CountryID2").value = IIf(val(Me.DcboCountryID2.BoundText) = 0, Null, val(Me.DcboCountryID2.BoundText))
         rs("Boxmil").value = TxtBox.text
      rs("ZipCode").value = Me.TxtZib.text
        rs("TypeCustomer").value = val(DcbDigCustomer.ListIndex)
       rs("Map").value = Trim$(Me.TxtMap.text)
       rs("Entry").value = Trim$(Me.TxtEntry.text)
       rs("JobName").value = Trim$(Me.txtJob.text)
       
        rs("GovernmentID").value = IIf(val(Me.DcboGovernmentID.BoundText) = 0, Null, val(Me.DcboGovernmentID.BoundText))
        rs("CityID").value = IIf(val(Me.DcboCityID.BoundText) = 0, Null, val(Me.DcboCityID.BoundText))
        rs("ResponsibleContact").value = Trim$(Me.TxtResponsibleContact.text)
        rs("Address").value = Trim$(Me.TxtAddress.text)
        rs("Sex").value = Trim$(Me.CboSex.text)
        '19 08 2013
        rs("ExpireDateH").value = Txt_DateExpLincH.value
        rs("Company").value = Trim(TxtCompany.text)
        rs("JobTitle").value = Trim(TXTJobTitle.text)
        rs("Salary").value = val(TxtSalary.text)
        rs("JobAddress").value = Trim(TxtJobAddress.text)
        rs("JobTel").value = Trim(TxtJobTel.text)
        rs("JobTelConvert").value = Trim(TxtJobTelConvert.text)
        rs("HomeTel").value = Trim(TxtHomeTel.text)
        rs("Mobile1").value = Trim(TxtMobile1.text)
        rs("Mobile2").value = Trim(TxtMobile2.text)
      
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
            
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code").value, Round(Me.TxtOpenBalance.text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText), , val(TxtShowQty1) - val(TxtShowQty1c), val(TxtshowPrice1) - val(TxtshowPrice1C), val(TxtshowPrice2) - val(TxtshowPrice2C), val(TxtSalaries1) - val(TxtSalaries1C), val(TxtSalaries2) - val(TxtSalaries2C)) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance.text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText), , val(TxtShowQty1c) - val(TxtShowQty1), val(TxtshowPrice1C) - val(TxtshowPrice1), val(TxtshowPrice2C) - val(TxtshowPrice2), val(TxtSalaries1C) - val(TxtSalaries1), val(TxtSalaries2C) - val(TxtSalaries2)) = False Then
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
                
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance.text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText), , val(TxtShowQty1c) - val(TxtShowQty1), val(TxtshowPrice1C) - val(TxtshowPrice1), val(TxtshowPrice2C) - val(TxtshowPrice2), val(TxtSalaries1C) - val(TxtSalaries1), val(TxtSalaries2C) - val(TxtSalaries2)) = False Then
                        GoTo ErrTrap
                    End If

                    ' If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
                      Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    
                    '       GoTo ErrTrap
                    'End If
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code").value, Round(Me.TxtOpenBalance.text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText), , val(TxtShowQty1) - val(TxtShowQty1c), val(TxtshowPrice1) - val(TxtshowPrice1C), val(TxtshowPrice2) - val(TxtshowPrice2C), val(TxtSalaries1) - val(TxtSalaries1C), val(TxtSalaries2) - val(TxtSalaries2C)) = False Then
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
                        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                            GoTo ErrTrap
                        End If
                    End If
            
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code1").value, Round(Me.TxtOpenBalance1.text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance1.text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
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
                
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance1.text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    ' If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
                      Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    
                    '       GoTo ErrTrap
                    'End If
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code1").value, Round(Me.TxtOpenBalance1.text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
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
            
                    If ModAccounts.AddNewDev(LngDevID, 1, rs("Account_Code2").value, Round(Me.TxtOpenBalance2.text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, Round(Me.TxtOpenBalance2.text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
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
                
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, Round(Me.TxtOpenBalance2.text, SystemOptions.SysDefCurrencyForamt), 0, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If

                    ' If ModAccounts.AddNewDev(LngDevID, 1, "a2a1a1", _
                      Val(Me.TxtOpenBalance.text), 0, "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.XPTxtCusName.text), LngOpenID, , , , Me.Dtp.value) = False Then
                    
                    '       GoTo ErrTrap
                    'End If
                    If ModAccounts.AddNewDev(LngDevID, 2, rs("Account_Code2").value, Round(Me.TxtOpenBalance2.text, SystemOptions.SysDefCurrencyForamt), 1, StrDes & Trim(Me.XPTxtCusName.text) & "  " & Trim$(Me.XPTxtCusNamee.text), LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text), , , , val(dcBranch.BoundText)) = False Then
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

        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·⁄„Ì· " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
 If index = 1 Then
 RSContract.ReloadCombos
            RSContract.dcCustomer.BoundText = val(XPTxtCusID.text)
            End If
            
            
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Done do you want new customer"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox "    Saved  ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If

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
            If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
                    Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
                    Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
            Else
            Msg = "Error  In Entry Data"
            End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
       If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Error During Saving"
    End If
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

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
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    ÿ·» ’Ì«‰…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(0), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    ’—› ﬁÿ⁄ €Ì«—   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(2), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…     ›« Ê—… ’Ì«‰…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(3), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   ‰»ÌÂ«  «·⁄„·«¡     ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(4), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   ﬁ«—Ì— «·’Ì«‰Â    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(5), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   ‘«‘… ›Õ’ ﬂ„»ÌÊ —  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(6), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   «·‰»ÌÂ«  «·„› ÊÕ…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(7), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    «· ‰»ÌÂ«  «·⁄«„…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(8), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    «·⁄„Ê·«  «·„” Õﬁ…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(9), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
            With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    «·⁄„·«¡   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(10), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
           With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…     ﬁ«—Ì— «·⁄„Ê·«    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(10), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·⁄„Ì·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„Ì· «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„Ì·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„Ì·" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·⁄„·«¡", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, BolRtl
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New..." & Wrap & "Add New Customer Data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print the current customer data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit the current customer data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the current editing or Save the new customer data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo..." & Wrap & "Undo in the adding new record" & Wrap & "OR undo editing current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete...." & Wrap & "Delete the current customer data." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search" & Wrap & "Search for a customer..." & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next..." & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last..." & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Customers Data", 1, 15204351, -2147483630, BolRtl
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
MySQL = MySQL & " Where (dbo.TblCustemers.CusID =" & val(XPTxtCusID.text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repRenter.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repRenter.rpt"
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
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
chkSendMessage.Caption = "Allow SMS"
    chkCustomerandVendor.Caption = "Customer / Supplier"
    Label1(2).Caption = "Type"
    Label3.Caption = "Branch"
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

    LngCusID = val(XPTxtCusID.text)
    OpenScreen PopUpShowCustomerBalanceScreen, LngCusID, 0
End Sub

Private Sub LoadDataCombos(Optional BolExceptCountries As Boolean = False, _
                           Optional BolExceptGovern As Boolean = False, _
                           Optional BolExceptCities As Boolean = False)
    Set Dcombo = New ClsDataCombos
   ' Dcombo.GetCountriesNames Me.DcboCountryID2
  Dcombo.GETNationality Me.DcboCountryID2
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


Function CREATEADDRESS()
If SystemOptions.IsBluee = True Then
TxtAddress = txtNoOFDigitUser(4) & " " & txtNoOFDigitUser(2) & " " & DcboCityID.text & " " & DcboGovernmentID.text & " " & DcboCountryID.text & " " & "«·—„“ «·»—ÌœÌ" & TxtZib
End If
End Function
Private Sub txtNoOFDigitUser_KeyPress(index As Integer, KeyAscii As Integer)
If index = 2 Or index = 10 Then Exit Sub
   KeyAscii = KeyAscii_Num(KeyAscii, txtNoOFDigitUser(4).text, 0)
End Sub




Function checkEeinvoice() As Boolean
   If Not SystemOptions.ApplyEinvoice Then checkEeinvoice = True: Exit Function
  'If chkTaxExempt.value = Checked Then checkEeinvoice = True: Exit Function
  'If creditlocked.value = Checked Then checkEeinvoice = True: Exit Function
checkEeinvoice = False

If TxtRecordNo.text = "" Then

    
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "—ﬁ„ «·”Ã· «·“«„Ì", vbCritical
                Else
                MsgBox "enter CRN ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If

If Not SystemOptions.CustVatNoMandatory Then
    If (TxtVATNO.text = "" Or Len(TxtVATNO) < 15 Or mId(TxtVATNO, 15, 1) <> 3) And Trim(TxtRecordNo) = "" Then
          If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "  «·—ﬁ„ «·÷—ÌÌ 15 Œ«‰…  «·“«„Ì ÊÌ‰ ÂÌ »«·—ﬁ„ 3", vbCritical
                    Else
                    MsgBox "Vat No 15 digit ", vbCritical
           End If
            checkEeinvoice = False
            Exit Function
    End If
    
    
'     If (TxtVATNO.text = "" Or Len(TxtVATNO) < 15 Or mId(TxtVATNO, 15, 1) <> 3) And Trim(txtNoOFDigitUser(0)) = "" Then
'          If SystemOptions.UserInterface = ArabicInterface Then
'                        MsgBox "  «·—ﬁ„ «·÷—ÌÌ 15 Œ«‰…  «·“«„Ì ÊÌ‰ ÂÌ »«·—ﬁ„ 3", vbCritical
'                    Else
'                    MsgBox "Vat No 15 digit ", vbCritical
'           End If
'            checkEeinvoice = False
'            Exit Function
'    End If
End If


If txtNoOFDigitUser(4).text = "" Or Len(txtNoOFDigitUser(4)) < 4 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     —ﬁ„ «·„»‰Ì 4 Œ«‰…  «·“«„Ì", vbCritical
                Else
                MsgBox "bulding no 4 digit ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If TxtZib.text = "" Or Len(TxtZib) < 5 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     «·—„“ «·»—ÌœÌ   5 Œ«‰…  «·“«„Ì", vbCritical
                Else
                MsgBox "Zib no 5 digit ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If txtNoOFDigitUser(2).text = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "    «”„ «·‘«—⁄  «·“«„Ì", vbCritical
                Else
                MsgBox "enter street name ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If

If txtNoOFDigitUser(10).text = "" Or Len(txtNoOFDigitUser(10)) < 2 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "   ﬂÊœ «·œÊ·…  «·“«„Ì 2 Œ«‰…", vbCritical
                Else
                MsgBox "must enter country code Code ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If val(DcboCountryID.BoundText) = 0 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     «·œÊ·…  «·“«„Ì   ", vbCritical
                Else
                MsgBox "must enter country  ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If val(DcboGovernmentID.BoundText) = 0 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     «·„œÌ‰…  «·“«„Ì   ", vbCritical
                Else
                MsgBox "must enter city  ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If

If val(DcboCityID.BoundText) = 0 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     «·ÕÌ  «·“«„Ì   ", vbCritical
                Else
                MsgBox "must enter distict  ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If
checkEeinvoice = True


End Function

