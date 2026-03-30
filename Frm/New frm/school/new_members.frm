VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form new_members 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·» «· Õ«ﬁ"
   ClientHeight    =   8085
   ClientLeft      =   180
   ClientTop       =   780
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   11415
   Begin VB.CommandButton Command4 
      Caption         =   " ⁄—÷  ﬂ· «·»Ì«‰« "
      Height          =   375
      Left            =   9840
      TabIndex        =   119
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Õ›Ÿ"
      Height          =   615
      Left            =   8640
      TabIndex        =   2
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÃœÌœ"
      Height          =   615
      Left            =   9960
      TabIndex        =   1
      Top             =   7440
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "»Ì«‰«  «·ÿ«·» «·«”«”Ì"
      TabPicture(0)   =   "new_members.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(10)=   "Label13"
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(12)=   "Label15"
      Tab(0).Control(13)=   "DataCombo1"
      Tab(0).Control(14)=   "Text1"
      Tab(0).Control(15)=   "Text3"
      Tab(0).Control(16)=   "Text4"
      Tab(0).Control(17)=   "TXTDOB"
      Tab(0).Control(18)=   "Text6"
      Tab(0).Control(19)=   "Text7"
      Tab(0).Control(20)=   "Text11"
      Tab(0).Control(21)=   "Text9"
      Tab(0).Control(22)=   "Text2"
      Tab(0).Control(23)=   "Text8"
      Tab(0).Control(24)=   "Text10"
      Tab(0).Control(25)=   "Text5"
      Tab(0).Control(26)=   "Command3(0)"
      Tab(0).Control(27)=   "Calendar1"
      Tab(0).Control(28)=   "Command3(1)"
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "»Ì«‰«  «· «»⁄ 1"
      TabPicture(1)   =   "new_members.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label16"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label17"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label18"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label19"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label20"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label21"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label22"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label23"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label24"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label25"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label26"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text12"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Text13"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text14"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Text15"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Text16"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text17"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Text18"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Text19"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Text20"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Text21"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Text22"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Command3(2)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Command3(3)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "»Ì«‰«   «· «»⁄ 2"
      TabPicture(2)   =   "new_members.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label27"
      Tab(2).Control(1)=   "Label28"
      Tab(2).Control(2)=   "Label29"
      Tab(2).Control(3)=   "Label30"
      Tab(2).Control(4)=   "Label31"
      Tab(2).Control(5)=   "Label32"
      Tab(2).Control(6)=   "Label33"
      Tab(2).Control(7)=   "Label34"
      Tab(2).Control(8)=   "Label35"
      Tab(2).Control(9)=   "Label36"
      Tab(2).Control(10)=   "Label37"
      Tab(2).Control(11)=   "Text23"
      Tab(2).Control(12)=   "Text24"
      Tab(2).Control(13)=   "Text25"
      Tab(2).Control(14)=   "Text26"
      Tab(2).Control(15)=   "Text27"
      Tab(2).Control(16)=   "Text28"
      Tab(2).Control(17)=   "Text29"
      Tab(2).Control(18)=   "Text30"
      Tab(2).Control(19)=   "Text31"
      Tab(2).Control(20)=   "Text32"
      Tab(2).Control(21)=   "Text33"
      Tab(2).Control(22)=   "Command3(4)"
      Tab(2).Control(23)=   "Command3(5)"
      Tab(2).ControlCount=   24
      TabCaption(3)   =   "»Ì«‰«   «· «»⁄ 3"
      TabPicture(3)   =   "new_members.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label38"
      Tab(3).Control(1)=   "Label39"
      Tab(3).Control(2)=   "Label40"
      Tab(3).Control(3)=   "Label41"
      Tab(3).Control(4)=   "Label42"
      Tab(3).Control(5)=   "Label43"
      Tab(3).Control(6)=   "Label44"
      Tab(3).Control(7)=   "Label45"
      Tab(3).Control(8)=   "Label46"
      Tab(3).Control(9)=   "Label47"
      Tab(3).Control(10)=   "Label48"
      Tab(3).Control(11)=   "Text34"
      Tab(3).Control(12)=   "Text35"
      Tab(3).Control(13)=   "Text36"
      Tab(3).Control(14)=   "Text37"
      Tab(3).Control(15)=   "Text38"
      Tab(3).Control(16)=   "Text39"
      Tab(3).Control(17)=   "Text40"
      Tab(3).Control(18)=   "Text41"
      Tab(3).Control(19)=   "Text42"
      Tab(3).Control(20)=   "Text43"
      Tab(3).Control(21)=   "Text44"
      Tab(3).Control(22)=   "Command3(6)"
      Tab(3).Control(23)=   "Command3(7)"
      Tab(3).ControlCount=   24
      TabCaption(4)   =   "»Ì«‰«   «· «»⁄ 4"
      TabPicture(4)   =   "new_members.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label49"
      Tab(4).Control(1)=   "Label50"
      Tab(4).Control(2)=   "Label51"
      Tab(4).Control(3)=   "Label52"
      Tab(4).Control(4)=   "Label53"
      Tab(4).Control(5)=   "Label54"
      Tab(4).Control(6)=   "Label55"
      Tab(4).Control(7)=   "Label56"
      Tab(4).Control(8)=   "Label57"
      Tab(4).Control(9)=   "Label58"
      Tab(4).Control(10)=   "Label59"
      Tab(4).Control(11)=   "Text45"
      Tab(4).Control(12)=   "Text46"
      Tab(4).Control(13)=   "Text47"
      Tab(4).Control(14)=   "Text48"
      Tab(4).Control(15)=   "Text49"
      Tab(4).Control(16)=   "Text50"
      Tab(4).Control(17)=   "Text51"
      Tab(4).Control(18)=   "Text52"
      Tab(4).Control(19)=   "Text53"
      Tab(4).Control(20)=   "Text54"
      Tab(4).Control(21)=   "Text55"
      Tab(4).Control(22)=   "Command3(8)"
      Tab(4).Control(23)=   "Command3(9)"
      Tab(4).ControlCount=   24
      Begin VB.CommandButton Command3 
         Height          =   255
         Index           =   9
         Left            =   -69840
         Picture         =   "new_members.frx":008C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Index           =   8
         Left            =   -74640
         Picture         =   "new_members.frx":08EE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Index           =   7
         Left            =   -69600
         Picture         =   "new_members.frx":1150
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Index           =   6
         Left            =   -74640
         Picture         =   "new_members.frx":19B2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Height          =   252
         Index           =   5
         Left            =   -69960
         Picture         =   "new_members.frx":2214
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   2760
         Width           =   252
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Index           =   4
         Left            =   -74640
         Picture         =   "new_members.frx":2A76
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Index           =   3
         Left            =   5520
         Picture         =   "new_members.frx":32D8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   2400
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Index           =   2
         Left            =   120
         Picture         =   "new_members.frx":3B3A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Index           =   1
         Left            =   -69120
         Picture         =   "new_members.frx":439C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   3120
         Width           =   255
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2775
         Left            =   -70200
         TabIndex        =   121
         Top             =   6480
         Visible         =   0   'False
         Width           =   3255
         _Version        =   524288
         _ExtentX        =   5741
         _ExtentY        =   4895
         _StockProps     =   1
         BackColor       =   -2147483641
         Year            =   2009
         Month           =   6
         Day             =   4
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   65535
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   0
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Index           =   0
         Left            =   -74400
         Picture         =   "new_members.frx":4BFE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox Text55 
         Alignment       =   2  'Center
         DataField       =   "grandmother_order_number"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -69480
         TabIndex        =   107
         Text            =   " "
         Top             =   1080
         Width           =   3200
      End
      Begin VB.TextBox Text54 
         Alignment       =   2  'Center
         DataField       =   "grandmother_NAME"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74280
         TabIndex        =   106
         Top             =   1560
         Width           =   3200
      End
      Begin VB.TextBox Text53 
         Alignment       =   2  'Center
         DataField       =   "grandmother_DOB"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   405
         Left            =   -69480
         TabIndex        =   105
         Top             =   2640
         Width           =   3200
      End
      Begin VB.TextBox Text52 
         Alignment       =   2  'Center
         DataField       =   "grandmother_born_place"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74280
         TabIndex        =   104
         Top             =   2640
         Width           =   3200
      End
      Begin VB.TextBox Text51 
         Alignment       =   2  'Center
         DataField       =   "grandmother_job"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74280
         TabIndex        =   103
         Top             =   2040
         Width           =   3200
      End
      Begin VB.TextBox Text50 
         Alignment       =   2  'Center
         DataField       =   "grandmother_NATIONAL_id"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -69480
         TabIndex        =   102
         Top             =   1560
         Width           =   3200
      End
      Begin VB.TextBox Text49 
         Alignment       =   2  'Center
         DataField       =   "grandmother_address"
         DataSource      =   "Adodc1"
         Height          =   1935
         Left            =   -69480
         MultiLine       =   -1  'True
         TabIndex        =   101
         Top             =   3720
         Width           =   3200
      End
      Begin VB.TextBox Text48 
         Alignment       =   2  'Center
         DataField       =   "grandmother_order_date"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   405
         Left            =   -74280
         TabIndex        =   100
         Top             =   1080
         Width           =   3200
      End
      Begin VB.TextBox Text47 
         Alignment       =   2  'Center
         DataField       =   "grandmother_job_address"
         DataSource      =   "Adodc1"
         Height          =   2295
         Left            =   -74280
         MultiLine       =   -1  'True
         TabIndex        =   99
         Top             =   3360
         Width           =   3200
      End
      Begin VB.TextBox Text46 
         Alignment       =   2  'Center
         DataField       =   "grandmother_certificate"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -69480
         TabIndex        =   98
         Top             =   2160
         Width           =   3200
      End
      Begin VB.TextBox Text45 
         Alignment       =   2  'Center
         DataField       =   "grandmother_telephone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -69480
         TabIndex        =   97
         Text            =   " "
         Top             =   3240
         Width           =   3200
      End
      Begin VB.TextBox Text44 
         Alignment       =   2  'Center
         DataField       =   "grandfather_order_number"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -69240
         TabIndex        =   85
         Text            =   " "
         Top             =   1320
         Width           =   3200
      End
      Begin VB.TextBox Text43 
         Alignment       =   2  'Center
         DataField       =   "grandfather_NAME"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74280
         TabIndex        =   84
         Top             =   1800
         Width           =   3200
      End
      Begin VB.TextBox Text42 
         Alignment       =   2  'Center
         DataField       =   "grandfather_DOB"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   405
         Left            =   -69240
         TabIndex        =   83
         Top             =   2880
         Width           =   3200
      End
      Begin VB.TextBox Text41 
         Alignment       =   2  'Center
         DataField       =   "grandfather_born_place"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74280
         TabIndex        =   82
         Top             =   2880
         Width           =   3200
      End
      Begin VB.TextBox Text40 
         Alignment       =   2  'Center
         DataField       =   "grandfather_job"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74280
         TabIndex        =   81
         Top             =   2280
         Width           =   3200
      End
      Begin VB.TextBox Text39 
         Alignment       =   2  'Center
         DataField       =   "grandfather_NATIONAL_id"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -69240
         TabIndex        =   80
         Top             =   1800
         Width           =   3200
      End
      Begin VB.TextBox Text38 
         Alignment       =   2  'Center
         DataField       =   "grandfather_address"
         DataSource      =   "Adodc1"
         Height          =   1935
         Left            =   -69240
         TabIndex        =   79
         Top             =   3840
         Width           =   3200
      End
      Begin VB.TextBox Text37 
         Alignment       =   2  'Center
         DataField       =   "grandfather_order_date"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   405
         Left            =   -74280
         TabIndex        =   78
         Top             =   1200
         Width           =   3200
      End
      Begin VB.TextBox Text36 
         Alignment       =   2  'Center
         DataField       =   "grandfather_job_address"
         DataSource      =   "Adodc1"
         Height          =   2295
         Left            =   -74280
         TabIndex        =   77
         Top             =   3600
         Width           =   3200
      End
      Begin VB.TextBox Text35 
         Alignment       =   2  'Center
         DataField       =   "grandfather_certificate"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -69240
         TabIndex        =   76
         Top             =   2400
         Width           =   3200
      End
      Begin VB.TextBox Text34 
         Alignment       =   2  'Center
         DataField       =   "grandfather_telephone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -69240
         TabIndex        =   75
         Text            =   " "
         Top             =   3360
         Width           =   3200
      End
      Begin VB.TextBox Text33 
         Alignment       =   2  'Center
         DataField       =   "sister_order_number"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -69600
         TabIndex        =   63
         Text            =   " "
         Top             =   840
         Width           =   3200
      End
      Begin VB.TextBox Text32 
         Alignment       =   2  'Center
         DataField       =   "sister_NAME"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74400
         TabIndex        =   62
         Top             =   1200
         Width           =   3200
      End
      Begin VB.TextBox Text31 
         Alignment       =   2  'Center
         DataField       =   "sister_DOB"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   405
         Left            =   -69600
         TabIndex        =   61
         Top             =   2640
         Width           =   3200
      End
      Begin VB.TextBox Text30 
         Alignment       =   2  'Center
         DataField       =   "sister_born_place"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74400
         TabIndex        =   60
         Top             =   2520
         Width           =   3200
      End
      Begin VB.TextBox Text29 
         Alignment       =   2  'Center
         DataField       =   "sister_job"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74400
         TabIndex        =   59
         Top             =   1680
         Width           =   3200
      End
      Begin VB.TextBox Text28 
         Alignment       =   2  'Center
         DataField       =   "sister_NATIONAL_id"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -69600
         TabIndex        =   58
         Top             =   1320
         Width           =   3200
      End
      Begin VB.TextBox Text27 
         Alignment       =   2  'Center
         DataField       =   "sister_address"
         DataSource      =   "Adodc1"
         Height          =   1935
         Left            =   -69600
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   3720
         Width           =   3200
      End
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         DataField       =   "sister_order_date"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   405
         Left            =   -74400
         TabIndex        =   56
         Top             =   720
         Width           =   3200
      End
      Begin VB.TextBox Text25 
         Alignment       =   2  'Center
         DataField       =   "sister_job_address"
         DataSource      =   "Adodc1"
         Height          =   2295
         Left            =   -74520
         MultiLine       =   -1  'True
         TabIndex        =   55
         Top             =   3360
         Width           =   3200
      End
      Begin VB.TextBox Text24 
         Alignment       =   2  'Center
         DataField       =   "sister_certificate"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -69600
         TabIndex        =   54
         Top             =   1920
         Width           =   3200
      End
      Begin VB.TextBox Text23 
         Alignment       =   2  'Center
         DataField       =   "sister_telephone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -69600
         TabIndex        =   53
         Text            =   " "
         Top             =   3240
         Width           =   3200
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         DataField       =   "wife_order_number"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   41
         Text            =   " "
         Top             =   840
         Width           =   3200
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         DataField       =   "wife_NAME"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   360
         TabIndex        =   40
         Top             =   2040
         Width           =   4000
      End
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         DataField       =   "wife_DOB"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   405
         Left            =   5880
         TabIndex        =   39
         Top             =   2280
         Width           =   3200
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         DataField       =   "wife_born_place"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   360
         TabIndex        =   38
         Top             =   3000
         Width           =   4000
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         DataField       =   "wife_job"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   360
         TabIndex        =   37
         Top             =   2520
         Width           =   4000
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         DataField       =   "WIFE_NATIONAL_id"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   36
         Top             =   1320
         Width           =   3200
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         DataField       =   "wife_address"
         DataSource      =   "Adodc1"
         Height          =   1935
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   3720
         Width           =   3195
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         DataField       =   "wife_order_date"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   405
         Left            =   360
         TabIndex        =   34
         Top             =   1560
         Width           =   4000
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         DataField       =   "wife_job_address"
         DataSource      =   "Adodc1"
         Height          =   1935
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   3600
         Width           =   4000
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         DataField       =   "wife_certificate"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   32
         Top             =   1800
         Width           =   3200
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         DataField       =   "wife_telephone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   31
         Text            =   " "
         Top             =   2880
         Width           =   3200
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_telephone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -68760
         TabIndex        =   29
         Text            =   " "
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_certificate"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -68760
         TabIndex        =   28
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_job_address"
         DataSource      =   "Adodc1"
         Height          =   2175
         Left            =   -74040
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   3960
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         DataField       =   "original_order_date"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74040
         TabIndex        =   26
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_address"
         DataSource      =   "Adodc1"
         Height          =   1935
         Left            =   -68760
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   4080
         Width           =   2895
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         DataField       =   "original_NATIONAL_id"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -68760
         TabIndex        =   24
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_job"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74040
         TabIndex        =   21
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_born_place"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74040
         TabIndex        =   20
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox TXTDOB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "MEMBER_DOB"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   405
         Left            =   -68760
         TabIndex        =   19
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_NAME"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   -74040
         TabIndex        =   18
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         DataField       =   "original_order_number"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -68760
         TabIndex        =   17
         Text            =   " "
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "index"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68760
         TabIndex        =   16
         Text            =   " "
         Top             =   1080
         Width           =   2775
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "new_members.frx":5460
         DataField       =   "member_type"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   -74040
         TabIndex        =   6
         Top             =   1560
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "MEMBER_NAME"
         BoundColumn     =   "MEMBER_ID"
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label59 
         Caption         =   "—ﬁ„ «·ÿ·» ··Ãœ…"
         Height          =   375
         Left            =   -66120
         TabIndex        =   118
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label58 
         Caption         =   "√”„ «·⁄÷Ê"
         Height          =   375
         Left            =   -70680
         TabIndex        =   117
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label57 
         Caption         =   "«·ÊŸÌ›…"
         Height          =   495
         Left            =   -70800
         TabIndex        =   116
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label56 
         Caption         =   "⁄‰Ê«‰ «·ÊŸÌ›…"
         Height          =   495
         Left            =   -71040
         TabIndex        =   115
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label55 
         Caption         =   " «—ÌŒ «·ÿ·»"
         Height          =   495
         Left            =   -70800
         TabIndex        =   114
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label54 
         Caption         =   " «—ÌŒ «·„Ì·«œ"
         Height          =   255
         Left            =   -66120
         TabIndex        =   113
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label53 
         Caption         =   "ÃÂ… «·„Ì·«œ"
         Height          =   375
         Left            =   -70800
         TabIndex        =   112
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label52 
         Caption         =   "⁄‰Ê«‰ «·”ﬂ‰"
         Height          =   375
         Left            =   -66120
         TabIndex        =   111
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label51 
         Caption         =   "«·„ƒÂ·"
         Height          =   375
         Left            =   -66000
         TabIndex        =   110
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label50 
         Caption         =   "«·—ﬁ„ «·ﬁÊ„Ì"
         Height          =   255
         Left            =   -66000
         TabIndex        =   109
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label49 
         Height          =   375
         Left            =   -66240
         TabIndex        =   108
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label48 
         Caption         =   "—ﬁ„ «·ÿ·» ··Ãœ"
         Height          =   375
         Left            =   -66000
         TabIndex        =   96
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label47 
         Caption         =   "√”„ «·⁄÷Ê"
         Height          =   375
         Left            =   -70920
         TabIndex        =   95
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label46 
         Caption         =   "«·ÊŸÌ›…"
         Height          =   495
         Left            =   -70800
         TabIndex        =   94
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label45 
         Caption         =   "⁄‰Ê«‰ «·ÊŸÌ›…"
         Height          =   495
         Left            =   -70920
         TabIndex        =   93
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label44 
         Caption         =   " «—ÌŒ «·ÿ·»"
         Height          =   495
         Left            =   -71040
         TabIndex        =   92
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label43 
         Caption         =   " «—ÌŒ «·„Ì·«œ"
         Height          =   255
         Left            =   -65880
         TabIndex        =   91
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label42 
         Caption         =   "ÃÂ… «·„Ì·«œ"
         Height          =   375
         Left            =   -70800
         TabIndex        =   90
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label41 
         Caption         =   "⁄‰Ê«‰ «·”ﬂ‰"
         Height          =   375
         Left            =   -65880
         TabIndex        =   89
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "«·„ƒÂ·"
         Height          =   375
         Left            =   -65760
         TabIndex        =   88
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label39 
         Caption         =   "«·—ﬁ„ «·ﬁÊ„Ì"
         Height          =   255
         Left            =   -65760
         TabIndex        =   87
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   " ·Ì›Ê‰ «·⁄÷Ê"
         Height          =   375
         Left            =   -65880
         TabIndex        =   86
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label37 
         Caption         =   "—ﬁ„ «·ÿ·»  ··«Œ "
         Height          =   375
         Left            =   -66360
         TabIndex        =   74
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label36 
         Caption         =   "√”„ «·⁄÷Ê"
         Height          =   375
         Left            =   -70920
         TabIndex        =   73
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label35 
         Caption         =   "«·ÊŸÌ›…"
         Height          =   495
         Left            =   -70920
         TabIndex        =   72
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "⁄‰Ê«‰ «·ÊŸÌ›…"
         Height          =   495
         Left            =   -70920
         TabIndex        =   71
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label33 
         Caption         =   " «—ÌŒ «·ÿ·»"
         Height          =   495
         Left            =   -71040
         TabIndex        =   70
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   " «—ÌŒ «·„Ì·«œ"
         Height          =   255
         Left            =   -66000
         TabIndex        =   69
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "ÃÂ… «·„Ì·«œ"
         Height          =   375
         Left            =   -71040
         TabIndex        =   68
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label30 
         Caption         =   "⁄‰Ê«‰ «·”ﬂ‰"
         Height          =   375
         Left            =   -65880
         TabIndex        =   67
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "«·„ƒÂ·"
         Height          =   375
         Left            =   -66240
         TabIndex        =   66
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "«·—ﬁ„ «·ﬁÊ„Ì"
         Height          =   255
         Left            =   -66240
         TabIndex        =   65
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   " ·Ì›Ê‰ «·⁄÷Ê"
         Height          =   375
         Left            =   -65880
         TabIndex        =   64
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "—ﬁ„ «·ÿ·»   "
         Height          =   375
         Left            =   9360
         TabIndex        =   52
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "√”„ «·ÿ«·»"
         Height          =   375
         Left            =   4560
         TabIndex        =   51
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "ÊŸÌ›… «·«»"
         Height          =   495
         Left            =   4440
         TabIndex        =   50
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "⁄‰Ê«‰ «·ÊŸÌ›…"
         Height          =   495
         Left            =   4440
         TabIndex        =   49
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   " «—ÌŒ «·ÿ·»"
         Height          =   375
         Left            =   4560
         TabIndex        =   48
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   " «—ÌŒ «·„Ì·«œ"
         Height          =   255
         Left            =   9240
         TabIndex        =   47
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "ÃÂ… «·„Ì·«œ"
         Height          =   375
         Left            =   4440
         TabIndex        =   46
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label19 
         Caption         =   "⁄‰Ê«‰ «·”ﬂ‰"
         Height          =   375
         Left            =   9240
         TabIndex        =   45
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "«Œ— ‘Â«œ… ··ÿ«·»"
         Height          =   375
         Left            =   9360
         TabIndex        =   44
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "—ﬁ„ ÂÊÌ… Ê·Ì «·«„—"
         Height          =   255
         Left            =   9240
         TabIndex        =   43
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   " ·Ì›Ê‰  Ê·Ì «·«„—"
         Height          =   375
         Left            =   9240
         TabIndex        =   42
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   " ·Ì›Ê‰ Ê·Ì «·«„—"
         Height          =   375
         Left            =   -65640
         TabIndex        =   30
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "—ﬁ„ ÂÊÌ… Ê·Ì «·«„—"
         Height          =   255
         Left            =   -65640
         TabIndex        =   23
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "«Œ— ‘Â«œÂ ··ÿ«·»"
         Height          =   375
         Left            =   -65640
         TabIndex        =   22
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "⁄‰Ê«‰ «·”ﬂ‰"
         Height          =   375
         Left            =   -65520
         TabIndex        =   15
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "ÃÂ… «·„Ì·«œ"
         Height          =   375
         Left            =   -70440
         TabIndex        =   14
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   " «—ÌŒ «·„Ì·«œ"
         Height          =   255
         Left            =   -65640
         TabIndex        =   13
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   " «—ÌŒ «·ÿ·»"
         Height          =   255
         Left            =   -70440
         TabIndex        =   12
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "⁄‰Ê«‰ «·ÊŸÌ›…"
         Height          =   495
         Left            =   -70320
         TabIndex        =   11
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "ÊŸÌ›… Ê·Ì «·«„—"
         Height          =   495
         Left            =   -70440
         TabIndex        =   10
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "√”„ «·ÿ«·»"
         Height          =   375
         Left            =   -70560
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "—ﬁ„  ÿ·» «·«· Õ«ﬁ"
         Height          =   375
         Left            =   -65640
         TabIndex        =   8
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "«·’› «·œ—«”Ì… "
         Height          =   615
         Left            =   -70560
         TabIndex        =   5
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "„”·”·"
         Height          =   375
         Left            =   -65760
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   360
      Top             =   7440
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   9240
      Top             =   8760
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "—ﬁ„ «·ÿ·» «·⁄÷Ê «·√”«”Ì"
      Height          =   615
      Left            =   8400
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "„”·”·"
      DataField       =   "index"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   4800
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Menu file 
      Caption         =   "„·›"
      Begin VB.Menu exit 
         Caption         =   "Œ—ÊÃ"
      End
   End
   Begin VB.Menu find 
      Caption         =   "»ÕÀ"
      Begin VB.Menu by_name 
         Caption         =   "»«·«”„"
         Shortcut        =   ^F
      End
      Begin VB.Menu by_no 
         Caption         =   "»«·—ﬁ„"
         Shortcut        =   ^G
      End
   End
End
Attribute VB_Name = "new_members"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim todayday, todaymonth, todayyear, DOBDAY, DOBmonth, DOBYEAR, Day, Month, year As Integer
Dim bindex As Integer
Dim check As Boolean

Private Sub Calendar1_Click()
    On Error Resume Next

    If bindex = 0 Then
        Text2.text = Calendar1.value
        'Adodc1.Recordset.Fields!original_order_date = Calendar1.Value
    End If

    If bindex = 1 Then

        '        If DataCombo1.Text = "" Then
        '        MsgBox "·« »œ „‰   «·’› «·œ—«”Ì", vbCritical
        '        TXTDOB.BackColor = 255
        '        DataCombo1.Text = ""
        '        Exit Sub
        '        End If

        calcage (Calendar1.value)

        '   If check = False And DataCombo1.Text = "⁄÷Ê ⁄«„·" Then
        '   MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ ⁄÷Ê ⁄«„· «ﬁ· „‰ 18 ”‰…", vbCritical
        '   DataCombo1.Text = "⁄÷Ê  «»⁄"
        '   TXTDOB.BackColor = 255
        '   Exit Sub
        '   End If

        TXTDOB.text = Calendar1.value
        'Adodc1.Recordset.Fields!MEMBER_DOB = Calendar1.Value
    End If

    If bindex = 2 Then
        Text15.text = Calendar1.value
        Adodc1.Recordset.Fields!wife_order_date = Calendar1.value
    End If

    If bindex = 3 Then
        Text20.text = Calendar1.value
        Adodc1.Recordset.Fields!wife_DOB = Calendar1.value
    End If

    If bindex = 4 Then
        Text26.text = Calendar1.value
        Adodc1.Recordset.Fields!sister_order_date = Calendar1.value
    End If

    If bindex = 5 Then
        Text31.text = Calendar1.value
        Adodc1.Recordset.Fields!SISTER_DOB = Calendar1.value
    End If

    If bindex = 6 Then
        Text37.text = Calendar1.value
        Adodc1.Recordset.Fields!grandfather_order_date = Calendar1.value
    End If

    If bindex = 7 Then
        Text42.text = Calendar1.value
        Adodc1.Recordset.Fields!GRANDFATHER_DOB = Calendar1.value
    End If

    If bindex = 8 Then
        Text48.text = Calendar1.value
        Adodc1.Recordset.Fields!grandmother_order_date = Calendar1.value
    End If

    If bindex = 9 Then
        Text53.text = Calendar1.value
        Adodc1.Recordset.Fields!GRANDMOTHER_DOB = Calendar1.value
    End If

    Adodc1.Recordset.update
    Calendar1.Visible = False
End Sub

Private Sub Command1_Click()
    Adodc1.Recordset.AddNew

End Sub

Private Sub Command2_Click()

    If DataCombo1.text = "" Then
        MsgBox "·«»œ „‰ «Œ Ì«— «·”‰… «·œ—«”Ì…"
        Exit Sub
    End If

    'If TXTDOB.Text = "" Then
    'MsgBox "·« »œ „‰ «œŒ«·  «—ÌŒ „Ì·«œ  «·ÿ«·» «·«”«”Ì «Ê·«", vbCritical
    TXTDOB.BackColor = 255
    'DataCombo1.Text = ""
    'Exit Sub
    'End If

    'calcage (TXTDOB.Text)
    'If check = False And DataCombo1.Text = "⁄÷Ê ⁄«„·" Then
    'MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ ⁄÷Ê ⁄«„· «ﬁ· „‰ 18 ”‰…", vbCritical

    'DataCombo1.Text = ""
    'TXTDOB.BackColor = 255
    '
    'Exit Sub
    'End If

    Adodc1.Recordset.Fields!MEMBER_TYPE = DataCombo1.BoundText
    Adodc1.Recordset.update
    'Adodc1.Refresh
End Sub

Function calcage(x As String)

    todayday = Mid(Date, 1, 2)
    todaymonth = Mid(Date, 4, 2)
    todayyear = Mid(Date, 7, 4)

    DOBDAY = Mid(x, 1, 2)
    DOBmonth = Mid(x, 4, 2)
    DOBYEAR = Mid(x, 7, 4)

    If todayday < DOBDAY Then
        todayday = todayday + 30
        TODAYMONT = TODAYMONT - 1
    End If

    Day = todayday - DOBDAY

    If todaymonth < DOBmonth Then
        todaymonth = todaymonth + 12
        todayyear = todayyear - 1
    End If

    Month = todaymonth - DOBmonth
    year = todayyear - DOBYEAR

    If year > 18 Then
        check = True
    Else
        check = False

    End If

End Function

Private Sub Command3_Click(Index As Integer)
    Calendar1.value = DateValue(Now)
    bindex = Index
    Calendar1.Visible = True
    Calendar1.top = Command3(Index).top + 360
    Calendar1.left = Command3(Index).left
End Sub

Private Sub Command4_Click()

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM new_members WHERE security=0 "
    Adodc1.Refresh
End Sub

Private Sub BY_NAME_Click()
    x = InputBox("«œŒ· «·«”„ «Ê Ã“¡ „‰ «·«”„ ··⁄÷Ê «·«”«”Ì", "‘«‘… «·»ÕÀ »«·«”„")

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM new_members where security=0 AND MEMBER_NAME LIKE'%" & x & "%'"
    Adodc1.Refresh
End Sub

Private Sub BY_NO_Click()
    x = InputBox(" «œŒ·  —ﬁ„ «Ê Ã“¡ „‰  —ﬁ„ «·ÿ·» ··⁄÷Ê «·«”«”Ì", "‘«‘… «·»ÕÀ »«·—ﬁ„")

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM new_members where security=0 AND original_order_number LIKE'%" & x & "%'"
    Adodc1.Refresh

End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM new_members WHERE security=0 "
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select *  FROM MEMBER_TYPES"
    Adodc2.Refresh

    check = False
    Calendar1.value = DateValue(Now)
    Adodc1.Recordset.AddNew
    Text22.text = 55

End Sub

