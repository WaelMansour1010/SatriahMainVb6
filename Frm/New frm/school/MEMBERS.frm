VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form MEMBERS 
   BackColor       =   &H80000006&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„·› «·ÿ«·»"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   105
   ClientWidth     =   18330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   18330
   Begin VB.TextBox txtcust_id 
      Alignment       =   2  'Center
      DataField       =   "cust_id"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7320
      TabIndex        =   137
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtaccount 
      Alignment       =   2  'Center
      DataField       =   "account_code"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9720
      TabIndex        =   136
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "⁄—÷ ”‰œ «·œÌ‰"
      Height          =   375
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   135
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton Command16 
      Caption         =   "⁄—÷ «·«ﬁ”«ÿ"
      Height          =   375
      Left            =   8760
      TabIndex        =   134
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox Text57 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   94
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox Text56 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11160
      TabIndex        =   92
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text55 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14880
      TabIndex        =   90
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text54 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   86
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text53 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   84
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text52 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14880
      TabIndex        =   82
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text51 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   80
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox Text50 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   78
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14880
      TabIndex        =   76
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text49 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   74
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text48 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   72
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text47 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14880
      TabIndex        =   70
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text46 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   68
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text45 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   66
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text43 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NATIONAL_id"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   63
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text42 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NATIONAL_id"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14880
      TabIndex        =   61
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text40 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NATIONAL_id"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   58
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text39 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NATIONAL_id"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14880
      TabIndex        =   56
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text38 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NATIONAL_id"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14880
      TabIndex        =   54
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text37 
      Alignment       =   2  'Center
      DataField       =   "KRABA"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7320
      TabIndex        =   52
      Top             =   1080
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000006&
      Caption         =   "„” Ãœ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   35160
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   -720
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
      Caption         =   "«·„—›ﬁ« "
      Height          =   375
      Left            =   8760
      TabIndex        =   49
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   360
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7800
      Top             =   840
   End
   Begin VB.CommandButton Command14 
      Caption         =   "⁄—÷ «·’Ê—…"
      Height          =   495
      Left            =   16200
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text27 
      Alignment       =   2  'Center
      DataField       =   "image_location"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9600
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text26 
      Alignment       =   1  'Right Justify
      DataField       =   "year"
      DataSource      =   "Adodc6"
      Height          =   285
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Text            =   "Text26"
      Top             =   -360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox CARD_VALUE 
      Alignment       =   1  'Right Justify
      DataField       =   "CARD_VALUE"
      DataSource      =   "Adodc5"
      Height          =   375
      Left            =   12480
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Text            =   "Text26"
      Top             =   11880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text25 
      Alignment       =   1  'Right Justify
      DataField       =   "INSTALLMENTS_TOTAL"
      DataSource      =   "Adodc8"
      Height          =   285
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Text            =   "Text25"
      Top             =   9840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command13 
      Height          =   735
      Left            =   14280
      Picture         =   "MEMBERS.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   360
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "SEX"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "MEMBERS.frx":1992
      Left            =   16560
      List            =   "MEMBERS.frx":199C
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   12720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_job"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14880
      TabIndex        =   20
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10920
      TabIndex        =   19
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&«‰Â«¡ «·»ÕÀ"
      Height          =   375
      Left            =   7080
      TabIndex        =   18
      Top             =   11400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   8400
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   " "
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      Caption         =   "5"
      Height          =   8880
      Left            =   -480
      TabIndex        =   7
      Top             =   -120
      Width           =   7455
      Begin VB.CommandButton Command7 
         Caption         =   "Õ›Ÿ"
         Height          =   495
         Left            =   480
         TabIndex        =   138
         Top             =   7080
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         DataField       =   "member_type"
         DataSource      =   "Adodc3"
         Height          =   315
         ItemData        =   "MEMBERS.frx":19AB
         Left            =   600
         List            =   "MEMBERS.frx":19C4
         RightToLeft     =   -1  'True
         TabIndex        =   133
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         DataField       =   "installmentsCount"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   132
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "„Ê«›ﬁ… «·Ê“«—…"
         DataField       =   "WEZARA"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   131
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_ID"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         TabIndex        =   130
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_job_address"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   2760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   120
         Top             =   7200
         Width           =   4095
      End
      Begin VB.OptionButton Option7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "–Â«» ›ﬁÿ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   118
         Top             =   6480
         Width           =   1455
      End
      Begin VB.OptionButton Option8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "⁄ÊœÂ ›ﬁÿ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   117
         Top             =   6480
         Width           =   1695
      End
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "»œÊ‰ ‰ﬁ·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   116
         Top             =   6120
         Width           =   1455
      End
      Begin VB.OptionButton Option6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "–Â«» Ê⁄ÊœÂ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   115
         Top             =   6120
         Width           =   1695
      End
      Begin VB.TextBox Text33 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_born_place"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   110
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_NATIONAL_id"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   108
         Top             =   5040
         Width           =   2175
      End
      Begin VB.TextBox Text32 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_born_place"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   105
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_born_place"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   103
         Top             =   4560
         Width           =   2175
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_born_place"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   100
         Top             =   4080
         Width           =   2175
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_born_place"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   99
         Top             =   4080
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "„” Ãœ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   97
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "„ÕÊ·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text28 
         Alignment       =   2  'Center
         DataField       =   "image_location"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5280
         TabIndex        =   47
         Top             =   7560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command17 
         Caption         =   "⁄—÷ «·’Ê—…"
         Height          =   495
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   7560
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command15 
         Caption         =   "€—«„« "
         Height          =   375
         Left            =   1680
         TabIndex        =   38
         Top             =   8520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "SEX"
         DataSource      =   "Adodc3"
         Height          =   315
         ItemData        =   "MEMBERS.frx":1A0B
         Left            =   600
         List            =   "MEMBERS.frx":1A15
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_certificate"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   26
         Top             =   3120
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         Caption         =   "...«œ—«Ã ’Ê—…"
         Height          =   495
         Left            =   600
         TabIndex        =   25
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_CHILD_ID"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         DataField       =   "MEMBER_CHILD_NAME"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   12
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "MEMBER_CHILD_iMAGE_PATH"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   8160
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   495
         Left            =   2400
         TabIndex        =   10
         Top             =   8280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "«÷«›… ÿ«·» ÃœÌœ"
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   7080
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "«‰”Õ«»"
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   7560
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   495
         Left            =   2040
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
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
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "MEMBERS.frx":1A24
         DataField       =   "acadyearID"
         DataSource      =   "Adodc3"
         Height          =   315
         Left            =   3600
         TabIndex        =   30
         Top             =   2400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "MEMBER_NAME"
         BoundColumn     =   "MEMBER_id"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin DBPIXLib.DBPix20 DBPiX2 
         DataSource      =   "Adodc3"
         Height          =   1455
         Left            =   600
         TabIndex        =   41
         Top             =   240
         Width           =   1455
         _Version        =   131072
         _ExtentX        =   2566
         _ExtentY        =   2566
         _StockProps     =   1
         BackColor       =   255
         _Image          =   "MEMBERS.frx":1A39
         ImageResampleWidth=   100
         ImageResampleHeight=   100
         ImageResampleMode=   0
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
      Begin MSComCtl2.DTPicker XPDtbBill 
         Height          =   330
         Left            =   600
         TabIndex        =   121
         Top             =   3120
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   582
         _Version        =   393216
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   3600
         TabIndex        =   122
         Top             =   3480
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   582
         _Version        =   393216
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   330
         Left            =   600
         TabIndex        =   123
         Top             =   5640
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   330
         Left            =   3600
         TabIndex        =   124
         Top             =   5520
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   582
         _Version        =   393216
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·«ﬁ”«ÿ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6120
         TabIndex        =   127
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " :«·Õ«·… «·’ÕÌ…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4320
         TabIndex        =   119
         Top             =   6840
         Width           =   3495
      End
      Begin VB.Label Label70 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ": «·‰ﬁ· Ê «·Õ—ﬂ…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   4920
         TabIndex        =   114
         Top             =   6120
         Width           =   2655
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·«‰ Â«¡"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2040
         TabIndex        =   113
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒÂ«"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   6000
         TabIndex        =   112
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„’œ—Â«"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1800
         TabIndex        =   111
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·ÂÊÌ…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5880
         TabIndex        =   109
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒÂ«    "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2160
         TabIndex        =   107
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·œÌ«‰…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1560
         TabIndex        =   106
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ã‰”Ì…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   6000
         TabIndex        =   104
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„œÌ‰… «·„Ì·«œ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5760
         TabIndex        =   102
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "œÊ·… «·„Ì·«œ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1920
         TabIndex        =   101
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Õ«· … «·œ—«”Ì…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5640
         TabIndex        =   98
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„”·”· «·’Ê—…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6720
         TabIndex        =   48
         Top             =   7680
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "  «·Ã‰”"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2040
         TabIndex        =   36
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·ÿ«·»"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2400
         TabIndex        =   32
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·”‰… «·œ—«”Ì…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5640
         TabIndex        =   31
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·«·· Õ«ﬁ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   852
         Left            =   3360
         TabIndex        =   29
         Top             =   720
         Width           =   1812
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         DataField       =   "MEMBER_NAME"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   28
         Top             =   -480
         Width           =   3135
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ "
         DataField       =   "MEMBER_TITLE"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5400
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«Œ— ‘Â«œÂ    "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5880
         TabIndex        =   24
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·„Ì·«œ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5760
         TabIndex        =   23
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ÿ«·»"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   5880
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„ "
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5640
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·’Ê—…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   5520
         TabIndex        =   15
         Top             =   8280
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·ÿ·«» «· «»⁄Ì‰ ·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5160
         TabIndex        =   14
         Top             =   -480
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&⁄—÷ «· «»⁄Ì‰"
      Height          =   375
      Left            =   10680
      TabIndex        =   6
      Top             =   8280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÃœÌœ"
      Height          =   375
      Left            =   8760
      TabIndex        =   4
      Top             =   7920
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "IMAGE_PATH"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   12840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14880
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9240
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Caption         =   "  Õ—Ìﬂ"
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
      Left            =   5520
      Top             =   4200
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   615
      Left            =   1200
      Top             =   0
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   615
      Left            =   240
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   615
      Left            =   360
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   615
      Left            =   5880
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   975
      Left            =   4680
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
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
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   330
      Left            =   7080
      TabIndex        =   125
      Top             =   2040
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   582
      _Version        =   393216
      Format          =   100073473
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   330
      Left            =   7080
      TabIndex        =   126
      Top             =   2400
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   582
      _Version        =   393216
      Format          =   100073473
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "MEMBERS.frx":1A51
      DataField       =   "MEMBER_TITLE"
      DataSource      =   "Adodc3"
      Height          =   315
      Left            =   14880
      TabIndex        =   128
      Top             =   5880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "relation_name"
      BoundColumn     =   "relation_name"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   495
      Left            =   8640
      Top             =   6600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   615
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·«ﬁ”«ÿ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   16920
      TabIndex        =   129
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label69 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·«ŒÊ… «·«‰«À"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9240
      TabIndex        =   95
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label68 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·«ŒÊ… «·–ﬂÊ—"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   12960
      TabIndex        =   93
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label67 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «›—«œ «·«”—…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   16440
      TabIndex        =   91
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label66 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "À«‰Ì« «·»Ì«‰«  «·«Ã „«⁄Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10680
      TabIndex        =   89
      Top             =   4920
      Width           =   4575
   End
   Begin VB.Label Label64 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ›Ì Õ«·…  ⁄–— «·« ’«· »ﬂ ⁄‰œ «·÷—Ê—… "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   13800
      TabIndex        =   88
      Top             =   4200
      Width           =   4695
   End
   Begin VB.Label Label63 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄Ì"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8640
      TabIndex        =   87
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label62 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«·‘«—⁄ «·—∆Ì”Ì"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   12360
      TabIndex        =   85
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label61 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " «·ÕÌ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   16200
      TabIndex        =   83
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label60 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÃÊ«·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8760
      TabIndex        =   81
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label59 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "’·… «·ﬁ—«»…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   12360
      TabIndex        =   79
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«·«”„"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   16560
      TabIndex        =   77
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÃÊ«·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8640
      TabIndex        =   75
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Â« › «·„‰“·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   12360
      TabIndex        =   73
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„Â‰…  «·«„"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   16440
      TabIndex        =   71
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÃÊ«·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8640
      TabIndex        =   69
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Â« › «·⁄„·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   12360
      TabIndex        =   67
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «‰ Â«∆Â«"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8880
      TabIndex        =   65
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label52 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„’œ—Â«"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   12480
      TabIndex        =   64
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„  «·«ﬁ«„…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   16440
      TabIndex        =   62
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label50 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «‰ Â«∆…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8760
      TabIndex        =   60
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„’œ—Â"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   12480
      TabIndex        =   59
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ ÃÊ«“ «·”›—"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   16440
      TabIndex        =   57
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«·Ã‰”Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   16320
      TabIndex        =   55
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "’·… «·ﬁ—«»…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9240
      TabIndex        =   53
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«Ê·« :„⁄·Ê„«  ⁄‰ Ê·Ì «·«„—"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9960
      TabIndex        =   51
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„”·”· «·’Ê—…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10320
      TabIndex        =   45
      Top             =   -120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  «·Ã‰”"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   17760
      TabIndex        =   34
      Top             =   12720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„Â‰… Ê·Ì «·«„—"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   16440
      TabIndex        =   22
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«·«”„ À«·ÀÌ«"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   13320
      TabIndex        =   21
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      FillColor       =   &H00FFFFFF&
      Height          =   8775
      Left            =   6720
      Top             =   0
      Width           =   12255
   End
   Begin VB.Image Image3 
      Height          =   1440
      Left            =   5160
      Picture         =   "MEMBERS.frx":1A66
      Stretch         =   -1  'True
      Top             =   6360
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«·’Ê—…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   12720
      TabIndex        =   3
      Top             =   12240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ Ê·Ì «·«„—"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   16440
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "MEMBERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bindex As Integer
Dim CHECKN As Integer
Dim CHECKN1 As Integer
Dim check  As Boolean
Dim todayday, todaymonth, todayyear, DOBDAY, DOBmonth, DOBYEAR, Day, Month, year As Integer

Private Sub BY_NAME_Click()
    x = InputBox("«œŒ· «·«”„ «Ê Ã“¡ „‰ «·«”„", "‘«‘… «·»ÕÀ »«·«”„")

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM MEMBERS where MEMBER_NAME LIKE'%" & x & "%'"
    Adodc1.Refresh
End Sub

Private Sub BY_NO_Click()
    x = InputBox("«œŒ· «·—ﬁ„ «Ê Ã“¡ „‰ «·—ﬁ„", "‘«‘… «·»ÕÀ »«·—ﬁ„")

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM MEMBERS where MEMBER_ID LIKE'%" & x & "%'"
    Adodc1.Refresh

End Sub

Private Sub Calendar1_Click()

    If bindex = 0 Then
        Text9.text = Calendar1.value
        calcage (Text9.text)
  
        If check = False Then
            MsgBox "Â–« «·⁄÷Ê «ﬁ· „‰ 18 ⁄«„ Ê·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ ⁄÷Ê ⁄«„·", vbInformation
        Else
            MsgBox "Â–« «·⁄÷Ê «ﬂ»— „‰ 18 ⁄«„ Ê·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ ⁄÷Ê  «»⁄ ", vbInformation
        End If

        Adodc1.Recordset.Fields!MEMBER_DOB = Calendar1.value

    End If

    If bindex = 1 Then
        Text20.text = Calendar1.value

        '  If check = True Then
        '  MsgBox "Â–« «·⁄÷Ê «ﬂ»— „‰ 18 ⁄«„ Ê·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ ⁄÷Ê  «»⁄ ", vbInformation
        '  Else
        '   MsgBox "Â–« «·⁄÷Ê «ﬁ· „‰ 18 ⁄«„ Ê·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ ⁄÷Ê ⁄«„·", vbInformation
        '  End If
        Adodc3.Recordset.Fields!MEMBER_DOB = Calendar1.value
  
    End If

    Calendar1.Visible = False
    'Adodc1.Recordset.Update
End Sub

Private Sub Command1_Click()
    Adodc1.Recordset.AddNew
    Me.Text2.text = CStr(new_id("MEMBERS", "MEMBER_ID", "", True))
  
    'Text23.text = DateValue(Now)
    CHECKN1 = 1
End Sub

Private Sub Command10_Click()

    If Text6.text = "" Then MsgBox "«Œ — ÿ«·» «Ê·«", vbCritical: Exit Sub
    DBPiX2.ImageClear
    x = MsgBox("Â·  —Ìœ ’Ê—… „‰ „·›", vbExclamation + vbYesNoCancel)

    If x = vbYes Then
        DBPiX2.ImageLoad

        DoEvents
        MsgBox " „  Õ„Ì· «·’Ê—…"
    Else

        If x = vbNo Then
            DBPix1.TWAINAcquire
            MsgBox " „ „”Õ ÷Ê∆Ì  ··’Ê—…"

            DoEvents
        Else

            Exit Sub
        End If
    End If

    DBPiX2.ImageSaveFile (system_path & "\images\std" & Text2.text & "-" & Text6.text & ".JPG")
    'NEW_IMAGE = False

    Adodc3.Recordset.Fields!image_location = "std" & Text2.text & "-" & Text6.text
    Adodc3.Recordset.update

    'Cd1.ShowOpen
    'Text28 = Mid(Cd1.FileTitle, 1, Len(Cd1.FileTitle) - 4)
    'Adodc3.Recordset.Fields!image_location = Text28
    'Adodc3.Recordset.update

    'DBPiX2.ImageViewFile (Text28 & ".JPG")
    'DBPiX2.ImageSaveFile (system_path & "\images\std" & Text28.text & ".JPG")

    'cd1.ShowOpen
    'Text1 = cd1.FileName
    'Adodc3.Recordset.Fields!MEMBER_CHILD_iMAGE_PATH = cd1.FileName

    'DBPiX2.ImageLoad

    'If Adodc3.Recordset.EOF <> True And Adodc3.Recordset.BOF <> True Then
    'Adodc3.Recordset.MoveNext
    'Adodc3.Recordset.MovePrevious
    'Else
    'Adodc3.Recordset.MoveLast
    'End If

    'DBPix2.ImageSave
    'Adodc3.Recordset.Update
End Sub

Private Sub Command11_Click()
    INSTALLMENT_DATA.Show
    INSTALLMENT_DATA.Adodc1.CommandType = adCmdText
    INSTALLMENT_DATA.Adodc1.RecordSource = "select *  FROM Installments where child_id=0 and MEMBER_ID ='" & Text2.text & "'"
    INSTALLMENT_DATA.Adodc1.Refresh

    INSTALLMENT_DATA.Adodc2.CommandType = adCmdText
    INSTALLMENT_DATA.Adodc2.RecordSource = "select INSTALLMENT_NO,INSTALLMENT_VALUE,DATE_OF_PAYED,ACTIVATED  FROM INSTALLMENT_DETAILS where MEMBER_ID LIKE'%" & Text2.text & "%' and child_id=0 "
    INSTALLMENT_DATA.Adodc2.Refresh
End Sub

Private Sub Command12_Click()
    FINES_DATA.Show
    FINES_DATA.Adodc1.CommandType = adCmdText
    FINES_DATA.Adodc1.RecordSource = "select *  FROM FINES where child_id=0 and MEMBER_ID LIKE'%" & Text2.text & "%'"
    FINES_DATA.Adodc1.Refresh

    FINES_DATA.Adodc2.CommandType = adCmdText
    FINES_DATA.Adodc2.RecordSource = "select FINES_NO,FINES_VALUE,PAYED_DATE,ACTIVATED  FROM FINES_DETAILS where child_id=0 and MEMBER_ID ='" & Text2.text & "'"
    FINES_DATA.Adodc2.Refresh

End Sub

Private Sub Command13_Click()
    ' BY_NO_Click
    member_search.Show
    member_search.from = 4
End Sub

Private Sub Command14_Click()
    DBPix1.ImageViewFile (system_path & "\images\std" & Text27.text & ".JPG")
End Sub

Private Sub Command15_Click()

    If Text6.text = 1 Then
        FINES_DATA.Show
        FINES_DATA.Adodc1.CommandType = adCmdText
        FINES_DATA.Adodc1.RecordSource = "select *  FROM FINES where child_id=1 and MEMBER_ID LIKE'%" & Text2.text & "%'"
        FINES_DATA.Adodc1.Refresh

        FINES_DATA.Adodc2.CommandType = adCmdText
        FINES_DATA.Adodc2.RecordSource = "select FINES_NO,FINES_VALUE,PAYED_DATE,ACTIVATED  FROM FINES_DETAILS where child_id=1 and MEMBER_ID LIKE'%" & Text2.text & "%'"
        FINES_DATA.Adodc2.Refresh
    End If

End Sub

Private Sub Command16_Click()

    If IsNumeric(txtcust_id.text) Then
        INSTALLMENT_DATA.Show
        'INSTALLMENT_DATA.Check1.Enabled = False
        INSTALLMENT_DATA.Adodc1.CommandType = adCmdText
        INSTALLMENT_DATA.Adodc1.RecordSource = "select *  FROM INSTALLMENT_DETAILS where  cust_id =" & Me.txtcust_id.text
        INSTALLMENT_DATA.Adodc1.Refresh

        INSTALLMENT_DATA.lblcustid = Me.txtcust_id.text
        INSTALLMENT_DATA.id.text = Text2.text
        INSTALLMENT_DATA.TxtName.text = Text11.text

    End If

End Sub

Private Sub Command17_Click()
    DBPiX2.ImageViewFile (system_path & "\images\std" & Text28.text & ".JPG")

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

Private Sub Command18_Click()
    On Error Resume Next

    If my_language = "E" Then
        If Text2.text = "" Then MsgBox "Select Voucher First": Exit Sub

    Else

        If Text2.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— ÿ«·»  «Ê·«": Exit Sub
    End If

    imaged.Show
    imaged.txtopeation_type = "‘Â«œ«  «·ÿ«·»"
    imaged.SUBJECT_NO = Text2.text

    If my_language = "E" Then
        imaged.Label6.Caption = "Voucher #"
        imaged.Caption = "Voucher Attachments"
    Else
        imaged.Label6.Caption = "—ﬁ„ «·ÿ«·»"
        imaged.Caption = "„—›ﬁ«  «·ÿ«·»"
    End If

    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type ='‘Â«œ«  «·ÿ«·»' and subject_no='" & Text2.text & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Sub

Private Sub Command2_Click()

    If CHECKN1 = 1 Then
        CHECKN1 = 0
 
        txtaccount.text = ModAccounts.AddNewAccount("a1a2a3", Trim$(Me.Text11.text), True, False)
 
        Adodc9.RecordSource = "select * from TblCustemers where CusID=0"
        Adodc9.Refresh
 
        Adodc9.Recordset.AddNew
        Adodc9.Recordset.Fields!CusID = CStr(new_id("TblCustemers", "CusID", "", True))
        txtcust_id.text = Adodc9.Recordset.Fields!CusID
        Adodc9.Recordset.Fields!CusName = Text11.text
        Adodc9.Recordset.Fields!OpenBalance = 0
        Adodc9.Recordset.Fields!OpenBalanceType = Null
        Adodc9.Recordset.Fields!type = 1
        Adodc9.Recordset.Fields!OpenBalanceDate = Date
 
        Adodc9.Recordset.Fields!OpenBalanceDate = Date
        Adodc9.Recordset.Fields!CreditlimitCredit = 0
        Adodc9.Recordset.Fields!SaleType = 0
        Adodc9.Recordset.Fields!Account_Code = txtaccount.text
        Adodc9.Recordset.Fields!Trans_DiscountPur = 0
        Adodc9.Recordset.Fields!Trans_DiscountTypePur = 0
 
        Adodc9.Recordset.update
 
        Adodc1.Recordset.update

        If Not IsNumeric(Text2.text) Then Exit Sub
        Adodc7.Recordset.AddNew
        Adodc7.Recordset.Fields!Sanad_No = CStr(new_id("sanad_dean", "sanad_no", "", True))
        Adodc7.Recordset.Fields!sanad_date = Date
        Adodc7.Recordset.Fields!member_id = Text2.text
        Adodc7.Recordset.Fields!member_name = Text11.text
        Adodc7.Recordset.Fields!cust_id = Me.txtcust_id.text
        Adodc7.Recordset.update
 
    Else
        Adodc1.Recordset.update
 
    End If
 
End Sub

Private Sub Command3_Click()

    If Not IsNumeric(txtcust_id.text) Then Exit Sub
    sanad_dean.Show

    sanad_dean.lblaccountcode.Caption = txtaccount.text
    sanad_dean.Adodc1.CommandType = adCmdText
    sanad_dean.Adodc1.RecordSource = "select*  FROM sanad_dean where cust_id=" & txtcust_id.text
    sanad_dean.Adodc1.Refresh

    sanad_dean.Adodc2.CommandType = adCmdText
    sanad_dean.Adodc2.RecordSource = "select *  FROM member_child where cust_id=" & txtcust_id.text
    sanad_dean.Adodc2.Refresh

    sanad_dean.lblid = Text2.text
    sanad_dean.LblName = Text11.text
    sanad_dean.txtcust_id.text = Me.txtcust_id.text

End Sub

Private Sub Command4_Click()
    Frame1.Visible = True

    If Text6.text = "" Then Exit Sub
    DBPiX2.ImageClear
 
    If Dir(system_path & "\images\std" & Text2.text & "-" & Text6.text & ".JPG") <> "" Then
        DBPiX2.ImageLoadFile (system_path & "\images\std" & Text2.text & "-" & Text6.text & ".JPG")
    End If
 
End Sub

Private Sub Command40_Click(Index As Integer)
    bindex = Index
    Calendar1.Visible = True
    'Calendar1.top = Command40(Index).top + 360
    'Calendar1.left = Command40(Index).left
End Sub

Private Sub Command6_Click()
    lastNo = 0

    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.MoveLast
        lastNo = Adodc3.Recordset.Fields!MEMBER_CHILD_ID
    End If

    Adodc3.Recordset.AddNew

    Adodc3.Recordset.Fields!member_id = Text2.text
    Adodc3.Recordset.Fields!MEMBER_CHILD_ID = lastNo + 1
    CHECKN = 1

End Sub

Private Sub Command7_Click()
    On Error Resume Next
    Dim cost As Double
    Dim discount As Double
    Dim NetCost As Double
    Dim slice As Double

    If Text4.text = "" Then MsgBox "«ﬂ » «”„ «·ÿ«·» «Ê·«", vbCritical: Exit Sub
   
    If DataCombo2.BoundText = "" Then
        MsgBox "Õœœ «·”‰… «·œ—«”Ì… «Ê·«", vbCritical: Exit Sub
    End If
 
    If Combo3.ListIndex < 0 Then
        MsgBox "Õœœ ‰Ê⁄ «·ÿ«·»", vbCritical: Exit Sub

    End If
 
    If Not IsNumeric(Text8.text) Then
        MsgBox "Õœœ ⁄œœ «·«ﬁ”«ÿ", vbCritical: Exit Sub

    End If
 
    Adodc11.CommandType = adCmdText
    Adodc11.RecordSource = "select * from  INSTALLMENT_DETAILS where member_id=" & Text2.text & "and child_id=" & Text6.text
    Adodc11.Refresh

    For i = 1 To Adodc11.Recordset.RecordCount
        Adodc11.Recordset.delete
        Adodc11.Recordset.MoveNext
    Next i

    Adodc11.CommandType = adCmdText
    Adodc11.RecordSource = "select*  FROM member_TYPES WHERE MEMBER_id= " & DataCombo2.BoundText
    Adodc11.Refresh

    If Adodc11.Recordset.RecordCount = 0 Then MsgBox "Â‰«ﬂ Œÿ√ ›Ì «·”‰… «·œ—«”Ì…":  CHECKN = 1: Exit Sub

    Select Case Combo3.ListIndex

        Case 0
            cost = Adodc11.Recordset.Fields!value
            discount = 0
            NetCost = cost

        Case 1
            cost = Adodc11.Recordset.Fields!value
            discount = (cost * Adodc11.Recordset.Fields!d1 / 100)
            NetCost = cost - discount

        Case 2
            cost = Adodc11.Recordset.Fields!value
            discount = (cost * Adodc11.Recordset.Fields!d2 / 100)
            NetCost = cost - discount

        Case 3
            cost = Adodc11.Recordset.Fields!value
            discount = (cost * Adodc11.Recordset.Fields!d3 / 100)
            NetCost = cost - discount

        Case 4
            cost = Adodc11.Recordset.Fields!value
            discount = (cost * Adodc11.Recordset.Fields!d4 / 100)
            NetCost = cost - discount

        Case 5
            cost = Adodc11.Recordset.Fields!value
            discount = (cost * Adodc11.Recordset.Fields!d4 / 100)
            NetCost = cost - discount

        Case 6
            cost = Adodc11.Recordset.Fields!value
            discount = (cost * Adodc11.Recordset.Fields!d4 / 100)
            NetCost = cost - discount
    End Select
 
    Adodc3.Recordset.Fields!cost = cost
    Adodc3.Recordset.Fields!discount = discount
    Adodc3.Recordset.Fields!NetCost = NetCost
    Adodc3.Recordset.Fields!acadyearname = DataCombo2.text
    Adodc3.Recordset.Fields!WEZARA = Check1.value
    Adodc3.Recordset.Fields!cust_id = txtcust_id.text
 
    Adodc3.Recordset.Fields!membership_value = Adodc11.Recordset.Fields!membership_value
    Adodc3.Recordset.Fields!total = NetCost + Adodc11.Recordset.Fields!membership_value

    'Adodc3.Recordset.MoveLast
            
    slice = NetCost / val(Text8.text)

    For i = 1 To Text8.text
        Adodc4.Recordset.AddNew
            
        Adodc4.Recordset.Fields!member_id = Text2.text
        Adodc4.Recordset.Fields!cust_id = txtcust_id.text

        Adodc4.Recordset.Fields!CHILD_ID = Text6.text
        Adodc4.Recordset.Fields!member_name = Text11.text
        Adodc4.Recordset.Fields!MEMBER_CHILD_NAME = Text4.text
        Adodc4.Recordset.Fields!INSTALLMENT_NO = i
        Adodc4.Recordset.Fields!installment_value = slice
            
    Next i

    Adodc4.Recordset.update
    Adodc3.Recordset.Fields!installment_value = slice

    Adodc3.Recordset.update
    ' Adodc3.Recordset.MoveLast
 
End Sub

Private Sub Command8_Click()

    If Adodc3.Recordset.RecordCount = 0 Then
        MsgBox "·« ÌÊÃœ ”Ã·«  ·Õ–›Â«", vbCritical
        Exit Sub
    End If

    x = MsgBox("Â· «‰  „ √ﬂœ „‰ Õ–› Â–« «·”Ã·", vbYesNo)
    Adodc3.Refresh

    If x = vbYes And Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.delete
        Adodc3.Refresh
    End If

End Sub

Private Sub Command9_Click()

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM MEMBERS "
    Adodc1.Refresh
End Sub

Private Sub DataCombo2_Click(Area As Integer)
 
    'If DataCombo2.Text = "⁄÷Ê ⁄«„·" Then
    'DataCombo3.Text = "«·“ÊÃ…"
    'Else
    'DataCombo3.Text = ""

    'End If
End Sub

Private Sub DataCombo3_Click(Area As Integer)
    'If DataCombo3.Text <> "«·“ÊÃ…" Then
    'DataCombo2.Text = "⁄÷Ê  «»⁄"
    'Else
    'DataCombo2.Text = "⁄÷Ê ⁄«„·"
    '
    'End If

End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    system_path = App.path
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    'Frame1.Visible = False
    CHECKN = 0
    CHECKN1 = 0

    'system_path = "D:\my works\accountant\28  01 2011\SourceCode\SourceCode"
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from  MEMBERS "
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from  MEMBER_TYPES"
    Adodc2.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    'Adodc3.RecordSource = "select * from  MEMBER_CHILD"
    'Adodc3.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "select * from INSTALLMENT_DETAILS where op_no=0  "
    Adodc4.Refresh

    Adodc7.ConnectionString = connection_string
    Adodc7.CommandType = adCmdText
    Adodc7.RecordSource = "select * from  sanad_dean where sanad_no=0"
    Adodc7.Refresh

    Adodc11.ConnectionString = connection_string
    Adodc11.CommandType = adCmdText

    Adodc9.ConnectionString = connection_string
    Adodc9.CommandType = adCmdText

End Sub

Private Sub Text1_Change()
    ' On Error GoTo ll
    ' Image2.Picture = LoadPicture(Text1.Text)
    ' Exit Sub

    'll:
    '  Image2.Picture = Image3.Picture
End Sub

Private Sub Text2_Change()

    If Text2.text = "" Then Exit Sub
    'DBPix1.ImageClear

    'If DBPix1.ImageLoadFile(system_path & "\images\std" & Text2.text & ".JPG") = True Then

    If Text2.text <> "" Then

        If Dir(system_path & "\images\std" & Text2.text & "-" & Text6.text & ".JPG") <> "" Then
            DBPiX2.ImageLoadFile (system_path & "\images\std" & Text2.text & "-" & Text6.text & ".JPG")
        End If

        Adodc3.ConnectionString = Cn.ConnectionString
        Adodc3.CommandType = adCmdText
        Adodc3.RecordSource = "select * from MEMBER_CHILD where   MEMBER_id=" & Text2.text & "order by member_id,member_child_id"
        Adodc3.Refresh

    End If

End Sub

Private Sub Text27_Change()
    'Timer1.Enabled = True
End Sub

Private Sub Text28_Change()
    'Timer2.Enabled = True
End Sub

Private Sub Text5_Change()
    ' On Error GoTo ll
    ' Image1.Picture = LoadPicture(Text5.Text)
    ' Exit Sub
    '
    'll:
    '  Image1.Picture = Image3.Picture
 
End Sub

Private Sub Text6_Change()

    If Text6.text = "" Then Exit Sub
    DBPiX2.ImageClear

    If Dir(system_path & "\images\std" & Text2.text & "-" & Text6.text & ".JPG") <> "" Then
        DBPiX2.ImageLoadFile (system_path & "\images\std" & Text2.text & "-" & Text6.text & ".JPG")
    End If

End Sub

Private Sub Text9_Change()
    DBPix1.ImageViewFile (system_path & "\images\std" & Text27.text & ".JPG")
End Sub

Private Sub Timer1_Timer()
    'DBPix1.ImageViewFile (system_path & "\images\std" & Text27.text & ".JPG")
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    DBPiX2.ImageViewFile (system_path & "\images\std" & Text28.text & ".JPG")
    Timer2.Enabled = False
End Sub
