VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{784C0C13-85E7-4E11-A8FB-F0243A135D03}#2.0#0"; "SuperLablel.ocx"
Begin VB.Form baranchesE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·›—Ê⁄"
   ClientHeight    =   7725
   ClientLeft      =   4800
   ClientTop       =   375
   ClientWidth     =   13140
   Icon            =   "baranchese.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   13140
   Begin VB.Frame Frame9 
      Caption         =   "Õ”«»«  «·«‰ «Ã"
      Height          =   1935
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   127
      Top             =   1920
      Width           =   11415
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":000C
         DataField       =   "a37"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   37
         Left            =   120
         TabIndex        =   128
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0021
         DataField       =   "a38"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   38
         Left            =   120
         TabIndex        =   129
         Top             =   840
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0036
         DataField       =   "a39"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   39
         Left            =   120
         TabIndex        =   130
         Top             =   1320
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
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
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„’«—Ì› «·«‰ «Ã „Ê«œ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   37
         Left            =   9285
         TabIndex        =   133
         Top             =   405
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„’«—Ì› «·«‰ «Ã  «ÃÊ—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   38
         Left            =   9285
         TabIndex        =   132
         Top             =   885
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "„’«—Ì› «·«‰ «Ã  „’«—Ì› ’‰«⁄Ì…"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   39
         Left            =   9285
         TabIndex        =   131
         Top             =   1365
         Width           =   1695
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Õ”«»«  „ ‰Ê⁄Â"
      Height          =   3855
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   102
      Top             =   1920
      Visible         =   0   'False
      Width           =   11415
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":004B
         DataField       =   "a18"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   18
         Left            =   240
         TabIndex        =   103
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0060
         DataField       =   "a19"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   19
         Left            =   240
         TabIndex        =   104
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0075
         DataField       =   "a6"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   105
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":008A
         DataField       =   "a20"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   20
         Left            =   240
         TabIndex        =   106
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":009F
         DataField       =   "a21"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   22
         Left            =   240
         TabIndex        =   107
         Top             =   1800
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":00B4
         DataField       =   "a22"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   21
         Left            =   240
         TabIndex        =   108
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":00C9
         DataField       =   "a33"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   33
         Left            =   240
         TabIndex        =   121
         Top             =   2520
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":00DE
         DataField       =   "a34"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   34
         Left            =   240
         TabIndex        =   122
         Top             =   2880
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":00F3
         DataField       =   "a35"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   35
         Left            =   240
         TabIndex        =   123
         Top             =   3240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Bety Cash"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   35
         Left            =   9570
         TabIndex        =   124
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«Ì—«œ« "
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   34
         Left            =   9600
         TabIndex        =   120
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„’—Ê›« "
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   33
         Left            =   9600
         TabIndex        =   119
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·»‰Êﬂ"
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   20
         Left            =   9525
         TabIndex        =   114
         Top             =   765
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» Ê”Ìÿ «›  «ÕÌ"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   19
         Left            =   9525
         TabIndex        =   113
         Top             =   1125
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "›—Êﬁ«  ⁄„·…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   9555
         TabIndex        =   112
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·’‰œÊﬁ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   9570
         TabIndex        =   111
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄Ã“ ›Ì «·‰ﬁœÌ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   21
         Left            =   9480
         TabIndex        =   110
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "“Ì«œ… ›Ì «·‰ﬁœÌ…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   22
         Left            =   9600
         TabIndex        =   109
         Top             =   2160
         Width           =   1575
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Õ”«»«  «·„‘«—Ì⁄"
      Height          =   2655
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   90
      Top             =   1920
      Width           =   11415
      Begin VB.Frame Frame8 
         Caption         =   "«·Ì… «·„‘«—Ì⁄"
         Height          =   2055
         Left            =   9240
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   240
         Width           =   2055
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» ‰Ÿ«„Ì «·„‘«—Ì⁄"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   32
            Left            =   240
            TabIndex        =   117
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» „’—Ê›«  «·„‘«—Ì⁄"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» «Ì—«œ«  «·„‘«—Ì⁄"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   99
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» „Êœ «·„‘«—Ì⁄"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   27
            Left            =   480
            TabIndex        =   98
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» «ÃÊ— «·„‘«—Ì⁄"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   28
            Left            =   480
            TabIndex        =   97
            Top             =   1320
            Width           =   1455
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0108
         DataField       =   "a23"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   23
         Left            =   240
         TabIndex        =   91
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":011D
         DataField       =   "a14"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   14
         Left            =   240
         TabIndex        =   92
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0132
         DataField       =   "a15"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   15
         Left            =   240
         TabIndex        =   93
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0147
         DataField       =   "a27"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   27
         Left            =   240
         TabIndex        =   94
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":015C
         DataField       =   "a28"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   28
         Left            =   240
         TabIndex        =   95
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0171
         DataField       =   "a32"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   32
         Left            =   240
         TabIndex        =   118
         Top             =   1800
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «Ì—«œ«  «·Œœ„« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   23
         Left            =   9240
         TabIndex        =   101
         Top             =   2280
         Width           =   1935
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Õ”«»«  «·«’Ê·"
      Height          =   2655
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   83
      Top             =   1920
      Visible         =   0   'False
      Width           =   11415
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0186
         DataField       =   "a24"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   24
         Left            =   120
         TabIndex        =   136
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":019B
         DataField       =   "a25"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   25
         Left            =   120
         TabIndex        =   137
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":01B0
         DataField       =   "a26"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   26
         Left            =   120
         TabIndex        =   138
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":01C5
         DataField       =   "a31"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   31
         Left            =   120
         TabIndex        =   139
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":01DA
         DataField       =   "a40"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   40
         Left            =   120
         TabIndex        =   140
         Top             =   1800
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":01EF
         DataField       =   "a41"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   41
         Left            =   120
         TabIndex        =   141
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_Name"
         BoundColumn     =   "Account_Code"
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
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Assets Opening Balances"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   41
         Left            =   9045
         TabIndex        =   142
         Top             =   2205
         Width           =   2055
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Assets Sales Losses"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   40
         Left            =   9240
         TabIndex        =   135
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Assets Sales Profir"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   31
         Left            =   9240
         TabIndex        =   87
         Top             =   1515
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   " „Ã„⁄ «·«Â·«ﬂ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   26
         Left            =   9240
         TabIndex        =   86
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·«Â·«ﬂ"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   25
         Left            =   9240
         TabIndex        =   85
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»  «·«’·"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   24
         Left            =   9240
         TabIndex        =   84
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Õ”«»«  «·–„„"
      Height          =   3135
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   1920
      Visible         =   0   'False
      Width           =   11415
      Begin VB.Frame Frame5 
         Caption         =   "«·Ì… «·„ÊŸ›Ì‰"
         Height          =   1215
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   240
         Width           =   1695
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "«·«ÃÊ— «·„” Õﬁ… "
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   79
            Top             =   550
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "–„„  «·„ÊŸ›Ì‰"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   78
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„Œ’’« "
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   77
            Top             =   840
            Width           =   1455
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0204
         DataField       =   "a16"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   16
         Left            =   240
         TabIndex        =   70
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0219
         DataField       =   "a7"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   7
         Left            =   240
         TabIndex        =   71
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":022E
         DataField       =   "a29"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   29
         Left            =   240
         TabIndex        =   72
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0243
         DataField       =   "a8"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   8
         Left            =   240
         TabIndex        =   73
         Top             =   1800
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0258
         DataField       =   "a9"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   9
         Left            =   240
         TabIndex        =   74
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":026D
         DataField       =   "a30"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   30
         Left            =   240
         TabIndex        =   75
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0282
         DataField       =   "a36"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   36
         Left            =   240
         TabIndex        =   125
         Top             =   2520
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»   „ﬁ«Ê·Ì «·»«ÿ‰"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   36
         Left            =   9360
         TabIndex        =   126
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»   «·⁄„·«¡"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   8
         Left            =   9360
         TabIndex        =   82
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»   «·„Ê—œÌ‰"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   9
         Left            =   9480
         TabIndex        =   81
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·«ÃÊ— "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   9480
         TabIndex        =   80
         Top             =   1440
         Width           =   1455
      End
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Left            =   10080
      TabIndex        =   44
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«»«  «·„Œ«“‰"
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
      MICON           =   "baranchese.frx":0297
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame10 
      Caption         =   "Õ”«»«  «·„Œ«“‰"
      Height          =   4575
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1920
      Width           =   11415
      Begin VB.Frame Frame4 
         Caption         =   "«·Ì… «·„Œ«“‰"
         Height          =   1575
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   1935
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "«· ”ÊÌ«  «·Ã—œÌ…"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   67
            ToolTipText     =   "Ì »⁄ «·«’Ê·"
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Œ”«∆— ›ﬁœ Ê ·›"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   66
            ToolTipText     =   "Ì »⁄ «·„’—Ê›« "
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "Õ”«» «·„Œ“Ê‰"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   65
            ToolTipText     =   "Ì »⁄ «·«’Ê·"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Labelx 
            Caption         =   "Õ”«»  Âœ«Ì« Ê⁄Ì‰« "
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   17
            Left            =   360
            TabIndex        =   64
            Top             =   1200
            Width           =   1455
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":02B3
         DataField       =   "a0"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   45
         ToolTipText     =   "Ì »⁄ «·«’Ê·"
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":02C8
         DataField       =   "a10"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   10
         Left            =   240
         TabIndex        =   46
         ToolTipText     =   "Ì »⁄ «·„’—Ê›« "
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":02DD
         DataField       =   "a11"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   11
         Left            =   240
         TabIndex        =   47
         ToolTipText     =   "Ì »⁄ «·«’Ê·"
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":02F2
         DataField       =   "a1"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   48
         ToolTipText     =   "Ì »⁄ «·„’—Ê›« "
         Top             =   1800
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0307
         DataField       =   "a2"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   49
         ToolTipText     =   "Ì »⁄ «·«Ì—«œ« "
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":031C
         DataField       =   "a3"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   50
         Top             =   2520
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0331
         DataField       =   "a4"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   51
         Top             =   2880
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0346
         DataField       =   "a5"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   52
         Top             =   3240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":035B
         DataField       =   "a12"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   12
         Left            =   240
         TabIndex        =   53
         Top             =   3600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0370
         DataField       =   "a13"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   13
         Left            =   240
         TabIndex        =   54
         Top             =   3960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "baranchese.frx":0385
         DataField       =   "a17"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   17
         Left            =   240
         TabIndex        =   55
         Top             =   1440
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "Account_NameEng"
         BoundColumn     =   "Account_Code"
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
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ’„ „ﬂ ”»"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   13
         Left            =   9240
         TabIndex        =   62
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ’„ „”„ÊÕ »…"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   9240
         TabIndex        =   61
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»  ﬂ·›… «·„»Ì⁄« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   9240
         TabIndex        =   60
         ToolTipText     =   "Ì »⁄ «·„’—Ê›« "
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·„»Ì⁄« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   9240
         TabIndex        =   59
         ToolTipText     =   "Ì »⁄ «·«Ì—«œ« "
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» „—œÊœ«  «·„»Ì⁄« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   9240
         TabIndex        =   58
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» «·„‘ —Ì« "
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   9240
         TabIndex        =   57
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» „—œÊœ«  «·„‘ —Ì« "
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   9240
         TabIndex        =   56
         Top             =   3240
         Width           =   1935
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "address"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4920
      TabIndex        =   39
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Frame Frame6 
      Caption         =   "œ·«·«  «·«·Ê«‰"
      ClipControls    =   0   'False
      Height          =   975
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   6600
      Width           =   3135
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» —∆Ì”Ì"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«» ‰Â«∆Ì"
         Height          =   255
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.TextBox txtnamee 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "branch_namee"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      TabIndex        =   32
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "tel"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   9600
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   4320
      TabIndex        =   22
      Top             =   6600
      Width           =   3975
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   23
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Õ›Ÿ"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "baranchese.frx":039A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Õ–›"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "baranchese.frx":03B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÿ»«⁄…"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "baranchese.frx":03D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   26
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ÃœÌœ"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "baranchese.frx":03EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   960
         Top             =   600
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   582
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
      Begin VB.Label Label5 
         Caption         =   "Label2"
         Height          =   15
         Left            =   -120
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   1200
      TabIndex        =   17
      Top             =   9000
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3840
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   1200
      TabIndex        =   12
      Top             =   9000
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   " «··€…"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "EN"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "baranchese.frx":040A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   -1  'True
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2175
      Left            =   600
      TabIndex        =   8
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Label Label9 
         Caption         =   "Tel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   0
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Branch#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   9480
      Width           =   7095
      Begin MSAdodcLib.Adodc user_priviliges_adodc 
         Height          =   495
         Left            =   120
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.Label screen_name 
         Alignment       =   1  'Right Justify
         Caption         =   "M30"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label adodc4error 
         Alignment       =   1  'Right Justify
         Caption         =   "Label4"
         DataField       =   "employee_id"
         DataSource      =   "user_priviliges_adodc"
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox txtnameA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "branch_name"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox txtcode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "branch_id"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   10440
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   480
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2175
      Left            =   9960
      TabIndex        =   7
      Top             =   -1920
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   480
      Top             =   8280
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   2400
      Top             =   8520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   735
      Left            =   4920
      TabIndex        =   6
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Text            =   "»Ì«‰«  «·›—Ê⁄"
      ColorGeneral    =   0
      ColorGeneral    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcemployee 
      DataField       =   "manger_id"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1200
      TabIndex        =   42
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ListField       =   ""
      BoundColumn     =   "Account_Code"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   4560
      Top             =   8280
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin ALLButtonS.ALLButton ALLButton3 
      Height          =   375
      Left            =   8400
      TabIndex        =   68
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«»«  «·–„„"
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
      MICON           =   "baranchese.frx":0426
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton4 
      Height          =   375
      Left            =   6720
      TabIndex        =   88
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«»«  «·«’Ê·"
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
      MICON           =   "baranchese.frx":0442
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton5 
      Height          =   375
      Left            =   5040
      TabIndex        =   89
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«»«  «·„‘«—Ì⁄"
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
      MICON           =   "baranchese.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   480
      TabIndex        =   115
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«·—»ÿ »«·„Ã„Ê⁄« "
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
      MICON           =   "baranchese.frx":047A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton6 
      Height          =   375
      Left            =   3360
      TabIndex        =   116
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«»«  „ ‰Ê⁄Â"
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
      MICON           =   "baranchese.frx":0496
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton7 
      Height          =   375
      Left            =   1800
      TabIndex        =   134
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«»«  «·«‰ «Ã"
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
      MICON           =   "baranchese.frx":04B2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "„œÌ— «·›—⁄"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3960
      TabIndex        =   41
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄‰Ê«‰ «·›—⁄"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8640
      TabIndex        =   40
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Eng Name"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   750
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·›—⁄"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   8640
      TabIndex        =   30
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   " ·Ì›Ê‰ «·›—⁄"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   11400
      TabIndex        =   31
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·›—⁄"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   11400
      TabIndex        =   29
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "baranchesE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag_mode As String

Private Sub ALLButton1_Click()
    Dim rsOut As New ADODB.Recordset
    Dim Current_case As Integer
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!opt_group = True And rsOut!opt_inv_and_branch_create_account = 1 Then
   
        Else

            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "to done this process change option in system manger", vbCritical: Exit Sub
            Else
                MsgBox "·« Ì„ﬂ‰ « „«„ Â–… «·⁄„·Ì… ·«‰ﬂ «Œ —  —»ÿ «·„Œ«“‰ »«·„Ã„Ê⁄«  ›ﬁÿ ›Ì „œÌ— «·‰Ÿ«„", vbCritical: Exit Sub
            End If
        End If
    End If

    If TXTCode.text = "" Then Exit Sub

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    sql = "Select * from groups_account_in_inventory where branch_id='" & TXTCode & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount > 0 Then
        If SystemOptions.UserInterface = EnglishInterface Then
            MsgBox "This Branch Already linked With groups", vbCritical: Exit Sub
        Else
            MsgBox " „ —»ÿ Â–« «·›—⁄ »«·„Ã„Ê⁄«  „‰ ﬁ»·", vbCritical: Exit Sub
        End If
    End If

    Rs3.Close
    sql = "Select * from Groups "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub

    For i = 1 To Rs3.RecordCount

        If create_Branch_group(TXTCode.text, Rs3("GroupID").value, Rs3("GroupName").value) = True Then
        End If

        Rs3.MoveNext
    Next i

    Rs3.Close

    If SystemOptions.UserInterface = EnglishInterface Then
        MsgBox "Linked Done", vbInformation
    Else
        MsgBox " „ «·—»ÿ", vbInformation
    End If

End Sub

Function hide_all_frame()
    Frame10.Visible = False
    Frame11.Visible = False
    Frame12.Visible = False
    Frame13.Visible = False
    Frame14.Visible = False

    Frame9.Visible = False

End Function

Private Sub ALLButton2_Click()
    hide_all_frame
    Frame10.Visible = True

End Sub

Private Sub ALLButton3_Click()
    hide_all_frame
    Frame11.Visible = True
End Sub

Private Sub ALLButton4_Click()
    hide_all_frame
    Frame12.Visible = True
End Sub

Private Sub ALLButton5_Click()
    hide_all_frame
    Frame13.Visible = True
End Sub

Private Sub ALLButton6_Click()
    hide_all_frame
    Frame14.Visible = True
End Sub

Private Sub ALLButton7_Click()
    hide_all_frame
    Frame9.Visible = True
End Sub

Private Sub CMD_language_Click()
    On Error Resume Next

    If CMD_language.Caption = "EN" Then
        my_language = "E"
 
        'Call Reload(Me)
 
    Else
        my_language = "A"
 
        'Call Reload(Me)
    End If

End Sub

Function create_accounts() As Boolean
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
    Dim i As Integer
    Dim StrNewAccountCode As String
    Dim namea As String
    Dim namee As String
    Dim currency_code As String
    Dim mowazna As Boolean
    Dim cost_center As Boolean
    Set rs = New ADODB.Recordset
    Set Rs1 = New ADODB.Recordset

    rs.Open "Select * from ACCOUNTS where Sum_account=1 ", Cn, adOpenStatic, adLockOptimistic, adCmdText

    If SystemOptions.UserInterface = EnglishInterface Then
        If rs.RecordCount = 0 Then MsgBox "·«»œ „‰  ⁄—Ì› Õ”«»«   Ã„Ì⁄Ì… «Ê·« ›Ì œ·»· «·Õ”«»« ", vbCritical, "": create_accounts = False: Exit Function
    Else

        If rs.RecordCount = 0 Then MsgBox "Must define Summary Accounts first ", vbCritical, "": create_accounts = False: Exit Function
    End If

    rs.MoveFirst
 
    Rs1.Open "ACCOUNTS", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    For i = 1 To rs.RecordCount
        namea = rs("Account_Name").value & "  ›—⁄   " & txtnameA.text
        namee = rs("Account_NameEng").value & " " & txtnamee.text & "  Branch"
        currency_code = IIf(IsNull(rs("currenct_code").value), 1, rs("currenct_code").value)
        mowazna = IIf(IsNull(rs("mowazna").value), 0, rs("mowazna").value)
        cost_center = IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)

        StrNewAccountCode = ModAccounts.AddNewAccount(rs("Account_Code").value, namea, 0, False, namee, currency_code, mowazna, cost_center, False, TXTCode.text)
        rs.MoveNext
    Next i

    If SystemOptions.UserInterface = EnglishInterface Then
        MsgBox "Branch Created With Accounts", vbInformation, ""
 
    Else
        MsgBox " „ «‰‘«¡ «·›—⁄ ÊÕ”«»« …", vbInformation, ""
    End If

    create_accounts = True
End Function

Private Sub Command1_Click(Index As Integer)
    On Error Resume Next

    If Index = 0 Then
        Adodc1.Recordset.AddNew
        TXTCode.text = CStr(new_id("branches", "branch_id", "", True))
        flag_mode = "N"

    Else

        If Index = 1 Then
     
            If txtnamee.text = "" Then MsgBox "write  branch name first", vbCritical: Exit Sub
      
            If txtnameA.text = "" Then MsgBox "«ﬂ »  «”„ «·›—⁄ «Ê·«  ", vbCritical: Exit Sub
    
            Adodc1.Recordset.Fields!inventory = DataCombo2.text
            Adodc1.Recordset.update
            Adodc1.Recordset.MoveLast
   
            If flag_mode = "N" Then
   
                If create_accounts = False Then
                    Exit Sub
                End If

                flag_mode = "E"
     
            End If
 
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "Saved", vbInformation, ""
            Else
                MsgBox " „ «·Õ›Ÿ", vbInformation, ""
            End If
  
        Else

            If Index = 2 Then
 
                Dim x As Integer

                If my_language = "E" Then
                    x = MsgBox("Confirm delete", vbCritical + vbYesNo)
                Else
                    x = MsgBox("Â· «‰  „ √ﬂœ „‰ «·Õ–›", vbCritical + vbYesNo)
              
                End If

                If x = vbNo Then
                    Exit Sub
                End If

                If Adodc1.Recordset.RecordCount > 0 Then
                    Adodc1.Recordset.delete
                    Adodc1.Refresh
                Else

                    If my_language = "E" Then
                        MsgBox "No Departement to delete", vbCritical
                    Else
                        MsgBox "·« ÌÊÃœ „« Ì„ﬂ‰ Õ–›…", vbCritical
                    End If
                
                End If

                Exit Sub

            End If
        End If
    End If

End Sub

Private Sub DataCombo1_Click(Index As Integer, _
                             Area As Integer)
    On Error Resume Next

    'Adodc2.Refresh
    'DataCombo1(Index).ReFill
End Sub

Private Sub DataCombo1_KeyUp(Index As Integer, _
                             KeyCode As Integer, _
                             Shift As Integer)
 
    'On Error Resume Next

    If KeyCode = vbKeyF3 Then
        Account_search.Show

        If Index = 37 Or Index = 38 Or Index = 39 Or Index = 19 Or Index = 18 Or Index = 22 Or Index = 21 Or Index = 23 Or Index = 41 Or Index = 16 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 4 Or Index = 5 Or Index = 12 Or Index = 13 Then 'Õ”«»«  ‰Â«∆Ì… ›ﬁÿ
            Account_search.case_id = 1700
        Else
            Account_search.case_id = 700
        End If

        Account_search.case_index = Index
    End If

    If KeyCode = vbKeyF6 Then
        account_index.Show
    End If

    If KeyCode = vbKeyF5 Then
        Adodc2.Refresh
        DataCombo1(Index).ReFill
    End If

End Sub

'Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF6 Then
'frmcities.Show
'End If
'End Sub

Private Sub Form_Activate()

End Sub

Private Sub Form_Load()
    Dim My_SQL As String

    My_SQL = "  select Emp_ID,Emp_Name from TblEmployee   "
    fill_combo DCEmployee, My_SQL 'On Error Resume Next

    hide_all_frame
    Frame10.Visible = True

    If my_language = "E" Then
        CMD_language.ToolTipText = "change Language"
 
        Me.dept_lbl = departement_name
        Me.emp_name_lbl = current_user_name
        InfoE.Visible = True
        infoA.Visible = False
    Else

        emp_a.Caption = current_user_name
        dep_a.Caption = departement_name
   
        infoA.Visible = True
        InfoE.Visible = False
    End If

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    '  On Error Resume Next
    If SystemOptions.UserInterface = EnglishInterface Then
    
        ALLButton7.Caption = "Production Account"
        Frame9.Caption = ALLButton7.Caption
        Labelx(37).Caption = "Production Account-Material"
        Labelx(38).Caption = "Production Account-Expenses"
        Labelx(39).Caption = "Production Account-F. Expenses"
    
        Label1.Caption = "Branch NO"
        Label2.Caption = "Branch Name"
        Label4.Caption = "Branch Tel"
        '    Label3.Caption = "Basic Store"
        Label14.Caption = "Address"
        Label15.Caption = "Manger"
    
        Labelx(0).Caption = "Store Account"
        Labelx(10).Caption = "Damage Account"
        Labelx(11).Caption = "Inventory adjustment"
     
        Labelx(1).Caption = "Sale cost Account"
     
        Labelx(2).Caption = "Sale Account"
      
        Labelx(3).Caption = "sale return Account"
        Labelx(4).Caption = "Purchase Account"
        Labelx(5).Caption = "purchase return Account"
       
        Labelx(6).Caption = "Box Account"
        Labelx(20).Caption = "Banks Account "
        Labelx(19).Caption = "Opening Balance "
        Labelx(7).Caption = "staff Accounts "
        Labelx(29).Caption = "Due salaries Acc."
        Labelx(16).Caption = "salaries"
            
        Labelx(30).Caption = "Apportionment"
        
        Frame10.Caption = "Store Accounts"
        Frame11.Caption = "receivables Accounts "
        Frame12.Caption = "Assets Accounts "
        Frame13.Caption = "Projects Accounts "
        Frame14.Caption = "Another Accounts "
        ALLButton2.Caption = "Store Accounts"
        ALLButton3.Caption = "receivables Accounts "
        ALLButton4.Caption = "Assets Accounts"
        ALLButton5.Caption = "Projects Accounts"
        ALLButton6.Caption = "Another Accounts"
         
        Labelx(8).Caption = "Customer Account"
        Labelx(9).Caption = "Vendor Account"
           
        Labelx(12).Caption = "Allowed discount"
        Labelx(13).Caption = "Unearned discount"
           
        Labelx(21).Caption = "Increase in cash "
        Labelx(22).Caption = "Shortfall in cash "
        Labelx(24).Caption = "Assets Account "
        Labelx(25).Caption = "Depreciation expense "
        Labelx(26).Caption = "Accumu. depreciation"
        
        Labelx(14).Caption = "Project Expanses"
        Labelx(15).Caption = "Projects Revenu"
        Labelx(27).Caption = "Project Materials"
        Labelx(28).Caption = "Projects salaries"
    
        Labelx(23).Caption = "Service revenue "
        Labelx(17).Caption = "Gifts and  Samples "
        Labelx(18).Caption = "Currency differences "
      
        ' TabControl1.Item(0).Caption = "Inventory"
         
        ALLButton1.Caption = "Link With Group"
        SetInterface Me
        Labelx(31).Caption = "Fixed Asset"
        Labelx(32).Caption = "Legal Accounts"

        Me.left = (mdifrmmain.Width - Me.Width) / 2
        Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

        CMD_language.Caption = "⁄—»Ì"
 
        ' Text1.Alignment = 0
        '  Text2.Alignment = 0
        ' DataCombo1.RightToLeft = False
  
        Frame2.Visible = False
        Frame3.Visible = True
        SuperLabel1.text = "Branches Data"
        Me.Caption = SuperLabel1.text
        Command1(0).Caption = "new"
        Command1(1).Caption = "save"
        Command1(2).Caption = "delete"
        Adodc1.Caption = "move"
        Frame8.Caption = "Projects"
        Frame4.Caption = "Stores"
        Frame5.Caption = "Employees"
        Frame6.Caption = "Colors"
        Label8.Caption = "Last Account"
        Label13.Caption = "Master account"
        Labelx(33).Caption = "Expenses"
        Labelx(34).Caption = "Revenues"
        Labelx(36).Caption = "Sub-contractor"
  
    End If

    connection_string = Cn.ConnectionString

    'Adodc5.ConnectionString = connection_string
    ' Adodc5.CommandType = adCmdText
    'Adodc5.RecordSource = "select * from cities where not(city_name is null) "
    'Adodc5.Refresh
    '

    'where  NOT (branch_name='')

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select *  from ACCOUNTS WHERE last_account=1" '
    Adodc2.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText

    Dim rsOut As New ADODB.Recordset
    Dim Msg As String
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!opt_group = True Then
            Adodc4.RecordSource = "select *  from ACCOUNTS WHERE last_account=0" '
            recolor
        Else
            Adodc4.RecordSource = "select *  from ACCOUNTS WHERE last_account=1" '
        End If
    End If

    Adodc4.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select *  from ACCOUNTS WHERE last_account=0" '
    Adodc5.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from  TblStore" '  where branch_no=" & branch_no
    Adodc3.Refresh

    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from   branches   " ' where departement_no=0"
    Adodc1.Refresh

End Sub

Function recolor()
    Labelx(1).ForeColor = &HFF&
    Labelx(2).ForeColor = &HFF&
    Labelx(3).ForeColor = &HFF&
    Labelx(4).ForeColor = &HFF&
    Labelx(5).ForeColor = &HFF&
    Labelx(17).ForeColor = &HFF&
    Labelx(12).ForeColor = &HFF&
    Labelx(13).ForeColor = &HFF&

End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
      
End Sub
 
Private Sub txtnameA_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtnamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
