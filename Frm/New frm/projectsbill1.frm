VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{784C0C13-85E7-4E11-A8FB-F0243A135D03}#2.0#0"; "SuperLablel.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form projectsbill1 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6825
   ClientLeft      =   3525
   ClientTop       =   1470
   ClientWidth     =   13800
   Icon            =   "projectsbill1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   13800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox note_id 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "note_id"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   0
      TabIndex        =   83
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtrevenue_account 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "revenue_account"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   0
      TabIndex        =   82
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox billto 
      DataField       =   "bill_to"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "projectsbill1.frx":000C
      Left            =   9720
      List            =   "projectsbill1.frx":0016
      RightToLeft     =   -1  'True
      TabIndex        =   79
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox bill_Type 
      DataField       =   "bill_type"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "projectsbill1.frx":0034
      Left            =   3480
      List            =   "projectsbill1.frx":003E
      TabIndex        =   78
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox txtsubaccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "Sub_user_account"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   0
      TabIndex        =   75
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtendaccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "End_user_account"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   120
      TabIndex        =   74
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "total"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   4320
      TabIndex        =   69
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   3360
      TabIndex        =   67
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   4320
      TabIndex        =   65
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtdate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "bill_date"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5640
      TabIndex        =   62
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   360
      Left            =   9720
      TabIndex        =   60
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5280
      TabIndex        =   55
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   6720
      TabIndex        =   42
      Top             =   6120
      Width           =   4935
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   43
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "projectsbill1.frx":0051
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
         Left            =   -1560
         TabIndex        =   44
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "«·„—›ﬁ« "
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
         BCOL            =   255
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "projectsbill1.frx":006D
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
         Left            =   3600
         TabIndex        =   45
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "projectsbill1.frx":0089
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
         Index           =   5
         Left            =   -1560
         TabIndex        =   47
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "»ÕÀ"
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
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "projectsbill1.frx":00A5
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
         Index           =   3
         Left            =   2520
         TabIndex        =   84
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   " ⁄œÌ·"
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
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "projectsbill1.frx":00C1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Caption         =   "Label2"
         Height          =   15
         Left            =   -120
         TabIndex        =   46
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   2280
      TabIndex        =   37
      Top             =   9120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2880
         TabIndex        =   41
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4440
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   120
      TabIndex        =   32
      Top             =   9120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   35
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   4095
      Left            =   6720
      TabIndex        =   28
      Top             =   9000
      Width           =   1455
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   0
         TabIndex        =   29
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
         Begin ALLButtonS.ALLButton Command1 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "»«·—ﬁ„"
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
            MICON           =   "projectsbill1.frx":00DD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SuperLablel.SuperLabel SuperLabel2 
            Height          =   615
            Left            =   240
            TabIndex        =   31
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1085
            Text            =   "»ÕÀ"
            ColorGeneral    =   16711680
            ColorGeneral    =   16711680
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
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   8400
      Width           =   6495
      Begin VB.Line Line1 
         Index           =   7
         X1              =   960
         X2              =   960
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„—ﬂ“"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„—ﬂ“"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5040
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   4460
         X2              =   4460
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·„—ﬂ“"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   1440
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·„” ÊÏ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   6180
         X2              =   6180
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   2700
         X2              =   2700
         Y1              =   120
         Y2              =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   9600
      Width           =   6495
      Begin VB.Line Line1 
         Index           =   3
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   5520
         X2              =   5520
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Center #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Center Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3780
         X2              =   3780
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Center Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   360
         X2              =   360
         Y1              =   120
         Y2              =   720
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   0
      TabIndex        =   12
      Top             =   8520
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label12 
         Caption         =   "Major Center"
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
         TabIndex        =   27
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label13 
         Caption         =   "Center Type"
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
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Center Name"
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
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Center#"
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
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   11400
      TabIndex        =   11
      Top             =   9000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   8760
      Visible         =   0   'False
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
         Caption         =   "M15"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox txtprojectname 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "project_name"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   3480
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   8520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      DataField       =   "last_root"
      DataSource      =   "Adodc5"
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text5"
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "projectsbill1.frx":00F9
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   8040
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "account_no"
         Caption         =   "—ﬁ„ «·„‘—Ê⁄"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "account_name"
         Caption         =   "«”„ «·„‘—Ê⁄"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "account_type"
         Caption         =   "‰Ê⁄ «·„‘—Ê⁄"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "parent_no"
         Caption         =   "parent_no"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "child_no"
         Caption         =   "child_no"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "level"
         Caption         =   "«·„” ÊÏ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4500.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   570
      Left            =   -240
      Top             =   9960
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1005
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
      Caption         =   " Õ—Ìﬂ"
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
      Height          =   330
      Left            =   10320
      Top             =   10080
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
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
      Caption         =   " Õ—Ìﬂ"
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
      Height          =   330
      Left            =   10320
      Top             =   10440
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
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
      Caption         =   " Õ—Ìﬂ"
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
      Height          =   330
      Left            =   10320
      Top             =   10800
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
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
      Caption         =   " Õ—Ìﬂ"
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
      Height          =   330
      Left            =   10320
      Top             =   11040
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
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
      Caption         =   " Õ—Ìﬂ"
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
      Height          =   330
      Left            =   10320
      Top             =   8400
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
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
      Caption         =   " Õ—Ìﬂ"
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
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   21
      ToolTipText     =   "Language  «··€…"
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
      MICON           =   "projectsbill1.frx":010E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   0
      Left            =   1665
      TabIndex        =   48
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
      ButtonImage     =   "projectsbill1.frx":012A
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
      Left            =   600
      TabIndex        =   49
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
      ButtonImage     =   "projectsbill1.frx":04C4
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
      Left            =   2190
      TabIndex        =   50
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
      ButtonStyle     =   1
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
      ButtonImage     =   "projectsbill1.frx":085E
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      Alignment       =   0
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      RightToLeft     =   -1  'True
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   3
      Left            =   1125
      TabIndex        =   51
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
      ButtonImage     =   "projectsbill1.frx":0BF8
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "projectsbill1.frx":0F92
      Height          =   2895
      Left            =   1920
      TabIndex        =   57
      Top             =   3120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "item"
         Caption         =   "«·»‰œ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cost"
         Caption         =   " ﬂ·›… «·»‰œ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "exe"
         Caption         =   " ﬂ·›… «·„‰›–"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "percentage"
         Caption         =   "‰”»… «· ‰›Ì–"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "exedate"
         Caption         =   " «—ÌŒ «·«‰ Â«¡"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   -1  'True
            ColumnWidth     =   7004.977
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            WrapText        =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1305.071
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   4080
      Top             =   8040
      Width           =   7695
      _ExtentX        =   13573
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
      Caption         =   "Adodc2"
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
      DataField       =   "project_no"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   9720
      TabIndex        =   58
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "account_no"
      BoundColumn     =   "Fullcode"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   6240
      TabIndex        =   63
      Top             =   2640
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "account_no"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcAccount2 
      DataField       =   "End_user_name"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   9720
      TabIndex        =   76
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BoundColumn     =   "End_user_name"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcAccount1 
      DataField       =   "Sub_user_name"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3480
      TabIndex        =   77
      Top             =   1440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BoundColumn     =   "End_user_name"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   360
      Left            =   1920
      TabIndex        =   80
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      Format          =   91095041
      CurrentDate     =   38784
   End
   Begin ALLButtonS.ALLButton Command2 
      Height          =   375
      Left            =   720
      TabIndex        =   81
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«œ—«Ã"
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
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "projectsbill1.frx":0FA7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ „ﬁ«Ê· «·»«ÿ‰"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7200
      TabIndex        =   73
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      Caption         =   "‰Ê⁄ «·„” Œ·’"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7680
      TabIndex        =   72
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   " «—ÌŒ «·«‰ Â«¡"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      TabIndex        =   71
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Caption         =   "«·«Ã„«·Ì"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5520
      TabIndex        =   70
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Caption         =   "‰”»… «· ‰›Ì–"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      TabIndex        =   68
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Caption         =   " ﬂ·›… «·„‰›–"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4320
      TabIndex        =   66
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "«·›« Ê—… «·Ï"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12600
      TabIndex        =   64
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      Caption         =   "«· «—ÌŒ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7680
      TabIndex        =   61
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·›« Ê—…"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12360
      TabIndex        =   59
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   " ﬂ·›… «·»‰œ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5400
      TabIndex        =   56
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "«·»‰Êœ «·„‰ÂÌ…"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12480
      TabIndex        =   54
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11760
      TabIndex        =   53
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄„Ì· «·‰Â«∆Ì"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12360
      TabIndex        =   52
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "    ›« Ê—… „‘—Ê⁄"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·„‘—Ê⁄"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7440
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   11880
      TabIndex        =   5
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·„‘—Ê⁄"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12480
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "projectsbill1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As String
Dim last_root As Integer
Dim last_geeral As Integer
Dim last_branch As Integer
Dim mod_flad As String
Dim first_run  As Boolean




Private Sub CMD_language_Click()
On Error Resume Next

If CMD_language.Caption = "EN" Then
my_language = "E"
 
''Call Reload(Me)

 
Else
my_language = "A"
 
''Call Reload(Me)
End If
End Sub
Function SaveData()
Dim accountdep As String
  total.text = gettotal(txtid.text)
  
If note_id.text = "" Then
Set Rs = New ADODB.Recordset
StrSQL = "select * From Notes where NoteType=5000 order by NoteID"
Rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
note_id.text = CStr(new_id("Notes", "NoteID", "", True))

    
        note_id.text = CStr(new_id("Notes", "NoteID", "", True))
       ' Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=5"))
        Rs.AddNew
        Rs("NoteID").value = Val(note_id.text)
   
  '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))

    Rs("Note_Value").value = IIf(total.text = "", Null, Val(total.text))
   '' Rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
   ' Rs("Remark").value = IIf(DCPROJECT.BoundText = "", "", Trim(DCPROJECT.BoundText))

    Rs("NoteType").value = 500
    Rs("NoteDate").value = DateValue(Now)
    Rs("UserID").value = user_id
    Rs.update
Else
  StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(note_id.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
 End If




        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
        Set RsDev = New ADODB.Recordset
        RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        If billto.ListIndex = -1 Then MsgBox "Specify bill to", vbCritical: Exit Function
        If billto.ListIndex = 0 Then
        accountdep = txtendaccount.text
        Else
        If billto.ListIndex = 1 Then
        accountdep = txtsubaccount.text
        End If
        End If
        
        If accountdep = "" Then GoTo ll
        '«·ÿ—› «·„œÌ‰
        RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 1
            RsDev("Account_Code").value = accountdep
            RsDev("Value").value = Val(Me.total.text)
            RsDev("Credit_Or_Debit").value = 0
           ' RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
             RsDev("Double_Entry_Vouchers_Description").value = "P" & TXTprojectname.text
            RsDev("Notes_ID").value = Val(note_id.text)
            RsDev("project_bill_no").value = Val(txtid.text)
            
            RsDev("RecordDate").value = DateValue(Now)
            RsDev("UserID").value = user_id
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev.update
ll:
        '«·ÿ—› «·œ«∆‰
        If Me.txtrevenue_account.text = "" Then Exit Function
        RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 2
            RsDev("Account_Code").value = Me.txtrevenue_account.text
            RsDev("Value").value = Val(Me.total.text)
            RsDev("Credit_Or_Debit").value = 1
            'RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
            RsDev("Double_Entry_Vouchers_Description").value = "P" & TXTprojectname.text
            RsDev("Notes_ID").value = Val(note_id.text)
            RsDev("project_bill_no").value = Val(txtid.text)
            RsDev("RecordDate").value = DateValue(Now)
            RsDev("UserID").value = user_id
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev.update




End Function
Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Select Case Index

Case 0
'mod_flad = "N"
'Adodc7.ConnectionString = connection_string
'Adodc7.CommandType = adCmdText
'Adodc7.RecordSource = "select * from projects_des where project_no='0'"
'Adodc7.Refresh

'Command1(1).Enabled = True
'Command1(0).Enabled = False
    Adodc1.Recordset.AddNew
 txtid.text = CStr(new_id("project_billl", "id", "", True))
  txtdate.text = DateValue(Now)
            
        'Adodc1.Recordset.Fields!account_name = ""
'Adodc1.Recordset.update
'Adodc1.Recordset.MoveLast
    
Case 1
If billto.ListIndex = -1 Then MsgBox "Õœœ «·›« Ê—… «·Ï «Ê·«", vbCritical: Exit Sub
If DcAccount1.text = "" And billto.ListIndex = 1 Then MsgBox "·«Ì„ﬂ‰ Õ›Ÿ «·›« Ê—… ·«‰ﬂ «Œ —  „ﬁ«Ê· »«ÿ‰ Ê«·„‘—Ê⁄ ·Ì” ·Â „ﬁ«Ê· »«ÿ‰", vbCritical: Exit Sub
 SaveData
 ''Adodc1.Recordset.Fields!  project_no = DataCombo2.text
  Adodc1.Recordset.update

    Exit Sub
ErrTrap:
MsgBox "error"
Case 2

'x = MsgBox("Â· «‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbCritical + vbYesNo)
'If x = vbNo Then
'Exit Sub
'End If

'    If Adodc1.Recordset.RecordCount > 0 Then
'    Adodc1.Recordset.Delete
'    Adodc1.Refresh
'    DataGrid1.Refresh
'    End If

On Error Resume Next
If SystemOptions.UserInterface = EnglishInterface Then
If Text1.text = "" Then MsgBox "Select Project firstly": Exit Sub

Else
If Text1.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „‘—Ê⁄ «Ê·«": Exit Sub

End If
 

 

imaged.Show
If SystemOptions.UserInterface = EnglishInterface Then

imaged.Label9.Caption = "Attachment For Project "
imaged.Caption = "Project Attachment  "
imaged.Label6.Caption = "   Project NO"
Label5.Caption = "Documents"
Label8.Caption = "Forms"

Else

imaged.Label9.Caption = "„—›ﬁ«    „‘—Ê⁄  —ﬁ„"
imaged.Caption = "„—›ﬁ«  „‘—Ê⁄  "
imaged.Label6.Caption = "—ﬁ„  «·„‘—Ê⁄"

End If
imaged.SUBJECT_NO = Text1.text
imaged.txtopeation_type = "„—›ﬁ«  „‘—Ê⁄"

imaged.Adodc1.CommandType = adCmdText
imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  „‘—Ê⁄' and subject_no='" & Text1.text & "'"
imaged.Adodc1.Refresh
If imaged.Adodc1.Recordset.RecordCount > 0 Then

imaged.DBPix201.Visible = True
Else
imaged.DBPix201.Visible = False
End If


Case 3
    If Adodc1.Recordset.RecordCount > 0 Then
    
    Form3.case_id = Me.name
   
    Form3.Show
    End If

Case 4
On Error Resume Next
X = InputBox("«œŒ· «·—ﬁ„ «·„ÿ·Ê» «·»ÕÀ ⁄‰…")

      
        Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from  projects where not (account_no is null)  and account_no='" & X & "'"
        Adodc1.Refresh
         

Case 5
    X = InputBox("«œŒ· ﬂ·„… «·»ÕÀ")
            Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from projects where not (account_no is null)  and account_name like '%" & X & "%'"
        Adodc1.Refresh


End Select

End Sub


Private Sub Combo1_Click()
 

End Sub

 

 

Private Sub Command2_Click()
On Error Resume Next
If SystemOptions.UserInterface = EnglishInterface Then

If txtid.text = "" Then MsgBox "Select bill first": Exit Sub
Else
If txtid.text = "" Then MsgBox "Õœœ «·›« Ê—… «·Ï": Exit Sub

End If

 Command1_Click (1)

Adodc7.Recordset.AddNew
Adodc7.Recordset.Fields!bill_id = txtid.text
Adodc7.Recordset.Fields!project_no = DataCombo2.BoundText
Adodc7.Recordset.Fields!Item = DataCombo5.text
Adodc7.Recordset.Fields!cost = Val(Text6.text)
Adodc7.Recordset.Fields!exe = Text9.text
Adodc7.Recordset.Fields!percentage = (Val(Text9.text) / Val(Text6.text)) * 100
 Adodc7.Recordset.Fields!exedate = DTPicker2.value
 
Adodc7.Recordset.update
'Text3.text = ""
Text6.text = ""
Adodc7.ConnectionString = connection_string
Adodc7.CommandType = adCmdText
Adodc7.RecordSource = "select * from project_bill_details where bill_id=" & Val(txtid.text)
Adodc7.Refresh
DataGrid2.Refresh

DataGrid2.Refresh
total.text = gettotal(txtid.text)
Command1_Click (1)
End Sub

Private Sub DataCombo2_Change()
On Error Resume Next
If DataCombo2.text = "" Then Exit Sub
Dim My_SQL As String

My_SQL = "select * from projects where id =" & DataCombo2.BoundText
 
Set rec = New ADODB.Recordset
rec.CursorLocation = adUseClient

rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 TXTprojectname.text = rec.Fields("Project_name").value
' txtprojecttype.text = rec.Fields("Contract_type").value
 
 
txtsubaccount.text = IIf(IsNull(rec.Fields("sub_contractor_Account").value), "", rec.Fields("sub_contractor_Account").value)

 DcAccount1.text = IIf(IsNull(rec.Fields("sub_contractor_name").value), "", rec.Fields("sub_contractor_name").value)
 
 
txtendaccount.text = IIf(IsNull(rec.Fields("End_user_Account").value), "", rec.Fields("End_user_Account").value)
DcAccount2.text = IIf(IsNull(rec.Fields("End_user_name").value), "", rec.Fields("End_user_name").value)

txtrevenue_account.text = IIf(IsNull(rec.Fields("REVENUE_account").value), "", rec.Fields("REVENUE_account").value)


My_SQL = "  select net,des from projects_des  where project_id='" & DataCombo2.BoundText & "'"
fill_combo DataCombo5, My_SQL
End Sub

Private Sub DataCombo2_Click(Area As Integer)
On Error Resume Next
If DataCombo2.text = "" Then Exit Sub
Dim My_SQL As String

My_SQL = "select * from projects where id =" & DataCombo2.BoundText
 
Set rec = New ADODB.Recordset
rec.CursorLocation = adUseClient

rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 TXTprojectname.text = rec.Fields("Project_name").value
 txtprojecttype.text = rec.Fields("Contract_type").value
 
 
txtsubaccount.text = IIf(IsNull(rec.Fields("sub_contractor_Account").value), "", rec.Fields("sub_contractor_Account").value)

 DcAccount1.text = IIf(IsNull(rec.Fields("sub_contractor_name").value), "", rec.Fields("sub_contractor_name").value)
 
 
txtendaccount.text = IIf(IsNull(rec.Fields("End_user_Account").value), "", rec.Fields("End_user_Account").value)
DcAccount2.text = IIf(IsNull(rec.Fields("End_user_name").value), "", rec.Fields("End_user_name").value)

txtrevenue_account.text = IIf(IsNull(rec.Fields("REVENUE_account").value), "", rec.Fields("REVENUE_account").value)


My_SQL = "  select net,des from projects_des  where project_id='" & DataCombo2.BoundText & "'"
fill_combo DataCombo5, My_SQL



End Sub

Private Sub DataCombo5_Click(Area As Integer)
If DataCombo5.BoundText <> "" Then
Text6.text = DataCombo5.BoundText
Text9.text = ""
Else
DataCombo5 = ""
End If
End Sub

Private Sub DataGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
If Adodc7.Recordset.RecordCount > 0 Then
Adodc7.Recordset.Delete
DataGrid2.Refresh
Command1_Click (1)
total.text = gettotal(txtid.text)



End If

End If
End Sub

Private Sub DcAccount1_Click(Area As Integer)
'On Error Resume Next
'If DcAccount1.text = "" Then Exit Sub
'Dim My_SQL As String
'
'My_SQL = "select Account_Name from ACCOUNTS where Account_Serial='" & DcAccount1.text & "'"
'
'Set rec = New ADODB.Recordset
'rec.CursorLocation = adUseClient
'
'rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
' DcAccount2.text = rec.Fields("Account_Name").value
End Sub

Private Sub DcAccount2_Change()
'On Error Resume Next
'If DcAccount2.text = "" Then Exit Sub
'Dim My_SQL As String

'My_SQL = "select Account_Serial from ACCOUNTS where Account_Name ='" & DcAccount2.text & "'"
 
'Set rec = New ADODB.Recordset
'rec.CursorLocation = adUseClient
'
'rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
' DcAccount1.text = rec.Fields("Account_Serial").value
'DcAccount2.SetFocus
'SendKeys "{f4}"
End Sub

Private Sub DcAccount2_Click(Area As Integer)
'On Error Resume Next
'If DcAccount2.text = "" Then Exit Sub
'Dim My_SQL As String

'My_SQL = "select Account_Serial from ACCOUNTS where Account_Name ='" & DcAccount2.text & "'"
 
'Set rec = New ADODB.Recordset
'rec.CursorLocation = adUseClient
'
'rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
' DcAccount1.text = rec.Fields("Account_Serial").value
End Sub

Function gettotal(X As String) As Double
Dim My_SQL As String

My_SQL = "  select Sum(exe) as total  from project_bill_details where bill_id=" & X

Set rec = New ADODB.Recordset
rec.CursorLocation = adUseClient

rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
gettotal = IIf(IsNull(rec.Fields("total").value), 0, rec.Fields("total").value)



End Function
Private Sub Form_Activate()
'On Error Resume Next
If first_run = True Then
first_run = False
connection_string = Cn.ConnectionString
Adodc7.ConnectionString = connection_string

If txtid.text <> "" Then
connection_string = Cn.ConnectionString
Adodc7.ConnectionString = connection_string
Adodc7.CommandType = adCmdText
Adodc7.RecordSource = "select * from project_bill_details where bill_id=" & txtid.text
Adodc7.Refresh
DataGrid2.Refresh
End If
DTPicker2.value = DateValue(Now)
End If




End Sub

Private Sub Form_Load()
'On Error Resume Next
   '

first_run = True
Dim My_SQL As String

'My_SQL = "  select Account_code,Account_Serial from ACCOUNTS  where last_account=1"

'fill_combo DcAccount1, My_SQL

My_SQL = "  select CusName,CusName from TblCustemers  "
fill_combo DcAccount1, My_SQL

My_SQL = "  select CusName,CusName from TblCustemers  "
fill_combo DcAccount2, My_SQL

My_SQL = "  select id,Fullcode from Projects"
fill_combo DataCombo2, My_SQL


If my_language = "E" Then
CMD_language.ToolTipText = "Change Language"

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

 Me.left = (MDIFrmMain.Width - Me.Width) / 2
    Me.top = (MDIFrmMain.Height - Me.Height) / 2 - 500
    
'LoadSettings

 


If SystemOptions.UserInterface = EnglishInterface Then
Temp = XPBtnMove(1).left
XPBtnMove(1).left = XPBtnMove(2).left
XPBtnMove(2).left = Temp

Temp = XPBtnMove(0).left
XPBtnMove(0).left = XPBtnMove(3).left
XPBtnMove(3).left = Temp
    SetInterface Me
Me.Caption = "Project Invoice"
Label9.Caption = Me.Caption

Label20.Caption = "Bill No."
Label25.Caption = "Date"

Label6.Caption = "Project Code"
Label1.Caption = "Project Name"
Label15.Caption = "End User"
Label23.Caption = "Sub-Contractor"
Label18.Caption = "Bill To"
Label30.Caption = "Bill Type"
Label17.Caption = "Finished Item"
Label19.Caption = "Item cost"
Label27.Caption = "Exe cost"
Label28.Caption = "Percentage"
Label5.Caption = "End date"

DataGrid2.Columns(0).Caption = "Item"
DataGrid2.Columns(1).Caption = "item Cost"
DataGrid2.Columns(2).Caption = "Exe Cose"
DataGrid2.Columns(3).Caption = "percentage"
DataGrid2.Columns(4).Caption = "End date"
 


 
DataGrid2.RightToLeft = False
 
Command2.Caption = "Insert "

 Label29.Caption = "Total"

 
 
  
 DataGrid1.RightToLeft = False
 CMD_language.Caption = "⁄—»Ì"
  Frame4.Visible = True
  Frame3.Visible = True
    Frame8.Visible = True
 
  
  Command1(0).Caption = "new"
  Command1(1).Caption = "save"
  Command1(2).Caption = "Attachments"
  SuperLabel2.text = "Search"
  Command1(4).Caption = "By ID"
  Command1(5).Caption = "Search"
  
  Adodc1.Caption = "move"
 ' Me.Width = 10000
 Else
 billto.Clear
 billto.AddItem "⁄„Ì· ‰Â«∆Ì"
  billto.AddItem "„ﬁ«Ê· »«ÿ‰"
 bill_Type.Clear
 bill_Type.AddItem "Ã“∆Ì"
 bill_Type.AddItem "‰Â«∆Ì"
 
 End If
connection_string = Cn.ConnectionString
Adodc1.ConnectionString = connection_string
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select *  from project_billl" '  order by account_no"
Adodc1.Refresh
End Sub

Private Sub txtid_Change()
If txtid.text <> "" Or txtid.text <> "txtid" Then
connection_string = Cn.ConnectionString
Adodc7.ConnectionString = connection_string
Adodc7.CommandType = adCmdText
Adodc7.RecordSource = "select * from project_bill_details where bill_id=" & Val(txtid.text)
Adodc7.Refresh
DataGrid2.Refresh
End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
'On Error GoTo ErrTrap
Select Case Index
    Case 0
        If Not (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
            Adodc1.Recordset.MovePrevious
            If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
        End If
    Case 1
        If Not (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
            Adodc1.Recordset.MoveFirst
        End If
    Case 2
        If Not (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
            Adodc1.Recordset.MoveLast
        End If
    Case 3
        If Not (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
            Adodc1.Recordset.MoveNext
            If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
        End If
End Select
 Adodc7.ConnectionString = connection_string
Adodc7.CommandType = adCmdText
Adodc7.RecordSource = "select * from project_bill_details where bill_id=" & Val(txtid.text)

Adodc7.Refresh
DataGrid2.Refresh
Exit Sub
ErrTrap:
End Sub

