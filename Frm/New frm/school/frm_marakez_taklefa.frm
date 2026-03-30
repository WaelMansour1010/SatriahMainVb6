VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{784C0C13-85E7-4E11-A8FB-F0243A135D03}#2.0#0"; "SuperLablel.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frm_marakez_taklefa 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7110
   ClientLeft      =   3525
   ClientTop       =   1470
   ClientWidth     =   8370
   Icon            =   "frm_marakez_taklefa.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   8370
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   0
      TabIndex        =   72
      Top             =   6650
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "„Ê«›ﬁ"
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
      MICON           =   "frm_marakez_taklefa.frx":000C
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
      Caption         =   "»ÕÀ"
      Height          =   1455
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   2280
      Width           =   4575
      Begin MSDataListLib.DataCombo dccategory 
         Height          =   315
         Left            =   600
         TabIndex        =   70
         Top             =   960
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "«·›∆…"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   71
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«”„"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   66
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ﬂÊœ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   65
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txt_mod_flag 
      Height          =   495
      Left            =   7320
      TabIndex        =   63
      Text            =   "Text6"
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "„›⁄·"
      DataField       =   "block"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2040
      TabIndex        =   62
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "code"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   2640
      TabIndex        =   50
      Top             =   6480
      Width           =   3135
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   51
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
         MICON           =   "frm_marakez_taklefa.frx":0028
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
         Left            =   0
         TabIndex        =   52
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "frm_marakez_taklefa.frx":0044
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
         Left            =   2040
         TabIndex        =   53
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
         MICON           =   "frm_marakez_taklefa.frx":0060
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
         TabIndex        =   54
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   2280
      TabIndex        =   43
      Top             =   9120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2880
         TabIndex        =   47
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4440
         TabIndex        =   45
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   120
      TabIndex        =   38
      Top             =   9120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   41
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   4095
      Left            =   6960
      TabIndex        =   34
      Top             =   9120
      Width           =   1455
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   0
         TabIndex        =   35
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
         Begin ALLButtonS.ALLButton Command1 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   36
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
            MICON           =   "frm_marakez_taklefa.frx":007C
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
            TabIndex        =   37
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
      Left            =   360
      TabIndex        =   28
      Top             =   8520
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
      Left            =   360
      TabIndex        =   21
      Top             =   9720
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   23
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
         TabIndex        =   22
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
      Left            =   240
      TabIndex        =   18
      Top             =   8760
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
         TabIndex        =   33
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
         TabIndex        =   24
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   9960
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   240
      TabIndex        =   14
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox txtID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "account_name"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "account_no"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "account_type"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "frm_marakez_taklefa.frx":0098
      Left            =   120
      List            =   "frm_marakez_taklefa.frx":00A5
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "—∆Ì”Ì"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "id"
      DataSource      =   "Adodc4"
      Height          =   285
      Left            =   9720
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      DataField       =   "last_root"
      DataSource      =   "Adodc5"
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_marakez_taklefa.frx":00BB
      Height          =   2775
      Left            =   0
      TabIndex        =   8
      Top             =   3840
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4895
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
         DataField       =   "Code"
         Caption         =   "ﬂÊœ «·„—ﬂ“"
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
         Caption         =   "«”„ «·„—ﬂ“"
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
         DataField       =   "Type_name"
         Caption         =   "Ì »⁄"
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
            ColumnWidth     =   1649.764
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
            Object.Visible         =   0   'False
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   570
      Left            =   4080
      Top             =   120
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
      Left            =   9840
      Top             =   360
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
      Left            =   9840
      Top             =   720
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
      Left            =   9840
      Top             =   1080
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
      Left            =   9840
      Top             =   1320
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
      Left            =   9840
      Top             =   1680
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
      TabIndex        =   27
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
      MICON           =   "frm_marakez_taklefa.frx":00D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   3840
      Width           =   7455
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frm_marakez_taklefa.frx":00EC
      DataField       =   "parent_no"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3240
      TabIndex        =   48
      Top             =   1920
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ListField       =   "account_name"
      BoundColumn     =   "account_no"
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
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   0
      Left            =   1665
      TabIndex        =   55
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
      ButtonImage     =   "frm_marakez_taklefa.frx":0101
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
      TabIndex        =   56
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
      ButtonImage     =   "frm_marakez_taklefa.frx":049B
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
      TabIndex        =   57
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
      ButtonImage     =   "frm_marakez_taklefa.frx":0835
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
      Left            =   1125
      TabIndex        =   58
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
      ButtonImage     =   "frm_marakez_taklefa.frx":0BCF
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin MSDataListLib.DataCombo dctype 
      Bindings        =   "frm_marakez_taklefa.frx":0F69
      DataField       =   "type"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   120
      TabIndex        =   61
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ListField       =   ""
      BoundColumn     =   "account_no"
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
   Begin ALLButtonS.ALLButton Command1 
      Height          =   495
      Index           =   5
      Left            =   0
      TabIndex        =   67
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
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
      MICON           =   "frm_marakez_taklefa.frx":0F7E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "«·›∆…"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   60
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ „—ﬂ“ «· ﬂ·›…"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6720
      TabIndex        =   59
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   " «·„—ﬂ“  «·—∆Ì”Ì"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      TabIndex        =   49
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "„—«ﬂ“ «· ﬂ·›…"
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
      TabIndex        =   13
      Top             =   0
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ „—ﬂ“ «· ﬂ·›…"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6720
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   11880
      TabIndex        =   11
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "—ﬁ„ „—ﬂ“ «· ﬂ·›…"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10080
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   " ‰Ê⁄ „—ﬂ“ «· ﬂ·›…"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1920
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frm_marakez_taklefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As String
Dim last_root As Integer
Dim last_geeral As Integer
Dim last_branch As Integer

Private Sub ALLButton1_Click()

    If Adodc1.Recordset.RecordCount > 0 Then

        marakes_taklefa_tawze3.DataCombo1.BoundText = Adodc1.Recordset.Fields!code
        Unload Me
    End If

End Sub

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

Private Sub Command1_Click(Index As Integer)
    On Error Resume Next

    Select Case Index

        Case 0
            txt_mod_flag.text = "N"
            'Command1(1).Enabled = True
            'Command1(0).Enabled = False
            Adodc1.Recordset.AddNew
            Adodc1.Recordset.Fields!account_name = ""
            Adodc1.Recordset.Fields!block = 1
            Adodc1.Recordset.update
            Adodc1.Recordset.MoveLast
    
        Case 1
  
            If my_language = "E" Then
                If txtid.text = "" Then MsgBox " Specify cost center name", vbCritical: Exit Sub
                If Combo1.text = "" Then MsgBox "Specify cost center Type ", vbCritical: Exit Sub
                If dctype.BoundText = "" Then MsgBox "Specify cost center Category ", vbCritical: Exit Sub
            
            Else

                If txtid.text = "" Then MsgBox " Õœœ «”„ „—ﬂ“ «· ﬂ·›…", vbCritical: Exit Sub
                If Combo1.text = "" Then MsgBox "Õœœ ‰Ê⁄ „—ﬂ“ «· ﬂ·›… ", vbCritical: Exit Sub
                If dctype.BoundText = "" Then MsgBox "Õœœ «·›∆… «· Ì Ì »⁄Â« «·„—ﬂ“", vbCritical: Exit Sub
            
            End If

            'Command1(1).Enabled = False
            'Command1(0).Enabled = True
            'Œ«’ »«·Õ”«»«  «·—∆Ì”Ì…
            If txt_mod_flag.text = "N" Then
                Adodc5.RecordSource = "select  max(account_NO) as last_root  from   markaas_taklefa where account_type='—∆Ì”Ì'  or account_type='major' "
                Adodc5.Refresh

                If Adodc5.Recordset.RecordCount = 0 Or IsNull(Adodc5.Recordset.Fields!last_root) Then
                    last_root = 1
                Else
                    Adodc5.Recordset.MoveLast
                    last_root = Adodc5.Recordset.Fields!last_root + 1
                End If

                '**************************************************************
                'Œ«’ »«·Õ”«» «·«»‰
                Adodc4.RecordSource = "select  max(child_no) as last_child  from   markaas_taklefa where parent_no='" & DataCombo1.BoundText & "'"
                Adodc4.Refresh

                If Adodc4.Recordset.RecordCount = 0 Or IsNull(Adodc4.Recordset.Fields!last_child) Then
                    last_child = 1
                Else
                    'Adodc4.Recordset.MoveLast
                    last_child = Adodc4.Recordset.Fields!last_child + 1
                End If

                '***************************************************************
                If Combo1.text <> "" And (Combo1.text = "—∆Ì”Ì" Or Combo1.text = "Major") Then
                    Text1.text = last_root

                End If

                If Combo1.text <> "" And (Combo1.text = "⁄«„" Or Combo1.text = "›—⁄Ì" Or Combo1.text = "General" Or Combo1.text = "Sub") Then
 
                    Text1.text = DataCombo1.BoundText & "-" & last_child
                End If

                Dim i As Integer
                Dim level_counter As Integer
                level_counter = 0

                For i = 1 To Len(Text1.text)

                    If Mid(Text1.text, i, 1) = "-" Then
                        level_counter = level_counter + 1
                    End If

                Next i

                Adodc1.Recordset.Fields!child_no = last_child
                Adodc1.Recordset.Fields!Level = level_counter + 1
                txt_mod_flag.text = "E"
            End If

            If Text3.text = "" Then Text3.text = Text1.text
            Adodc1.Recordset.Fields!Type_name = Me.dctype.text

            Adodc1.Recordset.update

        Case 2
            x = MsgBox("Â· «‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbCritical + vbYesNo)

            If x = vbNo Then
                Exit Sub
            End If

            If Adodc1.Recordset.RecordCount > 0 Then
                Adodc1.Recordset.delete
                Adodc1.Refresh
                DataGrid1.Refresh
            End If

        Case 3

            If Adodc1.Recordset.RecordCount > 0 Then
    
                Form3.case_id = Me.name
   
                Form3.Show
            End If

        Case 4
            On Error Resume Next
            x = InputBox("«œŒ· «·—ﬁ„ «·„ÿ·Ê» «·»ÕÀ ⁄‰…")
      
            Adodc1.CommandType = adCmdText
            Adodc1.RecordSource = "select * from  markaas_taklefa where not (account_no is null)  and account_no='" & x & "'"
            Adodc1.Refresh

        Case 5
            x = InputBox("«œŒ· ﬂ·„… «·»ÕÀ")
            Adodc1.CommandType = adCmdText
            Adodc1.RecordSource = "select * from markaas_taklefa where not (account_no is null)  and account_name like '%" & x & "%'"
            Adodc1.Refresh

    End Select

End Sub

Private Sub Combo1_Click()
    On Error Resume Next
    Frame2.Enabled = False
    ' Combo1.Clear
    'Combo1.AddItem "Major"
    'Combo1.AddItem "general"
    'Combo1.AddItem "sub"
    'Combo1.Text = "Major"

    If Combo1.text <> "" And (Combo1.text = "⁄«„" Or Combo1.text = "General") Then
 
        Frame2.Enabled = True

        Adodc2.RecordSource = " select * from   markaas_taklefa where account_type='Major'  or  account_type='General' or account_type='—∆Ì”Ì'  or  account_type='⁄«„'  order by account_no"
        Adodc2.Refresh

    End If

    If Combo1.text <> "" And (Combo1.text = "›—⁄Ì" Or Combo1.text = "Sub") Then
        'Frame4.Enabled = True
        Frame2.Enabled = True

        Adodc2.RecordSource = " select * from   markaas_taklefa where   account_type='⁄«„' or  account_type='general'  order by account_no"
        Adodc2.Refresh

    End If

End Sub

Private Sub DataCombo1_Click(Area As Integer)
    On Error Resume Next
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    If DataCombo1.BoundText <> "" Then

        sql = "select * from markaas_taklefa where account_no='" & DataCombo1.BoundText & "'"
 
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
        If Rs3.RecordCount > 0 Then
            dctype.BoundText = IIf(Not IsNumeric(Rs3("Type").value), 0, Rs3("Type").value)
        End If

        Rs3.Close

    End If

End Sub

Private Sub dccategory_Click(Area As Integer)

    If dccategory.BoundText = "" Then Exit Sub
    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
      
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from  markaas_taklefa where type =" & dccategory.BoundText
    Adodc1.Refresh

End Sub

Private Sub Form_Activate()
    On Error Resume Next

    'user_priviliges_adodc.ConnectionString = connection_string: user_priviliges_adodc.CommandType = adCmdText
    '    If my_language = "E" Then
    '    user_priviliges_adodc.RecordSource = "select * from USER_PRIVILIGES where employee_id=" & Val(current_user) & "and [no]='" & screen_name.Caption & "'"
    '    Else
    '    user_priviliges_adodc.RecordSource = "select * from USER_PRIVILIGES where employee_id=" & Val(current_user) & "and [no]='" & screen_name.Caption & "'"
    '
    '    End If
    'user_priviliges_adodc.Refresh
    '
    '    If user_priviliges_adodc.Recordset.RecordCount = 0 Then
    '            If my_language = "E" Then
    '        MsgBox "NOT allowed ", vbCritical
    '
    '        Else
    '        MsgBox "€Ì— „”„ÊÕ »«” Œœ«„ Â–… «·‘«‘…  ", vbCritical
    '        End If
    ' Unload Me
    '    End If

    'If user_priviliges_adodc.Recordset.Fields![View] = False Then
    '        If my_language = "E" Then
    '        MsgBox "NOT allowed ", vbCritical
    '
    '        Else
    '        MsgBox "€Ì— „”„ÊÕ »«” Œœ«„ Â–… «·‘«‘…  ", vbCritical
    '        End If
    '
    'Unload Me
    'End If

    'Command1(0).Enabled = user_priviliges_adodc.Recordset.Fields![add_new]
    'Command1(1).Enabled = user_priviliges_adodc.Recordset.Fields![Save]
    'Command1(2).Enabled = user_priviliges_adodc.Recordset.Fields![Delete]
    'Command1(3).Enabled = user_priviliges_adodc.Recordset.Fields![Print]

End Sub

Private Sub Form_Load()
    On Error Resume Next

    '

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

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    'LoadSettings
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
    
        Temp = XPBtnMove(1).left
        XPBtnMove(1).left = XPBtnMove(2).left
        XPBtnMove(2).left = Temp

        Temp = XPBtnMove(0).left
        XPBtnMove(0).left = XPBtnMove(3).left
        XPBtnMove(3).left = Temp

        Label6.Caption = "ID"
        Label5.Caption = "Type"
        Label1.Caption = "Name"

        Label15.Caption = "Code"
        Check1.Caption = "Active"
        Label16.Caption = "Category"

        Label8.Caption = "Primary"
        DataGrid1.Columns(1).Caption = "Code"
        DataGrid1.Columns(2).Caption = "Name"
        DataGrid1.Columns(3).Caption = "Category"
        'DataGrid1.Columns(6).Caption = "Level"
        Combo1.Clear
        Combo1.AddItem "Major"
        Combo1.AddItem "General"
        Combo1.AddItem "Sub"
        Combo1.text = "Major"
        Frame5.Visible = False

        Text1.Alignment = 0
        txtid.Alignment = 0
        DataCombo1.RightToLeft = False
        Combo1.RightToLeft = False
  
        DataGrid1.RightToLeft = False
        CMD_language.Caption = "⁄—»Ì"
        Frame4.Visible = True
        Frame3.Visible = True
        Frame8.Visible = True
    
        Label9.Caption = "Cost Center Data"
        Me.Caption = Label9.Caption
  
        Command1(0).Caption = "new"
        Command1(1).Caption = "save"
        Command1(2).Caption = "delete"
        SuperLabel2.text = "Search"
        Command1(4).Caption = "By ID"
        Command1(5).Caption = "ÚSearch"
  
        Adodc1.Caption = "move"
        ' Me.Width = 10000
        Frame10.Caption = "Search"
        Label19.Caption = "Category"
        Label17.Caption = "Code"
        Label18.Caption = "Name"

    End If
 
    Dim My_SQL As String
    My_SQL = "  select  id,name  from Marakes_taklefa_type  "

    fill_combo dctype, My_SQL

    My_SQL = "  select id,name from Marakes_taklefa_type   "
    fill_combo dccategory, My_SQL

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  from markaas_taklefa  order by account_no"
    Adodc1.Refresh
    'where not (account_no is null)
    'Command1(1).Enabled = True
    'Command1(0).Enabled = False

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst

        If IsNull(Adodc1.Recordset.Fields!account_no) Then
            GoTo ll
        End If

    End If

    ' Adodc1.Recordset.AddNew
    '     Adodc1.Recordset.Fields!account_name = ""
    'Adodc1.Recordset.update
    'Adodc1.Recordset.MoveLast
ll:
    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from   markaas_taklefa where  (account_type='—∆Ì”Ì'  or  account_type='⁄«„'  or account_type='Major'  or  account_type='general') and not (account_no is null)  order by account_no "
    Adodc2.Refresh

    'Adodc3.ConnectionString = connection_string
    'Adodc3.CommandType = adCmdText
    'Adodc3.RecordSource = "select * from   account_index where account_type='⁄«„' order by account_general_no"
    'Adodc3.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    'Adodc4.RecordSource = "select * from   account_index where black_list=0 and  (account_type='›—⁄Ì' or  account_type='Sub' ) and not (account_no is null)order by account_no "
    'Adodc4.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select  max(account_NO) as last_root  from   markaas_taklefa where (account_type='—∆Ì”Ì' or  account_type='general') and not (account_no is null) "
    Adodc5.Refresh

    'Adodc6.ConnectionString = connection_string
    'Adodc6.CommandType = adCmdText
    'Adodc6.RecordSource = "select *  from account_index  where  black_list=0 and not (account_no is null) order by account_no"
    'Adodc6.Refresh
    'If OPEN_NEW_SCREEN = True Then
    'Command1_Click (0)
    'End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Exit Sub
End Sub

Private Sub Text6_Change()
      
    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from  markaas_taklefa where code like'%" & Text6.text & "%'"
    Adodc1.Refresh
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, _
                        Shift As Integer)
    'If KeyCode = 13 Then
    '      Adodc1.CommandType = adCmdText
    '      Adodc1.RecordSource = "select * from  markaas_taklefa where code like'%" & Text6.text & "%'"
    '      Adodc1.Refresh
         
    'End If

End Sub

Private Sub Text7_Change()
    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from  markaas_taklefa where account_name like'%" & Text7.text & "%'"
    Adodc1.Refresh
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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
 
    Exit Sub
ErrTrap:
End Sub

