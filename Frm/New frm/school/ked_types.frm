VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{784C0C13-85E7-4E11-A8FB-F0243A135D03}#2.0#0"; "SuperLablel.ocx"
Begin VB.Form ked_types 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇäæÇÚ ÇáãÓÊäÏÇÊ"
   ClientHeight    =   6720
   ClientLeft      =   4800
   ClientTop       =   4215
   ClientWidth     =   7860
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   7860
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0C0&
      DataField       =   "TYPE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "ked_types.frx":0000
      Left            =   2280
      List            =   "ked_types.frx":0016
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Frame Frame7 
      Height          =   1215
      Left            =   1920
      TabIndex        =   34
      Top             =   5400
      Width           =   3975
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   35
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÍÝÙ"
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
         MICON           =   "ked_types.frx":0066
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
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÍÐÝ"
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
         BCOL            =   255
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "ked_types.frx":0082
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
         TabIndex        =   37
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ØÈÇÚÉ"
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
         MICON           =   "ked_types.frx":009E
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
         Index           =   0
         Left            =   2640
         TabIndex        =   38
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÌÏíÏ"
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
         MICON           =   "ked_types.frx":00BA
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
         Top             =   840
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
         Caption         =   " ÊÍÑíß"
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
         Index           =   0
         Left            =   -120
         TabIndex        =   39
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   1080
      TabIndex        =   29
      Top             =   7320
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label xx 
         Caption         =   "ÇáãæÙÝ ÇáÍÇáí"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3840
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label yy 
         Caption         =   "ÇáÞÓã"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1560
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2640
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   1080
      TabIndex        =   24
      Top             =   7320
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2640
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   23
      ToolTipText     =   " ÇááÛÉ"
      Top             =   0
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
      MICON           =   "ked_types.frx":00D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1935
      Left            =   0
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Label Label9 
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
         TabIndex        =   51
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Type"
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
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "ID"
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
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   240
      TabIndex        =   13
      Top             =   7800
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      DataField       =   "ked_name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      DataField       =   "ked_no"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   6720
      Top             =   5400
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
   Begin SuperLablel.SuperLabel SuperLabel1 
      Height          =   615
      Left            =   2280
      TabIndex        =   16
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      Text            =   "ÇäæÇÚ ÇáãÓÊäÏÇÊ"
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1935
      Left            =   5880
      TabIndex        =   17
      Top             =   600
      Width           =   1575
      Begin VB.Label Label3 
         Caption         =   "äæÚ ÇáÓäÏ"
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
         Left            =   360
         TabIndex        =   40
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "ãÓáÓá "
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
         Left            =   480
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "ÇÓã ÇáäæÚ"
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
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   -1440
      Top             =   3600
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
   Begin MSDataListLib.DataCombo M 
      Bindings        =   "ked_types.frx":00F2
      DataField       =   "M1"
      DataSource      =   "Adodc1"
      Height          =   480
      Index           =   1
      Left            =   4200
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "name"
      BoundColumn     =   "departement_name"
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin MSDataListLib.DataCombo M 
      Bindings        =   "ked_types.frx":0107
      DataField       =   "M2"
      DataSource      =   "Adodc1"
      Height          =   480
      Index           =   2
      Left            =   4200
      TabIndex        =   4
      Top             =   3240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "name"
      BoundColumn     =   "departement_name"
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin MSDataListLib.DataCombo M 
      Bindings        =   "ked_types.frx":011C
      DataField       =   "M3"
      DataSource      =   "Adodc1"
      Height          =   480
      Index           =   3
      Left            =   4200
      TabIndex        =   6
      Top             =   3720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "name"
      BoundColumn     =   "departement_name"
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin MSDataListLib.DataCombo M 
      Bindings        =   "ked_types.frx":0131
      DataField       =   "M4"
      DataSource      =   "Adodc1"
      Height          =   480
      Index           =   4
      Left            =   4200
      TabIndex        =   8
      Top             =   4200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "name"
      BoundColumn     =   "departement_name"
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin MSDataListLib.DataCombo M 
      Bindings        =   "ked_types.frx":0146
      DataField       =   "M5"
      DataSource      =   "Adodc1"
      Height          =   480
      Index           =   5
      Left            =   4200
      TabIndex        =   10
      Top             =   4680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "name"
      BoundColumn     =   "departement_name"
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin MSDataListLib.DataCombo d 
      Bindings        =   "ked_types.frx":015B
      DataField       =   "D1"
      DataSource      =   "Adodc1"
      Height          =   480
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "name"
      BoundColumn     =   "departement_name"
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin MSDataListLib.DataCombo d 
      Bindings        =   "ked_types.frx":0170
      DataField       =   "D2"
      DataSource      =   "Adodc1"
      Height          =   480
      Index           =   2
      Left            =   840
      TabIndex        =   5
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "name"
      BoundColumn     =   "departement_name"
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin MSDataListLib.DataCombo d 
      Bindings        =   "ked_types.frx":0185
      DataField       =   "D3"
      DataSource      =   "Adodc1"
      Height          =   480
      Index           =   3
      Left            =   840
      TabIndex        =   7
      Top             =   3600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "name"
      BoundColumn     =   "departement_name"
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin MSDataListLib.DataCombo d 
      Bindings        =   "ked_types.frx":019A
      DataField       =   "D4"
      DataSource      =   "Adodc1"
      Height          =   480
      Index           =   4
      Left            =   840
      TabIndex        =   9
      Top             =   4080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "name"
      BoundColumn     =   "departement_name"
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin MSDataListLib.DataCombo d 
      Bindings        =   "ked_types.frx":01AF
      DataField       =   "D5"
      DataSource      =   "Adodc1"
      Height          =   480
      Index           =   5
      Left            =   840
      TabIndex        =   11
      Top             =   4560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "name"
      BoundColumn     =   "departement_name"
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin VB.Label Label4 
      Caption         =   "ÏÇÆä5"
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
      Index           =   8
      Left            =   3000
      TabIndex        =   50
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "ÏÇÆä4"
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
      Index           =   7
      Left            =   3000
      TabIndex        =   49
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "ÏÇÆä3"
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
      Index           =   6
      Left            =   3000
      TabIndex        =   48
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "ÏÇÆä2"
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
      Index           =   5
      Left            =   3000
      TabIndex        =   47
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "ÏÇÆä1"
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
      Index           =   4
      Left            =   3000
      TabIndex        =   46
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "ãÏíä 5 "
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
      Index           =   3
      Left            =   6480
      TabIndex        =   45
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "ãÏíä 4"
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
      Index           =   2
      Left            =   6480
      TabIndex        =   44
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "ãÏíä 3"
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
      Index           =   1
      Left            =   6480
      TabIndex        =   43
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "ãÏíä 2"
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
      Left            =   6480
      TabIndex        =   42
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "ãÏíä 1"
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
      Index           =   0
      Left            =   6480
      TabIndex        =   41
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "ked_types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

    If Index = 0 Then
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields!ked_name = ""
        Adodc1.Recordset.update
        Adodc1.Recordset.MoveLast

    Else

        If Index = 1 Then
    
            If my_language = "E" Then
                If Text2.text = "" Then MsgBox "write type first", vbCritical: Exit Sub

            Else

                If Text2.text = "" Then MsgBox "ÇßÊÈ   äæÚ ÇáÞíÏ ", vbCritical: Exit Sub
            End If
        
            'Adodc1.Recordset.Fields!departement_manager_id = DataCombo1.Text
            Adodc1.Recordset.Fields!type_id = Combo1.ListIndex

            Adodc1.Recordset.update
        Else

            If Index = 2 Then
 
                Dim x As Integer

                If my_language = "E" Then
                    x = MsgBox("Confirm delete", vbCritical + vbYesNo)
                Else
                    x = MsgBox("åá ÇäÊ ãÊÃßÏ ãä ÇáÍÐÝ", vbCritical + vbYesNo)
              
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
                        MsgBox "áÇ íæÌÏ ãÇ íãßä ÍÐÝÉ", vbCritical
                    End If
                
                End If

                Exit Sub

            End If
        End If
    End If

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
    '        MsgBox "ÛíÑ ãÓãæÍ ÈÇÓÊÎÏÇã åÐÉ ÇáÔÇÔÉ  ", vbCritical
    '        End If
    '  Unload Me
    '    End If
    '
    'If user_priviliges_adodc.Recordset.Fields![View] = False Then
    '        If my_language = "E" Then
    '        MsgBox "NOT allowed ", vbCritical
    '
    '        Else
    '        MsgBox "ÛíÑ ãÓãæÍ ÈÇÓÊÎÏÇã åÐÉ ÇáÔÇÔÉ  ", vbCritical
    '        End If
    '
    'Unload Me
    'End If
    '
    'Command1(0).Enabled = user_priviliges_adodc.Recordset.Fields![add_new]
    'Command1(1).Enabled = user_priviliges_adodc.Recordset.Fields![Save]
    'Command1(2).Enabled = user_priviliges_adodc.Recordset.Fields![Delete]

End Sub

Private Sub Form_Load()
    On Error Resume Next

    If my_language = "E" Then
        CMD_language.ToolTipText = "change Language"
        'Command13.ToolTipText = "F3 Departement Search "
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
    
    On Error Resume Next

    If my_language = "E" Then
        CMD_language.Caption = "ÚÑÈí"
        Text1.Alignment = 0
        Text2.Alignment = 0
        Combo1.RightToLeft = False
  
        Frame2.Visible = False
        Frame3.Visible = True
        SuperLabel1.text = "Document Type"
        Me.Caption = SuperLabel1.text
        Command1(0).Caption = "new"
        Command1(1).Caption = "save"
        Command1(2).Caption = "delete"
        Adodc1.Caption = "move"
  
    End If
 
    'LoadSettings
    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from ked_types  where not(ked_name='') and not(ked_name is null) " ' where departement_no=0"
    Adodc1.Refresh

    '  where  NOT (ked_name ='')
    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveLast

        If Adodc1.Recordset.Fields!ked_name = "" Then
            GoTo ll
        End If

    End If

    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields!ked_name = ""
    Adodc1.Recordset.update
    Adodc1.Recordset.MoveLast

ll:
    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from fixed_accounts  " ' where departement_no=0"
    Adodc2.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        
End Sub

