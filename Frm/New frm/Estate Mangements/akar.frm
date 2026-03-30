VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form akar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·⁄ﬁ«—« "
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10845
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0C0&
      DataField       =   "nashat"
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
      ItemData        =   "akar.frx":0000
      Left            =   360
      List            =   "akar.frx":000D
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "units_trade_total"
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
      Left            =   360
      ScrollBars      =   2  'Vertical
      TabIndex        =   51
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   3495
      Left            =   6240
      TabIndex        =   29
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      Height          =   735
      Left            =   9840
      Picture         =   "akar.frx":0025
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "»ÕÀ ⁄‰ ”‰œF3"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "code"
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   720
      Width           =   1935
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   2160
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3120
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label yy 
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.Label xx 
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   1320
      TabIndex        =   17
      Top             =   7560
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   4920
      TabIndex        =   16
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   -3120
      TabIndex        =   13
      Top             =   8040
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
      Begin VB.Label screen_name 
         Alignment       =   1  'Right Justify
         Caption         =   "M35"
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   2520
      TabIndex        =   6
      Top             =   5400
      Width           =   5535
      Begin ALLButtonS.ALLButton Command1 
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "ÃœÌœ"
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
         MICON           =   "akar.frx":19B7
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
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Õ›Ÿ"
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
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "akar.frx":19D3
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
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Õ–›"
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
         BCOLO           =   255
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "akar.frx":19EF
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
         Left            =   1560
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
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
      Begin ALLButtonS.ALLButton Command1 
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
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
         MICON           =   "akar.frx":1A0B
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
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "„—›ﬁ« "
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
         MICON           =   "akar.frx":1A27
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   15
         Left            =   -120
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.TextBox txtname 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "floors"
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
      Left            =   -2280
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "building_no"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "no_of_entry"
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
      Left            =   6000
      TabIndex        =   3
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "units_total"
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
      Left            =   360
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "last_rent_price"
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
      Left            =   360
      TabIndex        =   1
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "last_price"
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
      Left            =   360
      TabIndex        =   0
      Top             =   4200
      Width           =   3135
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   -3000
      TabIndex        =   30
      ToolTipText     =   "Language  «··€…"
      Top             =   120
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
      MICON           =   "akar.frx":1A43
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "akar.frx":1A5F
      DataField       =   "owner_code"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   6000
      TabIndex        =   31
      Top             =   1320
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "vendor_name"
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
   Begin MSDataListLib.DataCombo prifix_combo 
      Bindings        =   "akar.frx":1A74
      DataField       =   "prifix"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   6000
      TabIndex        =   32
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   12632256
      ListField       =   "prifix"
      BoundColumn     =   "driver_name"
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
   Begin MSAdodcLib.Adodc Adodccode1 
      Height          =   465
      Left            =   -600
      Top             =   6960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   820
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
   Begin MSAdodcLib.Adodc Adodccode2 
      Height          =   345
      Left            =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
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
   Begin MSAdodcLib.Adodc Adodc_prifix 
      Height          =   465
      Left            =   720
      Top             =   6960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   820
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "akar.frx":1A8F
      DataField       =   "city_code"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   6000
      TabIndex        =   33
      Top             =   2280
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "city_name"
      BoundColumn     =   "id"
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
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "akar.frx":1AA4
      DataField       =   "hay_code"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   6000
      TabIndex        =   34
      Top             =   2760
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "hay_name"
      BoundColumn     =   "hay_name"
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "akar.frx":1AB9
      CausesValidation=   0   'False
      DataField       =   "street_code"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   6000
      TabIndex        =   35
      Top             =   3240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "street_name"
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
   Begin ALLButtonS.ALLButton Command100 
      Height          =   495
      Left            =   360
      TabIndex        =   50
      Top             =   1800
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   " ›’Ì·Ì ÊÕœ«  «·⁄ﬁ«—"
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
      MICON           =   "akar.frx":1ACE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "akar.frx":1AEA
      DataField       =   "type_code"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   360
      TabIndex        =   54
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   -2147483638
      ListField       =   "type"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8880
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Left            =   9000
      Top             =   6600
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Left            =   8880
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Left            =   9000
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.Label Label7 
      Caption         =   "‰Ê⁄ «·⁄ﬁ«—"
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
      Left            =   3600
      TabIndex        =   55
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "⁄œœ «·ÊÕœ«  «· Ã«—Ì…"
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
      Left            =   3720
      TabIndex        =   52
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "»Ì«‰«  «·⁄ﬁ«—« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3360
      TabIndex        =   49
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label10 
      Caption         =   "«·„«·ﬂ"
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
      Left            =   9480
      TabIndex        =   48
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label v 
      Caption         =   "ﬂÊœ «·⁄ﬁ«—"
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
      Left            =   9480
      TabIndex        =   47
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "⁄œœ «·«œÊ«—"
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
      Left            =   3840
      TabIndex        =   46
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "«·„œÌ‰…"
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
      Left            =   9480
      TabIndex        =   45
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "«·ÕÌ"
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
      Left            =   9480
      TabIndex        =   44
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "«·‘«—⁄"
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
      Left            =   9480
      TabIndex        =   43
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "«·⁄‰Ê«‰"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7320
      TabIndex        =   42
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label16 
      Caption         =   "—ﬁ„ «·⁄ﬁ«—"
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
      Left            =   9480
      TabIndex        =   41
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "⁄œœ «·„œ«Œ·"
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
      Left            =   9480
      TabIndex        =   40
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label19 
      Caption         =   "‰‘«ÿ «·⁄ﬁ«—"
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
      Left            =   3720
      TabIndex        =   39
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label20 
      Caption         =   "⁄œœ «·ÊÕœ«  «·”ﬂ‰Ì…"
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
      Height          =   735
      Left            =   3720
      TabIndex        =   38
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "«Œ— ﬁÌ„… «ÌÃ«—Ì… ”‰ÊÌ…"
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
      Height          =   735
      Left            =   3720
      TabIndex        =   37
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "”⁄— «·⁄ﬁ«—«·Õ«·Ì"
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
      Left            =   3720
      TabIndex        =   36
      Top             =   4320
      Width           =   2295
   End
End
Attribute VB_Name = "akar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As String
Dim fullcode As String
Dim code As Integer


Private Sub CMD_language_Click()
On Error Resume Next

If CMD_language.Caption = "EN" Then
my_language = "E"
 
Call Reload(Me)

 
Else
my_language = "A"
 
Call Reload(Me)
End If
End Sub

Private Sub Command2_Click()

       On Error Resume Next
box_search.Show
box_search.case_id = 0

End Sub

Private Sub Command100_Click()
If Adodc1.Recordset.RecordCount > 0 And Not IsNull(Adodc1.Recordset.Fields!fullcode) Then
Akar_details.Show
Akar_details.txtid = Adodc1.Recordset.Fields!fullcode


Akar_details.Adodc1.ConnectionString = connection_string
Akar_details.Adodc1.CommandType = adCmdText
Akar_details.Adodc1.RecordSource = "select *  from floors WHERE Akar_code='" & Adodc1.Recordset.Fields!fullcode & "'"
Akar_details.Adodc1.Refresh





End If

End Sub

Private Sub DataCombo2_Click(Area As Integer)
If DataCombo2.Text = "" Then Exit Sub
DataCombo3.Text = ""

   Adodc4.ConnectionString = connection_string
 Adodc4.CommandType = adCmdText
Adodc4.RecordSource = "select * from hay where city_id=" & DataCombo2.BoundText
Adodc4.Refresh

End Sub

Private Sub DataCombo2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF3 Then

emp_search.Show

emp_search.case_id = 8
End If

If KeyCode = vbKeyF6 Then
EMPLOYEES.Show
End If



If KeyCode = vbKeyF5 Then
Adodc5.Refresh
DataCombo2.ReFill
End If






End Sub

Private Sub DataCombo3_Click(Area As Integer)
If DataCombo3.Text = "" Then Exit Sub
DataCombo4.Text = ""

   Adodc5.ConnectionString = connection_string
 Adodc5.CommandType = adCmdText
Adodc5.RecordSource = "select * from streets where hay_id='" & DataCombo3.Text & "'"
Adodc5.Refresh



End Sub

Private Sub prifix_combo_Click(Area As Integer)

       On Error Resume Next
 '      Adodc_prifix.Refresh
'prifix_combo.ReFill

'txtID.Locked = True
End Sub

Private Sub prifix_combo_KeyUp(KeyCode As Integer, Shift As Integer)

       On Error Resume Next
       If KeyCode = vbKeyF6 Then
coding.Show
End If

prifix_combo.Text = ""
End Sub




Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Select Case Index

Case 0

    Adodc1.Recordset.AddNew
'        prifix_combo.Enabled = True
'Adodc1.Recordset.Fields!departement = departement_name
'Adodc1.Recordset.Fields!branch_no = branch_no
'prifix_combo.Enabled = True

'        Adodc1.Recordset.Fields!BOX_name = ""
'Adodc1.Recordset.Update
'Adodc1.Recordset.MoveLast
Case 1
 get_code
 
 
If my_language = "E" Then
If txtname.Text = "" Then MsgBox " Specify Box name", vbCritical: Exit Sub
If DataCombo2.Text = "" Then MsgBox "Specify  Box Manger", vbCritical: Exit Sub

If DataCombo1.Text = "" Then MsgBox "Specify Account no", vbCritical: Exit Sub


Else
If txtid.Text = "" Then MsgBox " Õœœ ﬂÊœ «·⁄ﬁ«—", vbCritical: Exit Sub
'If DataCombo2.Text = "" Then MsgBox "Õœœ  «„Ì‰ «·Œ“Ì‰…  ", vbCritical: Exit Sub

'If DataCombo1.Text = "" Then MsgBox "Õœœ —ﬁ„ «·Õ”«»", vbCritical: Exit Sub

End If

'Adodc1.Recordset.Fields!BOX_account-no = DataCombo1.Text

     If fullcode = "error" Then Exit Sub
 '    prifix_combo.Enabled = False
   Adodc1.Recordset.Fields!fullcode = prifix_combo.Text & txtid.Text
       Adodc1.Recordset.Fields!branch_no = branch_no
    Adodc1.Recordset.Fields!DEPARTEMENT = departement_name
    ' Adodc1.Recordset.Fields!manger = DataCombo2.Text
    Adodc1.Recordset.Update

Case 2
       If my_language = "E" Then
              x = MsgBox("Confirm delete", vbCritical + vbYesNo)
            Else
            x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
              
            End If
        If x = vbNo Then
Exit Sub
End If
    If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.Delete
    Adodc1.Refresh
    DataGrid1.Refresh
    End If

 

Case 4
On Error Resume Next
On Error Resume Next

If my_language = "E" Then
If Text1.Text = "" Then MsgBox "Select Voucher First": Exit Sub

Else
If Text1.Text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— ⁄ﬁ«—  «Ê·«": Exit Sub
End If

imaged.Show
imaged.txtopeation_type = "⁄ﬁ«—"
imaged.SUBJECT_NO = prifix_combo.Text & txtid.Text
 If my_language = "E" Then
 imaged.Label6.Caption = "Voucher #"
 imaged.Caption = "Voucher Attachments"
 Else
imaged.Label6.Caption = "—ﬁ„ «·⁄ﬁ«—"
 imaged.Caption = "„—›ﬁ«  «·⁄ﬁ«—"
End If
imaged.Adodc1.CommandType = adCmdText
imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '⁄ﬁ«—' and subject_no='" & prifix_combo.Text & txtid.Text & "'"
imaged.Adodc1.Refresh
If imaged.Adodc1.Recordset.RecordCount > 0 Then

imaged.DBPix201.Visible = True
Else
imaged.DBPix201.Visible = False
End If



Case 5
    If my_language = "E" Then
 x = InputBox("«œŒ· «”„ «·Œ“Ì‰…")
Else
x = InputBox("Enter box name")

End If


            Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from BOXs where  Not ([BOX_account-no] Is Null) and BOX_name like '%" & x & "%'"
        Adodc1.Refresh


End Select

End Sub

 

 

Private Sub Command13_Click()
On Error Resume Next

'Acccount_search.Show
'Acccount_search.case_id = 7
End Sub

Private Sub DataCombo1_DblClick(Area As Integer)
On Error Resume Next

Unload account_info_bar
account_info_bar.Show
account_info_bar.item_code = DataCombo1.Text
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF3 Then
Acccount_search.Show
Acccount_search.case_id = 7
End If

If KeyCode = vbKeyF6 Then
account_index.Show
End If



If KeyCode = vbKeyF5 Then
Adodc2.Refresh
DataCombo1.ReFill
End If


End Sub

Function get_code()

       On Error Resume Next
If txtid.Text <> "" Then Exit Function

Dim tempcode As String
Dim code As String

Dim coding_auto As Boolean

If prifix_combo.Text = "" Then Exit Function
Adodccode1.RecordSource = "select * from coding where FIELD_no=0 and  branch_no=" & branch_no & "and departement='" & departement_name & "' and prifix='" & prifix_combo.Text & "'"


Adodccode1.Refresh

If Adodccode1.Recordset.RecordCount > 0 Then

coding_auto = Adodccode1.Recordset.Fields!Auto
                        If coding_auto = False Then
                        
                        txtid.Locked = False
                        MsgBox "ÌÃ» «œŒ«· —ﬁ„ «·„⁄œÂ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·‘«‘« "
                        Exit Function
                        End If

                                 If coding_auto = True Then
                              no_of_digit = Adodccode1.Recordset.Fields!no_of_digit
                              zeros = Adodccode1.Recordset.Fields!zeros
                
                              prifix = prifix_combo.Text
                              End If
End If


Adodccode2.RecordSource = "select max(code)  as last_code from akar  where branch_no=" & branch_no & "and departement='" & departement_name & "' and prifix='" & prifix & "'"
Adodccode2.Refresh

If IsNull(Adodccode2.Recordset.Fields!last_code) Then

tempcode = prifix & "1"
            If Len(tempcode) < no_of_digit Then
            
             diffrent = no_of_digit - Len(tempcode)
             tempcode = prifix
                    If zeros = True Then
                                   For i = 1 To diffrent
                                   tempcode = tempcode & "0"
                                   code = code & "0"
                                   
                                   Next i
                       End If
                
            fullcode = tempcode & "1"
            code = code & "1"
            
            Else
              If Len(tempcode) > no_of_digit Then
              MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
               fullcode = "error"
              Exit Function
              
              Else
            fullcode = tempcode
            code = 1
            End If
             End If
Else
           tempcode = prifix & Val(Adodccode2.Recordset.Fields!last_code) + 1
            If Len(tempcode) < no_of_digit Then
            
             diffrent = no_of_digit - Len(tempcode)
             tempcode = prifix
                          If zeros = True Then

                        For i = 1 To diffrent
                        tempcode = tempcode & "0"
                              code = code & "0"
                        Next i
                        
                        End If
            fullcode = tempcode & Val(Adodccode2.Recordset.Fields!last_code) + 1
            code = code & Val(Adodccode2.Recordset.Fields!last_code + 1)
            Else
              If Len(tempcode) > no_of_digit Then
              MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ… ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â«"
               fullcode = "error"
              Exit Function
              
              Else
            fullcode = tempcode
              code = Val(Adodccode2.Recordset.Fields!last_code) + 1
            End If
             End If

End If
  txtid.Text = code
    If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.Fields!fullcode = fullcode
    End If
End Function
Private Sub Form_Activate()
On Error Resume Next


 
user_priviliges_adodc.ConnectionString = connection_string: user_priviliges_adodc.CommandType = adCmdText
    If my_language = "E" Then
    user_priviliges_adodc.RecordSource = "select * from USER_PRIVILIGES where employee_id=" & Val(current_user) & "and [no]='" & screen_name.Caption & "'"
    Else
    user_priviliges_adodc.RecordSource = "select * from USER_PRIVILIGES where employee_id=" & Val(current_user) & "and [no]='" & screen_name.Caption & "'"
    
    End If
user_priviliges_adodc.Refresh

    If user_priviliges_adodc.Recordset.RecordCount = 0 Then
            If my_language = "E" Then
        MsgBox "NOT allowed ", vbCritical
        
        Else
        MsgBox "€Ì— „”„ÊÕ »«” Œœ«„ Â–… «·‘«‘…  ", vbCritical
        End If
  Unload Me
    End If

If user_priviliges_adodc.Recordset.Fields![View] = False Then
        If my_language = "E" Then
        MsgBox "NOT allowed ", vbCritical
        
        Else
        MsgBox "€Ì— „”„ÊÕ »«” Œœ«„ Â–… «·‘«‘…  ", vbCritical
        End If

Unload Me
End If

Command1(0).Enabled = user_priviliges_adodc.Recordset.Fields![add_new]
Command1(1).Enabled = user_priviliges_adodc.Recordset.Fields![Save]
Command1(2).Enabled = user_priviliges_adodc.Recordset.Fields![Delete]

End Sub

Private Sub Form_Load()
On Error Resume Next


   ' login.SkinFramework.ApplyWindow Me.hWnd


If my_language = "E" Then
temp = Command1(0).Left
Command1(0).Left = Command1(2).Left
Command1(2).Left = temp
CMD_language.ToolTipText = "Change Language"
Command13.ToolTipText = "F3 Account Search "

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

 Me.Left = (MDIForm1.Width - Me.Width) / 2
    Me.Top = (MDIForm1.Height - Me.Height) / 2 - 500
    
 If my_language = "E" Then
 
       
 DataGrid1.RightToLeft = False
 DataCombo1.RightToLeft = False
  DataCombo2.RightToLeft = False


  txtid.Alignment = 0
  txtname.Alignment = 0
  DataCombo1.RightToLeft = False
  
 CMD_language.Caption = "⁄—»Ì"
  Frame2.Visible = True
  Frame3.Visible = True
    Frame8.Visible = True
    Frame4.Visible = False
  Label2.Caption = "Boxs Data"
  Me.Caption = Label42.Caption
  
  Command1(0).Caption = "new"
  Command1(1).Caption = "save"
  Command1(2).Caption = "delete"
  SuperLabel2.Text = "Search"
  Command1(4).Caption = "By ID"
  Command1(5).Caption = "By Name"
  
  Adodc1.Caption = "move"
   
  
 End If
 

LoadSettings
Adodc1.ConnectionString = connection_string
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select *  from Akar where branch_no=" & branch_no & "and departement='" & departement_name & "'"
Adodc1.Refresh
'Where Not ([BOX_account-no] Is Null)

If Adodc1.Recordset.RecordCount > 0 Then
Adodc1.Recordset.MoveLast

If IsNull(Adodc1.Recordset.Fields![fullcode]) Then
'GoTo LL
End If

End If
 


Adodc_prifix.ConnectionString = connection_string
Adodc_prifix.CommandType = adCmdText
Adodc_prifix.RecordSource = "select * from coding  where  FIELD_no=0 and branch_no=" & branch_no & "and departement='" & departement_name & "'"
Adodc_prifix.Refresh

Adodccode1.ConnectionString = connection_string
Adodccode1.CommandType = adCmdText
Adodccode2.ConnectionString = connection_string
Adodccode2.CommandType = adCmdText
     
     Adodc2.ConnectionString = connection_string
 Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "select * from city"
Adodc2.Refresh


   Adodc3.ConnectionString = connection_string
 Adodc3.CommandType = adCmdText
Adodc3.RecordSource = "select * from akar_type"
Adodc3.Refresh


   Adodc5.ConnectionString = connection_string
 Adodc5.CommandType = adCmdText
Adodc5.RecordSource = "select * from vendors"
Adodc5.Refresh



End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
End Sub


