VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form rent_notes 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«· –ﬂÌ— »„” √Ã—Ì‰  ÊÃœ ⁄·ÌÂ„ „·«ÕŸ«  "
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10815
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame5 
      Height          =   4935
      Left            =   120
      TabIndex        =   25
      Top             =   8400
      Visible         =   0   'False
      Width           =   10695
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "·« ÌÊÃœ  ‰»ÌÂ«  «·ÌÊ„"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2775
         Left            =   0
         TabIndex        =   26
         Top             =   1200
         Width           =   9975
      End
   End
   Begin VB.Frame InfoE 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   6960
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label zz 
         BackStyle       =   0  'Transparent
         Caption         =   "Departemnt"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dept_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Departement"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3960
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label vv 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee name"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame infoA 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3840
      TabIndex        =   15
      Top             =   6960
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label dep_a 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee name"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label xx 
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label yy 
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label emp_a 
         BackStyle       =   0  'Transparent
         Caption         =   "Departemnt"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3480
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   10665
      Begin VB.Line Line12 
         X1              =   1620
         X2              =   1620
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·ÕÌ"
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
         Height          =   735
         Left            =   6000
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.Line Line6 
         Index           =   1
         X1              =   10320
         X2              =   10320
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄„«—… —ﬁ„"
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
         Height          =   615
         Left            =   3480
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·„·«ÕŸ« "
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
         Height          =   735
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·„” √Ã—"
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
         Height          =   735
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·‘«—⁄"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   4800
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œÌ‰…"
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
         Height          =   735
         Left            =   7680
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ÊÕœ…"
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
         Height          =   735
         Left            =   9120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Line Line10 
         X1              =   3120
         X2              =   3120
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line9 
         X1              =   7635
         X2              =   7635
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line8 
         X1              =   4635
         X2              =   4635
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line7 
         X1              =   6135
         X2              =   6135
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line6 
         Index           =   0
         X1              =   9120
         X2              =   9120
         Y1              =   120
         Y2              =   960
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   7920
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
         Caption         =   "M29"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox Text5 
      DataField       =   "driver_name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "id"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text3 
      DataField       =   "alarm"
      DataSource      =   "qry_adoc"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      DataField       =   "CHILD_ID"
      DataSource      =   "Adodc5"
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   11280
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7320
      Top             =   3240
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "rent_notes.frx":0000
      Height          =   4695
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   24
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "car_no"
         Caption         =   "car_no"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "estmara_end"
         Caption         =   "estmara_end"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "insurance_end"
         Caption         =   "insurance_end"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "baky_estmara"
         Caption         =   "baky_estmara"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "baky_insurance"
         Caption         =   "baky_insurance"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "takher_estmara"
         Caption         =   "takher_estmara"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "takher_insurance"
         Caption         =   "takher_insurance"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1604.976
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   480
      Top             =   3480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1032
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
      Height          =   855
      Left            =   240
      Top             =   4440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1508
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
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   28
      ToolTipText     =   " €ÌÌ—  «··€…"
      Top             =   -120
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
      MICON           =   "rent_notes.frx":0015
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton close_all_alarms 
      Height          =   615
      Left            =   9600
      TabIndex        =   30
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "«€·«ﬁ ﬂ· «· ‰»ÌÂ« "
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
      MICON           =   "rent_notes.frx":0031
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "«· –ﬂÌ— »„” √Ã—Ì‰  ÊÃœ ⁄·ÌÂ„ „·«ÕŸ«  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   29
      Top             =   480
      Width           =   10695
   End
End
Attribute VB_Name = "rent_notes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_all_alarms_Click()
Unload messages_frm
Unload rent_notes
Unload risky_items_alarm
Unload car_out_warsha
Unload alarm_frm
Unload car_alarm
Unload ITEMS_ALARM
End Sub

Private Sub Form_Load()
  login.SkinFramework.ApplyWindow Me.hWnd
     Me.Left = (MDIForm1.Width - Me.Width) / 2
    Me.Top = (MDIForm1.Height - Me.Height) / 2 - 500
End Sub

