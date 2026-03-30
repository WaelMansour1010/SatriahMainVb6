VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form car_alarm 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " «· –ﬂÌ— »„” √Ã—Ì‰ ”Ê›  ‰ ÂÌ ⁄ﬁÊœÂ„ "
   ClientHeight    =   7845
   ClientLeft      =   3150
   ClientTop       =   930
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   10980
   Begin ALLButtonS.ALLButton close_all_alarms 
      Height          =   615
      Left            =   9960
      TabIndex        =   38
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "car_alarm.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame5 
      Height          =   4935
      Left            =   120
      TabIndex        =   36
      Top             =   8520
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
         TabIndex        =   37
         Top             =   1200
         Width           =   9975
      End
   End
   Begin VB.Frame InfoE 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   7080
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label zz 
         BackStyle       =   0  'Transparent
         Caption         =   "Departemnt"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3000
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dept_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Departement"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3960
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label vv 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee name"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame infoA 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3840
      TabIndex        =   29
      Top             =   7080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label dep_a 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee name"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label xx 
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   32
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label yy 
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label emp_a 
         BackStyle       =   0  'Transparent
         Caption         =   "Departemnt"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3480
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   1800
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
         TabIndex        =   35
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
         TabIndex        =   23
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·«Ì«„ «·„ »ﬁÌ…"
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   10650
      Begin VB.Line Line11 
         X1              =   9000
         X2              =   9000
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Days remaining for Insurance"
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
         Height          =   855
         Left            =   6120
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   300
         X2              =   300
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line5 
         X1              =   7500
         X2              =   7500
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line4 
         X1              =   4515
         X2              =   4515
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line3 
         X1              =   3015
         X2              =   3015
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line2 
         X1              =   6000
         X2              =   6000
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1500
         X2              =   1500
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Days delay"
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
         Left            =   9000
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Licens/warrentye Days delay"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   7560
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Days remaining for license/warrenty"
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
         Left            =   4560
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance End"
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
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "License/warrenty end"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment#"
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
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
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
      Begin VB.Label screen_name 
         Alignment       =   1  'Right Justify
         Caption         =   "M29"
         Height          =   255
         Left            =   3360
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox Text5 
      DataField       =   "driver_name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "id"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text3 
      DataField       =   "alarm"
      DataSource      =   "qry_adoc"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      DataField       =   "CHILD_ID"
      DataSource      =   "Adodc5"
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   11400
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7320
      Top             =   3360
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "car_alarm.frx":001C
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   2040
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
      Top             =   3600
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
      Top             =   4560
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
      TabIndex        =   16
      ToolTipText     =   " €ÌÌ—  «··€…"
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
      MICON           =   "car_alarm.frx":0031
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
      Caption         =   " «· –ﬂÌ— »„” √Ã—Ì‰ ”Ê›  ‰ ÂÌ ⁄ﬁÊœÂ„ "
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
      TabIndex        =   0
      Top             =   600
      Width           =   10695
   End
End
Attribute VB_Name = "car_alarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim start_load  As Boolean

Private Sub subject_no_Change()
On Error Resume Next
DATEDIILBL.Caption = DateDiff("D", predect_end_time, Now)
End Sub

Private Sub close_all_alarms_Click()
 
Unload messages_frm
Unload rent_notes
Unload risky_items_alarm
Unload car_out_warsha
Unload alarm_frm
Unload car_alarm
Unload ITEMS_ALARM
End Sub

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

 


Private Sub Form_Activate()
On Error Resume Next

If start_load = False Then
start_load = True
Adodc1.ConnectionString = connection_string
 Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select *  from cars  where branch_no=" & branch_no & " and departement='" & departement_name & "'and not(car_no is null) "
Adodc1.Refresh
 
 Adodc2.ConnectionString = connection_string
 Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "select *  from CAR_alarm  "
Adodc2.Refresh
 


Dim no_of_days1, no_of_days2 As Integer
Dim date1, date2 As Date
 
'empty alarm table
For i = 1 To Adodc2.Recordset.RecordCount

Adodc2.Recordset.Delete
Adodc2.Recordset.MoveNext
Next i


For i = 1 To Adodc1.Recordset.RecordCount
' date1 = DateAdd("d", -7, Adodc1.Recordset.Fields!estmara_date_end) 'day
 'date2 = DateAdd("d", -7, Adodc1.Recordset.Fields!end_insurance_date) 'day

 date1 = Adodc1.Recordset.Fields!estmara_date_end
 date2 = Adodc1.Recordset.Fields!end_insurance_date

  
 
'no_of_days1 = DateDiff("D", date1, Now)
'no_of_days2 = DateDiff("D", date1, Now)


If IsNull(Adodc1.Recordset.Fields!estmara_date_end) And IsNull(Adodc1.Recordset.Fields!end_insurance_date) Then
GoTo LL
End If
baky_estmara = 0
takher_estmara = 0
baky_INSURANCE = 0
takher_INSURANCE = 0
If Not IsNull(Adodc1.Recordset.Fields!estmara_date_end) Then
baky_estmara = DateDiff("D", Now, Adodc1.Recordset.Fields!estmara_date_end) '
takher_estmara = DateDiff("D", date1, Now)

End If

If Not IsNull(Adodc1.Recordset.Fields!end_insurance_date) Then
baky_INSURANCE = DateDiff("D", Now, Adodc1.Recordset.Fields!end_insurance_date) '

takher_INSURANCE = DateDiff("D", date2, Now)
End If


'If no_of_days1 >= -7 Or no_of_days2 >= -7 Then

'If qry_adoc.Recordset.Fields!no_of_day_to_alarm > no_of_days Then

 If takher_estmara > 0 Or takher_INSURANCE > 0 Then

Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields!Car_no = Adodc1.Recordset.Fields!fullcode
Adodc2.Recordset.Fields!estmara_end = Adodc1.Recordset.Fields!estmara_date_end
Adodc2.Recordset.Fields!insurance_end = Adodc1.Recordset.Fields!end_insurance_date
If baky_estmara > 0 Then
Adodc2.Recordset.Fields!baky_estmara = baky_estmara
Else
Adodc2.Recordset.Fields!baky_estmara = 0

End If

If baky_INSURANCE > 0 Then
Adodc2.Recordset.Fields!baky_INSURANCE = baky_INSURANCE
Else
Adodc2.Recordset.Fields!baky_INSURANCE = 0
End If

If takher_estmara > 0 Then
Adodc2.Recordset.Fields!takher_estmara = takher_estmara
Else
Adodc2.Recordset.Fields!takher_estmara = 0
End If
If takher_INSURANCE > 0 Then
Adodc2.Recordset.Fields!takher_INSURANCE = takher_INSURANCE
Else
Adodc2.Recordset.Fields!takher_INSURANCE = 0
End If

Adodc2.Recordset.Update
End If



'End If



LL:

Adodc1.Recordset.MoveNext
Next i

DoEvents

Adodc2.RecordSource = "select *  from CAR_alarm  "
Adodc2.Refresh


    If Adodc2.Recordset.RecordCount = 0 Then
Frame5.Visible = True

'   MsgBox "·« ÌÊÃœ  ‰»ÌÂ«  ··’Ì«‰… No Alarm", vbInformation
'   Unload Me
     End If
End If

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



End Sub

Private Sub Form_Load()
On Error Resume Next
Beep

    login.SkinFramework.ApplyWindow Me.hWnd

On Error Resume Next
CMD_language.ToolTipText = "Change Language"
 

If my_language = "E" Then
Label14.Caption = "No Alarm Today"
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
CMD_language.Caption = "⁄—»Ì"
Frame2.Visible = True
Frame3.Visible = False
Label17.Caption = "No Alarm today"
DataGrid1.RightToLeft = False

Label1.Caption = "End license,Insurance  and  warrenty dates  Alarm"
Me.Caption = Label1.Caption


End If


On Error Resume Next
LoadSettings





End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
start_load = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Label1.ForeColor = &H80000012 Then
Label1.ForeColor = &HFFFF&

Else
Label1.ForeColor = &H80000012


End If
End Sub
