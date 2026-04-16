VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form alarm_frm 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«· –ﬂÌ— »„” √Ã—Ì‰ ”ÌÕ· ⁄·ÌÂ„ ﬁ”ÿ ≈ÌÃ«— "
   ClientHeight    =   7800
   ClientLeft      =   3150
   ClientTop       =   930
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   10230
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame4 
      Height          =   5175
      Left            =   120
      TabIndex        =   37
      Top             =   7920
      Visible         =   0   'False
      Width           =   9975
      Begin VB.Label Label14 
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
         Height          =   4095
         Left            =   -240
         TabIndex        =   38
         Top             =   3480
         Width           =   9975
      End
   End
   Begin VB.Frame InfoE 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   7080
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label zz 
         BackStyle       =   0  'Transparent
         Caption         =   "Departemnt"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3000
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1440
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dept_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Departement"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3960
         TabIndex        =   29
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label vv 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee name"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame infoA 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3840
      TabIndex        =   32
      Top             =   7080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label dep_a 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee name"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label xx 
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ÊŸ› «·Õ«·Ì"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label yy 
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
      Begin VB.Label emp_a 
         BackStyle       =   0  'Transparent
         Caption         =   "Departemnt"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3480
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   9950
      Begin VB.Line Line10 
         Index           =   1
         X1              =   960
         X2              =   960
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁÌ„… «·„ÿ·Ê»…"
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
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
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
         Left            =   3240
         TabIndex        =   39
         Top             =   360
         Width           =   855
      End
      Begin VB.Line Line11 
         X1              =   6840
         X2              =   6840
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line6 
         Index           =   1
         X1              =   9600
         X2              =   9600
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Label Label13 
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
         Left            =   1200
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
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
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·‘«—⁄"
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
         Left            =   4800
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·ÕÌ"
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
         Left            =   5640
         TabIndex        =   23
         Top             =   360
         Width           =   1455
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
         Left            =   6480
         TabIndex        =   22
         Top             =   240
         Width           =   1695
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
         Left            =   8520
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Line Line10 
         Index           =   0
         X1              =   2100
         X2              =   2100
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line9 
         X1              =   5715
         X2              =   5715
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line8 
         X1              =   3200
         X2              =   3200
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line7 
         X1              =   4210
         X2              =   4210
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line6 
         Index           =   0
         X1              =   8205
         X2              =   8205
         Y1              =   120
         Y2              =   960
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   9930
      Begin VB.Line Line1 
         Index           =   1
         X1              =   300
         X2              =   300
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line5 
         X1              =   7800
         X2              =   7800
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line4 
         X1              =   5720
         X2              =   5720
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line3 
         X1              =   4220
         X2              =   4220
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line2 
         X1              =   6720
         X2              =   6720
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1710
         X2              =   1710
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "departemnt name"
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
         Left            =   7920
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Days delay"
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
         Left            =   6960
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Days remaining"
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
         Left            =   5760
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "date for maintenance"
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
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Type Maintenance"
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
         Height          =   495
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment#"
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
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   360
      TabIndex        =   9
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox Text5 
      DataField       =   "driver_name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "id"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text3 
      DataField       =   "alarm"
      DataSource      =   "qry_adoc"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      DataField       =   "car_no"
      DataSource      =   "alarm_adoc"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   " Ã«Â·"
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "CHILD_ID"
      DataSource      =   "Adodc5"
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   11400
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7320
      Top             =   3360
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   " ÕÊÌ· «·Ï «·’Ì«‰…"
      Height          =   615
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "alarm_frm.frx":0000
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   8705
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
         DataField       =   "opr_id"
         Caption         =   "opr_id"
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
         DataField       =   "car_no"
         Caption         =   "—ﬁ„ «·„⁄œÂ/«·”Ì«—…"
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
         DataField       =   "car_maintenace_type"
         Caption         =   "‰Ê⁄ «·’Ì«‰…"
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
         DataField       =   "car_maintenance_date"
         Caption         =   "«· «—ÌŒ «·„ﬁ —Õ ··’Ì«‰…"
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
         DataField       =   "alarm"
         Caption         =   "alarm"
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
         DataField       =   "ayam_motabkya"
         Caption         =   "«·«Ì«„ «·„ »ﬁÌ…"
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
         DataField       =   "ayam_ta2ker"
         Caption         =   "«Ì«„ «· √ŒÌ—"
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
      BeginProperty Column07 
         DataField       =   "departement_name"
         Caption         =   "«”„ «·ﬁ”„"
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
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2055.118
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc qry_adoc 
      Height          =   855
      Left            =   0
      Top             =   6000
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
   Begin MSAdodcLib.Adodc alarm_adoc 
      Height          =   855
      Left            =   0
      Top             =   6720
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   0
      Top             =   4920
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
      Left            =   0
      Top             =   5280
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
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   19
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
      MICON           =   "alarm_frm.frx":0019
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   585
      Left            =   0
      Top             =   0
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
   Begin ALLButtonS.ALLButton close_all_alarms 
      Height          =   615
      Left            =   9000
      TabIndex        =   41
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
      MICON           =   "alarm_frm.frx":0035
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
      Caption         =   "«· –ﬂÌ— »„” √Ã—Ì‰ ”ÌÕ· ⁄·ÌÂ„ ﬁ”ÿ ≈ÌÃ«— "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   9495
   End
End
Attribute VB_Name = "alarm_frm"
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

Private Sub Command1_Click()
On Error Resume Next
If alarm_adoc.Recordset.RecordCount = 0 Then
Exit Sub
End If

If alarm_adoc.Recordset.RecordCount <> 0 Then
     Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields!opr_date = DateValue(Now)
     Adodc1.Recordset.Fields!moghat_belt2men = vbTrue
     
     Adodc1.Recordset.Fields!Car_no = alarm_adoc.Recordset.Fields!Car_no
     If my_language = "E" Then
       Adodc1.Recordset.Fields!maintenance_type = "Periodic maintenance"
     Else
     Adodc1.Recordset.Fields!maintenance_type = "’Ì«‰… œÊ—Ì…"
     End If
     
     Adodc1.Recordset.Fields!error_description = alarm_adoc.Recordset.Fields!car_maintenace_type
     Adodc1.Recordset.Fields!departement_name_now = alarm_adoc.Recordset.Fields!departement_name
          Adodc1.Recordset.Fields!driver_name = alarm_adoc.Recordset.Fields!driver_name
  
         
    Adodc1.Recordset.Update
End If
        
        
     Adodc2.CommandType = adCmdText
     Adodc2.RecordSource = "select * from car_maintenaces where car_no='" & alarm_adoc.Recordset.Fields!Car_no & "' and car_maintenace_type='" & alarm_adoc.Recordset.Fields!car_maintenace_type & "'"
     Adodc2.Refresh
     
     If Adodc2.Recordset.RecordCount > 0 Then
     
     Adodc2.Recordset.Fields!alarm = 0
     Adodc2.Recordset.Update
     End If
     
        If my_language = "E" Then
             x = MsgBox("open maintenance Screen", vbExclamation + vbYesNo)
        Else
             x = MsgBox("Â·  —Ìœ › Õ ‘«‘… «·’Ì«‰…", vbExclamation + vbYesNo)
        End If
     If x = vbYes Then
     frm_maintenace.Show
     
        frm_maintenace.Adodc1.CommandType = adCmdText
        frm_maintenace.Adodc1.RecordSource = "select * from  maintenance where car_no='" & alarm_adoc.Recordset.Fields!Car_no & "'"
        frm_maintenace.Adodc1.Refresh
      frm_maintenace.Adodc1.Recordset.MoveLast

     End If
        
        
     alarm_adoc.Recordset.Delete
     alarm_adoc.Refresh
     DataGrid1.Refresh
        
End Sub

Private Sub Command2_Click()
On Error Resume Next
      If my_language = "E" Then
        x = MsgBox("confirm Ignore", vbCritical + vbYesNo)

      Else
        x = MsgBox(" √ﬂÌœ «· Ã«Â· ··⁄„·Ì…", vbCritical + vbYesNo)
      End If
If x = vbYes And alarm_adoc.Recordset.RecordCount > 0 Then
      
     Adodc2.CommandType = adCmdText
     Adodc2.RecordSource = "select * from car_maintenaces where car_no='" & alarm_adoc.Recordset.Fields!Car_no & "' and car_maintenace_type='" & alarm_adoc.Recordset.Fields!car_maintenace_type & "'"
     Adodc2.Refresh
     
     If Adodc2.Recordset.RecordCount > 0 Then
     
     Adodc2.Recordset.Fields!alarm = 0
     Adodc2.Recordset.Update
     End If


  If alarm_adoc.Recordset.RecordCount <> 0 Then
             alarm_adoc.Recordset.Delete
             alarm_adoc.Refresh
             DataGrid1.Refresh
        Else
        Exit Sub
        End If

End If

End Sub

Private Sub Form_Activate()
On Error Resume Next

If start_load = False Then
start_load = True
Adodc1.ConnectionString = connection_string
 Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select *  from maintenance  "
Adodc1.Refresh
 
 Adodc2.ConnectionString = connection_string
 Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "select *  from CARS  "
Adodc2.Refresh
 
 
  alarm_adoc.ConnectionString = connection_string
 alarm_adoc.CommandType = adCmdText
alarm_adoc.RecordSource = "select *  from alarm  "
alarm_adoc.Refresh

 
 
  qry_adoc.ConnectionString = connection_string
qry_adoc.CommandType = adCmdText

'  Adodc3.ConnectionString = connection_string

' Adodc3.CommandType = adCmdText
'Adodc3.RecordSource = "select * from departement where departement_no=" & current_user_departement
'Adodc3.Refresh
'If Adodc3.Recordset.RecordCount = 0 Then Frame4.Visible = True: Exit Sub
'If my_language = "E" Then
qry_adoc.RecordSource = "select * from car_maintenace_with_alarm_qry where    branch_no=" & branch_no & "and  departement='" & departement_name & "'"
'Else
'qry_adoc.RecordSource = "select * from car_maintenace_with_alarm_qry where departement='" & Adodc3.Recordset.Fields!departement_name & "'"
'End If

qry_adoc.Refresh


Dim no_of_days As Integer
Dim ayam_motabkya As Integer
Dim ayam_ta2ker As Integer
'empty alarm table
For i = 1 To alarm_adoc.Recordset.RecordCount

alarm_adoc.Recordset.Delete
alarm_adoc.Recordset.MoveNext
Next i


For i = 1 To qry_adoc.Recordset.RecordCount

 
no_of_days = DateDiff("D", Now, qry_adoc.Recordset.Fields!car_maintenance_date)

If no_of_days < 0 Then

'If qry_adoc.Recordset.Fields!no_of_day_to_alarm > no_of_days Then

ayam_motabkya = 0
ayam_ta2ker = Abs(no_of_days)

alarm_adoc.Recordset.AddNew
alarm_adoc.Recordset.Fields!Car_no = qry_adoc.Recordset.Fields!Car_no
alarm_adoc.Recordset.Fields!car_maintenace_type = qry_adoc.Recordset.Fields!car_maintenace_type
alarm_adoc.Recordset.Fields!car_maintenance_date = qry_adoc.Recordset.Fields!car_maintenance_date
alarm_adoc.Recordset.Fields!ayam_motabkya = ayam_motabkya
alarm_adoc.Recordset.Fields!ayam_ta2ker = ayam_ta2ker
alarm_adoc.Recordset.Fields!departement_name = qry_adoc.Recordset.Fields!DEPARTEMENT

alarm_adoc.Recordset.Fields!driver_name = qry_adoc.Recordset.Fields!driver_name


alarm_adoc.Recordset.Fields!branch_no = qry_adoc.Recordset.Fields!branch_no
alarm_adoc.Recordset.Update

'End If

Else
        
        If qry_adoc.Recordset.Fields!no_of_day_to_alarm <= no_of_days Then
 
        ayam_motabkya = no_of_days - -qry_adoc.Recordset.Fields!no_of_day_to_alarm
        ayam_ta2ker = 0
               
        End If
        
        If qry_adoc.Recordset.Fields!no_of_day_to_alarm > no_of_days Then
        
        ayam_motabkya = no_of_days
        ayam_ta2ker = 0
        
 
        End If
       alarm_adoc.Recordset.AddNew
        alarm_adoc.Recordset.Fields!Car_no = qry_adoc.Recordset.Fields!Car_no
        alarm_adoc.Recordset.Fields!car_maintenace_type = qry_adoc.Recordset.Fields!car_maintenace_type
        alarm_adoc.Recordset.Fields!car_maintenance_date = qry_adoc.Recordset.Fields!car_maintenance_date
        alarm_adoc.Recordset.Fields!ayam_motabkya = ayam_motabkya
        alarm_adoc.Recordset.Fields!ayam_ta2ker = ayam_ta2ker
        alarm_adoc.Recordset.Fields!departement_name = qry_adoc.Recordset.Fields!DEPARTEMENT
        alarm_adoc.Recordset.Fields!driver_name = qry_adoc.Recordset.Fields!driver_name
        alarm_adoc.Recordset.Fields!branch_no = qry_adoc.Recordset.Fields!branch_no
        alarm_adoc.Recordset.Update
        

GoTo LL


End If


'If qry_adoc.Recordset.Fields!no_of_day_to_alarm >= no_of_days Then

'ayam_motabkya = no_of_days
'ayam_ta2ker = 0
'End If




If qry_adoc.Recordset.Fields!no_of_day_to_alarm <= no_of_days Then

ayam_motabkya = 0
ayam_ta2ker = no_of_days - qry_adoc.Recordset.Fields!no_of_day_to_alarm

alarm_adoc.Recordset.AddNew
alarm_adoc.Recordset.Fields!Car_no = qry_adoc.Recordset.Fields!Car_no
alarm_adoc.Recordset.Fields!car_maintenace_type = qry_adoc.Recordset.Fields!car_maintenace_type
alarm_adoc.Recordset.Fields!car_maintenance_date = qry_adoc.Recordset.Fields!car_maintenance_date
alarm_adoc.Recordset.Fields!ayam_motabkya = ayam_motabkya
alarm_adoc.Recordset.Fields!ayam_ta2ker = ayam_ta2ker
alarm_adoc.Recordset.Fields!departement_name = qry_adoc.Recordset.Fields!DEPARTEMENT
 alarm_adoc.Recordset.Fields!driver_name = qry_adoc.Recordset.Fields!driver_name
        alarm_adoc.Recordset.Fields!branch_no = qry_adoc.Recordset.Fields!branch_no

alarm_adoc.Recordset.Update

End If


LL:
qry_adoc.Recordset.MoveNext
Next i

DoEvents
    If alarm_adoc.Recordset.RecordCount = 0 Then
Frame4.Visible = True

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

DataGrid1.RightToLeft = False

Label1.Caption = "Periodic maintenance schedules alarm"
Me.Caption = Label1.Caption
Command1.Caption = "Switch to maintenance"
Command2.Caption = "Ignore"

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
