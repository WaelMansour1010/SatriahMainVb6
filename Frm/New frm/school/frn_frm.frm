VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form operatiomn_update_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›« Ê—… «·«ﬁ”«ÿ-‘∆Ê‰ «·ÿ·«»"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   14085
   Begin VB.TextBox Text27 
      Alignment       =   2  'Center
      DataField       =   "member_type"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton Command9 
      Caption         =   "«·›—ﬁ"
      Height          =   375
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  'Center
      DataField       =   "last_update_year"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   66
      Text            =   " "
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text23 
      Alignment       =   1  'Right Justify
      DataField       =   "notes"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      DataField       =   "update_year"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   63
      Text            =   " "
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "€—«„«  «·“ÊÃ…"
      Height          =   2775
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   5520
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command5 
         Caption         =   "«Œ Ì«—"
         Height          =   495
         Left            =   120
         TabIndex        =   59
         Top             =   1560
         Width           =   852
      End
      Begin VB.CommandButton Command8 
         Caption         =   "«Œ Ì«—"
         Height          =   495
         Left            =   120
         TabIndex        =   58
         Top             =   960
         Width           =   852
      End
      Begin VB.TextBox Text25 
         Alignment       =   2  'Center
         DataField       =   "wife_FINES_TOTAL"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Text            =   "0"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         DataField       =   "wife_FINES_TOTAL1"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Text            =   "0"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   " Ã„⁄Ì… ⁄„Ê„Ì…  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         TabIndex        =   62
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "  √ŒÌ— "
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
         Left            =   2160
         TabIndex        =   61
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
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
         Left            =   2160
         TabIndex        =   60
         Top             =   360
         Width           =   2415
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   4200
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      DataField       =   "FINES_TOTAL1"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   360
      TabIndex        =   53
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "«Œ Ì«—"
      Height          =   375
      Left            =   2160
      TabIndex        =   52
      Top             =   3480
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      DataField       =   "USER_NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command60 
      Caption         =   "«Œ Ì«—"
      Height          =   372
      Left            =   11520
      TabIndex        =   49
      Top             =   3360
      Width           =   612
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      DataField       =   "no_of_person"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6720
      TabIndex        =   46
      Top             =   5040
      Width           =   1452
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      DataField       =   "ACTIVITY_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   3840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      DataField       =   "CHILD_NAME"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   41
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      DataField       =   "CHILD_ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   600
      TabIndex        =   40
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "«Œ Ì«—"
      Height          =   375
      Left            =   6960
      TabIndex        =   38
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Height          =   855
      Left            =   8400
      Picture         =   "frn_frm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "»ÕÀ"
      Height          =   375
      Left            =   8640
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Õ›Ÿ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8640
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_DATE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_NO"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_TYPE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "INSTALLMENTS_TOTAL"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9120
      TabIndex        =   10
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9120
      TabIndex        =   9
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataField       =   "FINES_TOTAL"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "ID_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9120
      TabIndex        =   7
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      DataField       =   "OTHERS_FEES"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      DataField       =   "TOTAL_VALUE"
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
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      DataField       =   "CHILD_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "«⁄«œ… «·Õ”«Ì« "
      Height          =   615
      Left            =   -360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      DataField       =   "discounts"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      DataField       =   "update_year"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text15 
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc4"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text15"
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   492
      Left            =   6240
      Top             =   8640
      Width           =   7692
      _ExtentX        =   13573
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
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
      Height          =   492
      Left            =   720
      Top             =   8640
      Visible         =   0   'False
      Width           =   1452
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frn_frm.frx":1992
      Height          =   2655
      Left            =   6960
      TabIndex        =   19
      Top             =   6000
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   0   'False
      HeadLines       =   1
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "op_id"
         Caption         =   "op_id"
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
         DataField       =   "CHILD_ID"
         Caption         =   "CHILD_ID"
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
         DataField       =   "CHILD_NAME"
         Caption         =   "CHILD_NAME"
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
         DataField       =   "MEMBER_VALUE"
         Caption         =   "MEMBER_VALUE"
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
         DataField       =   "member_card_value"
         Caption         =   "member_card_value"
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
         DataField       =   "FINES_value"
         Caption         =   "FINES_value"
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
         DataField       =   "activity_value"
         Caption         =   "activity_value"
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
         DataField       =   "PAYED"
         Caption         =   "PAYED"
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
      BeginProperty Column08 
         DataField       =   "FINES_value1"
         Caption         =   "FINES_value1"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   492
      Left            =   2040
      Top             =   0
      Visible         =   0   'False
      Width           =   3012
      _ExtentX        =   5318
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
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
      Height          =   495
      Left            =   480
      Top             =   2640
      Visible         =   0   'False
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
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
      Height          =   495
      Left            =   2160
      Top             =   2640
      Visible         =   0   'False
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
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
      Height          =   495
      Left            =   5040
      Top             =   3360
      Visible         =   0   'False
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
      Caption         =   "KEST"
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   495
      Left            =   4560
      Top             =   2520
      Visible         =   0   'False
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
      Caption         =   "GHRAMA"
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
      Height          =   492
      Left            =   3000
      Top             =   3960
      Visible         =   0   'False
      Width           =   1812
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
      Caption         =   "GHRAMA"
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
   Begin VB.Label Label33 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11520
      TabIndex        =   71
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   69
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Caption         =   "«Œ— ”‰…  ÃœÌœ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   6840
      TabIndex        =   67
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Caption         =   "„·«ÕŸ« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3120
      TabIndex        =   65
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Caption         =   " €—«„… Ã„⁄Ì… ⁄„Ê„Ì…"
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
      Left            =   2640
      TabIndex        =   54
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "»Ê«”ÿ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3120
      TabIndex        =   51
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·«‰‘ÿ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   48
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Caption         =   "⁄œœ «·«›—«œ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   732
      Left            =   8280
      TabIndex        =   47
      Top             =   4920
      Width           =   972
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·«‰‘ÿ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   7200
      TabIndex        =   45
      Top             =   3840
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·“ÊÃ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   43
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·“ÊÃ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   42
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "«·€—«„… "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   7440
      TabIndex        =   39
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   " —ﬁ„ «·ÿ«·»"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11520
      TabIndex        =   37
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·ÿ«·»"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   11160
      TabIndex        =   36
      Top             =   2160
      Width           =   3972
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "«·—”Ê„ «·œ—«”Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   8640
      TabIndex        =   35
      Top             =   5640
      Width           =   1812
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   " «—ÌŒ «·⁄„·Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   6960
      TabIndex        =   34
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   " —ﬁ„ «·⁄·„Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   11400
      TabIndex        =   33
      Top             =   480
      Width           =   3012
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   12240
      TabIndex        =   32
      Top             =   960
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·«ﬁ”«ÿ  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   12240
      TabIndex        =   31
      Top             =   3360
      Width           =   1812
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "«·—”Ê„ «·œ—«”Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   11520
      TabIndex        =   30
      Top             =   2760
      Width           =   3012
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "„’«—Ì› «·Õ«›·…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   11760
      TabIndex        =   29
      Top             =   3960
      Width           =   2412
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "„’«—Ì› «Œ—Ï"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   4440
      TabIndex        =   28
      Top             =   4440
      Width           =   2412
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "«Ã„«·Ì «·„ÿ·Ê»"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   4440
      TabIndex        =   27
      Top             =   5040
      Width           =   2412
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "«Ã„«·Ì „’—Ê›«  «· «»⁄Ì‰"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   11520
      TabIndex        =   26
      Top             =   4920
      Width           =   2652
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "«”„ «· «»⁄"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   10800
      TabIndex        =   25
      Top             =   5640
      Width           =   972
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·ﬂ«—‰Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   6840
      TabIndex        =   24
      Top             =   5640
      Width           =   1812
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   732
      Left            =   13320
      TabIndex        =   23
      Top             =   5280
      Width           =   972
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "›—Êﬁ«   Œ’„"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   11880
      TabIndex        =   22
      Top             =   4440
      Width           =   2412
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "”‰… «· ÃœÌœ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6960
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «· «»⁄"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   12120
      TabIndex        =   20
      Top             =   5640
      Width           =   972
   End
End
Attribute VB_Name = "operatiomn_update_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim check As Integer
Dim yy As Integer

Private Sub Command1_Click()

    Command2_Click

    If Text13.text = "" Then Text13.text = 0
    If Text10.text = "" Then Text10.text = 0

    Adodc1.Recordset.update

End Sub

Function update_date()
    Command2_Click
End Function

Private Sub Command2_Click()
    Adodc1.Recordset.Fields!total_value = val(Text12.text) + val(Text7.text) + val(Text5.text) + val(Text8.text) + val(Text21.text) + val(Text9.text) + val(Text10.text) - val(Text13.text) + val(Text18.text) + val(Text25.text) + val(Text26.text)
    Adodc1.Recordset.update
End Sub

Private Sub Command3_Click()
    member_search.Show
    member_search.from = 5

    'X = InputBox("«œŒ· «·—ﬁ„ «Ê Ã“¡ „‰ «·—ﬁ„", "‘«‘… «·»ÕÀ »«·—ﬁ„")

    'select * from operations where PAYED=0 and
    'Adodc1.CommandType = adCmdText
    'Adodc1.RecordSource = "select *  FROM OPERATIONS where  operation_type=' ÃœÌœ ⁄÷ÊÌ…' and  PAYED=0 and MEMBER_ID LIKE'%" & X & "%'"
    'Adodc1.Refresh
    'If Text1.Text = "" Then Text1.Text = 0
    '   Adodc2.CommandType = adCmdText
    '    Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,Fines_value ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND MEMBER_ID=" & Text1.Text
    '    Adodc2.Refresh
End Sub

Private Sub Command4_Click()
    x = InputBox("«œŒ· «·«”„ «Ê Ã“¡ „‰ «·«”„", "‘«‘… «·»ÕÀ »«·«”„")

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM OPERATIONS where  operation_type=' ÃœÌœ ⁄÷ÊÌ…' and operation_type=' ÃœÌœ ⁄÷ÊÌ…' and PAYED=0 and MEMBER_NAME LIKE'%" & x & "%'"
    Adodc1.Refresh

    If Text1.text = "" Then Text1.text = 0
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,Fines_value ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND MEMBER_ID=" & Text1.text
    Adodc2.Refresh
End Sub

Private Sub Command5_Click()
    fines_update.Show
    fines_update.wife.Caption = "1"

    fines_update.Adodc2.CommandType = adCmdText
    fines_update.Adodc2.RecordSource = "select FINES_NO,FINES_VALUE,PAYED_DATE,ACTIVATED,payed,child_id,MEMBER_ID,MEMBER_name,FINES_TYPE, FINES_TOTAL  FROM FINES_DETAILS where FINES_TYPE='€—«„… Ã„⁄Ì… ⁄„Ê„Ì…' and  payed =0 and MEMBER_ID =" & Text1.text & "and wife=1"
    fines_update.Adodc2.Refresh
    'Dim SUM As Single
    'Dim J As Integer
    'On Error GoTo ll
    '    Adodc2.Recordset.Fields!Fines_NAME = ""
    '     Adodc2.Recordset.Fields!FINES_VALUE = 0
    '   Adodc2.Recordset.Update
    '        SUM = 0
    '
    '    Adodc2.CommandType = adCmdText
    '    Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,fines_value ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND  MEMBER_ID=" & Text1.Text
    '   Adodc2.Refresh
    '    DataGrid2.Refresh
    '  For J = 1 To Adodc2.Recordset.RecordCount
    '  SUM = SUM + Adodc2.Recordset.Fields!MEMBER_VALUE + Adodc2.Recordset.Fields!member_card_value + Adodc2.Recordset.Fields!FINES_VALUE + Adodc2.Recordset.Fields!activity_value
    '
    '   Adodc2.Recordset.MoveNext
    '  Next J
    
    ' If Adodc1.Recordset.RecordCount > 0 Then
    '  Adodc1.Recordset.Fields!CHILD_VALUE = SUM
    '  End If
    
    ' Adodc1.Recordset.Fields!TOTAL_VALUE = SUM + Text7.Text + Text5.Text + Text8.Text + Text21.Text + Text9.Text + Text10.Text - Text13.Text
    ' Adodc1.Recordset.Update
    '    Adodc2.Refresh
    ' DataGrid2.Refresh
    ' updatedata
    'll:
End Sub

Private Sub Command6_Click()
    fines_update.Show

    fines_update.Adodc2.CommandType = adCmdText
    fines_update.Adodc2.RecordSource = "select FINES_NO,FINES_VALUE,PAYED_DATE,ACTIVATED,payed,child_id,MEMBER_ID,MEMBER_name,FINES_TYPE, FINES_TOTAL  FROM FINES_DETAILS where FINES_TYPE='€—«„… Ã„⁄Ì… ⁄„Ê„Ì…' and  payed =0 and MEMBER_ID =" & Text1.text & "and child_id=" & Text16.text
    fines_update.Adodc2.Refresh

End Sub

Private Sub Command60_Click()
    installments_update.Show
    installments_update.Adodc1.CommandType = adCmdText
    installments_update.Adodc1.RecordSource = "select *  FROM Installments where MEMBER_ID =" & Text1.text & " and child_id=" & Text16.text & " ORDER BY CHILD_ID DESC"
    installments_update.Adodc1.Refresh

    installments_update.Adodc2.CommandType = adCmdText
    installments_update.Adodc2.RecordSource = "select INSTALLMENT_NO,INSTALLMENT_VALUE,DATE_OF_PAYED,ACTIVATED,payed,child_id  FROM INSTALLMENT_DETAILS where payed=0 and MEMBER_ID =" & Text1.text & "and child_id=" & Text16.text
    installments_update.Adodc2.Refresh

End Sub

Private Sub Command7_Click()
    fines_update.Show
    fines_update.Adodc2.CommandType = adCmdText
    fines_update.Adodc2.RecordSource = "select  FINES_NO,FINES_VALUE,PAYED_DATE,ACTIVATED,payed,child_id,MEMBER_ID,MEMBER_name,FINES_TYPE,FINES_TOTAL   FROM FINES_DETAILS where   payed =0 and MEMBER_ID =" & Text1.text & "and child_id=" & Text16.text
    fines_update.Adodc2.Refresh

End Sub

Private Sub DataCombo1_Click(Area As Integer)

    If DataCombo1.text <> "" And Adodc2.Recordset.EOF = False And Adodc2.Recordset.BOF = False Then
        Adodc2.Recordset.Fields!Fines_NAME = DataCombo1.text
        Adodc2.Recordset.update
        Adodc2.Recordset.MoveFirst
        '    Adodc2.Refresh
        DataGrid2.Refresh
        Text3_Change
    End If
    
End Sub

Private Sub Command8_Click()
    fines_update.Show
    fines_update.wife.Caption = "1"
    fines_update.Adodc2.CommandType = adCmdText
    fines_update.Adodc2.RecordSource = "select  FINES_NO,FINES_VALUE,PAYED_DATE,ACTIVATED,payed,child_id,MEMBER_ID,MEMBER_name,FINES_TYPE,FINES_TOTAL   FROM FINES_DETAILS where FINES_TYPE='€—«„…  √ŒÌ—' and  payed =0 and MEMBER_ID =" & Text1.text & "and wife=1"
    fines_update.Adodc2.Refresh

    'On Error GoTo ll
    'If Not IsNull(Adodc2.Recordset.Fields!op_id) Then
    'Fine_TYPES.Show
    'Fine_TYPES.Text1.Enabled = False
    'Fine_TYPES.Text2.Enabled = False
    'Fine_TYPES.Text3.Enabled = False
    'Fine_TYPES.Text4.Enabled = False
    'Fine_TYPES.Command1.Visible = False
    'Fine_TYPES.Command2.Visible = False
    'Fine_TYPES.Command3.Visible = True
    'Fine_TYPES.Text5.Text = Adodc2.Recordset.Fields!MEMBER_VALUE
    'Fine_TYPES.Label5.Caption = "«Œ «— ‰Ê⁄  «·€—«„… „‰ ›÷·ﬂ"
    'End If

    'll:
End Sub

Public Function updatedata()
    Dim SUM As Single
    Dim J As Integer
    SUM = 0

    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,fines_value ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND  MEMBER_ID=" & Text1.text
    Adodc2.Refresh
    DataGrid2.Refresh

    For J = 1 To Adodc2.Recordset.RecordCount
        SUM = SUM + Adodc2.Recordset.Fields!MEMBER_VALUE + Adodc2.Recordset.Fields!member_card_value + Adodc2.Recordset.Fields!fines_value '+ Adodc2.Recordset.Fields!activity_value
    
        Adodc2.Recordset.MoveNext
    Next J
    
    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.Fields!CHILD_VALUE = SUM
    End If
    
    Adodc1.Recordset.Fields!total_value = SUM + Text7.text + Text5.text + Text8.text + Text21.text + Text9.text + Text10.text - Text13.text + Text18.text
    Adodc1.Recordset.update
    
    Adodc2.Recordset.MoveFirst
End Function

Private Sub Command9_Click()
    On Error GoTo ll

    Label32.Caption = val(Mid(Text22.text, 1, 4)) - val(Mid(Text24.text, 1, 4))
ll:

    If Text24.text = "" Then Label32.Caption = ""
End Sub

Private Sub Form_Activate()
    Dim i, J As Integer
    check = 1

    If yy = 0 Then
        yy = 1
        Dim SUM As Single

        If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.MoveFirst
        End If

        For i = 1 To Adodc1.Recordset.RecordCount

            SUM = 0

            If Adodc1.Recordset.Fields!CHILD_ID = 0 Then
                Adodc2.CommandType = adCmdText
                Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,fines_value,fines_value1 ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND  MEMBER_ID=" & Text1.text
                Adodc2.Refresh
                Text20.text = Adodc2.Recordset.RecordCount + 1

                For J = 1 To Adodc2.Recordset.RecordCount
                    SUM = SUM + Adodc2.Recordset.Fields!MEMBER_VALUE + Adodc2.Recordset.Fields!member_card_value + Adodc2.Recordset.Fields!fines_value '+ Adodc2.Recordset.Fields!activity_value
    
                    Adodc2.Recordset.MoveNext
                Next J
 
            Else
                Text20.text = 1
     
            End If

            If Adodc1.Recordset.RecordCount > 0 Then
                Adodc1.Recordset.Fields!CHILD_VALUE = SUM
            End If
      
            If Text18.text = "" Then Text18.text = 0
            Adodc1.Recordset.Fields!total_value = SUM + Text7.text + Text5.text + Text8.text + Text21.text + Text9.text + Text10.text - Text13.text + Text18.text + Text25.text + Text26.text
            Adodc1.Recordset.update

            Adodc1.Recordset.MoveNext
        Next i

        If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.MoveFirst
        End If
    End If

End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string

    check = 0

    yy = 0

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM OPERATIONS where PAYED=0 and operation_type=   ' ÃœÌœ ⁄÷ÊÌ…'"
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
    End If

    connection_string = Cn.ConnectionString
    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value  ,FINES_value ,PAYED ,FINES_value1 FROM OPERATION_DETAILS  WHERE MEMBER_ID=0  AND PAYED =0 "
    Adodc2.Refresh

    connection_string = Cn.ConnectionString
    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from  Fine_TYPES "
    Adodc3.Refresh

    connection_string = Cn.ConnectionString
    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "select * from  ready_to_print "
    Adodc4.Refresh

    connection_string = Cn.ConnectionString
    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select * from MEMBER_CHILD  "
    Adodc5.Refresh

    connection_string = Cn.ConnectionString
    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText
    Adodc6.RecordSource = "select * from  INSTALLMENT_DETAILS "
    Adodc6.Refresh

    connection_string = Cn.ConnectionString
    Adodc7.ConnectionString = connection_string
    Adodc7.CommandType = adCmdText
    Adodc7.RecordSource = "select * from FINES_DETAILS "
    Adodc7.Refresh

    connection_string = Cn.ConnectionString
    Adodc8.ConnectionString = connection_string
    Adodc8.CommandType = adCmdText
    Adodc8.RecordSource = "select * from  FINES_DETAILS "
    Adodc8.Refresh

End Sub

Private Sub Text10_Change()

    If IsNumeric(Text10.text) Then
 
    Else
        Text10.text = 0
    End If

End Sub

Private Sub Text13_Change()

    If IsNumeric(Text13.text) Then

    Else
        Text13.text = 0
    End If

End Sub

Private Sub Text21_Change()

    If IsNumeric(Text2.text) Then
 
        Adodc1.Recordset.Fields!FINES_TOTAL1 = Text21.text

    End If

End Sub

Private Sub Text3_Change()

    If Text1.text = "" Then Exit Sub
    Adodc2.ConnectionString = Cn.ConnectionString
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,fines_value,fines_value1 ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND  MEMBER_ID=" & Text1.text
    Adodc2.Refresh
    Command9_Click
 
End Sub

Private Sub Text5_Change()

    If IsNumeric(Text5.text) Then
        Adodc1.Recordset.Fields!INSTALLMENTS_TOTAL = Text5.text
    End If

End Sub

Private Sub Text8_Change()

    If IsNumeric(Text8.text) Then
        Adodc1.Recordset.Fields!FINES_TOTAL = Text8.text

    End If

End Sub

