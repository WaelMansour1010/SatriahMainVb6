VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form operation_from 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ›« Ê—… «·«ﬁ”«ÿ «·Œ“Ì‰…"
   ClientHeight    =   9945
   ClientLeft      =   180
   ClientTop       =   480
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   13770
   Begin VB.TextBox Text28 
      Alignment       =   2  'Center
      DataField       =   "member_type"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7800
      TabIndex        =   69
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text27 
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   66
      Text            =   " "
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "«·›—ﬁ"
      Height          =   375
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   2160
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "€—«„«  «·“ÊÃ…"
      Height          =   2415
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   5880
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         DataField       =   "wife_FINES_TOTAL1"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Text            =   "0"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text25 
         Alignment       =   2  'Center
         DataField       =   "wife_FINES_TOTAL"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label18 
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
         Left            =   2280
         TabIndex        =   64
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label30 
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
         Left            =   2040
         TabIndex        =   63
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "«Œ Ì«—"
      Height          =   375
      Left            =   1440
      TabIndex        =   57
      Top             =   3360
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  'Center
      DataField       =   "FINES_TOTAL1"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text23 
      Alignment       =   2  'Center
      DataField       =   "USER_NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   120
      TabIndex        =   55
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text22 
      DataField       =   "CENTER_MANAGER"
      DataSource      =   "Adodc9"
      Height          =   372
      Left            =   2040
      TabIndex        =   53
      Text            =   "Text22"
      Top             =   3960
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      DataField       =   "bill_no"
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
      Height          =   468
      Left            =   240
      TabIndex        =   51
      Top             =   1080
      Width           =   1572
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      DataField       =   "no_of_person"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   4920
      Width           =   1452
   End
   Begin VB.TextBox Text19 
      BorderStyle     =   0  'None
      DataField       =   "CHILD_ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   1440
      Width           =   1092
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      DataField       =   "ACTIVITY_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   3840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   6480
      TabIndex        =   44
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
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
      Left            =   0
      TabIndex        =   41
      Top             =   2760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      DataField       =   "CHILD_ID"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   2160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "«Œ Ì«—"
      Height          =   375
      Left            =   5640
      TabIndex        =   39
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "«Œ Ì«—"
      Height          =   375
      Left            =   10080
      TabIndex        =   38
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1440
      Width           =   1092
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   6960
      Picture         =   "operatiomn_update_frm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "»ÕÀ"
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "”œ«œ"
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
      Left            =   2280
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8640
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_DATE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_NO"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_TYPE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7560
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
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataField       =   "FINES_TOTAL"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "ID_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      DataField       =   "OTHERS_FEES"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
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
      Height          =   468
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5040
      Width           =   2052
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      DataField       =   "CHILD_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4920
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "«⁄«œ… «·Õ”«Ì« "
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      DataField       =   "discounts"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
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
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text15 
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc4"
      Height          =   615
      Left            =   14760
      TabIndex        =   0
      Text            =   "Text15"
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   492
      Left            =   5400
      Top             =   8640
      Width           =   6972
      _ExtentX        =   12303
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
      Left            =   360
      Top             =   8520
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
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
      Bindings        =   "operatiomn_update_frm.frx":1992
      Height          =   2655
      Left            =   5520
      TabIndex        =   19
      Top             =   6000
      Width           =   8175
      _ExtentX        =   14420
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
         DataField       =   "FINES_name"
         Caption         =   "FINES_name"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   854.929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   2040
      Top             =   9240
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   120
      Top             =   480
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
      LockType        =   4
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
      Height          =   492
      Left            =   840
      Top             =   1680
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
      Height          =   492
      Left            =   1680
      Top             =   2400
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
      Height          =   492
      Left            =   1800
      Top             =   2760
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   492
      Left            =   -840
      Top             =   4440
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
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   495
      Left            =   0
      Top             =   3840
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
      Caption         =   "„œÌ— «·„—ﬂ“"
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
   Begin MSAdodcLib.Adodc Adodc20 
      Height          =   375
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin VB.Label Label3 
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
      Height          =   615
      Left            =   11040
      TabIndex        =   70
      Top             =   1680
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
      Left            =   5400
      TabIndex        =   68
      Top             =   1680
      Width           =   1455
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
      Left            =   3840
      TabIndex        =   67
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   " €—«„… √ŒÌ—"
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
      Left            =   6000
      TabIndex        =   59
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label29 
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
      Left            =   1920
      TabIndex        =   58
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label28 
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
      Left            =   2400
      TabIndex        =   54
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·«Ì’«·"
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
      Left            =   1320
      TabIndex        =   52
      Top             =   1080
      Width           =   2412
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
      Left            =   3960
      TabIndex        =   50
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   7680
      TabIndex        =   49
      Top             =   4800
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
      Left            =   5400
      TabIndex        =   46
      Top             =   3840
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.Label Label15 
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
      Height          =   615
      Left            =   2760
      TabIndex        =   43
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label14 
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
      Height          =   615
      Left            =   2400
      TabIndex        =   42
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   10080
      TabIndex        =   37
      Top             =   1320
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
      Left            =   9720
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
      Left            =   7440
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
      Height          =   612
      Left            =   5880
      TabIndex        =   34
      Top             =   480
      Width           =   1452
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
      Left            =   9960
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
      Height          =   615
      Left            =   10800
      TabIndex        =   32
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   10920
      TabIndex        =   31
      Top             =   3360
      Width           =   1812
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "«·——”Ê„ «·œ—«”Ì…"
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
      Left            =   10080
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
      Height          =   495
      Left            =   9840
      TabIndex        =   29
      Top             =   3960
      Width           =   3615
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
      Height          =   372
      Left            =   3000
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
      Left            =   3360
      TabIndex        =   27
      Top             =   5040
      Width           =   2412
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "«Ã„«·Ì «‘ —«ﬂ«  «· «»⁄Ì‰"
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
      Left            =   10080
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
      Left            =   9600
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
      Left            =   5760
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
      Height          =   735
      Left            =   12480
      TabIndex        =   23
      Top             =   5280
      Width           =   975
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
      Left            =   10440
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
      Height          =   615
      Left            =   5280
      TabIndex        =   21
      Top             =   1080
      Width           =   1455
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
      Height          =   375
      Left            =   11400
      TabIndex        =   20
      Top             =   5640
      Width           =   975
   End
End
Attribute VB_Name = "operation_from"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim check As Integer
Dim yy As Integer

Private Sub Command1_Click()

    Dim dtmTest As Date
 
    If Text3.text = "" Then
        MsgBox "·«ÌÊÃœ ⁄„·Ì«  ·”œ«œÂ«", vbCritical

        Exit Sub
    End If
 
    If Text21.text = "" Then
        MsgBox "  ·«»œ „‰ ﬂ «»… —ﬁ„ «·«Ì’«· «Ê·«  ", vbCritical

        Exit Sub
    End If

    Adodc20.CommandType = adCmdText
    Adodc20.RecordSource = "select * from operations where bill_no=" & Text21.text
    Adodc20.Refresh

    If Adodc20.Recordset.RecordCount > 0 Then
        MsgBox "—ﬁ„ «·«Ì’«·  „” Œœ„ „‰ ﬁ»·", vbCritical
        Exit Sub
    End If

    Dim MEMBER_ID_V, MEMBER_NAME_V As String
    MEMBER_ID_V = Text1.text
    MEMBER_NAME_V = Text6.text

    If Text1.text = "" Then

        MsgBox "·«ÌÊÃœ «Ì ⁄„·Ì« ", vbCritical
        Exit Sub
    End If

    Text3_Change
    Adodc1.Recordset.Fields!payed = 1
    Adodc1.Recordset.Fields!ACTUAL_VALUE = Text11.text
    Adodc1.Recordset.update

    If Me.Caption = "‘«‘… «·⁄÷ÊÌ… «·ÃœÌœ…" Then

        Adodc6.CommandType = adCmdText
        Adodc6.RecordSource = "SELECT * FROM INSTALLMENT_DETAILS  WHERE MEMBER_ID=" & MEMBER_ID_V & "AND INSTALLMENT_NO=1" & "and child_id=" & Text16.text
        Adodc6.Refresh
    
        If Adodc6.Recordset.RecordCount > 0 Then
            Adodc6.Recordset.Fields!payed = 1
            Adodc6.Recordset.Fields!DATE_OF_PAYED = DateValue(Now)
            Adodc6.Recordset.Fields!ACTIVATED = 0
            Adodc6.Recordset.update

            DoEvents

        End If
    End If

    '   Adodc6.CommandType = adCmdText
    '    Adodc6.RecordSource = "SELECT * FROM INSTALLMENT_DETAILS  WHERE MEMBER_ID=" & MEMBER_ID_V & "AND INSTALLMENT_NO=" & Label14.Caption + 1
    '    Adodc6.Refresh
    '    If Adodc6.Recordset.RecordCount > 0 Then
    '    Adodc6.Recordset.Fields!ACTIVATED = 1
    '    Adodc6.Recordset.Update
    '    End If
    
    'End If

    'If Label15.Caption <> "" Then
    '    Adodc7.CommandType = adCmdText
    '    Adodc7.RecordSource = "SELECT * FROM fines_DETAILS  WHERE MEMBER_ID=" & MEMBER_ID_V & " AND fines_no=" & Label15.Caption
    '    Adodc7.Refresh
    '
    '    If Adodc7.Recordset.RecordCount > 0 Then
    '    Adodc7.Recordset.Fields!PAYED = 1
    '    Adodc7.Recordset.Fields!ACTIVATED = 0
    '   Adodc7.Recordset.Fields!PAYED_DATE = Date
    '    Adodc7.Recordset.Update
    '    End If
    '
    '      Adodc7.CommandType = adCmdText
    '    Adodc7.RecordSource = "SELECT * FROM fines_DETAILS  WHERE MEMBER_ID=" & MEMBER_ID_V & "AND fines_no=" & Label15.Caption + 1
    '    Adodc7.Refresh
    '
    '    If Adodc7.Recordset.RecordCount > 0 Then
    '    Adodc7.Recordset.Fields!ACTIVATED = 1
    '    Adodc7.Recordset.Update
    '    End If
    
    'End If

    Dim i As Integer

    'Adodc2.CommandType = adCmdText
    'Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,PAYED,date_of_payed FROM OPERATION_DETAILS  WHERE PAYED =0 AND  MEMBER_ID=14272"
    'Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
        Adodc2.Recordset.MoveFirst

        For i = 1 To Adodc2.Recordset.RecordCount

            Adodc4.Recordset.AddNew
            Adodc5.CommandType = adCmdText
            '  Adodc5.RecordSource = "SELECT * FROM MEMBER_CHILD  WHERE MEMBER_ID=" & MEMBER_ID_V & " AND MEMBER_CHILD_ID=" & Adodc2.Recordset.Fields!CHILD_ID
            Adodc5.RecordSource = "SELECT * FROM MEMBER_CHILD  WHERE MEMBER_ID=" & MEMBER_ID_V & " AND MEMBER_TITLE='" & Adodc2.Recordset.Fields!MEMBER_TITLE & "'"
  
            Adodc5.Refresh
   
            If Adodc5.Recordset.RecordCount > 0 Then Adodc4.Recordset.Fields!Sex = Adodc5.Recordset.Fields!Sex
    
            ' Adodc4.Recordset.Fields!IMAGE_PATH = Adodc5.Recordset.Fields!MEMBER_CHILD_iMAGE_PATH
            ' Adodc4.Recordset.Fields!IMAGE_PATH = LoadPicture(IMAGE_PATH_FRM.IMAGE_PATH  & "\IMAGES\" & Adodc5.Recordset.Fields!image_location & ".JPG")
            If Adodc5.Recordset.RecordCount > 0 Then Adodc4.Recordset.Fields!image_location = Adodc5.Recordset.Fields!image_location
            Adodc4.Recordset.Fields!OPERATION_DATE = DateValue(Now)

            If Text4.text = "⁄÷ÊÌ… ÃœÌœ…" Then
                Adodc4.Recordset.Fields!type = "⁄÷ÊÌ… ÃœÌœ…"
            End If
        
            If Text4.text = " ÃœÌœ ⁄÷ÊÌ…" Then
                Adodc4.Recordset.Fields!type = " ÃœÌœ ⁄÷ÊÌ…"
            End If

            Adodc4.Recordset.Fields!opr_type = "⁄«œÌ"
            Adodc4.Recordset.Fields!bill_no = Text21.text
            Adodc4.Recordset.Fields!CENTER_MANAGER = Text22.text
            Adodc4.Recordset.Fields!member_id = MEMBER_ID_V & "-" & Adodc2.Recordset.Fields!CHILD_ID
            Adodc4.Recordset.Fields!member_name = Adodc2.Recordset.Fields!CHILD_NAME
            Adodc4.Recordset.Fields!MEMBER_TYPE = "⁄«∆·Ì  «»⁄"
            ' Adodc4.Recordset.Fields!User_Name = Main.TxtUserName
            Adodc4.Recordset.Fields!update_year = Text14.text
    
            Adodc2.Recordset.Fields!payed = 1
            Adodc2.Recordset.Fields!DATE_OF_PAYED = DateValue(Now)
            Adodc2.Recordset.update
            Adodc4.Recordset.update
            Adodc2.Recordset.MoveNext
        Next i

    End If

    'Adodc2.CommandType = adCmdText
    'Adodc2.RecordSource = "SELECT * FROM members  WHERE MEMBER_ID=" & MEMBER_ID_V
    'Adodc2.Refresh

    Adodc4.Recordset.AddNew

    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "SELECT * FROM MEMBERS  WHERE MEMBER_ID=" & MEMBER_ID_V
    Adodc5.Refresh
    
    Adodc10.CommandType = adCmdText
    Adodc10.RecordSource = "SELECT * FROM MEMBER_CHILD  WHERE MEMBER_ID=" & MEMBER_ID_V
    Adodc10.Refresh
  
    If Text16.text = "0" Then
        '    Adodc4.Recordset.Fields!sex = Adodc5.Recordset.Fields!sex
        'Adodc4.Recordset.Fields!image_location = Adodc5.Recordset.Fields!image_location
        '  Adodc4.Recordset.Fields!IMAGE_PATH = Adodc5.Recordset.Fields!IMAGE_PATH
        'Adodc4.Recordset.Fields!IMAGE_PATH = LoadPicture(IMAGE_PATH_FRM.IMAGE_PATH  & "\IMAGES\" & Adodc5.Recordset.Fields!image_location & ".JPG")
    
        Adodc4.Recordset.Fields!OPERATION_DATE = DateValue(Now)
        Adodc4.Recordset.Fields!opr_type = "⁄«œÌ"

        If Text4.text = "⁄÷ÊÌ… ÃœÌœ…" Then
            Adodc4.Recordset.Fields!type = "⁄÷ÊÌ… ÃœÌœ…"
        End If
        
        If Text4.text = " ÃœÌœ ⁄÷ÊÌ…" Then
            Adodc4.Recordset.Fields!type = " ÃœÌœ ⁄÷ÊÌ…"
            Adodc5.Recordset.Fields!last_update_year = Text14.text
            Adodc5.Recordset.update
    
            For i = 1 To Adodc10.Recordset.RecordCount
                Adodc10.Recordset.Fields!last_update_year = Text14.text
                Adodc10.Recordset.MoveNext
            Next i
    
        End If

        ' Adodc4.Recordset.Fields!User_Name = Main.TxtUserName
        Adodc4.Recordset.Fields!bill_no = Text21.text
        Adodc4.Recordset.Fields!CENTER_MANAGER = Text22.text
        Adodc4.Recordset.Fields!member_id = MEMBER_ID_V
        Adodc4.Recordset.Fields!member_name = MEMBER_NAME_V
        '    Adodc4.Recordset.Fields!member_type = Adodc5.Recordset.Fields!member_type
        Adodc4.Recordset.Fields!update_year = Text14.text
        Adodc4.Recordset.update
    Else
     
        Adodc5.CommandType = adCmdText
        Adodc5.RecordSource = "SELECT * FROM MEMBER_CHILD  WHERE MEMBER_ID=" & MEMBER_ID_V & " and MEMBER_CHILD_ID=" & Text16.text
        Adodc5.Refresh
    
        Adodc4.Recordset.Fields!Sex = Adodc5.Recordset.Fields!Sex
        Adodc4.Recordset.Fields!OPERATION_DATE = DateValue(Now)
        Adodc4.Recordset.Fields!opr_type = "⁄«œÌ"

        If Text4.text = "⁄÷ÊÌ… ÃœÌœ…" Then
            Adodc4.Recordset.Fields!type = "⁄÷ÊÌ… ÃœÌœ…"
        End If
        
        If Text4.text = " ÃœÌœ ⁄÷ÊÌ…" Then
            Adodc4.Recordset.Fields!type = " ÃœÌœ ⁄÷ÊÌ…"
        End If

        ' Adodc4.Recordset.Fields!User_Name = Main.TxtUserName
        Adodc4.Recordset.Fields!image_location = Adodc5.Recordset.Fields!image_location
        '  Adodc4.Recordset.Fields!IMAGE_PATH = Adodc5.Recordset.Fields!MEMBER_CHILD_iMAGE_PATH
        ' Adodc4.Recordset.Fields!IMAGE_PATH = LoadPicture(IMAGE_PATH_FRM.IMAGE_PATH  & "\IMAGES\" & Adodc5.Recordset.Fields!image_location & ".JPG")
  
        Adodc4.Recordset.Fields!bill_no = Text21.text
        Adodc4.Recordset.Fields!CENTER_MANAGER = Text22.text
        Adodc4.Recordset.Fields!member_id = MEMBER_ID_V & "-" & Text16.text
        Adodc4.Recordset.Fields!member_name = Adodc5.Recordset.Fields!MEMBER_CHILD_NAME

        If Text16.text = "1" Then
            Adodc4.Recordset.Fields!MEMBER_TYPE = "⁄÷Ê ⁄«„·  «»⁄"
        Else
            Adodc4.Recordset.Fields!MEMBER_TYPE = "⁄÷Ê  «»⁄"
        End If

        Adodc4.Recordset.Fields!update_year = Text14.text
        Adodc4.Recordset.update
    
    End If

    Adodc1.Refresh

End Sub

Private Sub Command2_Click()
    Text3_Change
End Sub

Private Sub Command3_Click()
    member_search.Show
    member_search.from = 6

    'X = InputBox("«œŒ· «·—ﬁ„ «Ê Ã“¡ „‰ «·—ﬁ„", "‘«‘… «·»ÕÀ »«·—ﬁ„")

    'select * from operations where PAYED=0 and
    'Adodc1.CommandType = adCmdText
    'Adodc1.RecordSource = "select *  FROM OPERATIONS where   PAYED=0 and MEMBER_ID LIKE'%" & X & "%'" & " and operation_type LIKE'%" & Label25.Caption & "%'"
    'Adodc1.Refresh
    'If Text1.Text = "" Then Text1.Text = 0
    '   Adodc2.CommandType = adCmdText
    '    Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,Fines_value ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND MEMBER_ID=" & Text1.Text
    '    Adodc2.Refresh
End Sub

Private Sub Command4_Click()
    x = InputBox("«œŒ· «·«”„ «Ê Ã“¡ „‰ «·«”„", "‘«‘… «·»ÕÀ »«·«”„")

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select *  FROM OPERATIONS where    PAYED=0 and MEMBER_NAME LIKE'%" & x & "%'" & " and operation_type LIKE'%" & Label25.Caption & "%'"
    Adodc1.Refresh

    If Text1.text = "" Then Text1.text = 0
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT MEMBER_TITLE,op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,Fines_value ,PAYED,date_of_payed FROM OPERATION_DETAILS  WHERE PAYED =0 AND MEMBER_ID=" & Text1.text
    Adodc2.Refresh
End Sub

Private Sub Command5_Click()
    On Error GoTo ll

    Label32.Caption = val(Mid(Text14.text, 1, 4)) - val(Mid(Text27.text, 1, 4))
ll:

    If Text27.text = "" Then Label32.Caption = ""
    '    Adodc2.Recordset.Fields!Fines_NAME = ""
    ' Adodc2.Recordset.Fields!FINES_VALUE = 0
    ' Adodc2.Recordset.Update
    ' Text3_Change
    '    Adodc2.Refresh
    ' DataGrid2.Refresh

End Sub

Private Sub Command6_Click()
    installments_update.Show
    installments_update.Adodc1.CommandType = adCmdText
    installments_update.Adodc1.RecordSource = "select *  FROM Installments where MEMBER_ID =" & Text1.text & " and child_id=" & Text16.text & " ORDER BY CHILD_ID DESC"
    installments_update.Adodc1.Refresh

    installments_update.Adodc2.CommandType = adCmdText
    installments_update.Adodc2.RecordSource = "select INSTALLMENT_NO,INSTALLMENT_VALUE,DATE_OF_PAYED,ACTIVATED,payed  FROM INSTALLMENT_DETAILS where payed=0 and MEMBER_ID =" & Text1.text & "and child_id=" & Text16.text
    installments_update.Adodc2.Refresh

End Sub

Private Sub Command7_Click()
    'fines_update.Show
    'fines_update.Adodc1.CommandType = adCmdText
    'fines_update.Adodc1.RecordSource = "select *  FROM FINES where MEMBER_ID =" & Text1.Text & "and child_id=" & Text16.Text
    'fines_update.Adodc1.Refresh

    'fines_update.Adodc2.CommandType = adCmdText
    'fines_update.Adodc2.RecordSource = "select FINES_NO,FINES_VALUE,PAYED_DATE,ACTIVATED,payed  FROM FINES_DETAILS where payed =0 and MEMBER_ID =" & Text1.Text & "and child_id=" & Text16.Text
    'fines_update.Adodc2.Refresh

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
End Sub

Public Function updatedata()
    Text3_Change
End Function

Private Sub Command9_Click()
    Dim SUM As Single
    Dim i, J As Integer

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
    End If

    For i = 1 To Adodc1.Recordset.RecordCount

        SUM = 0

        If Adodc1.Recordset.Fields!CHILD_ID = 0 Then
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "SELECT MEMBER_TITLE,op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,fines_value ,PAYED,date_of_payed FROM OPERATION_DETAILS  WHERE PAYED =0 AND  MEMBER_ID=" & Text1.text
            Adodc2.Refresh
    
            Text20.text = Adodc2.Recordset.RecordCount + 1

            For J = 1 To Adodc2.Recordset.RecordCount
                SUM = SUM + Adodc2.Recordset.Fields!MEMBER_VALUE + Adodc2.Recordset.Fields!member_card_value + Adodc2.Recordset.Fields!fines_value ' + Adodc2.Recordset.Fields!activity_value
    
                Adodc2.Recordset.MoveNext
            Next J
    
        Else
            Text20.text = 1
     
        End If

        If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.Fields!CHILD_VALUE = SUM
        End If
      
        If Text18.text = "" Then Text18.text = 0
        Adodc1.Recordset.Fields!total_value = SUM + Text7.text + Text5.text + Text8.text + Text24.text + Text9.text + Text10.text - Text13.text + Text18.text + Text25.text + Text26.text
        Adodc1.Recordset.update

        Adodc1.Recordset.MoveNext
    Next i

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
    End If

End Sub

Private Sub Form_Activate()

    check = 1

    If yy = 0 Then
        yy = 1
        Command9_Click
    End If

End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select  *  from operations where PAYED=0  and operation_type=' ÃœÌœ ⁄÷ÊÌ…'  "
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME,MEMBER_TITLE,MEMBER_VALUE,member_card_value  ,PAYED ,FINES_value1,FINES_value FROM OPERATION_DETAILS  WHERE MEMBER_ID=0  AND PAYED =0 "
    Adodc2.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from  Fine_TYPES"
    Adodc3.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "select * from ready_to_print  "
    Adodc4.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select * from  MEMBER_CHILD"
    Adodc5.Refresh

    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText
    Adodc6.RecordSource = "select * from  INSTALLMENT_DETAILS"
    Adodc6.Refresh

    Adodc7.ConnectionString = connection_string
    Adodc7.CommandType = adCmdText
    Adodc7.RecordSource = "select * from   FINES_DETAILS"
    Adodc7.Refresh

    Adodc8.ConnectionString = connection_string
    Adodc8.CommandType = adCmdText
    Adodc8.RecordSource = "select * from FINES_DETAILS "
    Adodc8.Refresh

    Adodc9.ConnectionString = connection_string
    Adodc9.CommandType = adCmdText
    Adodc9.RecordSource = "select * from  CENTER_MANAGER"
    Adodc9.Refresh

    Adodc10.ConnectionString = connection_string
    Adodc10.CommandType = adCmdText
    Adodc10.RecordSource = "select * from  MEMBER_CHILD"
    Adodc10.Refresh

    Adodc20.ConnectionString = connection_string
    Adodc20.CommandType = adCmdText
    Adodc20.RecordSource = "select * from  OPERATIONS"
    Adodc20.Refresh

    check = 0
    yy = 0
End Sub

Private Sub Text13_Change()

    If IsNumeric(Text13.text) Then
        Text3_Change
 
    End If

End Sub

Private Sub Text14_Change()
    Command5_Click
End Sub

Private Sub Text21_Change()

    If Text21.text = "" Then Exit Sub
    If Not IsNumeric(Text21.text) Then
        MsgBox "  —ﬁ„ «·«Ì’«· „ﬂÊ‰ „‰ «—ﬁ«„ ›ﬁÿ    ", vbCritical
        Text21.text = ""
        Exit Sub
    End If

End Sub

Private Sub Text24_Change()
    'Adodc1.Recordset.Fields!FINES_TOTAL1 = Text24.Text
End Sub

Private Sub Text27_Change()
    Command5_Click
End Sub

Private Sub Text3_Change()

    If Text1.text = "" Then Exit Sub

    If Text19.text = "0" Then
        Text19.Visible = False
    Else
        Text19.Visible = True
    End If

    Adodc2.ConnectionString = Cn.ConnectionString
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT MEMBER_TITLE,op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,PAYED,date_of_payed FROM OPERATION_DETAILS  WHERE PAYED =0 AND  MEMBER_ID=" & Text1.text
    Adodc2.Refresh

    Command5_Click

End Sub

Private Sub Text5_Change()

    If IsNumeric(Text5.text) Then
        Adodc1.Recordset.Fields!INSTALLMENTS_TOTAL = Text5.text
        Adodc1.Recordset.update
        Text3_Change
    End If

End Sub

Private Sub Text8_Change()

    If IsNumeric(Text8.text) Then
        Adodc1.Recordset.Fields!FINES_TOTAL = Text8.text
        Adodc1.Recordset.update
        Text3_Change

    End If

End Sub

Private Sub Text9_Change()

    If IsNumeric(Text9.text) Then
        Text3_Change
    End If

End Sub
