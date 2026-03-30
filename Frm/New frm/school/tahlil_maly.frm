VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form tahlil_maly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«· Õ·Ì· «·„«·Ì"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13650
   Icon            =   "tahlil_maly.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   13650
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Text            =   "1"
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ÿ»«⁄Â"
      Height          =   375
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Õ–›"
      Height          =   375
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   " —«Ã⁄"
      Height          =   375
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   " ⁄œÌ·"
      Height          =   375
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   8280
      Width           =   1215
   End
   Begin VB.TextBox txtnumber 
      Alignment       =   1  'Right Justify
      DataField       =   "number"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   7680
      Width           =   1815
   End
   Begin VB.TextBox txtr3 
      Alignment       =   1  'Right Justify
      DataField       =   "r3"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   2280
      TabIndex        =   56
      Top             =   7200
      Width           =   2655
   End
   Begin VB.TextBox txttotal 
      Alignment       =   1  'Right Justify
      DataField       =   "r4"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   3840
      TabIndex        =   55
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtr1 
      Alignment       =   2  'Center
      DataField       =   "r1"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   10560
      TabIndex        =   54
      Top             =   7200
      Width           =   1695
   End
   Begin VB.TextBox txtr2 
      Alignment       =   2  'Center
      DataField       =   "r2"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6960
      TabIndex        =   53
      Top             =   7200
      Width           =   1815
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   50
      Top             =   7680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ”«» «·‰ ÌÃ…"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tahlil_maly.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "Equation_name"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      DataField       =   "details"
      DataSource      =   "Adodc1"
      Height          =   645
      Left            =   5040
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Top             =   1680
      Width           =   7215
   End
   Begin VB.CommandButton Command30 
      Caption         =   "ÃœÌœ"
      Height          =   375
      Index           =   12
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   8280
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   4920
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   794
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
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "account_name"
      Height          =   360
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   4920
      Width           =   4455
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "account_name"
      Height          =   360
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   10680
      TabIndex        =   23
      Top             =   7560
      Width           =   1575
      Begin VB.CommandButton Command5 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   1200
         TabIndex        =   27
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   840
         TabIndex        =   26
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   480
         TabIndex        =   25
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   2280
      TabIndex        =   20
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   4920
      Width           =   2055
      Begin VB.CommandButton Command2 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   1560
         TabIndex        =   46
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   18
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   840
         TabIndex        =   17
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1200
         TabIndex        =   16
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
      Begin VB.CommandButton Command1 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   1560
         TabIndex        =   45
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1200
         TabIndex        =   11
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   840
         TabIndex        =   10
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   9
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "«Õ„«·Ì «·„ﬁ«„"
         ForeColor       =   &H000000FF&
         Height          =   15
         Left            =   0
         TabIndex        =   51
         Top             =   2520
         Width           =   1935
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "tahlil_maly.frx":0028
      Height          =   1335
      Left            =   7080
      TabIndex        =   28
      Top             =   9000
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483648
      HeadLines       =   2
      RowHeight       =   19
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   18
      BeginProperty Column00 
         DataField       =   "subject_no"
         Caption         =   "„"
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
         DataField       =   "subject_date"
         Caption         =   "«·ﬁÌ„…"
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
         DataField       =   "subject"
         Caption         =   "«·«Õ„«·Ì"
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
         DataField       =   "by_employee"
         Caption         =   "«·⁄„·Ì…"
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
         DataField       =   "send"
         Caption         =   "«·«Ã„«·Ì"
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
         DataField       =   "interval_type"
         Caption         =   "·Â  √„Ì‰"
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
         DataField       =   "intterval_count"
         Caption         =   "—ﬁ„ «· √„Ì‰"
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
         DataField       =   "alarm"
         Caption         =   "alarm"
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
         DataField       =   "predect_end_time"
         Caption         =   "predect_end_time"
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
      BeginProperty Column09 
         DataField       =   "actual_end_time"
         Caption         =   "actual_end_time"
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
      BeginProperty Column10 
         DataField       =   "Subject_type"
         Caption         =   "‰Ê€ «·„” ‰œ"
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
      BeginProperty Column11 
         DataField       =   "subject_time"
         Caption         =   "subject_time"
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
      BeginProperty Column12 
         DataField       =   "egra2"
         Caption         =   "egra2"
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
      BeginProperty Column13 
         DataField       =   "Archive_name"
         Caption         =   "Archive_name"
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
      BeginProperty Column14 
         DataField       =   "room_name"
         Caption         =   "room_name"
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
      BeginProperty Column15 
         DataField       =   "box_name"
         Caption         =   "box_name"
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
      BeginProperty Column16 
         DataField       =   "shelf_name"
         Caption         =   "shelf_name"
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
      BeginProperty Column17 
         DataField       =   "folder_name"
         Caption         =   "folder_name"
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
            Object.Visible         =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "tahlil_maly.frx":003D
      DataField       =   "account_no"
      Height          =   315
      Left            =   10200
      TabIndex        =   31
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ListField       =   "Account_Serial"
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "tahlil_maly.frx":0052
      DataField       =   "account_no"
      Height          =   315
      Left            =   10200
      TabIndex        =   35
      Top             =   4920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ListField       =   "Account_Serial"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   120
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Adodc1"
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
      Left            =   120
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Adodc1"
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
      Left            =   120
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Adodc1"
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
      Left            =   120
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Adodc1"
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
      Left            =   120
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Adodc1"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "tahlil_maly.frx":0067
      Height          =   1695
      Left            =   2280
      TabIndex        =   48
      Top             =   3000
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
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
      ColumnCount     =   10
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
      BeginProperty Column02 
         DataField       =   "type"
         Caption         =   "type"
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
         DataField       =   "index"
         Caption         =   "index"
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
         DataField       =   "account_no"
         Caption         =   "—ﬁ„ «·Õ”«»"
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
         DataField       =   "account_name"
         Caption         =   "«”„ «·Õ”«»"
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
         DataField       =   "value"
         Caption         =   "«·ﬁÌ„…"
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
         DataField       =   "credit_or_depit"
         Caption         =   "«·Õ—ﬂ…"
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
      BeginProperty Column08 
         DataField       =   "operator"
         Caption         =   "«·⁄„·Ì…"
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
      BeginProperty Column09 
         DataField       =   "total"
         Caption         =   "total"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   4500.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "tahlil_maly.frx":007C
      Height          =   1695
      Left            =   2280
      TabIndex        =   49
      Top             =   5400
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
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
      ColumnCount     =   10
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
      BeginProperty Column02 
         DataField       =   "type"
         Caption         =   "type"
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
         DataField       =   "index"
         Caption         =   "index"
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
         DataField       =   "account_no"
         Caption         =   "—ﬁ„ «·Õ”«»"
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
         DataField       =   "account_name"
         Caption         =   "«”„ «·Õ”«»"
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
         DataField       =   "value"
         Caption         =   "«·ﬁÌ„…"
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
         DataField       =   "credit_or_depit"
         Caption         =   "«·Õ—ﬂ…"
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
      BeginProperty Column08 
         DataField       =   "operator"
         Caption         =   "«·⁄„·Ì…"
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
      BeginProperty Column09 
         DataField       =   "total"
         Caption         =   "total"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   4500.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   375
      Index           =   0
      Left            =   1185
      TabIndex        =   59
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "tahlil_maly.frx":0091
      ColorButton     =   16777215
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
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   60
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "tahlil_maly.frx":042B
      ColorButton     =   16777215
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
      Height          =   375
      Index           =   1
      Left            =   1710
      TabIndex        =   61
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "tahlil_maly.frx":07C5
      ColorButton     =   16777215
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
      Height          =   375
      Index           =   3
      Left            =   645
      TabIndex        =   62
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "tahlil_maly.frx":0B5F
      ColorButton     =   16777215
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "≈«·Ï"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   67
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label Label20 
      DataField       =   "operator"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10200
      TabIndex        =   58
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "«Õ„«·Ì «·»”ÿ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   52
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‰ ÌÃ… «·‰Â«∆Ì…"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5040
      TabIndex        =   47
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "«Õ„«·Ì «·„ﬁ«„"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8880
      TabIndex        =   41
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "«Õ„«·Ì «·»”ÿ"
      ForeColor       =   &H000000FF&
      Height          =   15
      Left            =   480
      TabIndex        =   40
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   1
      Left            =   8400
      TabIndex        =   37
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   36
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   9120
      TabIndex        =   33
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   32
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‰ ÌÃ…"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄œœ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9360
      TabIndex        =   22
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   12240
      TabIndex        =   21
      Top             =   8880
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "«·—’Ìœ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„ﬁ«„"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   13
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "«·—’Ìœ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ì”ÿ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‘—Õ"
      Height          =   255
      Left            =   12480
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "√”„ «·„⁄«œ·…"
      Height          =   255
      Left            =   8520
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·„⁄«œ·…"
      Height          =   255
      Index           =   0
      Left            =   12480
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«‰‘«¡ «·„⁄«œ·… «·„«·Ì…"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   11880
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "«· Õ·Ì· «·„«·Ì  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "tahlil_maly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private objScript As Object

Private Sub ALLButton1_Click()
    calc_result

    If Adodc1.Recordset.RecordCount > 0 And IsNumeric(Txttotal.text) Then
        Adodc1.Recordset.Fields!r4 = Txttotal.text
        Adodc1.Recordset.update
    End If

End Sub

Function calc_result()
    Dim r1 As Double
    Dim r2 As Double
    Dim r3 As Double
    Dim r4 As Double
    Dim total As Double
    Dim equ1 As String
    Dim equ2 As String
    Dim equ3 As String
    Dim equ4 As String
    Dim value As Double

    On Error Resume Next

    If Adodc2.Recordset.RecordCount > 0 Then
        Adodc2.Recordset.MoveFirst
    End If

    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.MoveFirst
    End If

    Set objScript = CreateObject("MSScriptControl.ScriptControl")
    objScript.Language = "VBScript"
    
    equ1 = "": equ2 = "": equ3 = "": equ4 = ""

    For i = 1 To Adodc2.Recordset.RecordCount
        value = Abs(get_balance(Adodc2.Recordset.Fields!account_no))

        If value < 0 Then

            Adodc2.Recordset.Fields!credit_or_depit = "„œÌ‰"
        Else
            Adodc2.Recordset.Fields!credit_or_depit = "œ«∆‰"

        End If

        Adodc2.Recordset.Fields!value = value

        If Adodc2.Recordset.Fields!Operator <> "=" Then

            If Adodc2.Recordset.Fields!credit_or_depit = "œ«∆‰" Then
                equ1 = equ1 & value & Adodc2.Recordset.Fields!Operator
        
                Adodc1.Recordset.Fields!value = x
        
            Else
                equ1 = equ1 & value * -1 & Adodc2.Recordset.Fields!Operator
            End If

        Else

            If Adodc2.Recordset.Fields!credit_or_depit = "œ«∆‰" Then
                equ1 = equ1 & value
            Else
                equ1 = equ1 & value * -1
            End If

            GoTo ll
        End If

        Adodc2.Recordset.MoveNext
    Next i

ll:

    If Adodc2.Recordset.RecordCount > 0 Then
        Adodc2.Recordset.update
        Adodc2.Refresh
    End If

    For i = 1 To Adodc3.Recordset.RecordCount

        value = Abs(get_balance(Adodc3.Recordset.Fields!account_no))

        If value < 0 Then

            Adodc3.Recordset.Fields!credit_or_depit = "„œÌ‰"
        Else
            Adodc3.Recordset.Fields!credit_or_depit = "œ«∆‰"

        End If

        Adodc3.Recordset.Fields!value = value

        If Adodc3.Recordset.Fields!Operator <> "=" Then

            If Adodc3.Recordset.Fields!credit_or_depit = "œ«∆‰" Then
                equ2 = equ2 & value & Adodc3.Recordset.Fields!Operator
            Else
                equ2 = equ2 & value * -1 & Adodc3.Recordset.Fields!Operator
            End If

        Else

            If Adodc3.Recordset.Fields!credit_or_depit = "œ«∆‰" Then
                equ2 = equ2 & Abs(get_balance(Adodc3.Recordset.Fields!account_no))
            Else
                equ2 = equ2 & value * -1
            End If

            GoTo mm
        End If

        Adodc3.Recordset.MoveNext
    Next i

mm:

    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.update
        Adodc3.Refresh
    End If

    If right(equ1, 1) = "+" Or right(equ1, 1) = "-" Or right(equ1, 1) = "*" Or right(equ1, 1) = "/" Then
        MsgBox "ÌÊÃœ Œÿ√ ›Ì „⁄«œ·… «·»”ÿ «Œ— ”ÿ— ·«»œ «‰ ÌÕ ÊÌ =", vbCritical: Exit Function
    End If

    If right(equ2, 1) = "+" Or right(equ2, 1) = "-" Or right(equ2, 1) = "*" Or right(equ2, 1) = "/" Then
        MsgBox "ÌÊÃœ Œÿ√ ›Ì „⁄«œ·… «·„ﬁ«„ «Œ— ”ÿ— ·«»œ «‰ ÌÕ ÊÌ =", vbCritical: Exit Function
    End If

    r1 = objScript.Eval(equ1)
    r2 = objScript.Eval(equ2)

    txtr1.text = r1
    txtr2.text = r2
    Adodc1.Recordset.update

    If r2 <> 0 Then
        txtr3.text = Round(r1 / r2, 2)
    Else
        'MsgBox " «Ã„«·Ì «·„ﬁ«„ Ì”«ÊÌ ’›— ·–·ﬂ ·« Ì„ﬂ‰ ﬁ”„… «·»”ÿ ⁄·Ï «·„ﬁ«„", vbCritical: Exit Function
    End If

    If txtnumber.text <> "" And Not IsNull(Adodc1.Recordset.Fields!Operator) Then
        equ3 = txtr3.text & Adodc1.Recordset.Fields!Operator & txtnumber.text
        Txttotal.text = objScript.Eval(equ3)
        Txttotal.text = Round(Txttotal.text, 2)
    End If

End Function

Private Sub Command1_Click(Index As Integer)
    On Error Resume Next

    If Text1.text = "" Then MsgBox "ﬁ„ »⁄„· ÃœÌœ «Ê·«", vbCritical
    Adodc2.Recordset.AddNew
    Adodc2.Recordset.Fields!id = Text1.text
    Adodc2.Recordset.Fields!type = 1
    'Adodc2.Recordset.Fields![Index] = Y
    Adodc2.Recordset.Fields!account_no = DataCombo1.text

    Adodc2.Recordset.Fields!account_name = Text4.text

    If val(Text3.text) < 0 Then

        Adodc2.Recordset.Fields!credit_or_depit = "„œÌ‰"
    Else
        Adodc2.Recordset.Fields!credit_or_depit = "œ«∆‰"

    End If

    Adodc2.Recordset.Fields![value] = Abs(Text3.text)

    Adodc2.Recordset.Fields!Operator = Command1(Index).Caption
    'Adodc2.Recordset.Fields!Total = X

    Adodc2.Recordset.update
    Adodc2.Refresh

End Sub

Function get_balance(account_serial As String) As Double
    Dim total_credit As Double
    Dim total_depit As Double
    Dim total As Double
    total_credit = 0: total_depit = 0: total = 0
    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText
    Adodc6.RecordSource = "select sum(DEV_Value) As total_credit from RptLedger_Sub where Credit_Or_Debit=0 and  Account_Serial='" & account_serial & "'"
    Adodc6.Refresh

    If Not IsNull(Adodc6.Recordset.Fields!total_credit) Then
        total_credit = Adodc6.Recordset.Fields!total_credit
    Else
        total_credit = 0
    End If

    Adodc6.RecordSource = "select sum(DEV_Value) As total_depit from RptLedger_Sub where Credit_Or_Debit=1 and  Account_Serial='" & account_serial & "'"
    Adodc6.Refresh

    If Not IsNull(Adodc6.Recordset.Fields!total_depit) Then
        total_depit = Adodc6.Recordset.Fields!total_depit
    Else
        total_depit = 0
    End If

    'Total = total_credit - total_depit
    get_balance = total_credit - total_depit

    '1 œ∆«∆‰
    '2„œÌ‰
End Function

Private Sub Command2_Click(Index As Integer)
    Adodc3.Recordset.AddNew
    Adodc3.Recordset.Fields!id = Text1.text
    Adodc3.Recordset.Fields!type = 2
    'Adodc2.Recordset.Fields![Index] = Y
    Adodc3.Recordset.Fields!account_no = DataCombo2.text
    Adodc3.Recordset.Fields!account_name = Text6.text

    If val(Text5.text) < 0 Then

        Adodc3.Recordset.Fields!credit_or_depit = "„œÌ‰"
    Else
        Adodc3.Recordset.Fields!credit_or_depit = "œ«∆‰"

    End If

    Adodc3.Recordset.Fields![value] = Abs(Text5.text)
    Adodc3.Recordset.Fields!Operator = Command2(Index).Caption
    'Adodc2.Recordset.Fields!Total = X

    Adodc3.Recordset.update
    Adodc3.Refresh
End Sub

Private Sub Command30_Click(Index As Integer)
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields!date_created = Now
    Adodc1.Recordset.MoveLast
 
    Adodc2.RecordSource = "select * from  equation_details where type=1 and id=" & Text1.text
    Adodc2.Refresh
 
    Adodc3.RecordSource = "select * from equation_details  where type=2 and id=" & Text1.text
    Adodc3.Refresh

End Sub

Private Sub Command40_Click()
    Adodc1.Recordset.update
End Sub

Private Sub Command5_Click(Index As Integer)
    On Error Resume Next

    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
    Adodc1.Recordset.Fields!Operator = Command5(Index).Caption
    Adodc1.Recordset.update
End Sub

Private Sub DataCombo1_Change()
    DataCombo1_Click (0)
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    Text3.text = 0
    Text4.text = 0
    On Error Resume Next

    If DataCombo1.text <> "" Then
        'Text4.text = ""
        'Text5.text = ""
        Adodc5.RecordSource = "select * from accounts where  Account_Serial='" & DataCombo1.text & "'"
        Adodc5.Refresh

        If Adodc5.Recordset.RecordCount > 0 And Not IsNull(Adodc5.Recordset.Fields!account_name) Then

            Text4.text = Adodc5.Recordset.Fields!account_name
            Text3.text = get_balance(DataCombo1.text)
        Else

            If my_language = "E" Then
                MsgBox "error in this account name to fix error goto account index screen", vbCritical
            Else
                MsgBox "Â‰«ﬂ Œÿ√ ›Ì «”„ «·Õ”«» —«Ã⁄ «·œ·Ì· «·„Õ«”»Ì", vbCritical
                DataCombo1.text = ""
            End If
        End If

    End If

End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Acccount_search.Show
        Acccount_search.case_id = 50
    End If

End Sub

Private Sub DataCombo2_Change()
    DataCombo2_Click (0)
End Sub

Private Sub DataCombo2_Click(Area As Integer)
    On Error Resume Next
    Text6.text = ""
    Text5.text = ""

    If DataCombo2.text <> "" Then
        'Text4.text = ""
        'Text5.text = ""
        Adodc5.RecordSource = "select * from accounts where  Account_Serial='" & DataCombo2.text & "'"
        Adodc5.Refresh

        If Adodc5.Recordset.RecordCount > 0 And Not IsNull(Adodc5.Recordset.Fields!account_name) Then

            Text6.text = Adodc5.Recordset.Fields!account_name
            Text5.text = get_balance(DataCombo2.text)
        Else

            If my_language = "E" Then
                MsgBox "error in this account name to fix error goto account index screen", vbCritical
            Else
                MsgBox "Â‰«ﬂ Œÿ√ ›Ì «”„ «·Õ”«» —«Ã⁄ «·œ·Ì· «·„Õ«”»Ì", vbCritical
                DataCombo1.text = ""
            End If
        End If

    End If

End Sub

Private Sub DataCombo2_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Acccount_search.Show
        Acccount_search.case_id = 60
    End If

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, _
                            Shift As Integer)

    Dim x As Integer
      
    If KeyCode = vbKeyDelete Then
           
        x = MsgBox("Â· «‰  „ √ﬂœ „‰ «·Õ–›", vbCritical + vbYesNo)
         
        If x = vbNo Then
            Exit Sub
        End If

        If Adodc2.Recordset.RecordCount > 0 Then
            Adodc2.Recordset.delete
            Adodc2.Refresh
            calc_result
        Else
            MsgBox "·« ÌÊÃœ „« Ì„ﬂ‰ Õ–›…", vbCritical
        End If
        
    End If
 
End Sub

Private Sub DataGrid2_KeyUp(KeyCode As Integer, _
                            Shift As Integer)
    Dim x As Integer
      
    If KeyCode = vbKeyDelete Then
           
        x = MsgBox("Â· «‰  „ √ﬂœ „‰ «·Õ–›", vbCritical + vbYesNo)
         
        If x = vbNo Then
            Exit Sub
        End If

        If Adodc3.Recordset.RecordCount > 0 Then
            Adodc3.Recordset.delete
            Adodc3.Refresh
            calc_result
        Else
            MsgBox "·« ÌÊÃœ „« Ì„ﬂ‰ Õ–›…", vbCritical
        End If
        
    End If

End Sub

Private Sub ChangeLang()
    Temp = XPBtnMove(1).left
    XPBtnMove(1).left = XPBtnMove(2).left
    XPBtnMove(2).left = Temp

    Temp = XPBtnMove(0).left
    XPBtnMove(0).left = XPBtnMove(3).left
    XPBtnMove(3).left = Temp

    Me.Caption = "Financial Analysis"
    Label1.Caption = Me.Caption
    Label2.Caption = "Create Financial equation"
    Label4.Caption = "Equation Name "
    Label3(0).Caption = "Equation # "
    Command30(12).Caption = "New"
    Command40.Caption = "Save"
    Command3.Caption = "ıEdit"
    Command4.Caption = "Undo"
    Command6.Caption = "Delete"
    Command7.Caption = "Print"

    Label5.Caption = "Description"
    Label6.Caption = "Rugs"
    Label7.Caption = "Account #"
    Label3(3).Caption = "Account Name "
    Label8.Caption = "Balance "
    Label9.Caption = "Primarily"
    Label10.Caption = "Account #"
    Label3(1).Caption = "Account Name "
    Label11.Caption = "Balance "
    Label14.Caption = "Result "
    Label17.Caption = "Final Result "
    Label13.Caption = "Number "
    ALLButton1.Caption = "Cal Result"

    Label19.Caption = "Rugs total "
    Label16.Caption = "Primarily total "

    DataGrid1.Columns(4).Caption = "account #"
    DataGrid1.Columns(5).Caption = "account name"
    DataGrid1.Columns(6).Caption = "value"
    DataGrid1.Columns(7).Caption = "type"
    DataGrid1.Columns(8).Caption = "opr"
    DataGrid1.RightToLeft = False

    DataGrid2.Columns(4).Caption = "account #"
    DataGrid2.Columns(5).Caption = "account name"
    DataGrid2.Columns(6).Caption = "value"
    DataGrid2.Columns(7).Caption = "type"
    DataGrid2.Columns(8).Caption = "opr"
    DataGrid2.RightToLeft = False

    'Command4.Caption = "Close"

End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    connection_string = Cn.ConnectionString

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from equation "
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from  equation_details where type=1 and id=1  "
    Adodc2.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from equation_details  where type=2 and id=1"
    Adodc3.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    'Adodc4.RecordSource = "select * from account_index where black_list=0 and ( account_type='›—⁄Ì' or account_type='sub' )"
    Adodc4.RecordSource = "select * from ACCOUNTS where last_account=1"
    Adodc4.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText
    Adodc5.RecordSource = "select * from ACCOUNTS  where last_account=1"
    Adodc5.Refresh

    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText
    'ALLButton1_Click
    XPBtnMove_Click (3)
End Sub

Private Sub Text1_Change()

    If Text1.text = "" Then Exit Sub
    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from  equation_details where type=1  AND ID=" & val(Text1.text)
    Adodc2.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "select * from equation_details  where type=2 AND ID=" & val(Text1.text)
    Adodc3.Refresh

    If val(Text1.text) <> 1 And val(Text1.text) <> 0 Then
        ALLButton1_Click
    End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap
    'Set RSRet = New ADODB.Recordset
    'Dim strsql As String
    '
    'strsql = "select * From  Notes where NoteType='300' order by NoteID"
    'RSAss.Open strsql, Cn, adOpenStatic, adLockOptimistic, adCmdText

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
