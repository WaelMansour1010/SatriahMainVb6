VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form user_priviliges 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7920
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   11400
      Top             =   2760
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "„” Õœ„ ÃœÌœ"
      TabPicture(0)   =   "user_priviliges.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Adodc5"
      Tab(0).Control(5)=   "Adodc4"
      Tab(0).Control(6)=   "Adodc3"
      Tab(0).Control(7)=   "Adodc2"
      Tab(0).Control(8)=   "Adodc1"
      Tab(0).Control(9)=   "DataCombo1"
      Tab(0).Control(10)=   "Text4"
      Tab(0).Control(11)=   "Command1"
      Tab(0).Control(12)=   "Option3"
      Tab(0).Control(13)=   "Option2"
      Tab(0).Control(14)=   "Option1"
      Tab(0).Control(15)=   "Command2"
      Tab(0).Control(16)=   "Text5"
      Tab(0).Control(17)=   "Text6"
      Tab(0).Control(18)=   "Frame2"
      Tab(0).ControlCount=   19
      TabCaption(1)   =   " ⁄œÌ· ’·«ÕÌ«  „” Õœ„"
      TabPicture(1)   =   "user_priviliges.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Adodc9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Adodc8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Adodc7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Adodc6"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DataCombo2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "DataGrid1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Check1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command4"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command5"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command6"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Command7"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame1"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   1095
         Left            =   -74760
         TabIndex        =   26
         Top             =   1080
         Width           =   10215
         Begin VB.TextBox Text3 
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   0
            PasswordChar    =   "*"
            TabIndex        =   29
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox Text2 
            Height          =   495
            Left            =   3720
            TabIndex        =   28
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Height          =   495
            Left            =   7200
            TabIndex        =   27
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "«·—ﬁ„ «·”—Ì"
            Height          =   495
            Left            =   720
            TabIndex        =   32
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "«”„ «·„” Œœ„"
            Height          =   495
            Left            =   4680
            TabIndex        =   31
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "«”„ «·„ÊŸ› ﬂ«„·«"
            Height          =   495
            Left            =   7800
            TabIndex        =   30
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ã⁄· ’·«ÕÌ«  «·„” Œœ„ „À·"
         Height          =   2175
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   4320
         Width           =   3135
         Begin VB.CommandButton Command8 
            Caption         =   " ‰›Ì–"
            Height          =   495
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   2775
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "user_priviliges.frx":0038
            Height          =   315
            Left            =   360
            TabIndex        =   25
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "username"
            BoundColumn     =   "user_id"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–› Â–« «·„” Œœ„"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Õ›Ÿ"
         Height          =   255
         Left            =   4440
         TabIndex        =   21
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "«⁄ÿ«¡ ﬂ· «·’·«ÕÌ« "
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   2880
         Width           =   2895
      End
      Begin VB.CommandButton Command4 
         Caption         =   "«·€«¡ ﬂ· «·’·«ÕÌ« "
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   " ⁄ÌÌ‰"
         DataField       =   "view"
         DataSource      =   "Adodc7"
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "⁄—÷ ’·«ÕÌ«  «·„” Œœ„"
         Height          =   375
         Left            =   3240
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "user_priviliges.frx":004D
         Height          =   4455
         Left            =   3960
         TabIndex        =   15
         Top             =   2280
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text6 
         DataField       =   "user_id"
         DataSource      =   "Adodc3"
         Height          =   375
         Left            =   -72960
         TabIndex        =   13
         Text            =   "Text6"
         Top             =   3720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text5 
         DataField       =   "no"
         DataSource      =   "Adodc2"
         Height          =   375
         Left            =   -72840
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Õ›Ÿ"
         Height          =   375
         Left            =   -69960
         TabIndex        =   7
         Top             =   4380
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Height          =   195
         Left            =   -68160
         TabIndex        =   6
         Top             =   2460
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Height          =   195
         Left            =   -68160
         TabIndex        =   5
         Top             =   3000
         Width           =   375
      End
      Begin VB.OptionButton Option3 
         Height          =   195
         Left            =   -68160
         TabIndex        =   4
         Top             =   3660
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÃœÌœ"
         Height          =   375
         Left            =   -66120
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         DataField       =   "user_id"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   -70200
         TabIndex        =   1
         Text            =   "Text4"
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "user_priviliges.frx":0062
         Height          =   315
         Left            =   -72720
         TabIndex        =   3
         Top             =   3000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "username"
         BoundColumn     =   "user_id"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   975
         Left            =   -72720
         Top             =   5400
         Visible         =   0   'False
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1720
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
         Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "users"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   975
         Left            =   -75240
         Top             =   4920
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1720
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from screens order by screen_id"
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   975
         Left            =   -75120
         Top             =   3360
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1720
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
         Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "user_priviliges"
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   375
         Left            =   -74880
         Top             =   2640
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "users"
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
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   375
         Left            =   -73800
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "user_priviliges"
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
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "user_priviliges.frx":0077
         Height          =   315
         Left            =   5040
         TabIndex        =   14
         Top             =   840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "username"
         BoundColumn     =   "user_id"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSAdodcLib.Adodc Adodc6 
         Height          =   375
         Left            =   960
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "users"
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
      Begin MSAdodcLib.Adodc Adodc7 
         Height          =   375
         Left            =   1080
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select screen_name,[view] from user_priviliges  where user_id=0"
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
      Begin MSAdodcLib.Adodc Adodc8 
         Height          =   375
         Left            =   600
         Top             =   3360
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "users"
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
      Begin MSAdodcLib.Adodc Adodc9 
         Height          =   375
         Left            =   960
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "users"
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
      Begin VB.Label Label8 
         Caption         =   "«Œ — «”„ «·„” Œœ„"
         Height          =   375
         Left            =   9000
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "«·’·«ÕÌ« "
         Height          =   495
         Left            =   -66600
         TabIndex        =   11
         Top             =   2220
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "’·«ÃÌ«  „À·"
         Height          =   495
         Left            =   -67800
         TabIndex        =   10
         Top             =   3060
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "„œÌ— «·‰Ÿ«„"
         Height          =   495
         Left            =   -67800
         TabIndex        =   9
         Top             =   2460
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   " ÕœÌœ «·’·«ÕÌ«  ·«Õﬁ« "
         Height          =   495
         Left            =   -67800
         TabIndex        =   8
         Top             =   3660
         Width           =   1935
      End
   End
End
Attribute VB_Name = "user_priviliges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Timer1.Enabled = True
Frame2.Enabled = True
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Frame2.Enabled = False

Adodc1.Recordset.Fields!Name = Text1.Text
Adodc1.Recordset.Fields!UserName = Text2.Text
Adodc1.Recordset.Fields!Password = Text3.Text
Adodc1.Recordset.Update


If Option1.Value = True Then 'dmin
For i = 1 To Adodc2.Recordset.RecordCount
Adodc3.Recordset.AddNew
Adodc3.Recordset.Fields!user_id = Adodc1.Recordset.Fields!user_id
Adodc3.Recordset.Fields!screen_name = Adodc2.Recordset.Fields!screen_name
Adodc3.Recordset.Fields!no = Adodc2.Recordset.Fields!no
Adodc3.Recordset.Fields![view] = True
Adodc3.Recordset.Update

Adodc2.Recordset.MoveNext
Next i

End If


If Option2.Value = True Then 'dmin
 Adodc5.CommandType = adCmdText
 Adodc5.RecordSource = "select * from user_priviliges where user_id=" & DataCombo1.BoundText
 Adodc5.Refresh


For i = 1 To Adodc5.Recordset.RecordCount
Adodc3.Recordset.AddNew
Adodc3.Recordset.Fields!user_id = Adodc1.Recordset.Fields!user_id
Adodc3.Recordset.Fields!screen_name = Adodc5.Recordset.Fields!screen_name
Adodc3.Recordset.Fields!no = Adodc5.Recordset.Fields!no
Adodc3.Recordset.Fields![view] = Adodc5.Recordset.Fields![view]
Adodc3.Recordset.Update

Adodc5.Recordset.MoveNext
Next i





End If


Timer1.Enabled = True

MsgBox " „", vbInformation
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Command3_Click()
 Adodc7.CommandType = adCmdText
 Adodc7.RecordSource = "select user_id,screen_name,[view] from user_priviliges  where user_id=" & DataCombo2.BoundText & "order by screen_name"
 Adodc7.Refresh
 DataGrid1.Refresh
End Sub

Private Sub Command4_Click()
If Adodc7.Recordset.RecordCount > 0 Then Adodc7.Recordset.MoveFirst
For i = 1 To Adodc7.Recordset.RecordCount

Adodc7.Recordset.Fields![view] = 0
 Adodc7.Recordset.Update

Adodc7.Recordset.MoveNext
Next i
 
DataGrid1.Refresh
End Sub

Private Sub Command5_Click()
If Adodc7.Recordset.RecordCount > 0 Then Adodc7.Recordset.MoveFirst
For i = 1 To Adodc7.Recordset.RecordCount

Adodc7.Recordset.Fields![view] = 1
 Adodc7.Recordset.Update

Adodc7.Recordset.MoveNext
Next i

DataGrid1.Refresh
End Sub

Private Sub Command6_Click()
Adodc7.Recordset.Update
End Sub

Private Sub Command7_Click()
 
Dim X As Integer
X = MsgBox("Â· «‰  „ √„œ „‰ Â–… «·⁄„·Ì…", vbExclamation + vbYesNo)
If X = vbNo Then
Exit Sub
End If
 
 
 
 Adodc8.CommandType = adCmdText
 Adodc8.RecordSource = "select * from users  where username='" & DataCombo2.Text & "'"
 Adodc8.Refresh
 
If Adodc8.Recordset.RecordCount > 0 Then
Adodc8.Recordset.Delete
Adodc8.Recordset.MoveFirst
DataCombo2.Text = ""
Timer1.Enabled = True
End If

End Sub

Private Sub Command8_Click()

 
 Adodc7.CommandType = adCmdText
 Adodc7.RecordSource = "select user_id,screen_name,[view] from user_priviliges  where user_id=" & DataCombo2.BoundText
 Adodc7.Refresh
 
 If Adodc7.Recordset.RecordCount > 0 Then ' delete old user_priviliges
    For i = 1 To Adodc7.Recordset.RecordCount
    Adodc7.Recordset.Delete
    Adodc7.Recordset.MoveNext
    Next i
 End If
 
 
 
 
 Adodc5.CommandType = adCmdText
 Adodc5.RecordSource = "select user_id,no,screen_name,[view] from user_priviliges where user_id=" & DataCombo3.BoundText
 Adodc5.Refresh
 

For i = 1 To Adodc5.Recordset.RecordCount
Adodc3.Recordset.AddNew
Adodc3.Recordset.Fields!user_id = DataCombo2.BoundText
Adodc3.Recordset.Fields!screen_name = Adodc5.Recordset.Fields!screen_name
Adodc3.Recordset.Fields!no = Adodc5.Recordset.Fields!no
Adodc3.Recordset.Fields![view] = Adodc5.Recordset.Fields![view]
Adodc3.Recordset.Update
Adodc5.Recordset.MoveNext
Next i
 
 Command3_Click
End Sub


Private Sub DataCombo2_Click(Area As Integer)
If DataCombo2.Text <> "" Then
Command3_Click
End If
End Sub

Private Sub DataCombo3_Click(Area As Integer)
If DataCombo2.BoundText = DataCombo3.BoundText Then
MsgBox "·« Ì„ﬂ‰ «Œ Ì«— «·„” Œœ„ Ê‰›”…", vbCritical
DataCombo3.BoundText = ""
DataCombo2.BackColor = &HFF&
DataCombo3.BackColor = &HFF&
Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
Adodc4.Refresh
DataCombo1.ReFill
Adodc6.Refresh
DataCombo2.ReFill
Adodc8.Refresh
DataCombo3.ReFill
Timer1.Enabled = False
End Sub
