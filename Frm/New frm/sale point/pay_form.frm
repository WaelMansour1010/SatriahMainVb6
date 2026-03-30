VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form pay_form 
   BackColor       =   &H80000008&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "—Ã«¡ «œŒ«· «·„»·€ «·„œ›Ê⁄"
   ClientHeight    =   7785
   ClientLeft      =   -195
   ClientTop       =   375
   ClientWidth     =   11685
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "«·—ﬁ„"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   2760
         TabIndex        =   47
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "pay_form.frx":0000
      Left            =   3960
      List            =   "pay_form.frx":0010
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Text            =   "‰ﬁœÌ"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   0
      Left            =   4800
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   1
      Left            =   3840
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   2
      Left            =   4800
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   3
      Left            =   5760
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   4
      Left            =   3840
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   5
      Left            =   4800
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   6
      Left            =   5760
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   7
      Left            =   3840
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   8
      Left            =   4800
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   9
      Left            =   5760
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   10
      Left            =   3840
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00808080&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   5760
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   ".75"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   11
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   ".5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   12
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   ".25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Index           =   13
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "«€·«ﬁ «·‘«‘…"
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Index           =   0
      Left            =   8520
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   3080
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Index           =   6
      Left            =   8520
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   3080
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Index           =   5
      Left            =   8520
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   3080
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Index           =   4
      Left            =   8520
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   3080
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Index           =   3
      Left            =   8520
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   3080
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Index           =   2
      Left            =   8520
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   3075
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Index           =   1
      Left            =   8520
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   3075
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   " ‰›Ì–"
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox payed_money 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1000
      Left            =   3960
      TabIndex        =   0
      Top             =   2400
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4320
      Top             =   7200
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=resturant"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=resturant"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "bill"
      Caption         =   "«·›« Ê—…"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Height          =   375
      Left            =   -360
      Top             =   8400
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=resturant"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=resturant"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "users"
      Caption         =   "«·„ÊŸ›Ì‰"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc tables 
      Height          =   375
      Left            =   -240
      Top             =   7800
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=resturant"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=resturant"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tables"
      Caption         =   "table"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Height          =   375
      Left            =   -600
      Top             =   10200
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=resturant"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=resturant"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"pay_form.frx":003A
      Caption         =   "waiter"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6840
      TabIndex        =   43
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«Œ — «”„ «·ÊÌ —"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   -360
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   6240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label table_id 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   720
      TabIndex        =   41
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·ÿ«Ê·…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2400
      TabIndex        =   40
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«Œ — «”„ «·„÷Ì›"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   -480
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   5280
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label TAXES_NOTE 
      Alignment       =   1  'Right Justify
      Caption         =   "Label9"
      Height          =   495
      Left            =   -480
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label TAXES_VALUE 
      Alignment       =   1  'Right Justify
      Caption         =   "Label8"
      Height          =   495
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   9000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label SHIFT_NO 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label USER_NAME 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label USER_ID 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   375
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label discount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label discount_note 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·›« Ê—…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5280
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label bill_id 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   3600
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label result 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1005
      Left            =   3960
      TabIndex        =   5
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«·„ »ﬁÌ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label bill_total 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1000
      Left            =   3960
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "’«›Ì «·›« Ê—…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6840
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "«·„»·€ «·„œ›Ê⁄"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
End
Attribute VB_Name = "pay_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()

    If Combo1.ListIndex <> 0 Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False

    End If

End Sub

Private Sub Command1_Click()
    Unload Me
 
    x = MsgBox("Â· «‰  „ √ﬂœ „‰ ⁄„·Ì… «·œ›⁄", vbExclamation + vbYesNo)

    If x = vbNo Then

        Exit Sub
    Else
        'Unload customer_screen
        'Load customer_screen
        'customer_screen.Show
        Unload Me
        Exit Sub

    End If

    If val(payed_money.text) >= val(bill_total.Caption) Then
        result.Caption = payed_money.text - bill_total.Caption
        Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from bill where bill_id=" & bill_id
        Adodc1.Refresh

        Adodc1.Recordset.Fields!total = bill_total.Caption
        Adodc1.Recordset.Fields!payed_money = payed_money.text
        Adodc1.Recordset.Fields!result = result.Caption
        'Adodc1.Recordset.Fields!NET = NET.Caption

        Adodc1.Recordset.Fields!TAXES_VALUE = TAXES_VALUE.Caption
        Adodc1.Recordset.Fields!TAXES_NOTE = TAXES_NOTE.Caption

        Adodc1.Recordset.Fields!discount_value = discount.Caption
        Adodc1.Recordset.Fields!Notes = discount_note.Caption
        Adodc1.Recordset.Fields!host_name = DataCombo1.text
        Adodc1.Recordset.Fields!waiter_name = DataCombo2.text
        Adodc1.Recordset.Fields!table_id = table_id.Caption

        Adodc1.Recordset.update
        'If order_form.table_id <> "" Then

        'tables.CommandType = adCmdText
        'tables.RecordSource = "select * from tables where table_id=" & order_form.table_id
        'tables.Refresh

        'tables.Recordset.Fields!bill_id = 0
        'tables.Recordset.Fields!Status = 0
        'tables.Recordset.Update
        'Adodc1.Recordset.Fields!table_id = order_form.table_id
        'Adodc1.Recordset.Update
        'End If
        'If frmprint.Visible = True Then
        'Unload frmprint
        'End If
        'frmprint.case_id = 0
        'frmprint.bill_id.Caption = bill_id.Caption
        'frmprint.Show
        'MsgBox " „  ‰›Ì– «·ÿ·» Ê «·ÿ»«⁄…", vbInformation

        Form3.Show
        Form3.case_id = 7
        Form3.bill_id = bill_id.Caption
        Form3.user_id = user_id.Caption
        Form3.user_name.Caption = user_name.Caption
        Form3.shift_no.Caption = shift_no.Caption

        'Unload order_form
        'order_form.Show
        'order_form.USER_ID.Caption = USER_ID.Caption
        'order_form.USER_NAME.Caption = USER_NAME.Caption
        'order_form.SHIFT_NO.Caption = SHIFT_NO.Caption
        'Call order_form.new_bill
        Unload Me
    Else
        MsgBox "·«»œ „‰ «œŒ«· „»·€ «ﬂ»— „‰ ﬁÌ„… «·›« Ê—…", vbCritical
        payed_money.text = ""
    End If

End Sub

Function check()

    For i = 0 To 6

        If val(Command2(i).Caption) < val(bill_total.Caption) Then
            Command2(i).Enabled = False
        End If

    Next i

End Function

Private Sub Command2_Click(Index As Integer)
    payed_money.text = Command2(Index).Caption
    payed_money.SetFocus
    'Command1_Click
End Sub

Private Sub Command2_KeyPress(Index As Integer, _
                              KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        order_form.Show
        Unload Me
    End If

End Sub

Private Sub Command3_Click()
    'order_form.Show
    Unload Me
End Sub

Private Sub Command4_Click(Index As Integer)
    payed_money.text = payed_money.text + Command4(Index).Caption
End Sub

Private Sub Command5_Click()
    payed_money.text = ""
    result.Caption = ""
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Form_Activate()

    'payed_money.SetFocus
    If discount_note.Caption = "100%" Then
        Label8.Visible = True
        DataCombo1.Visible = True
    End If

    'If order_form.table_id <> "" Then
    'DataCombo2.Visible = True
    'Label10.Visible = True

    'Me.table_id.Caption = order_form.table_id.Caption
    'End If
    check

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        order_form.Show
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
End Sub

Private Sub payed_money_Change()

    If payed_money.text <> "" Or val(payed_money.text) <> 0 Then
        If val(payed_money.text) >= val(bill_total.Caption) Then
            result.Caption = val(payed_money.text) - val(bill_total.Caption)
        Else
            result.Caption = ""
        End If

    End If

End Sub

Private Sub payed_money_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        order_form.Show
        Unload Me
    End If

    If KeyAscii = 13 Then
        Command1_Click
    End If

End Sub

Private Sub result_Change()
    'result.Caption = Round(result, 2)
End Sub

