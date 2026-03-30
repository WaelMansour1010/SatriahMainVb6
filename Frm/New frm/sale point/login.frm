VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form login 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "login"
   ClientHeight    =   3510
   ClientLeft      =   2370
   ClientTop       =   4245
   ClientWidth     =   6975
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0C0&
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
      ItemData        =   "login.frx":058A
      Left            =   1560
      List            =   "login.frx":059A
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   " Name"
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
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Style"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cashier ID"
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
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   0
         TabIndex        =   12
         Top             =   1440
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   4200
      TabIndex        =   8
      Top             =   0
      Width           =   1335
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·ﬂ«‘Ì—"
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
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " €ÌÌ— ‘ﬂ· «·»—‰«„Ã"
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
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ·„… «·„—Ê—"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·ﬂ«‘Ì—"
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
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
   End
   Begin ALLButtonS.ALLButton cmdconnect 
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "œŒÊ·"
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":05D5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtpassword 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtusername 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "name"
      DataSource      =   "Adodc7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   960
      Top             =   4680
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   1085
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
   Begin ALLButtonS.ALLButton Command1 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Œ—ÊÃ"
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
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":05F1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "«·« ’«·"
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":060D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   3720
      Top             =   4080
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   1085
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
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      ToolTipText     =   "Language  «··€…"
      Top             =   2760
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":0629
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc16 
      Height          =   615
      Left            =   840
      Top             =   6000
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   1085
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   615
      Left            =   0
      Top             =   3960
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   1085
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   615
      Left            =   2400
      Top             =   4440
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   1085
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
   Begin ALLButtonS.ALLButton ALLButton3 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   4320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "reset"
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "login.frx":0645
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   615
      Left            =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   1085
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   615
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   1085
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   615
      Left            =   0
      Top             =   3480
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   1085
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
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function CreateEllipticRgn _
                Lib "gdi32" (ByVal X1 As Long, _
                             ByVal Y1 As Long, _
                             ByVal X2 As Long, _
                             ByVal Y2 As Long) As Long

Private Declare Function SetWindowRgn _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal hRgn As Long, _
                              ByVal bRedraw As Boolean) As Long
Dim rgn As Long
Dim tmp As Long

Private Sub ALLButton1_Click()
    On Error Resume Next
    CONNECTION_FRM.Show
End Sub

Private Sub ALLButton2_Click()
    On Error Resume Next

    If ALLButton2.Caption = "EN" Then
        txtUserName.Alignment = 0
        txtPassword.Alignment = 0

        my_language = "E"
        Frame2.Visible = True
        Frame1.Visible = False
        CmdConnect.Caption = "Login"
        Command1.Caption = "Exit"
        ALLButton1.Caption = "Connectiion"
        ALLButton2.Caption = "⁄—»Ì"
    Else
        my_language = "A"
        Frame1.Visible = True
        Frame2.Visible = False
        CmdConnect.Caption = "œŒÊ·"
        Command1.Caption = "Œ—ÊÃ"
        ALLButton1.Caption = "« ’«·"
        ALLButton2.Caption = "EN"
    End If

End Sub

Private Sub cmdconnect_Click()
    On Error Resume Next
    connection_string = Cn.ConnectionString
  
    Adodc1.ConnectionString = connection_string

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from employees where fullcode='" & Text1.text & "' and password='" & txtPassword.text & "'"
    Adodc1.Refresh
  
    If Adodc1.Recordset.RecordCount > 0 Then
    
    End If
 
End Sub

Private Sub Combo1_Click()
    On Error Resume Next

    Dim x As String

    If Combo1.text = "«·‘ﬂ· «·«Ê·" Then
        x = "Vista"
    End If

    If Combo1.text = "«·‘ﬂ· «·À«‰Ì" Then
        x = "Office2007"
    End If

    If Combo1.text = "«·‘ﬂ· «·À«·À" Then
        x = "WinXP.Luna"
    End If

    If Combo1.text = "«·‘ﬂ· «·—«»⁄" Then
        x = "WinXP.Royale"
    End If

    SaveSetting "Transporter", "Setting", "Style", x

    ''login.SkinFramework.LoadSkin system_path + "\Styles\" + x + ".cjstyles", ""
    
    '  'login.SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics

End Sub

Private Sub Command1_Click()
    On Error Resume Next

    If my_language <> "E" Then
        x = MsgBox("Â·  —Ìœ «·Œ—ÊÃ »«·›⁄·", vbCritical + vbYesNo)

        If x = vbYes Then End
    Else
        x = MsgBox("Confirm exit", vbCritical + vbYesNo)

        If x = vbYes Then End
    End If

End Sub

Private Sub Form_Load()
    On Error Resume Next

    ' rgn = CreateEllipticRgn(0, 0, 350, 250)
    '    tmp = SetWindowRgn(Me.hWnd, rgn, True)
    
    'connection_string = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Transporter"

    'SaveSetting "Transporter", "Setting", "server_name", "127.0.0.1"
    'SaveSetting "Transporter", "Setting", "connection_string", connection_string

    'SaveSetting "Transporter", "Setting", "system_path", App.Path

    system_path = App.path
 
    If Style = "" Then
        'login.SkinFramework.LoadSkin system_path + "\Styles\" + "Vista.cjstyles", ""
    
        'login.SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
    Else
        'login.SkinFramework.LoadSkin system_path + "\Styles\" + Style + ".cjstyles", ""
    
        'login.SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics

    End If
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Text1_Change()
    Adodc7.ConnectionString = connection_string
    Adodc7.CommandType = adCmdText
    Adodc7.RecordSource = "select * from employees where fullcode='" & Text1.text & "'"
    Adodc7.Refresh

End Sub
