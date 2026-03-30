VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form marakes_taklefa_tawze3 
   Caption         =   "  ÇáÊæÒíÚ Úáì ãÑÇßÒ ÇáÊßáÝÉ   "
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12090
   Icon            =   "marakes_taklefa_tawze3.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   12090
   Begin VB.OptionButton Option4 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇíÑÇÏÇÊ"
      Height          =   255
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      Caption         =   "ÚãÇáÉ"
      Height          =   255
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      Caption         =   "ãæÇÏ"
      Height          =   255
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "ãÕÑæÝÇÊ"
      Height          =   255
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAccountSerial 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   615
      Left            =   960
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   50
      Top             =   2880
      Width           =   8895
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "Description"
      DataSource      =   "Adodc3"
      Height          =   495
      Left            =   3720
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   44
      Top             =   7200
      Width           =   7455
   End
   Begin VB.TextBox lineno 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox account_name 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox account_no 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox kedno 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   1080
      TabIndex        =   18
      Top             =   8160
      Width           =   6615
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÞíãÉ"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   6300
         X2              =   6300
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4560
         X2              =   4560
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2830
         X2              =   2830
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ÇÓã ÇáãÑßÒ"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "äæÚ ÇáãÑßÒ"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ÑÞã ÇáãÑßÒ"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3480
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -720
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1080
         X2              =   1080
         Y1              =   120
         Y2              =   720
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   1080
      TabIndex        =   25
      Top             =   8280
      Width           =   6615
      Begin VB.Line Line1 
         Index           =   7
         X1              =   300
         X2              =   300
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -720
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Center #"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4200
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Center  Type"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Center Name"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5640
         TabIndex        =   27
         Top             =   240
         Width           =   1695
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   3780
         X2              =   3780
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   5550
         X2              =   5550
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Center Name"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   10080
      TabIndex        =   17
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   5880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÍÝÙ"
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "marakes_taklefa_tawze3.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton Command2 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÇÏÑÇÌ"
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
      MICON           =   "marakes_taklefa_tawze3.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   9960
      TabIndex        =   14
      Top             =   600
      Width           =   2055
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "ØÑíÞÉ ÇáÊæÒíÚ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇÎÊÑ  ÇáãÔÑæÚ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "ÑÞã ÇáÓØÑ"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "ÑÞã ÇáÍÓÇÈ"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "ÑÞã ÇáÞíÏ"
         Height          =   255
         Left            =   -7920
         TabIndex        =   35
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇÌãÇáí ÇáãÈáÛ"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   3120
      TabIndex        =   11
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
      Begin VB.Label Label7 
         Caption         =   "Amount"
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
         Left            =   -360
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Distribution"
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
         Left            =   -1080
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "account_type"
      DataSource      =   "Adodc4"
      BeginProperty Font 
         Name            =   "Arabic Typesetting"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "account_name"
      DataSource      =   "Adodc4"
      BeginProperty Font 
         Name            =   "Arabic Typesetting"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "marakes_taklefa_tawze3.frx":0044
      Height          =   315
      Left            =   5880
      TabIndex        =   2
      Top             =   2400
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ListField       =   "account_name"
      BoundColumn     =   "code"
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
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "marakes_taklefa_tawze3.frx":0059
      Left            =   7800
      List            =   "marakes_taklefa_tawze3.frx":0063
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Çáí"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "account_no"
      DataSource      =   "Adodc4"
      BeginProperty Font 
         Name            =   "Arabic Typesetting"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   0
      Top             =   600
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
      Caption         =   "ÊÍÑíß"
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
      Height          =   585
      Left            =   0
      Top             =   840
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
      Caption         =   "ÊÍÑíß"
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
      Height          =   585
      Left            =   1320
      Top             =   1080
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
      Caption         =   "ÊÍÑíß"
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
      Height          =   585
      Left            =   0
      Top             =   960
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
      Caption         =   "ÊÍÑíß"
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
      TabIndex        =   23
      ToolTipText     =   "Language  ÇááÛÉ"
      Top             =   0
      Visible         =   0   'False
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
      MICON           =   "marakes_taklefa_tawze3.frx":0072
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
      Height          =   585
      Left            =   -360
      Top             =   7680
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
      Caption         =   "ÊÍÑíß"
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
      Caption         =   "ÊÍÑíß"
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
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   4200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÍÐÝ"
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "marakes_taklefa_tawze3.frx":008E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "marakes_taklefa_tawze3.frx":00AA
      Height          =   315
      Left            =   5880
      TabIndex        =   42
      Top             =   1920
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ListField       =   "Project_name"
      BoundColumn     =   "id"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "marakes_taklefa_tawze3.frx":00BF
      Height          =   2055
      Left            =   960
      TabIndex        =   43
      Top             =   3720
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "cost_center_id"
         Caption         =   "ßæÏ ãÑßÒ ÇáÊßáÝÉ"
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
         DataField       =   "cost_center"
         Caption         =   "ÇÓã ãÑßÒ ÇáÊßáÝÉ"
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
         DataField       =   "Project__code"
         Caption         =   "ßæÏ ÇáãÔÑæÚ"
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
         DataField       =   "Project_name"
         Caption         =   "ÇÓã ÇáãÔÑæÚ"
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
         DataField       =   "value"
         Caption         =   "ÇáÞíãÉ"
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
         DataField       =   "Description"
         Caption         =   "ÇáÔÑÍ"
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
      BeginProperty Column07 
         DataField       =   "opr_id"
         Caption         =   "opr_id"
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
         DataField       =   "opr_type"
         Caption         =   "opr_type"
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
         DataField       =   "account_name"
         Caption         =   "account_name"
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
         DataField       =   "account_no"
         Caption         =   "account_no"
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
         DataField       =   "depit_or_credit"
         Caption         =   "depit_or_credit"
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
         DataField       =   "account_type"
         Caption         =   "account_type"
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
         DataField       =   "line_no"
         Caption         =   "line_no"
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
         DataField       =   "kedno"
         Caption         =   "kedno"
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
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc7 
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
      Caption         =   "ÊÍÑíß"
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
      Height          =   375
      Left            =   5040
      TabIndex        =   57
      Top             =   1920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÇÏÑÇÌ"
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
      MICON           =   "marakes_taklefa_tawze3.frx":00D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTP_Date 
      Height          =   285
      Left            =   6000
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   10383715
      CustomFormat    =   "yyyy/M/d"
      Format          =   98762755
      CurrentDate     =   41640
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇáÔÑÍ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10080
      TabIndex        =   51
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇÎÊÑ ãÑßÒ ÇáäßáÝÉ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9960
      TabIndex        =   47
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "ÇáÔÑÍ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10680
      TabIndex        =   45
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label20 
      Caption         =   "ÇÓã ÇáÍÓÇÈ"
      Height          =   255
      Left            =   6120
      TabIndex        =   41
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "  ÇáÊæÒíÚ Úáì ãÑÇßÒ ÇáÊßáÝÉ              "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label opr_id 
      Height          =   735
      Left            =   7800
      TabIndex        =   33
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label depit_or_credit 
      Height          =   375
      Left            =   6120
      TabIndex        =   32
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label opr_type 
      Height          =   375
      Left            =   6720
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "ÇáãÑßÒ"
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
      Left            =   840
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label type 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "ÍÐÝ"
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
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label value 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   7800
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "marakes_taklefa_tawze3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As String
Dim Check As Boolean
 
Function distribute() As Boolean
    On Error Resume Next
    Check = False
    distribute = False
    Dim markas_value As Double

    If Combo1.Text = "Çáí" Or Combo1.Text = "Auto" Then

        If Adodc3.Recordset.RecordCount > 0 Then
            markas_value = value.Caption / Adodc3.Recordset.RecordCount
            Adodc3.Recordset.MoveFirst
                    
            For i = 1 To Adodc3.Recordset.RecordCount
                Adodc3.Recordset.Fields!value = markas_value
                Adodc3.Recordset.update
                Adodc3.Recordset.MoveNext
            Next i
                
        End If

        'Adodc5.RecordSource = "SELECT * FROM sandat_ked_details  where opr_id =" & Me.opr_id
        'Adodc5.Refresh
        'If Adodc5.Recordset.RecordCount > 0 Then
        'Adodc5.Recordset.Fields!dist = vbTrue
        'Adodc5.Recordset.update
        'End If
        distribute = True
    Else

        Adodc6.RecordSource = "SELECT SUM(VALUE) AS TOTAL FROM marakes_taklefa_temp  where line_no =" & Me.lineno.Text
        Adodc6.Refresh

        If Adodc6.Recordset.RecordCount > 0 Then
            If Round(Adodc6.Recordset.Fields!Total, 0) <> Round(Me.value.Caption, 0) Or IsNull(Adodc6.Recordset.Fields!Total) Then
                Check = True
                distribute = False

                If SystemOptions.UserInterface = EnglishInterface Then
                    X = MsgBox("error in manual distribution" & CHR(13) & "     continue with Automatic distribution ", vbCritical + vbYesNo)
                Else
                    X = MsgBox("åäÇß ÎØÃ Ýí ÇáÊæÒíÚ ÇáíÏæí ÇáãÈáÛ ÛíÑ ãæÒÚ ÈÔßá ÕÍíÍ" & CHR(13) & "  åá ÊÑíÏ ÇÚÇÏÉ ÇáÊæÒíÚ ÇáíÇ", vbCritical + vbYesNo)
                    
                End If
                            
                If X = vbYes Then
                    If Adodc3.Recordset.RecordCount > 0 Then
                        markas_value = value.Caption / Adodc3.Recordset.RecordCount
                        Adodc3.Recordset.MoveFirst
                                                   
                        For i = 1 To Adodc3.Recordset.RecordCount
                            Adodc3.Recordset.Fields!value = markas_value
                            Adodc3.Recordset.update
                            Adodc3.Recordset.MoveNext
                        Next i
                                               
                    End If
                                  
                    distribute = True
                                  
                End If

            Else
                distribute = True
            End If

        End If

    End If

End Function

Private Sub ALLButton1_Click()

    If distribute = True Then
        Unload Me
    End If

End Sub

Private Sub ALLButton2_Click()
    Label32_Click
End Sub

Private Sub ALLButton3_Click()

    If DataCombo2.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
                
            MsgBox "ÍÏÏ ãÔÑæÚ  ÇæáÇ "
        Else
            MsgBox "  Specify project "
        End If
        
        Exit Sub
    End If
        
    Command2_Click
    DataCombo2.BoundText = ""
        
End Sub

Private Sub CMD_language_Click()
    On Error Resume Next

    If CMD_language.Caption = "EN" Then
        my_language = "E"
 
        ''Call Reload(Me)
 
    Else
        my_language = "A"
 
        ''Call Reload(Me)
    End If

End Sub

Private Sub Combo1_Change()

    If Combo1.Text = "Çáí" Or Combo1.Text = "Auto" Then
        DataGrid1.AllowUpdate = False
    Else
        DataGrid1.AllowUpdate = True
    End If

End Sub

Private Sub Combo1_Click()

    If Combo1.Text = "Çáí" Or Combo1.Text = "Auto" Then
        DataGrid1.AllowUpdate = False
    Else
        DataGrid1.AllowUpdate = True
    End If

End Sub

Private Sub Command2_Click()

    If DataCombo1.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
                
            MsgBox "ÍÏÏ ãÑßÒ ÇæáÇ "
        Else
            MsgBox "  Specify cc "
        End If
        
        Exit Sub
        
    End If

    'On Error Resume Next
    If SystemOptions.UserInterface = EnglishInterface Then
        If Combo1.Text = "" Then MsgBox "Select the type of distribution, vbCritical: Exit Sub"
        If DataCombo1.Text = "" And DataCombo2.Text = "" Then MsgBox " select cost center or project", vbCritical: Exit Sub

    Else

        If Combo1.Text = "" Then MsgBox " ÍÏÏ äæÚ ÇáÊæÒíÚ", vbCritical: Exit Sub
        If DataCombo1.Text = "" And DataCombo2.Text = "" Then MsgBox " ÍÏÏ ãÑßÒ Çæ ãÔÑæÚ", vbCritical: Exit Sub
    End If

    Adodc3.Recordset.AddNew

    Adodc3.Recordset.Fields!opr_id = Me.opr_id
    Adodc3.Recordset.Fields!kedno = Me.kedno.Text
    Adodc3.Recordset.Fields!cost_center = DataCombo1.Text
    Adodc3.Recordset.Fields!cost_center_id = DataCombo1.BoundText
    Adodc3.Recordset.Fields!opr_type = Me.opr_type
    Adodc3.Recordset.Fields!depit_or_credit = Me.depit_or_credit.Caption
    Adodc3.Recordset.Fields!description = Text5.Text
    Adodc3.Recordset.Fields!account_no = Me.account_no.Text

    Adodc3.Recordset.Fields!account_name = Me.account_name.Text

    Adodc3.Recordset.Fields!Project__code = Me.DataCombo2.BoundText

    Adodc3.Recordset.Fields!Project_name = Me.DataCombo2.Text

    Adodc3.Recordset.Fields!line_no = Me.lineno.Text
    Adodc3.Recordset.Fields!record_date = DTP_Date.value

    'Adodc3.Recordset.Fields!account_type = Text4.text

    Adodc3.Recordset.update
    Text5.Text = ""

    If Me.Combo1.ListIndex = 0 Or Me.Combo1.ListIndex = -1 Then
        distribute
    End If

    Me.DataCombo1.Text = ""

End Sub

Private Sub DataCombo1_Click(Area As Integer)
    On Error Resume Next

    If DataCombo1.Text = "" Then Exit Sub
    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "select * from markaas_taklefa where account_name ='" & DataCombo1.Text & "' "
          
        
    Adodc4.Refresh
 
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 2
    End If
 
    If KeyCode = vbKeyF5 Then
        Adodc2.Refresh
        DataCombo1.ReFill
    End If

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, _
                            Shift As Integer)
    On Error Resume Next

    'If KeyCode = 13 Then
    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.update
    End If

    'End If
End Sub

Private Sub DataGrid1_LostFocus()
    On Error Resume Next

    'If KeyCode = 13 Then
    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.update
    End If

End Sub

Private Sub Form_Load()
    'On Error Resume Next
    'On Error Resume Next

    '
    Check = False
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        'ChangeLang
        depit_or_credit.Visible = False
   
        CMD_language.ToolTipText = "Change Language"

        'ALLButton1.left = 6120
        ' Command2.left = 6120
 
        '   Frame8.Visible = False
        '
        DataGrid1.RightToLeft = False
 
        CMD_language.Caption = "ÚÑÈí"
        '   value.Alignment = 0
 
        Label21.Caption = "Line#"
        Label19.Caption = "Account#"
        Label4.Caption = "Amount"
        Label9.Caption = "Dist. Type"
        Label22.Caption = "Project"
        Label5.Caption = "Cost Center"
        Label24.Caption = "Desc"
        Label20.Caption = "Account Name"

        ' Combo1.RightToLeft = False
        '  DataCombo1.RightToLeft = False
 
        'Frame3.Visible = True
        'Frame1.Visible = True

        Combo1.Clear
        Combo1.AddItem "Auto"
        Combo1.AddItem "Manual"
        Combo1.Text = "Auto"
      
        Label3.Caption = " cost centers Distribution"
        Me.Caption = Label3.Caption
        ' Label5.Caption = "select cost centers"
        Command2.Caption = "Add"
        ALLButton1.Caption = "Save"
        ALLButton3.Caption = "Add"
        ' ALLButton3.Caption = "Search"
        ALLButton2.Caption = "Delete"
        DataGrid1.Columns(0).Caption = "CC Code"
        DataGrid1.Columns(1).Caption = "CC Name"
        DataGrid1.Columns(2).Caption = "Project Code"
        DataGrid1.Columns(3).Caption = "Project Code"
        DataGrid1.Columns(4).Caption = "Value"
        DataGrid1.Columns(5).Caption = "Des"
 
        ' Label32.Caption = "delete"
  
    End If

    'LoadSettings
    connection_string = Cn.ConnectionString

    'Adodc1.ConnectionString = connection_string
    'Adodc1.CommandType = adCmdText
    'Adodc1.RecordSource = "SELECT * FROM BOXs  "
    'Adodc1.Refresh
 
  Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
 
           StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
           
        If SystemOptions.usertype <> UserAdminAll Then
            'StrSQL = StrSQL & " and  Branch_NO=0 or  Branch_NO=" & Current_branch
             StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
        End If
        
        StrSQL = StrSQL + " Order BY account_name"
    Else
         
        StrSQL = "  SELECT code ,EnglishName FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
        If SystemOptions.usertype <> UserAdminAll Then
            'StrSQL = StrSQL & " and  Branch_NO=0 or   Branch_NO=" & Current_branch
             StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
        End If
            
        StrSQL = StrSQL + " Order BY EnglishName"
End If
  
  
    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = StrSQL
    
    

    
    
    Adodc2.Refresh




    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    'Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where opr_id =" & Me.opr_id
    'Adodc3.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "SELECT * FROM markaas_taklefa   WHERE NOT(account_no IS NULL)  order by account_name "
    Adodc4.Refresh

    Adodc5.ConnectionString = connection_string
    Adodc5.CommandType = adCmdText

    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText

    Adodc7.ConnectionString = connection_string
    Adodc7.CommandType = adCmdText
    Adodc7.RecordSource = "SELECT * FROM projects  where     branch_no in(" & Current_branchSql & ")"
    Adodc7.Refresh

    'For i = 1 To Adodc3.Recordset.RecordCount
 
    'Adodc3.Recordset.Delete
    ' Adodc3.Recordset.MoveNext
    'Next i
    Combo1.ListIndex = 1
    Combo1.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    On Error Resume Next

    If Adodc3.Recordset.RecordCount = 0 Then
        Cancel = False
        Exit Sub
    End If

    If ALLButton1.Enabled = True Then
        ALLButton1_Click
    End If

    If Check = True Then
        Cancel = True
    End If

End Sub

Private Sub Label32_Click()
    On Error Resume Next

    If SystemOptions.UserInterface = EnglishInterface Then

        X = MsgBox("Confirm delete", vbCritical + vbYesNo)
    Else
        X = MsgBox("ÊÃßíÏ ÇáÍÐÝ", vbCritical + vbYesNo)
        
    End If
            
    If X = vbNo Then
        Exit Sub
    End If

    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.delete
        Adodc3.Refresh
        distribute
        'ALLButton1_Click
    End If

End Sub
