VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form mowazna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„Ê«“‰… «· ﬁœÌ—Ì…"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12930
   Icon            =   "mowazna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   12930
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CommandButton Command9 
      Caption         =   "Õ–›"
      Height          =   375
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   " —«Ã⁄"
      Height          =   375
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "»ÕÀ"
      Height          =   375
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   " ⁄œÌ·"
      Height          =   375
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "account_name"
      Height          =   360
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1080
      Width           =   5775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "«· Ê“Ì⁄"
      Height          =   375
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2040
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   1680
      Top             =   2520
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÃœÌœ"
      Height          =   375
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command4 
         Caption         =   "«Œ›«¡"
         Height          =   255
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "mowazna.frx":000C
         Left            =   0
         List            =   "mowazna.frx":001C
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   120
         Width           =   4815
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "mowazna.frx":0073
      Left            =   8160
      List            =   "mowazna.frx":007D
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox TXTcredit 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox TXTdepit 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox TXTtotalvalue 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox TXTFROM 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox TXTTO 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox TXTID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "mowazna.frx":008C
      Height          =   4095
      Left            =   600
      TabIndex        =   22
      Top             =   2640
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   7
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
         DataField       =   "index"
         Caption         =   "„"
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
      BeginProperty Column04 
         DataField       =   "transactions"
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
      BeginProperty Column05 
         DataField       =   "result"
         Caption         =   "«·›—ﬁ"
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
         DataField       =   "do"
         Caption         =   "«·«Ã—«¡"
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
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   4800.189
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   375
      Index           =   0
      Left            =   1185
      TabIndex        =   23
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
      ButtonImage     =   "mowazna.frx":00A1
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
      TabIndex        =   24
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
      ButtonImage     =   "mowazna.frx":043B
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
      TabIndex        =   25
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
      ButtonImage     =   "mowazna.frx":07D5
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
      TabIndex        =   26
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
      ButtonImage     =   "mowazna.frx":0B6F
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "mowazna.frx":0F09
      DataField       =   "account_no"
      Height          =   315
      Left            =   8160
      TabIndex        =   28
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   450
      Left            =   3480
      Top             =   0
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   450
      Left            =   3360
      Top             =   0
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   735
      Left            =   -600
      TabIndex        =   31
      Top             =   2880
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
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
      MICON           =   "mowazna.frx":0F1E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   450
      Left            =   2880
      Top             =   0
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·Õ”«»"
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   11640
      TabIndex        =   30
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·Õ”«»"
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   29
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "«· Ê“Ì⁄ ⁄·Ï «·› —« "
      Height          =   375
      Left            =   11280
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "œ«∆‰"
      Height          =   375
      Index           =   2
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "„œÌ‰"
      Height          =   375
      Index           =   1
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ã„«·Ì «·ﬁÌ„…"
      Height          =   375
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï"
      Height          =   375
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "«·› —… „‰"
      Height          =   375
      Index           =   0
      Left            =   12000
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·”Ì‰«—ÌÊ"
      Height          =   375
      Left            =   11640
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "«·„Ê«“‰… «· ﬁœÌ—Ì…"
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
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "mowazna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As String

Private Sub Command1_Click()
    Adodc1.Recordset.AddNew
    flag = "N"
    ClearData

End Sub

Private Sub Command2_Click()

    If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
 
    If flag = "N" Then
        save_data
        Adodc1.Recordset.MoveLast
    Else
        save_data
    End If

    RETRIVE_DATA

End Sub

Function save_data()
    On Error Resume Next
    Adodc1.Recordset.Fields!Form = txtfrom.text
    Adodc1.Recordset.Fields![To] = txtto.text
    Adodc1.Recordset.Fields!totalvalue = TXTtotalvalue.text

    If TXTdepit.text <> "" Then
        Adodc1.Recordset.Fields!Depit = TXTdepit.text
    End If

    If TxtCredit.text <> "" Then
        Adodc1.Recordset.Fields!Credit = TxtCredit.text
    End If

    Adodc1.Recordset.Fields!account_serial = DataCombo1.text
    Adodc1.Recordset.Fields!account_name = Text2.text

    Adodc1.Recordset.update
End Function

Function ClearData()
    On Error Resume Next

    If Adodc1.Recordset.RecordCount > 0 Then

        txtid.text = ""

        txtfrom.text = ""
        txtto.text = ""
        TXTtotalvalue.text = ""
        TXTdepit.text = ""
        Combo1.text = ""
        Text2.text = ""
        TxtCredit.text = ""
        Adodc2.RecordSource = "select * from mowazna_details WHERE ID=0"
        Adodc2.Refresh
    End If

End Function

Function RETRIVE_DATA()
    On Error Resume Next

    If Adodc1.Recordset.RecordCount > 0 Then

        txtid.text = Adodc1.Recordset.Fields!id

        txtfrom.text = Adodc1.Recordset.Fields!Form
        txtto.text = Adodc1.Recordset.Fields![To]
        TXTtotalvalue.text = Adodc1.Recordset.Fields!totalvalue

        If Not IsNull(Adodc1.Recordset.Fields!Depit) Then
            TXTdepit.text = Adodc1.Recordset.Fields!Depit
        End If

        If Not IsNull(Adodc1.Recordset.Fields!Credit) Then
            TxtCredit.text = Adodc1.Recordset.Fields!Credit
        End If
 
        DataCombo1.text = Adodc1.Recordset.Fields!account_serial
        Text2.text = Adodc1.Recordset.Fields!account_name

        Adodc2.RecordSource = "select * from mowazna_details WHERE ID=" & txtid.text
        Adodc2.Refresh
    End If

    For i = 1 To Adodc2.Recordset.RecordCount

        Adodc2.Recordset.Fields!transactions = get_balance(DataCombo1.text, Adodc2.Recordset.Fields!Index)
        Adodc2.Recordset.Fields!result = Abs(Adodc2.Recordset.Fields!transactions - Adodc2.Recordset.Fields![value])
        Adodc2.Recordset.update
        Adodc2.Recordset.MoveNext

    Next i

End Function

Private Sub Command3_Click()
    Frame1.Visible = True
End Sub

Private Sub Command4_Click()

    If Adodc2.Recordset.RecordCount > 0 Then
        'Adodc2.Recordset.Fields!Do = Combo1.text
        'Adodc2.Recordset.update
        'Adodc2.Refresh
    End If

    Frame1.Visible = False
End Sub

Private Sub Command5_Click()

    If Combo1.ListIndex < 0 Then MsgBox "«Œ —  ‰Ê⁄ «· Ê“Ì⁄ «Ê·« ", vbCritical: Exit Sub

    If Not IsNumeric(txtfrom.text) Or Not IsNumeric(txtto.text) Then MsgBox "·«»œ «‰ ÌﬂÊ‰ „‰ Ê«·Ï «—ﬁ«„ ›ﬁÿ", vbCritical: Exit Sub
    If Not IsNumeric(TXTdepit.text) Then MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… „œÌ‰… ’ÕÌÕ… ", vbCritical: Exit Sub

    If Adodc2.Recordset.RecordCount > 0 Then
        x = MsgBox(" „ «· Ê“Ì⁄ ⁄·Ï «·› —«  „‰ ﬁ»· Â·  —Ìœ Õ–›… Ê⁄„·  Ê“Ì⁄ ÃœÌœ", vbCritical + vbYesNo)

        If x = vbYes Then

            For i = 1 To Adodc2.Recordset.RecordCount
        
                Adodc2.Recordset.delete
                Adodc2.Recordset.MoveNext
            Next i
        
        Else
            Exit Sub
        End If
    End If

    Dim DIFF As Integer
    Dim value As Double
    DIFF = (txtto.text - txtfrom.text) + 1
    value = Round(TXTdepit / DIFF, 2)

    For i = 1 To DIFF

        Adodc2.Recordset.AddNew
        Adodc2.Recordset.Fields!id = txtid.text

        Adodc2.Recordset.Fields![Index] = txtfrom - 1 + i
        Adodc2.Recordset.Fields![value] = value

    Next i

    Adodc2.Recordset.update
    Adodc2.RecordSource = "select * from mowazna_details WHERE ID=" & txtid.text
    Adodc2.Refresh

    DataGrid1.Refresh
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    On Error Resume Next

    If DataCombo1.text <> "" Then
        Text4.text = ""
        Text5.text = ""
        Adodc5.RecordSource = "select * from accounts where  Account_Serial='" & DataCombo1.text & "'"
        Adodc5.Refresh

        If Adodc5.Recordset.RecordCount > 0 And Not IsNull(Adodc5.Recordset.Fields!account_no) Then

            Text2.text = Adodc5.Recordset.Fields!account_name
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
        Acccount_search.case_id = 40
    End If

End Sub

Function get_balance(account_serial As String, Month As Integer) As Double
    Dim total_credit As Double
    Dim total_depit As Double
    Dim total As Double
    total_credit = 0: total_depit = 0: total = 0
    Adodc6.RecordSource = "select sum(DEV_Value) As total_credit from RptLedger_Sub where Month(RecordDate)=" & Month & "and Year(RecordDate) = Year(GetDate()) and Credit_Or_Debit=0 and  Account_Serial='" & account_serial & "'"
    Adodc6.Refresh

    If Not IsNull(Adodc6.Recordset.Fields!total_credit) Then
        total_credit = Adodc6.Recordset.Fields!total_credit
    Else
        total_credit = 0
    End If

    Adodc6.RecordSource = "select sum(DEV_Value) As total_depit from RptLedger_Sub where  Month(RecordDate)=" & Month & "and Year(RecordDate) = Year(GetDate()) and Credit_Or_Debit=1 and  Account_Serial='" & account_serial & "'"
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

Private Sub ChangeLang()
    Temp = XPBtnMove(1).left
    XPBtnMove(1).left = XPBtnMove(2).left
    XPBtnMove(2).left = Temp

    Temp = XPBtnMove(0).left
    XPBtnMove(0).left = XPBtnMove(3).left
    XPBtnMove(3).left = Temp

    Me.Caption = "Budget"
    Label1.Caption = Me.Caption
    Label2.Caption = "Scenario #"
    Label9.Caption = "Account #"
    Label3(3).Caption = "Account Name"
    Label3(0).Caption = "From "
    Label4.Caption = "To"
    Label3(1).Caption = "Depit"
    Label3(2).Caption = "Credit"
    Label5.Caption = "Total Value"
    Command1.Caption = "New"
    Command2.Caption = "Save"
    Command6.Caption = "Edit"
    Command7.Caption = "Search"
    Command8.Caption = "Undo"
    Command9.Caption = "Delete"

    Command5.Caption = "Distribution"
    Label6.Caption = "Distribution in intervals "
    ALLButton1.Caption = "cal Result"

    Combo1.Clear
    Combo1.AddItem "Manual"
    Combo1.AddItem "Auto"
    DataGrid1.Columns(2).Caption = "Index"
    DataGrid1.Columns(3).Caption = "value"
    DataGrid1.Columns(4).Caption = "transactions"
    DataGrid1.Columns(5).Caption = "Result"
    DataGrid1.Columns(6).Caption = "Procedure"
    DataGrid1.RightToLeft = False

    Command4.Caption = "Close"

End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from mowazna  "
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from mowazna_details  "
    Adodc2.Refresh

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
 
    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        RETRIVE_DATA
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

    RETRIVE_DATA
    Exit Sub
ErrTrap:
End Sub
