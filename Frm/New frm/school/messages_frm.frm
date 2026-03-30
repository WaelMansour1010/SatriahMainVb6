VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form messages_frm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÔÇÔÉ ÑÓÇÆá ÇáÚãáÇÁ"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7560
   Icon            =   "messages_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   7560
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   0
      Width           =   7545
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   20
            Tag             =   "ãä ÝÖáß ÃÏÎá ÑÞã ÇáÞÖíÉ"
            Top             =   15
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483624
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "ÇáãÓÊÎÏã"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   13
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   45
            Width           =   855
         End
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "modflag"
         Top             =   90
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   510
         Visible         =   0   'False
         Width           =   945
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   3120
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "messages_frm.frx":000C
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "messages_frm.frx":03A6
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "messages_frm.frx":0740
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "messages_frm.frx":0ADA
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "messages_frm.frx":0E74
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "messages_frm.frx":120E
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "messages_frm.frx":15A8
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "messages_frm.frx":1B42
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÔÇÔÉ ÑÓÇÆá ÇáÚãáÇÁ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   3975
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   90
         Width           =   3240
      End
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   240
      TabIndex        =   12
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
      MICON           =   "messages_frm.frx":1EDC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text2 
      DataField       =   "id"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   6480
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1560
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8070
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "ÇáÕÇÏÑ"
      TabPicture(0)   =   "messages_frm.frx":1EF8
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ALLButton2"
      Tab(0).Control(1)=   "DataGrid2"
      Tab(0).Control(2)=   "Text4"
      Tab(0).Control(3)=   "DataGrid4"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "ÇáæÇÑÏ"
      TabPicture(1)   =   "messages_frm.frx":1F14
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(1)=   "Text3"
      Tab(1).Control(2)=   "DataGrid3"
      Tab(1).Control(3)=   "ALLButton1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "ÑÓÇáÉ ÌÏíÏÉ"
      TabPicture(2)   =   "messages_frm.frx":1F30
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DataCombo1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Text1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin ALLButtonS.ALLButton ALLButton1 
         Height          =   255
         Left            =   -71640
         TabIndex        =   14
         Top             =   2040
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   255
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "messages_frm.frx":1F4C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "messages_frm.frx":1F68
         Height          =   1455
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2566
         _Version        =   393216
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
            DataField       =   "to"
            Caption         =   "To"
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
            DataField       =   "message"
            Caption         =   "Subject"
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
            DataField       =   "date"
            Caption         =   "Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "time"
            Caption         =   "Time"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "H.mm.ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   4
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "from"
            Caption         =   "from"
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
            DataField       =   "to"
            Caption         =   "to"
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
            DataField       =   "message"
            Caption         =   "ÇáãæÖæÚ"
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
            DataField       =   "sender_show"
            Caption         =   "sender_show"
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
            DataField       =   "reciever_show"
            Caption         =   "reciever_show"
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
            DataField       =   "recived"
            Caption         =   "recived"
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
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "messages_frm.frx":1F7D
         Height          =   1455
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2566
         _Version        =   393216
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
            DataField       =   "from"
            Caption         =   "From"
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
            DataField       =   "message"
            Caption         =   "Subject"
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
            DataField       =   "date"
            Caption         =   "Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "time"
            Caption         =   "Time"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "H.mm.ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   4
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "from"
            Caption         =   "from"
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
            DataField       =   "to"
            Caption         =   "to"
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
            DataField       =   "message"
            Caption         =   "ÇáãæÖæÚ"
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
            DataField       =   "sender_show"
            Caption         =   "sender_show"
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
            DataField       =   "reciever_show"
            Caption         =   "reciever_show"
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
            DataField       =   "recived"
            Caption         =   "recived"
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
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text4 
         DataField       =   "message"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74880
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2400
         Width           =   7095
      End
      Begin VB.TextBox Text3 
         DataField       =   "message"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74880
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2400
         Width           =   7095
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "messages_frm.frx":1F92
         Height          =   1455
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2566
         _Version        =   393216
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
            DataField       =   "from"
            Caption         =   "ãä"
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
            DataField       =   "message"
            Caption         =   "ÇáãæÖæÚ"
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
            DataField       =   "date"
            Caption         =   "ÇáÊÇÑíÎ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "time"
            Caption         =   "ÇáæÞÊ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "H.mm.ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   4
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "from"
            Caption         =   "from"
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
            DataField       =   "to"
            Caption         =   "to"
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
            DataField       =   "message"
            Caption         =   "ÇáãæÖæÚ"
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
            DataField       =   "sender_show"
            Caption         =   "sender_show"
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
            DataField       =   "reciever_show"
            Caption         =   "reciever_show"
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
            DataField       =   "recived"
            Caption         =   "recived"
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
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÇÑÓÇá"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1560
         Width           =   4935
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "messages_frm.frx":1FA7
         Height          =   420
         Left            =   720
         TabIndex        =   1
         Top             =   1080
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   741
         _Version        =   393216
         ListField       =   "MEMBER_NAME"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "messages_frm.frx":1FBC
         Height          =   1455
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2566
         _Version        =   393216
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
            DataField       =   "to"
            Caption         =   "Çáí"
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
            DataField       =   "message"
            Caption         =   "ÇáãæÖæÚ"
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
            DataField       =   "date"
            Caption         =   "ÇáÊÇÑíÎ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "time"
            Caption         =   "ÇáæÞÊ"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "H.mm.ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   4
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "from"
            Caption         =   "from"
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
            DataField       =   "to"
            Caption         =   "to"
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
            DataField       =   "message"
            Caption         =   "ÇáãæÖæÚ"
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
            DataField       =   "sender_show"
            Caption         =   "sender_show"
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
            DataField       =   "reciever_show"
            Caption         =   "reciever_show"
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
            DataField       =   "recived"
            Caption         =   "recived"
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
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
         EndProperty
      End
      Begin ALLButtonS.ALLButton ALLButton2 
         Height          =   255
         Left            =   -71640
         TabIndex        =   15
         Top             =   2040
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   255
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "messages_frm.frx":1FD1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Çáì"
         Height          =   375
         Left            =   5760
         TabIndex        =   2
         Top             =   1080
         Width           =   375
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   960
      Top             =   6480
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   495
      Left            =   6120
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   495
      Left            =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
End
Attribute VB_Name = "messages_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub ALLButton1_Click()
    On Error Resume Next

    If Adodc3.Recordset.RecordCount = 0 Then Exit Sub
    Adodc3.Recordset.Fields!reciever_show = vbFalse
    Adodc3.Recordset.update
    Adodc3.Refresh

    DataGrid1.Refresh
    DataGrid3.Refresh

    Adodc4.Refresh

    DataGrid2.Refresh
    DataGrid4.Refresh

End Sub

Private Sub ALLButton2_Click()
    On Error Resume Next

    If Adodc4.Recordset.RecordCount = 0 Then Exit Sub
    Adodc4.Recordset.Fields!sender_show = vbFalse
    Adodc4.Recordset.update
    Adodc4.Refresh

    DataGrid2.Refresh
    DataGrid4.Refresh

    Adodc3.Refresh

    DataGrid1.Refresh
    DataGrid3.Refresh

End Sub

'You Must Follow The Letters If They Are Small Write Them In Small Letters If 'They Are Big Write Them In Big Letters So That This Code Works.

Private Sub CMD_language_Click()
    On Error Resume Next

    If CMD_language.Caption = "EN" Then
        my_language = "E"
 
        '''Call Reload(Me)
 
    Else
        my_language = "A"
 
        '''Call Reload(Me)
    End If

End Sub

Private Sub Command1_Click()
    On Error Resume Next

    If DataCombo1.text = "" Then
        MsgBox "ÇÎÊÑ ÇáãÑÓá ÇáíÉ" & Chr(13) & "select to ", vbCritical
        Exit Sub

    End If

    Adodc2.Recordset.AddNew

    Adodc2.Recordset.Fields![from] = current_user_name
    Adodc2.Recordset.Fields![To] = DataCombo1.text
    Adodc2.Recordset.Fields!Message = Text1.text
    Adodc2.Recordset.Fields!Date = Date
    Adodc2.Recordset.Fields!Time = Time

    Adodc2.Recordset.Fields!sender_show = vbTrue
    Adodc2.Recordset.Fields!reciever_show = vbTrue
    Adodc2.Recordset.Fields!recived = vbFalse

    Adodc2.Recordset.update
    MsgBox "Êã ÇÑÓÇá ÇáÑÓÇáÉ" & Chr(13) & "message was sent", vbInformation

    Adodc3.Refresh

    DataGrid1.Refresh
    DataGrid3.Refresh

    Adodc4.Refresh

    DataGrid2.Refresh
    DataGrid4.Refresh

End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF5 Then
        Adodc1.Refresh
        DataCombo1.ReFill
    End If

End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    '   Beep
    'Put This Code on a Command Button Or Anywhere You Want

    'Use Your Own Music File And Put The Inside Te Folder Where The Application is
    'This Is The Correct Code

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    If my_language = "E" Then
        CMD_language.Caption = "ÚÑÈí"
        ALLButton2.Caption = "Delete"
        ALLButton1.Caption = "Delete"
        SSTab1.TabCaption(0) = "Outbox "
        SSTab1.TabCaption(1) = "Inbox"
        SSTab1.TabCaption(2) = "New message"
        DataGrid3.Visible = True
        DataGrid4.Visible = True
        Label3.Visible = True
        Label2.Visible = False
        Me.Caption = "Message Screen"
        Command1.Caption = "Send"
        DataCombo1.RightToLeft = False
        Text1.Alignment = 0
        Text3.Alignment = 0
        Text4.Alignment = 0

    End If

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "SELECT * FROM MEMBERS  "
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT * FROM Messages  "
    Adodc2.Refresh

    Adodc3.ConnectionString = connection_string
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "SELECT  *  FROM  Messages  where   reciever_show=1 and  [to]='" & current_user_name & "' order by id desc"
    Adodc3.Refresh
    DataGrid1.Refresh

    Adodc4.ConnectionString = connection_string
    Adodc4.CommandType = adCmdText
    Adodc4.RecordSource = "SELECT * FROM Messages  where       sender_show=1 and [from]='" & current_user_name & "' order by id desc"
    Adodc4.Refresh

    DataGrid2.Refresh

End Sub

Private Sub Text3_Change()
    On Error Resume Next

    If Adodc3.Recordset.RecordCount > 0 Then
        Adodc3.Recordset.Fields!recived = vbTrue
        Adodc3.Recordset.update
    End If

End Sub

Private Sub Text4_Change()
    'If Adodc4.Recordset.RecordCount > 0 Then
    'Adodc4.Recordset.Fields!recived = vbTrue
    'Adodc4.Recordset.Update
    'End If
End Sub
