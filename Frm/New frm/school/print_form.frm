VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form READY_TO_PRINT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                                 ‘«‘…  ÕœÌœ «·„ÿ·Ê» ··ÿ»«⁄… Êÿ»«⁄ …"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   15285
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
   Begin VB.CommandButton Command9 
      Caption         =   " "
      Height          =   735
      Left            =   960
      Picture         =   "print_form.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   " "
      Height          =   735
      Left            =   8520
      Picture         =   "print_form.frx":1992
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H000000FF&
      Caption         =   ">>"
      Height          =   255
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "‰ﬁ· ﬂ· «·ﬂ«—‰ÌÂ«  „‰  «·»Ì«‰«   «·Ã«Â“… ··ÿ»«⁄…  «·Ï «·»Ì«‰«  «·„ «Õ…"
      Top             =   5400
      Width           =   732
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   20
      ToolTipText     =   "‰ﬁ· ﬂ«—‰Ì… Ê«Õœ „‰  «·»Ì«‰«   «·Ã«Â“… ··ÿ»«⁄…  «·Ï «·»Ì«‰«  «·„ «Õ…"
      Top             =   4680
      Width           =   732
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   19
      ToolTipText     =   "‰ﬁ· ﬂ«—‰Ì… Ê«Õœ „‰ «·»Ì«‰«  «·„ «Õ… «·Ï «·»Ì«‰«   «·Ã«Â“… ··ÿ»«⁄…"
      Top             =   4320
      Width           =   732
   End
   Begin VB.CommandButton Command4 
      Caption         =   " €ÌÌ— «·’Ê—…"
      Height          =   255
      Left            =   4440
      TabIndex        =   18
      ToolTipText     =   "«÷€ÿ Â‰« · €ÌÌ— «·’Ê—… «·Œ«’… »«·ﬂ«—‰Ì…"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc2"
      Height          =   288
      Left            =   6120
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   168
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   " „   «·ÿ»«⁄… »‰Ã«Õ"
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   1932
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "<<"
      Height          =   255
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "‰ﬁ· ﬂ· «·ﬂ«—‰ÌÂ«  „‰  «·»Ì«‰«   «·Ã«Â“… ··ÿ»«⁄…  «·Ï «·»Ì«‰«  «·„ «Õ…"
      Top             =   5040
      Width           =   732
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FF0000&
      Caption         =   "ÿ»«⁄‹‹‹‹‹‹‹‹‹‹‹‹…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "«÷€ÿ Â‰« ·ÿ»«⁄… «·ﬂ«—‰ÌÂ«  «·„Õœœ… ··ÿ»«⁄…"
      Top             =   6240
      Width           =   1932
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÿ»«⁄… «·ﬂ·"
      Height          =   192
      Left            =   4800
      TabIndex        =   1
      Top             =   10680
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "print_form.frx":3324
      Height          =   2532
      Left            =   8040
      TabIndex        =   0
      Top             =   3600
      Width           =   6492
      _ExtentX        =   11456
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "inedx"
         Caption         =   "inedx"
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
         DataField       =   "MEMBER_ID"
         Caption         =   "MEMBER_ID"
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
         DataField       =   "MEMBER_NAME"
         Caption         =   "MEMBER_NAME"
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
         DataField       =   "MEMBER_TYPE"
         Caption         =   "MEMBER_TYPE"
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
         DataField       =   "SELECTED"
         Caption         =   "SELECTED"
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
         DataField       =   "update_year"
         Caption         =   "update_year"
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
         DataField       =   "sex"
         Caption         =   "sex"
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
         DataField       =   "OPR_TYPE"
         Caption         =   "OPR_TYPE"
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
         DataField       =   "IMAGE_PATH"
         Caption         =   "IMAGE_PATH"
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
         DataField       =   "PRINTED"
         Caption         =   "PRINTED"
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
         DataField       =   "BILL_NO"
         Caption         =   "BILL_NO"
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
         DataField       =   "CENTER_MANAGER"
         Caption         =   "CENTER_MANAGER"
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
         DataField       =   "RECIVED"
         Caption         =   "RECIVED"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   764.787
         EndProperty
      EndProperty
   End
   Begin DBPIXLib.DBPix20 DBPIX1 
      DataField       =   "IMAGE_PATH"
      DataSource      =   "Adodc3"
      Height          =   1092
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   1931
      _StockProps     =   1
      _Image          =   "print_form.frx":3339
      ImageResampleWidth=   100
      ImageResampleHeight=   100
      ImageResampleMode=   1
      ImageSaveFormat =   0
      JPEGQuality     =   75
      JPEGEncoding    =   0
      JPEGColorMode   =   0
      JPEGNoRecompress=   -1  'True
      JPEGRotateWarning=   0
      PNGColorDepth   =   0
      PNGCompression  =   0
      PNGFilter       =   0
      PNGInterlace    =   1
      ImageDitherMethod=   3
      ImagePaletteMethod=   4
      ImagePreviewMode=   0   'False
      ImageKeepMetaData=   -1  'True
      UseAmbientBackcolor=   -1  'True
      ViewAsyncDecoding=   -1  'True
      ViewEnableMouseZoom=   -1  'True
      ViewInitialZoom =   0
      ViewHAlign      =   1
      ViewVAlign      =   1
      ViewMenuMode    =   0
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11040
      Top             =   120
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   873
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
      RecordSource    =   "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=0"
      Caption         =   ""
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
      Height          =   495
      Left            =   360
      Top             =   1560
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   873
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
      RecordSource    =   "selected_not_printed"
      Caption         =   ""
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
      Bindings        =   "print_form.frx":3351
      Height          =   2532
      Left            =   600
      TabIndex        =   27
      Top             =   3600
      Width           =   6492
      _ExtentX        =   11456
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "inedx"
         Caption         =   "inedx"
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
         DataField       =   "MEMBER_ID"
         Caption         =   "MEMBER_ID"
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
         DataField       =   "MEMBER_NAME"
         Caption         =   "MEMBER_NAME"
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
         DataField       =   "MEMBER_TYPE"
         Caption         =   "MEMBER_TYPE"
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
         DataField       =   "SELECTED"
         Caption         =   "SELECTED"
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
         DataField       =   "update_year"
         Caption         =   "update_year"
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
         DataField       =   "sex"
         Caption         =   "sex"
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
         DataField       =   "OPR_TYPE"
         Caption         =   "OPR_TYPE"
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
         DataField       =   "IMAGE_PATH"
         Caption         =   "IMAGE_PATH"
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
         DataField       =   "PRINTED"
         Caption         =   "PRINTED"
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
         DataField       =   "BILL_NO"
         Caption         =   "BILL_NO"
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
         DataField       =   "CENTER_MANAGER"
         Caption         =   "CENTER_MANAGER"
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
         DataField       =   "RECIVED"
         Caption         =   "RECIVED"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   764.787
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   492
      Left            =   8760
      Top             =   0
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   873
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
      RecordSource    =   "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
      Caption         =   ""
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
      Height          =   492
      Left            =   240
      Top             =   720
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   873
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
      RecordSource    =   "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
      Caption         =   ""
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
   Begin DBPIXLib.DBPix20 DBPix2 
      DataField       =   "IMAGE_PATH"
      DataSource      =   "Adodc3"
      Height          =   1095
      Left            =   1920
      TabIndex        =   36
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   1931
      _StockProps     =   1
      _Image          =   "print_form.frx":3366
      ImageResampleWidth=   100
      ImageResampleHeight=   100
      ImageResampleMode=   1
      ImageSaveFormat =   0
      JPEGQuality     =   75
      JPEGEncoding    =   0
      JPEGColorMode   =   0
      JPEGNoRecompress=   -1  'True
      JPEGRotateWarning=   0
      PNGColorDepth   =   0
      PNGCompression  =   0
      PNGFilter       =   0
      PNGInterlace    =   1
      ImageDitherMethod=   3
      ImagePaletteMethod=   4
      ImagePreviewMode=   0   'False
      ImageKeepMetaData=   -1  'True
      UseAmbientBackcolor=   -1  'True
      ViewAsyncDecoding=   -1  'True
      ViewEnableMouseZoom=   -1  'True
      ViewInitialZoom =   0
      ViewHAlign      =   1
      ViewVAlign      =   1
      ViewMenuMode    =   0
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   2880
      TabIndex        =   35
      Top             =   6240
      Width           =   1812
   End
   Begin VB.Label Label23 
      Caption         =   "⁄œœ «·ﬂ«—‰ÌÂ«  «·„ «Õ… ··ÿ»«⁄…"
      Height          =   372
      Left            =   4920
      TabIndex        =   34
      Top             =   6240
      Width           =   1812
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "«·»Ì«‰«  «·Ã«Â“… ··ÿ»«⁄Â"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   3000
      TabIndex        =   33
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "«·»Ì«‰«  «·„ «Õ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   9960
      TabIndex        =   32
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "„”·”·"
      Height          =   336
      Left            =   840
      TabIndex        =   31
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·ÿ«·»"
      Height          =   336
      Left            =   3840
      TabIndex        =   30
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·ÿ«·»"
      Height          =   336
      Left            =   2160
      TabIndex        =   29
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "«·”‰… «·œ—«”Ì…"
      Height          =   336
      Left            =   5760
      TabIndex        =   28
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "«·”‰… «·œ—«”Ì…"
      Height          =   336
      Left            =   13200
      TabIndex        =   26
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·ÿ«·»"
      Height          =   336
      Left            =   9600
      TabIndex        =   25
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·ÿ«·»"
      Height          =   336
      Left            =   11280
      TabIndex        =   24
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "„”·”·"
      Height          =   336
      Left            =   8280
      TabIndex        =   23
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label12 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "<<<<     ›Ì Õ«·… «·«‰ Â« „‰ «·ÿ»«⁄Â »œÊ‰ „‘«ﬂ· «÷€ÿ ⁄··Ï “—  „  «·ÿ»«⁄… »‰Ã«Õ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   2160
      TabIndex        =   22
      Top             =   7200
      Width           =   6372
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "MEMBER_TYPE"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   5640
      TabIndex        =   14
      Top             =   240
      Width           =   2052
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc3"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   6120
      TabIndex        =   13
      Top             =   720
      Width           =   2052
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc3"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   6000
      TabIndex        =   12
      Top             =   1200
      Width           =   2052
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·ÿ«·»"
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   8280
      TabIndex        =   11
      Top             =   1200
      Width           =   1212
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "inedx"
      DataSource      =   "Adodc3"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   6000
      TabIndex        =   10
      Top             =   1560
      Width           =   2052
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·„”·”·"
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   8280
      TabIndex        =   9
      Top             =   1560
      Width           =   1212
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "BILL_NO"
      DataSource      =   "Adodc3"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   6000
      TabIndex        =   8
      Top             =   1920
      Width           =   2052
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·«Ì’«·"
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   8280
      TabIndex        =   7
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   3960
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "update_year"
      DataSource      =   "Adodc3"
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
      TabIndex        =   6
      Top             =   2160
      Width           =   2532
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "„œÌ— ⁄«„ «·„—ﬂ“"
      DataSource      =   "Adodc3"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   4560
      TabIndex        =   5
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      DataField       =   "CENTER_MANAGER"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   4680
      TabIndex        =   4
      Top             =   2400
      Width           =   1212
   End
End
Attribute VB_Name = "READY_TO_PRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

'Adodc1.Recordset.Update


End Sub

Private Sub cmdprint_Click()
Dim X As Integer
X = MsgBox("Â· «‰  „ √„œ „‰ Â–… «·⁄„·Ì…", vbExclamation + vbYesNo)
If X = vbNo Then
Exit Sub
End If


For i = 1 To Adodc3.Recordset.RecordCount
DBPix2.ImageLoadFile (IMAGE_PATH_FRM.IMAGE_PATH & "\IMAGES\" & Adodc3.Recordset.Fields!image_location & ".JPG")

    DoEvents
Adodc3.Recordset.MoveNext
Next i


Form3.case_id = 2
Form3.Show

If Adodc1.Recordset.RecordCount > 0 Then

Adodc1.Recordset.Update
End If
If Adodc2.Recordset.RecordCount > 0 Then
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Update
End If
Adodc2.Refresh
DataGrid1.Refresh
End Sub

Private Sub Command1_Click()
Dim X, i As Integer
X = MsgBox("Â· «‰  „ √„œ „‰ Â–… «·⁄„·Ì…", vbExclamation + vbYesNo)
If X = vbNo Then
Exit Sub
End If


If Adodc3.Recordset.RecordCount = 0 Then Exit Sub
Adodc3.Refresh
Adodc3.Recordset.MoveFirst

For i = 1 To Adodc3.Recordset.RecordCount
Adodc3.Recordset.Fields!PRINTED = True
Adodc3.Recordset.Fields!DATE_OF_PRINT = DateValue(Now)
Adodc3.Recordset.Update

Adodc3.Recordset.MoveNext
Next i


'Adodc1.CommandType = adCmdText
'Adodc1.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=0 and selected=0"
'Adodc1.Recordset.Update


Adodc1.Refresh
DataGrid1.Refresh

Adodc1.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=0"
  Adodc1.Refresh
  DataGrid1.Refresh
Adodc3.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
  Adodc3.Refresh
  DataGrid2.Refresh
Adodc4.RecordSource = "SELECT COUNT(*) AS TOTAL FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
Adodc4.Refresh

Label24.Caption = Adodc4.Recordset.Fields!TOTAL
End Sub

Private Sub Command3_Click()
Dim X As Integer
X = MsgBox("Â· «‰  „ √„œ „‰ Â–… «·⁄„·Ì…", vbExclamation + vbYesNo)
If X = vbNo Then
Exit Sub
End If

If Adodc1.Recordset.RecordCount > 0 Then
Adodc1.Recordset.MoveFirst
 
End If

Dim i As Integer
For i = 1 To Adodc1.Recordset.RecordCount
Adodc1.Recordset.Fields!Selected = True
Adodc1.Recordset.Update

Adodc1.Recordset.MoveNext
Next i

If Adodc1.Recordset.RecordCount > 0 Then
Adodc1.Recordset.MoveFirst
End If

Adodc1.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=0"
  Adodc1.Refresh
  DataGrid1.Refresh
Adodc3.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
  Adodc3.Refresh
  DataGrid2.Refresh
  Adodc4.RecordSource = "SELECT COUNT(*) AS TOTAL FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
Adodc4.Refresh

Label24.Caption = Adodc4.Recordset.Fields!TOTAL

'For i = 1 To Adodc3.Recordset.RecordCount
'DBPix2.ImageLoadFile (IMAGE_PATH_FRM.IMAGE_PATH  & "\IMAGES\" & Adodc3.Recordset.Fields!image_location & ".JPG")
'
'    DoEvents
'Adodc3.Recordset.MoveNext
'Next i
End Sub

Private Sub Command4_Click()
DBPIX1.ImageLoad
If Adodc3.Recordset.EOF <> True Then
Adodc3.Recordset.MoveNext
Adodc3.Recordset.MovePrevious
Else
Adodc3.Recordset.MoveLast
End If
End Sub

Private Sub Command5_Click()
On Error GoTo ll
  Adodc1.Recordset.Fields!Selected = 1
  Adodc1.Recordset.Update
Adodc1.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=0"
  Adodc1.Refresh
  DataGrid1.Refresh
Adodc3.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
  Adodc3.Refresh
  DataGrid2.Refresh
Adodc4.RecordSource = "SELECT COUNT(*) AS TOTAL FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
Adodc4.Refresh
Label24.Caption = Adodc4.Recordset.Fields!TOTAL
ll:
End Sub

Private Sub Command6_Click()
On Error GoTo ll
'  Adodc1.Recordset.Fields!Selected = 0
'  Adodc1.Recordset.Update
'    DataGrid1.Refresh
  Adodc3.Recordset.Fields!Selected = 0
  Adodc3.Recordset.Update
    DataGrid2.Refresh

Adodc1.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=0"
  Adodc1.Refresh
  DataGrid1.Refresh
Adodc3.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
  Adodc3.Refresh
  DataGrid2.Refresh
  Adodc4.RecordSource = "SELECT COUNT(*) AS TOTAL FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
Adodc4.Refresh
Label24.Caption = Adodc4.Recordset.Fields!TOTAL
ll:
End Sub

Private Sub Command7_Click()

Dim X As Integer
X = MsgBox("Â· «‰  „ √„œ „‰ Â–… «·⁄„·Ì…", vbExclamation + vbYesNo)
If X = vbNo Then
Exit Sub
End If

Dim i As Integer
If Adodc3.Recordset.RecordCount > 0 Then
Adodc3.Recordset.MoveFirst
 
End If

For i = 1 To Adodc3.Recordset.RecordCount
Adodc3.Recordset.Fields!Selected = False
Adodc3.Recordset.Update

Adodc3.Recordset.MoveNext
Next i

 
 If Adodc1.Recordset.RecordCount > 0 Then
Adodc1.Recordset.MoveFirst
 
End If
Adodc1.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=0"
  Adodc1.Refresh
  DataGrid1.Refresh
Adodc3.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
  Adodc3.Refresh
  DataGrid2.Refresh
Adodc4.RecordSource = "SELECT COUNT(*) AS TOTAL FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
Adodc4.Refresh
Label24.Caption = Adodc4.Recordset.Fields!TOTAL
End Sub

 
Private Sub Command8_Click()
X = InputBox("„‰ ›÷·ﬂ «œŒ· —ﬁ„ «·⁄÷Ê ··»ÕÀ ⁄‰…")
Adodc1.RecordSource = "select * from ready_to_print where member_id like'" & X & "%' and RECIVED=0 and PRINTED=0 "
Adodc1.Refresh

If X = "" Then

Adodc1.RecordSource = "select * from ready_to_print where  RECIVED=0 and PRINTED=0"
Adodc1.Refresh
End If
End Sub

Private Sub Command9_Click()
X = InputBox("„‰ ›÷·ﬂ «œŒ· —ﬁ„ «·⁄÷Ê ··»ÕÀ ⁄‰…")
Adodc3.RecordSource = "select * from ready_to_print where member_id like'" & X & "%' and SELECTED=1 and PRINTED=0 "
Adodc3.Refresh

If X = "" Then

Adodc3.RecordSource = "select * from ready_to_print where  SELECTED=1  and PRINTED=0"
Adodc3.Refresh
End If
End Sub

Private Sub Form_Load()
Adodc4.RecordSource = "SELECT COUNT(*) AS TOTAL FROM ready_to_print WHERE PRINTED=0 AND SELECTED=1"
Adodc4.Refresh
Label24.Caption = Adodc4.Recordset.Fields!TOTAL

End Sub
