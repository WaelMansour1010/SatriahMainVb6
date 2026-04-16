VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Voucher_search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘… «·»ÕÀ ⁄‰ «·ﬁÌÊœ"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13830
   Icon            =   "Voucher_search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   13830
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox txtRowLimt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Text            =   "1000"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox Check6 
      Alignment       =   1  'Right Justify
      Caption         =   "·›—⁄ „Õœœ"
      Height          =   255
      Left            =   12360
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      Alignment       =   1  'Right Justify
      Caption         =   "·Õ—ﬂ… „Õœœ…"
      Height          =   255
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "«·„»·€"
      Height          =   615
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   1200
      Width           =   4095
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   5
         Left            =   0
         TabIndex        =   44
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   3
         Left            =   2040
         TabIndex        =   42
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "«·Ì „»·€"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   45
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„‰ „»·€"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3360
         TabIndex        =   43
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox TxtAccCode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   5
      Left            =   5760
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Height          =   495
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÊŸ›"
         Height          =   195
         Index           =   3
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ê—œ"
         Height          =   195
         Index           =   2
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄„Ì·"
         Height          =   195
         Index           =   1
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OpTcode 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ”«»"
         Height          =   195
         Index           =   0
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " ⁄œÌ· «·‘—Õ ·”ÿ— „⁄Ì‰"
      Height          =   1695
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   6600
      Width           =   9015
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         DataField       =   "DEV_DES"
         DataSource      =   "Adodc1"
         Height          =   1215
         Left            =   1080
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   240
         Width           =   7815
      End
      Begin ALLButtonS.ALLButton ALLButton5 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Õ›Ÿ"
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
         MICON           =   "Voucher_search.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "«” »œ«· ﬂ·„…"
      Height          =   1455
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   6600
      Width           =   3375
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "‰›–"
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
         MICON           =   "Voucher_search.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txt2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«” »œ«· »"
         Height          =   255
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "«·»ÕÀ ⁄‰"
         Height          =   255
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Õœœ «· «—ÌŒ"
      Height          =   735
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1440
      Width           =   5535
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "«·Ï"
         Height          =   195
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "„‰"
         Height          =   195
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTP_Date 
         Height          =   330
         Left            =   3000
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   206962691
         CurrentDate     =   37140
      End
      Begin MSComCtl2.DTPicker DTP_Date1 
         Height          =   330
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   206962691
         CurrentDate     =   37140
      End
   End
   Begin ALLButtonS.ALLButton ALLButton3 
      Height          =   375
      Left            =   5760
      TabIndex        =   18
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÿ»«⁄Â «·‘«‘…"
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
      MICON           =   "Voucher_search.frx":0044
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
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   2
      Left            =   9960
      TabIndex        =   17
      Tag             =   "Press F3 To Search"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   -480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   4
      Left            =   5760
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2280
      Width           =   6015
   End
   Begin VB.CheckBox Check2 
      Caption         =   " Õ ÊÏ «·ﬂ·„…"
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   -120
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "«·ﬂ·„… ›ﬁÿ"
      Height          =   375
      Left            =   10800
      TabIndex        =   5
      Top             =   -120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   1
      Left            =   5760
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2760
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   9960
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   8400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "„Ê«›ﬁ"
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
      MICON           =   "Voucher_search.frx":0060
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
      Height          =   585
      Left            =   1080
      Top             =   -120
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Voucher_search.frx":007C
      Height          =   2895
      Left            =   0
      TabIndex        =   8
      Top             =   3600
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      ColumnCount     =   27
      BeginProperty Column00 
         DataField       =   "NoteID"
         Caption         =   "NoteID"
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
         DataField       =   "NoteDate"
         Caption         =   "NoteDate"
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
         DataField       =   "NoteType"
         Caption         =   "NoteType"
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
         DataField       =   "NoteSerial"
         Caption         =   "—ﬁ„ «·ﬁÌœ"
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
         DataField       =   "Note_Value"
         Caption         =   "Note_Value"
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
         DataField       =   "Remark"
         Caption         =   "Remark"
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
         DataField       =   "Double_Entry_Vouchers_ID"
         Caption         =   "Double_Entry_Vouchers_ID"
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
         DataField       =   "DEV_ID_Line_No"
         Caption         =   "DEV_ID_Line_No"
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
         DataField       =   "Account_serial"
         Caption         =   "ﬂÊœ «·Õ”«»"
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
         DataField       =   "DEV_Value"
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
      BeginProperty Column10 
         DataField       =   "Credit_Or_Debit"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Credit"
            FalseValue      =   "Depit"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "DEV_DES"
         Caption         =   "«·‘—Õ"
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
      BeginProperty Column12 
         DataField       =   "RecordDate"
         Caption         =   "«· «—ÌŒ"
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
      BeginProperty Column13 
         DataField       =   "Notes_ID"
         Caption         =   "Notes_ID"
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
      BeginProperty Column14 
         DataField       =   "ReceiptID"
         Caption         =   "ReceiptID"
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
      BeginProperty Column15 
         DataField       =   "OperaID"
         Caption         =   "OperaID"
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
      BeginProperty Column16 
         DataField       =   "Transaction_ID"
         Caption         =   "Transaction_ID"
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
      BeginProperty Column17 
         DataField       =   "AdvanceID"
         Caption         =   "AdvanceID"
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
      BeginProperty Column18 
         DataField       =   "UserID"
         Caption         =   "UserID"
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
      BeginProperty Column19 
         DataField       =   "Posted"
         Caption         =   "Posted"
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
      BeginProperty Column20 
         DataField       =   "PostedDate"
         Caption         =   "PostedDate"
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
      BeginProperty Column21 
         DataField       =   "PostedUserID"
         Caption         =   "PostedUserID"
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
      BeginProperty Column22 
         DataField       =   "Account_Interval_ID"
         Caption         =   "Account_Interval_ID"
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
      BeginProperty Column23 
         DataField       =   "DEV_Serial"
         Caption         =   "NO"
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
      BeginProperty Column24 
         DataField       =   "credit_value"
         Caption         =   "credit_value"
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
      BeginProperty Column25 
         DataField       =   "depet_value"
         Caption         =   "depet_value"
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
      BeginProperty Column26 
         DataField       =   "des"
         Caption         =   "des"
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
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   4995.213
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   2489.953
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   -1  'True
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   5804.788
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column19 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column20 
            Object.Visible         =   0   'False
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column21 
            Object.Visible         =   0   'False
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column22 
            Object.Visible         =   0   'False
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column23 
            Object.Visible         =   0   'False
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column24 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column25 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column26 
            Object.Visible         =   0   'False
            ColumnWidth     =   4995.213
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   2520
      Top             =   9000
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin ALLButtonS.ALLButton ALLButton4 
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      Top             =   8400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÿ»«⁄Â ﬂ‘› Õ”«»"
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
      MICON           =   "Voucher_search.frx":0091
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DCNotesTypes 
      Height          =   315
      Left            =   120
      TabIndex        =   48
      Top             =   840
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Bindings        =   "Voucher_search.frx":00AD
      Height          =   315
      Left            =   5760
      TabIndex        =   49
      Top             =   840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin ALLButtonS.ALLButton ALLButton2 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   120
      TabIndex        =   50
      Top             =   2880
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "»ÕÀ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Voucher_search.frx":00C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DCExtraAccount 
      Height          =   315
      Left            =   5760
      TabIndex        =   52
      Top             =   1800
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
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
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 = »œÊ‰  ÕœÌœ"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1410
      TabIndex        =   54
      Top             =   2520
      Width           =   1020
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·Õœ «·«ﬁ’Ï ·‰ «∆Ã «·»ÕÀ"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3720
      TabIndex        =   53
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   8400
      Width           =   5535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·Õ”«»"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12360
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„’œ—"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   -1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "   ‘«‘… «·»ÕÀ ⁄‰ «·ﬁÌÊœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   705
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   13815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "‰ÿ«ﬁ «·»ÕÀ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12240
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‘—Õ «·⁄«„"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·‘—Õ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label title_lbl 
      Caption         =   "Label8"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   -840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label case_id 
      Caption         =   "0"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   -600
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Voucher_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim first_run As Boolean

Private Sub ALLButton1_Click()

    On Error Resume Next
 
    If case_id.Caption = 0 And Adodc1.Recordset.RecordCount >= 0 Then
    FrmAccEditJournal.TxtModFlg = "R"
        FrmAccEditJournal.Retrive Adodc1.Recordset.Fields!NoteSerial
        FrmAccEditJournal.show
        FrmAccEditJournal.StrOldTransID = Adodc1.Recordset.Fields!NoteSerial
        'UPDATE_RECORDS
    End If


    If case_id.Caption = 1 And Adodc1.Recordset.RecordCount >= 0 Then

        FrmAccEditJournal4.Retrive Adodc1.Recordset.Fields!NoteSerial
        FrmAccEditJournal4.show
        FrmAccEditJournal4.StrOldTransID = Adodc1.Recordset.Fields!NoteSerial
        'UPDATE_RECORDS
    End If
    
    
        If case_id.Caption = 2 And Adodc1.Recordset.RecordCount >= 0 Then

        FrmAccEditJournal3.Retrive Adodc1.Recordset.Fields!NoteSerial
        FrmAccEditJournal3.show
        FrmAccEditJournal3.StrOldTransID = Adodc1.Recordset.Fields!NoteSerial
        'UPDATE_RECORDS
    End If
    
    
    Unload Me
End Sub

Private Sub Calendar1_Click()
         
End Sub

Private Sub Command40_Click()
    On Error Resume Next
    Calendar1.Visible = True
End Sub

Private Sub DataGrid1_Click()
    On Error Resume Next
    ALLButton1_Click

End Sub

Private Sub ALLButton2_Click()
    On Error Resume Next
    buildSql
    Exit Sub
    '    If Check3.value = vbChecked And Check4.value = vbChecked Then
    '        sql = "select * from RptLedger_Sub where account_serial like'%" & Text1(2).Text & "%' and   DEV_DES like '%" & Text1(1).Text & "%'  and recorddate >= CONVERT(DATETIME, '" & DTP_Date.year & "-" & DTP_Date.Month & "-" & DTP_Date.day & " 00:00:00', 102) and recorddate <= CONVERT(DATETIME, '" & DTP_Date1.year & "-" & DTP_Date1.Month & "-" & DTP_Date1.day & " 00:00:00', 102)"
    '    ElseIf Check3.value = vbChecked And Check4.value = Unchecked Then
    '
    '        sql = "select * from RptLedger_Sub where account_serial like'%" & Text1(2).Text & "%' and   DEV_DES like '%" & Text1(1).Text & "%'  and recorddate >= CONVERT(DATETIME, '" & DTP_Date.year & "-" & DTP_Date.Month & "-" & DTP_Date.day & " 00:00:00', 102) "
    '    ElseIf Check3.value = Unchecked And Check4.value = vbChecked Then
    '        sql = "select * from RptLedger_Sub where account_serial like'%" & Text1(2).Text & "%' and   DEV_DES like '%" & Text1(1).Text & "%' and  recorddate <= CONVERT(DATETIME, '" & DTP_Date1.year & "-" & DTP_Date1.Month & "-" & DTP_Date1.day & " 00:00:00', 102)"
    '    Else
    '        sql = "select * from RptLedger_Sub where account_serial like'%" & Text1(2).Text & "%' and   DEV_DES like '%" & Text1(1).Text & "%' "
    '
    '    End If
    '
    '    'Sql = "select * from RptLedger_Sub where     recorddate = CONVERT(DATETIME, '" & DTP_Date.year & "-" & DTP_Date.Month & "-" & DTP_Date.Day & " 00:00:00', 102)"
    '
    '    Adodc1.RecordSource = sql
    '    Adodc1.Refresh
    '
    '    If my_language = "E" Then
    '        DataGrid1.Refresh
    '    Else
    '        DataGrid2.Refresh
    '    End If
    '
    '    If Adodc1.Recordset.RecordCount = 0 Then
    '        MsgBox "not fount  ·«ÌÊÃœ ‰ «∆Ã ··»ÕÀ", vbInformation
    '    End If
    '
    '    Calendar1.Visible = False
End Sub

Private Sub ALLButton4_Click()
 
    Dim StrAccountCode As String
    Dim StrAccountName As String
    Dim des As String
    
    Dim cAccountReport As New ClsAccReports
    cAccountReport.BegineDate = Me.DTP_Date.value
    cAccountReport.EndDate = Me.DTP_Date1.value

    If Text1(2).Text <> "" Then
        StrAccountCode = Get_Account_code(Text1(2).Text)
        StrAccountName = Get_Account_name(Text1(2).Text)
        des = Text1(1).Text
    Else

        If Adodc1.Recordset.RecordCount > 0 Then
            StrAccountCode = Get_Account_code(Adodc1.Recordset.Fields!account_serial)
            StrAccountName = Get_Account_name(Adodc1.Recordset.Fields!account_serial)
            des = Text1(1).Text
                
        End If
       
    End If
        
    cAccountReport.ShowLedger1 StrAccountCode, StrAccountName, des
        
    Set cAccountReport = Nothing
End Sub

Private Sub ALLButton5_Click()
    On Error Resume Next
    Adodc1.Recordset.update
    DataGrid2.Refresh
End Sub

Private Sub Check3_Click()
If Me.Check3.value = vbChecked Then
    DTP_Date.Enabled = True
Else
    DTP_Date.Enabled = False
End If

End Sub

Private Sub check4_Click()
If Me.Check4.value = vbChecked Then
    DTP_Date1.Enabled = True
Else
    DTP_Date1.Enabled = False
End If
End Sub

Private Sub check5_Click()
If Me.Check5.value = vbChecked Then
    DCNotesTypes.Enabled = True
Else
    DCNotesTypes.Enabled = False
End If

End Sub

Private Sub check6_Click()
If Me.Check6.value = vbChecked Then
    DcBranch.Enabled = True
Else
    DcBranch.Enabled = False
End If

End Sub

Private Sub Combo1_Click()
    Dim TopRows As String
    TopRows = ""
    If val(txtRowLimt) > 0 Then
        TopRows = " Top " & val(txtRowLimt) & " "
    End If
    If Combo1.ListIndex = 0 Then
        sql = "SELECT  " & TopRows & "   * from dbo.RptLedger_Sub  where NoteType=200 "
        Adodc1.RecordSource = sql
        Adodc1.Refresh
    ElseIf Combo1.ListIndex = 1 Then
        sql = "SELECT  " & TopRows & "    * from dbo.RptLedger_Sub  where NoteType<>200 "
        Adodc1.RecordSource = sql
        Adodc1.Refresh
    End If

End Sub

Private Sub Command1_Click()

    If replace_in_data_base("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_Description", txt1.Text, txt2.Text) = True Then
        MsgBox "Done"
      
        ALLButton2_Click
    End If

End Sub

Private Sub DataGrid2_Click()
ALLButton1_Click
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
      ALLButton2_Click
            
    End If
End Sub

Private Sub DCExtraAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.case_id = 110815
        Account_search.show vbModal
        
            
    End If
    
      If KeyCode = vbKeyReturn Then
      ALLButton2_Click
            
    End If
    
End Sub

Private Sub DCNotesTypes_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
      ALLButton2_Click
            
    End If
End Sub

Private Sub Form_Activate()

    If first_run = True Then

        Exit Sub
    Else
        first_run = True
        Dim TopRows As String
        TopRows = ""
        If val(txtRowLimt) > 0 Then
            TopRows = " Top " & val(txtRowLimt) & " "
        End If
        sql = "SELECT  " & TopRows & "   * from dbo.RptLedger_Sub   where Double_Entry_Vouchers_ID=0"
        Adodc1.RecordSource = sql
        Adodc1.Refresh
        DataGrid2.Refresh
 
    End If

End Sub
Function buildSql()

    sql = "select * from RptLedger_Sub where 1=1 "

    Dim RptLedger_Sub As String
    Dim TopRows       As String
    TopRows = ""
    If val(txtRowLimt) > 0 Then
        TopRows = " Top " & val(txtRowLimt) & " "
    End If
    RptLedger_Sub = "SELECT  " & TopRows & "    dbo.Notes.ChqueNum, dbo.Notes.ManualNo, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, "
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS DEV_Value, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS DEV_DES,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione AS DevDESE, dbo.ACCOUNTS.Account_Name,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.TblNotesTypes.NotesTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.TblUsers.UserName,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.ReceiptID,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.OperaID, dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID, dbo.Transactions.Transaction_Serial,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.Transactions.Transaction_Date, dbo.TransactionTypes.TransactionTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.PostedDate,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.PostedUserID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.Notes.NoteDate, dbo.Notes.NoteType,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.Notes.NoteSerial, dbo.Notes.Note_Value, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Parent_Account_Code,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.ACCOUNTS.Branch, dbo.ACCOUNTS.Sum_account, dbo.ACCOUNTS.cost_center,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.ACCOUNTS.currenct_code, dbo.Notes.Remark, dbo.Notes.note_value_by_characters, dbo.Notes.foxy_no, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.TblNotesTypes.NotesTypeNamee, dbo.TransactionTypes.TransactionEnglishName, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.TblBranchesData.ActivityTypeId, dbo.DOUBLE_ENTREY_VOUCHERS.notes_all, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.Posted, dbo.DOUBLE_ENTREY_VOUCHERS.valuee AS DEV_ValueE, dbo.DOUBLE_ENTREY_VOUCHERS.currency,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.rate, dbo.TblBranchesData.RegionID, dbo.TblSection.name, dbo.TblSection.namee,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.DescAccount, dbo.DOUBLE_ENTREY_VOUCHERS.NextAccount_Code, dbo.DOUBLE_ENTREY_VOUCHERS.project_id,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.DOUBLE_ENTREY_VOUCHERS.operid,"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS.pandid , dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid, dbo.TblAqar.aqarname, dbo.TblAqar.aqarNo"
    RptLedger_Sub = RptLedger_Sub & " FROM         dbo.TblAqar RIGHT OUTER JOIN"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.TblBranchesData INNER JOIN"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.TblUsers INNER JOIN"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.TblUsers.UserID = dbo.DOUBLE_ENTREY_VOUCHERS.UserID ON"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.TblBranchesData.branch_id = dbo.DOUBLE_ENTREY_VOUCHERS.branch_id ON"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.TblAqar.Aqarid = dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid LEFT OUTER JOIN"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.TblSection ON dbo.TblBranchesData.RegionID = dbo.TblSection.Id LEFT OUTER JOIN"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.Notes LEFT OUTER JOIN"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.Transactions ON dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
    RptLedger_Sub = RptLedger_Sub & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    RptLedger_Sub = RptLedger_Sub & "     where 1=1 "

    sql = RptLedger_Sub
    If (Text1(0).Text) <> "" Then
        sql = sql & " and   convert(decimal(30,10), Notes.noteSerial)   LIKE '%" & (Text1(0).Text) & "%'"

    End If

    If (Text1(1).Text) <> "" Then
        sql = sql & " and    Double_Entry_Vouchers_Description  like '%" & Text1(1).Text & "%'"

    End If

    If (Text1(2).Text) <> "" Then
        sql = sql & " and    account_serial like '%" & Text1(2).Text & "%'"

    End If
 
    If val((Text1(3).Text)) <> 0 Then
        sql = sql & " AND value >=" & val(Text1(3).Text)

    End If
 
    If val((Text1(5).Text)) <> 0 Then
        sql = sql & " AND value <=" & val(Text1(5).Text)

    End If

    If (Text1(4).Text) <> "" Then
        sql = sql & " and    Notes.remark like '%" & Text1(4).Text & "%'"

    End If

    If Check3.value = vbChecked Then
    
        sql = sql + " and RecordDate >=" & SQLDate(DTP_Date, True) & ""
    
    End If

    If Check4.value = vbChecked Then
    
        sql = sql & " and RecordDate <=" & SQLDate(DTP_Date1, True) & ""
    
    End If

    If Check5.value = vbChecked Then
    
        sql = sql & " and NoteType =" & val(DCNotesTypes.BoundText) & ""
    
    End If
 
    If Check6.value = vbChecked Then
    
        sql = sql & " and Notes.branch_no =" & val(DcBranch.BoundText) & ""
    
    End If

    If SystemOptions.usertype <> UserAdminAll Then
        sql = sql & " and Notes.branch_no =" & Current_branch & ""
    End If
    
    Adodc1.RecordSource = sql
    Adodc1.Refresh

    If my_language = "E" Then
        DataGrid1.Refresh
    Else
        DataGrid2.Refresh
    End If

    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "not fount  ·«ÌÊÃœ ‰ «∆Ã ··»ÕÀ", vbInformation
    End If
                
    If SystemOptions.UserInterface = ArabicInterface Then
        Label12.Caption = "‰ ÌÃ… «·»ÕÀ  " & Adodc1.Recordset.RecordCount & " " & IIf(val(txtRowLimt) > 0, "[" & txtRowLimt & " „  ÕœÌœ «·‰ «∆Ã » [", "")
    Else
        Label12.Caption = "Serarch Result  " & Adodc1.Recordset.RecordCount
    End If

End Function
Private Sub ChangeLang()
    Me.Caption = "Voucher Search"
    ALLButton5.Caption = "Save"
    Label3.Caption = "ACC Code"
    OpTcode(0).Caption = "Acc"
        OpTcode(1).Caption = "Customer"
            OpTcode(2).Caption = "Supp."
                OpTcode(3).Caption = "Employee"
              
                    
    Frame3.Caption = "Update Description Per Line "
    LblHeader.Caption = Me.Caption
    Label4.Caption = "Search"
    Label6.Caption = "Voucher#"
    Label7.Caption = "General Des"
    Label1.Caption = "Des"
    Label2.Caption = "Value"
    Label5.Caption = "Source"
    Combo1.Clear
    Combo1.AddItem "Manual"
    Combo1.AddItem "Auto"
    Label9.Caption = "Account#"
    Frame1.Caption = "Date"
    Check3.Caption = "From"
    Check4.Caption = "To"
    Frame5.Caption = "Value"
    ALLButton2.Caption = "Search"
    ALLButton4.Caption = "Print Statement of Acc."
    Frame2.Caption = "Replace Text "
    Label8.Caption = "Find"
    Label10.Caption = "Replace By"
    Command1.Caption = "Replace All"
    DataGrid2.Columns(3).Caption = "Voucher#"
    DataGrid2.Columns(8).Caption = "Account#"
    DataGrid2.Columns(9).Caption = "Value"
    DataGrid2.Columns(11).Caption = "Des"
    DataGrid2.Columns(12).Caption = "date"
    DataGrid2.RightToLeft = False
    Label2.Caption = "From"
    Label11.Caption = "To"
    Check5.Caption = "Transacion"
    Check6.Caption = "Branch"
    
    ALLButton1.Caption = "Ok"

End Sub

Private Sub Form_Load()
    On Error Resume Next
   
    DTP_Date.value = Date
    DTP_Date1.value = Date
         
    Combo1.Clear
    Combo1.AddItem "ÌœÊÌ"
    Combo1.AddItem "«·Ì"
         
    '
    Dim StrSQL  As String
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetAccountingCodes DCExtraAccount, True

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT NotesType,NotesTypeName From TblNotesTypes order by NotesTypeName "
    Else
        StrSQL = "SELECT NotesType,NotesTypeNamee From TblNotesTypes  order by NotesTypeNamee"
    End If

    fill_combo DCNotesTypes, StrSQL

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  select branch_id,branch_name from TblBranchesData   "
    Else
        StrSQL = "  select branch_id,branch_namee from TblBranchesData   "
    End If
    fill_combo DcBranch, StrSQL

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
 
    If my_language = "E" Then
        DataGrid1.Visible = True
        DataGrid2.Visible = False
    End If
   
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    first_run = False
End Sub

Private Sub Text1_Change(Index As Integer)
    'If KeyCode = vbKeyReturn Then
     DCExtraAccount.BoundText = ""
       DCExtraAccount.BoundText = Get_Account_code(Text1(2).Text, 1)

    'End If
    
End Sub

Private Sub Text1_KeyUp(Index As Integer, _
                        KeyCode As Integer, _
                        Shift As Integer)

    'On Error Resume Next
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 110815

    End If


  If KeyCode = vbKeyReturn Then
      ALLButton2_Click
            
    End If
    
    

Exit Sub
    If KeyCode = 13 Then

        '     On Error Resume Next
             
        If Index = 0 Then
 
            If Check1.value = 1 Then
                        
                sql = "select * from RptLedger_Sub where    noteSerial LIKE '%" & val(Text1(Index).Text) & "%'"
            Else
                sql = "select * from RptLedger_Sub where    noteSerial LIKE '%" & val(Text1(Index).Text) & "%'"

                'Sql = "select * from RptLedger_Sub where NoteType=200  and noteSerial  like '%" & Text1(Index).text & "%'"
            End If
        End If
                   
        If Index = 1 Then
            sql = "select * from RptLedger_Sub where    DEV_DES like '%" & Text1(Index).Text & "%'"
            '    If Check1.value = 1 Then
            '      Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_type = '" & Text1(Index).text & "' "
            '     Else
            '      Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_type like '%" & Text1(Index).text & "%'"
            '
    
            '                  End If
        End If
                    
        If Index = 2 Then
            sql = "select * from RptLedger_Sub where account_serial like '%" & Text1(2).Text & "%' and   DEV_DES like '%" & Text1(1).Text & "%'"
              
        End If
                    
        If Index = 3 Then
            If Not IsNumeric(Text1(Index).Text) Then Exit Sub
            sql = "select * from RptLedger_Sub where dev_value =" & val(Text1(Index).Text)
        End If
                    
        If Index = 4 Then
            'If Check1.value = 1 Then
            sql = "select * from RptLedger_Sub where    remark like '%" & Text1(Index).Text & "%'"
            'Else
            ' Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and description like '%" & Text1(Index).text & "%'"
    
            'End If
        End If
                  
        If Index = 5 Then
            If Check1.value = 1 Then
                '            Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_source = '" & Text1(Index).text & "' "
            Else
                '            Sql = "select * from sand_all_details_qry where  type='" & title_lbl.Caption & "' and sanad_source like '%" & Text1(Index).text & "%'"
    
            End If
        End If
             
        Adodc1.RecordSource = sql
        Adodc1.Refresh

                If my_language = "E" Then
                    DataGrid1.Refresh
                Else
                    DataGrid2.Refresh
                End If

                If Adodc1.Recordset.RecordCount = 0 Then
                    MsgBox "not fount  ·«ÌÊÃœ ‰ «∆Ã ··»ÕÀ", vbInformation
                End If
 
    End If

End Sub

