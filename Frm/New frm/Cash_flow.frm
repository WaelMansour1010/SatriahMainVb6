VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Cash_flow 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13260
   Icon            =   "Cash_flow.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   13260
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   8055
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   -120
      Width           =   1215
      Begin VB.Label lbl_Current_idex 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-1"
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   13
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   7560
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   12
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   7200
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   11
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   6600
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   10
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   6240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   9
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   5880
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   8
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   7
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   6
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   5
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   4
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   3
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7935
      Left            =   0
      TabIndex        =   15
      Top             =   -120
      Width           =   4815
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   480
         Top             =   4200
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSDataListLib.DataCombo dcAccount 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ALLButtonS.ALLButton ALLButton10 
         Height          =   375
         Left            =   0
         TabIndex        =   33
         Top             =   7440
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "‘—Õ ﬁ«∆„… «· œ›ﬁ «·‰ﬁœÌ"
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cash_flow.frx":6852
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton2 
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "«œ—«Ã"
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Cash_flow.frx":686E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Cash_flow.frx":688A
         Height          =   2775
         Left            =   0
         TabIndex        =   36
         Top             =   960
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   " "
         ColumnCount     =   5
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "Account_Code"
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
         BeginProperty Column02 
            DataField       =   "Account_Name"
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
         BeginProperty Column03 
            DataField       =   "Balance"
            Caption         =   "«·—’Ìœ"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   14
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   7080
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«Œ — «·Õ”«»"
         Height          =   255
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   255
      Index           =   0
      Left            =   11880
      TabIndex        =   0
      Top             =   1320
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":689F
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
      Height          =   255
      Index           =   1
      Left            =   11880
      TabIndex        =   1
      Top             =   1920
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":68BB
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
      Height          =   255
      Index           =   2
      Left            =   11880
      TabIndex        =   2
      Top             =   2280
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":68D7
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
      Height          =   255
      Index           =   3
      Left            =   11880
      TabIndex        =   3
      Top             =   3000
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":68F3
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
      Height          =   255
      Index           =   4
      Left            =   11880
      TabIndex        =   4
      Top             =   3360
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":690F
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
      Height          =   255
      Index           =   5
      Left            =   11880
      TabIndex        =   5
      Top             =   4440
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":692B
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
      Height          =   255
      Index           =   6
      Left            =   11880
      TabIndex        =   6
      Top             =   5040
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":6947
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
      Height          =   255
      Index           =   7
      Left            =   11880
      TabIndex        =   7
      Top             =   5400
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":6963
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
      Height          =   255
      Index           =   8
      Left            =   11880
      TabIndex        =   8
      Top             =   5760
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":697F
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
      Height          =   255
      Index           =   9
      Left            =   11880
      TabIndex        =   9
      Top             =   6120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":699B
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
      Height          =   255
      Index           =   10
      Left            =   11880
      TabIndex        =   10
      Top             =   7080
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":69B7
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
      Height          =   255
      Index           =   11
      Left            =   11880
      TabIndex        =   34
      Top             =   7440
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "«Œ Ì«— «·Õ”«»« "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Cash_flow.frx":69D3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "«Ê·«"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Index           =   0
      Left            =   11880
      TabIndex        =   11
      Top             =   930
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Index           =   3
      Left            =   11880
      TabIndex        =   14
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "À«·À«"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Index           =   2
      Left            =   12000
      TabIndex        =   13
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "À«‰Ì«"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Index           =   1
      Left            =   12000
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   7800
      Left            =   5880
      Picture         =   "Cash_flow.frx":69EF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7305
   End
End
Attribute VB_Name = "Cash_flow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ALLButton1_Click(Index As Integer)
    lbl_Current_idex.Caption = Index
    Adodc1.RecordSource = "select * from  Cash_flow where [index]=" & Index
    Adodc1.Refresh
    Label3(lbl_Current_idex.Caption).Caption = get_sum(lbl_Current_idex.Caption)

End Sub
Private Function get_sum(Index As Integer)
     Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "select sum(Balance) As total from Cash_flow where [index]=" & Index
     Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
      If Not IsNull(Rs3("total").value) Then
        get_sum = Rs3("total").value
    Else
        get_sum = 0
    End If
    Rs3.Close
End Function
Private Sub ALLButton10_Click()
    Dim xApp As New CRAXDRT.Application
    Dim EmpReport As ClsEmployeeReport
    Dim xReport As New CRAXDRT.Report
    Set xReport = xApp.OpenReport(App.path & "\reports\New Reports\REPORT1.rpt")
    ' xReport.Database.SetDataSource Rs
     Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    '   xReport.ReportTitle = X
    SendKeys "{RIGHT}"
End Sub
Private Sub ALLButton2_Click()
    If dcAccount.BoundText <> "" And lbl_Current_idex.Caption <> "-1" Then
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields!Account_Code = dcAccount.BoundText
        Adodc1.Recordset.Fields!account_name = dcAccount.text
        Adodc1.Recordset.Fields!Balance = get_balance_P(dcAccount.BoundText)
        Adodc1.Recordset.Fields![Index] = lbl_Current_idex.Caption

        Adodc1.Recordset.update
        DataGrid2.Refresh
        Label3(lbl_Current_idex.Caption).Caption = get_sum(lbl_Current_idex.Caption)
        Label3(5).Caption = (val(Label3(0).Caption) + val(Label3(1).Caption) + val(Label3(2).Caption)) - val((Label3(3).Caption) + val(Label3(4).Caption))

        Label3(11).Caption = (val(Label3(6).Caption) + val(Label3(7).Caption) + val(Label3(8).Caption)) + (val(Label3(9).Caption) + val(Label3(10).Caption))

    End If

End Sub
Private Sub DataGrid2_KeyUp(KeyCode As Integer, _
                            Shift As Integer)

    If KeyCode = 46 Then
        If Adodc1.Recordset.RecordCount > 0 Then
            Adodc1.Recordset.delete
            DataGrid2.Refresh
            Label3(lbl_Current_idex.Caption).Caption = get_sum(lbl_Current_idex.Caption)
            Label3(5).Caption = (val(Label3(0).Caption) + val(Label3(1).Caption) + val(Label3(2).Caption)) - val((Label3(3).Caption) + val(Label3(4).Caption))

            Label3(11).Caption = (val(Label3(6).Caption) + val(Label3(7).Caption) + val(Label3(8).Caption)) + (val(Label3(9).Caption) + val(Label3(10).Caption))
        End If
    End If
End Sub
Private Sub Form_Load()
    Dim My_SQL As String
    Dim i As Integer
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    connection_string = Cn.ConnectionString
    My_SQL = "select Account_Serial,Account_Name from ACCOUNTS  where last_account=1"
    fill_combo dcAccount, My_SQL
    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from  Cash_flow where [index]=-1"
    Adodc1.Refresh
     For i = 0 To Label3.count - 1
        Label3(i).Caption = get_sum(i)
    Next i
    Label3(5).Caption = (val(Label3(0).Caption) + val(Label3(1).Caption) + val(Label3(2).Caption)) - val((Label3(3).Caption) + val(Label3(4).Caption))
    Label3(11).Caption = (val(Label3(6).Caption) + val(Label3(7).Caption) + val(Label3(8).Caption)) + (val(Label3(9).Caption) + val(Label3(10).Caption))
       If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
End Sub
Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Cash Flow Ctatement"
    ' labell name
    Me.ALLButton1(0).Caption = "Choose Accounts"
    Me.ALLButton1(1).Caption = "Choose Accounts"
    Me.ALLButton1(2).Caption = "Choose Accounts"
    Me.ALLButton1(3).Caption = "Choose Accounts"
    Me.ALLButton1(4).Caption = "Choose Accounts"
    Me.ALLButton1(5).Caption = "Choose Accounts"
    Me.ALLButton1(6).Caption = "Choose Accounts"
    Me.ALLButton1(7).Caption = "Choose Accounts"
    Me.ALLButton1(8).Caption = "Choose Accounts"
    Me.ALLButton1(9).Caption = "Choose Accounts"
    Me.ALLButton1(10).Caption = "Choose Accounts"
    Me.ALLButton1(11).Caption = "Choose Accounts"
    Me.ALLButton2.Caption = "insert"
    Me.ALLButton10.Caption = "Explain Cash Flow Ctatement"
    
    Me.Label1(0).Caption = "First"
    Me.Label1(1).Caption = "Second"
    Me.Label1(2).Caption = "Third"
    Me.Label2.Caption = "Choose Account"
        

    With Me.DataGrid2
    .Columns(0).Caption = "id "
    .Columns(1).Caption = "Acounting Code"
    .Columns(2).Caption = "Acounting Name"
    .Columns(3).Caption = "Balance"
    .Columns(4).Caption = "index"
    End With
ErrTrap:
End Sub
