VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form imaged 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÚÑÖ ãÑÝÞÇÊ ã"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   11910
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
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   5520
      TabIndex        =   33
      Top             =   1680
      Width           =   4935
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   34
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ÊßÈíÑ"
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
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   5
         Left            =   1200
         TabIndex        =   35
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ÊÕÛíÑ"
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
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ÏæÑÇä"
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
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   37
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ÊÛííÑ ÕæÑÉ"
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
         MICON           =   "imaged.frx":0054
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
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   2760
      TabIndex        =   28
      Top             =   8760
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   2520
      TabIndex        =   23
      Top             =   8760
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label xx 
         Caption         =   "ÇáãæÙÝ ÇáÍÇáí"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label yy 
         Caption         =   "ÇáÞÓã"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      Height          =   2535
      Left            =   600
      TabIndex        =   19
      Top             =   0
      Width           =   1455
      Begin ALLButtonS.ALLButton Command10 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÌÏíÏ"
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
         BCOLO           =   16777215
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":0070
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
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16777215
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command3 
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         MICON           =   "imaged.frx":00A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ØÈÇÚå"
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
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "imaged.frx":00C4
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
   Begin VB.TextBox txtopeation_type 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   -720
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0C0C0&
      DataField       =   "operation_type"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6000
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ÊÍãíá ÕæÑÉ"
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
      MICON           =   "imaged.frx":00E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      DataField       =   "image_no"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9000
      TabIndex        =   10
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      DataField       =   "subject_no"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Text            =   "Text6"
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      DataField       =   "image_NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5640
      TabIndex        =   5
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      DataField       =   "subject_no"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8880
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      DataField       =   "image_date"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin DBPIXLib.DBPix20 DBPix201 
      Height          =   5415
      Left            =   720
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   10575
      _Version        =   131072
      _ExtentX        =   18653
      _ExtentY        =   9551
      _StockProps     =   1
      _Image          =   "imaged.frx":00FC
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
      Height          =   375
      Left            =   3120
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
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
      Caption         =   " ÊÍÑíß"
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
      Height          =   375
      Left            =   -720
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   2040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "ãÓÍ ÖæÆí"
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
      MICON           =   "imaged.frx":0114
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   22
      ToolTipText     =   "Language  ÇááÛÉ"
      Top             =   0
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
      MICON           =   "imaged.frx":0130
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "äæÚ ÇáãÑÝÞ"
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
      Left            =   6240
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label screen_name 
      Alignment       =   1  'Right Justify
      Caption         =   "Label4"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ÇÓã ÇáãÑÝÞ"
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
      Left            =   5880
      TabIndex        =   11
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label DEPARTEMENT 
      Caption         =   "Label5"
      Height          =   255
      Left            =   10080
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÇÑíÎ ÇáãÑÝÞ"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label SUBJECT_NO 
      Caption         =   "Label3"
      Height          =   255
      Left            =   12240
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ÑÞã ÇáãÑÝÞ"
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
      Left            =   9120
      TabIndex        =   4
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ßæÏ ÇáãÚÏÉ"
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
      Left            =   9360
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   6015
      Left            =   360
      Top             =   2640
      Width           =   11055
   End
End
Attribute VB_Name = "imaged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NEW_IMAGE As Boolean

Private Sub ALLButton1_Click()
On Error Resume Next
DBPix201.ImageLoad
DBPix201.ImageSaveFile (system_path & "\images\" & Text2.Text & ".JPG")
NEW_IMAGE = False
End Sub

Private Sub ALLButton2_Click()
On Error Resume Next
DBPix201.TWAINAcquire
DBPix201.ImageSaveFile (system_path & "\images\" & Text2.Text & ".JPG")
NEW_IMAGE = False
End Sub

Private Sub CMD_language_Click()
On Error Resume Next

If CMD_language.Caption = "EN" Then
my_language = "E"
 
'Call Reload(Me)

 
Else
my_language = "A"
 
'Call Reload(Me)
End If
End Sub

Private Sub Command1_Click(Index As Integer)


If Index = 3 Then
'Dim x As Integer
X = MsgBox("åá ÊÑíÏ ÕæÑÉ ãä ãáÝ", vbExclamation + vbYesNoCancel)
If X = vbYes Then
DBPix201.ImageLoad

Else
If X = vbNo Then
DBPix201.TWAINAcquire
Else

Exit Sub
End If
End If
DBPix201.ImageSaveFile (App.Path & "\images\" & Text2.Text & ".JPG")
End If


If Index = 4 Then
DBPix201.ViewZoomIn
End If

If Index = 5 Then
DBPix201.ViewZoomOut
End If

If Index = 6 Then
DBPix201.ImageRotate ImageRotate90
End If


If Index = 7 Then
On Error Resume Next
loading_temolates.Show
loading_temolates.Frame2.Visible = False
loading_temolates.Frame3.Visible = False
loading_temolates.Frame4.Visible = False
loading_temolates.Frame5.Visible = False
 loading_temolates.Frame6.Top = 4800
loading_temolates.Image1.Picture = LoadPicture(App.Path & "\images\" & Text2.Text & ".JPG")
loading_temolates.Image1.Enabled = False
End If

'&&&&&&&&&&&&&&&&&&&&

End Sub

Private Sub Command10_Click(Index As Integer)
On Error Resume Next
NEW_IMAGE = True
DBPix201.ImageClear
DBPix201.Visible = True
Dim LASTIMAGENO As Integer
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "SELECT MAX(image_no)  AS LASTIMAGENO FROM subjects_images WHERE subject_no='" & SUBJECT_NO.Caption & "'"
Adodc2.Refresh

If Adodc2.Recordset.RecordCount = 0 Or IsNull(Adodc2.Recordset.Fields!LASTIMAGENO) Then
LASTIMAGENO = 1
Else
LASTIMAGENO = (Adodc2.Recordset.Fields!LASTIMAGENO) + 1
End If

Adodc1.Recordset.AddNew
Text1.Text = SUBJECT_NO.Caption
Text2.Text = SUBJECT_NO & "-" & LASTIMAGENO & "#" & Mid(Date, 1, 2) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 7, 4)
Text3.Text = Now
Text4.Text = LASTIMAGENO
Text5.Text = txtopeation_type.Text
'Text5.Text = DEPARTEMENT.Caption
Adodc1.Recordset.Update


End Sub

Private Sub Command2_Click()
On Error Resume Next
system_path = "D:\my works\school"

DBPix201.ImageSaveFile (system_path & "\images\" & Text2.Text & ".JPG")
NEW_IMAGE = False
End Sub

Private Sub Command3_Click()
On Error Resume Next

           If my_language = "E" Then
              X = MsgBox("Confirm delete", vbCritical + vbYesNo)
            Else
            X = MsgBox("ÊÃßíÏ ÇáÍÐÝ", vbCritical + vbYesNo)
              
            End If

If X = vbNo Then
Exit Sub
End If
If Adodc1.Recordset.RecordCount > 0 Then
Adodc1.Recordset.Delete
Adodc1.Refresh
End If
End Sub

 

Private Sub Form_Activate()
On Error Resume Next

 

End Sub

Private Sub Form_Load()
On Error Resume Next

 
   ' login.SkinFramework.ApplyWindow Me.hWnd
    
If my_language = "E" Then
CMD_language.ToolTipText = "Change Language"

Me.dept_lbl = departement_name
Me.emp_name_lbl = current_user_name
   InfoE.Visible = True
   infoA.Visible = False
Else

emp_a.Caption = current_user_name
dep_a.Caption = departement_name

   
   infoA.Visible = True
   InfoE.Visible = False
End If

 Me.Left = (MDIForm1.Width - Me.Width) / 2
    Me.Top = (MDIForm1.Height - Me.Height) / 2 - 500
    
'LoadSettings
 If my_language = "E" Then
 CMD_language.Caption = "ÚÑÈí"
 Label6.Caption = "Car Code"
 Label1.Caption = "Attachment no"
  Label2.Caption = "Attachment Name"
   Label3.Caption = "Attachment date"
    Label4.Caption = "Attachment type"
    
   Adodc1.Caption = "move"
  
 
  Me.Caption = "View Attachments"
  
  Command10(0).Caption = "new"
  Command2.Caption = "save"
  Command3.Caption = "delete"
  ALLButton1.Caption = "load image from files"
  ALLButton2.Caption = "Acquire image from Scanner"
  
 End If
 
connection_string = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"

Adodc1.ConnectionString = connection_string
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE subject_no=''  "
Adodc1.Refresh

Adodc2.ConnectionString = connection_string
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "select * from subjects_images   "
Adodc2.Refresh


NEW_IMAGE = False
'Command4_Click
End Sub

Private Sub Text2_Change()
On Error Resume Next
system_path = "D:\my works\school"
If Text2.Text = "" Or NEW_IMAGE = True Then Exit Sub
If DBPix201.ImageLoadFile(system_path & "\images\" & Text2.Text & ".JPG") = True Then

End If
End Sub

