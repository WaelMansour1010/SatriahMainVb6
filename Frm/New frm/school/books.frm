VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form books 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·ﬂ » «·œ—«”Ì…"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   12450
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5160
      TabIndex        =   13
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "books.frx":0000
      Left            =   5760
      List            =   "books.frx":000A
      TabIndex        =   12
      Top             =   840
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   8880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "ÃœÌœ"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   8520
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "books.frx":0039
      Left            =   3840
      List            =   "books.frx":0043
      RightToLeft     =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Top             =   1560
      Width           =   5775
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "books.frx":0056
      Left            =   7800
      List            =   "books.frx":0060
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "«÷«›…"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   2280
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "books.frx":0071
      Height          =   2415
      Left            =   4080
      TabIndex        =   5
      Top             =   3000
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "FINES_NO"
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
      BeginProperty Column01 
         DataField       =   "FINES_VALUE"
         Caption         =   "«”„ «·ﬂ «»"
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
         DataField       =   "PAYED_DATE"
         Caption         =   "ÿ—ﬁ… «· ”·Ì„"
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
         DataField       =   "ACTIVATED"
         Caption         =   "ACTIVATED"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "«·ﬁÌ„…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   14
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label70 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label60 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "«·ﬂ » «·œ—«”Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "«·’› «·œ—«”Ì"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·ﬂ «»"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "ÿ—Ìﬁ… «· ”·Ì„"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "books"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
