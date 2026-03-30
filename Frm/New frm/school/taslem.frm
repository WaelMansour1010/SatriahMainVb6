VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form taslem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ…  ”·Ì„ «·ﬂ » «·œ—«”Ì…"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   11190
   Begin VB.CommandButton Command5 
      Caption         =   " ”·Ì„   «·ﬂ·"
      Height          =   615
      Left            =   600
      TabIndex        =   16
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   " ÕœÌœ"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ÃœÌœ"
      Height          =   615
      Left            =   600
      TabIndex        =   14
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   " ”·Ì„ «·„Õœœ ›ﬁÿ"
      Height          =   615
      Left            =   600
      TabIndex        =   13
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Command13 
      Height          =   735
      Left            =   5160
      Picture         =   "taslem.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "taslem.frx":1992
      Left            =   6120
      List            =   "taslem.frx":199C
      Style           =   1  'Simple Combo
      TabIndex        =   11
      Top             =   840
      Width           =   3855
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
      ItemData        =   "taslem.frx":19AF
      Left            =   120
      List            =   "taslem.frx":19B9
      Style           =   1  'Simple Combo
      TabIndex        =   9
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   8880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "ÃœÌœ"
      Height          =   375
      Left            =   120
      TabIndex        =   1
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
      ItemData        =   "taslem.frx":19CC
      Left            =   3720
      List            =   "taslem.frx":19D6
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   1560
      Width           =   5775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "taslem.frx":19E9
      Height          =   2415
      Left            =   3240
      TabIndex        =   3
      Top             =   2280
      Width           =   7575
      _ExtentX        =   13361
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
         Caption         =   " „ «· ”·Ì„"
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
            Object.Visible         =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "«·ÿ«·»"
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
      Left            =   9480
      TabIndex        =   10
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label70 
      Caption         =   "0"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label60 
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Õ—ﬂ…  ”·Ì„ «·ﬂ » «·œ—«”Ì…"
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
      Left            =   3840
      TabIndex        =   6
      Top             =   0
      Width           =   4455
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
      Left            =   3360
      TabIndex        =   5
      Top             =   840
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
      Left            =   9480
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "taslem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command13_Click()
    member_search.Show
    member_search.from = 40
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
