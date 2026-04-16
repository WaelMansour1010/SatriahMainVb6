VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form exams_details 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· »Ì«‰«  «·«„ Õ«‰« "
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   10125
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CommandButton Command1 
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "exams_details.frx":0000
      DataField       =   "MEMBER_TYPE"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   5880
      TabIndex        =   1
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "MEMBER_NAME"
      BoundColumn     =   "MEMBER_NAME"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "exams_details.frx":0015
      DataField       =   "MEMBER_TYPE"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   5880
      TabIndex        =   6
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "MEMBER_NAME"
      BoundColumn     =   "MEMBER_NAME"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "exams_details.frx":002A
      Height          =   2655
      Left            =   0
      TabIndex        =   14
      Top             =   2880
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   24
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "item_code"
         Caption         =   "«”„ «·ÿ«·»"
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
         DataField       =   "part_no"
         Caption         =   "œ—Ã… «⁄„«· «·”‰…/«·‘Â—"
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
         DataField       =   "items_name"
         Caption         =   "«·«„ Õ«‰ «· Õ—Ì—Ì"
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
         DataField       =   "risk"
         Caption         =   "œ—Ã… «·‰‘«ÿ"
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
         DataField       =   "sum1"
         Caption         =   "«·„Ã„Ê⁄"
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
         DataField       =   "akher_s3r_shera"
         Caption         =   "akher_s3r_shera"
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
         DataField       =   "departement"
         Caption         =   "departement"
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
         DataField       =   "branch_no"
         Caption         =   "branch_no"
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
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1635.024
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "exams_details.frx":003F
      DataField       =   "MEMBER_TYPE"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   5880
      TabIndex        =   16
      Top             =   1440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "MEMBER_NAME"
      BoundColumn     =   "MEMBER_NAME"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "„Ã„Ê⁄Â —ﬁ„"
      Height          =   375
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "œ—Ã«  «·ÿ·«»"
      Height          =   375
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«·œ—Ã… «·⁄Ÿ„Ï"
      Height          =   375
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "«·œ—Ã… «·’€—Ï"
      Height          =   375
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«· Œ’’"
      Height          =   375
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·’› «·œ—«”Ì"
      Height          =   375
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   " ”ÃÌ· »Ì«‰«  «·«„ Õ«‰« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„œ—”"
      Height          =   375
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·«„ Õ«‰"
      Height          =   375
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "exams_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
