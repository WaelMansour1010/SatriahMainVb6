VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form iso_departement 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "priviligies"
      Height          =   1215
      Left            =   0
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   7095
      Begin MSAdodcLib.Adodc user_priviliges_adodc 
         Height          =   495
         Left            =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.Label adodc4error 
         Alignment       =   1  'Right Justify
         Caption         =   "Label4"
         DataField       =   "employee_id"
         DataSource      =   "user_priviliges_adodc"
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label screen_name 
         Alignment       =   1  'Right Justify
         Caption         =   "M14"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ÃœÌœ"
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
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "iso_departement.frx":0000
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
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "iso_departement.frx":001C
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
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Õ–›"
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
         MICON           =   "iso_departement.frx":0038
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "id"
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
      Height          =   615
      Left            =   -2400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "iso_departement.frx":0054
      Height          =   3855
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   19
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
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
         DataField       =   "iso_departement_name"
         Caption         =   "«· ’‰Ì›« "
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
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   0
      Top             =   6120
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
   Begin VB.TextBox txtID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "iso_departement_name"
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
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "«· ’‰Ì›"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   " ’‰Ì›«  «·‰„«–Ã Ê «·⁄ﬁÊœ"
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
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "„”·”·"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   -360
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "iso_departement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As String

Private Sub Command1_Click(Index As Integer)
Select Case Index

Case 0
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields!iso_departement_name = ""
    Adodc1.Recordset.update
    Adodc1.Recordset.MoveLast
Case 1
    Adodc1.Recordset.update
Call Main.FillMenu
Case 2
X = MsgBox("Â· «‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbCritical + vbYesNo)
If X = vbNo Then
Exit Sub
End If

    If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.Delete
    Adodc1.Refresh
    DataGrid1.Refresh
    Call Main.FillMenu
    End If

Case 3
 

Case 4
On Error Resume Next
X = InputBox("«œŒ· «·—ﬁ„ «·„ÿ·Ê» «·»ÕÀ ⁄‰…")

        If IsNumeric(X) Then
        Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from  BOXs where BOX_id=" & X
        Adodc1.Refresh
        Else
        MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ ›ﬁÿ", vbCritical
        End If

Case 5
    X = InputBox("«œŒ· ﬂ·„… «·»ÕÀ")
            Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from BOXs where BOX_name like '%" & X & "%'"
        Adodc1.Refresh


End Select

End Sub

Private Sub Form_Activate()
user_priviliges_adodc.CommandType = adCmdText
If my_language = "E" Then
user_priviliges_adodc.RecordSource = "select * from USER_PRIVILIGES where employee_id=" & mainE.employee_id.text & "and [no]='" & screen_name.Caption & "'"

Else

user_priviliges_adodc.RecordSource = "select * from USER_PRIVILIGES where employee_id=" & Main.employee_id.text & "and [no]='" & screen_name.Caption & "'"
End If

user_priviliges_adodc.Refresh

If user_priviliges_adodc.Recordset.RecordCount = 0 Then
Exit Sub
End If

If user_priviliges_adodc.Recordset.Fields![View] = False Then
MsgBox "€Ì— „”„ÊÕ »«” Œœ«„ Â–… «·‘«‘…  ", vbCritical
 
Unload Me
End If

Command1(0).Enabled = user_priviliges_adodc.Recordset.Fields![add_new]
Command1(1).Enabled = user_priviliges_adodc.Recordset.Fields![Save]
Command1(2).Enabled = user_priviliges_adodc.Recordset.Fields![Delete]

End Sub

Private Sub Form_Load()
On Error Resume Next
'LoadSettings
Adodc1.ConnectionString = connection_string
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from  iso_department"
Adodc1.Refresh


 

 

user_priviliges_adodc.ConnectionString = connection_string
user_priviliges_adodc.CommandType = adCmdText
user_priviliges_adodc.RecordSource = "select * from  USER_PRIVILIGES"
user_priviliges_adodc.Refresh
End Sub
