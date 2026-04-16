VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmmaintenace_type 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12180
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   12180
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   1455
      Begin VB.CommandButton Command1 
         Caption         =   "»«·—ﬁ„"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "»«·«”„"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "«·»ÕÀ"
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
         Left            =   360
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1455
      Begin VB.CommandButton Command1 
         Caption         =   "ÃœÌœ"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Õ›Ÿ"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Õ–›"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÿ»«⁄…"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   15
         Left            =   -120
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      DataField       =   "maintenance_id"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arabic Typesetting"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmmaintenace_type.frx":0000
      Height          =   3855
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "maintenance_id"
         Caption         =   "—ﬁ„"
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
         DataField       =   "maintenance_type"
         Caption         =   "‰Ê⁄ «·’Ì«‰…"
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
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   7560
      Top             =   2520
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   1032
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Transporter"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Transporter"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "maintenance_type"
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
      DataField       =   "maintenance_type"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arabic Typesetting"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "«‰Ê«⁄ «·’Ì«‰…"
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
      Left            =   4080
      TabIndex        =   15
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "‰Ê⁄ «·’Ì«‰…"
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
      Left            =   10680
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "—ﬁ„"
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
      Left            =   10800
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "frmmaintenace_type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As String

Private Sub Command1_Click(Index As Integer)
Select Case Index

Case 0
    Adodc1.Recordset.AddNew
Case 1
    Adodc1.Recordset.Update

Case 2
x = MsgBox("Â· «‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbCritical + vbYesNo)
If x = vbNo Then
Exit Sub
End If

    If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.Delete
    Adodc1.Refresh
    DataGrid1.Refresh
    End If

Case 3
    If Adodc1.Recordset.RecordCount > 0 Then
    
    Form3.case_id = Me.Name
   
    Form3.Show
    End If

Case 4
On Error Resume Next
x = InputBox("«œŒ· «·—ﬁ„ «·„ÿ·Ê» «·»ÕÀ ⁄‰…")

        If IsNumeric(x) Then
        Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from  maintenance_type where maintenance_id=" & x
        Adodc1.Refresh
        Else
        MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ ›ﬁÿ", vbCritical
        End If

Case 5
    x = InputBox("«œŒ· ﬂ·„… «·»ÕÀ")
            Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from maintenance_type where maintenance_type like '%" & x & "%'"
        Adodc1.Refresh


End Select

End Sub

