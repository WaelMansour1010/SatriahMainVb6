VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form new_templates 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4740
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4740
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox id 
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Text            =   "0"
      Top             =   2880
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "new_tmplates_dat1.frx":0000
      Height          =   480
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   847
      _Version        =   393216
      BackColor       =   12632256
      ListField       =   "DepartmentName"
      BoundColumn     =   ""
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   585
      Left            =   960
      Top             =   3360
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
      Caption         =   "ÊÍÑíß"
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
   Begin ALLButtonS.ALLButton Command2 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "new_tmplates_dat1.frx":0015
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
      Caption         =   "ÊÍÏíÏ ÞÓã ááäãæÐÌ"
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
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "ÇÓã ÇáäãæÐÌ"
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
      Left            =   3480
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ÇáÞÓã ÇáÊÇÈÚ áÉ"
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
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "new_templates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As String

Private Sub Command2_Click()

    If loading_temolates.Visible = False Then Exit Sub
    If DataCombo1.text = "" Then MsgBox "  ÇÎÊÑ ÇáÞÓã ÇæáÇ", vbCritical: Exit Sub
 
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from templatesnew where templates_id= " & id.text
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.Fields!departement_name = DataCombo1.text
        Adodc1.Recordset.update
    End If

    Unload Me

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Form_Load()
    On Error Resume Next
    'LoadSettings
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from  TblEmpDepartments"
    Adodc1.Refresh
End Sub
