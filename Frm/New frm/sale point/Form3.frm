VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form Form3 
   Caption         =   "ØÈÇÚÉ ÇáÊÞÇÑíÑ"
   ClientHeight    =   9045
   ClientLeft      =   195
   ClientTop       =   495
   ClientWidth     =   15120
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "äØÇÞ ÇáÈÍË"
      Height          =   855
      Left            =   10440
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form3.frx":0000
         Left            =   240
         List            =   "Form3.frx":000D
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   12360
      TabIndex        =   27
      Top             =   0
      Width           =   2775
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÊÞÑíÑ ÑÞã"
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
         Height          =   615
         Left            =   840
         TabIndex        =   29
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label case_id 
         Alignment       =   2  'Center
         Caption         =   "14"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc17 
      Height          =   375
      Left            =   7560
      Top             =   120
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
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   16
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
      MICON           =   "Form3.frx":003A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame infoA 
      Height          =   735
      Left            =   8640
      TabIndex        =   22
      Top             =   13320
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label dep_a 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label xx 
         Caption         =   "ÇáãæÙÝ ÇáÍÇáí"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4560
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label yy 
         Caption         =   "ÇáÞÓã"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label emp_a 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2880
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame InfoE 
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   13320
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label zz 
         Caption         =   "Departemnt"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.Label emp_name_lbl 
         Caption         =   "Label7"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label dept_lbl 
         Caption         =   "Departement"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label vv 
         Caption         =   "Employee name"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   600
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   10335
      Begin VB.TextBox TxtEmp_Code 
         Height          =   375
         Left            =   3720
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label SHOWPICTURE 
         Height          =   495
         Left            =   3960
         TabIndex        =   33
         Top             =   480
         Width           =   975
      End
      Begin VB.Label noofmonth 
         Height          =   255
         Left            =   4680
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Åáì"
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
         Left            =   6600
         TabIndex        =   14
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ãä"
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
         Left            =   10080
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label from_date 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   8400
         TabIndex        =   12
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label to_date 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   4200
         TabIndex        =   11
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "ÑÞã ÇáÊÞÑíÑ"
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
         Left            =   2160
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRV1 
      Height          =   8895
      Left            =   -120
      TabIndex        =   8
      Top             =   120
      Width           =   18885
      lastProp        =   600
      _cx             =   33311
      _cy             =   15690
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ØÈÇÚÉ "
      Height          =   735
      Left            =   7080
      TabIndex        =   0
      Top             =   12840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label sqllbl 
      Caption         =   "sql"
      Height          =   135
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label table 
      Caption         =   "0120"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   13200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label caller 
      Caption         =   "1"
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label criteria_user_id 
      Caption         =   "criteria_user_id"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   12600
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label shift_no 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   11880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label user_name 
      Caption         =   "Label1"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label user_id 
      Caption         =   "Label1"
      Height          =   735
      Left            =   7680
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label bill_id 
      Caption         =   "197"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rep As CRAXDRT.Report
Dim cry As New CRAXDRT.Application
Dim found As Boolean
    
Dim StrRS As New ADODB.Recordset
Dim rs As New ADODB.Recordset

'Option Explicit

Dim sql As String
Dim strConnect As String

'Dim m_Connection As adodb.Connection
'Dim adoRS As adodb.Recordset

Dim user_r_id  As String
Dim user_r_pw  As String

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

Private Sub Command2_Click()
    'Set rs = New adodb.Recordset
    On Error Resume Next

    sql = "select * from maintenance_all_details_qry where opr_id=183"
 
    rs.Open sql, connection_string, adOpenStatic, adLockReadOnly

    'below opening the rpt file
    Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT4.rpt")

    'below for setting the logon information. it works even if
    'ur database (access or sql server) is password protected.
    'if u r using access database  then the server name can be "Sql server" or even blank it does not matter as if it is blank it will use the server which was used at design time.

    'for access u can use
    '        CrxReport.Database.Tables(i).SetLogOnInfo "", "database", "usernm", "password"

    For i = 1 To rep.Database.tables.count
        rep.Database.tables(i).SetLogOnInfo server_name, database_name, "salim", Password
    Next

    With rep
        .DiscardSavedData
        .Database.SetDataSource rs
    End With

    CRV1.ReportSource = rep
    CRV1.viewReport
        
    While CRV1.IsBusy

        DoEvents
    Wend

End Sub

Private Sub Combo1_Click()
    Dim X As String
    X = Combo1.text

    If Combo1.text = "opening balance" Then

        X = "ÑÕíÏ ÇÝÊÊÇÍí"
 
    Else

        If Combo1.text = "issue voucher" Then
            X = "ÓäÏ ÕÑÝ ãÎÒäí"

        Else

            If Combo1.text = "recive voucher" Then
                X = "ÓäÏ ÇÓÊáÇã"

            End If
        End If
    End If

    If case_id.Caption = 5 Then
    
        ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, "x", ""

        sql = "SELECT * from inventory WHERE  transaction_type='" & X & "' AND branch_no=" & Branch_NO & " and  inventory_id like'%" & REPORTSFRM.DataCombo1.BoundText & "%' and (transaction_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND transaction_date <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) )  order by    item_code"
        rep.SQLQueryString = sql
        rep.DiscardSavedData
     
    End If
     
    If case_id.Caption = 6 Then
        sql = "SELECT * from inventory WHERE   transaction_type='" & X & "' AND branch_no=" & Branch_NO & " and   (transaction_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND transaction_date <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102)  and item_code like'%" & REPORTSFRM.DataCombo2.text & "%' and inventory_id like'%" & REPORTSFRM.DataCombo1.BoundText & "%')  order by    item_code"

        rep.SQLQueryString = sql
        rep.DiscardSavedData
 
    End If
    
    If my_language = "E" Then
        rep.reporttitle = Adodc1.Recordset.Fields!company_name_E
    Else
        rep.reporttitle = Adodc1.Recordset.Fields!Company_Name
    End If

    CRV1.ReportSource = rep
    '  CRV1.RefreshEx True
    CRV1.viewReport
    CRV1.Zoom 90
End Sub

Function SHOWPIC(PICNAME As String)
    Dim xLogo As CRAXDRT.OLEObject
    StrFileName = App.path & "\Images\" & PICNAME & ".JPG"

    Set xLogo = rep.Areas(3).Sections(1).AddPictureObject(StrFileName, 120, 300)
    xLogo.Width = 1700
    xLogo.Height = 1700
    xLogo.backcolor = vbWhite
    xLogo.BorderColor = 255
    xLogo.CloseAtPageBreak = True
    '  xLogo.HyperlinkText = "BYTE"
    '  xLogo.HyperlinkType = crHyperlinkWebsite
    '  rep.Areas(1).Sections(1).SuppressIfBlank = True
    '  rep.Areas(1).Sections(1).Height = xLogo.Height + 250
 
End Function

Private Sub Form_Activate()

    On Error Resume Next
    CRV1.Width = Me.Width - 70
    CRV1.Height = Me.Height - 600

    'user_r_id = "salimman2003"
    'user_r_pw = "salimman2003"

    If found = True Then
        Exit Sub
    End If

    'cry.LogOnServer , server_name, database_name
 
    Select Case case_id

        Case 1
        
            sql = "SELECT * from emp_all_details  WHERE Fullcode='" & FrmEmployee.DCPreFix.BoundText & Trim(txtid.text) & "'"
            rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

            Set rep = cry.OpenReport(system_path & "\reports\emp\REPORT1.rpt")
            rep.Database.SetDataSource rs
            Set FrmReport = New FrmReportViewer
            FrmReport.CRViewer.ReportSource = rep
            FrmReport.CRViewer.viewReport
            FrmReport.show
            Screen.MousePointer = vbDefault
            SendKeys "{RIGHT}"
            End
            '   CRV1.RefreshEx True
   
            '    rep.SQLQueryString = Sql
            '      rep.DiscardSavedData
            '     rep.ReportTitle = noofmonth.Caption
     
            '         CRV1.ReportSource = rep
    
            '    CRV1.ViewReport
    
            Exit Sub
     
        Case 2

            Set rep = cry.OpenReport(system_path & "\reports\emp\REPORT2.rpt")
            sql = "SELECT * from emp_all_details WHERE    Fullcode='" & FrmEmployee.DCPreFix.BoundText & Trim(txtid.text) & "'"

            rep.SQLQueryString = sql
            rep.DiscardSavedData
            rep.reporttitle = noofmonth.Caption

        Case 3

            Set rep = cry.OpenReport(system_path & "\reports\emp\REPORT2.rpt")
            sql = "SELECT * from emp_all_details WHERE   Fullcode='" & FrmEmployee.DCPreFix.BoundText & Trim(txtid.text) & "'"

            rep.SQLQueryString = sql
            rep.DiscardSavedData

        Case 4

            Set rep = cry.OpenReport(system_path & "\reports\emp\REPORT2.rpt")
            sql = "SELECT * from emp_all_details WHERE   Fullcode='" & FrmEmployee.DCPreFix.BoundText & Trim(txtid.text) & "'"

            rep.SQLQueryString = sql
            rep.DiscardSavedData
 
        Case 5

            Set rep = cry.OpenReport(system_path & "\reports\emp\REPORT5.rpt")
            sql = "SELECT * from emp_all_details WHERE   Fullcode='" & FrmEmployee.DCPreFix.BoundText & Trim(txtid.text) & "'"

            rep.SQLQueryString = sql
            rep.DiscardSavedData
            rep.reporttitle = noofmonth.Caption
    
        Case 6

            Set rep = cry.OpenReport(system_path & "\reports\emp\REPORT6.rpt")
            sql = "SELECT * from emp_all_details WHERE   Fullcode='" & FrmEmployee.DCPreFix.BoundText & Trim(txtid.text) & "'"

            rep.SQLQueryString = sql
            rep.DiscardSavedData
   
        Case 7

            Set rep = cry.OpenReport(system_path & "\reports\emp\REPORT7.rpt")
            sql = "SELECT * from emp_all_details WHERE   Fullcode='" & FrmEmployee.DCPreFix.BoundText & Trim(txtid.text) & "'"

            rep.SQLQueryString = sql
            rep.DiscardSavedData
    
        Case 8

            Set rep = cry.OpenReport(system_path & "\reports\emp\REPORT8.rpt")
            sql = "SELECT * from emp_all_details WHERE   Fullcode='" & FrmEmployee.DCPreFix.BoundText & Trim(txtid.text) & "'"

            rep.SQLQueryString = sql
            rep.DiscardSavedData
     
        Case 9

            Set rep = cry.OpenReport(system_path & "\reports\emp\REPORT9.rpt")
            sql = Me.sqllbl

            rep.SQLQueryString = sql
            rep.DiscardSavedData
    
        Case 10

            Set rep = cry.OpenReport(system_path & "\reports\emp\REPORT10.rpt")
            '  Sql = Me.sqllbl
            '
            '     rep.SQLQueryString = Sql
            '     rep.DiscardSavedData
     
        Case 11
  
            Set rep = cry.OpenReport(system_path & "\reports\emp\REPORT11.rpt")

        Case 12

            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT12.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT12.rpt")
    
            End If
    
            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw
     
            sql = "SELECT * from car_not_work WHERE   (DATE1 >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND DATE1 <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102)  and cars_NO like'%" & REPORTSFRM.DataCombo4.text & "%' and trip_diffrent_days>0)"

            rep.SQLQueryString = sql
            rep.DiscardSavedData
    
        Case 13

            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT13.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT13.rpt")
    
            End If
    
            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw
     
            sql = "SELECT * from time_lose_qry WHERE   (DATE >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND DATE <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102)  and cars_NO ='" & REPORTSFRM.DataCombo4.text & "' and km_differenr>0)"

            rep.SQLQueryString = sql
            rep.DiscardSavedData
    
        Case 14

            If my_language = "E" Then
                If REPORTSFRM.Check1.value = 0 Then

                    Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT14.rpt")
                Else
                    Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT14m.rpt")

                End If

            Else

                If REPORTSFRM.Check1.value = 0 Then
                    Set rep = cry.OpenReport(system_path & "\reports\REPORT14.rpt")

                Else
                    Set rep = cry.OpenReport(system_path & "\reports\REPORT14m.rpt")
 
                End If
    
            End If
    
            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw
     
            sql = "SELECT * from maintenance WHERE   (OPr_DATE >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND OPr_DATE <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and (car_no like'%" & REPORTSFRM.DataCombo4.text & "%' and Driver_name like'%" & REPORTSFRM.DataCombo3.text & "%')) ORDER BY  car_no"
 
            rep.SQLQueryString = sql
            rep.DiscardSavedData
    
        Case 15

            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT4.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT4.rpt")
    
            End If
    
            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw
     
            Adodc17.ConnectionString = connection_string
            Adodc17.CommandType = adCmdText
            Adodc17.RecordSource = "select * from maintenance_all_details_qry  WHERE   (OPr_DATE >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND OPr_DATE <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and (car_NO like'%" & REPORTSFRM.DataCombo11.text & "%' and Driver_name like'%" & REPORTSFRM.DataCombo10.text & "%'))"
            Adodc17.Refresh

            If Adodc17.Recordset.RecordCount > 1 Then
                sql = "SELECT * from  maintenance_all_details_qry  WHERE not (item_code is null) and   (OPr_DATE >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND OPr_DATE <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and (car_NO like'%" & REPORTSFRM.DataCombo11.text & "%' and Driver_name like'%" & REPORTSFRM.DataCombo10.text & "%') ) ORDER By opr_id"

                ' Form3.sqllbl = "select * from maintenance_all_details_qry where not (item_code is null) and opr_id=" & Text1.Text

            Else
                sql = "SELECT * from  maintenance_all_details_qry  WHERE   (OPr_DATE >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND OPr_DATE <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and (car_NO like'%" & REPORTSFRM.DataCombo11.text & "%' and Driver_name like'%" & REPORTSFRM.DataCombo10.text & "%') ) ORDER By opr_id"
                '    Form3.sqllbl = "select * from maintenance_all_details_qry where opr_id=" & Text1.Text

            End If

            If REPORTSFRM.DataCombo12.text <> "" And IsNumeric(REPORTSFRM.DataCombo12.text) Then
         
                Adodc17.ConnectionString = connection_string
                Adodc17.CommandType = adCmdText
                Adodc17.RecordSource = "SELECT * from maintenance_all_details_qry WHERE    Amr_shogl_no=" & REPORTSFRM.DataCombo12.text
                Adodc17.Refresh

                If Adodc17.Recordset.RecordCount > 1 Then
                    sql = "SELECT * from maintenance_all_details_qry WHERE not (item_code is null) and    Amr_shogl_no=" & REPORTSFRM.DataCombo12.text
 
                    ' Form3.sqllbl = "select * from maintenance_all_details_qry where not (item_code is null) and opr_id=" & Text1.Text

                Else
                    sql = "SELECT * from maintenance_all_details_qry WHERE    Amr_shogl_no=" & REPORTSFRM.DataCombo12.text
                    '    Form3.sqllbl = "select * from maintenance_all_details_qry where opr_id=" & Text1.Text

                End If

            End If
    
            '     Sql = "SELECT * from  maintenance_all_details_qry  WHERE   (OPr_DATE >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND OPr_DATE <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and (car_NO like'%" & REPORTSFRM.DataCombo11.Text & "%' and Driver_name like'%" & REPORTSFRM.DataCombo10.Text & "%'))"

            rep.SQLQueryString = sql
            rep.DiscardSavedData
    
        Case 16
  
            system_path = App.path

            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT16.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT16.rpt")
    
            End If

            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw
     
            sql = "SELECT * from sand_all_details_qry  WHERE  sanad_no= " & frmsandat_ked.Text1.text

            rep.SQLQueryString = sql
            rep.DiscardSavedData
     
        Case 18

            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT18.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT18.rpt")
    
            End If

            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw
     
            sql = "SELECT * from sand_all_details_qry  WHERE  sandat_pc_no= " & frmsandat_sarf.Text7.text

            rep.SQLQueryString = sql
            rep.DiscardSavedData
     
        Case 17

            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT17.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT17.rpt")
    
            End If

            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw
     
            sql = "SELECT * from sand_all_details_qry  WHERE  sandat_pc_no= " & frmsandat_kabd.Text7.text

            rep.SQLQueryString = sql
            rep.DiscardSavedData
     
        Case 19

            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT19.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT19.rpt")
    
            End If

            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw
            sql = "select *  from items_risk_alarm where branch_no=" & Branch_NO & " and departement='" & departement_name & "'"
            '  Sql = "SELECT * from sand_all_details_qry  WHERE  sandat_pc_no= " & frmsandat_sarf.Text7.Text

            rep.SQLQueryString = sql
            rep.DiscardSavedData
     
        Case 20

            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT8.rpt")
                sql = "SELECT * from travel_all_details WHERE    transactions_id=" & transactionsE.Text1.text

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT8.rpt")

                sql = "SELECT * from travel_all_details WHERE    transactions_id=" & frmtransactions.Text1.text
            End If
    
            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw

            rep.SQLQueryString = sql
            rep.DiscardSavedData
     
        Case 21
  
            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT7.rpt")
      
                sql = "SELECT * from transactions WHERE    transactions_id=" & transactionsE.Text1.text

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT7.rpt")

                sql = "SELECT * from transactions WHERE    transactions_id=" & frmtransactions.Text1.text

            End If

            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw

            rep.SQLQueryString = sql
            rep.DiscardSavedData
     
        Case 25

            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT25.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT25.rpt")
    
            End If

            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw
      
            sql = "SELECT * from car_maintenaces  WHERE  car_no='" & FrmCars.DCPreFix.text & FrmCars.txtid.text & "'"

            rep.SQLQueryString = sql
            rep.DiscardSavedData
     
        Case 26

            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT26.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT26.rpt")
    
            End If

            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw
      
            sql = "SELECT * from items  WHERE  group_name='" & items_groups.Text3.text & "'"

            rep.SQLQueryString = sql
            rep.DiscardSavedData
    
        Case 330

            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT330.rpt")
            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT330.rpt")
            End If

            '  Sql = "SELECT * from [log_files] WHERE   (log_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND log_date <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) )"
            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, user_r_id, user_r_pw

            sql = "SELECT opr_id, log_date, log_time, user_name, process_name, process_text, SUBJECT_NO  From  dbo.[log_files]  WHERE     (log_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (log_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and user_name='" & REPORTSFRM.DataCombo9.text & "')"

            '  Sql = sqllbl.Caption
            rep.SQLQueryString = sql
            rep.DiscardSavedData
            'CRV1.DisplayGroupTree = True
            '   CRV1.ReportSource = rep
            '  CRV1.RefreshEx True
            '  CRV1.ViewReport
            ' CRV1.Zoom 100
 
        Case 27
 
            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT27.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT27.rpt")
    
            End If
    
            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, "x", ""

            sql = "SELECT * from inventory WHERE  branch_no=" & Branch_NO & " and sanad_no='" & sand_ESTLAM_inventory.Text3.text & "'"
            rep.SQLQueryString = sql
            rep.DiscardSavedData
   
        Case 28
 
            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT27.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT27.rpt")
    
            End If
    
            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, "x", ""

            sql = "SELECT * from inventory WHERE  branch_no=" & Branch_NO & " and sanad_no='" & sand_sarf_inventory.Text3.text & "'"
            rep.SQLQueryString = sql
            rep.DiscardSavedData
      
        Case 40
 
            If my_language = "E" Then
                Set rep = cry.OpenReport(system_path & "\reports\EN\REPORT40.rpt")

            Else
                Set rep = cry.OpenReport(system_path & "\reports\REPORT40.rpt")
    
            End If
    
            ' rep.Database.Tables(1).SetLogOnInfo server_name, database_name, "x", ""
            'Sql = "SELECT  cars_NO, SUM(trip_value) AS TRIP_SUM, SUM(driver_value) AS DRIVER_ALL, SUM(another_masrouf) AS ANOTHER_ALL, SUM(commision) AS COMMISION_ALL,   COUNT(*) AS TOTAL   FROM         (SELECT     * From transactions Where branch_no = " & branch_no & " AND ([date] >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND [date] <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102))) DERIVEDTBL GROUP BY cars_NO"
            'Sql = "SELECT * from inventory WHERE  branch_no=" & branch_no & " and sanad_no='" & sand_sarf_inventory.Text3.text & "'"
     
            'Sql = "SELECT     cars_NO, SUM(TRIP_SUM) AS trip_sum1, SUM(DRIVER_ALL) AS driver_sum, SUM(ANOTHER_ALL) AS another_all1, SUM(COMMISION_ALL) AS commision_all1,  SUM(total) As total1 From dbo.xx WHERE     ([date] >= CONVERT(DATETIME, '2011-01-22 00:00:00', 102)) AND ([date] <= CONVERT(DATETIME, '2011-01-30 00:00:00', 102)) GROUP BY cars_NO"
            'Sql = "SELECT     cars_NO, SUM(TRIP_SUM) AS trip_sum1, SUM(DRIVER_ALL) AS driver_sum, SUM(ANOTHER_ALL) AS another_all1, SUM(COMMISION_ALL) AS commision_all1,  SUM(total) As total1 From dbo.xx WHERE     ([date] >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & "  00:00:00', 102)) AND ([date] <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102)) AND (branch_no =" & branch_no & ") GROUP BY cars_NO order by cars_NO "
            sql = "SELECT     * from xx WHERE  not cars_no='' and   ([date] >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & "  00:00:00', 102)) AND ([date] <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102)) GROUP BY cars_NO,driver_name order by cars_NO,driver_name "
            rep.SQLQueryString = sql
            rep.DiscardSavedData
            rep.ReportComments = REPORTSFRM.from_date
            rep.ReportAuthor = REPORTSFRM.to_date
   
    End Select
   
    If SHOWPICTURE.Caption = 1 And Me.case_id < 10 Then
        SHOWPIC (Me.TxtEmp_Code.text)
    End If

    CRV1.ReportSource = rep
    CRV1.RefreshEx True
    CRV1.viewReport
   
    CRV1.Zoom 90
    found = True
    
End Sub

' *************************************************************
' Load the Report in the viewer
'
Private Sub Form_Load()
    found = False
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 700
    
    ' Me.Width = MDIFrmMain.Width - 500
    '    Me.Height = MDIFrmMain.Height - 500
    On Error Resume Next
    '
 
    If my_language = "E" Then
        Frame2.Caption = "Filter"
        Combo1.Clear
 
        Combo1.AddItem "opening balance"
        Combo1.AddItem "issue voucher"
        Combo1.AddItem "recive voucher"

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

    Me.left = (MDIForm1.Width - Me.Width) / 2
    Me.top = (MDIForm1.Height - Me.Height) / 2 - 500

    If my_language = "E" Then
        CMD_language.Caption = "ÚÑÈí"
    End If
 
    'LoadSettings

    'Adodc1.ConnectionString = connection_string
    'Adodc1.CommandType = adCmdText
    'Adodc1.RecordSource = "select *  from info"
    'Adodc1.Refresh
 
    ' Set m_Connection = New adodb.Connection
    '  Set adoRS = New adodb.Recordset

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
 
End Sub
 
