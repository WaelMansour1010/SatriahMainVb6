VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmıEmpMoveReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ﬁ«—Ì—  ‰ﬁ·«  «·„ÊŸ›"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   Icon            =   "FrmEmpMoveReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3915
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeClient1 
      Height          =   495
      Left            =   8280
      TabIndex        =   14
      Top             =   2760
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·„ÊŸ›"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   11640
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox ComGranty 
      Height          =   315
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   2205
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   390
      Width           =   10155
      Begin VB.CheckBox chkTransferReport 
         Alignment       =   1  'Right Justify
         Caption         =   " ﬁ—Ì— ‰ﬁ· «·„ÊŸ›Ì‰ »Ì‰ «·„‘«—Ì⁄"
         Height          =   195
         Left            =   4650
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1980
         Width           =   2535
      End
      Begin VB.CheckBox ChkStatus 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈ŸÂ«— ﬂ· «·„ÊŸ›Ì‰ „⁄ «·„‰ ÂÌ… Œœ„« Â„"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox TxtRemarks 
         Alignment       =   1  'Right Justify
         Height          =   765
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1200
         Width           =   3915
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   7680
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   118685699
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   5160
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   118685699
         CurrentDate     =   38887
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   3240
         TabIndex        =   21
         Top             =   240
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcmbToDepart 
         Height          =   315
         Left            =   5160
         TabIndex        =   22
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcmbToProject 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   735
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcmbToJob 
         Height          =   315
         Left            =   5160
         TabIndex        =   24
         Top             =   1140
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Height          =   195
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   45
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›…«·Õ«·Ì…"
         Height          =   195
         Index           =   5
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„·«ÕŸ…"
         Height          =   195
         Index           =   2
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Êﬁ⁄ «·Õ«·Ì"
         Height          =   195
         Index           =   0
         Left            =   3990
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï  «—ÌŒ"
         Height          =   195
         Index           =   3
         Left            =   6870
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1530
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰  «—ÌŒ"
         Height          =   195
         Index           =   4
         Left            =   9210
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1530
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁ”„ «·Õ«·Ì"
         Height          =   195
         Index           =   7
         Left            =   8895
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ›"
         Height          =   195
         Left            =   9300
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   720
      Width           =   915
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   3240
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "⁄—÷ «· ﬁ—Ì—"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   1
      Left            =   1410
      TabIndex        =   1
      Top             =   3240
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   495
      Index           =   2
      Left            =   30
      TabIndex        =   2
      Top             =   3240
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "Œ—ÊÃ"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeCar 
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   2760
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·ﬁ”„"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeModel 
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Top             =   2760
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·„Êﬁ⁄"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypePlate 
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Top             =   2760
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·ÊŸÌ›…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton RadioButton1 
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·„ÊŸ›"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   " ﬁ«—Ì—  ‰ﬁ·«  «·„ÊŸ›"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   120
      Width           =   2805
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3090
      Width           =   1785
   End
End
Attribute VB_Name = "FrmıEmpMoveReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch



Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       
    If chkTransferReport Then
        GetData2
    Else
        GetData
    End If
            
        Case 1
            clear_all Me
DtpDateFrom.value = ""
DtpDateTo.value = ""

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
            Case 3
'print_report
    End Select

End Sub



Private Sub GetData2()
  
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim StrWhere As String
    Dim Msg As String
    Dim s As String
       s = " select distinct P1.Id  Transaction_ID,IsNull(Pro1.project_name ,p1.Project_name) Emp_Comm , p1.Project_name,P11.Project_name,PDes.[des],PDes1.[des] des2,T1.Start_date Transaction_Date,"
       s = s & "T2.to_opr,t2.to_term "
       s = s & " ,T2.emp_name CusName,T2.[Start_date],T2.person_name  ,T2.to_project_name  , T2.opreration_type PaymentType ,T1.OpraID"
       s = s & "   From opr_Employee T1"
       s = s & " LEFT OUTER JOIN opr_employee_details T2"
       s = s & " ON T2.pk_id = T1.id"
       s = s & " AND T2.opr_type = T1.opr_type"
       s = s & " LEFT OUTER JOIN projects AS p1 ON T1.Project_id = p1.id"
       s = s & " LEFT OUTER JOIN projects AS p11 ON T1.Project_id1  = p11.id"
       s = s & "        LEFT OUTER JOIN projects Pro1"
       s = s & "        ON  Pro1.Id = T2.FromProjectID "
        s = s & "        LEFT OUTER JOIN TblEmployee"
         s = s & "        On TblEmployee.Emp_Id = T2.Emp_Id"
        
       s = s & " LEFT OUTER JOIN projects_des PDes ON  T1.term_Fullcode = PDes.project_id"
       s = s & " LEFT OUTER JOIN projects_des PDes1 ON  T1.term_Fullcode1  = PDes1.project_id"

       

       s = s & " Where T1.opr_type = 1"
      ' s = s & " And (T1.id = " & val(Me.xptxtid.Text) & ")"
       
If (Me.TxtRemarks.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblEmployee.Fullcode like '%" & Me.TxtRemarks.Text & "%'"
        
    End If
   If Me.DcboEmpName.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID=" & Me.DcboEmpName.BoundText & ""
      
    End If
    
    If Me.dcmbToProject.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND T1.Project_id = " & Me.dcmbToProject.BoundText & ""
      
    End If
'  If Me.DcmbToJob.BoundText <> "" Then
'
'            StrWhere = StrWhere & " AND dbo.TblMoveEmp1.JobTo=" & val(Me.DcmbToJob.BoundText) & ""
'
'    End If

    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND T1.Start_date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If

    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND  T1.Start_date <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     
    End If





        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EmpSalaryTransfer.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EmpSalaryTransfer.rpt"
        End If

        ''''''


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open s & StrWhere, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "?CE??I E?C?CE ?????"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo
    Dim oorderdate As Date
    Dim CBoBasedON As Integer
    Dim PONo As String

     
    
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " EIC?E ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(12).AddCurrentValueval (lbTotalMente.Caption)
  'xReport.ParameterFields(12).AddCurrentValue (dcproject.Text)
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
End Sub



Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Private Sub ChangeLang()

 
ChkStatus.Caption = "All Employees With End Service"
 
 Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "View Report"
   Cmd(2).Caption = "Exit"
  Me.Caption = "Reports of Movement Of Employee "
Label5.Caption = Me.Caption
Label1.Caption = "Emp"


lbl(7).Caption = "Current Dept"
lbl(0).Caption = "Current Location"
lbl(2).Caption = "Remraks"
lbl(5).Caption = "Current Job"
XPChkSearchTypeCar.RightToLeft = False
Me.XPChkSearchTypeCar.Caption = "By Dept"
Me.XPChkSearchTypeClient1.RightToLeft = False
Me.XPChkSearchTypeClient1.Caption = "By Emp"
Me.XPChkSearchTypePlate.RightToLeft = False
Me.XPChkSearchTypePlate.Caption = "By Job"
Me.XPChkSearchTypeModel.RightToLeft = False
Me.XPChkSearchTypeModel.Caption = "By Location"

lbl(3).Caption = "To Date"
lbl(4).Caption = "From Date"
End Sub

Private Sub Form_Load()
    'Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    
        Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500




    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
  Set Dcombos = New ClsDataCombos
     'Dcombos.GetUsers Me.DCboUserName
     Dcombos.GetEmployees Me.DcboEmpName
    
     Dcombos.GetEmpDepartments Me.DcmbToDepart
    
   
   Dcombos.GetEmpJobsTypes Me.DcmbToJob
   
   Dcombos.GetEmpLocations Me.dcmbToProject ' locatione
    Set DCboSearch = New clsDCboSearch
  '  Set DCboSearch.Client = Me.DcbClientname
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture


 
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
StrSQL = " SELECT     TOP 100 PERCENT dbo.TblMoveEmp1.ID, dbo.TblMoveEmp1.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
StrSQL = StrSQL & "                      dbo.TblMoveEmp1.RecordDate, dbo.TblMoveEmp1.EmpID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblMoveEmp1.FromDepart,"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblMoveEmp1.ToDepart,"
StrSQL = StrSQL & "                      TblEmpDepartments_1.DepartmentName AS DepartmentNameTo, TblEmpDepartments_1.DepartmentNamee AS DepartmentNameeTo, dbo.TblMoveEmp1.JobID,"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblMoveEmp1.JobTo, TblEmpJobsTypes_1.JobTypeName AS JobTypeNameTo,"
StrSQL = StrSQL & "                      TblEmpJobsTypes_1.JobTypeNamee AS JobTypeNameeTo, dbo.TblMoveEmp1.ProjectFrom, dbo.EmpGroupDep.GroupName, dbo.TblMoveEmp1.ProjectTo,"
StrSQL = StrSQL & "                      EmpGroupDep_1.GroupName AS GroupNameTo, dbo.TblMoveEmp1.basicSalary, dbo.TblMoveEmp1.Reson, dbo.TblMoveEmp1.PostedDate,"
StrSQL = StrSQL & "                      dbo.TblMoveEmp1.posted , dbo.TblMoveEmp1.Approved, dbo.TblMoveEmp1.DiffDate , dbo.TblEmployee.jopstatusid"
StrSQL = StrSQL & " FROM         dbo.TblMoveEmp1 LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.EmpGroupDep EmpGroupDep_1 ON dbo.TblMoveEmp1.ProjectTo = EmpGroupDep_1.GroupID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.EmpGroupDep ON dbo.TblMoveEmp1.ProjectFrom = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_1 ON dbo.TblMoveEmp1.JobTo = TblEmpJobsTypes_1.JobTypeID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes ON dbo.TblMoveEmp1.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments TblEmpDepartments_1 ON dbo.TblMoveEmp1.ToDepart = TblEmpDepartments_1.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblMoveEmp1.FromDepart = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblMoveEmp1.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblMoveEmp1.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " WHERE     (1 = 1) "
    BolBegine = False
    StrWhere = ""

   If ChkStatus.value = vbUnchecked Then
 StrSQL = StrSQL & " and dbo.TblEmployee.jopstatusid <> 2"
 StrSQL = StrSQL & " and dbo.TblEmployee.jopstatusid <> 5"
 StrSQL = StrSQL & " and dbo.TblEmployee.jopstatusid <> 6"
 End If

 
 If (Me.TxtRemarks.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblEmployee.Fullcode like '%" & Me.TxtRemarks.Text & "%'"
        
    End If
   If Me.DcboEmpName.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID=" & Me.DcboEmpName.BoundText & ""
      
    End If
    If Me.DcmbToDepart.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblMoveEmp1.ToDepart=" & Me.DcmbToDepart.BoundText & ""
      
    End If
    If Me.dcmbToProject.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblMoveEmp1.ProjectTo = " & Me.dcmbToProject.BoundText & ""
      
    End If
  If Me.DcmbToJob.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblMoveEmp1.JobTo=" & val(Me.DcmbToJob.BoundText) & ""
      
    End If

    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND dbo.TblMoveEmp1.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If

    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND  dbo.TblMoveEmp1.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     
    End If


    

    '-----------------------------------
StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " Order By dbo.TblMoveEmp1.ID"
  
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
    ' Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value
 print_report StrSQL
'print_report StrSQL
       ' With Me.Fg
       '     .Clear flexClearScrollable, flexClearEverything
       '     .Rows = .FixedRows
       '     .Rows = rs.RecordCount + .FixedRows
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           
 

    End If

End Sub
Function print_report(Optional NoteSerial As String)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



        If SystemOptions.UserInterface = ArabicInterface Then
        If Me.XPChkSearchTypeClient1.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpMove.rpt"
            Else
            If Me.XPChkSearchTypeCar.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpMoveDept.rpt"
            Else
            If Me.XPChkSearchTypeModel.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpMoveLocation.rpt"
            Else
             If Me.XPChkSearchTypePlate.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpMoveJob.rpt"
            Else
            
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpMoveAll.rpt"
            
            End If
            End If
            
            
            End If
             End If
        Else
               If Me.XPChkSearchTypeClient1.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpMove.rpt"
            Else
            If Me.XPChkSearchTypeCar.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpMoveDept.rpt"
            Else
            If Me.XPChkSearchTypeModel.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpMoveLocation.rpt"
            Else
             If Me.XPChkSearchTypePlate.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpMoveJob.rpt"
            Else
            
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpMoveAll.rpt"
        
            End If
            End If
            
            
            End If
             End If
           
        End If



    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  Dim Total As String
  Dim totl As Double
 ' totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
 ' total = totl
 '  xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
 '     xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
 '       xReport.ParameterFields(14).AddCurrentValue total
   ' xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function


 
Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub
