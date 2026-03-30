VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ReportApp As New CRAXDRT.Application

Private ReportShow As CRAXDRT.Report

Private FrmShowRep As FrmReportViewer
Dim RPTCompany_Name_Arabic As String
Dim RPTComment_Arabic As String
Dim RPTCompany_Name_Eng As String
Dim RPTComment_Eng As String

  
  
   Public Function ShowEmpDataRep(Optional empid As Double = 0, Optional ByVal mWhere As String = "")
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.GetEmployeeSalary(dbo.TblEmployee.Emp_ID, GETDATE()) AS Salary, dbo.TblEmployee.*"
MySQL = MySQL & " FROM         dbo.TblEmployee INNER JOIN"
MySQL = MySQL & "                      dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id"
MySQL = MySQL & " Where (dbo.jopstatus.resignationInt Is Null)"
If empid <> 0 Then
MySQL = MySQL & " and dbo.TblEmployee.Emp_ID=" & empid & " "
End If
If mWhere <> "" Then
    MySQL = MySQL & mWhere
End If

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmp_Data.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepEmp_DataE.rpt"
        End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
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
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng

        StrReportTitle = ""
    End If

    xReport.ParameterFields(2).AddCurrentValue user_name

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






Private Sub Class_Initialize()
    Dim RsOpt As ADODB.Recordset
    Set RsOpt = New ADODB.Recordset

    RsOpt.Open "TBLOPTIONS", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    RPTCompany_Name_Arabic = IIf(IsNull(RsOpt!Company_Arabic_Name), "", RsOpt!Company_Arabic_Name)   'rs!Company_Name_Arabic
    RPTComment_Arabic = IIf(IsNull(RsOpt!Company_Comment), "", RsOpt!Company_Comment)    'rs!Comment_Arabic
    RPTCompany_Name_Eng = IIf(IsNull(RsOpt!Company_Name), "", RsOpt!Company_Name)   'rs!Company_Name_Eng
    RPTComment_Eng = IIf(IsNull(RsOpt!Company_Comment), "", RsOpt!Company_Comment)   'rs!Comment_Eng
    Set RsOpt = Nothing
End Sub

'*******************************************************************************
Public Function ShowKhasmMkafeaRep(ByVal Wher As String, _
                                   Optional StrReportTitle As String)
    On Error GoTo ErrTrap
    Dim StrRepFile As String
    Dim RepNotExist As String
    Dim RsRep As ADODB.Recordset
    Dim MySQL As String

    'StrRepFile = App.Path & "\Reports\RepKhsmMkafea.rpt"

    StrRepFile = App.path & "\Reports\KHASMEDAFA.rpt"
    RepNotExist = Dir(StrRepFile, vbNormal)

    If RepNotExist = "" Then
        MsgBox "⁄›Ê« „·› «· ﬁ—Ì— €Ì— „ÊÃÊœ", vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    Set RsRep = New ADODB.Recordset
    'MySQL = "select * From QryUnionKhsmMkfea() where Emp_ID is not null " & Wher
    MySQL = "select * From khasmedafaview where Emp_ID is not null " & Wher

    RsRep.Open MySQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    If RsRep.RecordCount = 0 Then
        MsgBox "⁄›Ê« ·«  ÊÃœ »Ì«‰«  ··⁄—÷", vbOKOnly + vbMsgBoxRight, App.title
        GoTo EXT
    End If

    Class_Initialize
    Set ReportShow = ReportApp.OpenReport(StrRepFile)
    ReportShow.Database.SetDataSource RsRep

    If SystemOptions.UserInterface = ArabicInterface Then
        ReportShow.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ReportShow.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        '  StrReportTitle = " ﬁ—Ì— «·Œ’Ê„«  Ê «·«÷«›Ì " '& StrAccountName
 
    Else
        ReportShow.ParameterFields(1).AddCurrentValue RPTCompany_Name_Eng
        ReportShow.ParameterFields(2).AddCurrentValue RPTComment_Eng
        '  StrReportTitle = "    " '& StrAccountName & " Account"
 
    End If

    ReportShow.ParameterFields(3).AddCurrentValue user_name
    ReportShow.reporttitle = StrReportTitle
    ReportShow.EnableParameterPrompting = False
    ReportShow.ApplicationName = App.title
    ReportShow.ReportAuthor = App.title

    ReportShow.EnableParameterPrompting = False

    Set FrmShowRep = New FrmReportViewer

    With FrmShowRep
        .CRViewer.ReportSource = ReportShow
        .CRViewer.viewReport

        Do While .CRViewer.IsBusy = True
            DoEvents
        Loop

        Screen.MousePointer = vbDefault
        .show vbModal
    End With

EXT:

    RsRep.Close
    Set RsRep = Nothing
    Set ReportShow = Nothing
    Set ReportApp = Nothing
ErrTrap:
    Screen.MousePointer = vbDefault
End Function

'RepPresentTime.rpt
Public Function ShowRepPresentTime(ByVal Wher As String)
    On Error GoTo ErrTrap
    Dim StrRepFile As String
    Dim RepNotExist As String
    Dim RsRep As ADODB.Recordset
    Dim MySQL As String

    StrRepFile = App.path & "\Reports\RepPresentTime.rpt"
    RepNotExist = Dir(StrRepFile, vbNormal)

    If RepNotExist = "" Then
        MsgBox "⁄›Ê« „·› «· ﬁ—Ì— €Ì— „ÊÃÊœ", vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    Set RsRep = New ADODB.Recordset
    MySQL = "select * From QryRepPresentAbsent where Emp_ID is not null " & Wher
    RsRep.Open MySQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    If RsRep.RecordCount = 0 Then
        MsgBox "⁄›Ê« ·«  ÊÃœ »Ì«‰«  ··⁄—÷", vbOKOnly + vbMsgBoxRight, App.title
        GoTo EXT
    End If

    Set ReportShow = ReportApp.OpenReport(StrRepFile)
    ReportShow.Database.SetDataSource RsRep
    ReportShow.EnableParameterPrompting = False

    Set FrmShowRep = New FrmReportViewer

    With FrmShowRep
        .CRViewer.ReportSource = ReportShow
        .CRViewer.viewReport

        Do While .CRViewer.IsBusy = True
            DoEvents
        Loop

        Screen.MousePointer = vbDefault

        .show vbModal
    End With

EXT:

    RsRep.Close
    Set RsRep = Nothing
    Set ReportShow = Nothing
    Set ReportApp = Nothing
ErrTrap:
    Screen.MousePointer = vbDefault
End Function

'RepPresentTime.rpt
Public Function ShowRepSalary(Wher As String)
    'On Error GoTo ErrTrap
    Dim StrRepFile As String
    Dim RepNotExist As String
    Dim RsRep As ADODB.Recordset
    Dim MySQL As String
    Dim RsExp As ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim SQLExp2 As String
    'Dim SQLExp As String
    StrRepFile = App.path & "\Reports\RepEmpAllValues.rpt"

    RepNotExist = Dir(StrRepFile, vbNormal)

    If RepNotExist = "" Then
        MsgBox "⁄›Ê« „·› «· ﬁ—Ì— €Ì— „ÊÃÊœ", vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    Set RsRep = New ADODB.Recordset
    MySQL = "select * From QrySalaryvalua where Emp_ID is not null " & Wher
    RsRep.Open MySQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    If RsRep.RecordCount = 0 Then
        MsgBox "⁄›Ê« ·«  ÊÃœ »Ì«‰«  ··⁄—÷", vbOKOnly + vbMsgBoxRight, App.title
        GoTo EXT
    End If

    Set ReportShow = ReportApp.OpenReport(StrRepFile)
    ReportShow.Database.SetDataSource RsRep
    ReportShow.EnableParameterPrompting = False

    Set FrmShowRep = New FrmReportViewer

    Set RsExp = New ADODB.Recordset

    Dim SQLExp As String
 
    SQLExp = "SELECT (Go_HourTime+12)-Bring_HourTime as whour from tblTimeSetting"
    RsExp.Open SQLExp, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If RsExp.RecordCount > 0 Then

        ReportShow.ParameterFields(1).AddCurrentValue IIf(IsNull(Format(RsExp("whour").value, "00.00")), 0, Format(RsExp("whour").value, "00.00"))
    Else
        ReportShow.ParameterFields(1).AddCurrentValue CStr(0)
    End If

    RsExp.Close

    'SQLExp2 = "select * from VIEW13 where Emp_ID  IS NOT NULL " & Wher
    'RsExp.Open SQLExp2, Cn, adOpenStatic, adLockPessimistic, adCmdText
    'If RsExp.RecordCount > 0 Then
    '    xReport.ParameterFields(0).AddCurrentValue IIf(IsNull(Format(RsExp("TotaDay").Value, "00.00")), 0, Format(RsExp("TotaDay").Value, "00.00"))
    '    'xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(RsExp("TotaDay").Value), "0", RsExp("TotaDay").Value)
    '' IIf(IsNull(RsExp("TotaDay").Value), 0, RsExp("TotaDay").Value)
    'Else
    '    xReport.ParameterFields(1).AddCurrentValue "0"
    'End If
    'RsExp.Close

    With FrmShowRep
        .CRViewer.ReportSource = ReportShow
        .CRViewer.viewReport

        Do While .CRViewer.IsBusy = True
            DoEvents
        Loop

        Screen.MousePointer = vbDefault

        .show
    End With

EXT:

    RsRep.Close

    Set RsRep = Nothing

    Set ReportShow = Nothing

    'Set ReportApp = Nothing
ErrTrap:
    Screen.MousePointer = vbDefault
End Function

'RepPresentTime.rpt
Public Function ShowRepAbsent(ByVal Wher As String)
    On Error GoTo ErrTrap
    Dim StrRepFile As String
    Dim RepNotExist As String
    Dim RsRep As ADODB.Recordset
    Dim MySQL As String

    StrRepFile = App.path & "\Reports\RepAbsentReport.rpt"
    RepNotExist = Dir(StrRepFile, vbNormal)

    If RepNotExist = "" Then
        MsgBox "⁄›Ê« „·› «· ﬁ—Ì— €Ì— „ÊÃÊœ", vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    Set RsRep = New ADODB.Recordset
    MySQL = "select * From QryAbsentEmp where Emp_ID is not null " & Wher
    RsRep.Open MySQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    If RsRep.RecordCount = 0 Then
        MsgBox "⁄›Ê« ·«  ÊÃœ »Ì«‰«  ··⁄—÷", vbOKOnly + vbMsgBoxRight, App.title
        GoTo EXT
    End If

    Set ReportShow = ReportApp.OpenReport(StrRepFile)
    ReportShow.Database.SetDataSource RsRep
    ReportShow.EnableParameterPrompting = False

    Set FrmShowRep = New FrmReportViewer

    With FrmShowRep
        .CRViewer.ReportSource = ReportShow
        .CRViewer.viewReport

        Do While .CRViewer.IsBusy = True
            DoEvents
        Loop

        Screen.MousePointer = vbDefault

        .show vbModal
    End With

EXT:

    RsRep.Close
    Set RsRep = Nothing
    Set ReportShow = Nothing
    Set ReportApp = Nothing
ErrTrap:
    Screen.MousePointer = vbDefault
End Function

