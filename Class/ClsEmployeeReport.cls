VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEmployeeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function EmpData(EmpID As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrTrap

    If EmpID = Null Then
        Exit Function
    End If

    StrSQL = "select * From View8 where Emp_ID=" & EmpID
    'StrSQL = "SELECT dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,dbo.TblEmployee.Emp_Mail, dbo.TblEmpSpecifications.SpecificationName, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile,dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.NumEkama, dbo.TblEmployee.DateExpoekama,dbo.TblEmployee.DateEndekama, dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc,dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.NumPasp,dbo.TblEmployee.DateExpPasp , dbo.TblEmployee.DateEndPasp, dbo.TblEmployee.BignDateWork FROM dbo.TblEmpDepartments INNER JOIN dbo.TblEmployee ON dbo.TblEmpDepartments.DeparmentID = dbo.TblEmployee.DepartmentID INNER JOIN dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID INNER JOIN dbo.TblEmpSpecifications ON dbo.TblEmployee.SpecificationID = dbo.TblEmpSpecifications.SpecificationID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If Dir(App.path & "\Reports\EmplReport.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = "\Reports\EmplReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\EmplReport.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & Chr(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Function EmpPyment(EmpID As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrTrap

    If EmpID = Null Then
        Exit Function
    End If

    StrSQL = "select * From QryPyment where AdvanceID=" & EmpID
    'StrSQL = "SELECT dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,dbo.TblEmployee.Emp_Mail, dbo.TblEmpSpecifications.SpecificationName, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile,dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.NumEkama, dbo.TblEmployee.DateExpoekama,dbo.TblEmployee.DateEndekama, dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc, dbo.TblEmployee.DateEndLinc,dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket, dbo.TblEmployee.dateendpoket, dbo.TblEmployee.NumPasp,dbo.TblEmployee.DateExpPasp , dbo.TblEmployee.DateEndPasp, dbo.TblEmployee.BignDateWork FROM dbo.TblEmpDepartments INNER JOIN dbo.TblEmployee ON dbo.TblEmpDepartments.DeparmentID = dbo.TblEmployee.DepartmentID INNER JOIN dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID INNER JOIN dbo.TblEmpSpecifications ON dbo.TblEmployee.SpecificationID = dbo.TblEmpSpecifications.SpecificationID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If Dir(App.path & "\Reports\EmpPyment.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = "\Reports\EmplReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\EmpPyment.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & Chr(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

