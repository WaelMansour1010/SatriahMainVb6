VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMaintananceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function MaintenanceData(MaintanenceID As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrTrap

    If MaintanenceID = Null Then
        Exit Function
    End If

    StrSQL = "select * From QryMaintananceReport where MaintananceID=" & MaintanenceID
    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If Dir(App.path & "\Reports\MaintananceReport.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = "\Reports\MaintananceReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\MaintananceReport.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        StrSQL = "select * From Notes where MaintananceID=" & MaintanenceID
        StrSQL = StrSQL + " and NoteType=0"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
        End If

        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.Show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & Chr(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Function MaintenanceDataShort(MaintanenceID As Integer, _
                                     Optional IntType As Integer = 0)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim StrFileName As String

    On Error GoTo ErrTrap

    If MaintanenceID = Null Then
        Exit Function
    End If

    If IntType = 0 Then
        StrFileName = App.path & "\Reports\MaintananceReportShort.rpt"
    Else
        StrFileName = App.path & "\Reports\MaintananceReportShortRecive.rpt"
    End If

    StrSQL = "select * From QryMaintananceReport where MaintananceID=" & MaintanenceID
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = Msg & Chr(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = Msg & Chr(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        StrSQL = "select * From Notes where MaintananceID=" & MaintanenceID
        StrSQL = StrSQL + " and NoteType=0"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
        End If

        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.Show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & Chr(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

