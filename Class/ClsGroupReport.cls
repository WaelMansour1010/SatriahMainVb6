VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGroupReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function GroupData(GroupID As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrTrap

    If GroupID = Null Then
        Exit Function
    End If

    StrSQL = "select * From Groups where GroupID=" & GroupID
    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If Dir(App.path & "\Reports\GroupsReport.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = "\Reports\GroupsReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\GroupsReport.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        
        If Not (IsNull(rs("ParentID").value)) Then
            StrSQL = "select * From Groups where GroupID=" & rs("ParentID").value
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("GroupName").value), "", Trim(RsTemp("GroupName").value))
            Else
                xReport.ParameterFields(4).AddCurrentValue ""
            End If

        Else
            xReport.ParameterFields(4).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

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

