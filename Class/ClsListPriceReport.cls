VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsListPriceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function ShowListPrice(Optional IntStyle As Integer = 0)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim cOptions As ClsCompanyInfo

    On Error GoTo ErrTrap

    'IntStyle=0 --- Normal Style
    'IntStyle=1 --- Tree Report Style
    If IntStyle = 0 Then
        StrFileName = App.path & "\Reports\ListPrice.rpt"
        StrSQL = "Select * From QryPriceList Order By ItemName"
    ElseIf IntStyle = 1 Then
        StrFileName = App.path & "\Reports\ListPriceTreee.rpt"
        StrSQL = "Select * From QryPriceListTree"
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = "\Reports\ListPrice.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        xReport.reporttitle = "ﬁ«∆„… √”⁄«— «·√’‰«›"
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & Chr(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

