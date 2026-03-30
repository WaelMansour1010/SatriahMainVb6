VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBoxesReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
clsName ClsBoxesReports
updated function


Public Function BoxIncapcity(LngBoxID As Long, _
                             FromDate As Variant, _
                             toDate As Variant, _
                             Optional m_Traget As PrintTarget = WindowTarget, Optional ByVal mNoteType As Integer = -1, Optional ByVal mReportName As String = "")
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim FrmReport As FrmReportViewer
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim cReportViewer As ClsReportViewer
    Dim StrFileName As String
 
    On Error GoTo ErrTrap
    If mReportName <> "" Then
        StrFileName = App.Path & "\Reports\REPORTS NEW\" & Trim(mReportName)
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.Path & "\Reports\BoxIncapcityReport.rpt"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            StrFileName = App.Path & "\Reports\BoxIncapcityReport_Eng.rpt"
        End If
    End If
    
    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    StrSQL = "Select * From QryBoxReports Where (NoteType=11 Or NoteType=12) "
    If LngBoxID <> -1 Then
        StrSQL = StrSQL & " and BoxID = " & LngBoxID
    End If
    
    
    If Not IsNull(FromDate) Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = StrSQL + " And NoteDate >=#" & SQLDate(CDate(FromDate)) & "#"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = StrSQL + " And NoteDate >='" & SQLDate(CDate(FromDate)) & "'"
        End If
    End If

    If Not IsNull(toDate) Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = StrSQL + " And NoteDate <=#" & SQLDate(CDate(toDate)) & "#"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = StrSQL + " And NoteDate <='" & SQLDate(CDate(toDate)) & "'"
        End If
    End If

    If mNoteType <> -1 Then
        StrSQL = StrSQL + " AND NoteType =" & mNoteType
    End If
    
    StrSQL = StrSQL + " Order By NoteID "
    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.reporttitle = " ﬁ—Ì— ⁄Ã“ «·‰ﬁœÌ… Ê“Ì«œ… «·‰ﬁœÌ… ›Ï «·Œ“‰…"
    xReport.Database.SetDataSource RsData
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").Value), "", StrRS("Company_Arabic_Name").Value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").Value), "", StrRS("Company_Comment").Value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), FromDate)
    xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(toDate), CDate("01/01/1000"), toDate)
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_Traget, " ﬁ—Ì— ⁄Ã“ «·‰ﬁœÌ… Ê“Ì«œ… «·‰ﬁœÌ… ›Ï «·Œ“‰…"
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Screen.MousePointer = vbDefault
    'Resume
End Function


ClassName ClsRepoerts
Updated Function


Public Function ItemsReport(StrSQL As String, Optional ByVal mReportName As String = "")
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Msg As String

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If
Dim filename As String

If mReportName = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    filename = App.Path & "\Reports\ReportItemsData.rpt"
     Else
     filename = App.Path & "\Reports\ReportItemsDataE.rpt"
     End If
Else
    filename = App.Path & "\Reports\REPORTS NEW\" & Trim(mReportName)
End If
    If Dir(filename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = Msg + filename
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(filename)
        xReport.Database.SetDataSource rs
        xReport.reporttitle = "»Ì«‰«  «·√’‰«›"
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , filename
    Screen.MousePointer = vbDefault
    
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & Chr(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

