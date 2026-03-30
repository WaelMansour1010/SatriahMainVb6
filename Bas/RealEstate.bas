Attribute VB_Name = "RealEstate"

Function print_report1(Optional NoteSerial As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From Expanses_Order  where noteserial='" & NoteSerial & "'"
 
    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Real Etstae\" & "Order_form1.rpt"
    Else
        StrFileName = App.path & "\Reports\Real Etstae\" & "Order_form1.rpt"
    End If

    If Dir(StrFileName) = "" Then

        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If RsData.BOF Or RsData.EOF Then

    '    Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    RsData.Close
    '    Set RsData = Nothing
    '    Screen.MousePointer = vbDefault
    '    Exit Function
    'End If
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    'xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    'If SystemOptions.UserInterface = ArabicInterface Then
    '    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    '
    '    StrReportTitle = "" '& StrAccountName
    ' Else
    '
    '    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    '
    '     xReport.ParameterFields(4).AddCurrentValue get_branch_name(Val(my_branch))
    '    StrReportTitle = ""
    ' End If
    'xReport.ParameterFields(3).AddCurrentValue user_name
    'xReport.ReportTitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    'xReport.ApplicationName = App.Title
    'xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Function print_report(Optional NoteSerial As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From Expanses_Order  where noteserial='" & NoteSerial & "'"
 
    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Real Etstae\" & "Contract_form.rpt"
    Else
        StrFileName = App.path & "\Reports\Real Etstae\" & "Contract_form.rpt"
    End If

    If Dir(StrFileName) = "" Then

        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If RsData.BOF Or RsData.EOF Then

    '    Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    RsData.Close
    '    Set RsData = Nothing
    '    Screen.MousePointer = vbDefault
    '    Exit Function
    'End If
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    'xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    'If SystemOptions.UserInterface = ArabicInterface Then
    '    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    '
    '    StrReportTitle = "" '& StrAccountName
    ' Else
    '
    '    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    '
    '     xReport.ParameterFields(4).AddCurrentValue get_branch_name(Val(my_branch))
    '    StrReportTitle = ""
    ' End If
    'xReport.ParameterFields(3).AddCurrentValue user_name
    'xReport.ReportTitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    'xReport.ApplicationName = App.Title
    'xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

