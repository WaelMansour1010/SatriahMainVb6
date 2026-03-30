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
Option Explicit

Public Function BoxBalance(LngBoxID As Long, _
                           FromDate As Variant, _
                           toDate As Variant, _
                           Optional BolShort As Boolean = False)
    
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsBegine As ADODB.Recordset
    Dim RsEndQty As ADODB.Recordset
    Dim DblBegineAccount As Double
    Dim DblEndAccount As Double
    Dim RsData As ADODB.Recordset
    Dim FrmReport As FrmReportViewer
    Dim StrRS As ADODB.Recordset
    Dim cReportViewer As ClsReportViewer
    Dim StrBoxStockData As String
    Dim Msg As String
    Dim cProgress As ClsProgress

    Dim StrFileName As String
    On Error GoTo ErrTrap

    If BolShort = False Then
        StrFileName = App.path & "\Reports\BoxBalance.rpt"
    Else
        StrFileName = App.path & "\Reports\BoxBalanceShort.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
    cProgress.StartProgress

    If Not IsNull(FromDate) Then
    
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT QryBoxBalance.BoxID, QryBoxBalance.BoxName," & "Sum(QryBoxBalance.Note_Value  * TransDir) AS BoxAccount " & "From QryBoxBalance "
            StrSQL = StrSQL + " Where (QryBoxBalance.BoxID=" & LngBoxID & ") "
            StrSQL = StrSQL + " AND QryBoxBalance.NoteDate < #" & SQLDate(CDate(FromDate), False) & "#"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT QryBoxBalance.BoxID, QryBoxBalance.BoxName," & "Sum(QryBoxBalance.Note_Value  * TransDir) AS BoxAccount " & "From dbo.QryBoxBalance()QryBoxBalance "
            StrSQL = StrSQL + " Where (QryBoxBalance.BoxID=" & LngBoxID & ") "
            StrSQL = StrSQL + " AND QryBoxBalance.NoteDate < '" & SQLDate(CDate(FromDate), False) & "'"
        End If

        StrSQL = StrSQL + " GROUP BY QryBoxBalance.BoxID, QryBoxBalance.BoxName;"
        Set RsBegine = New ADODB.Recordset
        RsBegine.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute

        Do While RsBegine.State = adStateExecuting
            DoEvents
        Loop

        If Not (RsBegine.BOF Or RsBegine.EOF) Then
            DblBegineAccount = IIf(IsNull(RsBegine("BoxAccount").value), "", RsBegine("BoxAccount").value)
        End If

    Else
        DblBegineAccount = 0
    End If

    Set RsEndQty = New ADODB.Recordset

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT QryBoxBalance.BoxID, QryBoxBalance.BoxName," & "Sum(QryBoxBalance.Note_Value  * TransDir) AS BoxAccount " & "From QryBoxBalance "
        StrSQL = StrSQL + " Where (QryBoxBalance.BoxID=" & LngBoxID & ") "
        StrSQL = StrSQL + " GROUP BY QryBoxBalance.BoxID, QryBoxBalance.BoxName;"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT QryBoxBalance.BoxID, QryBoxBalance.BoxName," & "Sum(QryBoxBalance.Note_Value  * TransDir) AS BoxAccount " & "From dbo.QryBoxBalance()QryBoxBalance "
        StrSQL = StrSQL + " Where (QryBoxBalance.BoxID=" & LngBoxID & ") "
        StrSQL = StrSQL + " GROUP BY QryBoxBalance.BoxID, QryBoxBalance.BoxName;"
    End If

    RsEndQty.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute

    Do While RsEndQty.State = adStateExecuting
        DoEvents
    Loop

    If Not (RsEndQty.BOF Or RsEndQty.EOF) Then
        DblEndAccount = RsEndQty("BoxAccount").value
    End If

    If BolShort = False Then
        GetBoxStockData LngBoxID, FromDate, toDate

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT QryBoxBalance.*, TempBoxStockTable.Comment,TempBoxStockTable.StockSum " & "FROM QryBoxBalance LEFT JOIN TempBoxStockTable ON " & "(QryBoxBalance.NoteDate = TempBoxStockTable.BoxStockDate) " & " AND (QryBoxBalance.BoxID = TempBoxStockTable.BoxID)"
            StrSQL = StrSQL + " Where QryBoxBalance.BoxID=" & LngBoxID & ""
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "Select QryBoxBalance.*, dbo.TempBoxStockTable.Comment,dbo.TempBoxStockTable.StockSum " & "From dbo.QryBoxBalance()QryBoxBalance LEFT OUTER JOIN " & "dbo.TempBoxStockTable ON QryBoxBalance.NoteDate = " & "dbo.TempBoxStockTable.BoxStockDate " & "AND QryBoxBalance.BoxID = dbo.TempBoxStockTable.BoxID"
            StrSQL = StrSQL + " Where QryBoxBalance.BoxID=" & LngBoxID & ""
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

        StrSQL = StrSQL + " Order By NoteDate,NoteID"
    
    Else
        StrSQL = "SELECT QryBoxBalance.BoxID, QryBoxBalance.BoxName," & "QryBoxBalance.TransactionTypeName, QryBoxBalance.TransDir," & "Sum(QryBoxBalance.Note_Value) AS SumNote_Value,Count(QryBoxBalance.TransactionTypeName) as TransCount "

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = StrSQL + " From QryBoxBalance"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = StrSQL + " From dbo.QryBoxBalance()QryBoxBalance"
        End If

        StrSQL = StrSQL + " Where QryBoxBalance.BoxID=" & LngBoxID & ""

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

        StrSQL = StrSQL + " GROUP BY QryBoxBalance.BoxID, QryBoxBalance.BoxName," & "QryBoxBalance.TransactionTypeName, QryBoxBalance.TransDir "
        StrSQL = StrSQL + " Order by QryBoxBalance.TransDir DESC "
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute

    Do While RsData.State = adStateExecuting
        DoEvents
    Loop

    If RsData.BOF Or RsData.EOF Then
        If Not cProgress Is Nothing Then
            cProgress.StopProgess
            Set cProgress = Nothing
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    If Not (IsNull(toDate) And IsNull(FromDate)) Then
        If toDate = FromDate Then
            'StrBoxStockData = GetBoxStockData(LngBoxID, CDate(FromDate))
        End If
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue DblBegineAccount
    xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), FromDate)
    xReport.ParameterFields(6).AddCurrentValue IIf(IsNull(toDate), CDate("01/01/1000"), toDate)
    xReport.ParameterFields(7).AddCurrentValue DblEndAccount
    'xReport.ParameterFields(8).AddCurrentValue StrBoxStockData

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, WindowTarget, "ﬂ‘› Õ”«» «·Œ“‰…"

    If Not cProgress Is Nothing Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If

    Set cReportViewer = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Screen.MousePointer = vbDefault

    If Not cProgress Is Nothing Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If

End Function

Public Function BoxDrawingDeposit(LngBoxID As Long, _
                                  FromDate As Variant, _
                                  toDate As Variant, _
                                  Optional m_Traget As PrintTarget = WindowTarget)
    
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

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\BoxDrawingReport.rpt"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrFileName = App.path & "\Reports\BoxDrawingReport_Eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    StrSQL = "Select * From QryBoxReports Where BoxID = " & LngBoxID & " AND(NoteType=7 Or NoteType=8)"
    
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

    StrSQL = StrSQL + " Order By NoteID "
    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.reporttitle = " ﬁ—Ì— «·”Õ» Ê«·≈Ìœ«⁄ ›Ï «·Œ“‰…"
    xReport.Database.SetDataSource RsData
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), FromDate)
    xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(toDate), CDate("01/01/1000"), toDate)
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_Traget, " ﬁ—Ì— «·”Õ» Ê«·≈Ìœ«⁄ ›Ï «·Œ“‰…"
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Screen.MousePointer = vbDefault
    'Resume
End Function

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
        StrFileName = App.path & "\Reports\REPORTS NEW\" & Trim(mReportName)
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\BoxIncapcityReport.rpt"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            StrFileName = App.path & "\Reports\BoxIncapcityReport_Eng.rpt"
        End If
    End If
    
    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.reporttitle = " ﬁ—Ì— ⁄Ã“ «·‰ﬁœÌ… Ê“Ì«œ… «·‰ﬁœÌ… ›Ï «·Œ“‰…"
    xReport.Database.SetDataSource RsData
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
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
Public Function BoxStock(LngBoxID As Long, _
                         FromDate As Variant, _
                         toDate As Variant, _
                         Optional m_Traget As PrintTarget = WindowTarget)
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

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\BoxStockReport.rpt"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrFileName = App.path & "\Reports\BoxStockReport_Eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    StrSQL = "Select * From QryBoxStockReport Where BoxID =" & LngBoxID & ""
    
    If Not IsNull(FromDate) Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = StrSQL + " And BoxStockDate >=#" & SQLDate(CDate(FromDate)) & "#"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = StrSQL + " And BoxStockDate >='" & SQLDate(CDate(FromDate)) & "'"
        End If
    End If

    If Not IsNull(toDate) Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = StrSQL + " And BoxStockDate <=#" & SQLDate(CDate(toDate)) & "#"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = StrSQL + " And BoxStockDate <='" & SQLDate(CDate(toDate)) & "'"
        End If
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.reporttitle = " ﬁ—Ì— »⁄„·Ì«  Ã—œ «·Œ“‰…"
    xReport.Database.SetDataSource RsData
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), FromDate)
    xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(toDate), CDate("01/01/1000"), toDate)

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_Traget, xReport.reporttitle
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Screen.MousePointer = vbDefault
    'Resume
End Function

Private Function GetBoxStockData(LngBoxID As Long, _
                                 FromDate As Variant, _
                                 toDate As Variant) As String
        
    Dim rs As ADODB.Recordset
    Dim StrTemp As String
    Dim StrTempDits As String
    Dim i As Integer
    Dim StrSQL As String
    Dim OldID As Long
    Dim M_Date As Date
    Dim RsNew As ADODB.Recordset
    Dim DblTotal As Single

    StrSQL = "Delete From TempBoxStockTable"
    Cn.Execute StrSQL
    StrSQL = "Select * From QryBoxStockReport Where BoxID =" & LngBoxID & ""

    If Not (IsNull(FromDate)) Then
        StrSQL = StrSQL + " And BoxStockDate >=" & SQLDate(CDate(FromDate), True) & ""
    End If

    If Not (IsNull(toDate)) Then
        StrSQL = StrSQL + " And BoxStockDate <=" & SQLDate(CDate(toDate), True) & ""
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Set RsNew = New ADODB.Recordset
        RsNew.Open "TempBoxStockTable", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        StrTemp = ""
        StrTempDits = ""
        OldID = 0
        rs.MoveFirst

        'OldID = Rs("BoxStockID").Value
        For i = 0 To rs.RecordCount - 1

            If OldID <> rs("BoxStockID").value Then
                OldID = rs("BoxStockID").value

                If StrTemp <> "" And StrTempDits <> "" Then
                    StrTemp = StrTemp & vbCrLf & StrTempDits & vbCrLf
                    StrTemp = StrTemp & "-------------------------" & vbCrLf
                    StrTemp = StrTemp & "«·≈Ã„«·Ï:-" & DblTotal
                
                    RsNew.AddNew
                    RsNew("BoxID").value = LngBoxID
                    RsNew("BoxStockDate").value = M_Date
                    RsNew("Comment").value = StrTemp
                    RsNew("StockSum").value = DblTotal
                    DblTotal = 0
                    RsNew.update
                End If

                'Read Header
                StrTemp = "»Ì«‰«  ⁄„·Ì… «·Ã—œ ›Ï ÌÊ„ " & DisplayDate(rs("BoxStockDate").value) & vbCrLf
                M_Date = rs("BoxStockDate").value
                StrTemp = StrTemp & "-------------------------" & vbCrLf
                StrTemp = StrTemp & "—ﬁ„ ⁄„·Ì… «·Ã—œ:" & rs("BoxStockID").value & vbCrLf
                StrTemp = StrTemp & "`" & "«”„ «·„” Œœ„ «·„Õ——:" & "`" & rs("UserName").value & "`" & vbCrLf
                StrTemp = StrTemp & "---- ›«’Ì· ⁄„·Ì… «·Ã—œ----" & vbCrLf
                StrTemp = StrTemp & String(5, vbKeySpace) & "«·›∆…" & String(30, vbKeySpace)
                StrTemp = StrTemp & "«·⁄œœ" & String(10, vbKeySpace)
                StrTemp = StrTemp & "«·ﬁÌ„…" & String(5, vbKeySpace)
                StrTempDits = ""
            End If

            StrTempDits = StrTempDits & "" & rs("CurrencyBandName").value & String(30 - Len(rs("CurrencyBandName").value), vbKeySpace) & rs("CurrencyBandCount").value
            StrTempDits = StrTempDits & String(10, " ") & (rs("CurrencyBandCount").value * rs("CurrenyBandValue").value)
            DblTotal = DblTotal + (rs("CurrencyBandCount").value * rs("CurrenyBandValue").value)
            StrTempDits = StrTempDits & vbCrLf
            rs.MoveNext
        Next i

        If StrTemp <> "" And StrTempDits <> "" Then
            StrTemp = StrTemp & vbCrLf & StrTempDits & vbCrLf
            StrTemp = StrTemp & "-------------------------" & vbCrLf
            StrTemp = StrTemp & "«·≈Ã„«·Ï:- " & DblTotal
            RsNew.AddNew
            RsNew("BoxID").value = LngBoxID
            RsNew("BoxStockDate").value = M_Date
            RsNew("Comment").value = StrTemp
            RsNew("StockSum").value = DblTotal
            DblTotal = 0
            RsNew.update
        End If

    Else
        StrTemp = ""
    End If

    rs.Close
    Set rs = Nothing
    GetBoxStockData = StrTemp
End Function

Public Sub BoxDayReport(LngBoxID As Long, _
                        FromDate As Variant, _
                        toDate As Variant, _
                        Optional BolShort As Boolean = False)
    
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsBegine As ADODB.Recordset
    Dim RsEndQty As ADODB.Recordset
    Dim DblBegineAccount As Double
    Dim DblEndAccount As Double
    Dim RsData As ADODB.Recordset
    Dim StrRS As ADODB.Recordset
    Dim cReportViewer As ClsReportViewer
    Dim StrBoxStockData As String
    Dim Msg As String
    Dim StrFileName As String
    Dim i As Integer
    Dim StrReportTitle As String
    Dim RsQryPar As ADODB.Recordset

    Dim SngPar1 As Double, SngPar2 As Double, SngPar3 As Double, SngPar4 As Double, SngPar5 As Double, SngPar6 As Double
    Dim SngPar7 As Double, SngPar8 As Double, SngPar9 As Double, SngPar10 As Double, SngPar11 As Double, SngPar12 As Double
    Dim SngPar13 As Double, SngPar14 As Double, SngPar15 As Double, SngPar16 As Double, SngPar17 As Double, SngPar18 As Double
    Dim SngPar19 As Double, SngPar20 As Double, SngPar21 As Double, SngPar22 As Double, SngPar23 As Double, SngPar24 As Double
    Dim SngPar25 As Double

    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\BoxDayReport.rpt"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrFileName = App.path & "\Reports\BoxDayReport_Eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If Not IsNull(FromDate) Then
        DblBegineAccount = GetBoxBegainAccount(LngBoxID, CDate(FromDate))
    Else
        DblBegineAccount = 0
    End If

    Set RsEndQty = New ADODB.Recordset

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT QryBoxBalance.BoxID, QryBoxBalance.BoxName," & "Sum(QryBoxBalance.Note_Value  * TransDir) AS BoxAccount " & "From QryBoxBalance "
        StrSQL = StrSQL + " Where (QryBoxBalance.BoxID=" & LngBoxID & ") "
        StrSQL = StrSQL + " GROUP BY QryBoxBalance.BoxID, QryBoxBalance.BoxName;"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT QryBoxBalance.BoxID, QryBoxBalance.BoxName," & "Sum(QryBoxBalance.Note_Value  * TransDir) AS BoxAccount " & "From dbo.QryBoxBalance()QryBoxBalance "
        StrSQL = StrSQL + " Where (QryBoxBalance.BoxID=" & LngBoxID & ") "
        StrSQL = StrSQL + " GROUP BY QryBoxBalance.BoxID, QryBoxBalance.BoxName;"
    End If

    RsEndQty.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEndQty.BOF Or RsEndQty.EOF) Then
        DblEndAccount = RsEndQty("BoxAccount").value
    End If

    '---------------SngPar1---------------------------
    'Õ”«» ﬁÌ„… «·„»Ì⁄«  «·‰ﬁœÌ…
    Set RsQryPar = New ADODB.Recordset

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT SUM(QryTransactionsTotal.TotalAfterTax) AS ParValue"
        StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal INNER JOIN"
        StrSQL = StrSQL + " dbo.Notes ON QryTransactionsTotal.Transaction_ID = dbo.Notes.Transaction_ID"
        StrSQL = StrSQL + " Where (dbo.Notes.BoxID=" & LngBoxID & ""
        StrSQL = StrSQL + " AND QryTransactionsTotal.PaymentType=0 AND Transaction_Type=2)"

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date >=" & SQLDate(CDate(FromDate), True)
        End If

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date <=" & SQLDate(CDate(toDate), True)
        End If

    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
    
    End If

    RsQryPar.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsQryPar.BOF Or RsQryPar.EOF) Then
        SngPar1 = IIf(IsNull(RsQryPar("ParValue").value), 0, RsQryPar("ParValue").value)
    Else
        SngPar1 = 0
    End If

    '----------------SngPar2---------------------------------
    '„‰  Õ  Õ”«» „»Ì⁄«  «Ã·…

    StrSQL = " SELECT     SUM(Notes.Note_Value) AS ParValue"
    StrSQL = StrSQL + " FROM Transactions INNER JOIN"
    StrSQL = StrSQL + " Notes ON Transactions.Transaction_ID = Notes.Transaction_ID"
    StrSQL = StrSQL + " Where (Notes.BoxID=" & LngBoxID & ") AND (Notes.NoteType = 0) " & "And (Transactions.PaymentType = 1)" & "And (Transactions.Transaction_Type = 2)"

    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " AND Transactions.Transaction_Date >=" & SQLDate(CDate(FromDate), True)
    End If

    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " AND Transactions.Transaction_Date <=" & SQLDate(CDate(toDate), True)
    End If

    Set RsQryPar = New ADODB.Recordset
    RsQryPar.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsQryPar.BOF Or RsQryPar.EOF) Then
        SngPar2 = IIf(IsNull(RsQryPar("ParValue").value), 0, RsQryPar("ParValue").value)
    Else
        SngPar2 = 0
    End If

    '------------SngPar3------------------------------
    'Õ”«» ﬁÌ„… „—œÊœ«  „‘ —Ì«   «·‰ﬁœÌ…

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT SUM(QryTransactionsTotal.TotalAfterTax) AS ParValue"
        StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal INNER JOIN"
        StrSQL = StrSQL + " dbo.Notes ON QryTransactionsTotal.Transaction_ID = dbo.Notes.Transaction_ID"
        StrSQL = StrSQL + " Where (dbo.Notes.BoxID=" & LngBoxID & ""
        StrSQL = StrSQL + " AND QryTransactionsTotal.PaymentType=0 AND Transaction_Type=5)"

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date >=" & SQLDate(CDate(FromDate), True)
        End If

        If Not IsNull(FromDate) Then
            '        StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date <=" & SQLDate(CDate(ToDate), True)
        End If

    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
    
    End If

    Set RsQryPar = New ADODB.Recordset
    RsQryPar.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsQryPar.BOF Or RsQryPar.EOF) Then
        SngPar3 = IIf(IsNull(RsQryPar("ParValue").value), 0, RsQryPar("ParValue").value)
    Else
        SngPar3 = 0
    End If

    '------------SngPar5------------------------------
    '«·“Ì«œ… ›Ï «·‰ﬁœÌ…
    StrSQL = "Select sum(Note_Value) as ParValue From QryBoxReports Where BoxID = " & LngBoxID & " AND(NoteType=11)"
    
    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " And NoteDate >=" & SQLDate(CDate(FromDate), True) & ""
    End If

    If Not IsNull(toDate) Then
        StrSQL = StrSQL + " And NoteDate <=" & SQLDate(CDate(toDate), True) & ""
    End If

    Set RsQryPar = New ADODB.Recordset
    RsQryPar.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsQryPar.BOF Or RsQryPar.EOF) Then
        SngPar5 = IIf(IsNull(RsQryPar("ParValue").value), 0, RsQryPar("ParValue").value)
    Else
        SngPar5 = 0
    End If

    '---------------SngPar8--------------------
    '«·„’—Ê›« 
    Set RsQryPar = New ADODB.Recordset
    StrSQL = "SELECT SUM(Note_Value) AS ParValue FROM ExpensesReport Where BoxID=" & LngBoxID & ""

    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(CDate(FromDate), True)
    End If

    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(CDate(toDate), True)
    End If

    RsQryPar.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsQryPar.BOF Or RsQryPar.EOF) Then
        SngPar8 = IIf(IsNull(RsQryPar("ParValue").value), 0, RsQryPar("ParValue").value)
    Else
        SngPar8 = 0
    End If

    '---------------SngPar9---------------------------
    'Õ”«» ﬁÌ„… «·„‘ —Ì«  «·‰ﬁœÌ…
    Set RsQryPar = New ADODB.Recordset

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT SUM(QryTransactionsTotal.TotalAfterTax) AS ParValue"
        StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal INNER JOIN"
        StrSQL = StrSQL + " dbo.Notes ON QryTransactionsTotal.Transaction_ID = dbo.Notes.Transaction_ID"
        StrSQL = StrSQL + " Where (dbo.Notes.BoxID=" & LngBoxID & ""
        StrSQL = StrSQL + " AND QryTransactionsTotal.PaymentType=0 AND Transaction_Type=1)"

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date >=" & SQLDate(CDate(FromDate), True)
        End If

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date <=" & SQLDate(CDate(toDate), True)
        End If

    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
    
    End If

    RsQryPar.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsQryPar.BOF Or RsQryPar.EOF) Then
        SngPar9 = IIf(IsNull(RsQryPar("ParValue").value), 0, RsQryPar("ParValue").value)
    Else
        SngPar9 = 0
    End If

    '----------------SngPar10---------------------------------
    '„‰  Õ  Õ”«» „‘ —Ì«  «Ã·…
    'œ›⁄«  „ﬁœ„…
    StrSQL = " SELECT     SUM(Notes.Note_Value) AS ParValue"
    StrSQL = StrSQL + " FROM Transactions INNER JOIN"
    StrSQL = StrSQL + " Notes ON Transactions.Transaction_ID = Notes.Transaction_ID"
    StrSQL = StrSQL + " Where (Notes.BoxID=" & LngBoxID & ") AND (Notes.NoteType = 0) " & "And (Transactions.PaymentType = 1)" & "And (Transactions.Transaction_Type = 1)"

    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " AND Transactions.Transaction_Date >=" & SQLDate(CDate(FromDate), True)
    End If

    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " AND Transactions.Transaction_Date <=" & SQLDate(CDate(toDate), True)
    End If

    Set RsQryPar = New ADODB.Recordset
    RsQryPar.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsQryPar.BOF Or RsQryPar.EOF) Then
        SngPar10 = IIf(IsNull(RsQryPar("ParValue").value), 0, RsQryPar("ParValue").value)
    Else
        SngPar10 = 0
    End If

    '------------SngPar11------------------------------
    'Õ”«» ﬁÌ„… „—œÊœ«  „»Ì⁄«   «·‰ﬁœÌ…
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT SUM(QryTransactionsTotal.TotalAfterTax) AS ParValue"
        StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal INNER JOIN"
        StrSQL = StrSQL + " dbo.Notes ON QryTransactionsTotal.Transaction_ID = dbo.Notes.Transaction_ID"
        StrSQL = StrSQL + " Where (dbo.Notes.BoxID=" & LngBoxID & ""
        StrSQL = StrSQL + " AND QryTransactionsTotal.PaymentType=0 AND Transaction_Type=9)"

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date >=" & SQLDate(CDate(FromDate), True)
        End If

        If Not IsNull(FromDate) Then
            '        StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date <=" & SQLDate(CDate(ToDate), True)
        End If

    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
    
    End If

    Set RsQryPar = New ADODB.Recordset
    RsQryPar.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsQryPar.BOF Or RsQryPar.EOF) Then
        SngPar11 = IIf(IsNull(RsQryPar("ParValue").value), 0, RsQryPar("ParValue").value)
    Else
        SngPar11 = 0
    End If

    '------------SngPar13------------------------------
    '·⁄Ã“ ›Ï «·‰ﬁœÌ…
    StrSQL = "Select sum(Note_Value) as ParValue From QryBoxReports Where BoxID = " & LngBoxID & " AND(NoteType=12)"
    
    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " And NoteDate >=" & SQLDate(CDate(FromDate), True) & ""
    End If

    If Not IsNull(toDate) Then
        StrSQL = StrSQL + " And NoteDate <=" & SQLDate(CDate(toDate), True) & ""
    End If

    Set RsQryPar = New ADODB.Recordset
    RsQryPar.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsQryPar.BOF Or RsQryPar.EOF) Then
        SngPar13 = IIf(IsNull(RsQryPar("ParValue").value), 0, RsQryPar("ParValue").value)
    Else
        SngPar13 = 0
    End If

    '---------------------------------

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    'xReport.Database.SetDataSource RsData
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue DblBegineAccount

    xReport.ParameterFields(5).AddCurrentValue CDbl(SngPar1)
    xReport.ParameterFields(6).AddCurrentValue SngPar2
    xReport.ParameterFields(7).AddCurrentValue SngPar3
    xReport.ParameterFields(8).AddCurrentValue SngPar4
    xReport.ParameterFields(9).AddCurrentValue SngPar5
    xReport.ParameterFields(10).AddCurrentValue SngPar6
    xReport.ParameterFields(11).AddCurrentValue SngPar7
    xReport.ParameterFields(12).AddCurrentValue SngPar8
    xReport.ParameterFields(13).AddCurrentValue SngPar9
    xReport.ParameterFields(14).AddCurrentValue SngPar10
    xReport.ParameterFields(15).AddCurrentValue SngPar11
    xReport.ParameterFields(16).AddCurrentValue SngPar12
    xReport.ParameterFields(17).AddCurrentValue SngPar13
    xReport.ParameterFields(18).AddCurrentValue SngPar14
    xReport.ParameterFields(19).AddCurrentValue SngPar15
    xReport.ParameterFields(20).AddCurrentValue SngPar16
    xReport.ParameterFields(21).AddCurrentValue SngPar17
    xReport.ParameterFields(22).AddCurrentValue SngPar18
    xReport.ParameterFields(23).AddCurrentValue SngPar19
    xReport.ParameterFields(24).AddCurrentValue SngPar20
    xReport.ParameterFields(25).AddCurrentValue SngPar21
    xReport.ParameterFields(26).AddCurrentValue SngPar22
    xReport.ParameterFields(27).AddCurrentValue SngPar23
    xReport.ParameterFields(28).AddCurrentValue SngPar24
    xReport.ParameterFields(29).AddCurrentValue SngPar25
    StrReportTitle = " ﬁ—Ì— »ÌÊ„Ì… «·Œ“‰… "

    If Not (IsNull(FromDate)) Then
        StrReportTitle = StrReportTitle & "„‰ " & DisplayDate(CDate(FromDate))
    End If

    If Not (IsNull(toDate)) Then
        StrReportTitle = StrReportTitle & "≈·Ï " & DisplayDate(CDate(toDate))
    End If

    xReport.reporttitle = StrReportTitle

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, WindowTarget, " ﬁ—— ÌÊ„Ì… «·Œ“‰…"
    Set cReportViewer = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
    'Resume
End Sub

Private Function GetBoxBegainAccount(LngBoxID As Long, _
                                     FromDate As Date) As Single
    Dim StrSQL As String
    Dim RsBegine As ADODB.Recordset
    Dim DblBegineAccount As Double
    On Error GoTo ErrTrap

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT QryBoxBalance.BoxID, QryBoxBalance.BoxName," & "Sum(QryBoxBalance.Note_Value  * TransDir) AS BoxAccount " & "From QryBoxBalance "
        StrSQL = StrSQL + " Where (QryBoxBalance.BoxID=" & LngBoxID & ") "
        StrSQL = StrSQL + " AND QryBoxBalance.NoteDate < #" & SQLDate(FromDate, False) & "#"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT QryBoxBalance.BoxID, QryBoxBalance.BoxName," & "Sum(QryBoxBalance.Note_Value  * TransDir) AS BoxAccount " & "From dbo.QryBoxBalance()QryBoxBalance "
        StrSQL = StrSQL + " Where (QryBoxBalance.BoxID=" & LngBoxID & ") "
        StrSQL = StrSQL + " AND QryBoxBalance.NoteDate < '" & SQLDate(FromDate, False) & "'"
    End If

    StrSQL = StrSQL + " GROUP BY QryBoxBalance.BoxID, QryBoxBalance.BoxName;"
    Set RsBegine = New ADODB.Recordset
    RsBegine.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsBegine.BOF Or RsBegine.EOF) Then
        DblBegineAccount = IIf(IsNull(RsBegine("BoxAccount").value), "", RsBegine("BoxAccount").value)
    End If

    GetBoxBegainAccount = DblBegineAccount
    Exit Function
ErrTrap:
    GetBoxBegainAccount = 0
End Function

Public Function BoxDailyTrans(LngBoxID As Long, _
                              FromDate As Variant, _
                              toDate As Variant) As Boolean

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim cReportViewer As ClsReportViewer
    Dim Msg As String
    Dim cProgress As ClsProgress
    Dim cOptions As ClsCompanyInfo
    Dim StrFileName As String
    Dim i As Long
    Dim StrFromDate As String, StrToDate As String

    On Error GoTo ErrTrap
    StrFileName = App.path & "\Reports\BoxDailyTrans.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        Exit Function
    Else

        If Not IsNull(FromDate) Then
            StrFromDate = "" & SQLDate(CDate(FromDate), True) & ""
        Else
            StrFromDate = "" & SQLDate(CDate(#1/1/1753#), True) & ""
        End If

        If Not IsNull(toDate) Then
            StrToDate = "" & SQLDate(CDate(toDate), True) & ""
        Else
            StrToDate = "" & SQLDate(CDate(#1/1/9999#), True) & ""
        End If

        StrSQL = "Select * From dbo.QryBoxDailyTrans(" & StrFromDate & "," & StrToDate & ")QryBoxDailyTrans"
        StrSQL = StrSQL + " Where BoxID <> 0"

        If LngBoxID <> 0 Then
            StrSQL = StrSQL + " AND BoxID =" & LngBoxID
        End If
    End If

    Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
    cProgress.StartProgress
    'For I = I To 1000
    '    Dim TempDate As Date
    '    Set RsData = New ADODB.Recordset
    '    TempDate = Date - I
    '    StrSQL = "Select  dbo.QryBoxCreditUptoDate(1," & SQLDate(TempDate, True) & ") as StartBal"
    '    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '    Debug.Print I, RsData("StartBal").Value
    '    DoEvents
    'Next I
    Set RsData = New ADODB.Recordset
    Cn.CommandTimeout = 180
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute

    Do While RsData.State = adStateExecuting
        DoEvents
    Loop

    If RsData.BOF Or RsData.EOF Then
        If Not cProgress Is Nothing Then
            cProgress.StopProgess
            Set cProgress = Nothing
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Set cOptions = New ClsCompanyInfo
    xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), FromDate)
    xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(toDate), CDate("01/01/1000"), toDate)
    xReport.reporttitle = " ﬁ—Ì— „ «»⁄… ÌÊ„Ì… ·—’Ìœ «·Œ“‰"
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, WindowTarget

    If Not cProgress Is Nothing Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If

    Set cReportViewer = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Screen.MousePointer = vbDefault

    If Not cProgress Is Nothing Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If

    Set xReport = Nothing
    Msg = "·«Ì„ﬂ‰ ⁄—÷ «· ﬁ—Ì— .."
    Msg = Msg + Chr(13) + Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Function
