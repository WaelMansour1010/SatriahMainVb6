VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNotesReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function ShowCheckReport(IntReportType As Integer, _
                                FromDate As Variant, _
                                ToDate As Variant, _
                                Optional m_PrintTarget As PrintTarget = WindowTarget)

    'IntReportType=0 '«·‘Ìﬂ«  «· Ï  „  Õ’Ì·Â«
    'IntReportType=1 ‘Ìﬂ«   Õ  «· Õ’Ì·
    'IntReportType = 2 ‘Ìﬂ«   „  ”œÌœÂ«
    'IntReportType=3 ‘Ìﬂ«  „ÿ·Ê» ”œ«œÂ«

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim StrFileName As String
    Dim StrSQL As String
    Dim cReportViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo
    Dim StrReportTitle As String

    On Error GoTo ErrTrap

    If IntReportType = 0 Or IntReportType = 2 Then
        StrFileName = App.path & "\Reports\CheckReleased.rpt"
    Else
        StrFileName = App.path & "\Reports\CheckUnReleased.rpt"
    End If

    If Dir(StrFileName) = "" Or StrFileName = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    Set rs = New ADODB.Recordset

    If IntReportType = 0 Then
        StrReportTitle = " ﬁ—Ì— »«·‘Ìﬂ«  «· Ï  „  Õ’Ì·Â«"
        StrSQL = "Select * From QryChecksReports "
        StrSQL = StrSQL + " Where OperaType=0"
    ElseIf IntReportType = 2 Then
        StrReportTitle = " ﬁ—Ì— »«·‘Ìﬂ«  «· Ï  „ ”œ«œÂ«"
        StrSQL = "Select * From QryChecksReports "
        StrSQL = StrSQL + " Where OperaType=1"
    ElseIf IntReportType = 1 Then
        StrReportTitle = " ﬁ—Ì— »«·‘Ìﬂ«   Õ  «· Õ’Ì·"
        StrSQL = "Select * From QryChecksReports "
        StrSQL = StrSQL + " Where (OperaID is NUll) AND (NoteType=2)"
    ElseIf IntReportType = 3 Then
        StrReportTitle = " ﬁ—Ì— »«·‘Ìﬂ«  «·„ÿ·Ê» ”œ«œÂ«"
        StrSQL = "Select * From QryChecksReports "
        StrSQL = StrSQL + " Where (OperaID is NUll) AND (NoteType=13)"
    End If

    If IntReportType = 0 Or IntReportType = 2 Then
        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND OperaDate >=" & SQLDate(CDate(FromDate), True) & ""
        End If

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND OperaDate <=" & SQLDate(CDate(ToDate), True) & ""
        End If

    ElseIf IntReportType = 1 Or IntReportType = 2 Then

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND DueDate >=" & SQLDate(CDate(FromDate), True) & ""
        End If

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND DueDate <=" & SQLDate(CDate(ToDate), True) & ""
        End If
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        rs.Close
        Set rs = Nothing
        Exit Function
    ElseIf rs.RecordCount = 0 Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        rs.Close
        Set rs = Nothing
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = StrReportTitle
        Screen.MousePointer = vbDefault
    End If

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_PrintTarget, StrReportTitle

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

 
 Public Sub ShowAllowsDiscounts(IntReportType As Integer, _
                               FromDate As Variant, _
                               ToDate As Variant, _
                               Optional m_PrintTarget As PrintTarget = WindowTarget, _
                               Optional LngCusID As Long = 0, Optional CurrenrEmployeeIDs As String = "", Optional CurrenrEmployeeNames As String = "", Optional ByVal BrcnActivety As String = "", Optional ByVal mbranch_id As Long = 0)

    'IntReportType=0 'Œ’Ê„«  „ﬂ ”»…
    'IntReportType=1 Œ’Ê„«  „”„ÊÕ…
    'IntReportType=2 Œ’Ê„«  „”„ÊÕ… «Ê „ﬂ ”»…
    
    'IntReportType=4 8034 œÌÊ‰ „⁄œÊ„…
    'IntReportType=5 9082 «‘⁄«— „œÌ‰
    'IntReportType=6 9083 «‘⁄«— œ«∆‰
    'IntReportType=7 9089 ﬁÌ„… «÷«›… «÷«›…
    'IntReportType=8 9090 ﬁÌ„… „÷«›… Œ’„
    'IntReportType=9 9099  ’›Ì… «„·«ﬂ

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim StrFileName As String
    Dim StrSQL As String
    Dim cReportViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo
    Dim StrReportTitle As String

    On Error GoTo ErrTrap
    StrFileName = App.path & "\Reports\AllowsDiscounts.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
StrSQL = " SELECT     dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.Note_Value, dbo.Notes.Transaction_ID,"
  StrSQL = StrSQL + " dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Type, dbo.TransactionTypes.TransactionTypeName, dbo.TblCustemers.CusName,"
   StrSQL = StrSQL + "dbo.TblUsers.UserName, dbo.TblNotesTypes.NotesTypeName, dbo.Notes.Remark, dbo.Notes.CusID, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
  StrSQL = StrSQL + " dbo.Notes.NoteSerial1 , dbo.Notes.EmpID, dbo.TblEmployee.emp_code, dbo.TblEmployee.emp_name, dbo.TblEmployee.Emp_Namee"
  
 StrSQL = StrSQL + " FROM            TblEmployee RIGHT OUTER JOIN"
 StrSQL = StrSQL + "                         TblCustemers RIGHT OUTER JOIN"
 StrSQL = StrSQL + "                         TblNotesTypes RIGHT OUTER JOIN"
 StrSQL = StrSQL + "                         Notes ON TblNotesTypes.NotesType = Notes.NoteType ON TblCustemers.CusID = Notes.CusID ON TblEmployee.Emp_ID = Notes.EmpId LEFT OUTER JOIN"
 StrSQL = StrSQL + "                         Transactions LEFT OUTER JOIN"
 StrSQL = StrSQL + "                         TransactionTypes ON Transactions.Transaction_Type = TransactionTypes.Transaction_Type ON"
 StrSQL = StrSQL + "                         Notes.Transaction_ID = Transactions.Transaction_ID RIGHT OUTER JOIN"
 StrSQL = StrSQL + "                         TblUsers ON Notes.UserID = TblUsers.UserID"
  
  

                      
    'IntReportType=4 8034 œÌÊ‰ „⁄œÊ„…
    'IntReportType=5 9082 «‘⁄«— „œÌ‰
    'IntReportType=6 9083 «‘⁄«— œ«∆‰
    'IntReportType=7 9089 ﬁÌ„… «÷«›… «÷«›…
    'IntReportType=8 9090 ﬁÌ„… „÷«›… Œ’„
    'IntReportType=9 9099  ’›Ì… «„·«ﬂ

 
    If IntReportType = 0 Then
        StrReportTitle = " ﬁ—Ì— »«·Œ’Ê„«  «·„ﬂ ”»…"
        'StrSQL = "Select * From QryAllowsDiscounts "
        StrSQL = StrSQL + " Where NoteType=10"
    ElseIf IntReportType = 1 Then
        StrReportTitle = " ﬁ—Ì— »«·Œ’Ê„«  «·„”„ÊÕ…"
        'StrSQL = "Select * From QryAllowsDiscounts "
        StrSQL = StrSQL + " Where NoteType=9"
        
    ElseIf IntReportType = 4 Then
        StrReportTitle = " ﬁ—Ì— »«·œÌÊ‰ «·„⁄œÊ„…"
        'StrSQL = "Select * From QryAllowsDiscounts "
        StrSQL = StrSQL + " Where NoteType=8034"
        
    ElseIf IntReportType = 5 Then
        StrReportTitle = " ﬁ—Ì— »«·«‘⁄«—«  «·„œÌ‰…"
        'StrSQL = "Select * From QryAllowsDiscounts "
        StrSQL = StrSQL + " Where NoteType=9082"
        
    ElseIf IntReportType = 6 Then
        StrReportTitle = " ﬁ—Ì— »«·«‘⁄«—«  «·œ«∆‰…"
        'StrSQL = "Select * From QryAllowsDiscounts "
        StrSQL = StrSQL + " Where NoteType=9083"
        
    ElseIf IntReportType = 7 Then
        StrReportTitle = " ﬁ—Ì— »«·ﬁÌ„… «·„÷«›… - «÷«›…"
        'StrSQL = "Select * From QryAllowsDiscounts "
        StrSQL = StrSQL + " Where NoteType=9089"
        
    ElseIf IntReportType = 8 Then
        StrReportTitle = "  ﬁ—Ì— »«·ﬁÌ„… «·„÷«›… - Œ’„"
        'StrSQL = "Select * From QryAllowsDiscounts "
        StrSQL = StrSQL + " Where NoteType=9090"
        
    ElseIf IntReportType = 9 Then
        StrReportTitle = "  ﬁ—Ì—  ’›Ì… «„·«ﬂ"
        'StrSQL = "Select * From QryAllowsDiscounts "
        StrSQL = StrSQL + " Where NoteType=9099"
        
    Else
        StrReportTitle = " ﬁ—Ì— »«·Œ’Ê„«  «·„”„ÊÕ… √Ê«·„ﬂ ”»… "
        'StrSQL = "Select * From QryAllowsDiscounts "
        StrSQL = StrSQL + " Where (NoteType=9 OR NoteType=10 )"
    End If

If CurrenrEmployeeNames <> "" And (IntReportType = 0 Or IntReportType = 1 Or IntReportType = 4) Then
StrReportTitle = StrReportTitle & CHR(13) & " ··„‰œÊ»  " & CurrenrEmployeeNames
End If

      If val(mbranch_id) <> 0 Then
            StrSQL = StrSQL & " and Notes.branch_no  =" & val(mbranch_id) & ""
        End If
   

    If LngCusID <> 0 Then
        StrSQL = StrSQL + " AND CusID=" & LngCusID & ""
    End If

  If CurrenrEmployeeIDs <> "" And (IntReportType = 0 Or IntReportType = 1 Or IntReportType = 4) Then
  
  StrSQL = StrSQL + "  and  Notes.EmpId in(" & CurrenrEmployeeIDs & ")"
  
  End If
  
    If BrcnActivety <> "" Then
        StrSQL = StrSQL & " and Notes.branch_no  in (" & BrcnActivety & ")"
    End If
  
    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(CDate(FromDate), True) & ""
    End If

    If Not IsNull(ToDate) Then
        StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(CDate(ToDate), True) & ""
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        rs.Close
        Set rs = Nothing
        Exit Sub
    ElseIf rs.RecordCount = 0 Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        rs.Close
        Set rs = Nothing
        Exit Sub
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), FromDate)
        xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(ToDate), CDate("01/01/1000"), ToDate)
        xReport.reporttitle = StrReportTitle
        Screen.MousePointer = vbDefault
    End If

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_PrintTarget, StrReportTitle, , , , StrFileName

    Exit Sub
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Sub




Public Sub ShowCompanyDebitValues(IntReportType As Integer, _
                                  FromDate As Variant, _
                                  ToDate As Variant, _
                                  BolShort As Boolean, _
                                  Optional m_PrintTarget As PrintTarget = WindowTarget, _
                                  Optional LngCusID As Long = 0)

    'IntReportType=0 Show Company Debit Values Report
    'IntReportType=1 Show Company Credit Values Report
    'IntReportType=2 Show Company Debit & Credit Values Report

    Dim xApp As CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim rs As ADODB.Recordset
    Dim XCmd As ADODB.Command
    Dim XPar As ADODB.Parameter
    Dim Msg As String
    Dim StrFileName As String
    Dim StrSQL As String
    Dim cReportViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo
    Dim StrReportTitle As String
    Dim StrPostedTransID As String

    'AdCmd.CommandType = adCmdStoredProc
    'AdCmd.CommandText = StrSQL
    'AdCmd.Parameters.Append AdPar

    On Error GoTo ErrTrap
    StrFileName = App.path & "\Reports\CompanyDebitCreditValues.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Set rs = New ADODB.Recordset

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        If IntReportType = 0 Then
            StrReportTitle = " ﬁ—Ì— »«·ﬁÌ„ «·„«·Ì… «·√Ã·… ⁄·Ï «·‘—ﬂ…"
            Set rs = GetDebitCreditValues(0, FromDate, ToDate, LngCusID)
        ElseIf IntReportType = 1 Then
            StrReportTitle = " ﬁ—Ì— »«·ﬁÌ„ «·„«·Ì… «·√Ã·… ··‘—ﬂ…"
            Set rs = GetDebitCreditValues(1, FromDate, ToDate, LngCusID)
        End If

    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then

        If IntReportType = 0 Then
            StrReportTitle = " ﬁ—Ì— »«·ﬁÌ„ «·„«·Ì… «·√Ã·… ⁄·Ï «·‘—ﬂ…"
            Set rs = GetDebitCreditValues(0, FromDate, ToDate, LngCusID)
        ElseIf IntReportType = 1 Then
            StrReportTitle = " ﬁ—Ì— »«·ﬁÌ„ «·„«·Ì… «·√Ã·… ··‘—ﬂ…"
            Set rs = GetDebitCreditValues(1, FromDate, ToDate, LngCusID)
        End If
    End If

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        rs.Close
        Set rs = Nothing
        Exit Sub
    ElseIf rs.RecordCount = 0 Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        rs.Close
        Set rs = Nothing
        Exit Sub
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xApp = New CRAXDRT.Application
        Set xReport = New CRAXDRT.Report
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        'xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), FromDate)
        'xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(ToDate), CDate("01/01/1000"), ToDate)
        xReport.reporttitle = StrReportTitle
        Screen.MousePointer = vbDefault
    End If

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_PrintTarget, StrReportTitle

    Exit Sub
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Sub

Public Function PrintReceipt(NoteID As Long, _
                             Optional m_Target As PrintTarget = WindowTarget) As Boolean

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Msg As String
    Dim StrTemp As String
    Dim StrFileName As String
    On Error GoTo ErrTrap

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, TblNotesTypes.NotesTypeName" & ", Notes.NoteSerial,Cdbl( Notes.Note_Value) as Note_Value, Notes.BankID, Notes.ChqueNum, Notes.DueDat" & "e, Transactions.Transaction_Serial, TransactionTypes.TransactionTypeName, Transa" & "ctions.Transaction_ID, TblCustemers.CusName, TblEmployee.Emp_Name, Notes.Remark," & " BanksData.BankName FROM TransactionTypes RIGHT JOIN ((TblEmployee RIGHT JOIN Tr" & "ansactions ON TblEmployee.Emp_ID = Transactions.Emp_ID) RIGHT JOIN (TblNotesType" & "s INNER JOIN (TblCustemers INNER JOIN (BanksData RIGHT JOIN Notes ON BanksData.B" & "ankID = Notes.BankID) ON TblCustemers.CusID = Notes.CusID) ON TblNotesTypes.Note" & "sType = Notes.NoteType) ON Transactions.Transaction_ID = Notes.Transaction_ID) O" & "N TransactionTypes.Transaction_Type = Transactions.Transaction_Type"
        StrSQL = StrSQL + " WHERE (((Notes.NoteID)=" & NoteID & "))"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT     dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.TblNote" & _
           "sTypes.NotesTypeName, dbo.Notes.NoteSerial, dbo.Notes.Note_Value,               " & _
           "         dbo.Notes.BankID, dbo.BanksData.BankName, dbo.Notes.ChqueNum, dbo.Notes" & _
           ".DueDate, dbo.Transactions.Transaction_Serial,                        dbo.Transa" & _
           "ctionTypes.TransactionTypeName, dbo.Transactions.Transaction_ID, dbo.TblCustemer" & _
           "s.CusName, dbo.TblEmployee.Emp_Name,Notes.Remark FROM         dbo.TblEmployee RIGHT OUTER JOI" & _
           "N                       dbo.TransactionTypes RIGHT OUTER JOIN                   " & _
           "    dbo.Transactions ON dbo.TransactionTypes.Transaction_Type = dbo.Transactions" & _
           ".Transaction_Type ON                        dbo.TblEmployee.Emp_ID = dbo.Transac" & _
           "tions.Emp_ID RIGHT OUTER JOIN                       dbo.TblCustemers INNER JOIN " & _
           "                      dbo.TblNotesTypes INNER JOIN                       dbo.Ban" & _
           "ksData RIGHT OUTER JOIN                       dbo.Notes ON dbo.BanksData.BankID " & _
           "= dbo.Notes.BankID ON dbo.TblNotesTypes.NotesType = dbo.Notes.NoteType ON       " & _
           "                 dbo.TblCustemers.CusID = dbo.Notes.CusID ON dbo.Transactions.Tr" & _
           "ansaction_ID = dbo.Notes.Transaction_ID"
        StrSQL = StrSQL + " WHERE (((Notes.NoteID)=" & NoteID & "))"
    End If

    'End If
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Function
    End If

    StrFileName = App.path & "\Reports\NoteReceipt.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        StrTemp = IIf(IsNull(rs("Note_Value").value), "", WriteNo(rs("Note_Value").value, 0))
        ' StrTemp = StrTemp & " " & " —Ì«· ”⁄ÊœÌ ›ﬁÿ ·«€Ì— "
        xReport.ParameterFields(3).AddCurrentValue StrTemp
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, m_Target
    Set CViewer = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function
