VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAccReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private RPTCompany_Name_Arabic As String

Private RPTComment_Arabic As String

Private RPTCompany_Name_Eng As String

Private RPTComment_Eng As String

Private RPTCurrency As String

Private IntZoomLevel As Integer
Dim m_EndDate As Variant
Dim m_BegineDate As Variant
 Public Function ShowLedger40(StrAccountCode As String, _
                           StrAccountName As String, _
                           Optional manyString As Boolean = False, _
                           Optional d1 As Date, _
                           Optional d2 As Date, _
                           Optional ActivityId As Integer = 0, _
                           Optional BranchID As Integer = 0, _
                           Optional CustomerAgeingData As String, _
                           Optional salesPersonName As String, _
                           Optional ShowAgingReport As String, _
                           Optional ByRef account_serial As String, Optional project_id As Integer, Optional ProjectName As String, Optional detailedtransaction As Integer = 0, Optional ShowValuee As Integer = 0, Optional RegionID As Integer = 0, Optional REtype As Integer = -1) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String


 MySQL = " SELECT        dbo.RptLedger_Sub.ChqueNum, dbo.RptLedger_Sub.ManualNo, dbo.RptLedger_Sub.RecordDateH, dbo.RptLedger_Sub.Double_Entry_Vouchers_ID, dbo.RptLedger_Sub.DEV_Value, "
 MySQL = MySQL & "                         dbo.RptLedger_Sub.Credit_Or_Debit, dbo.ACCOUNTS.Account_Code, dbo.RptLedger_Sub.DEV_DES, dbo.RptLedger_Sub.DevDESE, dbo.RptLedger_Sub.Account_Name, dbo.RptLedger_Sub.DEV_ID_Line_No,"
 MySQL = MySQL & "                         dbo.RptLedger_Sub.NotesTypeName, dbo.RptLedger_Sub.UserID, dbo.RptLedger_Sub.UserName, dbo.RptLedger_Sub.RecordDate, dbo.RptLedger_Sub.ReceiptID, dbo.RptLedger_Sub.Notes_ID,"
 MySQL = MySQL & "                         dbo.RptLedger_Sub.OperaID, dbo.RptLedger_Sub.Transaction_ID, dbo.RptLedger_Sub.Transaction_Serial, dbo.RptLedger_Sub.Transaction_Date, dbo.RptLedger_Sub.Posted,"
 MySQL = MySQL & "                         dbo.RptLedger_Sub.TransactionTypeName, dbo.RptLedger_Sub.PostedDate, dbo.RptLedger_Sub.PostedUserID, dbo.RptLedger_Sub.Account_Interval_ID, dbo.RptLedger_Sub.NoteDate,"
 MySQL = MySQL & "                         dbo.RptLedger_Sub.NoteType, dbo.RptLedger_Sub.NoteSerial, dbo.RptLedger_Sub.Note_Value, dbo.RptLedger_Sub.Account_Serial, dbo.RptLedger_Sub.Account_NameEng,"
 MySQL = MySQL & "                         dbo.RptLedger_Sub.Parent_Account_Code, dbo.RptLedger_Sub.Branch, dbo.RptLedger_Sub.Sum_account, dbo.RptLedger_Sub.cost_center, dbo.RptLedger_Sub.currenct_code, dbo.RptLedger_Sub.Remark,"
 MySQL = MySQL & "                         dbo.RptLedger_Sub.note_value_by_characters, dbo.RptLedger_Sub.foxy_no, dbo.RptLedger_Sub.DEV_ID_Line_No1, dbo.RptLedger_Sub.project_id, dbo.RptLedger_Sub.NotesTypeNamee,"
 MySQL = MySQL & "                         dbo.RptLedger_Sub.TransactionEnglishName, dbo.RptLedger_Sub.opr_fullcode, dbo.RptLedger_Sub.NoteSerial1, dbo.RptLedger_Sub.branch_id, dbo.RptLedger_Sub.ActivityTypeId,"
 MySQL = MySQL & "                         dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.RptLedger_Sub.notes_all, dbo.RptLedger_Sub.DEV_ValueE, dbo.RptLedger_Sub.currency, dbo.RptLedger_Sub.rate,"
 MySQL = MySQL & "                         dbo.RptLedger_Sub.name, dbo.RptLedger_Sub.namee, dbo.RptLedger_Sub.branch_name, dbo.RptLedger_Sub.branch_namee, dbo.RptLedger_Sub.RegionID, dbo.RptLedger_Sub.DescAccount,"
 MySQL = MySQL & "                         dbo.RptLedger_Sub.aqarname , dbo.RptLedger_Sub.Aqarid, dbo.RptLedger_Sub.aqarNo"
 MySQL = MySQL & "  FROM            dbo.ACCOUNTS LEFT OUTER JOIN"
 MySQL = MySQL & "                         dbo.RptLedger_Sub ON dbo.ACCOUNTS.Account_Code = dbo.RptLedger_Sub.Account_Code"
 MySQL = MySQL & "   WHERE  1=1"
If REtype = 0 Then
  MySQL = MySQL & " and RptLedger_Sub.Credit_Or_Debit =0 "
ElseIf REtype = 1 Then
MySQL = MySQL & " and RptLedger_Sub.Credit_Or_Debit =1 "
End If

    If BranchID <> 0 Then
        MySQL = MySQL & " and RptLedger_Sub.branch_id = " & BranchID 'Val(FrmAccountingReport.dcBranch.BoundText)
    End If

    If ActivityId <> 0 Then
        MySQL = MySQL & " and dbo.RptLedger_Sub.ActivityTypeId = " & ActivityId 'Val(FrmAccountingReport.DCActivity.BoundText)
    End If
    If RegionID <> 0 Then
        MySQL = MySQL & " and dbo.RptLedger_Sub.RegionID = " & RegionID 'Val(FrmAccountingReport.DCActivity.BoundText)
    End If
  If project_id <> 0 Then
        MySQL = MySQL & " and dbo.RptLedger_Sub.Aqarid = " & project_id 'Val(FrmAccountingReport.DCActivity.BoundText)
    End If
    
    
    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and (RptLedger_Sub.RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and RptLedger_Sub.RecordDate <=" & SQLDate(EndDate, True) & ""
    End If
 
    MySQL = MySQL + " or  RptLedger_Sub.RecordDate is null)"
     
   ' updateProfitsToall


        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\REPORTS NEW\" & "Sub-MassterAqar.rpt"
        Else
            StrFileName = App.path & "\Reports\REPORTS NEW\" & "Sub-MassterAqar.rpt"
        End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
        
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        RsData.Close

Exit Function

    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If d1 <> Null Then
        Me.BegineDate = d1
    End If

    If d2 <> Null Then
        Me.EndDate = d2
    End If

    Dim activityName As String
    Dim branchname As String
    Dim fullcode As String
    branchname = get_branch_name(val(BranchID), activityName)
GetCustomerIdByAccountCodeLong StrAccountCode, fullcode

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(branch_id))
        xReport.ParameterFields(7).AddCurrentValue openingbalanceDes
        'CustomerAgeingData
        xReport.ParameterFields(8).AddCurrentValue CustomerAgeingData





        
       If fullcode <> "" Then
            xReport.ParameterFields(13).AddCurrentValue "ﬂÊœ «·⁄„Ì· : " & fullcode
        Else
            xReport.ParameterFields(13).AddCurrentValue ""
        
     End If
     
        
        If salesPersonName <> "" Then
            xReport.ParameterFields(9).AddCurrentValue "  «·„‰œÊ» " & salesPersonName
        Else
            xReport.ParameterFields(9).AddCurrentValue "    "
        End If

        xReport.ParameterFields(10).AddCurrentValue ShowAgingReport

        If ChartPrintinAS = True Then
            xReport.ParameterFields(11).AddCurrentValue "1"
        Else
            xReport.ParameterFields(11).AddCurrentValue "0"
        End If

If REtype = 0 Then
   StrReportTitle = "ﬂ‘› „’—Ê›«  ··⁄ﬁ«—   "
ElseIf REtype = 1 Then
StrReportTitle = "ﬂ‘› «Ì—«œ«   ··⁄ﬁ«—   "
Else
 StrReportTitle = "  ﬂ‘› Õ”«»  ··⁄ﬁ«— "
End If

        'ShowAgingReport
        'StrReportTitle = "ﬂ‘› Õ”«»   " '& StrAccountName



        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        
            If SystemOptions.DateOpt = 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰ " & Format(Me.BegineDate, "yyyy/Mm/dd") & "  "
            Else
                StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰ " & ToHijriDate(Format(Me.BegineDate, "yyyy/Mm/dd")) & "  "

            End If
    
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            If SystemOptions.DateOpt = 0 Then

                StrReportTitle = StrReportTitle & CHR(13) & "   ≈·Ï " & "   " & Format(EndDate, "yyyy/Mm/dd")
            Else
                StrReportTitle = StrReportTitle & CHR(13) & "   ≈·Ï " & "   " & ToHijriDate(Format(EndDate, "yyyy/Mm/dd"))
            End If
        End If
     
        If ActivityId <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & activityName
        End If

        If BranchID <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & branchname
        End If
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue branchname  'get_branch_name(Val(branchId))
        xReport.ParameterFields(7).AddCurrentValue openingbalanceDes
        xReport.ParameterFields(8).AddCurrentValue CustomerAgeingData
     
       If fullcode <> "" Then
            xReport.ParameterFields(13).AddCurrentValue "Code " & fullcode
        Else
          xReport.ParameterFields(13).AddCurrentValue ""
        
     End If
     
        StrReportTitle = "Statement of ACC :" & account_serial & CHR(13) & StrAccountName & CHR(13)    '& " Account "
    
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From : " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To   :  " & (EndDate)
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(5).AddCurrentValue account_serial
    xReport.ParameterFields(6).AddCurrentValue StrAccountName
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    'RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

 Public Function ShowLedger22(StrAccountCode As String, _
                           StrAccountName As String, _
                           Optional manyString As Boolean = False, _
                           Optional d1 As Date, _
                           Optional d2 As Date, _
                           Optional ActivityId As Integer = 0, _
                           Optional BranchID As Integer = 0, _
                           Optional CustomerAgeingData As String, _
                           Optional salesPersonName As String, _
                           Optional ShowAgingReport As String, _
                           Optional ByRef account_serial As String, Optional project_id As Integer, Optional ProjectName As String, Optional detailedtransaction As Integer = 0, Optional ShowValuee As Integer = 0, Optional RegionID As Integer = 0) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

      MySQL = " SELECT        dbo.RptLedger_Sub.ChqueNum, dbo.RptLedger_Sub.ManualNo, dbo.RptLedger_Sub.RecordDateH, dbo.RptLedger_Sub.Double_Entry_Vouchers_ID, dbo.RptLedger_Sub.DEV_Value, "
      MySQL = MySQL + "                   dbo.RptLedger_Sub.Credit_Or_Debit, ACCOUNTS_1.Account_Code, dbo.RptLedger_Sub.DEV_DES, dbo.RptLedger_Sub.DevDESE, dbo.RptLedger_Sub.Account_Name, dbo.RptLedger_Sub.DEV_ID_Line_No,"
      MySQL = MySQL + "                   dbo.RptLedger_Sub.NotesTypeName, dbo.RptLedger_Sub.UserID, dbo.RptLedger_Sub.UserName, dbo.RptLedger_Sub.RecordDate, dbo.RptLedger_Sub.ReceiptID, dbo.RptLedger_Sub.Notes_ID,"
      MySQL = MySQL + "                   dbo.RptLedger_Sub.OperaID, dbo.RptLedger_Sub.Transaction_ID, dbo.RptLedger_Sub.Transaction_Serial, dbo.RptLedger_Sub.Transaction_Date, dbo.RptLedger_Sub.Posted,"
      MySQL = MySQL + "                   dbo.RptLedger_Sub.TransactionTypeName, dbo.RptLedger_Sub.PostedDate, dbo.RptLedger_Sub.PostedUserID, dbo.RptLedger_Sub.Account_Interval_ID, dbo.RptLedger_Sub.NoteDate,"
      MySQL = MySQL + "                   dbo.RptLedger_Sub.NoteType, dbo.RptLedger_Sub.NoteSerial, dbo.RptLedger_Sub.Note_Value, dbo.RptLedger_Sub.Account_Serial, dbo.RptLedger_Sub.Account_NameEng,"
      MySQL = MySQL + "                   dbo.RptLedger_Sub.Parent_Account_Code, dbo.RptLedger_Sub.Branch, dbo.RptLedger_Sub.Sum_account, dbo.RptLedger_Sub.cost_center, dbo.RptLedger_Sub.currenct_code, dbo.RptLedger_Sub.Remark,"
      MySQL = MySQL + "                   dbo.RptLedger_Sub.note_value_by_characters, dbo.RptLedger_Sub.foxy_no, dbo.RptLedger_Sub.DEV_ID_Line_No1, dbo.RptLedger_Sub.project_id, dbo.RptLedger_Sub.NotesTypeNamee,"
      MySQL = MySQL + "                   dbo.RptLedger_Sub.TransactionEnglishName, dbo.RptLedger_Sub.opr_fullcode, dbo.RptLedger_Sub.NoteSerial1, dbo.RptLedger_Sub.branch_id, dbo.RptLedger_Sub.ActivityTypeId,"
      MySQL = MySQL + "                   ACCOUNTS_1.opening_balance, ACCOUNTS_1.opening_balance_type, dbo.RptLedger_Sub.notes_all, dbo.RptLedger_Sub.DEV_ValueE, dbo.RptLedger_Sub.currency, dbo.RptLedger_Sub.rate,"
      MySQL = MySQL + "                   dbo.RptLedger_Sub.name, dbo.RptLedger_Sub.NameE, dbo.RptLedger_Sub.branch_name, dbo.RptLedger_Sub.branch_namee, dbo.RptLedger_Sub.RegionID, dbo.Transactions.CashCustomerName,"
      MySQL = MySQL + "                   dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.RptLedger_Sub.DescAccount, dbo.RptLedger_Sub.NextAccount_Code,"
      MySQL = MySQL + "                   dbo.ACCOUNTS.Account_Name AS NextAccount_Name, dbo.ACCOUNTS.Account_NameEng AS NextAccount_NameE"
      MySQL = MySQL + "      FROM            dbo.ACCOUNTS RIGHT OUTER JOIN"
      MySQL = MySQL + "                   dbo.RptLedger_Sub ON dbo.ACCOUNTS.Account_Code = dbo.RptLedger_Sub.NextAccount_Code LEFT OUTER JOIN"
      MySQL = MySQL + "                   dbo.TblCustemers RIGHT OUTER JOIN"
      MySQL = MySQL + "                   dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID ON dbo.RptLedger_Sub.Transaction_ID = dbo.Transactions.Transaction_ID RIGHT OUTER JOIN"
      MySQL = MySQL + "                     dbo.ACCOUNTS AS ACCOUNTS_1 ON dbo.RptLedger_Sub.Account_Code = ACCOUNTS_1.Account_Code"
      MySQL = MySQL + " WHERE   ACCOUNTS_1.Account_Code ='" & StrAccountCode & "'"

    If BranchID <> 0 Then
        MySQL = MySQL & " and dbo.RptLedger_Sub.branch_id = " & BranchID 'Val(FrmAccountingReport.dcBranch.BoundText)
    End If

    If ActivityId <> 0 Then
        MySQL = MySQL & " and dbo.RptLedger_Sub.ActivityTypeId = " & ActivityId 'Val(FrmAccountingReport.DCActivity.BoundText)
    End If
    If RegionID <> 0 Then
        MySQL = MySQL & " and dbo.RptLedger_Sub.RegionID = " & RegionID 'Val(FrmAccountingReport.DCActivity.BoundText)
    End If
    If project_id <> 0 Then
        MySQL = MySQL & " and dbo.RptLedger_Sub.project_id = " & project_id 'Val(FrmAccountingReport.DCActivity.BoundText)
    End If
    
    
    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and (dbo.RptLedger_Sub.RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and dbo.RptLedger_Sub.RecordDate <=" & SQLDate(EndDate, True) & ""
    End If
 
    MySQL = MySQL + " or  dbo.RptLedger_Sub.RecordDate is null)"
     
    updateProfitsToall

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\REPORTS NEW\" & "Sub-MassterEng22.rpt"
        Else
            StrFileName = App.path & "\Reports\REPORTS NEW\" & "Sub-MassterEng22.rpt"
        End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
        
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText



    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If d1 <> Null Then
        Me.BegineDate = d1
    End If

    If d2 <> Null Then
        Me.EndDate = d2
    End If

    Dim activityName As String
    Dim branchname As String
    Dim fullcode As String
    branchname = get_branch_name(val(BranchID), activityName)
GetCustomerIdByAccountCodeLong StrAccountCode, fullcode

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(branch_id))
        xReport.ParameterFields(7).AddCurrentValue openingbalanceDes
        'CustomerAgeingData
        xReport.ParameterFields(8).AddCurrentValue CustomerAgeingData
        
       If fullcode <> "" Then
            xReport.ParameterFields(13).AddCurrentValue "ﬂÊœ «·⁄„Ì· : " & fullcode
        Else
            xReport.ParameterFields(13).AddCurrentValue ""
        
     End If
     
        
        If salesPersonName <> "" Then
            xReport.ParameterFields(9).AddCurrentValue "  «·„‰œÊ» " & salesPersonName
        Else
            xReport.ParameterFields(9).AddCurrentValue "    "
        End If

        xReport.ParameterFields(10).AddCurrentValue ShowAgingReport

        If ChartPrintinAS = True Then
            xReport.ParameterFields(11).AddCurrentValue "1"
        Else
            xReport.ParameterFields(11).AddCurrentValue "0"
        End If

        'ShowAgingReport
        StrReportTitle = "ﬂ‘› Õ”«»   " & StrAccountName

If ProjectName <> "" Then
StrReportTitle = StrReportTitle & " ··„‘—Ê⁄  " & ProjectName
End If

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        
            If SystemOptions.DateOpt = 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰ " & Format(Me.BegineDate, "yyyy/Mm/dd") & "  "
            Else
                StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰ " & ToHijriDate(Format(Me.BegineDate, "yyyy/Mm/dd")) & "  "

            End If
    
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            If SystemOptions.DateOpt = 0 Then

                StrReportTitle = StrReportTitle & CHR(13) & "   ≈·Ï " & "   " & Format(EndDate, "yyyy/Mm/dd")
            Else
                StrReportTitle = StrReportTitle & CHR(13) & "   ≈·Ï " & "   " & ToHijriDate(Format(EndDate, "yyyy/Mm/dd"))
            End If
        End If
     
        If ActivityId <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & activityName
        End If

        If BranchID <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & branchname
        End If
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue branchname  'get_branch_name(Val(branchId))
        xReport.ParameterFields(7).AddCurrentValue openingbalanceDes
        xReport.ParameterFields(8).AddCurrentValue CustomerAgeingData
     
       If fullcode <> "" Then
            xReport.ParameterFields(13).AddCurrentValue "Code " & fullcode
        Else
          xReport.ParameterFields(13).AddCurrentValue ""
        
     End If
     
        StrReportTitle = "Statement of ACC :" & account_serial & CHR(13) & StrAccountName & CHR(13)    '& " Account "
    
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From : " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To   :  " & (EndDate)
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(5).AddCurrentValue account_serial
    xReport.ParameterFields(6).AddCurrentValue StrAccountName
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function


 Public Function ShowLedger(StrAccountCode As String, _
                           StrAccountName As String, _
                           Optional manyString As Boolean = False, _
                           Optional d1 As Date, _
                           Optional d2 As Date, _
                           Optional ActivityId As Integer = 0, _
                           Optional BranchID As Integer = 0, _
                           Optional CustomerAgeingData As String, _
                           Optional salesPersonName As String, _
                           Optional ShowAgingReport As String, _
                           Optional ByRef account_serial As String, Optional project_id As Integer, Optional ProjectName As String, Optional detailedtransaction As Integer = 0, Optional ShowValuee As Integer = 0, Optional RegionID As Integer = 0, Optional withouopenening As Boolean, Optional notesType As Double, Optional detailedtransactionBasi As Integer = 0) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    If Not manyString Then
    If CheckUserNotPermAccounts(CDbl(user_id), StrAccountCode) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·« Ì„ﬂ‰ «·«ÿ·«⁄ ⁄·Ì Â–« «·Õ”«»", vbCritical
        Else
        MsgBox "can't view this acc statement", vbCritical
        End If
        Exit Function
    End If
    End If
 
    
'    End If
    
    If manyString = False Then
        MySQL = "Select * From RptLedger_Sub "
        MySQL = MySQL + " WHERE Account_Code='" & StrAccountCode & "'"
    Else
        MySQL = "Select * From RptLedger_Sub "
        'MySQL = MySQL + StrAccountCode
    End If

    If manyString = False Then

        MySQL = " SELECT        dbo.RptLedger_Sub.ChqueNum, dbo.RptLedger_Sub.ManualNo, dbo.RptLedger_Sub.RecordDateH, dbo.RptLedger_Sub.Double_Entry_Vouchers_ID, dbo.RptLedger_Sub.DEV_Value,RptLedger_Sub.order_no, "
        MySQL = MySQL + "                 dbo.RptLedger_Sub.Credit_Or_Debit, dbo.ACCOUNTS.Account_Code, dbo.RptLedger_Sub.DEV_DES, dbo.RptLedger_Sub.DevDESE, dbo.RptLedger_Sub.Account_Name, dbo.RptLedger_Sub.DEV_ID_Line_No,"
        MySQL = MySQL + "                 dbo.RptLedger_Sub.NotesTypeName, dbo.RptLedger_Sub.UserID, dbo.RptLedger_Sub.UserName, dbo.RptLedger_Sub.RecordDate, dbo.RptLedger_Sub.ReceiptID, dbo.RptLedger_Sub.Notes_ID,"
        MySQL = MySQL + "                 dbo.RptLedger_Sub.OperaID, dbo.RptLedger_Sub.Transaction_ID, dbo.RptLedger_Sub.Transaction_Serial, dbo.RptLedger_Sub.Transaction_Date, dbo.RptLedger_Sub.Posted,"
        MySQL = MySQL + "                 dbo.RptLedger_Sub.TransactionTypeName, dbo.RptLedger_Sub.PostedDate, dbo.RptLedger_Sub.PostedUserID, dbo.RptLedger_Sub.Account_Interval_ID, dbo.RptLedger_Sub.NoteDate,"
        MySQL = MySQL + "                 dbo.RptLedger_Sub.NoteType, dbo.RptLedger_Sub.NoteSerial, dbo.RptLedger_Sub.Note_Value, dbo.RptLedger_Sub.Account_Serial, dbo.RptLedger_Sub.Account_NameEng,"
        MySQL = MySQL + "                 dbo.RptLedger_Sub.Parent_Account_Code, dbo.RptLedger_Sub.Branch, dbo.RptLedger_Sub.Sum_account, dbo.RptLedger_Sub.cost_center, dbo.RptLedger_Sub.currenct_code, dbo.RptLedger_Sub.Remark,"
        MySQL = MySQL + "                 dbo.RptLedger_Sub.note_value_by_characters, dbo.RptLedger_Sub.foxy_no, dbo.RptLedger_Sub.DEV_ID_Line_No1, dbo.RptLedger_Sub.project_id, dbo.RptLedger_Sub.NotesTypeNamee,"
        MySQL = MySQL + "                 dbo.RptLedger_Sub.TransactionEnglishName, dbo.RptLedger_Sub.opr_fullcode, dbo.RptLedger_Sub.NoteSerial1, dbo.RptLedger_Sub.branch_id, dbo.RptLedger_Sub.ActivityTypeId,"
        MySQL = MySQL + "                 dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.RptLedger_Sub.notes_all, dbo.RptLedger_Sub.DEV_ValueE, dbo.RptLedger_Sub.currency, dbo.RptLedger_Sub.rate,"
        MySQL = MySQL + "                 dbo.RptLedger_Sub.Name , dbo.RptLedger_Sub.NameE, dbo.RptLedger_Sub.branch_name, dbo.RptLedger_Sub.branch_namee, dbo.RptLedger_Sub.RegionID"
        MySQL = MySQL + "           FROM            dbo.RptLedger_Sub RIGHT OUTER JOIN"
        MySQL = MySQL + "                 dbo.ACCOUNTS ON dbo.RptLedger_Sub.Account_Code = dbo.ACCOUNTS.Account_Code"
        MySQL = MySQL + "   WHERE  ACCOUNTS.Account_Code='" & StrAccountCode & "'"

 MySQL = " SELECT     dbo.RptLedger_Sub.ChqueNum, dbo.RptLedger_Sub.ManualNo, dbo.RptLedger_Sub.RecordDateH, dbo.RptLedger_Sub.Double_Entry_Vouchers_ID,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.DEV_Value, dbo.RptLedger_Sub.Credit_Or_Debit, dbo.ACCOUNTS.Account_Code, dbo.RptLedger_Sub.DEV_DES, dbo.RptLedger_Sub.DevDESE,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.Account_Name, dbo.RptLedger_Sub.DEV_ID_Line_No, dbo.RptLedger_Sub.NotesTypeName, dbo.RptLedger_Sub.UserID,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.UserName, dbo.RptLedger_Sub.RecordDate, dbo.RptLedger_Sub.ReceiptID, dbo.RptLedger_Sub.Notes_ID, dbo.RptLedger_Sub.OperaID,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.Transaction_ID, dbo.RptLedger_Sub.Transaction_Serial, dbo.RptLedger_Sub.Transaction_Date, dbo.RptLedger_Sub.Posted,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.TransactionTypeName, dbo.RptLedger_Sub.PostedDate, dbo.RptLedger_Sub.PostedUserID, dbo.RptLedger_Sub.Account_Interval_ID,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.NoteDate, dbo.RptLedger_Sub.NoteType, dbo.RptLedger_Sub.NoteSerial, dbo.RptLedger_Sub.Note_Value, dbo.RptLedger_Sub.Account_Serial,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.Account_NameEng, dbo.RptLedger_Sub.Parent_Account_Code, dbo.RptLedger_Sub.Branch, dbo.RptLedger_Sub.Sum_account,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.cost_center, dbo.RptLedger_Sub.currenct_code, dbo.RptLedger_Sub.Remark, dbo.RptLedger_Sub.note_value_by_characters,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.foxy_no, dbo.RptLedger_Sub.DEV_ID_Line_No1, dbo.RptLedger_Sub.project_id, dbo.RptLedger_Sub.NotesTypeNamee,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.TransactionEnglishName, dbo.RptLedger_Sub.opr_fullcode, dbo.RptLedger_Sub.NoteSerial1, dbo.RptLedger_Sub.branch_id,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.ActivityTypeId, dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.RptLedger_Sub.notes_all,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.DEV_ValueE, dbo.RptLedger_Sub.currency, dbo.RptLedger_Sub.rate, dbo.RptLedger_Sub.name, dbo.RptLedger_Sub.NameE,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.branch_name, dbo.RptLedger_Sub.branch_namee, dbo.RptLedger_Sub.RegionID, dbo.RptLedger_Sub.DescAccount,"
 MySQL = MySQL + "                        dbo.TblCustemers.CusNamee , dbo.TblCustemers.CusName,"
 MySQL = MySQL + "                        TblEmployee.Emp_Name,TblEmployee.Emp_Namee,"
 MySQL = MySQL + "       Emp_Code = (Select Emp_Code from TblEmployee where IsNull(TblEmployee.Account_code,'' ) = ACCOUNTS.Account_Code Or  IsNull(TblEmployee.Account_code2,'' ) = ACCOUNTS.Account_Code or IsNull(TblEmployee.Account_code3,'' ) = ACCOUNTS.Account_Code "
 MySQL = MySQL + "       or IsNull(TblEmployee.Account_code4,'' ) = ACCOUNTS.Account_Code or IsNull(TblEmployee.Account_code5,'' ) = ACCOUNTS.Account_Code  or IsNull(TblEmployee.Account_code1,'' ) = ACCOUNTS.Account_Code  ) "
MySQL = MySQL + "   FROM         dbo.TblCustemers RIGHT OUTER JOIN"
MySQL = MySQL + "                         dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID RIGHT OUTER JOIN"
MySQL = MySQL + "                         dbo.RptLedger_Sub ON dbo.Transactions.Transaction_ID = dbo.RptLedger_Sub.Transaction_ID RIGHT OUTER JOIN"
MySQL = MySQL + "                         dbo.ACCOUNTS ON dbo.RptLedger_Sub.Account_Code = dbo.ACCOUNTS.Account_Code"
MySQL = MySQL + "                             left OUTER JOIN dbo.TblEmployee"
MySQL = MySQL + "                             ON dbo.RptLedger_Sub.Account_Code = dbo.TblEmployee.Account_Code"
MySQL = MySQL + " WHERE  ACCOUNTS.Account_Code='" & StrAccountCode & "'"






If SystemOptions.IsHiddenUser = False Then
    MySQL = MySQL + " and  IsNull(RptLedger_Sub.IsHiddenInv,0) = 0 "
End If

    Else
        'MySQL = MySQL + StrAccountCode
         MySQL = MySQL & "  where  Account_Code in (" & StrAccountCode & ")"

    End If

If manyString = False Then

    MySQL = ""
    MySQL = MySQL + " SELECT     L.ChqueNum, L.ManualNo, L.RecordDateH, L.Double_Entry_Vouchers_ID, L.DEV_Value, L.order_no,L.CBoBasedON, " & vbCrLf
    MySQL = MySQL + "            L.Credit_Or_Debit, A.Account_Code, L.DEV_DES, L.DevDESE, L.Account_Name, L.DEV_ID_Line_No, " & vbCrLf
    MySQL = MySQL + "            L.NotesTypeName, L.UserID, L.UserName, L.RecordDate, L.ReceiptID, L.Notes_ID, " & vbCrLf
    MySQL = MySQL + "            L.OperaID, L.Transaction_ID, L.Transaction_Serial, L.Transaction_Date, L.Posted, " & vbCrLf
    MySQL = MySQL + "            L.TransactionTypeName, L.PostedDate, L.PostedUserID, L.Account_Interval_ID, L.NoteDate, " & vbCrLf
    MySQL = MySQL + "            L.NoteType, L.NoteSerial, L.Note_Value, L.Account_Serial, L.Account_NameEng, " & vbCrLf
    MySQL = MySQL + "            L.Parent_Account_Code, L.Branch, L.Sum_account, L.cost_center, L.currenct_code, L.Remark, " & vbCrLf
    MySQL = MySQL + "            L.note_value_by_characters, L.foxy_no, L.DEV_ID_Line_No1, L.project_id, L.NotesTypeNamee, " & vbCrLf
    MySQL = MySQL + "            L.TransactionEnglishName, L.opr_fullcode, L.NoteSerial1, L.branch_id, L.ActivityTypeId, " & vbCrLf
    MySQL = MySQL + "            A.opening_balance, A.opening_balance_type, L.notes_all, L.DEV_ValueE, L.currency, L.rate, " & vbCrLf
    MySQL = MySQL + "            L.name, L.NameE, L.branch_name, L.branch_namee, L.RegionID, L.DescAccount, " & vbCrLf
    MySQL = MySQL + "            C.CusNamee, C.CusName, E.Emp_Name, E.Emp_Namee, EmpMap.Emp_Code " & vbCrLf

    MySQL = MySQL + " FROM       dbo.RptLedger_Sub L " & vbCrLf
    MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON L.Account_Code = A.Account_Code " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.Transactions T ON T.Transaction_ID = L.Transaction_ID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblCustemers C ON C.CusID = T.CusID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblEmployee  E ON E.Account_Code = L.Account_Code " & vbCrLf

    '»œÌ· √”—⁄ Ê¬„‰ „‰ «·‹ Subquery «··Ì ›ÌÂ OR ﬂ Ì— (Ê»Ì—Ã⁄ √ﬂ — „‰ ’› √ÕÌ«‰«)
    MySQL = MySQL + " OUTER APPLY ( " & vbCrLf
    MySQL = MySQL + "     SELECT TOP (1) E2.Emp_Code " & vbCrLf
    MySQL = MySQL + "     FROM dbo.TblEmployee E2 " & vbCrLf
    MySQL = MySQL + "     WHERE A.Account_Code IN ( " & vbCrLf
    MySQL = MySQL + "         ISNULL(E2.Account_code,''), ISNULL(E2.Account_code1,''), ISNULL(E2.Account_code2,''), " & vbCrLf
    MySQL = MySQL + "         ISNULL(E2.Account_code3,''), ISNULL(E2.Account_code4,''), ISNULL(E2.Account_code5,'') " & vbCrLf
    MySQL = MySQL + "     ) " & vbCrLf
    MySQL = MySQL + "     ORDER BY E2.Emp_ID " & vbCrLf
    MySQL = MySQL + " ) EmpMap " & vbCrLf

    MySQL = MySQL + " WHERE      A.Account_Code='" & StrAccountCode & "' " & vbCrLf

    If SystemOptions.IsHiddenUser = False Then
        MySQL = MySQL + " AND ISNULL(L.IsHiddenInv,0) = 0 " & vbCrLf
    End If

Else
    MySQL = MySQL & " WHERE  Account_Code in (" & StrAccountCode & ")"
End If


If detailedtransaction = 1 Then
MySQL = " SELECT        dbo.RptLedger_Sub.ChqueNum,RptLedger_Sub.order_no, dbo.RptLedger_Sub.ManualNo, dbo.RptLedger_Sub.RecordDateH, dbo.RptLedger_Sub.Double_Entry_Vouchers_ID, dbo.RptLedger_Sub.DEV_Value, "
MySQL = MySQL + "                         dbo.RptLedger_Sub.Credit_Or_Debit, dbo.ACCOUNTS.Account_Code, dbo.RptLedger_Sub.DEV_DES, dbo.RptLedger_Sub.DevDESE, dbo.RptLedger_Sub.Account_Name, dbo.RptLedger_Sub.DEV_ID_Line_No,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.NotesTypeName, dbo.RptLedger_Sub.UserID, dbo.RptLedger_Sub.UserName, dbo.RptLedger_Sub.RecordDate, dbo.RptLedger_Sub.ReceiptID, dbo.RptLedger_Sub.Notes_ID,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.OperaID, dbo.RptLedger_Sub.Transaction_Serial, dbo.RptLedger_Sub.Transaction_Date, dbo.RptLedger_Sub.Posted, dbo.RptLedger_Sub.TransactionTypeName,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.PostedDate, dbo.RptLedger_Sub.PostedUserID, dbo.RptLedger_Sub.Account_Interval_ID, dbo.RptLedger_Sub.NoteDate, dbo.RptLedger_Sub.NoteType, dbo.RptLedger_Sub.NoteSerial,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.Note_Value, dbo.RptLedger_Sub.Account_Serial, dbo.RptLedger_Sub.Account_NameEng, dbo.RptLedger_Sub.Parent_Account_Code, dbo.RptLedger_Sub.Branch,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.Sum_account, dbo.RptLedger_Sub.cost_center, dbo.RptLedger_Sub.currenct_code, dbo.RptLedger_Sub.Remark, dbo.RptLedger_Sub.note_value_by_characters, dbo.RptLedger_Sub.foxy_no,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.DEV_ID_Line_No1, dbo.RptLedger_Sub.project_id, dbo.RptLedger_Sub.NotesTypeNamee, dbo.RptLedger_Sub.TransactionEnglishName, dbo.RptLedger_Sub.opr_fullcode,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.NoteSerial1, dbo.RptLedger_Sub.branch_id, dbo.RptLedger_Sub.ActivityTypeId, dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.RptLedger_Sub.notes_all,"
MySQL = MySQL + "                         dbo.RptLedger_Sub.Transaction_ID, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transactions.Transaction_Type,"
MySQL = MySQL + "                         dbo.TransactionTypes.StockEffect, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee,"
MySQL = MySQL + "                         dbo.Transaction_Details.discountvalue , dbo.RptLedger_Sub.branch_name, dbo.RptLedger_Sub.branch_namee, dbo.RptLedger_Sub.RegionID, dbo.RptLedger_Sub.Name, dbo.RptLedger_Sub.NameE"
MySQL = MySQL + " FROM            dbo.TransactionTypes INNER JOIN"
MySQL = MySQL + "                         dbo.Transactions ON dbo.TransactionTypes.Transaction_Type = dbo.Transactions.Transaction_Type RIGHT OUTER JOIN"
MySQL = MySQL + "                         dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
MySQL = MySQL + "                         dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
MySQL = MySQL + "                         dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID RIGHT OUTER JOIN"
MySQL = MySQL + "                         dbo.RptLedger_Sub ON dbo.Transactions.Transaction_ID = dbo.RptLedger_Sub.Transaction_ID RIGHT OUTER JOIN"
MySQL = MySQL + "                         dbo.ACCOUNTS ON dbo.RptLedger_Sub.Account_Code = dbo.ACCOUNTS.Account_Code"

 MySQL = MySQL + " WHERE     (dbo.ACCOUNTS.Account_Code = '" & StrAccountCode & "') "
 If SystemOptions.IsHiddenUser = False Then
    MySQL = MySQL + " and  IsNull(RptLedger_Sub.IsHiddenInv,0) = 0 "
End If

End If


If detailedtransaction = 1 Then

    MySQL = ""
    MySQL = MySQL + " SELECT     L.ChqueNum, L.order_no, L.CBoBasedON, L.ManualNo, L.RecordDateH, L.Double_Entry_Vouchers_ID, L.DEV_Value, " & vbCrLf
    MySQL = MySQL + "            L.Credit_Or_Debit, A.Account_Code, L.DEV_DES, L.DevDESE, L.Account_Name, L.DEV_ID_Line_No, " & vbCrLf
    MySQL = MySQL + "            L.NotesTypeName, L.UserID, L.UserName, L.RecordDate, L.ReceiptID, L.Notes_ID, " & vbCrLf
    MySQL = MySQL + "            L.OperaID, L.Transaction_Serial, L.Transaction_Date, L.Posted, L.TransactionTypeName, " & vbCrLf
    MySQL = MySQL + "            L.PostedDate, L.PostedUserID, L.Account_Interval_ID, L.NoteDate, L.NoteType, L.NoteSerial, " & vbCrLf
    MySQL = MySQL + "            L.Note_Value, L.Account_Serial, L.Account_NameEng, L.Parent_Account_Code, L.Branch, " & vbCrLf
    MySQL = MySQL + "            L.Sum_account, L.cost_center, L.currenct_code, L.Remark, L.note_value_by_characters, L.foxy_no, " & vbCrLf
    MySQL = MySQL + "            L.DEV_ID_Line_No1, L.project_id, L.NotesTypeNamee, L.TransactionEnglishName, L.opr_fullcode, " & vbCrLf
    MySQL = MySQL + "            L.NoteSerial1, L.branch_id, L.ActivityTypeId, A.opening_balance, A.opening_balance_type, L.notes_all, " & vbCrLf
    MySQL = MySQL + "                CASE  WHEN l.NotesTypeName Like N'%„ﬁ»Ê÷« %'" & vbCrLf
MySQL = MySQL + "                       AND ISNULL(L.Transaction_ID,0) = 0   " & vbCrLf
MySQL = MySQL + "                  THEN CAST(-1000000000000 - ROW_NUMBER() OVER(  " & vbCrLf
MySQL = MySQL + "                                ORDER BY L.RecordDate, L.Notes_ID, L.DEV_ID_Line_No, L.DEV_ID_Line_No1   " & vbCrLf
MySQL = MySQL + "                            ) AS float)   " & vbCrLf
MySQL = MySQL + "                  Else l.Transaction_ID  " & vbCrLf
MySQL = MySQL + "                END AS Transaction_ID  , TD.Item_ID, TD.ShowQty, TD.showPrice, T.Transaction_Type, " & vbCrLf

    MySQL = MySQL + "            TT.StockEffect, I.ItemCode, I.ItemName, I.ItemNamee, TD.UnitId, U.UnitName, U.UnitNamee, " & vbCrLf
    MySQL = MySQL + "            TD.discountvalue, L.branch_name, L.branch_namee, L.RegionID, L.Name, L.NameE " & vbCrLf

    MySQL = MySQL + " FROM       dbo.RptLedger_Sub L " & vbCrLf
    MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON A.Account_Code = L.Account_Code " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.Transactions T ON T.Transaction_ID = L.Transaction_ID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.Transaction_Details TD ON TD.Transaction_ID = T.Transaction_ID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblItems I ON I.ItemID = TD.Item_ID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblUnites U ON U.UnitID = TD.UnitId " & vbCrLf

    MySQL = MySQL + " WHERE      (A.Account_Code = '" & StrAccountCode & "') " & vbCrLf

    If SystemOptions.IsHiddenUser = False Then
        MySQL = MySQL + " AND ISNULL(L.IsHiddenInv,0) = 0 " & vbCrLf
    End If

End If


'
'    If BranchID <> 0 Then
'        MySQL = MySQL & " and RptLedger_Sub.branch_id = " & BranchID 'Val(FrmAccountingReport.dcBranch.BoundText)
'    End If
'
'    If ActivityId <> 0 Then
'        MySQL = MySQL & " and dbo.RptLedger_Sub.ActivityTypeId = " & ActivityId 'Val(FrmAccountingReport.DCActivity.BoundText)
'    End If
'    If RegionID <> 0 Then
'        MySQL = MySQL & " and dbo.RptLedger_Sub.RegionID = " & RegionID 'Val(FrmAccountingReport.DCActivity.BoundText)
'    End If
'  If project_id <> 0 Then
'        MySQL = MySQL & " and dbo.RptLedger_Sub.project_id = " & project_id 'Val(FrmAccountingReport.DCActivity.BoundText)
'    End If
'
'If NotesType <> 0 Then
'        MySQL = MySQL & " and dbo.RptLedger_Sub.NoteType = " & NotesType 'Val(FrmAccountingReport.DCActivity.BoundText)
'    End If
'

If BranchID <> 0 Then
    MySQL = MySQL & " AND L.branch_id = " & BranchID
End If

If ActivityId <> 0 Then
    MySQL = MySQL & " AND L.ActivityTypeId = " & ActivityId
End If

If RegionID <> 0 Then
    MySQL = MySQL & " AND L.RegionID = " & RegionID
End If

If project_id <> 0 Then
    MySQL = MySQL & " AND L.project_id = " & project_id
End If

If notesType <> 0 Then
    MySQL = MySQL & " AND L.NoteType = " & notesType
End If

        
    
'
'    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
'        MySQL = MySQL + " and (RptLedger_Sub.RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
'    End If
'
'    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
'        MySQL = MySQL + " and RptLedger_Sub.RecordDate <=" & SQLDate(EndDate, True) & ""
'    End If
'
'    MySQL = MySQL + " or  RptLedger_Sub.RecordDate is null)"
'

Dim hasFrom As Boolean, hasTo As Boolean
hasFrom = (Not IsNull(Me.BegineDate)) And (Me.BegineDate <> Empty)
hasTo = (Not IsNull(Me.EndDate)) And (Me.EndDate <> Empty)

If hasFrom Or hasTo Then
    MySQL = MySQL & " AND ( ("  '‰› Õ „Ã„Ê⁄… «· «—ÌŒ

    If hasFrom Then
        MySQL = MySQL & " L.RecordDate >= " & SQLDate(Me.BegineDate, True) & " "
    Else
        MySQL = MySQL & " 1=1 "
    End If

    If hasTo Then
        MySQL = MySQL & " AND L.RecordDate <= " & SQLDate(Me.EndDate, True) & " "
    End If

    '‰›” „‰ÿﬁﬂ: √Ê ·Ê «· «—ÌŒ NULL
    MySQL = MySQL & " ) OR L.RecordDate IS NULL )"
End If


If detailedtransactionBasi = 1 Then

    MySQL = ""

    '=========================================================
    ' (A) «·ﬁÌÊœ „‰ RptLedger_Sub
    '=========================================================
    MySQL = MySQL + " SELECT     L.ChqueNum, L.order_no, L.CBoBasedON, L.ManualNo, L.RecordDateH, L.Double_Entry_Vouchers_ID, L.DEV_Value, " & vbCrLf
    MySQL = MySQL + "            L.Credit_Or_Debit, A.Account_Code, L.DEV_DES, L.DevDESE, L.Account_Name, L.DEV_ID_Line_No, " & vbCrLf
    MySQL = MySQL + "            L.NotesTypeName, L.UserID, L.UserName, L.RecordDate, L.ReceiptID, L.Notes_ID, " & vbCrLf
    MySQL = MySQL + "            L.OperaID, L.Transaction_Serial, L.Transaction_Date, L.Posted, L.TransactionTypeName, " & vbCrLf
    MySQL = MySQL + "            L.PostedDate, L.PostedUserID, L.Account_Interval_ID, L.NoteDate, L.NoteType, L.NoteSerial, " & vbCrLf
    MySQL = MySQL + "            L.Note_Value, L.Account_Serial, L.Account_NameEng, L.Parent_Account_Code, L.Branch, " & vbCrLf
    MySQL = MySQL + "            L.Sum_account, L.cost_center, L.currenct_code, L.Remark, L.note_value_by_characters, L.foxy_no, " & vbCrLf
    MySQL = MySQL + "            L.DEV_ID_Line_No1, L.project_id, L.NotesTypeNamee, L.TransactionEnglishName, L.opr_fullcode, " & vbCrLf
    MySQL = MySQL + "            L.NoteSerial1, L.branch_id, L.ActivityTypeId, A.opening_balance, A.opening_balance_type, L.notes_all, " & vbCrLf
    
      MySQL = MySQL + "                CASE  WHEN l.NotesTypeName Like N'%„ﬁ»Ê÷« %'" & vbCrLf
MySQL = MySQL + "                       AND ISNULL(L.Transaction_ID,0) = 0   " & vbCrLf
MySQL = MySQL + "                  THEN CAST(-1000000000000 - ROW_NUMBER() OVER(  " & vbCrLf
MySQL = MySQL + "                                ORDER BY L.RecordDate, L.Notes_ID, L.DEV_ID_Line_No, L.DEV_ID_Line_No1   " & vbCrLf
MySQL = MySQL + "                            ) AS float)   " & vbCrLf
MySQL = MySQL + "                  Else l.Transaction_ID  " & vbCrLf
MySQL = MySQL + "                END AS Transaction_ID  ," & vbCrLf

    MySQL = MySQL + "            TD.Item_ID, TD.ShowQty, TD.showPrice, T.Transaction_Type, " & vbCrLf
    MySQL = MySQL + "            TT.StockEffect, I.ItemCode, I.ItemName, I.ItemNamee, TD.UnitId, U.UnitName, U.UnitNamee, " & vbCrLf
    MySQL = MySQL + "            TD.discountvalue, L.branch_name, L.branch_namee, L.RegionID, L.Name, L.NameE " & vbCrLf

    MySQL = MySQL + " FROM       dbo.RptLedger_Sub L " & vbCrLf
    MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON A.Account_Code = L.Account_Code " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.Transactions T ON T.Transaction_ID = L.Transaction_ID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.Transaction_Details TD ON TD.Transaction_ID = T.Transaction_ID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblItems I ON I.ItemID = TD.Item_ID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblUnites U ON U.UnitID = TD.UnitId " & vbCrLf

    MySQL = MySQL + " WHERE      (A.Account_Code = '" & StrAccountCode & "') " & vbCrLf

    If SystemOptions.IsHiddenUser = False Then
        MySQL = MySQL + " AND ISNULL(L.IsHiddenInv,0) = 0 " & vbCrLf
    End If

    If BranchID <> 0 Then
        MySQL = MySQL & " AND L.branch_id = " & BranchID & vbCrLf
    End If

    If ActivityId <> 0 Then
        MySQL = MySQL & " AND L.ActivityTypeId = " & ActivityId & vbCrLf
    End If

    If RegionID <> 0 Then
        MySQL = MySQL & " AND L.RegionID = " & RegionID & vbCrLf
    End If

    If project_id <> 0 Then
        MySQL = MySQL & " AND L.project_id = " & project_id & vbCrLf
    End If

    If notesType <> 0 Then
        MySQL = MySQL & " AND L.NoteType = " & notesType & vbCrLf
    End If


    '========================
    ' ›· — «· «—ÌŒ (··ﬁÌÊœ ›ﬁÿ)
    '========================
   ' Dim hasFrom As Boolean, hasTo As Boolean
    hasFrom = (Not IsNull(Me.BegineDate)) And (Me.BegineDate <> Empty)
    hasTo = (Not IsNull(Me.EndDate)) And (Me.EndDate <> Empty)

    If hasFrom Or hasTo Then
        MySQL = MySQL & " AND ( 1=1 " & vbCrLf
        If hasFrom Then MySQL = MySQL & "     AND L.RecordDate >= " & SQLDate(Me.BegineDate, True) & vbCrLf
        If hasTo Then MySQL = MySQL & "     AND L.RecordDate <= " & SQLDate(Me.EndDate, True) & vbCrLf
        MySQL = MySQL & " ) " & vbCrLf
    End If


'    '=========================================================
'    ' UNION ALL (B) ›Ê« Ì— „»œ∆Ì… „‰ Transactions »œÊ‰ ﬁÌœ
'    '=========================================================
'    MySQL = MySQL & vbCrLf & " UNION ALL " & vbCrLf & vbCrLf
'
'    MySQL = MySQL + " SELECT     " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS ChqueNum, " & vbCrLf
'    MySQL = MySQL + "            T.order_no, T.CBoBasedON, T.ManualNo, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS RecordDateH, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS float)         AS Double_Entry_Vouchers_ID, " & vbCrLf
'
'    ' ? „»·€ «·”ÿ— „‰  ›«’Ì· «·›« Ê—… ( ›’Ì·Ì)
'    MySQL = MySQL + "            ( (ISNULL(TD.ShowQty,0) * ISNULL(TD.showPrice,0)) - ISNULL(TD.discountvalue,0) ) AS DEV_Value, " & vbCrLf
'
'    ' ? « Ã«Â „œÌ‰/œ«∆‰ (€«·»« ›Ê« Ì— «·⁄„Ì· = „œÌ‰)
'    MySQL = MySQL + "            CAST(0 AS int)             AS Credit_Or_Debit, " & vbCrLf
'
'    MySQL = MySQL + "            A.Account_Code, " & vbCrLf
'    MySQL = MySQL + "            CAST(N'›« Ê—… „»œ∆Ì… (»œÊ‰ ﬁÌœ)' AS nvarchar(4000)) AS DEV_DES, " & vbCrLf
'    MySQL = MySQL + "            CAST(N'Provisional Invoice (No Voucher)' AS nvarchar(4000)) AS DevDESE, " & vbCrLf
'    MySQL = MySQL + "            A.Account_Name, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No, " & vbCrLf
'
'    MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS NotesTypeName, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS int)          AS UserID, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS UserName, " & vbCrLf
'
'    ' ?  «—ÌŒ «·›« Ê—… Â‰« ÂÊ «··Ì Ì„À¯· RecordDate
'    MySQL = MySQL + "            T.Transaction_Date         AS RecordDate, " & vbCrLf
'
'    MySQL = MySQL + "            CAST(NULL AS int)          AS ReceiptID, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS float)        AS Notes_ID, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS int)          AS OperaID, " & vbCrLf
'
'    MySQL = MySQL + "            T.Transaction_Serial, T.Transaction_Date, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS bit)          AS Posted, " & vbCrLf
'    MySQL = MySQL + "            TT.TransactionTypeName, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS datetime)     AS PostedDate, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS int)          AS PostedUserID, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS int)          AS Account_Interval_ID, " & vbCrLf
'
'    MySQL = MySQL + "            CAST(NULL AS datetime)     AS NoteDate, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS int)          AS NoteType, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS nvarchar(50)) AS NoteSerial, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS money)        AS Note_Value, " & vbCrLf
'
'    MySQL = MySQL + "            A.Account_Serial, A.Account_NameEng, A.Parent_Account_Code, A.Branch, " & vbCrLf
'    MySQL = MySQL + "            A.Sum_account, A.cost_center, A.currenct_code, " & vbCrLf
'
'    MySQL = MySQL + "            T.Remark, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS note_value_by_characters, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS nvarchar(50)) AS foxy_no, " & vbCrLf
'
'    MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No1, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS int)          AS project_id, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS NotesTypeNamee, " & vbCrLf
'    MySQL = MySQL + "            TT.TransactionEnglishName, " & vbCrLf
'    MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS opr_fullcode, " & vbCrLf
'
'    MySQL = MySQL + "            CAST(NULL AS float)        AS NoteSerial1, " & vbCrLf
'
'    MySQL = MySQL + "            B.branch_id, B.ActivityTypeId, " & vbCrLf
'    MySQL = MySQL + "            A.opening_balance, A.opening_balance_type, " & vbCrLf
'
'    MySQL = MySQL + "            CAST(NULL AS bit)          AS notes_all, " & vbCrLf
'
'    MySQL = MySQL + "            T.Transaction_ID, TD.Item_ID, TD.ShowQty, TD.showPrice, T.Transaction_Type, " & vbCrLf
'    MySQL = MySQL + "            TT.StockEffect, I.ItemCode, I.ItemName, I.ItemNamee, TD.UnitId, U.UnitName, U.UnitNamee, " & vbCrLf
'    MySQL = MySQL + "            TD.discountvalue, B.branch_name, B.branch_namee, B.RegionID, S.Name, S.NameE " & vbCrLf
'
'    MySQL = MySQL + " FROM       dbo.Transactions T " & vbCrLf
'    MySQL = MySQL + " INNER JOIN dbo.TblCustemers C ON C.CusID = T.CusID " & vbCrLf
'    MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON A.Account_Code = C.Account_Code " & vbCrLf
'    MySQL = MySQL + " LEFT  JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type " & vbCrLf
'    MySQL = MySQL + " LEFT  JOIN dbo.Transaction_Details TD ON TD.Transaction_ID = T.Transaction_ID " & vbCrLf
'    MySQL = MySQL + " LEFT  JOIN dbo.TblItems I ON I.ItemID = TD.Item_ID " & vbCrLf
'    MySQL = MySQL + " LEFT  JOIN dbo.TblUnites U ON U.UnitID = TD.UnitId " & vbCrLf
'    MySQL = MySQL + " LEFT  JOIN dbo.TblBranchesData B ON B.branch_id = T.BranchId " & vbCrLf
'    MySQL = MySQL + " LEFT  JOIN dbo.TblSection S ON S.Id = B.RegionID " & vbCrLf
'
'    MySQL = MySQL + " WHERE      (A.Account_Code = '" & StrAccountCode & "') " & vbCrLf
'
'    ' ? ‘—ÿ: „›Ì‘ ﬁÌœ „— »ÿ »«·›« Ê—…
'    MySQL = MySQL + " AND NOT EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS DEV WHERE DEV.Transaction_ID = T.Transaction_ID) " & vbCrLf
'
'    ' ? ‰›” ›·« — «·›—Ê⁄/«·‰‘«ÿ/«·„‰ÿﬁ… («··Ì ‰ﬁœ— ‰ÿ»ﬁÂ« Â‰«)
'    If BranchID <> 0 Then
'        MySQL = MySQL & " AND T.BranchId = " & BranchID & vbCrLf
'    End If
'
'    If ActivityId <> 0 Then
'        MySQL = MySQL & " AND B.ActivityTypeId = " & ActivityId & vbCrLf
'    End If
'
'    If RegionID <> 0 Then
'        MySQL = MySQL & " AND B.RegionID = " & RegionID & vbCrLf
'    End If
'
'    ' ?? project_id Ê NotesType Â‰« „‘ „ÊÃÊœÌ‰ »‰›” «·„⁄‰Ï („‘ Notes)
'    ' ›·Ê «·„” Œœ„ „Œ «—Â„° ‰” »⁄œ «·›Ê« Ì— ⁄‘«‰ „« ‰ﬂ”—‘ «·«” ⁄·«„
'    If project_id <> 0 Then
'        MySQL = MySQL & " AND 1=0 " & vbCrLf
'    End If
'
'    If NotesType <> 0 Then
'        MySQL = MySQL & " AND 1=0 " & vbCrLf
'    End If
'
'    '========================
'    ' ›· — «· «—ÌŒ (··›Ê« Ì—)
'    '========================
'    If hasFrom Or hasTo Then
'        MySQL = MySQL & " AND ( 1=1 " & vbCrLf
'        If hasFrom Then MySQL = MySQL & "     AND T.Transaction_Date >= " & SQLDate(Me.BegineDate, True) & vbCrLf
'        If hasTo Then MySQL = MySQL & "     AND T.Transaction_Date <= " & SQLDate(Me.EndDate, True) & vbCrLf
'        MySQL = MySQL & " ) " & vbCrLf
'    End If

'=========================================================
' UNION ALL (B) ›Ê« Ì— „‰ Transactions »œÊ‰ ﬁÌœ
'   - (B1) ‰Â«∆Ì… 21 (ﬂ·Â«)
'   - (B2) „»œ∆Ì… 6 »‘—ÿ „« ÌﬂÊ‰‘ « ⁄„· „‰Â« ‰Â«∆Ì (21 Ê CBoBasedON=10)
'=========================================================
MySQL = MySQL & vbCrLf & " UNION ALL " & vbCrLf & vbCrLf

'--------------------------
' (B1) «·‰Â«∆Ì… 21
'--------------------------
MySQL = MySQL + " SELECT     " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS ChqueNum, " & vbCrLf
MySQL = MySQL + "            T.order_no, T.CBoBasedON, T.ManualNo, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS RecordDateH, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS float)         AS Double_Entry_Vouchers_ID, " & vbCrLf
MySQL = MySQL + "            ( (ISNULL(TD.ShowQty,0) * ISNULL(TD.showPrice,0)) - ISNULL(TD.discountvalue,0) ) + Isnull(td.Vat,0)  AS DEV_Value, " & vbCrLf
MySQL = MySQL + "            CAST(0 AS int)             AS Credit_Or_Debit, " & vbCrLf
MySQL = MySQL + "            A.Account_Code, " & vbCrLf
MySQL = MySQL + "           T.TransactionComment AS DEV_DES, " & vbCrLf
MySQL = MySQL + "            CAST(N'Sales Invoice (No Voucher)' AS nvarchar(4000)) AS DevDESE, " & vbCrLf
MySQL = MySQL + "            A.Account_Name, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No, " & vbCrLf
MySQL = MySQL + "            CAST('›« Ê—… „»Ì⁄«  „' AS nvarchar(200)) AS NotesTypeName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS UserID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS UserName, " & vbCrLf
MySQL = MySQL + "            T.Transaction_Date         AS RecordDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS ReceiptID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS float)        AS Notes_ID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS OperaID, " & vbCrLf
MySQL = MySQL + "            T.Transaction_Serial, T.Transaction_Date, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS bit)          AS Posted, " & vbCrLf
MySQL = MySQL + "            TT.TransactionTypeName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS datetime)     AS PostedDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS PostedUserID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS Account_Interval_ID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS datetime)     AS NoteDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS NoteType, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50)) AS NoteSerial, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS money)        AS Note_Value, " & vbCrLf
MySQL = MySQL + "            A.Account_Serial, A.Account_NameEng, A.Parent_Account_Code, A.Branch, " & vbCrLf
MySQL = MySQL + "            A.Sum_account, A.cost_center, A.currenct_code, " & vbCrLf
MySQL = MySQL + "            T.Remark, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS note_value_by_characters, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50)) AS foxy_no, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No1, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS project_id, " & vbCrLf
MySQL = MySQL + "            CAST('Invoice' AS nvarchar(200)) AS NotesTypeNamee, " & vbCrLf
MySQL = MySQL + "            TT.TransactionEnglishName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS opr_fullcode, " & vbCrLf
MySQL = MySQL + "            CAST(T.NoteSerial1 AS float)        AS NoteSerial1, " & vbCrLf
MySQL = MySQL + "            B.branch_id, B.ActivityTypeId, " & vbCrLf
MySQL = MySQL + "            A.opening_balance, A.opening_balance_type, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS bit)          AS notes_all, " & vbCrLf
MySQL = MySQL + "            T.Transaction_ID, TD.Item_ID, TD.ShowQty, TD.showPrice, T.Transaction_Type, " & vbCrLf
MySQL = MySQL + "            TT.StockEffect, I.ItemCode, I.ItemName, I.ItemNamee, TD.UnitId, U.UnitName, U.UnitNamee, " & vbCrLf
MySQL = MySQL + "            TD.discountvalue, B.branch_name, B.branch_namee, B.RegionID, S.Name, S.NameE " & vbCrLf

MySQL = MySQL + " FROM       dbo.Transactions T " & vbCrLf
MySQL = MySQL + " INNER JOIN dbo.TblCustemers C ON C.CusID = T.CusID " & vbCrLf
MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON A.Account_Code = C.Account_Code " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.Transaction_Details TD ON TD.Transaction_ID = T.Transaction_ID " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblItems I ON I.ItemID = TD.Item_ID " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblUnites U ON U.UnitID = TD.UnitId " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblBranchesData B ON B.branch_id = T.BranchId " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblSection S ON S.Id = B.RegionID " & vbCrLf

MySQL = MySQL + " WHERE      (A.Account_Code = '" & StrAccountCode & "') " & vbCrLf
MySQL = MySQL + " AND T.Transaction_Type = 21 " & vbCrLf
MySQL = MySQL + " AND NOT EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS DEV WHERE DEV.Transaction_ID = T.Transaction_ID) " & vbCrLf

If BranchID <> 0 Then MySQL = MySQL & " AND T.BranchId = " & BranchID & vbCrLf
If ActivityId <> 0 Then MySQL = MySQL & " AND B.ActivityTypeId = " & ActivityId & vbCrLf
If RegionID <> 0 Then MySQL = MySQL & " AND B.RegionID = " & RegionID & vbCrLf

If hasFrom Or hasTo Then
    MySQL = MySQL & " AND ( 1=1 " & vbCrLf
    If hasFrom Then MySQL = MySQL & "     AND T.Transaction_Date >= " & SQLDate(Me.BegineDate, True) & vbCrLf
    If hasTo Then MySQL = MySQL & "     AND T.Transaction_Date <= " & SQLDate(Me.EndDate, True) & vbCrLf
    MySQL = MySQL & " ) " & vbCrLf
End If


'--------------------------
' UNION ALL (B2) «·„»œ∆Ì… 6 (›ﬁÿ ·Ê ·”Â „« « ⁄„·‘ „‰Â« ‰Â«∆Ì)
'--------------------------
MySQL = MySQL & vbCrLf & " UNION ALL " & vbCrLf & vbCrLf

MySQL = MySQL + " SELECT     " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS ChqueNum, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS float)        AS order_no, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS CBoBasedON, " & vbCrLf
MySQL = MySQL + "            T.ManualNo, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS RecordDateH, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS float)         AS Double_Entry_Vouchers_ID, " & vbCrLf
MySQL = MySQL + "            ( (ISNULL(TD.ShowQty,0) * ISNULL(TD.showPrice,0)) - ISNULL(TD.discountvalue,0) ) + Isnull(td.Vat,0)  AS DEV_Value, " & vbCrLf
MySQL = MySQL + "            CAST(0 AS int)             AS Credit_Or_Debit, " & vbCrLf
MySQL = MySQL + "            A.Account_Code, " & vbCrLf
MySQL = MySQL + "            t.TransactionComment AS DEV_DES, " & vbCrLf
MySQL = MySQL + "            CAST(N'Provisional Invoice (No Final Yet)' AS nvarchar(4000)) AS DevDESE, " & vbCrLf
MySQL = MySQL + "            A.Account_Name, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No, " & vbCrLf
MySQL = MySQL + "            CAST('›« Ê—… „»Ì⁄«  „' AS nvarchar(200)) AS NotesTypeName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS UserID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS UserName, " & vbCrLf
MySQL = MySQL + "            T.Transaction_Date         AS RecordDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS ReceiptID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS float)        AS Notes_ID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS OperaID, " & vbCrLf
MySQL = MySQL + "            T.Transaction_Serial, T.Transaction_Date, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS bit)          AS Posted, " & vbCrLf
MySQL = MySQL + "            TT.TransactionTypeName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS datetime)     AS PostedDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS PostedUserID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS Account_Interval_ID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS datetime)     AS NoteDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS NoteType, " & vbCrLf
MySQL = MySQL + "            CAST(t.Noteserial AS nvarchar(50)) AS NoteSerial, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS money)        AS Note_Value, " & vbCrLf
MySQL = MySQL + "            A.Account_Serial, A.Account_NameEng, A.Parent_Account_Code, A.Branch, " & vbCrLf
MySQL = MySQL + "            A.Sum_account, A.cost_center, A.currenct_code, " & vbCrLf
MySQL = MySQL + "            T.Remark, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS note_value_by_characters, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50)) AS foxy_no, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No1, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS project_id, " & vbCrLf
MySQL = MySQL + "            CAST('Invoice' AS nvarchar(200)) AS NotesTypeNamee, " & vbCrLf
MySQL = MySQL + "            TT.TransactionEnglishName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS opr_fullcode, " & vbCrLf

'? Â‰« —ﬁ„ «·„»œ∆Ì… «·ÕﬁÌﬁÌ
MySQL = MySQL + "            T.NoteSerial1              AS NoteSerial1, " & vbCrLf

MySQL = MySQL + "            B.branch_id, B.ActivityTypeId, " & vbCrLf
MySQL = MySQL + "            A.opening_balance, A.opening_balance_type, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS bit)          AS notes_all, " & vbCrLf
MySQL = MySQL + "            T.Transaction_ID, TD.Item_ID, TD.ShowQty, TD.showPrice, T.Transaction_Type, " & vbCrLf
MySQL = MySQL + "            TT.StockEffect, I.ItemCode, I.ItemName, I.ItemNamee, TD.UnitId, U.UnitName, U.UnitNamee, " & vbCrLf
MySQL = MySQL + "            TD.discountvalue, B.branch_name, B.branch_namee, B.RegionID, S.Name, S.NameE " & vbCrLf

MySQL = MySQL + " FROM       dbo.Transactions T " & vbCrLf
MySQL = MySQL + " INNER JOIN dbo.TblCustemers C ON C.CusID = T.CusID " & vbCrLf
MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON A.Account_Code = C.Account_Code " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.Transaction_Details TD ON TD.Transaction_ID = T.Transaction_ID " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblItems I ON I.ItemID = TD.Item_ID " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblUnites U ON U.UnitID = TD.UnitId " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblBranchesData B ON B.branch_id = T.BranchId " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblSection S ON S.Id = B.RegionID " & vbCrLf

MySQL = MySQL + " WHERE      (A.Account_Code = '" & StrAccountCode & "') " & vbCrLf
MySQL = MySQL + " AND T.Transaction_Type = 6 " & vbCrLf
MySQL = MySQL + " AND NOT EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS DEV WHERE DEV.Transaction_ID = T.Transaction_ID) " & vbCrLf

'? «” »⁄«œ «·„»œ∆Ì… ·Ê « ⁄„· „‰Â« ‰Â«∆Ì (21 Ê CBoBasedON=10) Ê«·—»ÿ order_no = NoteSerial1
MySQL = MySQL + " AND NOT EXISTS ( " & vbCrLf
MySQL = MySQL + "     SELECT 1 " & vbCrLf
MySQL = MySQL + "     FROM dbo.Transactions T2 " & vbCrLf
MySQL = MySQL + "     WHERE T2.Transaction_Type = 21 " & vbCrLf
MySQL = MySQL + "       AND ISNULL(T2.CBoBasedON,0) = 10 " & vbCrLf
MySQL = MySQL + "       AND T2.order_no = T.NoteSerial1 " & vbCrLf
MySQL = MySQL + " ) " & vbCrLf

If BranchID <> 0 Then MySQL = MySQL & " AND T.BranchId = " & BranchID & vbCrLf
If ActivityId <> 0 Then MySQL = MySQL & " AND B.ActivityTypeId = " & ActivityId & vbCrLf
If RegionID <> 0 Then MySQL = MySQL & " AND B.RegionID = " & RegionID & vbCrLf

If hasFrom Or hasTo Then
    MySQL = MySQL & " AND ( 1=1 " & vbCrLf
    If hasFrom Then MySQL = MySQL & "     AND T.Transaction_Date >= " & SQLDate(Me.BegineDate, True) & vbCrLf
    If hasTo Then MySQL = MySQL & "     AND T.Transaction_Date <= " & SQLDate(Me.EndDate, True) & vbCrLf
    MySQL = MySQL & " ) " & vbCrLf
End If

End If



If detailedtransactionBasi = 1 Then

    MySQL = ""

    '=========================================================
    ' (A) «·ﬁÌÊœ „‰ RptLedger_Sub
    '=========================================================
    MySQL = MySQL + " SELECT     L.ChqueNum, L.order_no, L.CBoBasedON, L.ManualNo, L.RecordDateH, L.Double_Entry_Vouchers_ID, L.DEV_Value, " & vbCrLf
    MySQL = MySQL + "            L.Credit_Or_Debit, A.Account_Code, L.DEV_DES, L.DevDESE, L.Account_Name, L.DEV_ID_Line_No, " & vbCrLf
    MySQL = MySQL + "            L.NotesTypeName, L.UserID, L.UserName, L.RecordDate, L.ReceiptID, L.Notes_ID, " & vbCrLf
    MySQL = MySQL + "            L.OperaID, L.Transaction_Serial, L.Transaction_Date, L.Posted, L.TransactionTypeName, " & vbCrLf
    MySQL = MySQL + "            L.PostedDate, L.PostedUserID, L.Account_Interval_ID, L.NoteDate, L.NoteType, L.NoteSerial, " & vbCrLf
    MySQL = MySQL + "            L.Note_Value, L.Account_Serial, L.Account_NameEng, L.Parent_Account_Code, L.Branch, " & vbCrLf
    MySQL = MySQL + "            L.Sum_account, L.cost_center, L.currenct_code, L.Remark, L.note_value_by_characters, L.foxy_no, " & vbCrLf
    MySQL = MySQL + "            L.DEV_ID_Line_No1, L.project_id, L.NotesTypeNamee, L.TransactionEnglishName, L.opr_fullcode, " & vbCrLf
    MySQL = MySQL + "            L.NoteSerial1, L.branch_id, L.ActivityTypeId, A.opening_balance, A.opening_balance_type, L.notes_all, " & vbCrLf
    
    MySQL = MySQL + "                CASE  WHEN l.NotesTypeName Like N'%„ﬁ»Ê÷« %'" & vbCrLf
    MySQL = MySQL + "                       AND ISNULL(L.Transaction_ID,0) = 0   " & vbCrLf
    MySQL = MySQL + "                  THEN CAST(-1000000000000 - ROW_NUMBER() OVER(  " & vbCrLf
    MySQL = MySQL + "                                ORDER BY L.RecordDate, L.Notes_ID, L.DEV_ID_Line_No, L.DEV_ID_Line_No1   " & vbCrLf
    MySQL = MySQL + "                            ) AS float)   " & vbCrLf
    MySQL = MySQL + "                  Else l.Transaction_ID  " & vbCrLf
    MySQL = MySQL + "                END AS Transaction_ID  ," & vbCrLf

    MySQL = MySQL + "            TD.Item_ID, TD.ShowQty, TD.showPrice, T.Transaction_Type, " & vbCrLf
    MySQL = MySQL + "            TT.StockEffect, I.ItemCode, I.ItemName, I.ItemNamee, TD.UnitId, U.UnitName, U.UnitNamee, " & vbCrLf
    MySQL = MySQL + "            TD.discountvalue, L.branch_name, L.branch_namee, L.RegionID, L.Name, L.NameE " & vbCrLf

    MySQL = MySQL + " FROM       dbo.RptLedger_Sub L " & vbCrLf
    MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON A.Account_Code = L.Account_Code " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.Transactions T ON T.Transaction_ID = L.Transaction_ID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.Transaction_Details TD ON TD.Transaction_ID = T.Transaction_ID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblItems I ON I.ItemID = TD.Item_ID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblUnites U ON U.UnitID = TD.UnitId " & vbCrLf

    MySQL = MySQL + " WHERE      (A.Account_Code = '" & StrAccountCode & "') " & vbCrLf

    If SystemOptions.IsHiddenUser = False Then
        MySQL = MySQL + " AND ISNULL(L.IsHiddenInv,0) = 0 " & vbCrLf
    End If

    If BranchID <> 0 Then
        MySQL = MySQL & " AND L.branch_id = " & BranchID & vbCrLf
    End If

    If ActivityId <> 0 Then
        MySQL = MySQL & " AND L.ActivityTypeId = " & ActivityId & vbCrLf
    End If

    If RegionID <> 0 Then
        MySQL = MySQL & " AND L.RegionID = " & RegionID & vbCrLf
    End If

    If project_id <> 0 Then
        MySQL = MySQL & " AND L.project_id = " & project_id & vbCrLf
    End If

    If notesType <> 0 Then
        MySQL = MySQL & " AND L.NoteType = " & notesType & vbCrLf
    End If

    hasFrom = (Not IsNull(Me.BegineDate)) And (Me.BegineDate <> Empty)
    hasTo = (Not IsNull(Me.EndDate)) And (Me.EndDate <> Empty)

    If hasFrom Or hasTo Then
        MySQL = MySQL & " AND ( 1=1 " & vbCrLf
        If hasFrom Then MySQL = MySQL & "     AND L.RecordDate >= " & SQLDate(Me.BegineDate, True) & vbCrLf
        If hasTo Then MySQL = MySQL & "     AND L.RecordDate <= " & SQLDate(Me.EndDate, True) & vbCrLf
        MySQL = MySQL & " ) " & vbCrLf
    End If

'=========================================================
' UNION ALL (B) ›Ê« Ì— „‰ Transactions »œÊ‰ ﬁÌœ
'   - (B1) ‰Â«∆Ì… 21 (ﬂ·Â«)
'   - (B2) „»œ∆Ì… 6 »‘—ÿ „« ÌﬂÊ‰‘ « ⁄„· „‰Â« ‰Â«∆Ì (21 Ê CBoBasedON=10)
'=========================================================
MySQL = MySQL & vbCrLf & " UNION ALL " & vbCrLf & vbCrLf

'--------------------------
' (B1) «·‰Â«∆Ì… 21
'--------------------------
MySQL = MySQL + " SELECT     " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS ChqueNum, " & vbCrLf
MySQL = MySQL + "            T.order_no, T.CBoBasedON, T.ManualNo, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS RecordDateH, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS float)         AS Double_Entry_Vouchers_ID, " & vbCrLf
MySQL = MySQL + "            ISNULL(( " & vbCrLf
MySQL = MySQL + "                SELECT SUM( " & vbCrLf
MySQL = MySQL + "                       ((ISNULL(TD2.ShowQty,0) * ISNULL(TD2.showPrice,0)) - ISNULL(TD2.discountvalue,0)) " & vbCrLf
MySQL = MySQL + "                       + ISNULL(TD2.Vat,0) " & vbCrLf
MySQL = MySQL + "                ) " & vbCrLf
MySQL = MySQL + "                FROM dbo.Transaction_Details TD2 " & vbCrLf
MySQL = MySQL + "                WHERE TD2.Transaction_ID = T.Transaction_ID " & vbCrLf
MySQL = MySQL + "            ),0) AS DEV_Value, " & vbCrLf
MySQL = MySQL + "            CAST(0 AS int)             AS Credit_Or_Debit, " & vbCrLf
MySQL = MySQL + "            A.Account_Code, " & vbCrLf
MySQL = MySQL + "           T.TransactionComment AS DEV_DES, " & vbCrLf
MySQL = MySQL + "            CAST(N'Sales Invoice (No Voucher)' AS nvarchar(4000)) AS DevDESE, " & vbCrLf
MySQL = MySQL + "            A.Account_Name, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No, " & vbCrLf
MySQL = MySQL + "            CAST('›« Ê—… „»Ì⁄«  „' AS nvarchar(200)) AS NotesTypeName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS UserID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS UserName, " & vbCrLf
MySQL = MySQL + "            T.Transaction_Date         AS RecordDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS ReceiptID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS float)        AS Notes_ID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS OperaID, " & vbCrLf
MySQL = MySQL + "            T.Transaction_Serial, T.Transaction_Date, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS bit)          AS Posted, " & vbCrLf
MySQL = MySQL + "            TT.TransactionTypeName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS datetime)     AS PostedDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS PostedUserID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS Account_Interval_ID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS datetime)     AS NoteDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS NoteType, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50)) AS NoteSerial, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS money)        AS Note_Value, " & vbCrLf
MySQL = MySQL + "            A.Account_Serial, A.Account_NameEng, A.Parent_Account_Code, A.Branch, " & vbCrLf
MySQL = MySQL + "            A.Sum_account, A.cost_center, A.currenct_code, " & vbCrLf
MySQL = MySQL + "            T.Remark, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS note_value_by_characters, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50)) AS foxy_no, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No1, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS project_id, " & vbCrLf
MySQL = MySQL + "            CAST('Invoice' AS nvarchar(200)) AS NotesTypeNamee, " & vbCrLf
MySQL = MySQL + "            TT.TransactionEnglishName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS opr_fullcode, " & vbCrLf
MySQL = MySQL + "            CAST(T.NoteSerial1 AS float)        AS NoteSerial1, " & vbCrLf
MySQL = MySQL + "            B.branch_id, B.ActivityTypeId, " & vbCrLf
MySQL = MySQL + "            A.opening_balance, A.opening_balance_type, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS bit)          AS notes_all, " & vbCrLf
MySQL = MySQL + "            T.Transaction_ID, TD.Item_ID, TD.ShowQty, TD.showPrice, T.Transaction_Type, " & vbCrLf
MySQL = MySQL + "            TT.StockEffect, I.ItemCode, I.ItemName, I.ItemNamee, TD.UnitId, U.UnitName, U.UnitNamee, " & vbCrLf
MySQL = MySQL + "            TD.discountvalue, B.branch_name, B.branch_namee, B.RegionID, S.Name, S.NameE " & vbCrLf

MySQL = MySQL + " FROM       dbo.Transactions T " & vbCrLf
MySQL = MySQL + " INNER JOIN dbo.TblCustemers C ON C.CusID = T.CusID " & vbCrLf
MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON A.Account_Code = C.Account_Code " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.Transaction_Details TD ON TD.Transaction_ID = T.Transaction_ID " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblItems I ON I.ItemID = TD.Item_ID " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblUnites U ON U.UnitID = TD.UnitId " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblBranchesData B ON B.branch_id = T.BranchId " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblSection S ON S.Id = B.RegionID " & vbCrLf

MySQL = MySQL + " WHERE      (A.Account_Code = '" & StrAccountCode & "') " & vbCrLf
MySQL = MySQL + " AND T.Transaction_Type = 21 " & vbCrLf
MySQL = MySQL + " AND NOT EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS DEV WHERE DEV.Transaction_ID = T.Transaction_ID) " & vbCrLf

If BranchID <> 0 Then MySQL = MySQL & " AND T.BranchId = " & BranchID & vbCrLf
If ActivityId <> 0 Then MySQL = MySQL & " AND B.ActivityTypeId = " & ActivityId & vbCrLf
If RegionID <> 0 Then MySQL = MySQL & " AND B.RegionID = " & RegionID & vbCrLf

If hasFrom Or hasTo Then
    MySQL = MySQL & " AND ( 1=1 " & vbCrLf
    If hasFrom Then MySQL = MySQL & "     AND T.Transaction_Date >= " & SQLDate(Me.BegineDate, True) & vbCrLf
    If hasTo Then MySQL = MySQL & "     AND T.Transaction_Date <= " & SQLDate(Me.EndDate, True) & vbCrLf
    MySQL = MySQL & " ) " & vbCrLf
End If

'--------------------------
' UNION ALL (B2) «·„»œ∆Ì… 6 (›ﬁÿ ·Ê ·”Â „« « ⁄„·‘ „‰Â« ‰Â«∆Ì)
'--------------------------
MySQL = MySQL & vbCrLf & " UNION ALL " & vbCrLf & vbCrLf

MySQL = MySQL + " SELECT     " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS ChqueNum, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS float)        AS order_no, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS CBoBasedON, " & vbCrLf
MySQL = MySQL + "            T.ManualNo, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS RecordDateH, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS float)         AS Double_Entry_Vouchers_ID, " & vbCrLf
MySQL = MySQL + "            ISNULL(( " & vbCrLf
MySQL = MySQL + "                SELECT SUM( " & vbCrLf
MySQL = MySQL + "                       ((ISNULL(TD2.ShowQty,0) * ISNULL(TD2.showPrice,0)) - ISNULL(TD2.discountvalue,0)) " & vbCrLf
MySQL = MySQL + "                       + ISNULL(TD2.Vat,0) " & vbCrLf
MySQL = MySQL + "                ) " & vbCrLf
MySQL = MySQL + "                FROM dbo.Transaction_Details TD2 " & vbCrLf
MySQL = MySQL + "                WHERE TD2.Transaction_ID = T.Transaction_ID " & vbCrLf
MySQL = MySQL + "            ),0) AS DEV_Value, " & vbCrLf
MySQL = MySQL + "            CAST(0 AS int)             AS Credit_Or_Debit, " & vbCrLf
MySQL = MySQL + "            A.Account_Code, " & vbCrLf
MySQL = MySQL + "            t.TransactionComment AS DEV_DES, " & vbCrLf
MySQL = MySQL + "            CAST(N'Provisional Invoice (No Final Yet)' AS nvarchar(4000)) AS DevDESE, " & vbCrLf
MySQL = MySQL + "            A.Account_Name, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No, " & vbCrLf
MySQL = MySQL + "            CAST('›« Ê—… „»Ì⁄«  „' AS nvarchar(200)) AS NotesTypeName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS UserID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS UserName, " & vbCrLf
MySQL = MySQL + "            T.Transaction_Date         AS RecordDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS ReceiptID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS float)        AS Notes_ID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS OperaID, " & vbCrLf
MySQL = MySQL + "            T.Transaction_Serial, T.Transaction_Date, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS bit)          AS Posted, " & vbCrLf
MySQL = MySQL + "            TT.TransactionTypeName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS datetime)     AS PostedDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS PostedUserID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS Account_Interval_ID, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS datetime)     AS NoteDate, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS NoteType, " & vbCrLf
MySQL = MySQL + "            CAST(t.Noteserial AS nvarchar(50)) AS NoteSerial, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS money)        AS Note_Value, " & vbCrLf
MySQL = MySQL + "            A.Account_Serial, A.Account_NameEng, A.Parent_Account_Code, A.Branch, " & vbCrLf
MySQL = MySQL + "            A.Sum_account, A.cost_center, A.currenct_code, " & vbCrLf
MySQL = MySQL + "            T.Remark, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS note_value_by_characters, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(50)) AS foxy_no, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No1, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS int)          AS project_id, " & vbCrLf
MySQL = MySQL + "            CAST('Invoice' AS nvarchar(200)) AS NotesTypeNamee, " & vbCrLf
MySQL = MySQL + "            TT.TransactionEnglishName, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS opr_fullcode, " & vbCrLf
MySQL = MySQL + "            T.NoteSerial1              AS NoteSerial1, " & vbCrLf
MySQL = MySQL + "            B.branch_id, B.ActivityTypeId, " & vbCrLf
MySQL = MySQL + "            A.opening_balance, A.opening_balance_type, " & vbCrLf
MySQL = MySQL + "            CAST(NULL AS bit)          AS notes_all, " & vbCrLf
MySQL = MySQL + "            T.Transaction_ID, TD.Item_ID, TD.ShowQty, TD.showPrice, T.Transaction_Type, " & vbCrLf
MySQL = MySQL + "            TT.StockEffect, I.ItemCode, I.ItemName, I.ItemNamee, TD.UnitId, U.UnitName, U.UnitNamee, " & vbCrLf
MySQL = MySQL + "            TD.discountvalue, B.branch_name, B.branch_namee, B.RegionID, S.Name, S.NameE " & vbCrLf

MySQL = MySQL + " FROM       dbo.Transactions T " & vbCrLf
MySQL = MySQL + " INNER JOIN dbo.TblCustemers C ON C.CusID = T.CusID " & vbCrLf
MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON A.Account_Code = C.Account_Code " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.Transaction_Details TD ON TD.Transaction_ID = T.Transaction_ID " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblItems I ON I.ItemID = TD.Item_ID " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblUnites U ON U.UnitID = TD.UnitId " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblBranchesData B ON B.branch_id = T.BranchId " & vbCrLf
MySQL = MySQL + " LEFT  JOIN dbo.TblSection S ON S.Id = B.RegionID " & vbCrLf

MySQL = MySQL + " WHERE      (A.Account_Code = '" & StrAccountCode & "') " & vbCrLf
MySQL = MySQL + " AND T.Transaction_Type = 6 " & vbCrLf
MySQL = MySQL + " AND NOT EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS DEV WHERE DEV.Transaction_ID = T.Transaction_ID) " & vbCrLf

MySQL = MySQL + " AND NOT EXISTS ( " & vbCrLf
MySQL = MySQL + "     SELECT 1 " & vbCrLf
MySQL = MySQL + "     FROM dbo.Transactions T2 " & vbCrLf
MySQL = MySQL + "     WHERE T2.Transaction_Type = 21 " & vbCrLf
MySQL = MySQL + "       AND ISNULL(T2.CBoBasedON,0) = 10 " & vbCrLf
MySQL = MySQL + "       AND T2.order_no = T.NoteSerial1 " & vbCrLf
MySQL = MySQL + " ) " & vbCrLf

If BranchID <> 0 Then MySQL = MySQL & " AND T.BranchId = " & BranchID & vbCrLf
If ActivityId <> 0 Then MySQL = MySQL & " AND B.ActivityTypeId = " & ActivityId & vbCrLf
If RegionID <> 0 Then MySQL = MySQL & " AND B.RegionID = " & RegionID & vbCrLf

If hasFrom Or hasTo Then
    MySQL = MySQL & " AND ( 1=1 " & vbCrLf
    If hasFrom Then MySQL = MySQL & "     AND T.Transaction_Date >= " & SQLDate(Me.BegineDate, True) & vbCrLf
    If hasTo Then MySQL = MySQL & "     AND T.Transaction_Date <= " & SQLDate(Me.EndDate, True) & vbCrLf
    MySQL = MySQL & " ) " & vbCrLf
End If
MySQL = MySQL & vbCrLf & " ORDER BY RecordDate, Transaction_Date, NoteSerial, Transaction_ID " & vbCrLf
End If

   If detailedtransactionBasi = 1 Then

    MySQL = ""

    '=========================================================
    ' (A) «·ﬁÌÊœ „‰ RptLedger_Sub
    '   - »œÊ‰ √’‰«›
    '=========================================================
    MySQL = MySQL + " SELECT     L.ChqueNum, L.order_no, L.CBoBasedON, L.ManualNo, L.RecordDateH, L.Double_Entry_Vouchers_ID, L.DEV_Value, " & vbCrLf
    MySQL = MySQL + "            L.Credit_Or_Debit, A.Account_Code, L.DEV_DES, L.DevDESE, L.Account_Name, L.DEV_ID_Line_No, " & vbCrLf
    MySQL = MySQL + "            L.NotesTypeName, L.UserID, L.UserName, L.RecordDate, L.ReceiptID, L.Notes_ID, " & vbCrLf
    MySQL = MySQL + "            L.OperaID, L.Transaction_Serial, L.Transaction_Date, L.Posted, L.TransactionTypeName, " & vbCrLf
    MySQL = MySQL + "            L.PostedDate, L.PostedUserID, L.Account_Interval_ID, L.NoteDate, L.NoteType, L.NoteSerial, " & vbCrLf
    MySQL = MySQL + "            L.Note_Value, L.Account_Serial, L.Account_NameEng, L.Parent_Account_Code, L.Branch, " & vbCrLf
    MySQL = MySQL + "            L.Sum_account, L.cost_center, L.currenct_code, L.Remark, L.note_value_by_characters, L.foxy_no, " & vbCrLf
    MySQL = MySQL + "            L.DEV_ID_Line_No1, L.project_id, L.NotesTypeNamee, L.TransactionEnglishName, L.opr_fullcode, " & vbCrLf
    MySQL = MySQL + "            L.NoteSerial1, L.branch_id, L.ActivityTypeId, A.opening_balance, A.opening_balance_type, L.notes_all, " & vbCrLf

    MySQL = MySQL + "            CASE WHEN L.NotesTypeName Like N'%„ﬁ»Ê÷« %'" & vbCrLf
    MySQL = MySQL + "                      AND ISNULL(L.Transaction_ID,0) = 0 " & vbCrLf
    MySQL = MySQL + "                 THEN CAST(-1000000000000 - ROW_NUMBER() OVER(ORDER BY L.RecordDate, L.Notes_ID, L.DEV_ID_Line_No, L.DEV_ID_Line_No1) AS float) " & vbCrLf
    MySQL = MySQL + "                 ELSE L.Transaction_ID " & vbCrLf
    MySQL = MySQL + "            END AS Transaction_ID, " & vbCrLf

    MySQL = MySQL + "            CAST(NULL AS int)            AS Item_ID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS float)          AS ShowQty, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS money)          AS showPrice, " & vbCrLf
    MySQL = MySQL + "            T.Transaction_Type, " & vbCrLf
    MySQL = MySQL + "            TT.StockEffect, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(50))   AS ItemCode, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS ItemName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS ItemNamee, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)            AS UnitId, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(200))  AS UnitName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(200))  AS UnitNamee, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS money)          AS discountvalue, " & vbCrLf
    MySQL = MySQL + "            L.branch_name, L.branch_namee, L.RegionID, L.Name, L.NameE " & vbCrLf

    MySQL = MySQL + " FROM       dbo.RptLedger_Sub L " & vbCrLf
    MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON A.Account_Code = L.Account_Code " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.Transactions T ON T.Transaction_ID = L.Transaction_ID " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type " & vbCrLf

    MySQL = MySQL + " WHERE      (A.Account_Code = '" & StrAccountCode & "') " & vbCrLf

    If SystemOptions.IsHiddenUser = False Then
        MySQL = MySQL + " AND ISNULL(L.IsHiddenInv,0) = 0 " & vbCrLf
    End If

    If BranchID <> 0 Then
        MySQL = MySQL & " AND L.branch_id = " & BranchID & vbCrLf
    End If

    If ActivityId <> 0 Then
        MySQL = MySQL & " AND L.ActivityTypeId = " & ActivityId & vbCrLf
    End If

    If RegionID <> 0 Then
        MySQL = MySQL & " AND L.RegionID = " & RegionID & vbCrLf
    End If

    If project_id <> 0 Then
        MySQL = MySQL & " AND L.project_id = " & project_id & vbCrLf
    End If

    If notesType <> 0 Then
        MySQL = MySQL & " AND L.NoteType = " & notesType & vbCrLf
    End If

    hasFrom = (Not IsNull(Me.BegineDate)) And (Me.BegineDate <> Empty)
    hasTo = (Not IsNull(Me.EndDate)) And (Me.EndDate <> Empty)

    If hasFrom Or hasTo Then
        MySQL = MySQL & " AND ( 1=1 " & vbCrLf
        If hasFrom Then MySQL = MySQL & "     AND L.RecordDate >= " & SQLDate(Me.BegineDate, True) & vbCrLf
        If hasTo Then MySQL = MySQL & "     AND L.RecordDate <= " & SQLDate(Me.EndDate, True) & vbCrLf
        MySQL = MySQL & " ) " & vbCrLf
    End If

    '=========================================================
    ' UNION ALL (B) ›Ê« Ì— „‰ Transactions »œÊ‰ ﬁÌœ
    '   - (B1) ‰Â«∆Ì… 21
    '   - (B2) „»œ∆Ì… 6 »‘—ÿ „« ÌﬂÊ‰‘ « ⁄„· „‰Â« ‰Â«∆Ì
    '   - »œÊ‰ √’‰«›
    '=========================================================
    MySQL = MySQL & vbCrLf & " UNION ALL " & vbCrLf & vbCrLf

    '--------------------------
    ' (B1) «·‰Â«∆Ì… 21
    '--------------------------
    MySQL = MySQL + " SELECT     " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS ChqueNum, " & vbCrLf
    MySQL = MySQL + "            T.order_no, T.CBoBasedON, T.ManualNo, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS RecordDateH, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS float)         AS Double_Entry_Vouchers_ID, " & vbCrLf
    MySQL = MySQL + "            ISNULL(( " & vbCrLf
    MySQL = MySQL + "                SELECT SUM( " & vbCrLf
    MySQL = MySQL + "                       ((ISNULL(TD2.ShowQty,0) * ISNULL(TD2.showPrice,0)) - ISNULL(TD2.discountvalue,0)) " & vbCrLf
    MySQL = MySQL + "                       + ISNULL(TD2.Vat,0) " & vbCrLf
    MySQL = MySQL + "                ) " & vbCrLf
    MySQL = MySQL + "                FROM dbo.Transaction_Details TD2 " & vbCrLf
    MySQL = MySQL + "                WHERE TD2.Transaction_ID = T.Transaction_ID " & vbCrLf
    MySQL = MySQL + "            ),0) AS DEV_Value, " & vbCrLf
    MySQL = MySQL + "            CAST(0 AS int)             AS Credit_Or_Debit, " & vbCrLf
    MySQL = MySQL + "            A.Account_Code, " & vbCrLf
    MySQL = MySQL + "            T.TransactionComment AS DEV_DES, " & vbCrLf
    MySQL = MySQL + "            CAST(N'Sales Invoice (No Voucher)' AS nvarchar(4000)) AS DevDESE, " & vbCrLf
    MySQL = MySQL + "            A.Account_Name, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No, " & vbCrLf
    MySQL = MySQL + "            CAST(N'›« Ê—… „»Ì⁄« ' AS nvarchar(200)) AS NotesTypeName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS UserID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS UserName, " & vbCrLf
    MySQL = MySQL + "            T.Transaction_Date         AS RecordDate, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS ReceiptID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS float)        AS Notes_ID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS OperaID, " & vbCrLf
    MySQL = MySQL + "            T.Transaction_Serial, T.Transaction_Date, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS bit)          AS Posted, " & vbCrLf
    MySQL = MySQL + "            TT.TransactionTypeName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS datetime)     AS PostedDate, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS PostedUserID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS Account_Interval_ID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS datetime)     AS NoteDate, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS NoteType, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(50)) AS NoteSerial, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS money)        AS Note_Value, " & vbCrLf
    MySQL = MySQL + "            A.Account_Serial, A.Account_NameEng, A.Parent_Account_Code, A.Branch, " & vbCrLf
    MySQL = MySQL + "            A.Sum_account, A.cost_center, A.currenct_code, " & vbCrLf
    MySQL = MySQL + "            T.Remark, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS note_value_by_characters, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(50)) AS foxy_no, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No1, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS project_id, " & vbCrLf
    MySQL = MySQL + "            CAST('Invoice' AS nvarchar(200)) AS NotesTypeNamee, " & vbCrLf
    MySQL = MySQL + "            TT.TransactionEnglishName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS opr_fullcode, " & vbCrLf
    MySQL = MySQL + "            CAST(T.NoteSerial1 AS float) AS NoteSerial1, " & vbCrLf
    MySQL = MySQL + "            B.branch_id, B.ActivityTypeId, " & vbCrLf
    MySQL = MySQL + "            A.opening_balance, A.opening_balance_type, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS bit)          AS notes_all, " & vbCrLf
    MySQL = MySQL + "            T.Transaction_ID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)            AS Item_ID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS float)          AS ShowQty, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS money)          AS showPrice, " & vbCrLf
    MySQL = MySQL + "            T.Transaction_Type, " & vbCrLf
    MySQL = MySQL + "            TT.StockEffect, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(50))   AS ItemCode, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS ItemName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS ItemNamee, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)            AS UnitId, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(200))  AS UnitName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(200))  AS UnitNamee, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS money)          AS discountvalue, " & vbCrLf
    MySQL = MySQL + "            B.branch_name, B.branch_namee, B.RegionID, S.Name, S.NameE " & vbCrLf

    MySQL = MySQL + " FROM       dbo.Transactions T " & vbCrLf
    MySQL = MySQL + " INNER JOIN dbo.TblCustemers C ON C.CusID = T.CusID " & vbCrLf
    MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON A.Account_Code = C.Account_Code " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblBranchesData B ON B.branch_id = T.BranchId " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblSection S ON S.Id = B.RegionID " & vbCrLf

    MySQL = MySQL + " WHERE      (A.Account_Code = '" & StrAccountCode & "') " & vbCrLf
    MySQL = MySQL + " AND T.Transaction_Type = 21 " & vbCrLf
    MySQL = MySQL + " AND NOT EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS DEV WHERE DEV.Transaction_ID = T.Transaction_ID) " & vbCrLf

    If BranchID <> 0 Then MySQL = MySQL & " AND T.BranchId = " & BranchID & vbCrLf
    If ActivityId <> 0 Then MySQL = MySQL & " AND B.ActivityTypeId = " & ActivityId & vbCrLf
    If RegionID <> 0 Then MySQL = MySQL & " AND B.RegionID = " & RegionID & vbCrLf

    If hasFrom Or hasTo Then
        MySQL = MySQL & " AND ( 1=1 " & vbCrLf
        If hasFrom Then MySQL = MySQL & "     AND T.Transaction_Date >= " & SQLDate(Me.BegineDate, True) & vbCrLf
        If hasTo Then MySQL = MySQL & "     AND T.Transaction_Date <= " & SQLDate(Me.EndDate, True) & vbCrLf
        MySQL = MySQL & " ) " & vbCrLf
    End If

    '--------------------------
    ' (B2) «·„»œ∆Ì… 6
    '--------------------------
    MySQL = MySQL & vbCrLf & " UNION ALL " & vbCrLf & vbCrLf

    MySQL = MySQL + " SELECT     " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS ChqueNum, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS float)         AS order_no, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)           AS CBoBasedON, " & vbCrLf
    MySQL = MySQL + "            T.ManualNo, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(50))  AS RecordDateH, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS float)         AS Double_Entry_Vouchers_ID, " & vbCrLf
    MySQL = MySQL + "            ISNULL(( " & vbCrLf
    MySQL = MySQL + "                SELECT SUM( " & vbCrLf
    MySQL = MySQL + "                       ((ISNULL(TD2.ShowQty,0) * ISNULL(TD2.showPrice,0)) - ISNULL(TD2.discountvalue,0)) " & vbCrLf
    MySQL = MySQL + "                       + ISNULL(TD2.Vat,0) " & vbCrLf
    MySQL = MySQL + "                ) " & vbCrLf
    MySQL = MySQL + "                FROM dbo.Transaction_Details TD2 " & vbCrLf
    MySQL = MySQL + "                WHERE TD2.Transaction_ID = T.Transaction_ID " & vbCrLf
    MySQL = MySQL + "            ),0) AS DEV_Value, " & vbCrLf
    MySQL = MySQL + "            CAST(0 AS int)             AS Credit_Or_Debit, " & vbCrLf
    MySQL = MySQL + "            A.Account_Code, " & vbCrLf
    MySQL = MySQL + "            T.TransactionComment AS DEV_DES, " & vbCrLf
    MySQL = MySQL + "            CAST(N'Provisional Invoice (No Final Yet)' AS nvarchar(4000)) AS DevDESE, " & vbCrLf
    MySQL = MySQL + "            A.Account_Name, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No, " & vbCrLf
    MySQL = MySQL + "            CAST(N'›« Ê—… „»Ì⁄«  „' AS nvarchar(200)) AS NotesTypeName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS UserID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS UserName, " & vbCrLf
    MySQL = MySQL + "            T.Transaction_Date         AS RecordDate, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS ReceiptID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS float)        AS Notes_ID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS OperaID, " & vbCrLf
    MySQL = MySQL + "            T.Transaction_Serial, T.Transaction_Date, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS bit)          AS Posted, " & vbCrLf
    MySQL = MySQL + "            TT.TransactionTypeName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS datetime)     AS PostedDate, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS PostedUserID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS Account_Interval_ID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS datetime)     AS NoteDate, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS NoteType, " & vbCrLf
    MySQL = MySQL + "            CAST(T.Noteserial AS nvarchar(50)) AS NoteSerial, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS money)        AS Note_Value, " & vbCrLf
    MySQL = MySQL + "            A.Account_Serial, A.Account_NameEng, A.Parent_Account_Code, A.Branch, " & vbCrLf
    MySQL = MySQL + "            A.Sum_account, A.cost_center, A.currenct_code, " & vbCrLf
    MySQL = MySQL + "            T.Remark, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS note_value_by_characters, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(50)) AS foxy_no, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS DEV_ID_Line_No1, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)          AS project_id, " & vbCrLf
    MySQL = MySQL + "            CAST('Invoice' AS nvarchar(200)) AS NotesTypeNamee, " & vbCrLf
    MySQL = MySQL + "            TT.TransactionEnglishName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(200)) AS opr_fullcode, " & vbCrLf
    MySQL = MySQL + "            T.NoteSerial1              AS NoteSerial1, " & vbCrLf
    MySQL = MySQL + "            B.branch_id, B.ActivityTypeId, " & vbCrLf
    MySQL = MySQL + "            A.opening_balance, A.opening_balance_type, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS bit)          AS notes_all, " & vbCrLf
    MySQL = MySQL + "            T.Transaction_ID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)            AS Item_ID, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS float)          AS ShowQty, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS money)          AS showPrice, " & vbCrLf
    MySQL = MySQL + "            T.Transaction_Type, " & vbCrLf
    MySQL = MySQL + "            TT.StockEffect, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(50))   AS ItemCode, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS ItemName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(4000)) AS ItemNamee, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS int)            AS UnitId, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(200))  AS UnitName, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS nvarchar(200))  AS UnitNamee, " & vbCrLf
    MySQL = MySQL + "            CAST(NULL AS money)          AS discountvalue, " & vbCrLf
    MySQL = MySQL + "            B.branch_name, B.branch_namee, B.RegionID, S.Name, S.NameE " & vbCrLf

    MySQL = MySQL + " FROM       dbo.Transactions T " & vbCrLf
    MySQL = MySQL + " INNER JOIN dbo.TblCustemers C ON C.CusID = T.CusID " & vbCrLf
    MySQL = MySQL + " INNER JOIN dbo.ACCOUNTS A ON A.Account_Code = C.Account_Code " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblBranchesData B ON B.branch_id = T.BranchId " & vbCrLf
    MySQL = MySQL + " LEFT  JOIN dbo.TblSection S ON S.Id = B.RegionID " & vbCrLf

    MySQL = MySQL + " WHERE      (A.Account_Code = '" & StrAccountCode & "') " & vbCrLf
    MySQL = MySQL + " AND T.Transaction_Type = 6 " & vbCrLf
    MySQL = MySQL + " AND NOT EXISTS (SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS DEV WHERE DEV.Transaction_ID = T.Transaction_ID) " & vbCrLf
    MySQL = MySQL + " AND NOT EXISTS ( " & vbCrLf
    MySQL = MySQL + "     SELECT 1 " & vbCrLf
    MySQL = MySQL + "     FROM dbo.Transactions T2 " & vbCrLf
    MySQL = MySQL + "     WHERE T2.Transaction_Type = 21 " & vbCrLf
    MySQL = MySQL + "       AND ISNULL(T2.CBoBasedON,0) = 10 " & vbCrLf
    MySQL = MySQL + "       AND T2.order_no = T.NoteSerial1 " & vbCrLf
    MySQL = MySQL + " ) " & vbCrLf

    If BranchID <> 0 Then MySQL = MySQL & " AND T.BranchId = " & BranchID & vbCrLf
    If ActivityId <> 0 Then MySQL = MySQL & " AND B.ActivityTypeId = " & ActivityId & vbCrLf
    If RegionID <> 0 Then MySQL = MySQL & " AND B.RegionID = " & RegionID & vbCrLf

    If hasFrom Or hasTo Then
        MySQL = MySQL & " AND ( 1=1 " & vbCrLf
        If hasFrom Then MySQL = MySQL & "     AND T.Transaction_Date >= " & SQLDate(Me.BegineDate, True) & vbCrLf
        If hasTo Then MySQL = MySQL & "     AND T.Transaction_Date <= " & SQLDate(Me.EndDate, True) & vbCrLf
        MySQL = MySQL & " ) " & vbCrLf
    End If

    MySQL = MySQL & vbCrLf & " ORDER BY RecordDate, Transaction_Date, NoteSerial, Transaction_ID " & vbCrLf

End If
    updateProfitsToall

    If SystemOptions.DateOpt = 0 Then
        
If ShowValuee = 1 Then
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Sub-MassterE.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Sub-MassterEngE.rpt"
        End If
      Else
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Sub-Masster.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Sub-MassterEng.rpt"
        End If
 End If
        If detailedtransaction = 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                     StrFileName = App.path & "\Reports\" & "Sub-MassterdetailedTransactions.rpt"
                 Else
                     StrFileName = App.path & "\Reports\" & "Sub-MassterdetailedTransactions.rpt"
                 End If
        End If


    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Sub-MassterH.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Sub-MassterEngH.rpt"
        End If

    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
        
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        RsData.Close

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\New Reports\" & "Sub-Masster1.rpt"
        Else
            StrFileName = App.path & "\Reports\New Reports\" & "Sub-MassterEng1.rpt"
        End If

        MySQL = "SELECT     Account_Code, Account_Name, Account_Serial, Account_NameEng, DepitBalance, CreditBalance, Balance, opening_balance, opening_balance_type"
        MySQL = MySQL & " from dbo.ACCOUNTS"
        MySQL = MySQL & "  WHERE     (Account_Code = '" & StrAccountCode & "')"



        RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        'GetMsgs 138, vbExclamation
        '   Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        '   MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '   RsData.Close
        '   Set RsData = Nothing
        '   Screen.MousePointer = vbDefault
        '   Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData



     Dim s As String
     s = " select TblChecqueBoxContent1.*,TblCustemers.Account_Code from TblChecqueBoxContent1"
s = s & " Inner join Notes On Notes.NoteID = TblChecqueBoxContent1.NoteID"
s = s & " left outer join TblCustemers On TblCustemers.CusID = notes.CusID"
s = s & " left outer join TblBoxesData On TblBoxesData.BoxId = notes.BoxId"


s = s & " Where (payed = 0 Or payed Is Null)"
 If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        s = s + " and TblChecqueBoxContent1.RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        s = s + " and TblChecqueBoxContent1.RecordDate <=" & SQLDate(EndDate, True) & ""
    End If
     s = s & " and (TblCustemers.Account_Code ='" & StrAccountCode & "' or TblBoxesData.Account_Code ='" & StrAccountCode & "' or Notes.BTCashAccountcode ='" & StrAccountCode & "')"
     Dim RsData2 As New ADODB.Recordset
    Set RsData2 = New ADODB.Recordset
      
      If SystemOptions.SuppCreat4Acc And Not manyString Then
         RsData2.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
         If Not RsData2.EOF And ShowValuee = 0 Then
           
            
        End If
        If StrFileName = App.path & "\Reports\" & "Sub-Masster.rpt" Or StrFileName = App.path & "\Reports\" & "Sub-MassterEng.rpt" Then
            xReport.OpenSubreport("aa").Database.SetDataSource RsData2
        End If
        If Not RsData2.EOF Then
            
        End If
    End If
 
    Dim cCompanyInfo As New ClsCompanyInfo

    If d1 <> Null Then
        Me.BegineDate = d1
    End If

    If d2 <> Null Then
        Me.EndDate = d2
    End If

    Dim activityName As String
    Dim branchname As String
    Dim fullcode As String
    branchname = get_branch_name(val(BranchID), activityName)
If Not manyString Then
GetCustomerIdByAccountCodeLong StrAccountCode, fullcode
End If
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(4).AddCurrentValue branchname  'get_branch_name(val(branch_id))
        xReport.ParameterFields(7).AddCurrentValue openingbalanceDes
        'CustomerAgeingData
        xReport.ParameterFields(8).AddCurrentValue CustomerAgeingData





        
       If fullcode <> "" Then
            xReport.ParameterFields(13).AddCurrentValue "ﬂÊœ «·⁄„Ì· : " & fullcode
        Else
            xReport.ParameterFields(13).AddCurrentValue ""
        
     End If
     
        
        If salesPersonName <> "" Then
            xReport.ParameterFields(9).AddCurrentValue "  «·„‰œÊ» " & salesPersonName
        Else
            xReport.ParameterFields(9).AddCurrentValue "    "
        End If

        xReport.ParameterFields(10).AddCurrentValue ShowAgingReport

        If ChartPrintinAS = True Then
            xReport.ParameterFields(11).AddCurrentValue "1"
        Else
            xReport.ParameterFields(11).AddCurrentValue "0"
        End If

        'ShowAgingReport
        StrReportTitle = "ﬂ‘› Õ”«»   " & StrAccountName
If withouopenening = True Then
        StrReportTitle = "ﬂ‘› Õ—ﬂ…   " & StrAccountName
End If

If ProjectName <> "" Then
StrReportTitle = StrReportTitle & " ··„‘—Ê⁄  " & ProjectName
End If

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        
            If SystemOptions.DateOpt = 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰ " & Format(Me.BegineDate, "yyyy/Mm/dd") & "  "
            Else
                StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰ " & ToHijriDate(Format(Me.BegineDate, "yyyy/Mm/dd")) & "  "

            End If
    
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            If SystemOptions.DateOpt = 0 Then

                StrReportTitle = StrReportTitle & CHR(13) & "   ≈·Ï " & "   " & Format(EndDate, "yyyy/Mm/dd")
            Else
                StrReportTitle = StrReportTitle & CHR(13) & "   ≈·Ï " & "   " & ToHijriDate(Format(EndDate, "yyyy/Mm/dd"))
            End If
        End If
     
        If ActivityId <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & activityName
        End If

        If BranchID <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & branchname
        End If
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue branchname  'get_branch_name(Val(branchId))
        xReport.ParameterFields(7).AddCurrentValue openingbalanceDes
        xReport.ParameterFields(8).AddCurrentValue CustomerAgeingData
     
       If fullcode <> "" Then
            xReport.ParameterFields(13).AddCurrentValue "Code " & fullcode
        Else
          xReport.ParameterFields(13).AddCurrentValue ""
        
     End If
     
        StrReportTitle = "Statement of ACC :" & account_serial & CHR(13) & StrAccountName & CHR(13)    '& " Account "
    
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From : " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To   :  " & (EndDate)
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(5).AddCurrentValue account_serial
    xReport.ParameterFields(6).AddCurrentValue StrAccountName
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function



Public Function ShowLedger1(StrAccountCode As String, _
                            StrAccountName As String, _
                            des As String, _
                            Optional manyaccounts As Boolean = False, Optional ShowValuee As Integer = 0) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From RptLedger_Sub "

    If manyaccounts = True Then
        MySQL = MySQL & "  where  Account_Code in (" & StrAccountCode & ")"
    Else
        MySQL = MySQL + " WHERE Account_Code='" & StrAccountCode & "'"
    End If

    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(EndDate, True) & ""
    End If

    If FrmAccountingReport.dcBranch.BoundText <> "" Then
        MySQL = MySQL & " and branch_id = " & val(FrmAccountingReport.dcBranch.BoundText)
    End If

    If FrmAccountingReport.DCActivity.BoundText <> "" Then
        MySQL = MySQL & " and ActivityTypeId = " & val(FrmAccountingReport.DCActivity.BoundText)
    End If

    If des <> "" Then
        MySQL = MySQL + " and DEV_DES like'%" & des & "%'"
    End If

    updateProfitsToall
If ShowValuee = 1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        '  StrFileName = App.Path & "\Reports\" & "Sub-Masster _with_des.rpt"
        StrFileName = App.path & "\Reports\" & "Sub-MassterE.rpt"
    Else
        'StrFileName = App.Path & "\Reports\" & "Sub-Masster_Eng _with_des.rpt"
        StrFileName = App.path & "\Reports\" & "Sub-Masster_EngE.rpt"
    End If
Else
    If SystemOptions.UserInterface = ArabicInterface Then
        '  StrFileName = App.Path & "\Reports\" & "Sub-Masster _with_des.rpt"
        StrFileName = App.path & "\Reports\" & "Sub-Masster.rpt"
    Else
        'StrFileName = App.Path & "\Reports\" & "Sub-Masster_Eng _with_des.rpt"
        StrFileName = App.path & "\Reports\" & "Sub-Masster_Eng.rpt"
    End If
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "ﬂ‘› Õ”«» ÿ»ﬁ« ··‘—Õ " & StrAccountName

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰    " & Format(Me.BegineDate, "yyyy/M/d") & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle & CHR(13) & " ≈·Ï " & Format(EndDate, "yyyy/M/d")
        End If
    
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If
    
    Else
        xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To " & (EndDate) & "∫"
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function


Public Function ShowLedger1x08102017(StrAccountCode As String, _
                            StrAccountName As String, _
                            des As String, _
                            Optional manyaccounts As Boolean = False) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From RptLedger_Sub "

    If manyaccounts = True Then
        MySQL = MySQL & "  where  Account_Code in (" & StrAccountCode & ")"
    Else
        MySQL = MySQL + " WHERE Account_Code='" & StrAccountCode & "'"
    End If

    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(EndDate, True) & ""
    End If

    If FrmAccountingReport.dcBranch.BoundText <> "" Then
        MySQL = MySQL & " and branch_id = " & val(FrmAccountingReport.dcBranch.BoundText)
    End If

    If FrmAccountingReport.DCActivity.BoundText <> "" Then
        MySQL = MySQL & " and ActivityTypeId = " & val(FrmAccountingReport.DCActivity.BoundText)
    End If

    If des <> "" Then
        MySQL = MySQL + " and DEV_DES like'%" & des & "%'"
    End If

    updateProfitsToall

    If SystemOptions.UserInterface = ArabicInterface Then
        '  StrFileName = App.Path & "\Reports\" & "Sub-Masster _with_des.rpt"
        StrFileName = App.path & "\Reports\" & "Sub-Masster.rpt"
    Else
        'StrFileName = App.Path & "\Reports\" & "Sub-Masster_Eng _with_des.rpt"
        StrFileName = App.path & "\Reports\" & "Sub-Masster_Eng.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "ﬂ‘› Õ”«» ÿ»ﬁ« ··‘—Õ " & StrAccountName

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰    " & Format(Me.BegineDate, "yyyy/M/d") & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle & CHR(13) & " ≈·Ï " & Format(EndDate, "yyyy/M/d")
        End If
    
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If
    
    Else
        xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To " & (EndDate) & "∫"
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function ShowLedgers(StrAccountCode As String, _
                            StrAccountName As String, _
                            des As String, _
                            Optional manyaccounts As Boolean = False) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From RptLedger_Sub "

MySQL = "  SELECT     *"
MySQL = MySQL & "  FROM         dbo.RptLedger_Sub RIGHT OUTER JOIN"
MySQL = MySQL & "                        dbo.ACCOUNTS ON dbo.RptLedger_Sub.Account_Code = dbo.ACCOUNTS.Account_Code"
MySQL = MySQL & "   or dbo.ACCOUNTS.Account_Code in (" & StrAccountCode & ")"

'MySQL = MySQL & "  WHERE     (dbo.RptLedger_Sub.RecordDate IS NULL OR"
'MySQL = MySQL & "                        dbo.RptLedger_Sub.RecordDate >= '30-Oct-2016' AND dbo.RptLedger_Sub.RecordDate <= '30-Oct-2016') AND (dbo.ACCOUNTS.Account_Code IN (N'a1a2a1a1a1a1',"
'MySQL = MySQL & "                        n 'a1a2a1a1a1a2', N'a1a2a1a1a2a1', N'a1a2a1a1a2a2', N'a1a2a1a2a1a1', N'a1a2a1a2a1a2', N'a1a2a1a2a1a3', N'a1a2a1a2a2a1', N'a1a2a1a2a2a2',"
'MySQL = MySQL & "                        n 'a1a2a1a2a2a3', N'a1a2a1a1a3a1', N'a1a2a1a1a3a2', N'a1a2a1a1a4a1', N'a1a2a1a1a4a2', N'a1a2a1a2a3')"
'MySQL = MySQL & "  )"
'
    If manyaccounts = True Then
        MySQL = MySQL & "  where  ACCOUNTS.Account_Code in (" & StrAccountCode & ")"
  '  Else
  '      MySQL = MySQL + " WHERE ACCOUNTS.Account_Code='" & StrAccountCode & "'"
    End If

    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " AND (  dbo.RptLedger_Sub.RecordDate IS NULL OR  (   RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(EndDate, True) & "))"
    End If

    If FrmAccountingReport.dcBranch.BoundText <> "" Then
        MySQL = MySQL & " and branch_id = " & val(FrmAccountingReport.dcBranch.BoundText)
    End If

    If FrmAccountingReport.DCActivity.BoundText <> "" Then
        MySQL = MySQL & " and ActivityTypeId = " & val(FrmAccountingReport.DCActivity.BoundText)
    End If

    If des <> "" Then
        MySQL = MySQL + " and DEV_DES like'%" & des & "%'"
    End If

    updateProfitsToall

    If SystemOptions.UserInterface = ArabicInterface Then
        '  StrFileName = App.Path & "\Reports\" & "Sub-Masster _with_des.rpt"
        StrFileName = App.path & "\Reports\" & "Sub-Massters.rpt"
    Else
        'StrFileName = App.Path & "\Reports\" & "Sub-Masster_Eng _with_des.rpt"
        StrFileName = App.path & "\Reports\" & "Sub-Masster_Engs.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "ﬂ‘› Õ”«» ÿ»ﬁ« ··‘—Õ " & StrAccountName

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰    " & Format(Me.BegineDate, "yyyy/M/d") & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle & CHR(13) & " ≈·Ï " & Format(EndDate, "yyyy/M/d")
        End If
    
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If
    
    Else
        xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To " & (EndDate) & "∫"
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function ShowTripsData(Optional DriverID As Integer = 0, _
                              Optional CityFromId As Integer = 0, _
                              Optional CityToId As Integer = 0, _
                              Optional CarID As Integer = 0, _
                              Optional totals As Boolean, _
                              Optional ReportType As Integer = 0) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
    MySQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], "
    MySQL = MySQL & "  dbo.ACCOUNTS.Account_Name, dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteSerial, TblCountriesGovernments_2.GovernmentName AS FromCity,"
    MySQL = MySQL & " TblCountriesGovernments_1.GovernmentName AS ToCity, dbo.notes_all.NoteType, dbo.notes_all.CityFromId, dbo.notes_all.CityToId, dbo.notes_all.Location,"
    MySQL = MySQL & " dbo.notes_all.CarId, dbo.notes_all.TravelPrice, dbo.notes_all.DriverValue, dbo.notes_all.DriverEra, dbo.notes_all.totalExpenses AS Desil, dbo.TblCarsData.BoardNO,"
    MySQL = MySQL & " dbo.TblCarsData.Model, dbo.TblCarsData.Name, dbo.TblCarsData.LicenseNO, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.notes_all.Distance,"
    MySQL = MySQL & " dbo.notes_all.KmPrice, dbo.notes_all.DriverId, dbo.notes_all.DriverPercentage, dbo.notes_all.startDate, dbo.notes_all.EndDate, dbo.notes_all.StartTime,"
    MySQL = MySQL & " dbo.notes_all.EndTime, dbo.notes_all.KMCounterBeforeStart, dbo.notes_all.KMCounterAtEnd, dbo.notes_all.NoteCashingType, dbo.TblCustemers.CusName,"
    MySQL = MySQL & " dbo.TblCustemers.CusNamee, dbo.notes_all.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ExpensesType.ID,"
    MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.notes_all.general_des"
    MySQL = MySQL & " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
    MySQL = MySQL & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code RIGHT OUTER JOIN"
    MySQL = MySQL & " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID RIGHT OUTER JOIN"
    MySQL = MySQL & " dbo.ExpensesType ON dbo.ACCOUNTS.Account_Code = dbo.ExpensesType.Account_Code RIGHT OUTER JOIN"
    MySQL = MySQL & " dbo.notes_all INNER JOIN"
    MySQL = MySQL & " dbo.TblCarsData ON dbo.notes_all.CarId = dbo.TblCarsData.id INNER JOIN"
    MySQL = MySQL & " dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.notes_all.CityFromId = TblCountriesGovernments_2.GovernmentID ON"
    MySQL = MySQL & " dbo.Notes.notes_all = dbo.notes_all.NoteID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblCustemers ON dbo.notes_all.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmployee ON dbo.notes_all.DriverId = dbo.TblEmployee.Emp_Code LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.notes_all.CityToId = TblCountriesGovernments_1.GovernmentID"
    MySQL = MySQL & "  WHERE     (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"

    If DriverID <> 0 Then
        MySQL = MySQL & " and dbo.notes_all.DriverId=" & DriverID
    End If

    If CarID <> 0 Then
        MySQL = MySQL & " and dbo.notes_all.CarID=" & CarID
    End If

    If CityFromId <> 0 Then
        MySQL = MySQL & " and dbo.notes_all.CityFromId=" & CityFromId
    End If

    If CityToId <> 0 Then
        MySQL = MySQL & " and dbo.notes_all.CityToId=" & CityToId
    End If

    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and dbo.notes_all.NoteDate  >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and dbo.notes_all.NoteDate  <=" & SQLDate(EndDate, True) & ""
    End If
 
    If totals = False Then
        StrReportTitle = " ﬁ—Ì—«·—Õ·«    Õ·Ì·Ì  ⁄‰ › —… „Õœœ… " '

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\Transporter\" & "Trips.rpt"
        Else
            StrFileName = App.path & "\Reports\Transporter\" & "Trips.rpt"
        End If

    Else
        
        If ReportType = 0 Then
            StrReportTitle = " ﬁ—Ì—«·—Õ·«  ÿ»ﬁ« ··‘«Õ‰… «Ã„«·Ì«  ⁄‰ › —… „Õœœ… " '

            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\Transporter\" & "Tripstotals.rpt"
            Else
                StrFileName = App.path & "\Reports\Transporter\" & "Tripstotals.rpt"
            End If

        ElseIf ReportType = 1 Then
            StrReportTitle = " ﬁ—Ì—«—»«Õ «·‘«Õ‰«  ⁄‰ › —… „Õœœ… " '

            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\Transporter\" & "Tripstotals1.rpt"
            Else
                StrFileName = App.path & "\Reports\Transporter\" & "Tripstotals1.rpt"
            End If

        End If

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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        'StrReportTitle = " ﬁ—Ì—«·—Õ·«  ⁄‰ › —… „Õœœ… " '& StrAccountName
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + CHR(13) + "„‰:      " & Format(Me.BegineDate, "yyyy/M/d") & "   "
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + CHR(13) + " ≈·Ï:      " & Format(EndDate, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))

        ' StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (EndDate) & "∫"
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
 
   Public Function ProductionAuditing1(Optional noreciveQty As Boolean = False) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    MySQL = "  SELECT     TOP 100 PERCENT  dbo.Transaction_Details.Item_ID, SUM(dbo.Transaction_Details.ShowQty) AS ShowQty, "
    MySQL = MySQL & "   dbo.GetOrderActualReciveqty(dbo.Transaction_Details.Item_ID, dbo.Transactions.Transaction_Serial) AS RecivQty,  CAST (dbo.Transactions.Transaction_Serial AS INT )"
    MySQL = MySQL & "   AS orderNo, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee"
    MySQL = MySQL & "  FROM         dbo.Transactions INNER JOIN"
    MySQL = MySQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    MySQL = MySQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    MySQL = MySQL & "  Where (dbo.Transactions.Transaction_Type = 26)"
 
    If noreciveQty = True Then
        MySQL = MySQL & "  and  ( dbo.Transactions.Product_Receive_voucher_Serial =0 or dbo.Transactions.Product_Receive_voucher_Serial is null)"
    End If
 
    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and dbo.Transactions.Transaction_Date >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(EndDate, True) & ""
    End If
 
    MySQL = MySQL & "  GROUP BY dbo.Transaction_Details.Item_ID, dbo.GetOrderActualReciveqty(dbo.Transaction_Details.Item_ID, dbo.Transactions.Transaction_Serial),"
    MySQL = MySQL & "   (dbo.Transactions.Transaction_Serial    ), dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee"
 
    MySQL = MySQL & "  ORDER BY  (dbo.Transactions.Transaction_Serial    )  "
  
    If noreciveQty = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\Production\" & "ProductionAuditing1.rpt"
        Else
            StrFileName = App.path & "\Reports\Production\" & "ProductionAuditing1.rpt"
        End If

    Else
 
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\Production\" & "ProductionAuditing2.rpt"
        Else
            StrFileName = App.path & "\Reports\Production\" & "ProductionAuditing2.rpt"
        End If

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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        If noreciveQty = False Then
            StrReportTitle = " ﬁ—Ì— ÌÊ÷Õ «·›—ﬁ »Ì‰  «·„ Êﬁ⁄ Ê «·›⁄·Ï ··«‰ «Ã " '& StrAccountName
        Else
            StrReportTitle = " ﬁ—Ì— ÌÊ÷Õ «Ê«„— «·«‰ «Ã «·„› ÊÕ…   " '& StrAccountName
        End If
  
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + CHR(13) + "From: " & Format(Me.BegineDate, "dd/mm/yyyy") & "   "
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + "To: " & Format(EndDate, "dd/mm/yyyy") & " "
        End If
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
   
        ' StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (EndDate) & "∫"
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function



 
 Public Function ProductionAuditing() As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    MySQL = " SELECT     dbo.GetOHForPO(dbo.Transactions.WorkOrderNO)  as FactoryExpenses   ,  dbo.Transactions.WorkOrderNO, dbo.TblDoCumentsTypes.name, dbo.TblDoCumentsTypes.namee, SUM(dbo.Transaction_Details.Quantity) AS totalqty, "
    MySQL = " SELECT     205 as FactoryExpenses   ,  dbo.Transactions.WorkOrderNO, dbo.TblDoCumentsTypes.name, dbo.TblDoCumentsTypes.namee, SUM(dbo.Transaction_Details.Quantity) AS totalqty, "
    MySQL = MySQL & "    SUM(dbo.Transaction_Details.Price * dbo.Transaction_Details.Quantity) AS totalprice, SUM(dbo.Transaction_Details.ShowQty) AS totalshowqty,"
    MySQL = MySQL & " SUM(dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) AS totalshoprice, MAX(Transactions_1.Transaction_Date) AS Transaction_Date,"
    MySQL = MySQL & " MAX(Transactions_1.EndDate) AS OrderEndDate, MAX(Transaction_Details_1.Quantity) AS orderQtySF"
    MySQL = MySQL & " FROM         dbo.Transactions Left Outer JOIN"
    MySQL = MySQL & " dbo.TblDoCumentsTypes ON dbo.Transactions.Doctype = dbo.TblDoCumentsTypes.id INNER JOIN"
    MySQL = MySQL & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    MySQL = MySQL & " dbo.Transactions Transactions_1 ON  dbo.Transactions.Transaction_ID = Transactions_1.nots2 INNER JOIN"
    MySQL = MySQL & "  dbo.Transaction_Details Transaction_Details_1 ON Transactions_1.Transaction_ID = Transaction_Details_1.Transaction_ID"
    MySQL = MySQL & "  Where (dbo.Transactions.Transaction_Type = 27)"
 
    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and dbo.Transactions.Transaction_Date >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(EndDate, True) & ""
    End If

    MySQL = MySQL & "  GROUP BY dbo.Transactions.WorkOrderNO, dbo.TblDoCumentsTypes.name, dbo.TblDoCumentsTypes.namee , dbo.Transactions.FactoryExpenses"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Production\" & "ProductionAuditing.rpt"
    Else
        StrFileName = App.path & "\Reports\Production\" & "ProductionAuditing.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "Job Orders " '& StrAccountName

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + CHR(13) + "From: " & Format(Me.BegineDate, "dd/mm/yyyy") & "   "
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + "To: " & Format(EndDate, "dd/mm/yyyy") & " "
        End If
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
   
        ' StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (EndDate) & "∫"
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function



Public Function ShowProductItems(Optional bigunit As Boolean, _
                                 Optional GroupID As Double = 0, _
                                 Optional GroupName As String, _
                                 Optional StoreID As Integer = 0, _
                                 Optional StoreName As String, Optional ByRef reportno As Integer, Optional FromDate As String, Optional ToDate As String) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    If reportno = 1 Then
   
 
 
        MySQL = "SELECT     dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, "
MySQL = MySQL + "                       dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transactions.NoteSerial1, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName,"
MySQL = MySQL + "                       dbo.TblItems.ItemNamee , dbo.TblItems.fullcode"
MySQL = MySQL + "  FROM         dbo.Transactions INNER JOIN"
MySQL = MySQL + "                       dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
MySQL = MySQL + "                       dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
MySQL = MySQL + "                       dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
MySQL = MySQL + "  Where (dbo.Transactions.Transaction_Type = 26)"


          
          
 



Else

    MySQL = "SELECT   dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice,  dbo.Transactions.WorkOrderNO,   dbo.Transactions.NoteSerial1 ,dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_Type_Sub, dbo.Transactions.order_no, " & "        dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, " & "        dbo.Transactions.Transaction_Date " & " FROM         dbo.Transactions INNER JOIN " & "                      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN " & "                      dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID  " & "  WHERE     (dbo.Transactions.Transaction_Type = 28)  "
End If

    If StoreID <> 0 Then
        MySQL = MySQL + " and  Transactions.storeid=" & StoreID
    End If
 
 
 
    If GroupID <> 0 Then
        MySQL = MySQL + " and  GroupID=" & GroupID
    End If
 
    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and dbo.Transactions.Transaction_Date >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(EndDate, True) & ""
    End If

  MySQL = MySQL + " order by dbo.Transactions.Transaction_Date "
  If reportno = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "Â·  —Ìœ«· ﬁ—Ì— ÿ»ﬁ« ··Õ—ﬂ…«÷€ÿ Yes . Ã„Ì⁄ ⁄·Ï „” ÊÏ «·’‰› «÷€ÿ No "
    Else
    Msg = "Do you want group by Transection press Yes. by item press No"
    End If
    Dim IntRes As Integer
    IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
    If IntRes = vbYes Then
    If bigunit = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\Production\" & "ProductedItems.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "ProductedItems.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\Production\" & "ProductedItems1.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "ProductedItems1.rpt"
        End If
    End If
   Else
  ''//////////
      If bigunit = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
          '  StrFileName = App.path & "\Reports\Production\" & "ProductedItems.rpt"
            StrFileName = App.path & "\Reports\REPORTS NEW\" & "ProductedItemsByitems.rpt"
        Else
          '  StrFileName = App.path & "\Reports\" & "ProductedItems.rpt"
            StrFileName = App.path & "\Reports\REPORTS NEW\" & "ProductedItemsByitemsE.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            'StrFileName = App.path & "\Reports\Production\" & "ProductedItems1.rpt"
            StrFileName = App.path & "\Reports\REPORTS NEW\" & "ProductedItems1ByItem.rpt"
        Else
          '  StrFileName = App.path & "\Reports\" & "ProductedItems1.rpt"
             StrFileName = App.path & "\Reports\REPORTS NEW\" & "ProductedItems1ByItemE.rpt"
        End If
    End If
  '/////
   End If
    End If
    If reportno = 1 Then
  
    StrFileName = App.path & "\Reports\REPORTS NEW\" & "productionAuditing.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.RecordCount = 0 Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        If reportno = 1 Then
 xReport.ParameterFields(4).AddCurrentValue FromDate
 xReport.ParameterFields(5).AddCurrentValue ToDate
 

End If


    If reportno = 1 Then
       StrReportTitle = "„—«Ã⁄Â √Ê«„— «·«‰ «Ã " '& StrAccountName
       
       If GroupName <> "" Then
       StrReportTitle = StrReportTitle & " ··„Ã„Ê⁄Â  " & GroupName
       
       End If
       
       If StoreName <> "" Then
       StrReportTitle = StrReportTitle & " ··„Œ“‰  " & StoreName
       
       End If
       
       
        Else
       
        
        StrReportTitle = " ﬁ—Ì—«·«’‰«› «·„‰ Ã… Œ·«· › —… " '& StrAccountName
        If reportno = 1 Then
 xReport.ParameterFields(4).AddCurrentValue FromDate
 xReport.ParameterFields(5).AddCurrentValue ToDate
 

End If

End If


        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + CHR(13) + " »œ«Ì… „‰:      " & Format(Me.BegineDate, "yyyy/M/d") & "   "
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " ≈·Ï:      " & Format(EndDate, "yyyy/M/d") & " "
        End If
    
        If GroupID <> 0 Then
            StrReportTitle = StrReportTitle + CHR(13) + "  «·„Ã„Ê⁄Â " + GroupName
        End If
    
        If StoreID <> 0 Then
            StrReportTitle = StrReportTitle + CHR(13) + "  «·„Œ“‰ " + StoreName
        End If
 

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        
      '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))


 

        ' StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (EndDate) & "∫"
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function ShowDamagesItems(Optional StoreID As Integer = 0) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    MySQL = "SELECT    dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice,  dbo.Transactions.WorkOrderNO,   dbo.Transactions.NoteSerial1 ,dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_Type_Sub, dbo.Transactions.order_no, " & "        dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, " & "        dbo.Transactions.Transaction_Date , dbo.Transaction_Details.DamageshowQty, dbo.Transaction_Details.DamageQty" & " FROM         dbo.Transactions INNER JOIN " & "                      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN " & "                      dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID  " & "  WHERE     (dbo.Transactions.Transaction_Type = 28 or dbo.Transactions.Transaction_Type = 34)  AND (NOT (dbo.Transaction_Details.DamageQty IS NULL)) "

    If StoreID <> 0 Then
        MySQL = MySQL + " and   Transactions.storeid=" & StoreID
    End If
 
    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and dbo.Transactions.Transaction_Date >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(EndDate, True) & ""
    End If
 
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Production\" & "damagesItems.rpt"
    Else
        StrFileName = App.path & "\Reports\Production\" & "Sub-damagesItems.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = " ﬁ—Ì—«·«’‰«› «·Â«·ﬂ   Œ·«· › —… " '& StrAccountName

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " »œ«Ì… „‰:      " & Format(Me.BegineDate, "yyyy/M/d") & "   "
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " ≈·Ï:      " & Format(EndDate, "yyyy/M/d") & " "
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))

        ' StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (EndDate) & "∫"
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function ShowProductionSummury() As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    MySQL = " select * from TblTypicalProduction where 1=1 "
 
    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(EndDate, True) & ""
    End If
 
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Inventory\" & "TypicalProductionsSummury.rpt"
    Else
        StrFileName = App.path & "\Reports\Inventory\" & "TypicalProductionsSummury.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "      „·Œ’ «·«‰«Ã Œ·«· › —… " '& StrAccountName

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = CHR(13) & StrReportTitle + " »œ«Ì… „‰:      " & Format(Me.BegineDate, "yyyy/mM/dd") & CHR(13)
            xReport.ParameterFields(4).AddCurrentValue Format(Me.BegineDate, "yyyy/mm/dd")
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " ≈·Ï:      " & Format(EndDate, "yyyy/mM/dd") & " "
            xReport.ParameterFields(5).AddCurrentValue Format(EndDate, "yyyy/mm/dd")
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng

        '     xReport.ParameterFields(5).AddCurrentValue CStr(Format(Enddate, "yyyy/mM/dd") & " ")
        ' StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.BegineDate) & ""
            xReport.ParameterFields(4).AddCurrentValue Format(Me.BegineDate, "yyyy/mm/dd")
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (EndDate) & "∫"
            xReport.ParameterFields(5).AddCurrentValue Format(EndDate, "yyyy/mm/dd")
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function ShowProductionSummury2() As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    MySQL = " SELECT  MaterialIssueVoucherTotals,AccDep,   AccountIntervalID, StartDate, EndDate, SalaryVouchersTotals, Expenses, Allocations, Allocations1, AdvancedPayments, total, SaleValue, SaLePayValue, "
    MySQL = MySQL & "  ServicesValue , profit"
    MySQL = MySQL & " FROM         dbo.TblAccountIntervals where 1=1 "
 
    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and StartDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and EndDate <=" & SQLDate(EndDate, True) & ""
    End If
 
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Inventory\" & "TypicalProductionsSummury2.rpt"
    Else
        StrFileName = App.path & "\Reports\Inventory\" & "TypicalProductionsSummury2.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "      „·Œ’ «·«‰«Ã Œ·«· › —… " '& StrAccountName

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = CHR(13) & StrReportTitle + " »œ«Ì… „‰:      " & Format(Me.BegineDate, "yyyy/mM/dd") & CHR(13)
            xReport.ParameterFields(4).AddCurrentValue Format(Me.BegineDate, "yyyy/mm/dd")
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " ≈·Ï:      " & Format(EndDate, "yyyy/mM/dd") & " "
            xReport.ParameterFields(5).AddCurrentValue Format(EndDate, "yyyy/mm/dd")
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng

        '     xReport.ParameterFields(5).AddCurrentValue CStr(Format(Enddate, "yyyy/mM/dd") & " ")
        ' StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.BegineDate) & ""
            xReport.ParameterFields(4).AddCurrentValue Format(Me.BegineDate, "yyyy/mm/dd")
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (EndDate) & "∫"
            xReport.ParameterFields(5).AddCurrentValue Format(EndDate, "yyyy/mm/dd")
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function ShowProductOrderExpenses(order_no As Integer) As Boolean
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_Type_Sub, dbo.Transaction_Details.Item_ID, " & "                   dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transactions.Transaction_Date, " & "                    dbo.Transactions.Transaction_Serial, dbo.Transactions.FactoryExpenses, dbo.Transactions.FinacilaTotal, dbo.Transactions.Expenses, dbo.Transactions.workerTotal, " & "                    dbo.Transactions.TotalMaterials , dbo.Transactions.LineExpenses " & " FROM         dbo.Transactions INNER JOIN " & "                     dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN " & "                     dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID " & " WHERE     (dbo.Transactions.Transaction_Type = 26) AND (dbo.Transactions.Transaction_Serial = N'" & order_no & "') "
 
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Production\" & "ProductionOrderExpenses.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Sub-Masster_Eng.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "«·„’«—Ì› «·«Ã„«·Ì… ·«„— «‰ «Ã —ﬁ„   " & order_no
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        ' StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

'3333333333333333
Public Function ShowLogFile(wherestr As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = " SELECT  dbo.LogFile.ConnectionData,    dbo.LogFile.Remarks, dbo.LogFile.NotesSerial1, dbo.LogFile.NotesSerial, dbo.LogFile.TransactionType, dbo.LogFile.DescriptionE, dbo.LogFile.Description, "
    MySQL = MySQL & "  dbo.LogFile.LogTime, dbo.LogFile.LogDate, dbo.LogFile.NotesType, dbo.LogFile.UserID, dbo.TblUsers.UserName, dbo.TblNotesTypes.NotesTypeName,"
    MySQL = MySQL & "  dbo.TblNotesTypes.NotesTypeNameE"
    MySQL = MySQL & " FROM         dbo.LogFile INNER JOIN"
    MySQL = MySQL & "  dbo.TblUsers ON dbo.LogFile.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
  MySQL = MySQL & "  dbo.TblNotesTypes ON dbo.LogFile.NotesType = dbo.TblNotesTypes.NotesType"
                      
  'MySQL = "SELECT     dbo.LogFile.Remarks, dbo.LogFile.NotesSerial1, dbo.LogFile.NotesSerial, dbo.LogFile.TransactionType, dbo.LogFile.DescriptionE, dbo.LogFile.Description, "
 'MySQL = MySQL & "  dbo.LogFile.LogTime, dbo.LogFile.LogDate, dbo.LogFile.NotesType, dbo.LogFile.UserID, dbo.TblUsers.UserName, dbo.TblNotesTypes.NotesTypeName,"
 'MySQL = MySQL & "                       dbo.TblNotesTypes.NotesTypeNameE , dbo.Screens.ScreenTitleEng, dbo.Screens.ScreenCaption"
 'MySQL = MySQL & " FROM         dbo.LogFile INNER JOIN"
 'MySQL = MySQL & "                       dbo.TblUsers ON dbo.LogFile.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
 'MySQL = MySQL & "                       dbo.Screens ON dbo.LogFile.Remarks = dbo.Screens.ScreenName LEFT OUTER JOIN"
 'MySQL = MySQL & "                       dbo.TblNotesTypes ON dbo.LogFile.NotesType = dbo.TblNotesTypes.NotesType"
 '
    MySQL = MySQL & " Where  "
    MySQL = MySQL & wherestr
 
    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and   LogDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and LogDate <=" & SQLDate(EndDate, True)
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\New Reports\" & "LogFile.rpt"
    Else
        StrFileName = App.path & "\Reports\New Reports\" & "LogFilee.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = " ﬁ—Ì— ‰‘«ÿ «·„” Œœ„Ì‰  " & CHR(13)
        StrReportTitle = StrReportTitle & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd")
        StrReportTitle = StrReportTitle & "   «·Ï   " & Format(Me.EndDate, "yyyy/mm/dd")
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
'        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = "Log File " & CHR(13)
        StrReportTitle = StrReportTitle & "  From   " & Me.BegineDate
        StrReportTitle = StrReportTitle & "To   " & Me.EndDate
      
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Public Function ShowOrdersiTEMS(order_no As String) As Boolean
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_Type_Sub, dbo.Transaction_Details.Item_ID, " & "                   dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transactions.Transaction_Date, " & "                    dbo.Transactions.Transaction_Serial, dbo.Transactions.FactoryExpenses, dbo.Transactions.FinacilaTotal, dbo.Transactions.Expenses, dbo.Transactions.workerTotal, " & "                    dbo.Transactions.TotalMaterials , dbo.Transactions.LineExpenses " & " FROM         dbo.Transactions INNER JOIN " & "                     dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN " & "                     dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID " & " WHERE     (dbo.Transactions.Transaction_Type = 20) AND (dbo.Transactions.order_no= '" & order_no & "') "
 
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Production\" & "OrdersiTEMS.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Sub-Masster_Eng.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "«·«’‰«› «·„‰ Ã… ›Ì ÿ·»Ì… —ﬁ„   " & order_no
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        ' StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function ShowOrdersStatus(order_no As String) As Boolean
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "SELECT     dbo.Transactions.Transaction_Type, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.Transaction_Details.Quantity, " & "                      dbo.Transaction_Details.Price , dbo.Transactions.Transaction_Date, dbo.Transactions.ORDER_NO, dbo.Transactions.NoteSerial, dbo.Transactions.noteserial1 " & " FROM         dbo.Transactions INNER JOIN " & "                     dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN " & "                     dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID " & " WHERE     (dbo.Transactions.Transaction_Type = 21) AND (dbo.Transactions.order_no = '" & order_no & "')"
 
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Production\" & "OrdersStatus.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Sub-Masster_Eng.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "«·«’‰«› «·„‰ Ã… ›Ì ÿ·»Ì… —ﬁ„   " & order_no
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        ' StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function ShowLedgerx08102018(StrAccountCode As String, _
                           StrAccountName As String, _
                           Optional manyString As Boolean = False, _
                           Optional d1 As Date, _
                           Optional d2 As Date, _
                           Optional ActivityId As Integer = 0, _
                           Optional BranchID As Integer = 0, _
                           Optional CustomerAgeingData As String, _
                           Optional salesPersonName As String, _
                           Optional ShowAgingReport As String, _
                           Optional ByRef account_serial As String, Optional project_id As Integer, Optional ProjectName As String, Optional detailedtransaction As Integer = 0) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    If manyString = False Then
        MySQL = "Select * From RptLedger_Sub "
        MySQL = MySQL + " WHERE Account_Code='" & StrAccountCode & "'"
    Else
        MySQL = "Select * From RptLedger_Sub "
        MySQL = MySQL + StrAccountCode
    End If

    If manyString = False Then

        MySQL = "SELECT   dbo.RptLedger_Sub.ChqueNum,dbo.RptLedger_Sub.ManualNo, dbo.RptLedger_Sub.RecordDateH,   dbo.RptLedger_Sub.Double_Entry_Vouchers_ID, dbo.RptLedger_Sub.DEV_Value, dbo.RptLedger_Sub.Credit_Or_Debit, dbo.ACCOUNTS.Account_Code, "
        MySQL = MySQL & "   dbo.RptLedger_Sub.DEV_DES,dbo.RptLedger_Sub.DevDESE, dbo.RptLedger_Sub.Account_Name, dbo.RptLedger_Sub.DEV_ID_Line_No, dbo.RptLedger_Sub.NotesTypeName,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.UserID, dbo.RptLedger_Sub.UserName, dbo.RptLedger_Sub.RecordDate, dbo.RptLedger_Sub.ReceiptID, dbo.RptLedger_Sub.Notes_ID,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.OperaID, dbo.RptLedger_Sub.Transaction_ID, dbo.RptLedger_Sub.Transaction_Serial, dbo.RptLedger_Sub.Transaction_Date,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Posted, dbo.RptLedger_Sub.TransactionTypeName, dbo.RptLedger_Sub.PostedDate, dbo.RptLedger_Sub.PostedUserID,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Account_Interval_ID, dbo.RptLedger_Sub.NoteDate, dbo.RptLedger_Sub.NoteType, dbo.RptLedger_Sub.NoteSerial,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Note_Value, dbo.RptLedger_Sub.Account_Serial, dbo.RptLedger_Sub.Account_NameEng, dbo.RptLedger_Sub.Parent_Account_Code,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Branch, dbo.RptLedger_Sub.Sum_account, dbo.RptLedger_Sub.cost_center, dbo.RptLedger_Sub.currenct_code, dbo.RptLedger_Sub.Remark,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.note_value_by_characters, dbo.RptLedger_Sub.foxy_no, dbo.RptLedger_Sub.DEV_ID_Line_No1, dbo.RptLedger_Sub.project_id,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.NotesTypeNamee, dbo.RptLedger_Sub.TransactionEnglishName, dbo.RptLedger_Sub.opr_fullcode, dbo.RptLedger_Sub.NoteSerial1,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.branch_id , dbo.RptLedger_Sub.ActivityTypeId, dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type"
        MySQL = MySQL & " , dbo.RptLedger_Sub.notes_all FROM         dbo.RptLedger_Sub RIGHT OUTER JOIN"
        MySQL = MySQL & "  dbo.ACCOUNTS ON dbo.RptLedger_Sub.Account_Code = dbo.ACCOUNTS.Account_Code"
        MySQL = MySQL + " WHERE  ACCOUNTS.Account_Code='" & StrAccountCode & "'"

  
    Else
        MySQL = MySQL + StrAccountCode

    End If



If detailedtransaction = 1 Then
MySQL = "SELECT     dbo.RptLedger_Sub.ChqueNum, dbo.RptLedger_Sub.ManualNo, dbo.RptLedger_Sub.RecordDateH, dbo.RptLedger_Sub.Double_Entry_Vouchers_ID, dbo.RptLedger_Sub.DEV_Value, "
 MySQL = MySQL + " dbo.RptLedger_Sub.Credit_Or_Debit, dbo.ACCOUNTS.Account_Code, dbo.RptLedger_Sub.DEV_DES, dbo.RptLedger_Sub.DevDESE,"
 MySQL = MySQL + " dbo.RptLedger_Sub.Account_Name, dbo.RptLedger_Sub.DEV_ID_Line_No, dbo.RptLedger_Sub.NotesTypeName, dbo.RptLedger_Sub.UserID,"
 MySQL = MySQL + " dbo.RptLedger_Sub.UserName, dbo.RptLedger_Sub.RecordDate, dbo.RptLedger_Sub.ReceiptID, dbo.RptLedger_Sub.Notes_ID, dbo.RptLedger_Sub.OperaID,"
 MySQL = MySQL + " dbo.RptLedger_Sub.Transaction_Serial, dbo.RptLedger_Sub.Transaction_Date, dbo.RptLedger_Sub.Posted, dbo.RptLedger_Sub.TransactionTypeName,"
 MySQL = MySQL + " dbo.RptLedger_Sub.PostedDate, dbo.RptLedger_Sub.PostedUserID, dbo.RptLedger_Sub.Account_Interval_ID, dbo.RptLedger_Sub.NoteDate,"
 MySQL = MySQL + " dbo.RptLedger_Sub.NoteType, dbo.RptLedger_Sub.NoteSerial, dbo.RptLedger_Sub.Note_Value, dbo.RptLedger_Sub.Account_Serial,"
 MySQL = MySQL + " dbo.RptLedger_Sub.Account_NameEng, dbo.RptLedger_Sub.Parent_Account_Code, dbo.RptLedger_Sub.Branch, dbo.RptLedger_Sub.Sum_account,"
 MySQL = MySQL + " dbo.RptLedger_Sub.cost_center, dbo.RptLedger_Sub.currenct_code, dbo.RptLedger_Sub.Remark, dbo.RptLedger_Sub.note_value_by_characters,"
 MySQL = MySQL + " dbo.RptLedger_Sub.foxy_no, dbo.RptLedger_Sub.DEV_ID_Line_No1, dbo.RptLedger_Sub.project_id, dbo.RptLedger_Sub.NotesTypeNamee,"
 MySQL = MySQL + " dbo.RptLedger_Sub.TransactionEnglishName, dbo.RptLedger_Sub.opr_fullcode, dbo.RptLedger_Sub.NoteSerial1, dbo.RptLedger_Sub.branch_id,"
 MySQL = MySQL + " dbo.RptLedger_Sub.ActivityTypeId, dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.RptLedger_Sub.notes_all,"
 MySQL = MySQL + " dbo.RptLedger_Sub.Transaction_ID, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice,"
 MySQL = MySQL + " dbo.Transactions.Transaction_Type, dbo.TransactionTypes.StockEffect, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
 MySQL = MySQL + " dbo.Transaction_Details.unitid , dbo.TblUnites.Unitname, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.discountvalue"
 MySQL = MySQL + " FROM         dbo.TransactionTypes INNER JOIN"
 MySQL = MySQL + " dbo.Transactions ON dbo.TransactionTypes.Transaction_Type = dbo.Transactions.Transaction_Type RIGHT OUTER JOIN"
 MySQL = MySQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
 MySQL = MySQL + " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
 MySQL = MySQL + " dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID RIGHT OUTER JOIN"
 MySQL = MySQL + " dbo.RptLedger_Sub ON dbo.Transactions.Transaction_ID = dbo.RptLedger_Sub.Transaction_ID RIGHT OUTER JOIN"
 MySQL = MySQL + " dbo.ACCOUNTS ON dbo.RptLedger_Sub.Account_Code = dbo.ACCOUNTS.Account_Code"
 MySQL = MySQL + " WHERE     (dbo.ACCOUNTS.Account_Code = '" & StrAccountCode & "') "
 
End If

    If BranchID <> 0 Then
        MySQL = MySQL & " and branch_id = " & BranchID 'Val(FrmAccountingReport.dcBranch.BoundText)
    End If

    If ActivityId <> 0 Then
        MySQL = MySQL & " and dbo.RptLedger_Sub.ActivityTypeId = " & ActivityId 'Val(FrmAccountingReport.DCActivity.BoundText)
    End If

  If project_id <> 0 Then
        MySQL = MySQL & " and dbo.RptLedger_Sub.project_id = " & project_id 'Val(FrmAccountingReport.DCActivity.BoundText)
    End If
    
    
    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and (RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(EndDate, True) & ""
    End If
 
    MySQL = MySQL + " or  RecordDate is null)"
     
    updateProfitsToall

    If SystemOptions.DateOpt = 0 Then
        

        
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Sub-Masster.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Sub-MassterEng.rpt"
        End If
        
        If detailedtransaction = 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                     StrFileName = App.path & "\Reports\" & "Sub-MassterdetailedTransactions.rpt"
                 Else
                     StrFileName = App.path & "\Reports\" & "Sub-MassterdetailedTransactions.rpt"
                 End If
        End If


    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Sub-MassterH.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Sub-MassterEngH.rpt"
        End If

    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
        
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        RsData.Close

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\New Reports\" & "Sub-Masster1.rpt"
        Else
            StrFileName = App.path & "\Reports\New Reports\" & "Sub-MassterEng1.rpt"
        End If

        MySQL = "SELECT     Account_Code, Account_Name, Account_Serial, Account_NameEng, DepitBalance, CreditBalance, Balance, opening_balance, opening_balance_type"
        MySQL = MySQL & " from dbo.ACCOUNTS"
        MySQL = MySQL & "  WHERE     (Account_Code = '" & StrAccountCode & "')"



        RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        'GetMsgs 138, vbExclamation
        '   Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        '   MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '   RsData.Close
        '   Set RsData = Nothing
        '   Screen.MousePointer = vbDefault
        '   Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If d1 <> Null Then
        Me.BegineDate = d1
    End If

    If d2 <> Null Then
        Me.EndDate = d2
    End If

    Dim activityName As String
    Dim branchname As String
    Dim fullcode As String
    branchname = get_branch_name(val(BranchID), activityName)
GetCustomerIdByAccountCodeLong StrAccountCode, fullcode

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(branch_id))
        xReport.ParameterFields(7).AddCurrentValue openingbalanceDes
        'CustomerAgeingData
        xReport.ParameterFields(8).AddCurrentValue CustomerAgeingData





        
       If fullcode <> "" Then
            xReport.ParameterFields(13).AddCurrentValue "ﬂÊœ «·⁄„Ì· : " & fullcode
        Else
            xReport.ParameterFields(13).AddCurrentValue ""
        
     End If
     
        
        If salesPersonName <> "" Then
            xReport.ParameterFields(9).AddCurrentValue "  «·„‰œÊ» " & salesPersonName
        Else
            xReport.ParameterFields(9).AddCurrentValue "    "
        End If

        xReport.ParameterFields(10).AddCurrentValue ShowAgingReport

        If ChartPrintinAS = True Then
            xReport.ParameterFields(11).AddCurrentValue "1"
        Else
            xReport.ParameterFields(11).AddCurrentValue "0"
        End If

        'ShowAgingReport
        StrReportTitle = "ﬂ‘› Õ”«»   " & StrAccountName

If ProjectName <> "" Then
StrReportTitle = StrReportTitle & " ··„‘—Ê⁄  " & ProjectName
End If

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        
            If SystemOptions.DateOpt = 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰ " & Format(Me.BegineDate, "yyyy/Mm/dd") & "  "
            Else
                StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰ " & ToHijriDate(Format(Me.BegineDate, "yyyy/Mm/dd")) & "  "

            End If
    
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            If SystemOptions.DateOpt = 0 Then

                StrReportTitle = StrReportTitle & CHR(13) & "   ≈·Ï " & "   " & Format(EndDate, "yyyy/Mm/dd")
            Else
                StrReportTitle = StrReportTitle & CHR(13) & "   ≈·Ï " & "   " & ToHijriDate(Format(EndDate, "yyyy/Mm/dd"))
            End If
        End If
     
        If ActivityId <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & activityName
        End If

        If BranchID <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & branchname
        End If
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue branchname  'get_branch_name(Val(branchId))
        xReport.ParameterFields(7).AddCurrentValue openingbalanceDes
        xReport.ParameterFields(8).AddCurrentValue CustomerAgeingData
     
       If fullcode <> "" Then
            xReport.ParameterFields(13).AddCurrentValue "Code " & fullcode
        Else
          xReport.ParameterFields(13).AddCurrentValue ""
        
     End If
     
        StrReportTitle = "Statement of ACC :" & account_serial & CHR(13) & StrAccountName & CHR(13)    '& " Account "
    
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From : " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To   :  " & (EndDate)
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(5).AddCurrentValue account_serial
    xReport.ParameterFields(6).AddCurrentValue StrAccountName
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function ShowLedgerComposite(StrAccountCode As String, _
                                    StrAccountName As String, _
                                    Optional manyString As Boolean = False, _
                                    Optional d1 As Date, _
                                    Optional d2 As Date, _
                                    Optional ActivityId As Integer = 0, _
                                    Optional BranchID As Integer = 0, _
                                    Optional CustomerAgeingData As String, _
                                    Optional salesPersonName As String, _
                                    Optional ShowAgingReport As String, _
                                    Optional ByRef account_serial As String) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    'If manyString = False Then
    'MySQL = "Select * From RptLedger_Sub "
    'MySQL = MySQL + " WHERE Account_Code='" & _
    'StrAccountCode & "'"
    'Else
    'MySQL = "Select * From RptLedger_Sub "
    'MySQL = MySQL + StrAccountCode
    'End If

    If manyString = False Then

        MySQL = "SELECT   dbo.ACCOUNTS.Account_Name AS SUBACCOUNTNAME   ,  dbo.RptLedger_Sub.RecordDateH,   dbo.RptLedger_Sub.Double_Entry_Vouchers_ID, dbo.RptLedger_Sub.DEV_Value, dbo.RptLedger_Sub.Credit_Or_Debit, dbo.ACCOUNTS.Account_Code, "
        MySQL = MySQL & "   dbo.RptLedger_Sub.DEV_DES,dbo.RptLedger_Sub.DevDESE, dbo.RptLedger_Sub.Account_Name, dbo.RptLedger_Sub.DEV_ID_Line_No, dbo.RptLedger_Sub.NotesTypeName,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.UserID, dbo.RptLedger_Sub.UserName, dbo.RptLedger_Sub.RecordDate, dbo.RptLedger_Sub.ReceiptID, dbo.RptLedger_Sub.Notes_ID,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.OperaID, dbo.RptLedger_Sub.Transaction_ID, dbo.RptLedger_Sub.Transaction_Serial, dbo.RptLedger_Sub.Transaction_Date,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Posted, dbo.RptLedger_Sub.TransactionTypeName, dbo.RptLedger_Sub.PostedDate, dbo.RptLedger_Sub.PostedUserID,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Account_Interval_ID, dbo.RptLedger_Sub.NoteDate, dbo.RptLedger_Sub.NoteType, dbo.RptLedger_Sub.NoteSerial,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Note_Value, dbo.RptLedger_Sub.Account_Serial, dbo.RptLedger_Sub.Account_NameEng, dbo.RptLedger_Sub.Parent_Account_Code,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Branch, dbo.RptLedger_Sub.Sum_account, dbo.RptLedger_Sub.cost_center, dbo.RptLedger_Sub.currenct_code, dbo.RptLedger_Sub.Remark,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.note_value_by_characters, dbo.RptLedger_Sub.foxy_no, dbo.RptLedger_Sub.DEV_ID_Line_No1, dbo.RptLedger_Sub.project_id,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.NotesTypeNamee, dbo.RptLedger_Sub.TransactionEnglishName, dbo.RptLedger_Sub.opr_fullcode, dbo.RptLedger_Sub.NoteSerial1,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.branch_id , dbo.RptLedger_Sub.ActivityTypeId, dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type"
        MySQL = MySQL & "  FROM         dbo.RptLedger_Sub RIGHT OUTER JOIN"
        MySQL = MySQL & "  dbo.ACCOUNTS ON dbo.RptLedger_Sub.Account_Code = dbo.ACCOUNTS.Account_Code"
        MySQL = MySQL + " WHERE  ACCOUNTS.Account_Code='" & StrAccountCode & "'"
 
    Else
        MySQL = "SELECT  dbo.ACCOUNTS.Account_Name AS SUBACCOUNTNAME   ,   dbo.RptLedger_Sub.RecordDateH,   dbo.RptLedger_Sub.Double_Entry_Vouchers_ID, dbo.RptLedger_Sub.DEV_Value, dbo.RptLedger_Sub.Credit_Or_Debit, dbo.ACCOUNTS.Account_Code, "
        MySQL = MySQL & "   dbo.RptLedger_Sub.DEV_DES,dbo.RptLedger_Sub.DevDESE, dbo.RptLedger_Sub.Account_Name, dbo.RptLedger_Sub.DEV_ID_Line_No, dbo.RptLedger_Sub.NotesTypeName,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.UserID, dbo.RptLedger_Sub.UserName, dbo.RptLedger_Sub.RecordDate, dbo.RptLedger_Sub.ReceiptID, dbo.RptLedger_Sub.Notes_ID,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.OperaID, dbo.RptLedger_Sub.Transaction_ID, dbo.RptLedger_Sub.Transaction_Serial, dbo.RptLedger_Sub.Transaction_Date,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Posted, dbo.RptLedger_Sub.TransactionTypeName, dbo.RptLedger_Sub.PostedDate, dbo.RptLedger_Sub.PostedUserID,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Account_Interval_ID, dbo.RptLedger_Sub.NoteDate, dbo.RptLedger_Sub.NoteType, dbo.RptLedger_Sub.NoteSerial,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Note_Value, dbo.RptLedger_Sub.Account_Serial, dbo.RptLedger_Sub.Account_NameEng, dbo.RptLedger_Sub.Parent_Account_Code,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.Branch, dbo.RptLedger_Sub.Sum_account, dbo.RptLedger_Sub.cost_center, dbo.RptLedger_Sub.currenct_code, dbo.RptLedger_Sub.Remark,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.note_value_by_characters, dbo.RptLedger_Sub.foxy_no, dbo.RptLedger_Sub.DEV_ID_Line_No1, dbo.RptLedger_Sub.project_id,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.NotesTypeNamee, dbo.RptLedger_Sub.TransactionEnglishName, dbo.RptLedger_Sub.opr_fullcode, dbo.RptLedger_Sub.NoteSerial1,"
        MySQL = MySQL & "  dbo.RptLedger_Sub.branch_id , dbo.RptLedger_Sub.ActivityTypeId, dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type"
        MySQL = MySQL & "  FROM         dbo.RptLedger_Sub RIGHT OUTER JOIN"
        MySQL = MySQL & "  dbo.ACCOUNTS ON dbo.RptLedger_Sub.Account_Code = dbo.ACCOUNTS.Account_Code"
        MySQL = MySQL & "  where ACCOUNTS. Account_Code in (" & StrAccountCode & ")"

    End If

    If BranchID <> 0 Then
        MySQL = MySQL & " and branch_id = " & BranchID 'Val(FrmAccountingReport.dcBranch.BoundText)
    End If

    If ActivityId <> 0 Then
        MySQL = MySQL & " and dbo.RptLedger_Sub.ActivityTypeId = " & ActivityId 'Val(FrmAccountingReport.DCActivity.BoundText)
    End If

    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and (RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(EndDate, True) & ""
    End If
 
    MySQL = MySQL + " or  RecordDate is null)"
     
    updateProfitsToall
 
    Dim ss As String
    ss = "update accounts  set opening_balance=" & updatecompositeAccounts(StrAccountCode) & "   where        Account_Code  in (" & StrAccountCode & ")"
    Cn.Execute ss
 
    If SystemOptions.DateOpt = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Sub-Masster2.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Sub-MassterEng2.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Sub-MassterH.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "Sub-MassterEngH.rpt"
        End If

    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        RsData.Close

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "Sub-Masster3.rpt"
        Else
            StrFileName = App.path & "\Reports\ " & "Sub-MassterEng1.rpt"
        End If

        MySQL = "SELECT     Account_Code, Account_Name, Account_Serial, Account_NameEng, DepitBalance, CreditBalance, Balance, opening_balance, opening_balance_type"
        MySQL = MySQL & " from dbo.ACCOUNTS"
        MySQL = MySQL & "  WHERE     (Account_Code in ( " & StrAccountCode & "))"

        RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        'GetMsgs 138, vbExclamation
        '   Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        '   MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '   RsData.Close
        '   Set RsData = Nothing
        '   Screen.MousePointer = vbDefault
        '   Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If d1 <> Null Then
        Me.BegineDate = d1
    End If

    If d2 <> Null Then
        Me.EndDate = d2
    End If

    Dim activityName As String
    Dim branchname As String
    branchname = get_branch_name(val(BranchID), activityName)

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(branch_id))
        xReport.ParameterFields(7).AddCurrentValue openingbalanceDes
        'CustomerAgeingData
        xReport.ParameterFields(8).AddCurrentValue CustomerAgeingData

        If salesPersonName <> "" Then
            xReport.ParameterFields(9).AddCurrentValue "  «·„‰œÊ» " & salesPersonName
        Else
            xReport.ParameterFields(9).AddCurrentValue "    "
        End If

        xReport.ParameterFields(10).AddCurrentValue ShowAgingReport

        If ChartPrintinAS = True Then
            xReport.ParameterFields(11).AddCurrentValue "1"
        Else
            xReport.ParameterFields(11).AddCurrentValue "0"
        End If

        'ShowAgingReport
        StrReportTitle = "ﬂ‘› Õ”«»  „Ã„⁄ " & StrAccountName

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        
            If SystemOptions.DateOpt = 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰ " & Format(Me.BegineDate, "yyyy/Mm/dd") & "  "
            Else
                StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰ " & ToHijriDate(Format(Me.BegineDate, "yyyy/Mm/dd")) & "  "

            End If
    
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            If SystemOptions.DateOpt = 0 Then

                StrReportTitle = StrReportTitle & CHR(13) & "   ≈·Ï " & "   " & Format(EndDate, "yyyy/Mm/dd")
            Else
                StrReportTitle = StrReportTitle & CHR(13) & "   ≈·Ï " & "   " & ToHijriDate(Format(EndDate, "yyyy/Mm/dd"))
            End If
        End If
     
        If ActivityId <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & activityName
        End If

        If BranchID <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & branchname
        End If
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue branchname  'get_branch_name(Val(branchId))
        xReport.ParameterFields(7).AddCurrentValue openingbalanceDes
        xReport.ParameterFields(8).AddCurrentValue CustomerAgeingData
     
        StrReportTitle = "Statement of ACC :" & account_serial & CHR(13) & StrAccountName & CHR(13)    '& " Account "
    
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From : " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To   :  " & (EndDate)
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function ShowLedger2(StrAccountCode As String, _
                            StrAccountName As String, _
                            Optional manyString As Boolean = False, _
                            Optional d1 As Date, _
                            Optional d2 As Date, _
                            Optional StrAccountSerial As String) As Boolean
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    If manyString = False Then
        MySQL = "Select * From RptLedger_Sub "
        MySQL = MySQL + " WHERE Account_Code LIKE '" & StrAccountCode & "%'"
    Else
        MySQL = "Select * From RptLedger_Sub "
        MySQL = MySQL + StrAccountCode
    End If

    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and RecordDate <=" & SQLDate(EndDate, True) & ""
    End If
 
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "Sub-Masster.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Sub-Masster_Eng.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    'xReport.OpenSubreport("GL_cc.rpt").Database.SetDataSource RsData
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If d1 <> Null Then
        Me.BegineDate = d1
    End If

    If d2 <> Null Then
        Me.EndDate = d2
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "ﬂ‘› Õ”«»   " & StrAccountName

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.BegineDate, "yyyy/M/d") & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " ≈·Ï " & Format(EndDate, "yyyy/M/d") & "∫"
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
     
        ' StrReportTitle = "Ledger Report   " '& StrAccountName & " Account"
        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle + " From Date " & (Me.BegineDate) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle + " To Date :  " & (EndDate) & "∫"
        End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(5).AddCurrentValue StrAccountSerial
    xReport.ParameterFields(6).AddCurrentValue StrAccountName
 
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub Class_Initialize()
    Dim RsOpt As ADODB.Recordset
    Set RsOpt = New ADODB.Recordset

    'rsopt.Open "OPTIONS", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    'RPTCompany_Name_Arabic = IIf(IsNull(rsopt!Company_Name_Arabic), "", rsopt!Company_Name_Arabic)   'rs!Company_Name_Arabic
    'RPTComment_Arabic = IIf(IsNull(rsopt!Comment_Arabic), "", rsopt!Comment_Arabic)    'rs!Comment_Arabic
    'RPTCompany_Name_Eng = IIf(IsNull(rsopt!Company_Name_Eng), "", rsopt!Company_Name_Eng)   'rs!Company_Name_Eng
    'RPTComment_Eng = IIf(IsNull(rsopt!Comment_Eng), "", rsopt!Comment_Eng)   'rs!Comment_Eng
    'IntZoomLevel = IIf(IsNull(rsopt("RptZoom").Value), 100, rsopt("RptZoom").Value)     'rs!Comment_Eng
    'RPTCurrency = IIf(IsNull(rsopt("currency_unite").Value), "", rsopt("currency_unite").Value)    'rs!Comment_Eng
    'rsopt.Close
    'Set rsopt = Nothing
End Sub

Public Property Get BegineDate() As Variant
    BegineDate = m_BegineDate
End Property

Public Property Let BegineDate(ByVal vNewValue As Variant)
    m_BegineDate = vNewValue
End Property

Public Property Get EndDate() As Variant
    EndDate = m_EndDate
End Property

Public Property Let EndDate(ByVal vNewValue As Variant)
    m_EndDate = vNewValue
End Property
Public Sub ShowChartAccounts(m_PrintTarget As PrintTarget, Optional account_level As Integer, Optional ByVal IsAll As Boolean = False)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo
Dim FileName As String
FileName = App.path & "\Reports\" & "AccoutGuide.rpt"

    If Dir(FileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    MySQL = "Select * From Accounts "
        If account_level <> 0 Then
        MySQL = MySQL & " WHERE len(Account_Code) - len(replace(Account_Code,'a','')) <= " & account_level
        Else
            MySQL = MySQL & " WHERE 1 = 1 "
        End If
    
        If Not IsAll Then
            MySQL = MySQL + GetAccountCodeHiding
        End If
    
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.UserInterface = EnglishInterface Then
    FileName = App.path & "\Reports\" & "AccoutGuidee.rpt"
        Set xReport = xApp.OpenReport(FileName)
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngComment
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.EngCompanyName
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = "Chart of Accounts"
    Else
FileName = App.path & "\Reports\" & "AccoutGuide.rpt"
        Set xReport = xApp.OpenReport(FileName)
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabComment
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabCompanyName
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = "œ·Ì· «·Õ”«»« "

    End If

    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, FileName, , MySQL ' App.path & "\Reports\" & "AccoutGuide.rpt"
    Set xApp = Nothing
    Set xReport = Nothing
    Screen.MousePointer = vbDefault
End Sub



                
 



Function updateFuncSql()

 On Error Resume Next
 
    Dim sql As String

    sql = "    DROP FUNCTION FN_MAIN_ACCOUNT_SUB_CODES " & CHR(13)

    Cn.Execute sql
Dim s As String
s = " Create FUNCTION [dbo].[FN_MAIN_ACCOUNT_SUB_CODES]" & CHR(13)
s = s & " ("
s = s & "     @FROMCODE       NVARCHAR(20), " & CHR(13)
s = s & "     @TOCODE         NVARCHAR(20), " & CHR(13)
s = s & "     @WithParent     AS BIT " & CHR(13)
s = s & " )"
s = s & " RETURNS @TABLE TABLE (CODE NVARCHAR(20)) " & CHR(13)
s = s & " AS"


s = s & " Begin " & CHR(13)
s = s & "     IF @WithParent = 0 " & CHR(13)
s = s & "     Begin " & CHR(13)
s = s & "         DECLARE @IsMain BIT " & CHR(13)
s = s & "         DECLARE @Level AS SMALLINT " & CHR(13)
s = s & "         SET @IsMain = @WithParent " & CHR(13)
s = s & "         DECLARE @Child_Code NVARCHAR(20) " & CHR(13)
s = s & "         SET @Child_Code = '' " & CHR(13)
s = s & "         DECLARE @NextLevel SMALLINT " & CHR(13)
s = s & "         DECLARE my_Cursor CURSOR " & CHR(13)
s = s & "         FOR"
s = s & "             SELECT T1.Account_Code, " & CHR(13)
s = s & "                    T1.last_account " & CHR(13)
s = s & "             From Accounts " & CHR(13)
s = s & "                    LEFT OUTER JOIN ( " & CHR(13)
s = s & "                             SELECT Parent_Account_Code, " & CHR(13)
s = s & "                                    Account_Code, " & CHR(13)
s = s & "                                    last_account " & CHR(13)
s = s & "                             From Accounts " & CHR(13)

s = s & "                         )T1 " & CHR(13)
s = s & "                         ON  T1.Parent_Account_Code = ACCOUNTS.Account_Code " & CHR(13)
s = s & "             Where Accounts.last_account = 0 " & CHR(13)

s = s & "                    AND ( " & CHR(13)
s = s & "                            (ACCOUNTS.Account_Code >= @FromCode AND @FromCode <> '') " & CHR(13)
s = s & "                            OR (1 = 1 AND @FromCode = '') " & CHR(13)
s = s & "                        ) " & CHR(13)
s = s & "                    AND ( " & CHR(13)
s = s & "                            (ACCOUNTS.Account_Code <= @ToCode AND @ToCode <> '') " & CHR(13)
s = s & "                            OR (1 = 1 AND @ToCode = '') " & CHR(13)
s = s & "                        ) " & CHR(13)
        
s = s & "         OPEN my_Cursor " & CHR(13)
        
s = s & "         FETCH NEXT FROM my_Cursor INTO @Child_Code,@IsMain " & CHR(13)
        
s = s & "         WHILE @@FETCH_STATUS = 0 " & CHR(13)
         
s = s & "         Begin " & CHR(13)
s = s & "             IF @IsMain = 1 " & CHR(13)
s = s & "             Begin " & CHR(13)
s = s & "                 INSERT INTO @Table " & CHR(13)
s = s & "                 Values " & CHR(13)
s = s & "                   ( " & CHR(13)
s = s & "                     @Child_Code " & CHR(13)
s = s & "                   ) " & CHR(13)
s = s & "             End " & CHR(13)
            
s = s & "             IF @IsMain = 0 " & CHR(13)
s = s & "             Begin " & CHR(13)
s = s & "                 SET @NextLevel = @Level + 1 " & CHR(13)
                
s = s & "                 INSERT @Table " & CHR(13)
s = s & "                 SELECT * " & CHR(13)
s = s & "                 FROM   dbo.fn_Main_Account_Sub_Codes (@Child_Code, @Child_Code, @IsMain) " & CHR(13)
s = s & "             End " & CHR(13)
            
s = s & "             FETCH NEXT FROM my_Cursor INTO @Child_Code,@IsMain " & CHR(13)
s = s & "         End " & CHR(13)
        
s = s & "         Close my_Cursor " & CHR(13)

s = s & "         DEALLOCATE my_Cursor " & CHR(13)

s = s & "         Return " & CHR(13)
s = s & "     End " & CHR(13)
s = s & "     Else " & CHR(13)
s = s & "     Begin " & CHR(13)
s = s & "         DECLARE @TblRet TABLE(Code NVARCHAR(20)) " & CHR(13)
s = s & "         DECLARE @vCurrentNodeCode VARCHAR(50) " & CHR(13)
s = s & "         IF @FROMCODE IS NULL " & CHR(13)
s = s & "            OR @FROMCODE = '' " & CHR(13)
s = s & "             Return " & CHR(13)
        
s = s & "         DECLARE CostCentersCurChildCur CURSOR " & CHR(13)
s = s & "         FOR " & CHR(13)
s = s & "             SELECT Account_Code " & CHR(13)
s = s & "             From [dbo].[Accounts] " & CHR(13)
s = s & "             WHERE  Parent_Account_Code = @FROMCODE " & CHR(13)
        
s = s & "         OPEN CostCentersCurChildCur " & CHR(13)
s = s & "         FETCH NEXT FROM CostCentersCurChildCur " & CHR(13)
s = s & "         INTO @vCurrentNodeCode " & CHR(13)
s = s & "         WHILE @@FETCH_STATUS = 0 " & CHR(13)
s = s & "         Begin " & CHR(13)
s = s & "             INSERT INTO @TABLE " & CHR(13)
s = s & "               ( " & CHR(13)
s = s & "                 code " & CHR(13)
s = s & "               ) " & CHR(13)
s = s & "             Values " & CHR(13)
s = s & "               ( " & CHR(13)
s = s & "                 @vCurrentNodeCode " & CHR(13)
s = s & "               ) " & CHR(13)
s = s & "             INSERT INTO @TABLE " & CHR(13)
s = s & "             SELECT Code " & CHR(13)
s = s & "             FROM   dbo.FN_MAIN_ACCOUNT_SUB_CODES(@vCurrentNodeCode,@vCurrentNodeCode,@WithParent) " & CHR(13)
            
s = s & "             FETCH NEXT FROM CostCentersCurChildCur " & CHR(13)
s = s & "             INTO @vCurrentNodeCode " & CHR(13)
s = s & "         End " & CHR(13)
s = s & "         Close CostCentersCurChildCur " & CHR(13)
s = s & "         DEALLOCATE CostCentersCurChildCur " & CHR(13)
s = s & "     End " & CHR(13)
s = s & "     Return " & CHR(13)
    

s = s & " End " & CHR(13)

 Cn.Execute s
End Function





Public Sub ShowChartAccountsxx(m_PrintTarget As PrintTarget, Optional account_level As Integer)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo
Dim FileName As String
FileName = App.path & "\Reports\" & "AccoutGuide.rpt"

    If Dir(FileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    MySQL = "Select * From Accounts "
        If account_level <> 0 Then
        MySQL = MySQL & " WHERE len(Account_Code) - len(replace(Account_Code,'a','')) <= " & account_level
    End If
    
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.UserInterface = EnglishInterface Then
    FileName = App.path & "\Reports\" & "AccoutGuidee.rpt"
        Set xReport = xApp.OpenReport(FileName)
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngComment
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.EngCompanyName
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = "Chart of Accounts"
    Else
FileName = App.path & "\Reports\" & "AccoutGuide.rpt"
        Set xReport = xApp.OpenReport(FileName)
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabComment
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabCompanyName
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = "œ·Ì· «·Õ”«»« "

    End If

    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, FileName, , MySQL ' App.path & "\Reports\" & "AccoutGuide.rpt"
    Set xApp = Nothing
    Set xReport = Nothing
    Screen.MousePointer = vbDefault
End Sub

Public Sub ShowAgingReport(m_PrintTarget As PrintTarget)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo

    If Dir(App.path & "\Reports\" & "AgingREport.rpt") = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
    MySQL = "SELECT     TOP 100 PERCENT dbo.Ageng_type.id ,dbo.TblTempCustomerAging.CustD, SUM(dbo.TblTempCustomerAging.DueValue) AS DuevalueSum, dbo.Ageng_type.Name, CONVERT(varchar(5), "
    MySQL = MySQL & " dbo.Ageng_type.[From]) + ' - ' + CONVERT(varchar(5), dbo.Ageng_type.[To]) AS DES, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee"
    MySQL = MySQL & " FROM         dbo.TblTempCustomerAging LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblCustemers ON dbo.TblTempCustomerAging.CustD = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.Ageng_type ON dbo.TblTempCustomerAging.LateID = dbo.Ageng_type.id"
    MySQL = MySQL & " GROUP BY dbo.TblTempCustomerAging.CustD, dbo.Ageng_type.Name, dbo.Ageng_type.[From], dbo.Ageng_type.[To], dbo.Ageng_type.id, dbo.TblCustemers.CusName,"
    MySQL = MySQL & " dbo.TblCustemers.CusNamee"
    MySQL = MySQL & " ORDER BY dbo.Ageng_type.id "

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.UserInterface = EnglishInterface Then
        Set xReport = xApp.OpenReport(App.path & "\Reports\" & "AgingREporte.rpt")
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngComment
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.EngCompanyName
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = "Aging Report"
    Else

        Set xReport = xApp.OpenReport(App.path & "\Reports\" & "AgingREport.rpt")
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabComment
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabCompanyName
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = "«⁄„«—« ·œÌÊ‰  "

    End If

    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    If SystemOptions.UserInterface = ArabicInterface Then
    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, App.path & "\Reports\" & "AgingREport.rpt"
    Else
    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, App.path & "\Reports\" & "AgingREporte.rpt"
    End If
    Set xApp = Nothing
    Set xReport = Nothing
    Screen.MousePointer = vbDefault
End Sub

Public Function ShowTrialBalance(Optional branch_id As Integer = 0, _
                                 Optional ActivityId As Integer = 0) As Boolean
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "TrialBalance.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "TrialBalance_Eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        MySQL = "SELECT sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as DebitValue ," & "sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as CreditValue, ACCOUNTS.Account_Code," & "ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID =" & "DOUBLE_ENTREY_VOUCHERS.Notes_Id " & " where (((notes.NotePosted) = True) AND (Note_Date <=#" & SQLDate(Me.EndDate) & "#)) " & "Group By ACCOUNTS.Account_Code, ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " Order By ACCOUNTS.Account_Serial"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        MySQL = "Select Sum(DEV_Value1) as DebitValue,Sum(DEV_Value2) as CreditValue,Account_Code, Account_Name,"
        MySQL = MySQL + " Account_Serial , Account_NameEng,opening_balance,opening_balance_type,branch_id,ActivityTypeId "
        MySQL = MySQL + " From"
        MySQL = MySQL + " ( "
        MySQL = MySQL + " SELECT"
        MySQL = MySQL + " DEV_Value1=Case"
        MySQL = MySQL + " When Credit_Or_Debit=0   Then DEV_Value * 1"
        MySQL = MySQL + " Else 0 "
        MySQL = MySQL + " END,"
        MySQL = MySQL + " DEV_Value2=Case"
        MySQL = MySQL + " When Credit_Or_Debit=1  Then DEV_Value * 1"
        MySQL = MySQL + " Else 0"
        MySQL = MySQL + " End"
        MySQL = MySQL + ", Account_Code, Account_Name, Account_Serial,opening_balance,opening_balance_type,branch_id,ActivityTypeId," & "Account_NameEng,Parent_Account_Code"
        MySQL = MySQL + " From dbo.RptLedger_Sub where NoteType<>1000  "

        If ActivityId <> 0 Then
            MySQL = MySQL & " AND  ActivityTypeId=" & ActivityId
        End If

        If branch_id <> 0 Then
            MySQL = MySQL & " AND  branch_id=" & branch_id
        End If
   
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            MySQL = MySQL + " and    RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.EndDate, True) & ""
        End If

        MySQL = MySQL + ")XTable"
        MySQL = MySQL + " Group by Account_Code, Account_Name, Account_Serial, Account_NameEng,opening_balance,opening_balance_type ,branch_id,ActivityTypeId"
    
    End If

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
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
        '   xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        '   xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        Dim StrReportTitle As String

        StrReportTitle = "  „Ì“«‰ «·„—«Ã⁄…    " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
    
        xReport.reporttitle = "Trial Balance Report(Totals and Balances) As AT " & Format(Now, "d/M/yyyy")
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "„Ì“«‰ «·„—«Ã⁄…", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    ShowTrialBalance = False
End Function

Public Function ShowMaterLedgar(StrAccountCode As String, _
                                StrAccountName As String) As Boolean
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer  As ClsReportViewer
    Dim StrFileName As String

    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "Master.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Master_Eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        Exit Function
        'MySQL = "SELECT Sum(iif( DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0," & _
        '"DOUBLE_ENTREY_VOUCHERS.Value * -1 ,DOUBLE_ENTREY_VOUCHERS.Value)) " & _
        '" AS SumValue, DOUBLE_ENTREY_VOUCHERS.Account_Code, ACCOUNTS.Account_Name," & _
        '"IIf( SumValue < 0 ,0,1) AS Credit_Or_Debit ,ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & _
        '" FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & _
        '"ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON " & _
        '"NOTES.Note_ID = DOUBLE_ENTREY_VOUCHERS.Notes_Id " & _
        '" where (((notes.NotePosted) = True))" & _
        '" GROUP BY DOUBLE_ENTREY_VOUCHERS.Account_Code, ACCOUNTS.Account_Name," & _
        '"Account_NameEng,ACCOUNTS.Account_Serial " & _
        '" HAVING (((DOUBLE_ENTREY_VOUCHERS.Account_Code) Like '" & _
        'StrAccountCode & "' & '%'))"
    Else
        MySQL = "Select Sum(DEV_Value1) as SumValue,Account_Code, Account_Name, "
        MySQL = MySQL + " Credit_Or_Debit=Case "
        MySQL = MySQL + " WHEN Sum(DEV_Value1) > 0 THEN 1 "
        MySQL = MySQL + " WHEN Sum(DEV_Value1) < 0 THEN 0 "
        MySQL = MySQL + " ELSE 0 "
        MySQL = MySQL + " End"
        MySQL = MySQL + ", Account_Serial, Account_NameEng"
        MySQL = MySQL + " From "
        MySQL = MySQL + " ( "
        MySQL = MySQL + " SELECT"
        MySQL = MySQL + " DEV_Value1=Case"
        MySQL = MySQL + " When Credit_Or_Debit=0   Then DEV_Value * -1"
        MySQL = MySQL + " When Credit_Or_Debit=1  Then DEV_Value * 1"
        MySQL = MySQL + " End"
        MySQL = MySQL + ", Account_Code, Account_Name, Account_Serial, Account_NameEng,Parent_Account_Code"
        MySQL = MySQL + " From dbo.RptLedger_Sub"
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            MySQL = MySQL + " where     RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.EndDate, True) & ""
        End If
    
        If FrmAccountingReport.dcBranch.BoundText <> "" Then
            MySQL = MySQL & " and branch_id = " & val(FrmAccountingReport.dcBranch.BoundText)
        End If

        If FrmAccountingReport.DCActivity.BoundText <> "" Then
            MySQL = MySQL & " and ActivityTypeId = " & val(FrmAccountingReport.DCActivity.BoundText)
        End If

        MySQL = MySQL + " )XTable"
        MySQL = MySQL + " Where Parent_Account_Code like '" & StrAccountCode & "%'"
    
        MySQL = MySQL + " Group by Account_Code, Account_Name, Account_Serial, Account_NameEng"
    End If

    '
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        ' GetMsgs 138, vbExclamation
        MsgBox "NO data in this Interval", vbCritical
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue "  „œÌ‰ "
        xReport.ParameterFields(2).AddCurrentValue "  œ«∆‰ "
        xReport.ParameterFields(3).AddCurrentValue "  Õ”«» "
        xReport.ParameterFields(4).AddCurrentValue " »«·ﬁÌ„… "
        xReport.ParameterFields(7).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(8).AddCurrentValue RPTComment_Arabic

        '    xReport.ParameterFields(7).AddCurrentValue RPTComment_Arabic
        '    xReport.ParameterFields(8).AddCurrentValue RPTCompany_Name_Arabic
        Dim StrReportTitle As String
        StrReportTitle = ""
   
        StrReportTitle = " ﬁ—Ì— «” «– ⁄«„ ·Õ”«» " & StrAccountName & CHR(13)

        If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
            StrReportTitle = StrReportTitle & CHR(13) & " »œ«Ì… „‰ " & " " & Format(Me.BegineDate, "yyyy/Mm/dd") & "  "
        End If

        If Me.EndDate <> Empty Or Me.EndDate <> Null Then
            StrReportTitle = StrReportTitle & CHR(13) & "   ≈·Ï " & "   " & Format(EndDate, "yyyy/Mm/dd")
        End If
     
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
    Else

        xReport.ParameterFields(1).AddCurrentValue "  Debit "
        xReport.ParameterFields(2).AddCurrentValue "  Credit "
        xReport.ParameterFields(3).AddCurrentValue "  Account "
        xReport.ParameterFields(4).AddCurrentValue " By "
        xReport.ParameterFields(7).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(8).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(10).AddCurrentValue get_branch_name(val(my_branch))
    
        xReport.reporttitle = " General Ledger Report For " & StrAccountName & " Account. From :" & Format(Me.BegineDate, "d/M/yyyy") & " To:" & Format(Me.EndDate, "d/M/yyyy")
    End If

    xReport.ParameterFields(5).AddCurrentValue StrAccountName
    xReport.ParameterFields(6).AddCurrentValue RPTCurrency
    xReport.ParameterFields(9).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "ﬂ‘› ≈” «– ⁄«„", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    MsgBox Err.Description
    ShowMaterLedgar = False
End Function

Public Function ShowDealerBalance(Optional StrDealerCode As String, _
                                  Optional StrAccCode As String) As Boolean

End Function

Public Function ShowIncomeStatment(Optional branch_id As Integer = 0, _
                                   Optional ActivityId As Integer = 0)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\IncomeStatement.rpt"
    Else
        StrFileName = App.path & "\Reports\IncomeStatement_eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        MySQL = "SELECT sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as DebitValue ," & "sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as CreditValue, ACCOUNTS.Account_Code," & "ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID =" & "DOUBLE_ENTREY_VOUCHERS.Notes_Id " & " where (((notes.NotePosted) = True) AND (Note_Date <=#" & SQLDate(Me.EndDate) & "#)) " & "Group By ACCOUNTS.Account_Code, ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " Order By ACCOUNTS.Account_Serial"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then

        'If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        'MySQL = "select * from QryIncomeStatement1('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Date) & "')"
        'End If

        'If Me.Enddate <> Empty Or Me.Enddate <> Null Then
        '   MySQL = "select * from QryIncomeStatement1('" & SQLDate(1 / 1 / 2000) & "','" & SQLDate(Me.Enddate) & "')"
        'End If
        If (Me.BegineDate <> Empty Or Me.BegineDate <> Null) And Me.EndDate <> Empty Or Me.EndDate <> Null Then

            MySQL = "select * from QryIncomeStatement1('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "')"

            If ActivityId <> 0 Then
                MySQL = "select * from QryIncomeStatementActivity('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "' ," & ActivityId & ")"
            End If

            If branch_id <> 0 Then
                MySQL = "select * from QryIncomeStatementBranch('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "' ," & branch_id & ")"
            End If

            'Else
            'MySQL = "Select * From RptQryIncomeStatement"
        End If

        '    MySQL = "Select * From RptQryIncomeStatement "
    
    End If

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        '    xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        '    xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        Dim StrReportTitle As String
        StrReportTitle = "  ﬁ«∆„… «·œŒ·      " & CHR(13) & "   ›Ì «·› —… „‰ " & "   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & " " & Format(Me.EndDate, "yyyy/mm/dd")

        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If
     
        xReport.reporttitle = StrReportTitle
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
     
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo
  
        xReport.reporttitle = "Income Statment :  " & Format(Me.BegineDate, "d/M/yyyy") & " To: " & Format(Me.EndDate, "d/M/yyyy")
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "ﬁ«∆„… «·œŒ·", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

End Function

Public Function ShowBalanceSheetNew()
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\BalanceSheet.rpt"
    Else
        StrFileName = App.path & "\Reports\BalanceSheet_eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        MySQL = "SELECT sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as DebitValue ," & "sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as CreditValue, ACCOUNTS.Account_Code," & "ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID =" & "DOUBLE_ENTREY_VOUCHERS.Notes_Id " & " where (((notes.NotePosted) = True) AND (Note_Date <=#" & SQLDate(Me.EndDate) & "#)) " & "Group By ACCOUNTS.Account_Code, ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " Order By ACCOUNTS.Account_Serial"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
 
        MySQL = "SELECT     dbo.BalanceSheetReport.ComparisonYearCredit, dbo.BalanceSheetReport.ComparisonYearTotal, dbo.BalanceSheetReport.ComparisonYearDepit, "
        MySQL = MySQL & "   dbo.BalanceSheetReport.Balance, dbo.BalanceSheetReport.CreditValue, dbo.BalanceSheetReport.DebitValue, dbo.BalanceSheetReport.Account_Code, "
        MySQL = MySQL & " dbo.BalanceSheetReport.id, dbo.BalanceSheetReport.totals, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Branch, dbo.ACCOUNTS.Parent_Account_Code, dbo.ACCOUNTS.last_account,"
        MySQL = MySQL & " dbo.ACCOUNTS.account_serial , dbo.ACCOUNTS.BasicAccount, dbo.ACCOUNTS.Account_NameEng"
        MySQL = MySQL & " FROM         dbo.ACCOUNTS INNER JOIN"
        MySQL = MySQL & "  dbo.BalanceSheetReport ON dbo.ACCOUNTS.Account_Code = dbo.BalanceSheetReport.Account_Code"
    End If

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        ' xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        xReport.reporttitle = "«·„Ì“«‰Ì… «·⁄„Ê„ÌÂ "
   '     Me.BegineDate = BaklanceSheetvIEW.DTPickerAccFrom
   '     Me.EndDate = BaklanceSheetvIEW.DTPickerAccTo
   '     xReport.reporttitle = "«·„Ì“«‰Ì… «·⁄„Ê„ÌÂ  „‰   " & Format(Me.BegineDate, "d/M/yyyy") & "   «·Ï  :" & Format(Me.EndDate, "d/M/yyyy")

    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
     
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo
        xReport.reporttitle = "Balance Sheet  From:  " & Format(Me.BegineDate, "d/M/yyyy") & "  To :" & Format(Me.EndDate, "d/M/yyyy")
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "«·„Ì“«‰Ì… «·⁄„Ê„ÌÂ", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function

Public Function ShowFinancialEquationsHistory(FinancialEquationsId As Integer)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\FinancialEquations\f1.rpt"
    Else
        StrFileName = App.path & "\Reports\FinancialEquations\f1.rpt"
    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        MySQL = "SELECT sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as DebitValue ," & "sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as CreditValue, ACCOUNTS.Account_Code," & "ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID =" & "DOUBLE_ENTREY_VOUCHERS.Notes_Id " & " where (((notes.NotePosted) = True) AND (Note_Date <=#" & SQLDate(Me.EndDate) & "#)) " & "Group By ACCOUNTS.Account_Code, ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " Order By ACCOUNTS.Account_Serial"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
 
        MySQL = "SELECT     dbo.FinancialEquationsHistory.FinancialEquationsId, dbo.FinancialEquationsHistory.fromdate, dbo.FinancialEquationsHistory.ToDate, "
        MySQL = MySQL & "    dbo.FinancialEquationsHistory.Finalvalue, dbo.FinancialEquations.FinancialEquationsName, dbo.FinancialEquations.FinancialEquationsDes,"
        MySQL = MySQL & "          dbo.FinancialEquations.FinancialEquationsCode"
        MySQL = MySQL & "  FROM         dbo.FinancialEquationsHistory INNER JOIN"
        MySQL = MySQL & "  dbo.FinancialEquations ON dbo.FinancialEquationsHistory.FinancialEquationsId = dbo.FinancialEquations.FinancialEquationsId"
        MySQL = MySQL & "  Where (dbo.FinancialEquationsHistory.FinancialEquationsId = " & FinancialEquationsId & ")"
    End If

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData
    xReport.reporttitle = "«· Õ·Ì· «·„«·Ì   "
    Me.BegineDate = FinancialAnalysisView.DTPickerAccFrom
    Me.EndDate = FinancialAnalysisView.DTPickerAccTo

    If SystemOptions.UserInterface = ArabicInterface Then
        ' xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        xReport.reporttitle = "««· Õ·Ì· «·„«·Ì   " & Format(Me.BegineDate, "d/M/yyyy") & "   «·Ï  :" & Format(Me.EndDate, "d/M/yyyy")

    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
     
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo
        xReport.reporttitle = "Financial Equations    " & Format(Me.BegineDate, "d/M/yyyy") & "  To :" & Format(Me.EndDate, "d/M/yyyy")
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "«· Õ·Ì· «·„«·Ì  ", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function

Public Function ArrowsHistorydata(sql As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String

    'On Error GoTo ErrTrap
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ArrowsFollow\Repor1.rpt"
    Else
        StrFileName = App.path & "\Reports\ArrowsFollow\Repor1.rpt"
    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    MySQL = sql

    If SystemOptions.SysDataBaseType = AccessDataBase Then
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        MySQL = sql
        'Me.BegineDate = ArrowsHistory.DTPickerAccFrom.value
        'Me.EndDate = ArrowsHistory.DTPickerAccTo.value
 
    End If

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        '    xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        xReport.reporttitle = "   ﬁ—Ì— „ «»⁄Â «·«”Â„ "
        xReport.reporttitle = " ﬁ—Ì— „ «»⁄Â «·«”Â„   " & Format(Me.BegineDate, "d/M/yyyy") & "   «·Ï  :" & Format(Me.EndDate, "d/M/yyyy")

    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
     
        ' Me.BegineDate = ArrowsHistory.DTPickerAccFrom
        '  Me.Enddate = ArrowsHistory.DTPickerAccTo
        xReport.reporttitle = " Arrows Price History " & Format(Me.BegineDate, "d/M/yyyy") & "  To :" & Format(Me.EndDate, "d/M/yyyy")
    End If

    'xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, " ﬁ—Ì— „ «»⁄Â «·«”Â„", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function

Public Function ShowOpeningBalances(Optional branch_id As Integer = 0, _
                                    Optional ActivityId As Integer = 0, _
                                    Optional StrAccountCode As String, _
                                    Optional account_level As Integer = 0, _
                                    Optional HideZeroBalance As Integer = True)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ShowOpeningBalances.rpt"
    Else
        StrFileName = App.path & "\Reports\ShowOpeningBalancesENG.rpt"
    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
 
    MySQL = "SELECT     TOP 100 PERCENT Account_Code, Account_Name, last_account, Account_NameEng,  ISNULL(opening_balance, 0) AS opening_balance  ,   ISNULL(Balance, 0) AS Balance ,Account_Serial,Parent_Account_Code"
    MySQL = MySQL & " from dbo.ACCOUNTS"
    'If StrAccountCode = "" Then
    'MySQL = MySQL & " Where (last_account<>1 ) and  (AccountTypes = 1)"
    'Else
    'MySQL = MySQL & " Where Account_Code like'" & StrAccountCode & "%'"
    'End If
    MySQL = MySQL & " Where 1=1"

    If account_level <> 0 Then
        MySQL = MySQL & " and len(account_code) - len(replace(account_code,'a','')) <= " & account_level
    End If

    If HideZeroBalance = vbYes Then
        MySQL = MySQL & " and (opening_balance)<>0"
    End If

    MySQL = MySQL & "  ORDER BY Account_Serial"

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        ' xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
    
        Dim StrReportTitle As String

        StrReportTitle = "  «·«—’œ… «·«ﬁ  «ÕÌ…   " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
     
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo

        StrReportTitle = "     Opening Balances Auditing   " & CHR(13) & "  From " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    TO     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Activity  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Branch  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "«·„Ì“«‰Ì…", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function

Public Function ShowOpeningBalances2(Optional branch_id As Integer = 0, _
                                    Optional ActivityId As Integer = 0, _
                                    Optional StrAccountCode As String, _
                                    Optional account_level As Integer = 0, _
                                    Optional HideZeroBalance As Integer = True)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ShowOpeningBalances.rpt"
    Else
        StrFileName = App.path & "\Reports\ShowOpeningBalancesENG.rpt"
    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
 
    MySQL = "SELECT     dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_Name, dbo.DOUBLE_ENTREY_VOUCHERS1.[Value], dbo.DOUBLE_ENTREY_VOUCHERS1.Credit_Or_Debit, "
                      MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id , dbo.TblBranchesData.ActivityTypeId, dbo.ACCOUNTS.Account_Code, dbo.ACCOUNTS.Parent_Account_Code"
MySQL = MySQL & "  FROM         dbo.ACCOUNTS INNER JOIN"
MySQL = MySQL & "                       dbo.DOUBLE_ENTREY_VOUCHERS1 ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code INNER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id = dbo.TblBranchesData.branch_id"
 


    MySQL = MySQL & " Where 1=1"
If branch_id <> 0 Then
MySQL = MySQL & "  and  (dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id = " & branch_id & ")"
End If
 
 
 If ActivityId <> 0 Then
MySQL = MySQL & "   And (dbo.TblBranchesData.ActivityTypeId =  " & ActivityId & ")"
End If

    MySQL = MySQL & "  ORDER BY Account_Serial"

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        ' xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
    
        Dim StrReportTitle As String

        StrReportTitle = "  «·«—’œ… «·«ﬁ  «ÕÌ…   " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
     
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo

        StrReportTitle = "     Opening Balances Auditing   " & CHR(13) & "  From " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    TO     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Activity  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Branch  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "«·„Ì“«‰Ì…", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function
Public Function RecursiveUpdate(StrAccountCode As String)
    Dim ParentAccount As String
    Dim StrSQL As String
    ParentAccount = Get_Account_Parent_code(StrAccountCode)

    If ParentAccount = "r" Then
        Exit Function
    Else
        StrSQL = " update ACCOUNTS"

        If totalprofit > 0 Then
            StrSQL = StrSQL & " SET opening_balance=opening_balance+" & totalprofit * -1
        Else
            StrSQL = StrSQL & " SET opening_balance=opening_balance+" & totalprofit * -1
        End If
               
        StrSQL = StrSQL & "  where  Account_code ='" & ParentAccount & "'"
        Cn.Execute StrSQL
            
        RecursiveUpdate (ParentAccount)
    End If

End Function

Public Function updatecompositeAccounts(StrAccounts As String) As Double
    Dim StrSQL As String

    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = "SELECT      SUM(opening_balance) AS totalopeningBalance"
    sql = sql + "  from dbo.ACCOUNTS"
 
    sql = sql + "   where       Account_Code  in (" & StrAccounts & ")"
    '       MySQL = MySQL & "  where ACCOUNTS. Account_Code in (" & StrAccountCode & ")"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        updatecompositeAccounts = IIf(IsNull(rs("totalopeningBalance").value), 0, rs("totalopeningBalance").value)
    Else
        updatecompositeAccounts = 0
    End If

End Function
 
Public Function updateProfitsToall()
    Dim StrSQL As String
    Dim Account_code  As String
    Account_code = get_account_code_branch(49, my_branch)

    If Account_code = "NO branch" Or Account_code = "NO account" Then
         
    Else
        StrSQL = " update ACCOUNTS"

        If totalprofit > 0 Then
            StrSQL = StrSQL & " SET opening_balance=opening_balance+" & totalprofit * -1
        Else
            StrSQL = StrSQL & " SET opening_balance=opening_balance+" & totalprofit * -1
        End If
               
        StrSQL = StrSQL & "  where  Account_code ='" & Account_code & "'"
        Cn.Execute StrSQL
         
        RecursiveUpdate (Account_code)
         
        '     StrSQL = " update ACCOUNTS"
        '           If totalprofit > 0 Then
        '       StrSQL = StrSQL & " SET opening_balance=opening_balance+" & totalprofit * -1
        '       Else
        '       StrSQL = StrSQL & " SET opening_balance=opening_balance+" & totalprofit * -1
        '       End If
        '
        '       StrSQL = StrSQL & "  where  Account_code ='" & "a2" & "'"
        '    Cn.Execute StrSQL
            
    End If

End Function

Public Function ShowBalanceSheet(Optional branch_id As Integer = 0, _
                                 Optional ActivityId As Integer = 0, _
                                 Optional StrAccountCode As String, _
                                 Optional account_level As Integer = 0, _
                                 Optional HideZeroBalance As Integer = True)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\BalanceSheetold.rpt"
    Else
        StrFileName = App.path & "\Reports\BalanceSheetoldeng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        MySQL = "SELECT sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as DebitValue ," & "sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as CreditValue, ACCOUNTS.Account_Code," & "ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID =" & "DOUBLE_ENTREY_VOUCHERS.Notes_Id " & " where (((notes.NotePosted) = True) AND (Note_Date <=#" & SQLDate(Me.EndDate) & "#)) " & "Group By ACCOUNTS.Account_Code, ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " Order By ACCOUNTS.Account_Serial"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        '  If account_level = 0 Then
        MySQL = "SELECT QryBalanceSheet.* FROM dbo.QryBalanceSheet() QryBalanceSheet"

        If (Me.BegineDate <> Empty Or Me.BegineDate <> Null) And Me.EndDate <> Empty Or Me.EndDate <> Null Then

            MySQL = "select * from QryBalanceSheet('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "')"

            If ActivityId <> 0 Then
 
                MySQL = "select * from QryBalanceSheetActivity('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "' ," & ActivityId & ")"
            End If

            If branch_id <> 0 Then
                MySQL = "select * from QryBalanceSheetBranch('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "' ," & branch_id & ")"
            End If
        End If

        '  Else
        '  MySQL = "SELECT QryBalanceSheet.* FROM dbo.QryBalanceSheet() QryBalanceSheet  where len(account_code)<=" & account_level * 2
        '  End If
 
    End If

    MySQL = "SELECT     TOP 100 PERCENT Account_Code, Account_Name, last_account, Account_NameEng,  ISNULL(opening_balance, 0) AS opening_balance  ,   ISNULL(Balance, 0) AS Balance ,Account_Serial,Parent_Account_Code"
    MySQL = MySQL & " from dbo.ACCOUNTS"

    If StrAccountCode = "" Then
        MySQL = MySQL & " Where (last_account<>1 ) and  (AccountTypes = 1)"
    Else
        MySQL = MySQL & " Where Account_Code like'" & StrAccountCode & "%'"
    End If

    If account_level <> 0 Then
        MySQL = MySQL & " and len(account_code) - len(replace(account_code,'a','')) <= " & account_level
    End If

    If HideZeroBalance = vbYes Then
        MySQL = MySQL & " and (opening_balance+Balance)<>0 AND   Balance<>0 "
    End If

    MySQL = MySQL & "  ORDER BY Account_Serial"

    updateProfitsToall
            
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        ' xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
    
        Dim StrReportTitle As String

        StrReportTitle = "  «·„Ì“«‰Ì…   " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)

        StrReportTitle = "  Balance Sheet   " & CHR(13) & " From   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    To     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Activity  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Branch  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
     
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "«·„Ì“«‰Ì…", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function
 
 
 Public Function ShowTrialBalanceNew(Optional branch_id As Integer = 0, _
                                    Optional ActivityId As Integer = 0, _
                                    Optional StrAccountCode As String, _
                                    Optional account_level As Integer = 0, _
                                    Optional HideZeroBalance As Integer = True, _
                                    Optional Contins As String = "0", _
                                    Optional laseAcc As Boolean = True, Optional Account_code As String = "", Optional showlastlevel As Boolean = False, Optional HideMasterAcc As Integer = 0)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap
If laseAcc = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\TrialBalanceNew.rpt"
            Else
                StrFileName = App.path & "\Reports\TrialBalanceNewEng.rpt"
            End If
Else
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\TrialBalanceNew1.rpt"
            Else
                StrFileName = App.path & "\Reports\TrialBalanceNewEng1.rpt"
            End If


End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

   ' updateProfitsToall
    MySQL = "SELECT     TOP 100 PERCENT Account_Code, Account_Name, last_account, Account_NameEng,ISNULL(CreditBalance, 0) AS CreditBalance,   ISNULL(DepitBalance, 0) AS DepitBalance ,  ISNULL(opening_balance, 0) AS opening_balance  ,   ISNULL(Balance, 0) AS Balance ,Account_Serial,Parent_Account_Code"
    MySQL = MySQL & " from dbo.ACCOUNTS where 1=1"

If Account_code <> "" Then
MySQL = MySQL & "and  Account_Code like'" & Account_code & "%'"
End If

If showlastlevel = False Then
    If laseAcc = True Then
        MySQL = MySQL & " and (last_account=1 )  "
    Else
        MySQL = MySQL & " and (last_account=0 )  "
    End If
End If

 
    If account_level <> 0 Then
        MySQL = MySQL & " and len(account_code) - len(replace(account_code,'a',''))  <= " & account_level
    End If

    If HideZeroBalance = vbYes Then
        MySQL = MySQL & "    and not (  opening_balance =0  and DepitBalance=0 and CreditBalance=0)"
    End If

If HideMasterAcc = vbYes Then
   MySQL = MySQL & " and (last_account=1 )  "
End If
MySQL = MySQL & GetAccountByBarnchUser
    'MySQL = MySQL & "    and not (  Account_Code in (" & GetAllProfitsAccounts & ") )"
 
    MySQL = MySQL & "  ORDER BY Account_Serial"

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass

    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        ' xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
    
        Dim StrReportTitle As String

        If laseAcc = True Then
            StrReportTitle = "  „Ì“«‰ «·„—«Ã⁄Â    " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
        Else
            StrReportTitle = "  „Ì“«‰ «·„—«Ã⁄Â  »«·„” ÊÌ«  " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
        End If
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If
        If val(FrmAccountingReport.DCRegionID.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··„‰ÿﬁ…  " & FrmAccountingReport.DCRegionID.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
     
        If Contins = "0" Then
            xReport.ParameterFields(2).AddCurrentValue "0"
        Else
            xReport.ParameterFields(2).AddCurrentValue "1"
        End If
     
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo

        If laseAcc = True Then
            StrReportTitle = "  Trial Balance  " & CHR(13) & "  From  " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    To     " & Format(Me.EndDate, "yyyy/mm/dd")
        Else
            StrReportTitle = "  Trial Balance  by Levels " & CHR(13) & "  From  " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    To     " & Format(Me.EndDate, "yyyy/mm/dd")

        End If
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Activity  " & FrmAccountingReport.DCActivity.Text
        End If
        If val(FrmAccountingReport.DCRegionID.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Region  " & FrmAccountingReport.DCRegionID.Text
        End If
        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Branch  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "„Ì“«‰ «·„—«Ã⁄Â", , , , StrFileName, , MySQL
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function


Public Function ShowTrialBalanceNew2(Optional branch_id As Integer = 0, _
                                     Optional ActivityId As Integer = 0, _
                                     Optional StrAccountCode As String, _
                                     Optional account_level As Integer = 0, _
                                     Optional HideZeroBalance As Integer = True, _
                                     Optional Parent_Account_Code As String, _
                                     Optional AccountsCodes As String, Optional AccountName As String = "")

    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\TrialBalanceNew.rpt"
    Else
        StrFileName = App.path & "\Reports\TrialBalanceNeweng.rpt"
    End If

    If AccountsCodes <> "" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\New Reports\EmployeeStatement.rpt"
        Else
            StrFileName = App.path & "\Reports\New Reports\EmployeeStatementeng.rpt"
        End If

    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
 
    MySQL = "SELECT     TOP 100 PERCENT Account_Code, Account_Name, last_account, Account_NameEng,ISNULL(CreditBalance, 0) AS CreditBalance,   ISNULL(DepitBalance, 0) AS DepitBalance ,  ISNULL(opening_balance, 0) AS opening_balance  ,   ISNULL(Balance, 0) AS Balance ,Account_Serial,Parent_Account_Code"
    MySQL = MySQL & " from dbo.ACCOUNTS"

    If AccountsCodes <> "" Then
        MySQL = MySQL & " Where (last_account=1 )   and Account_Code in (" & AccountsCodes & ")"
    Else
 
        MySQL = MySQL & " Where (last_account=1 )   and Parent_Account_Code='" & Parent_Account_Code & "'"
    End If

    If HideZeroBalance = vbYes Then
        MySQL = MySQL & "    and not (  opening_balance =0 and Balance=0  and DepitBalance=0 and CreditBalance=0)"
    End If
 
    MySQL = MySQL & "  ORDER BY Account_Serial"

    updateProfitsToall

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·« ÌÊÃœ »Ì«‰« ", vbCritical
Else
MsgBox "No Data in this interval", vbCritical
End If
'        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        ' xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
    
        Dim StrReportTitle As String

        StrReportTitle = "    «” «– ⁄«„ »«·Õ—ﬂ«    : " & AccountName & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
     
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo
     
        StrReportTitle = "   General Ledger By Transactions  " & CHR(13) & "  From   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    To     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Activity  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Branch  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    End If

    If AccountsCodes <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
  
            StrReportTitle = "     ﬂ‘› Õ”«»«  „ÊŸ› " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
  
            If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
            End If

            If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
            End If

            xReport.reporttitle = StrReportTitle
 
        Else
            StrReportTitle = "   Employee Accounts " & CHR(13) & "  From  " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    T0     " & Format(Me.EndDate, "yyyy/mm/dd")
  
            If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & "  Activity  " & FrmAccountingReport.DCActivity.Text
            End If

            If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & "  Branches  " & FrmAccountingReport.dcBranch.Text
            End If

            xReport.reporttitle = StrReportTitle
        End If
 
    End If
 
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "    ", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function


Public Function ShowTrialBalanceNew3(Optional branch_id As Integer = 0, _
                                     Optional ActivityId As Integer = 0, _
                                     Optional StrAccountCode As String, _
                                     Optional account_level As Integer = 0, _
                                     Optional HideZeroBalance As Integer = True, _
                                     Optional Parent_Account_Code As String, _
                                     Optional AccountsCodes As String, Optional ProjectName As String)

    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\construction\TrialBalanceNew.rpt"
    Else
        StrFileName = App.path & "\Reports\construction\TrialBalanceNeweng.rpt"
    End If

    If AccountsCodes <> "" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\construction\EmployeeStatement.rpt"
        Else
            StrFileName = App.path & "\Reports\construction\EmployeeStatementeng.rpt"
        End If

    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
 
    MySQL = "SELECT     TOP 100 PERCENT Account_Code, Account_Name, last_account, Account_NameEng,ISNULL(CreditBalance, 0) AS CreditBalance,   ISNULL(DepitBalance, 0) AS DepitBalance ,  ISNULL(opening_balance, 0) AS opening_balance  ,   ISNULL(Balance, 0) AS Balance ,Account_Serial,Parent_Account_Code"
    MySQL = MySQL & " from dbo.ACCOUNTS"

    If AccountsCodes <> "" Then
        MySQL = MySQL & " Where (last_account=1 )   and Account_Code in (" & AccountsCodes & ")"
    Else
 
        MySQL = MySQL & " Where (last_account=1 )   and Parent_Account_Code='" & Parent_Account_Code & "'"
    End If

    If HideZeroBalance = vbYes Then
        MySQL = MySQL & "    and not (  opening_balance =0 and Balance=0  and DepitBalance=0 and CreditBalance=0)"
    End If
 
    MySQL = MySQL & "  ORDER BY Account_Serial"

    updateProfitsToall

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        ' xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

     xReport.ParameterFields(5).AddCurrentValue ProjectName
    
        Dim StrReportTitle As String

        StrReportTitle = "    «” «– ⁄«„ »«·Õ—ﬂ«   " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
    '    xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
     
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo
     
        StrReportTitle = "   General Ledger By Transactions  " & CHR(13) & "  From   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    To     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Activity  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Branch  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    End If

    If AccountsCodes <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
  
            StrReportTitle = "     ﬂ‘› Õ”«»«  „‘—Ê⁄  " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
  
            If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
            End If

            If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
            End If

            xReport.reporttitle = StrReportTitle
 
        Else
            StrReportTitle = "   Employee Accounts " & CHR(13) & "  From  " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    T0     " & Format(Me.EndDate, "yyyy/mm/dd")
  
            If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & "  Activity  " & FrmAccountingReport.DCActivity.Text
            End If

            If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
                StrReportTitle = StrReportTitle & CHR(13) & "  Branches  " & FrmAccountingReport.dcBranch.Text
            End If

            xReport.reporttitle = StrReportTitle
        End If
 
    End If
 
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "    ", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function
 Public Function ShowIncomeStatementnew(Optional branch_id As Integer = 0, _
                                       Optional ActivityId As Integer = 0, _
                                       Optional StrAccountCode As String, Optional Details As Boolean = False)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap
If Details = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\IncomeStatementNew.rpt"
    Else
        StrFileName = App.path & "\Reports\IncomeStatementNewEng.rpt"
    End If

Else
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\IncomeStatementNew1.rpt"
    Else
        StrFileName = App.path & "\Reports\IncomeStatementNewEng1.rpt"
    End If



End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        MySQL = "SELECT sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as DebitValue ," & "sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as CreditValue, ACCOUNTS.Account_Code," & "ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID =" & "DOUBLE_ENTREY_VOUCHERS.Notes_Id " & " where (((notes.NotePosted) = True) AND (Note_Date <=#" & SQLDate(Me.EndDate) & "#)) " & "Group By ACCOUNTS.Account_Code, ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " Order By ACCOUNTS.Account_Serial"
    
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        '  If account_level = 0 Then
        MySQL = "SELECT QryBalanceSheet.* FROM dbo.QryBalanceSheet() QryBalanceSheet"

        If (Me.BegineDate <> Empty Or Me.BegineDate <> Null) And Me.EndDate <> Empty Or Me.EndDate <> Null Then

            MySQL = "select * from QryBalanceSheet('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "')"

            If ActivityId <> 0 Then
 
                MySQL = "select * from QryBalanceSheetActivity('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "' ," & ActivityId & ")"
            End If

            If branch_id <> 0 Then
                MySQL = "select * from QryBalanceSheetBranch('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "' ," & branch_id & ")"
            End If
        End If

        '  Else
        '  MySQL = "SELECT QryBalanceSheet.* FROM dbo.QryBalanceSheet() QryBalanceSheet  where len(account_code)<=" & account_level * 2
        '  End If
 
    End If
 
    MySQL = "SELECT     TOP 100 PERCENT Account_Code, Account_Name, last_account, Account_NameEng,  ISNULL(opening_balance, 0) AS opening_balance  ,   ISNULL(Balance, 0) AS Balance ,Account_Serial,Parent_Account_Code"
    MySQL = MySQL & " from dbo.ACCOUNTS"

    If StrAccountCode = "" Then
        MySQL = MySQL & " Where (last_account<>1 ) and  (AccountTypes = 2)"
    Else
        MySQL = MySQL & " Where Account_Code like'" & StrAccountCode & "%'"
    End If
   MySQL = MySQL & GetAccountByBarnchUser
    MySQL = MySQL & "  ORDER BY Account_Serial"
    
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        ' xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
    
        Dim StrReportTitle As String

        
    If Details = False Then
    StrReportTitle = "  ﬁ«∆„… «·œŒ· „Ã„⁄Â  " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
    Else
    StrReportTitle = "  ﬁ«∆„… «·œŒ· ·› —… „Õœœ…  " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
    End If
    
    
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
            xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  'RPTCompany_Name_Arabic
   
       ' xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
     
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo
      If Details = False Then
        StrReportTitle = "   Income Statement     " & CHR(13) & "   From   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    To     " & Format(Me.EndDate, "yyyy/mm/dd")
    Else
    
    StrReportTitle = "   Income Statement  For Interval   " & CHR(13) & "   From   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    To     " & Format(Me.EndDate, "yyyy/mm/dd")
    End If
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Activity  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Branch  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "  ", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function

 
 

Public Function ShowIncomeStatementmonthly2(Optional branch_id As Integer = 0, _
                                       Optional ActivityId As Integer = 0, _
                                       Optional StrAccountCode As String, Optional Details As Boolean = False)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
     On Error GoTo ErrTrap
If Details = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\IncomeStatementintervalsExp.rpt"
    Else
        StrFileName = App.path & "\Reports\IncomeStatementintervalsExp.rpt"
    End If

Else
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\0.rpt"
    Else
        StrFileName = App.path & "\Reports\0.rpt"
    End If



End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

 
    MySQL = "SELECT    dbo.ACCOUNTS.Parent_Account_Code, dbo.ACCOUNTS.Account_Code, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial , "
   MySQL = MySQL & "                       dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.TblBranchesData.branch_name,"
   MySQL = MySQL & "                       dbo.TblBranchesData.branch_namee , dbo.tblActivitesType.Name, dbo.tblActivitesType.NameE"

   MySQL = MySQL & "  ,debit=Case"
   MySQL = MySQL & "  When Credit_Or_Debit=0   Then Value * 1"
   MySQL = MySQL & "  Else 0"
   MySQL = MySQL & "  END,"
   MySQL = MySQL & "  Credit=Case"
   MySQL = MySQL & "  When Credit_Or_Debit=1  Then Value * 1"
   MySQL = MySQL & "  Else 0"
   MySQL = MySQL & "  End"

   MySQL = MySQL & " FROM         dbo.ACCOUNTS INNER JOIN"
   MySQL = MySQL & "                       dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
   MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id INNER JOIN"
   MySQL = MySQL & "                       dbo.tblActivitesType ON dbo.TblBranchesData.ActivityTypeId = dbo.tblActivitesType.id"
   MySQL = MySQL & " WHERE     ( dbo.ACCOUNTS.AccountTab = 3)  "
      If val(branch_id) <> 0 Then
        MySQL = MySQL & "AND (dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = " & branch_id & ")  "
   
    End If

  If val(ActivityId) <> 0 Then
        MySQL = MySQL & " AND (dbo.TblBranchesData.ActivityTypeId = " & ActivityId & ")"
   
    End If

    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and DOUBLE_ENTREY_VOUCHERS.RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and DOUBLE_ENTREY_VOUCHERS.RecordDate <=" & SQLDate(EndDate, True) & ""
    End If

  
  
    MySQL = MySQL & "  ORDER BY Account_Serial"

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData
Dim rs As ADODB.Recordset
    Dim sql As String
    Dim strinteval As String
  Dim i As Integer
        Dim StrReportTitle As String
    If SystemOptions.UserInterface = ArabicInterface Then
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
  
     
        
    If Details = False Then
    StrReportTitle = "  «·„’—Ê›«  ·› —… „Õœœ…  " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
    Else
    StrReportTitle = "  «·„’—Ê›«  ·› —… „Õœœ…  " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
    End If
    
    
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
         
         
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
     
xy:
    If SystemOptions.UserInterface = ArabicInterface Then
    StrReportTitle = "  ﬁ«∆„Â œŒ· „ﬁ«—‰Â ‘Â—Ì«" & CHR(13) & "    „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ì     " & Format(Me.EndDate, "yyyy/mm/dd")
    Else
    StrReportTitle = " Income Statement " & CHR(13) & " From  " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    To     " & Format(Me.EndDate, "yyyy/mm/dd")
    End If
    
    
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Activity  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Branch  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
       
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "  ", , , , StrFileName, , MySQL
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function


 

Public Function ShowIncomeStatementmonthly(Optional branch_id As Integer = 0, _
                                       Optional ActivityId As Integer = 0, _
                                       Optional StrAccountCode As String, Optional Details As Boolean = False)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
     On Error GoTo ErrTrap
If Details = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\IncomeStatementintervals.rpt"
    Else
        StrFileName = App.path & "\Reports\IncomeStatementintervalsE.rpt"
    End If

Else
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\0.rpt"
    Else
        StrFileName = App.path & "\Reports\0.rpt"
    End If



End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

 
    MySQL = "SELECT    dbo.ACCOUNTS.Parent_Account_Code, dbo.ACCOUNTS.Account_Code, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial , "
   MySQL = MySQL & "                       dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.TblBranchesData.branch_name,"
   MySQL = MySQL & "                       dbo.TblBranchesData.branch_namee , dbo.tblActivitesType.Name, dbo.tblActivitesType.NameE"

   MySQL = MySQL & "  ,debit=Case"
   MySQL = MySQL & "  When Credit_Or_Debit=0   Then Value * 1"
   MySQL = MySQL & "  Else 0"
   MySQL = MySQL & "  END,"
   MySQL = MySQL & "  Credit=Case"
   MySQL = MySQL & "  When Credit_Or_Debit=1  Then Value * 1"
   MySQL = MySQL & "  Else 0"
   MySQL = MySQL & "  End"

   MySQL = MySQL & " FROM         dbo.ACCOUNTS INNER JOIN"
   MySQL = MySQL & "                       dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
   MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id INNER JOIN"
   MySQL = MySQL & "                       dbo.tblActivitesType ON dbo.TblBranchesData.ActivityTypeId = dbo.tblActivitesType.id"
   MySQL = MySQL & " WHERE     (dbo.ACCOUNTS.AccountTab = 2 or dbo.ACCOUNTS.AccountTab = 3)  "
      If val(branch_id) <> 0 Then
        MySQL = MySQL & "AND (dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = " & branch_id & ")  "
   
    End If

  If val(ActivityId) <> 0 Then
        MySQL = MySQL & " AND (dbo.TblBranchesData.ActivityTypeId = " & ActivityId & ")"
   
    End If

    If Me.BegineDate <> Empty Or Me.BegineDate <> Null Then
        MySQL = MySQL + " and DOUBLE_ENTREY_VOUCHERS.RecordDate >=" & SQLDate(Me.BegineDate, True) & ""
    End If

    If Me.EndDate <> Empty Or Me.EndDate <> Null Then
        MySQL = MySQL + " and DOUBLE_ENTREY_VOUCHERS.RecordDate <=" & SQLDate(EndDate, True) & ""
    End If

  
  
    MySQL = MySQL & "  ORDER BY Account_Serial"

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData
Dim rs As ADODB.Recordset
    Dim sql As String
    Dim strinteval As String
  Dim i As Integer
        Dim StrReportTitle As String
    If SystemOptions.UserInterface = ArabicInterface Then
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
  
     
        
    If Details = False Then
    StrReportTitle = "  ﬁ«∆„… «·œŒ· „ﬁ«—‰… «·› —«   " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
    Else
    StrReportTitle = "  ﬁ«∆„… «·œŒ· ·› —… „Õœœ…  " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
    End If
    
    
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
         
         
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
     
xy:
    If SystemOptions.UserInterface = ArabicInterface Then
    StrReportTitle = "  ﬁ«∆„Â œŒ· „ﬁ«—‰Â ‘Â—Ì«" & CHR(13) & "    „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ì     " & Format(Me.EndDate, "yyyy/mm/dd")
    Else
    StrReportTitle = " Income Statement " & CHR(13) & " From  " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    To     " & Format(Me.EndDate, "yyyy/mm/dd")
    End If
    
    
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Activity  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  Branch  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
       
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "  ", , , , StrFileName, , MySQL
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function



Public Function ShowGenrealLedgertnew(Optional branch_id As Integer = 0, _
                                      Optional ActivityId As Integer = 0, _
                                      Optional StrAccountCode As String, _
                                      Optional StrAccountName As String)
    Dim MySQL       As String
    Dim RsData      As New ADODB.Recordset
    Dim xApp        As New CRAXDRT.Application
    Dim xReport     As CRAXDRT.Report
    Dim CViewer     As ClsReportViewer
    Dim StrFileName As String
    Dim Msg         As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\GenrealLedger.rpt"
    Else
        StrFileName = App.path & "\Reports\GenrealLedgerEng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        MySQL = "SELECT sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as DebitValue ," & "sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as CreditValue, ACCOUNTS.Account_Code," & "ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID =" & "DOUBLE_ENTREY_VOUCHERS.Notes_Id " & " where (((notes.NotePosted) = True) AND (Note_Date <=#" & SQLDate(Me.EndDate) & "#)) " & "Group By ACCOUNTS.Account_Code, ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " Order By ACCOUNTS.Account_Serial"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        '  If account_level = 0 Then
        MySQL = "SELECT QryBalanceSheet.* FROM dbo.QryBalanceSheet() QryBalanceSheet"

        If (Me.BegineDate <> Empty Or Me.BegineDate <> Null) And Me.EndDate <> Empty Or Me.EndDate <> Null Then

            MySQL = "select * from QryBalanceSheet('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "')"

            If ActivityId <> 0 Then
 
                MySQL = "select * from QryBalanceSheetActivity('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "' ," & ActivityId & ")"
            End If

            If branch_id <> 0 Then
                MySQL = "select * from QryBalanceSheetBranch('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "' ," & branch_id & ")"
            End If
        End If

        '  Else
        '  MySQL = "SELECT QryBalanceSheet.* FROM dbo.QryBalanceSheet() QryBalanceSheet  where len(account_code)<=" & account_level * 2
        '  End If
 
    End If
 
    MySQL = "SELECT     TOP 100 PERCENT Account_Code, Account_Name, last_account, Account_NameEng,  ISNULL(opening_balance, 0) AS opening_balance  , level,  ISNULL(Balance, 0) AS Balance ,Account_Serial,Parent_Account_Code"
    MySQL = MySQL & " from dbo.ACCOUNTS"

    If StrAccountCode = "" Then
        'MySQL = MySQL & " Where (last_account<>1 ) and  (AccountTypes = 2)"
    ElseIf InStr(1, StrAccountCode, ",") > 0 Then
        MySQL = MySQL & " Where Parent_Account_Code in (" & StrAccountCode & ") or Account_Code in (" & StrAccountCode & ") "
    Else
        MySQL = MySQL & " Where Parent_Account_Code ='" & StrAccountCode & "' or Account_Code='" & StrAccountCode & "'"
    End If

    MySQL = MySQL & "  ORDER BY Account_Serial"
    
    
    updateProfitsToall

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        ' xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic

        '     xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.Enddate, Activityid, branch_id)
    
        Dim StrReportTitle As String

        StrReportTitle = "    «” «– ⁄«„  »«·√—’œ… :  " & StrAccountName & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue getprofitValue(Me.BegineDate, Me.EndDate, ActivityId, branch_id)
     
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo
        xReport.reporttitle = "General Ledger    From:  " & Format(Me.BegineDate, "d/M/yyyy") & "  To :" & Format(Me.EndDate, "d/M/yyyy")
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "«” «– ⁄«„  »«·√—’œ… ", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function

Public Function ShowMasterAccountsBalances(Optional branch_id As Integer = 0, _
                                           Optional ActivityId As Integer = 0, _
                                           Optional ByRef StrAccountCode As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\BalanceSheetold.rpt"
    Else
        StrFileName = App.path & "\Reports\BalanceSheet_eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        MySQL = "SELECT sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as DebitValue ," & "sum(iif(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1," & "DOUBLE_ENTREY_VOUCHERS.Value,0) )as CreditValue, ACCOUNTS.Account_Code," & "ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " FROM NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS ON " & "ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID =" & "DOUBLE_ENTREY_VOUCHERS.Notes_Id " & " where (((notes.NotePosted) = True) AND (Note_Date <=#" & SQLDate(Me.EndDate) & "#)) " & "Group By ACCOUNTS.Account_Code, ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial,ACCOUNTS.Account_NameEng " & " Order By ACCOUNTS.Account_Serial"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        '  If account_level = 0 Then
        MySQL = "SELECT QryBalanceSheet.* FROM dbo.QryBalanceSheet() QryBalanceSheet"

        If (Me.BegineDate <> Empty Or Me.BegineDate <> Null) And Me.EndDate <> Empty Or Me.EndDate <> Null Then

            MySQL = "select * from QryBalanceSheet('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "')"

            If ActivityId <> 0 Then
                MySQL = "select * from QryBalanceSheetActivity('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "' ," & ActivityId & ")"
            End If

            If branch_id <> 0 Then
                MySQL = "select * from QryBalanceSheetBranch('" & SQLDate(Me.BegineDate) & "','" & SQLDate(Me.EndDate) & "' ," & branch_id & ")"
            End If
        End If

        '  Else
        '  MySQL = "SELECT QryBalanceSheet.* FROM dbo.QryBalanceSheet() QryBalanceSheet  where len(account_code)<=" & account_level * 2
        '  End If
 
    End If

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.Database.SetDataSource RsData

    If SystemOptions.UserInterface = ArabicInterface Then
        ' xReport.ParameterFields(1).AddCurrentValue RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    
        Dim StrReportTitle As String

        StrReportTitle = "  «·„Ì“«‰Ì…   " & CHR(13) & "   «·› —… „‰   " & Format(Me.BegineDate, "yyyy/mm/dd") & CHR(13) & "    «·Ï     " & Format(Me.EndDate, "yyyy/mm/dd")
  
        If val(FrmAccountingReport.DCActivity.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··‰‘«ÿ  " & FrmAccountingReport.DCActivity.Text
        End If

        If val(FrmAccountingReport.dcBranch.BoundText) <> 0 Then
            StrReportTitle = StrReportTitle & CHR(13) & "  ··›—⁄  " & FrmAccountingReport.dcBranch.Text
        End If

        xReport.reporttitle = StrReportTitle
       
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
     
        Me.BegineDate = FrmAccountingReport.DTPickerAccFrom
        Me.EndDate = FrmAccountingReport.DTPickerAccTo
        xReport.reporttitle = "Balance Sheet  From:  " & Format(Me.BegineDate, "d/M/yyyy") & "  To :" & Format(Me.EndDate, "d/M/yyyy")
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "«·„Ì“«‰Ì…", , , , StrFileName
    Set xApp = Nothing
    Set xReport = Nothing
    'RsData.Close
    'Set RsData = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ ⁄„·Ì… «·ÿ»«⁄…"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
    Msg = Msg & CHR(13) & "—ﬁ„ «·Œÿ« " & Err.Number
    Msg = Msg & CHR(13) & "‰’ «·Œÿ« " & Err.Description
    Msg = Msg & CHR(13) & "„’œ— «·Œÿ« " & Err.Source
    Msg = Msg & CHR(13) & "ClsAccReports-ShowTrialBalance"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function

