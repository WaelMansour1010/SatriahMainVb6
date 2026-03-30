Attribute VB_Name = "SalimNew"


 'Public workWithBarcode As Boolean

Dim sql As String
Public PPointID As Integer
Public CurrentCashireID As Integer
Public X2600 As Date
Public groupcodesPublic As String
Public Strforitems   As String
Public StrforitemsCodes   As String
Public Strforitemsnames   As String
Public groupcodesAll As String
Public firstrun As Boolean
Public Report_Folder As String
Public myGrid As VSFlexGrid
Public P_DTPickerAccFrom As Date
Public P_DTPickerAccTo As Date
Public P_DCActivity As Integer
Public P_DCRegionID As Integer
Public P_dcBranch As Integer
Public onLineMOde As Boolean
 Public onlineservername As String
 Public onlineDataBasename As String
  Public onlinusername As String
   Public onlinepassword As String
    Public onlinebackground As String


Public TempPath As String

Private Const LOCALE_SSHORTDATE = &H1F
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Dim lLocal As Long
Dim Length As Long
Dim lLocal2 As Long
Dim buf As String * 1024

Dim length2 As Long

Dim buf2 As String * 1024

Dim a

Private Const E_POINTER As Long = &H80004003
Private Const S_OK As Long = 0
Private Const INTERNET_MAX_URL_LENGTH As Long = 2048
Private Const URL_ESCAPE_PERCENT As Long = &H1000&

Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeA" ( _
    ByVal pszUrl As String, _
    ByVal pszEscaped As String, _
    ByRef pcchEscaped As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function UrlUnescape Lib "shlwapi" Alias "UrlUnescapeA" ( _
    ByVal pszUrl As String, _
    ByVal pszUnescaped As String, _
    ByRef pcchUnescaped As Long, _
    ByVal dwFlags As Long) As Long
  
  
Public Function GenerateGUID() As String
    ' ≈‰‘«¡ GUID »«” Œœ«„ CreateObject
    GenerateGUID = mId$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
End Function

Function printAnyreport(Optional sql As String, Optional Reportname As String, Optional StrReportTitle As String)

    'Set rs = New ADODB.Recordset
    'rs.Open SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer

    Dim StrFileName As String
    Dim Msg As String



        If SystemOptions.UserInterface = ArabicInterface Then

            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "" & Reportname & ".rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "" & Reportname & "e.rpt"

       End If





    If Dir(StrFileName) = "" Then
     MsgBox " not found reports " & StrFileName
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

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

        StrReportTitle = "" '& StrAccountName

    Else

        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng

        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If


  Dim total As String
  Dim totl As Double

 xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault





End Function


Function getcostbuylastinvoice(Item_ID As Double, _
                               Transaction_Date As Date, _
                               Optional ByVal LngUnitID As Long = 0, _
                                Optional ByRef UnitFactor As Double = 1, _
                                Optional ByRef SecOrder As Integer = 1) As Double
    Dim sql As String
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset


    Dim RsUnitData As ADODB.Recordset
    Dim LngCurItemID As Long
  '  Dim LngUnitID As Long
    Dim QtyBySmalltUnit As Double
'     Dim UnitFactor As Double
'    Dim SecOrder As Integer

        
    LngCurItemID = Item_ID
    LngUnitID = LngUnitID
    QtyBySmalltUnit = 1
    QtyBySmalltUnit = 1
    LngUnitID = 1
    UnitFactor = 1
    SecOrder = 1
    If LngUnitID = 0 Then
        StrSQL = "Select * From TblItemsUnits Where  DefaultUnit = 1 and ItemID=" & LngCurItemID
        'StrSQL = StrSQL + " AND UnitID=" & LngUnitID
        Set RsUnitData = New ADODB.Recordset
        RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
            QtyBySmalltUnit = RsUnitData("UnitFactor").value
            LngUnitID = RsUnitData("UnitID").value
            UnitFactor = RsUnitData("UnitFactor").value
            SecOrder = RsUnitData("SecOrder").value


        End If
    Else
         StrSQL = "Select * From TblItemsUnits Where  UnitId = " & LngUnitID & " and ItemID=" & LngCurItemID
        'StrSQL = StrSQL + " AND UnitID=" & LngUnitID
        Set RsUnitData = New ADODB.Recordset
        RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
            QtyBySmalltUnit = RsUnitData("UnitFactor").value
            LngUnitID = RsUnitData("UnitID").value
            UnitFactor = RsUnitData("UnitFactor").value
            SecOrder = val(RsUnitData("SecOrder").value & "")
        Else
            QtyBySmalltUnit = 1
            LngUnitID = 1
        End If
    
    End If
    
    sql = "  SELECT          ISNULL(dbo.Transaction_Details.Price, 0) AS Price,TblItemsUnits.UnitFactor,TblItemsUnits.SecOrder   "
    sql = sql & "  FROM            dbo.Transactions INNER JOIN                           dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    sql = sql & "  Inner join TblItemsUnits On TblItemsUnits.ItemID =Transaction_Details.Item_ID and Transaction_Details.UnitId = Transaction_Details.UnitID "
    sql = sql & "   WHERE        (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND"
    If LngUnitID <> 0 Then
        'sql = sql & "         (Transaction_Details.UnitID = " & LngUnitID & ") and "
    End If
    sql = sql & "   dbo.Transactions.Transaction_Date IN"
    sql = sql & "                             ("
    sql = sql & "  SELECT        MAX(Transactions_1.Transaction_Date) AS date"
    sql = sql & "             FROM            dbo.Transactions AS Transactions_1 INNER JOIN"
    sql = sql & " dbo.Transaction_Details AS Transaction_Details_1"
    sql = sql & "  ON Transactions_1.Transaction_ID = Transaction_Details_1.Transaction_ID"
    sql = sql & "        Where (Transaction_Details_1.Item_ID = " & Item_ID & ")"
    If LngUnitID <> 0 Then
     '   sql = sql & "        and (Transaction_Details_1.UnitID = " & LngUnitID & ")"
    End If

    sql = sql & "  AND (Transactions_1.Transaction_Type = 20)"
    sql = sql & "  and  Transactions_1.Transaction_Date <=" & SQLDate(Transaction_Date, True)
    sql = sql & " )"
    sql = sql & "  AND (dbo.Transactions.CBoBasedON = 5)"

    sql = sql & "  ORDER BY dbo.Transactions.Transaction_Date DESC, dbo.Transactions.NoteSerial1 DESC"

    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs2.RecordCount > 0 Then
        UnitFactor = val(rs2("UnitFactor").value & "")
        SecOrder = val(rs2("SecOrder").value & "")

        getcostbuylastinvoice = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value) '* QtyBySmalltUnit
    Else
        sql = "  SELECT          ISNULL(dbo.Transaction_Details.Price, 0) AS Price ,TblItemsUnits.UnitFactor,TblItemsUnits.SecOrder   "
        sql = sql & "  FROM            dbo.Transactions INNER JOIN                           dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
        sql = sql & "  Inner join TblItemsUnits On TblItemsUnits.ItemID =Transaction_Details.Item_ID and Transaction_Details.UnitId = Transaction_Details.UnitID "
        sql = sql & "   WHERE        (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND"
        If LngUnitID <> 0 Then
        '    sql = sql & "         (Transaction_Details.UnitID = " & LngUnitID & ") and "
        End If

        sql = sql & "   dbo.Transactions.Transaction_Date IN"
        sql = sql & "                             ("
        sql = sql & "  SELECT        MAX(Transactions_1.Transaction_Date) AS date"
        sql = sql & "             FROM            dbo.Transactions AS Transactions_1 INNER JOIN"
        sql = sql & " dbo.Transaction_Details AS Transaction_Details_1"
        sql = sql & "  ON Transactions_1.Transaction_ID = Transaction_Details_1.Transaction_ID"
        sql = sql & "        Where (Transaction_Details_1.Item_ID = " & Item_ID & ")"

        If LngUnitID <> 0 Then
         '   sql = sql & "        and (Transaction_Details_1.UnitID = " & LngUnitID & ")"
        End If
        sql = sql & "  AND (Transactions_1.Transaction_Type = 28)"
        sql = sql & "  and  Transactions_1.Transaction_Date <=" & SQLDate(Transaction_Date, True)
        sql = sql & " )"

        sql = sql & "  ORDER BY dbo.Transactions.Transaction_Date DESC, dbo.Transactions.NoteSerial1 DESC"
        rs2.Close
        rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If rs2.RecordCount > 0 Then
            UnitFactor = val(rs2("UnitFactor").value & "")
            SecOrder = val(rs2("SecOrder").value & "")

            getcostbuylastinvoice = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value) '* QtyBySmalltUnit
        Else
            getcostbuylastinvoice = 0
        End If

    End If
'    If getcostbuylastinvoice = 0 Then
'        Set rs2 = New ADODB.Recordset
'
'        sql = "  SELECT          ISNULL(dbo.Transaction_Details.Price, 0) AS Price   "
'        sql = sql & "  FROM            dbo.Transactions INNER JOIN                           dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'        sql = sql & "   WHERE        (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND"
'        If LngUnitID <> 0 Then
'            sql = sql & "         (Transaction_Details.UnitID = " & LngUnitID & ") and "
'        End If
'        sql = sql & "  (Transactions.Transaction_Type = 20) and "
'        sql = sql & "  (dbo.Transactions.CBoBasedON = 5)"
'        sql = sql & "  and ISNULL(dbo.Transaction_Details.Price, 0) <> 0 and "
'        sql = sql & "   dbo.Transactions.Transaction_Date IN"
'        sql = sql & "                             ("
'        sql = sql & "  SELECT        MAX(Transactions_1.Transaction_Date) AS date"
'        sql = sql & "             FROM            dbo.Transactions AS Transactions_1 INNER JOIN"
'        sql = sql & " dbo.Transaction_Details AS Transaction_Details_1"
'        sql = sql & "  ON Transactions_1.Transaction_ID = Transaction_Details_1.Transaction_ID"
'        sql = sql & "        Where (Transaction_Details_1.Item_ID = " & Item_ID & ")"
'        If LngUnitID <> 0 Then
'        '    sql = sql & "        and (Transaction_Details_1.UnitID = " & LngUnitID & ")"
'        End If
'
'        sql = sql & "  AND (Transactions_1.Transaction_Type = 22)"
'        sql = sql & "  and  Transactions_1.Transaction_Date <=" & SQLDate(Transaction_Date, True)
'        sql = sql & " )"
'
'        sql = sql & "  ORDER BY dbo.Transactions.Transaction_Date DESC, dbo.Transactions.NoteSerial1 DESC"
'
'        rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If rs2.RecordCount > 0 Then
'            getcostbuylastinvoice = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value) * QtyBySmalltUnit
'        Else
'            getcostbuylastinvoice = 0
'        End If
'
'    End If
End Function

Public Function ChkDateFormat() As Boolean
 If SystemOptions.CheckDateFormatCorrect = False Then
 ChkDateFormat = True
 Exit Function
 End If

    lLocal = GetSystemDefaultLCID()
    Length = GetLocaleInfo(3073, LOCALE_SSHORTDATE, buf, Len(buf))
   ChkDateFormat = True

    a = left$(buf, Length - 1)
    If SetLocaleInfo(3073, LOCALE_SSHORTDATE, "dd/MM/yyyy") = False Then

                If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "dd/mm/yyyy  ÌÃ» Ÿ»ÿ  ‰”Ìﬁ «· «—ÌŒ"
                Else
                MsgBox "  Date Formate Must Changed To : dd/mm/yyyy       "
                End If
        ChkDateFormat = False
        Exit Function
    End If

    length2 = GetLocaleInfo(3073, 32, buf2, Len(buf2))
    a = left$(buf2, length2 - 1)
    If SetLocaleInfo(3073, 32, "dd MMMM, yyyy") = False Then

        'MsgBox "ÌÃ» Ÿ»ÿ  ‰”Ìﬁ «· «—ÌŒ"
                        If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "dd/mm/yyyy  ÌÃ» Ÿ»ÿ  ‰”Ìﬁ «· «—ÌŒ"
                Else
                MsgBox "  Date Formate Must Changed To : dd/mm/yyyy       "
                End If

ChkDateFormat = False
        Exit Function

    End If


End Function



Function checkonlinedate() As Boolean
On Error Resume Next
Dim FileName As String
FileName = App.path & "\OnLineServer.txt"
If Dir(FileName, vbNormal) = "" Then checkonlinedate = False: Exit Function

            Open FileName For Input As #1

            Do Until EOF(1)
            Line Input #1, a


        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then



                   onlineservername = (VarSet(0))
                   onlineDataBasename = (VarSet(1))
    onlinusername = (VarSet(2))
     onlinepassword = (VarSet(3))
      onlinebackground = (VarSet(4))

                onLineMOde = True


checkonlinedate = True
Exit Function
            End If
        End If

    Loop

    Close #1
    checkonlinedate = False
End Function


Function print_report3_HyperLink(Optional innerStrAccountCode As String, Optional innerStrAccountnAME As String)
On Error Resume Next
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 Dim AccountTypes As Integer

  Dim OpeningBalancebeformdateMinus1 As Double
  Dim OpeningBalancebeformStartCurrentyearTOFromDAteminus1 As Double
  Dim NewOpinning As Double
  Dim OpeningBalance As Double
  Dim ProfitBalance As Double
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset

  Dim i As Integer
  Dim BranchID As String
  Dim HideZeroBalance As Integer
   Dim openingBalanceDate As Date
   Dim FromdateMinus1 As Date
   Dim StartCurrentDate As Date
   Dim BrcnActivety As String
   FromdateMinus1 = DateAdd("d", -1, P_DTPickerAccFrom)
    getFirstPeriodDateInthisYear2 openingBalanceDate
    getFirstPeriodDateInthisYear StartCurrentDate

         If SystemOptions.UserInterface = ArabicInterface Then
                HideZeroBalance = MsgBox("Â·  —Ìœ «Œ›«¡ Õ”«»«  ’›—ÌÂ ‰⁄„ «„ ·« ", vbInformation + vbYesNoCancel)
            Else
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            End If

            If HideZeroBalance = 2 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
          Dim BranshesReg As String

         If val(P_DCRegionID) <> 0 Then
         BranshesReg = BranchRegion(CDbl(P_DCRegionID))
         End If
         If val(P_DCActivity) <> 0 Then
         BrcnActivety = BrcnhActivityType(CDbl(P_DCActivity))
         End If


  updateprofitAccount val(P_DCActivity), val(P_dcBranch), P_DTPickerAccTo, BranshesReg

  sql = " SELECT    ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name, Account_NameEng , debitBalance ="
  sql = sql & "                         (SELECT     SUM(DEV_Value1)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d"
  sql = sql & "                                              WHERE      (d.Credit_Or_Debit = 0 AND d.RecordDate >= " & SQLDate(P_DTPickerAccFrom, True) & " AND d.RecordDate <= " & SQLDate(P_DTPickerAccTo, True) & ") AND d.Account_Code = A.Account_Code  and(d.Posted IS NULL)"
 If val(P_DCActivity) <> 0 Then
  sql = sql & " and d.branch_id in (" & BrcnActivety & ")"
  End If

  If val(P_DCRegionID) <> 0 Then
  sql = sql & " and d.branch_id in (" & BranshesReg & ")"
  End If

  If val(P_dcBranch) <> 0 Then
  sql = sql & " and d.branch_id =" & val(P_dcBranch) & ""
  End If
 sql = sql & "  ) x),"
  sql = sql & "                    CreditBalance ="
  sql = sql & "                        (SELECT     SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d1"
  sql = sql & "                                                   WHERE     (d1.Credit_Or_Debit = 1 AND d1.RecordDate >= " & SQLDate(P_DTPickerAccFrom, True) & "  AND d1.RecordDate <= " & SQLDate(P_DTPickerAccTo, True) & ") AND d1.Account_Code = A.Account_Code and(d1.Posted IS NULL)"
  If val(P_DCActivity) <> 0 Then
  sql = sql & " and d1.branch_id in (" & BrcnActivety & ")"
  End If
  If val(P_DCRegionID) <> 0 Then
  sql = sql & " and d1.branch_id in (" & BranshesReg & ")"
  End If
 If val(P_dcBranch) <> 0 Then
  sql = sql & " and d1.branch_id =" & val(P_dcBranch) & ""
  End If
  sql = sql & " ) x),"
  sql = sql & "                     OpeningBalance ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 AS do"
  sql = sql & "                                                   WHERE     (  do.Account_Code = A.Account_Code and(do.Posted IS NULL)"
  If val(P_DCActivity) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(P_DCRegionID) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(P_dcBranch) <> 0 Then
 sql = sql & " and do.branch_id =" & val(P_dcBranch) & ""
 End If
sql = sql & "  )) x),"
  sql = sql & "    OpeningBalancebeformdateMinus1 ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  sql = sql & "                                                   WHERE     ( do.RecordDate >=" & SQLDate(openingBalanceDate, True) & " and   do.RecordDate <= " & SQLDate(FromdateMinus1, True) & ") AND do.Account_Code = A.Account_Code and(do.Posted IS NULL)"
  If val(P_DCActivity) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(P_DCRegionID) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(P_dcBranch) <> 0 Then
  sql = sql & " and do.branch_id =" & val(P_dcBranch) & ""
  End If
  sql = sql & " ) x),"
  sql = sql & "                    OpeningBalancebeformStartCurrentyearTOFromDAteminus1 ="
  sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  sql = sql & "                                                   WHERE     (do.RecordDate >= " & SQLDate(StartCurrentDate, True) & " AND do.RecordDate < " & SQLDate(P_DTPickerAccFrom, True) & ") AND do.Account_Code = A.Account_Code and(do.Posted IS NULL) "
  If val(P_DCActivity) <> 0 Then
  sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
  End If
  If val(P_DCRegionID) <> 0 Then
  sql = sql & " and do.branch_id in (" & BranshesReg & ")"
  End If
  If val(P_dcBranch) <> 0 Then
  sql = sql & " and do.branch_id =" & val(P_dcBranch) & ""
  End If
  sql = sql & " ) x)"
  sql = sql & " FROM         ACCOUNTS A"
  sql = sql & " WHERE     A.last_account = 1   "



  sql = sql & " and (A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS"
  sql = sql & "    Where 1 = 1"
  StrAccountCode = (innerStrAccountCode)
        If mId(StrAccountCode, Len(StrAccountCode), 1) = "G" Then
                    StrAccountCode = mId(StrAccountCode, 1, Len(StrAccountCode) - 1)

                    End If

    If StrAccountCode <> "" Then


 sql = sql & " and A.Account_Code like'" & StrAccountCode & "a%'"
  End If

    If val(P_DCActivity) <> 0 Then
  sql = sql & " and branch_id in (" & BrcnActivety & ")"
  End If
  If val(P_DCRegionID) <> 0 Then
  sql = sql & " and branch_id in (" & BranshesReg & ")"
  End If
  If val(P_dcBranch) <> 0 Then
  sql = sql & " and branch_id =" & val(P_dcBranch) & ""
  End If
   sql = sql & "   )"
  sql = sql & " or A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS1"
    sql = sql & "    Where 1 = 1"


      If StrAccountCode <> "" Then
 sql = sql & " and A.Account_Code like'" & StrAccountCode & "a%'"
  End If


    If val(P_DCActivity) <> 0 Then
  sql = sql & " and branch_id in (" & BrcnActivety & ")"
  End If
  If val(P_DCRegionID) <> 0 Then
  sql = sql & " and branch_id in (" & BranshesReg & ")"
  End If
  If val(P_dcBranch) <> 0 Then
  sql = sql & " and branch_id =" & val(P_dcBranch) & ""
  End If
   sql = sql & "   ))"


    sql = sql & "order by Account_Serial "
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewSahyper.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewSaEhyper.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
     If SystemOptions.UserInterface = ArabicInterface Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
     Else
     Msg = "No Data"
     End If
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
   Dim desc As String
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
    desc = ""
    If val(P_DCActivity) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "«·‰‘«ÿ " & ": " & P_DCActivity & CHR(13)
   Else
   desc = desc & "Region" & ": " & P_DCActivity & CHR(13)
   End If
   End If

   If val(P_DCRegionID) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "··„‰ÿﬁ…" & ": " & P_DCRegionID & CHR(13)
   Else
   desc = desc & "Activity" & ": " & P_DCRegionID & CHR(13)
   End If
   End If
  If val(P_dcBranch) <> 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   desc = desc & "··›—⁄" & ": " & P_dcBranch & CHR(13)
   Else
   desc = desc & "Branch" & ": " & P_dcBranch & CHR(13)
   End If
   End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If HideZeroBalance = 6 Then
    xReport.ParameterFields(6).AddCurrentValue 1
    Else
    xReport.ParameterFields(6).AddCurrentValue 0
    End If
    If Not IsNull(P_DTPickerAccFrom) Then
    'xReport.ParameterFields(4).AddCurrentValue "" + CStr(P_DTPickerAccFrom)
    End If
    If Not IsNull(DTPickerAccTo) Then
   ' xReport.ParameterFields(5).AddCurrentValue ToDate(P_DTPickerAccTo)
    End If
  '  xReport.ParameterFields(7).AddCurrentValue desc
    xReport.reporttitle = "  Õ·Ì· «·Õ”«» " & CHR(13) & innerStrAccountnAME & CHR(13) & "   «·› —… „‰   " & P_DTPickerAccFrom & CHR(13) & "   «·Ì " & P_DTPickerAccTo & CHR(13) & desc
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:
  End Function




Public Function updateAccountsmanully(Ayear As Double)
    Dim StrSQL As String
    Dim Account_code  As String





          StrSQL = " update ACCOUNTS   "


            StrSQL = StrSQL & " SET opening_balance= "

StrSQL = StrSQL & " ("
StrSQL = StrSQL & " SELECT       dbo.TblBalanceSheetDetails.AValue"
StrSQL = StrSQL & " FROM            dbo.TblBalanceSheetHeader INNER JOIN"
StrSQL = StrSQL & "                                                   dbo.TblBalanceSheetDetails ON dbo.TblBalanceSheetHeader.BalanceSheetHeaderid = dbo.TblBalanceSheetDetails.BalanceSheetHeaderid"
StrSQL = StrSQL & " Where ACCOUNTS.Account_Code = TblBalanceSheetDetails.Account_Code and  (dbo.TblBalanceSheetHeader.DYear = " & Ayear & ")"
StrSQL = StrSQL & "  )"
StrSQL = StrSQL & "   where  ACCOUNTS.Account_Code in ("
StrSQL = StrSQL & "   SELECT       dbo.TblBalanceSheetDetails.Account_Code"
 StrSQL = StrSQL & "   FROM            dbo.TblBalanceSheetHeader INNER JOIN                                                   dbo.TblBalanceSheetDetails"
 StrSQL = StrSQL & "   ON dbo.TblBalanceSheetHeader.BalanceSheetHeaderid = dbo.TblBalanceSheetDetails.BalanceSheetHeaderid"
' StrSQL = StrSQL & "  Where (dbo.TblBalanceSheetHeader.DYear = " & Ayear & "  and Avalue<>0 )"
StrSQL = StrSQL & "  Where (dbo.TblBalanceSheetHeader.DYear = " & Ayear & "     )"

StrSQL = StrSQL & "  )"



Cn.Execute StrSQL




End Function




 Public Function GetCarsREbenueAcountCode(Optional ID As Double) As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim ownerid As Long
Dim sql As String
sql = "select AccountPaym from TblCarsData where id=" & ID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetCarsREbenueAcountCode = IIf(IsNull(rs2("AccountPaym").value), "", rs2("AccountPaym").value)
'ownerid = IIf(IsNull(rs2("ownerid").value), "", rs2("ownerid").value)
    '    If GetAqarAcountCode = "" Then
    '           GetAqarAcountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", ownerid)
    '    End If
Else


GetCarsREbenueAcountCode = ""
End If


End Function
  Public Function GetCarsREbenueAcountCode2(Optional ID As Double) As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim ownerid As Long
Dim sql As String
sql = "select DCOwner from TblCarsData where id=" & ID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
'rs2.Close

GetCarsREbenueAcountCode2 = IIf(IsNull(rs2("DCOwner").value), "", rs2("DCOwner").value)
'ownerid = IIf(IsNull(rs2("ownerid").value), "", rs2("ownerid").value)
    '    If GetAqarAcountCode = "" Then
    '           GetAqarAcountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", ownerid)
    '    End If
Else


GetCarsREbenueAcountCode2 = ""
End If


End Function

Public Function checkManulanoisExist(Optional Transaction_Type As Double, Optional Transaction_ID As Double, Optional CusID As Double, Optional ManualNO As String, Optional ByRef NoteSerial1 As String) As Boolean

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim ownerid As Long
Dim sql As String
sql = "SELECT     Transaction_ID, Transaction_Type, CusID, ManualNO, NoteSerial1"
sql = sql & "  From dbo.transactions"
sql = sql & "  WHERE     (Transaction_Type = " & Transaction_Type & ") "
sql = sql & "   AND (Transaction_ID <> " & Transaction_ID & ")"
sql = sql & "   AND (CusID = " & CusID & ")"
sql = sql & "   AND (ManualNO = '" & ManualNO & "')"

rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
checkManulanoisExist = True
 NoteSerial1 = IIf(IsNull(rs2("NoteSerial1").value), "", rs2("NoteSerial1").value)
Else

NoteSerial1 = ""
checkManulanoisExist = False
End If


End Function


 Public Function GetCarsFixedAssetID(Optional ID As Double) As Double
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim ownerid As Long
Dim sql As String
sql = " SELECT     fixedAssetid From dbo.TblCarsData  Where ID =" & ID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetCarsFixedAssetID = IIf(IsNull(rs2("fixedAssetid").value), 0, rs2("fixedAssetid").value)

Else


GetCarsFixedAssetID = 0
End If


End Function



 Public Function GetAqarAcountCode(Optional ID As Double) As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim ownerid As Long
Dim sql As String
sql = "select * from TblAqar where Aqarid=" & ID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetAqarAcountCode = IIf(IsNull(rs2("AccounCode").value), "", rs2("AccounCode").value)
'ownerid = IIf(IsNull(rs2("ownerid").value), "", rs2("ownerid").value)
    '    If GetAqarAcountCode = "" Then
    '           GetAqarAcountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", ownerid)
    '    End If
Else


GetAqarAcountCode = ""
End If


End Function


Public Function GetCurrencyCode(Optional ID As Double, Optional filed As String) As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = " SELECT        " & filed & " AS RetuensF"
sql = sql & " From dbo.currency"
sql = sql & " WHERE        (id = " & ID & ") "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetCurrencyCode = IIf(IsNull(rs2("RetuensF").value), 0, rs2("RetuensF").value)
Else
GetCurrencyCode = 0
End If
End Function

Public Function GetValueFiter(Optional ID As Double, Optional filed As String) As Double
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = " SELECT        SUM(" & filed & ") AS Value"
sql = sql & " From dbo.TblFiterWaiverDet2"
sql = sql & " WHERE        (MasterID = " & ID & ") "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetValueFiter = IIf(IsNull(rs2("Value").value), 0, rs2("Value").value)
Else
GetValueFiter = 0
End If
End Function
Public Function GetValueFiterHeader(Optional ID As Double, Optional filed As String) As Double
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = " SELECT        SUM(" & filed & ") AS Value"
sql = sql & " From dbo.TblFiterWaiver"
sql = sql & " WHERE        (ID = " & ID & ") "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetValueFiterHeader = IIf(IsNull(rs2("Value").value), 0, rs2("Value").value)
Else
GetValueFiterHeader = 0
End If
End Function
Public Function CheckAkarPayments(NoteID As Double) As Boolean
       Dim s As String
       Dim rs As New ADODB.Recordset

    s = "select * From Notes where NoteType=5   "

s = s & "   and not (  (akarid is null )  and   (IqarID2 is null )  and   (NoteOrBonID is null ) )  "

      s = s & " and NoteID= " & NoteID

         rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then

        CheckAkarPayments = True
        Else
        CheckAkarPayments = False
          End If

    rs.Close
End Function



Public Function CheckAkarCashes(NoteID As Double) As Boolean
       Dim s As String
       Dim rs As New ADODB.Recordset

    s = "select * From Notes where NoteType=4    "

s = s & " and CashingType >= 7"

        s = s & " and NoteID= " & NoteID

         rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then

        CheckAkarCashes = True
        Else
        CheckAkarCashes = False
          End If

    rs.Close
End Function


Public Function CheckUserNotPermAccounts(UserID As Double, AccountCode As String) As Boolean
       Dim s As String
       Dim rs As New ADODB.Recordset

    s = "select * From tblUserPermAccounts where UserId=" & UserID & "  and AccountCode='" & AccountCode & "'"

         rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
         If rs.RecordCount > 0 Then

        CheckUserNotPermAccounts = True
        Else
        CheckUserNotPermAccounts = False
          End If

    rs.Close
End Function

Public Function CheckAkarExpenses(NoteID As Double) As Boolean
       Dim s As String
       Dim rs As New ADODB.Recordset

         s = "select * From notes_all where notetype=3"
         s = s & " and  not (ToPriodDateH is null)"
        s = s & " and NoteID= " & NoteID

         rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then

        CheckAkarExpenses = True
        Else
        CheckAkarExpenses = False
          End If

    rs.Close
End Function



Public Function GetDepAccByEmp(ByVal mEmp As Integer, Optional ByVal mAccountIndex As Integer = 1) As String
       Dim s As String
       Dim rs As New ADODB.Recordset

         s = "SELECT TblEmpDepartments.Account_Code" & mAccountIndex & " as AccountName"

         s = s & " From TblEmpDepartments"
         s = s & " LEFT OUTER JOIN TblEmployee AS te"
         s = s & " ON  te.DepartmentID = TblEmpDepartments.DeparmentID"
         s = s & " Where te.Emp_ID = " & val(mEmp)

         rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then
             'GetDepAccByEmp = (rs!AccountName & "")

          If IsNull(rs!AccountName & "") Or Trim(rs!AccountName & "") = "" Then GetDepAccByEmp = "NO account": Exit Function

          If Not IsNull(rs!AccountName) Then
              If CheckAccountToJE(rs!AccountName & "") = True Then
                  GetDepAccByEmp = Trim(rs!AccountName) & "": Exit Function
              Else
                  GetDepAccByEmp = "NO account": Exit Function
              End If

          End If
    End If
    rs.Close
End Function



Public Function CheckWORKINposvATsCREEN() As Boolean
If SystemOptions.GeneralVoucherCreateSalesGE = True Then
CheckWORKINposvATsCREEN = True
Exit Function
End If

    Dim rs As ADODB.Recordset
    Dim StrSQL  As String

    StrSQL = "SELECT    * FROM TblReCalVATPO "


    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        CheckWORKINposvATsCREEN = True
    Else
        CheckWORKINposvATsCREEN = False
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function CheckintenalRequstQty(Item_ID As Double, order_no As String) As Double

    Dim rs As ADODB.Recordset
    Dim StrSQL  As String

    StrSQL = "SELECT     SUM(dbo.Transaction_Details.ShowQty) AS totalMoving"
StrSQL = StrSQL + "FROM         dbo.Transaction_Details INNER JOIN"
StrSQL = StrSQL + "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
StrSQL = StrSQL + "WHERE     (dbo.Transactions.Transaction_Type = 10) AND (dbo.Transactions.BillBasedOn = 1) AND (dbo.Transactions.order_no = '" & order_no & "' and Item_ID=" & Item_ID & ")"


    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckintenalRequstQty = IIf(IsNull(rs("totalMoving").value), 0, rs("totalMoving").value)
    Else
        CheckintenalRequstQty = 0
    End If

    rs.Close
    Set rs = Nothing
End Function



Public Function GetPaymentBank() As Long

    Dim rs As ADODB.Recordset
    Dim StrSQL  As String

    StrSQL = "Select *  From  TblPaymentType "
    StrSQL = StrSQL + " Where bankid<>0"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        GetPaymentBank = IIf(IsNull(rs("bankid").value), 0, rs("bankid").value)
    Else
        GetPaymentBank = 0
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function GetTotalSales(Optional Transaction_Type As Integer, Optional FromDate As Variant, Optional ToDate As Variant, Optional BranshesActiv As String, Optional BrnchIDes As String, Optional BranchID As Integer = 0) As Double
Dim StrSQL As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    StrSQL = " SELECT      sum(dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice) AS totalsale"
    StrSQL = StrSQL & "   FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    StrSQL = StrSQL & "  WHERE     (dbo.Transactions.Transaction_Type = " & Transaction_Type & ")"
 If BrnchIDes <> "-1" And BrnchIDes <> "" Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId in (" & BrnchIDes & ")"
  End If
  If BranshesActiv <> "-1" And BranshesActiv <> "" Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId in (" & BranshesActiv & ")"
  End If

   If BranchID <> 0 Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId= " & BranchID
  End If

    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
    End If

    If Not IsNull(ToDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
    End If
    rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs2.RecordCount > 0 Then
    GetTotalSales = IIf(IsNull(rs2("totalsale").value), 0, rs2("totalsale").value)
    Else
    GetTotalSales = 0
    End If
End Function
Public Function get_StoreBYPurchasePerson(PurchasePersonid As Double) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select StoreID from TblStore where PurchasePersonid=" & PurchasePersonid

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount = 0 Then get_StoreBYPurchasePerson = 0: Exit Function
    If IsNull(Rs3("StoreID").value) Then get_StoreBYPurchasePerson = 0: Exit Function
    If Not IsNull(Rs3("StoreID").value) Then get_StoreBYPurchasePerson = Rs3("StoreID").value: Exit Function
    Rs3.Close

End Function

Public Function GetTblProcessDEF(ProcessDEFID As Long, _
                            Optional ByRef ProcessName As String, _
                            Optional ByRef ProcessNameE As String, _
                            Optional ByRef UnitID As Integer)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
On Error Resume Next
    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from TblProcessDEF where TblProcessDEFID=" & ProcessDEFID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then

       ProcessName = IIf(IsNull(Rs1("ProcessName").value), "", Rs1("ProcessName").value)
       ProcessNameE = IIf(IsNull(Rs1("ProcessNameE").value), "", Rs1("ProcessNameE").value)
       UnitID = IIf(IsNull(Rs1("UnitID").value), "", Rs1("UnitID").value)

        If ProcessDEFID = 1 Then
            BranchID = 1
            StoreID = 1
            BoxID = 1
            BankID = 1
        End If

    Else

        If UserID = 1 Then
            BranchID = 1
            StoreID = 1
            BoxID = 1
            BankID = 1

        Else
            BranchID = 0
            StoreID = 0
            BoxID = 0
            BankID = 0
            EmpID = 0

        End If



    End If
'If checkmanyBranches("") = True Then usertype = 0
'If checkmanyStores("") = True Then usertype = 0
    Rs1.Close
End Function

Public Function PercentgValueAddedAllToBarcode(Optional RecDate As Date, Optional ItemID As Double, Optional Transe As Integer) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt "
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE     (TblSettsReqLimK.SelectType=0 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) )"
sql = sql + " AND (  dbo.TblSettsReqLimKDet.typ = 1)  "
sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
PercentgValueAddedAllToBarcode = IIf(IsNull(rs2("MultiPercentTxt").value), 0, rs2("MultiPercentTxt").value)
Else
PercentgValueAddedAllToBarcode = 0
End If
End Function
Public Function PercentgValueAddedBarcode(Optional RecDate As Date, Optional ItemID As Double, Optional Transe As Integer) As Double
Dim Percent As Double
Percent = 0
If CheckItemFreeVATToBarcode(RecDate, ItemID, Transe) = True Then
PercentgValueAddedBarcode = -1
Else
Percent = PercentgValueAddedAllToBarcode(RecDate, ItemID, Transe)
If Percent > 0 Then
PercentgValueAddedBarcode = Percent
Else
Percent = PercentgValueAddedGroupToBarcode(RecDate, ItemID, Transe)
If Percent > 0 Then
PercentgValueAddedBarcode = Percent
Else
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimKDet.PercentD"
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE   TblSettsReqLimK.SelectType=2 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) and   (dbo.TblSettsReqLimKDet.Typ = 0 AND (dbo.TblSettsReqLimKDet.ItemID = " & ItemID & ")) "

sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
PercentgValueAddedBarcode = IIf(IsNull(rs2("PercentD").value), 0, rs2("PercentD").value)
Else
PercentgValueAddedBarcode = 0
End If
End If
End If
End If
End Function

Public Function CheckItemFreeVATToBarcode(Optional RecDate As Date, Optional ItemID As Double, Optional Transe As Integer) As Boolean
Dim sql As String
If PercentgValueAddedGroupFreeToBarcode(RecDate, ItemID, Transe) = True Then
CheckItemFreeVATToBarcode = True
Else
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimKDet.PercentD"
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE     (dbo.TblSettsReqLimKDet.Typ = 5 AND (dbo.TblSettsReqLimKDet.ItemID = " & ItemID & ")) "
sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckItemFreeVATToBarcode = True
Else
CheckItemFreeVATToBarcode = False
End If
End If
End Function
Public Function PercentgValueAddedGroupFreeToBarcode(Optional RecDate As Date, Optional ItemID As Double, Optional Transe As Integer) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt"
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE     (TblSettsReqLimK.SelectType=1 and (TblSettsReqLimK.ItemStust=0 ) and  " & ItemID & " in ( SELECT     dbo.TblItems.ItemID"
sql = sql + " FROM         dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
sql = sql + "                      dbo.TblItems ON dbo.TblSettsReqLimKDet.GroupID = dbo.TblItems.GroupID"
sql = sql + " Where (dbo.TblSettsReqLimKDet.typ = 2)  And (dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID)))"
sql = sql + " AND ( dbo.TblSettsReqLimKDet.Typ = 1)  "
sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
PercentgValueAddedGroupFreeToBarcode = True
Else
PercentgValueAddedGroupFreeToBarcode = False
End If
End Function
 Public Function Get_movingreciveTransaction_ID(Optional ByRef ReturnID As Double) As Double
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = "SELECT     Transaction_ID"
sql = sql & " From dbo.transactions"
sql = sql & " Where (Transaction_Type = 11) And (ReturnID = " & ReturnID & ")"

rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
Get_movingreciveTransaction_ID = IIf(IsNull(rs2("Transaction_ID").value), 0, rs2("Transaction_ID").value)

Else
Get_movingreciveTransaction_ID = 0
 End If
End Function




Public Sub Get_TradingContractinfo(Optional ByRef TradingContractID As Double, Optional ByRef TContractCustID As Double, Optional Typed As Integer = 0)
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = "select ID,TContract_CustID  from Tbl_TradingContract"

sql = sql & " "
If Typed = 0 Then
sql = sql & " where id=" & TradingContractID & ""
Else
sql = sql & " where TContract_CustID='" & TContractCustID & "'"
End If
sql = sql & " And IsNull(IsCanceld,0) <> 1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
    TradingContractID = IIf(IsNull(rs2("id").value), 0, rs2("id").value)
    TContractCustID = IIf(IsNull(rs2("TContract_CustID").value), "", rs2("TContract_CustID").value)
Else
    TContractCustID = 0
    TradingContractID = 0
End If
End Sub

Public Function CheckCusCredit2(LngCusID As Long, _
                               SngOutValue As Single, _
                               IntCheckType As Integer, Optional Transaction_ID As Double, Optional ByRef MsgRe As String, Optional Typd As Integer = 0, Optional TransDate As Date, Optional IssueDate As Date) As Boolean

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim SngCreditLiimt As Single
    Dim SngCreditLimitCredit As Single
    Dim SngCusAccount As Single
    Dim Msg As String
    Dim StrTemp As String
    Dim IntRes As Integer
    Dim DepitInterval As Integer
    Dim DepitIntervalID As Integer
    Dim NoDay As Integer
    'On Local Error GoTo ErrTra

    StrSQL = "Select DepitIntervalID, DepitInterval,Account_Code,CreditLimit,CreditLimitCredit From TblCustemers Where CusID=" & LngCusID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        SngCreditLiimt = IIf(IsNull(rs("CreditLimit").value), 0, rs("CreditLimit").value)
        SngCreditLimitCredit = IIf(IsNull(rs("CreditLimitCredit").value), 0, rs("CreditLimitCredit").value)
           DepitInterval = IIf(IsNull(rs("DepitInterval").value), 0, rs("DepitInterval").value)
        DepitIntervalID = IIf(IsNull(rs("DepitIntervalID").value), 0, rs("DepitIntervalID").value)
    Else
        CheckCusCredit2 = False
        Exit Function
    End If

    If IntCheckType = 0 Then

        '«·ﬂ‘› ⁄·Ï «‰ „œÌÊ‰Ì… «·⁄„Ì· ·‰  “Ìœ ⁄‰ «·Õœ «·„Õœœ ·Â
        If SngCreditLiimt = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Ì—ÃÏ  ”ÃÌ· »Ì«‰«  Õœ «·«∆ „«‰ ·Â– «·⁄„Ì·"
        Else
        Msg = "Please enter data of credit limit"
        End If
            'NO CreditLimit For this customer
            MsgRe = Msg
            CheckCusCredit2 = False

            Exit Function
        Else

            '------------------------------------------------
            '»⁄œ «·√” ⁄·«„ ⁄‰ —’Ìœ «·⁄„Ì·
'*******new code********************************************
 Dim Account_code As String
     Dim FirstPeriod As Date
   getFirstPeriodDateInthisYear FirstPeriod

 Account_code = GetMyAccountCode("TblCustemers", "CusID", LngCusID)  '
 SngCusAccount = GetActualAccountBalance(Account_code, 0, FirstPeriod, Date)
 SngCusAccount = SngCusAccount - GetSumOfGeForOneAccount(Account_code, Transaction_ID, 0)
  If DepitIntervalID = 1 Then
DepitInterval = DepitInterval * 30
ElseIf DepitIntervalID = 2 Then
DepitInterval = DepitInterval * 365
End If
NoDay = DateDiff("d", IssueDate, TransDate)
NoDay = Abs(NoDay)
'***************************************************\

            If SngCusAccount >= 0 Then  '„œÌ‰
                If (Abs(SngCusAccount) + SngOutValue) > SngCreditLiimt Or (NoDay > DepitInterval) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
                     Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê›   ŒÿÏ  «·√∆ „«‰ «·Œ«’ »«·⁄„Ì·...!!!"
                 Else
                     Msg = "This process can not be allowed...!!!"
                     Msg = Msg & CHR(13) & "will exceed the credit limit...!!!"
                End If


                 If (Abs(SngCusAccount) + SngOutValue) > SngCreditLiimt Then
                   ' Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
                  '  Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê›   ŒÿÏ Õœ «·√∆ „«‰ «·Œ«’ »«·⁄„Ì·...!!!"
                    Msg = Msg & CHR(13) & "------------------------------------------------"
                    If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = Msg & CHR(13) & "Õœ ≈∆ „«‰ «·⁄„Ì· : " & SngCreditLiimt & " " & WriteNo(CStr(SngCreditLiimt), 0)
                    Msg = Msg & CHR(13) & "«·—’Ìœ «·Õ«·Ï ··⁄„Ì·  ﬁ»· Â–Â «·Õ—ﬂ…: "
                    Else
                    Msg = Msg & CHR(13) & "credit limit of customer : " & SngCreditLiimt & " " & WriteNo(CStr(SngCreditLiimt), 0)
                    Msg = Msg & CHR(13) & "current balance before this is process: "
                    End If
               If SystemOptions.UserInterface = ArabicInterface Then
                    If SngCusAccount > 0 Then
                        StrTemp = Abs(SngCusAccount) & "(„œÌ‰)"
                     ElseIf SngCusAccount < 0 Then
                     StrTemp = Abs(SngCusAccount) & "(œ«∆‰)"
                    Else
                        StrTemp = "(Œ«·’)"
                    End If
                Else
                     If SngCusAccount > 0 Then
                        StrTemp = Abs(SngCusAccount) & "(debt)"
                     ElseIf SngCusAccount < 0 Then
                     StrTemp = Abs(SngCusAccount) & "(credit)"
                    Else
                        StrTemp = "(Zero)"
                    End If
                End If


                    Msg = Msg & StrTemp
                  If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = Msg & CHR(13) & "«·„»·€ «·„—«œ  ”ÃÌ·Â ⁄·Ï «·⁄„Ì· : " & SngOutValue
                    Else
                    Msg = Msg & CHR(13) & "The amount to be recorded on the customer : " & SngOutValue
                  End If
                   ' Msg = Msg & Chr(13) & ""
                   End If
                 '//////////////////
                   If (NoDay > DepitInterval) Then
                   ' Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
                   If SystemOptions.UserInterface = EnglishInterface Then
                   ' Msg = Msg & Chr(13) & "------------------------------------------------"
                    Msg = Msg & CHR(13) & "credit period  : " & " " & DepitInterval
                   ' Msg = Msg & StrTemp
                    Msg = Msg & CHR(13) & "The period to be recorded on the customer  : " & NoDay
                 '   Msg = Msg & Chr(13) & ""
                 Else
                  '   Msg = Msg & Chr(13) & "------------------------------------------------"
                    Msg = Msg & CHR(13) & "„œ… ≈∆ „«‰ «·⁄„Ì· : " & " " & DepitInterval
                  '  Msg = Msg & StrTemp
                    Msg = Msg & CHR(13) & "«·„œ… «·„—«œ  ”ÃÌ·Â«  ··⁄„Ì· : " & NoDay

                   End If
                 End If
                 Msg = Msg & CHR(13) & ""
                 '/////////////

        If SystemOptions.SendToAprovedSalesBill = False And Typd = 1 Then
     Else
           MsgRe = Msg
           CheckCusCredit2 = False
           Exit Function
    End If

            End If

            '------------------------------------------------
        End If

        End If

    ElseIf IntCheckType = 1 Then

        '«·ﬂ‘› ⁄·Ï «‰ œ«∆‰Ì… «·⁄„Ì· ·‰  “Ìœ ⁄‰ «·Õœ «·„Õœœ ·Â
        If SngCreditLimitCredit = 0 Then
            'NO CreditLimit For this customer
            CheckCusCredit2 = True
            Exit Function
        Else
            'Set Rs = New ADODB.Recordset
            SngCusAccount = GetCustomerAccount(LngCusID, True)

            '------------------------------------------------
            '»⁄œ «·√” ⁄·«„ ⁄‰ —’Ìœ «·⁄„Ì·
            If SngCusAccount >= 0 Then '„œÌ‰
                If (Abs(SngCusAccount) + SngOutValue) > SngCreditLimitCredit Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê›   ŒÿÏ Õœ «·√∆ „«‰ («·œ«∆‰) «·Œ«’ »«·⁄„Ì· ...!!!"
                    Msg = Msg & CHR(13) & "------------------------------------------------"
                    Msg = Msg & CHR(13) & "Õœ ≈∆ „«‰ «·⁄„Ì· : " & SngCreditLimitCredit & " " & WriteNo(CStr(SngCreditLimitCredit), 0)
                    Msg = Msg & CHR(13) & "«·—’Ìœ «·Õ«·Ï ··⁄„Ì· : "

                    If SngCusAccount < 0 Then
                        StrTemp = Abs(SngCusAccount) & "(œ«∆‰)"
                    Else
                        StrTemp = "(Œ«·’)"
                    End If

                    Msg = Msg & StrTemp
                    Msg = Msg & CHR(13) & "«·„»·€ «·„—«œ  ”ÃÌ·Â ⁄·Ï «·⁄„Ì· : " & SngOutValue
               Else
                   ' Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
                   ' Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê›   ŒÿÏ Õœ «·√∆ „«‰ («·œ«∆‰) «·Œ«’ »«·⁄„Ì· ...!!!"
                     Msg = "This process can not be allowed...!!!"
                     Msg = Msg & CHR(13) & "will exceed the credit limit...!!!"

                       Msg = Msg & CHR(13) & "credit limit of customer : " & SngCreditLiimt & " " & WriteNo(CStr(SngCreditLiimt), 0)
                    Msg = Msg & CHR(13) & "current balance before this is process: "

                    Msg = Msg & CHR(13) & "------------------------------------------------"
                    Msg = Msg & CHR(13) & "credit limit of customer : " & SngCreditLimitCredit & " " & WriteNo(CStr(SngCreditLimitCredit), 0)
                    Msg = Msg & CHR(13) & "current balance : "

                    If SngCusAccount < 0 Then
                        StrTemp = Abs(SngCusAccount) & "(credit)"
                    Else
                        StrTemp = "(zero)"
                    End If

                    Msg = Msg & StrTemp
                    Msg = Msg & CHR(13) & "The amount to be recorded on the customer : " & SngOutValue

               End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    CheckCusCredit2 = False
                    Exit Function
                End If
            End If

            '------------------------------------------------
        End If
    End If

    CheckCusCredit2 = True
    Exit Function
ErrTrap:
    CheckCusCredit2 = False
End Function


Public Sub GetProjectInf(Optional ByRef projectId As Double, Optional ByRef ProjectCode As String, Optional Typed As Integer = 0)
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = "select id,Fullcode  from projects"
If Typed = 0 Then
sql = sql & " where id=" & projectId & ""
Else
sql = sql & " where Fullcode='" & ProjectCode & "'"
End If
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
projectId = IIf(IsNull(rs2("id").value), 0, rs2("id").value)
ProjectCode = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
Else
ProjectCode = ""
projectId = 0
End If
End Sub


Public Function BrcnhActivityType(Optional ActivityTypeId As Double) As String
Dim i As Integer
Dim BrnchIDes As String
BrnchIDes = "-1"
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = "  SELECT     branch_id"
sql = sql & " From dbo.TblBranchesData"
sql = sql & " Where (ActivityTypeId = " & ActivityTypeId & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
For i = 1 To rs2.RecordCount
BrnchIDes = BrnchIDes & "," & IIf(IsNull(rs2("branch_id").value), -1, rs2("branch_id").value)
rs2.MoveNext
Next i
End If
BrcnhActivityType = BrnchIDes
End Function


Public Function BranchRegion(Optional RegionID As Double) As String
Dim i As Integer
Dim BrnchIDes As String
BrnchIDes = "-1"
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = "  SELECT     branch_id"
sql = sql & " From dbo.TblBranchesData"
sql = sql & " Where (RegionID = " & RegionID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
For i = 1 To rs2.RecordCount
BrnchIDes = BrnchIDes & "," & IIf(IsNull(rs2("branch_id").value), -1, rs2("branch_id").value)
rs2.MoveNext
Next i
End If
BranchRegion = BrnchIDes
End Function

Public Function PercentgValueAddedAccounProject(Optional RecDate As Date, Optional ByRef flg As Integer, Optional BranchID As Double, Optional ByRef ForcedFlg As Integer) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimK.PercentH "
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE   (dbo.TblSettsReqLimK.ProjAccount=1 ) and (dbo.TblSettsReqLimKDet.Typ = 9) AND (dbo.TblSettsReqLimKDet.BranchID = " & BranchID & ")   "
sql = sql + " and      (" & SQLDate(RecDate, True) & "BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) and TblSettsReqLimK.AccOrTran=0 "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
flg = 1
PercentgValueAddedAccounProject = IIf(IsNull(rs2("PercentH").value), 0, rs2("PercentH").value)
ForcedFlg = 0
Else
PercentgValueAddedAccounProject = 0
flg = 0
ForcedFlg = 0
End If
End Function

 Public Function CheckProjectAccountDept(Optional ByRef Account_code As String) As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     expanses_account, Material_account, Salary_account, legal, AccountUnderImp, AcountGood"
sql = sql & " From dbo.Projects"
sql = sql & " WHERE    (AcountGood = N'" & Account_code & "') or  (expanses_account = N'" & Account_code & "') or (Salary_account = N'" & Account_code & "') or (legal = N'" & Account_code & "')  or (Material_account = N'" & Account_code & "')or (AccountUnderImp = N'" & Account_code & "')"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckProjectAccountDept = 0
Else
CheckProjectAccountDept = CheckProjectAccountCredit(Account_code)
End If
End Function
 Public Function CheckProjectAccountCredit(Optional ByRef Account_code As String) As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     REVENUE_account"
sql = sql & " From dbo.Projects"
sql = sql & " WHERE     (REVENUE_account = N'" & Account_code & "')"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckProjectAccountCredit = 1
Else
CheckProjectAccountCredit = -1
End If
End Function

 Public Function GetIssuedQty(order_no As String, Optional Transaction_ID As Double, Optional StoreId2 As Double, Optional Item_ID As Double, Optional OldID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "SELECT     SUM(dbo.Transaction_Details.ShowQty) AS ISSUEDQTY"
sql = sql & "  FROM         dbo.Transactions INNER JOIN"
sql = sql & "                        dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
sql = sql & "  WHERE     (dbo.Transactions.Transaction_Type = 19) AND (dbo.Transactions.order_no = '" & order_no & "') AND (dbo.Transactions.BillBasedOn = 2) AND"
sql = sql & "                        (dbo.Transactions.Transaction_ID <> " & Transaction_ID & ")"
sql = sql & "        and                 (dbo.Transactions.StoreID = " & StoreId2 & ")"
If Item_ID <> 0 Then
sql = sql & "        and                 (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
End If
If OldID <> 0 Then
 sql = sql & "        and                 (dbo.Transaction_Details.OldID = " & OldID & ")"
End If

rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetIssuedQty = IIf(IsNull(rs2("ISSUEDQTY").value), 0, rs2("ISSUEDQTY").value)
Else
GetIssuedQty = 0
End If

End Function

 Public Function GetIssuedQty2(order_no As String, Optional Transaction_ID As Double, Optional StoreId2 As Double, Optional Item_ID As Double, Optional OldID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "SELECT     SUM(dbo.Transaction_Details.ShowQty) AS ISSUEDQTY"
sql = sql & "  FROM         dbo.Transactions INNER JOIN"
sql = sql & "                        dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
sql = sql & "  WHERE     (dbo.Transactions.Transaction_Type = 19) AND (dbo.Transactions.order_no = '" & order_no & "') AND (dbo.Transactions.BillBasedOn = 13) AND"
sql = sql & "                        (dbo.Transactions.Transaction_ID <> " & Transaction_ID & ")"
sql = sql & "        and                 (dbo.Transactions.StoreID = " & StoreId2 & ")"
If Item_ID <> 0 Then
sql = sql & "        and                 (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
End If
If OldID <> 0 Then
 sql = sql & "        and                 (dbo.Transaction_Details.OldID = " & OldID & ")"
End If

rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetIssuedQty2 = IIf(IsNull(rs2("ISSUEDQTY").value), 0, rs2("ISSUEDQTY").value)
Else
GetIssuedQty2 = 0
End If

End Function
Public Function GetCusIDByCarID(Optional ID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from TblVendorCars where ID =" & ID & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetCusIDByCarID = IIf(IsNull(rs2("CustomerID").value), 0, rs2("CustomerID").value)
Else
GetCusIDByCarID = 0
End If
End Function

Public Sub GetAccountTypeTrans(Optional ID As Double, Optional ByRef AccountRevenue As String, Optional ByRef AccountExpense As String)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from TblTypesTransport where ID =" & ID & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
AccountRevenue = IIf(IsNull(rs2("AccountRevenue").value), "", rs2("AccountRevenue").value)
AccountExpense = IIf(IsNull(rs2("AccountExpense").value), "", rs2("AccountExpense").value)
Else
AccountExpense = ""
AccountRevenue = ""
End If
End Sub

  Public Sub DeletInvoiceofCustomer(Optional CusID As Double, Optional Transaction_Date As Date)
Cn.Execute "delete from Transactions where Transaction_Date=" & SQLDate(Transaction_Date, True) & " and CusID=" & CusID & " and Transaction_Type=21"
End Sub
Public Function CheckCustomerTrans(Optional CusID As Double, Optional Transaction_Date As Date) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     CusID2"
sql = sql & " FROM         dbo.Transaction_Details where Transaction_ID in (select Transaction_ID from Transactions "
sql = sql & " WHERE     (Transaction_Type=21  and Transaction_Date = " & SQLDate(Transaction_Date, True) & ") ) and CusID2=" & CusID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckCustomerTrans = True
Else
CheckCustomerTrans = False
End If
End Function
Public Function GetMaxIDTransection(Optional Item_ID As Long, Optional UnitID As Long) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     MAX(dbo.Transaction_Details.ID) AS MaxID"
sql = sql & " FROM         dbo.Transaction_Details INNER JOIN"
sql = sql & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
sql = sql & "                      dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
sql = sql & " Where (dbo.Transaction_Details.Item_ID = " & Item_ID & ") And (dbo.Transaction_Details.UnitID = " & UnitID & ") And (dbo.transactions.Transaction_Type = 21)"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxIDTransection = IIf(IsNull(rs2("MaxID").value), 0, rs2("MaxID").value)
Else
GetMaxIDTransection = 0
End If
End Function
Public Function GetLastPrice(Optional Item_ID As Long, Optional UnitID As Long) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "SELECT     showPrice"
sql = sql & " From dbo.Transaction_Details"
sql = sql & " WHERE     (ID = " & GetMaxIDTransection(Item_ID, UnitID) & ") "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetLastPrice = IIf(IsNull(rs2("showPrice").value), 0, Round(rs2("showPrice").value, 2))
Else
GetLastPrice = 0
End If
End Function
 Public Sub GetItemsInformation(Optional fullcode As String, Optional ByRef ItemID As Double, Optional ByRef Name As String)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     ItemID, ItemName, ItemNamee"
sql = sql & " From dbo.TblItems"
sql = sql & " WHERE     (Fullcode = N'" & fullcode & "') or (barCodeNO = N'" & fullcode & "')"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
ItemID = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
If SystemOptions.UserInterface = ArabicInterface Then
Name = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
Else
Name = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
End If
Else
Name = ""
ItemID = 0
End If
End Sub
Public Sub GetUnitID(Optional UnitName As String, Optional ByRef UnitID As Double)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     UnitID, UnitName, UnitNamee"
sql = sql & " FROM         dbo.TblUnites "
sql = sql & " WHERE     (UnitNamee = N'" & UnitName & "') or (UnitName = N'" & UnitName & "')"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
UnitID = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
Else
UnitID = 0
End If
End Sub

 Public Function GetRegVATNo(Optional branch_id As Integer) As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     VATNO"
sql = sql & " From dbo.TblBranchesData"
sql = sql & " Where (branch_id = " & branch_id & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetRegVATNo = IIf(IsNull(rs2("VATNO").value), "", rs2("VATNO").value)
Else
GetRegVATNo = ""
End If
End Function
Public Function checkmanyBoxes(Optional ByRef str As String = "") As Boolean

    Dim sql As String
    Dim rs As New ADODB.Recordset
If SystemOptions.UserInterface = ArabicInterface Then
  sql = " SELECT     dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName "
Else
sql = " SELECT     dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxNameE "
End If
sql = sql & " FROM         dbo.TblUsersBoxes  LEFT OUTER JOIN "
 sql = sql & "                     dbo.TblBoxesData ON dbo.TblUsersBoxes.BoxId = dbo.TblBoxesData.BoxID"
If user_id <> 1 Then
sql = sql & "    Where (dbo.TblUsersBoxes.userid = " & user_id & ")"
Else
   checkmanyBoxes = False
   Exit Function
  End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
str = sql
        checkmanyBoxes = True


    Else
   checkmanyBoxes = False
    End If

End Function
 Public Function CheckItemFreeVAT(Optional RecDate As Date, Optional StoreID As Long, Optional ItemID As Double, Optional Transe As Integer) As Boolean
Dim sql As String
If PercentgValueAddedGroupFree(RecDate, StoreID, ItemID, Transe) = True Then
CheckItemFreeVAT = True
Else
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimKDet.PercentD"
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE     (dbo.TblSettsReqLimKDet.Typ = 5 AND (dbo.TblSettsReqLimKDet.ItemID = " & ItemID & ")) "
sql = sql + " AND (dbo.TblSettsReqLimKDet.StoreID = " & StoreID & ")  "
sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckItemFreeVAT = True
Else
CheckItemFreeVAT = False
End If
End If
End Function

Public Function PercentgValueAddedAll(Optional RecDate As Date, Optional StoreID As Long, Optional ItemID As Double, Optional Transe As Integer) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt "
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE     (TblSettsReqLimK.SelectType=0 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) )"
sql = sql + " AND (dbo.TblSettsReqLimKDet.StoreID = " & StoreID & " and dbo.TblSettsReqLimKDet.typ = 1)  "
sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
PercentgValueAddedAll = IIf(IsNull(rs2("MultiPercentTxt").value), 0, rs2("MultiPercentTxt").value)
Else
PercentgValueAddedAll = 0
End If
End Function
Public Function PercentgValueAddedGroupFree(Optional RecDate As Date, Optional StoreID As Long, Optional ItemID As Double, Optional Transe As Integer) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt"
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE     (TblSettsReqLimK.SelectType=1 and (TblSettsReqLimK.ItemStust=0 ) and  " & ItemID & " in ( SELECT     dbo.TblItems.ItemID"
sql = sql + " FROM         dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
sql = sql + "                      dbo.TblItems ON dbo.TblSettsReqLimKDet.GroupID = dbo.TblItems.GroupID"
sql = sql + " Where (dbo.TblSettsReqLimKDet.typ = 2)  And (dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID)))"
sql = sql + " AND (dbo.TblSettsReqLimKDet.StoreID = " & StoreID & " and dbo.TblSettsReqLimKDet.Typ = 1)  "
sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
PercentgValueAddedGroupFree = True
Else
PercentgValueAddedGroupFree = False
End If
End Function

Public Function PercentgValueAddedGroup(Optional RecDate As Date, Optional StoreID As Long, Optional ItemID As Double, Optional Transe As Integer) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt"
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE     (TblSettsReqLimK.SelectType=1 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) and  " & ItemID & " in ( SELECT     dbo.TblItems.ItemID"
sql = sql + " FROM         dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
sql = sql + "                      dbo.TblItems ON dbo.TblSettsReqLimKDet.GroupID = dbo.TblItems.GroupID"
sql = sql + " Where (dbo.TblSettsReqLimKDet.typ = 2)  And (dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID)))"
sql = sql + " AND (dbo.TblSettsReqLimKDet.StoreID = " & StoreID & " and dbo.TblSettsReqLimKDet.Typ = 1)  "
sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
PercentgValueAddedGroup = IIf(IsNull(rs2("MultiPercentTxt").value), 0, rs2("MultiPercentTxt").value)
Else
PercentgValueAddedGroup = 0
End If
End Function

   Public Function PercentgValueAddedGroupToBarcode(Optional RecDate As Date, Optional ItemID As Double, Optional Transe As Integer) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt"
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE     (TblSettsReqLimK.SelectType=1 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) and  " & ItemID & " in ( SELECT     dbo.TblItems.ItemID"
sql = sql + " FROM         dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
sql = sql + "                      dbo.TblItems ON dbo.TblSettsReqLimKDet.GroupID = dbo.TblItems.GroupID"
sql = sql + " Where (dbo.TblSettsReqLimKDet.typ = 2)  And (dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID)))"
sql = sql + " AND ( dbo.TblSettsReqLimKDet.Typ = 1)  "
sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
PercentgValueAddedGroupToBarcode = IIf(IsNull(rs2("MultiPercentTxt").value), 0, rs2("MultiPercentTxt").value)
Else
PercentgValueAddedGroupToBarcode = 0
End If
End Function


  Public Sub GetAssestMoveYearly(Optional RecDate As Date, Optional ByRef YearMove As Double, Optional ByRef YearNotMove As Double, Optional ByRef MonthMove As Double, Optional ByRef MonthNotMove As Double)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " select * from TblSettsReqLimK "
sql = sql + " where       (" & SQLDate(RecDate, True) & "BETWEEN RecordDate AND RecordDateTo) "
sql = sql + " and  AccOrTran = 1 and TransType= 11"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
YearMove = IIf(IsNull(rs2("YearMove").value), 0, rs2("YearMove").value)
YearNotMove = IIf(IsNull(rs2("YearNotMove").value), 0, rs2("YearNotMove").value)
MonthMove = IIf(IsNull(rs2("MonthMove").value), 0, rs2("MonthMove").value)
MonthNotMove = IIf(IsNull(rs2("MonthNotMove").value), 0, rs2("MonthNotMove").value)
Else
MonthMove = 0
MonthNotMove = 0
YearMove = 0
YearNotMove = 0
End If
End Sub
 Public Function GetCashCustomerPhoneByName(CashCustomerName As String) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
        sql = "SELECT     CashCustomerName, CashCustomerPhone From dbo.Transactions  WHERE     (CashCustomerName = '" & CashCustomerName & "')"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
        GetCashCustomerPhoneByName = IIf(IsNull(rs("CashCustomerPhone").value), "", rs("CashCustomerPhone").value)
    Else
         GetCashCustomerPhoneByName = ""
    End If
    rs.Close
End Function
Public Function GetItemUnitsId(Optional UnitName As String) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     UnitID, UnitName, UnitNamee"
sql = sql & " FROM         dbo.TblUnites where UnitName='" & UnitName & " '"
sql = sql & " or    UnitNamee='" & UnitName & " '"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetItemUnitsId = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
Else
GetItemUnitsId = 0
End If
End Function

 Public Sub PercentgValueAddedAccount_Transec(Optional RecDate As Date, Optional TransType As Integer, Optional Dept_Credit As Integer, Optional ByRef AccountCode As String, Optional ByRef Percentage As Double)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " select * from TblSettsReqLimK "
sql = sql + " where       (" & SQLDate(RecDate, True) & "BETWEEN RecordDate AND RecordDateTo) "
sql = sql + " and  AccOrTran = 1 and TransType= " & TransType & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

If rs2.RecordCount > 0 Then
Percentage = IIf(IsNull(rs2("PercentH").value), 0, rs2("PercentH").value)
If Dept_Credit = 0 Then
AccountCode = IIf(IsNull(rs2("AccDep").value), "", rs2("AccDep").value)
Else
AccountCode = IIf(IsNull(rs2("AccCir").value), "", rs2("AccCir").value)
End If
Else
AccountCode = ""
Percentage = 0
End If
End Sub
Public Function GetItemUnits(Optional UnitID As Double) As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     UnitID, UnitName, UnitNamee"
sql = sql & " FROM         dbo.TblUnites where UnitID=" & UnitID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
GetItemUnits = IIf(IsNull(rs2("UnitName").value), "", rs2("UnitName").value)
Else
GetItemUnits = IIf(IsNull(rs2("UnitNamee").value), "", rs2("UnitNamee").value)
End If
Else
GetItemUnits = ""
End If
End Function
Public Function CheckCustomerCont(Optional customerid As Double, Optional RecDate As Date) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     TblCustomerContractD"
sql = sql & " From dbo.TblCustomerContract"
sql = sql & " WHERE     (Locked = 0 OR Locked IS NULL)"
sql = sql & " and CustomerId=" & customerid & ""
sql = sql & " and FromDate <=" & SQLDate(RecDate, True) & ""
sql = sql & " and Todate >=" & SQLDate(RecDate, True) & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckCustomerCont = IIf(IsNull(rs2("TblCustomerContractD").value), 0, rs2("TblCustomerContractD").value)
Else
CheckCustomerCont = False
End If
End Function
 Public Function CheckWorkState(UserID As Integer) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     dbo.TblUsers.UserID, dbo.TblEmployee.workstate"
sql = sql & " FROM         dbo.TblUsers LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
sql = sql & " Where (dbo.TblUsers.UserID = " & UserID & ") And (dbo.TblEmployee.WorkState = 1)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckWorkState = True
Else
CheckWorkState = False
End If
End Function

Public Function GetItemsTotalByStore(Optional Transaction_ID As Long, Optional StoreID As Integer) As Double
    Dim DblTemp As Double
    Dim RowNum As Long
    Dim Msg  As String
    Dim sql As String
    Dim linetotl As Double
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
     On Local Error GoTo ErrTrap
     sql = " SELECT     SUM(ShowQty * showPrice) AS Price"
     sql = sql & "      From dbo.Transaction_Details"
     sql = sql & "  WHERE     (StoreID2 = " & StoreID & ") AND (Transaction_ID = " & Transaction_ID & ")"
     rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
     If rs2.RecordCount > 0 Then
            linetotl = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
                     If SystemOptions.PursgaseWithoutDecimal = True And GridTrans = PurchaseTransaction Then
                         DblTemp = DblTemp + Int(linetotl)
                     Else
                           DblTemp = DblTemp + Round(linetotl, SystemOptions.SysDefCurrencyForamt)
                    End If
    Else
    linetotl = 0
    End If

    GetItemsTotalByStore = linetotl
    Exit Function
ErrTrap:
    Msg = "ERROR "
    GetItemsTotalByStore = linetotl
End Function

Public Function CheckExpeIqar(Optional NoteID As Double) As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim StrSQL As String
 StrSQL = "select * From notes_all where notetype=3"
 StrSQL = StrSQL & " and  not (ToPriodDateH is null) and NoteID=" & NoteID & ""
 rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If rs2.RecordCount > 0 Then
 CheckExpeIqar = 1
 Else
 CheckExpeIqar = 0
 End If
End Function
 Public Function GetIDesUnpadiVacation(Optional Emp_id As Double) As String
Dim sql As String
Dim StrIDes As String
Dim i As Integer
Dim NoDay As Double
NoDay = 0
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
StrIDes = "0,0"
sql = " SELECT     id, MoveVacBalance"
sql = sql & " From dbo.TblEmbarkation"
sql = sql & " WHERE     (TypeVacation = 1) AND (VacationPaied IS NULL OR  VacationPaied = 0)"
sql = sql & "  AND (Emp_ID = " & Emp_id & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
For i = 1 To rs2.RecordCount
StrIDes = StrIDes & "," & IIf(IsNull(rs2("id").value), 0, rs2("id").value)
rs2.MoveNext
Next i
End If
GetIDesUnpadiVacation = StrIDes
End Function
 Public Function GetNoDayUnpadiVacation2(Optional Emp_id As Double, Optional RdTypeVaction As Integer = 0) As Double
Dim sql As String
Dim i As Integer
Dim NoDay As Double
NoDay = 0
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

sql = " SELECT     id, MoveVacBalance"
sql = sql & " From dbo.TblEmbarkation"
sql = sql & " WHERE     (TypeVacation = 1) AND (VacationPaied IS NULL OR  VacationPaied = 0)"
sql = sql & "  AND (Emp_ID = " & Emp_id & ")"
If RdTypeVaction = 1 Then
sql = sql & "  AND (RdTypeVaction = 1)"
Else
sql = sql & "  AND (RdTypeVaction = 0 or RdTypeVaction is null)"
End If
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
For i = 1 To rs2.RecordCount
NoDay = NoDay + IIf(IsNull(rs2("MoveVacBalance").value), 0, rs2("MoveVacBalance").value)
rs2.MoveNext
Next i
End If
GetNoDayUnpadiVacation2 = NoDay
End Function

 Public Function GetMaxIDVation(Optional EmpID As Double) As Double
If EmpID = 0 Then GetMaxIDVation = 0: Exit Function
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     MAX(ID) AS MaxID"
sql = sql & " From dbo.TblVocationEntitlements"
sql = sql & " Where (EmpID = " & EmpID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxIDVation = IIf(IsNull(rs2("MaxID").value), 0, rs2("MaxID").value)
Else
GetMaxIDVation = 0
End If
End Function


 Public Function GetMaxIDVation2(Optional EmpID As Double) As Double
If EmpID = 0 Then GetMaxIDVation2 = 0: Exit Function
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     MAX(ID) AS MaxID"
sql = sql & " From dbo.TblInstalVacationDet"
sql = sql & " Where (EmpID = " & EmpID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxIDVation2 = IIf(IsNull(rs2("MaxID").value), 0, rs2("MaxID").value)
Else
GetMaxIDVation2 = 0
End If
End Function

Public Function GetLastBalanceMonthVaction(Optional EmpID As Double, Optional ID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     LastBalanceMonth"
sql = sql & " From dbo.TblVocationEntitlements"
sql = sql & " WHERE     (ID = " & GetMaxIDVation(EmpID) & ") and ID<>" & ID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetLastBalanceMonthVaction = IIf(IsNull(rs2("LastBalanceMonth").value), 0, rs2("LastBalanceMonth").value)
Else
GetLastBalanceMonthVaction = 0
End If
rs2.Close
sql = " Select VacBalance,EmpID from TblInstalVacationDet"
sql = sql & " WHERE     (ID = " & GetMaxIDVation2(EmpID) & ") and ID<>" & ID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetLastBalanceMonthVaction = IIf(IsNull(rs2("VacBalance").value), 0, rs2("VacBalance").value) + GetLastBalanceMonthVaction
Else
'GetLastBalanceMonthVaction = 0
End If

End Function

 Public Function GetEmIDUnpaidVacation(Optional ID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     Emp_id"
sql = sql & " From dbo.TblEmpPassOver"
sql = sql & " Where (TypeTrans = 3) And (advanceID = " & ID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText



If rs2.RecordCount > 0 Then
GetEmIDUnpaidVacation = IIf(IsNull(rs2("Emp_id").value), 0, rs2("Emp_id").value)
Else
GetEmIDUnpaidVacation = 0
End If


End Function
 Public Sub GetNoDayUnpadiVacation(Optional Emp_id As Double, Optional ByRef IDes As String, Optional ByRef NoVaction As Double)
Dim sql As String
Dim StrIDes As String
Dim i As Integer
Dim NoDay As Double
NoDay = 0
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
StrIDes = "0,0"
sql = " SELECT     id, MoveVacBalance"
sql = sql & " From dbo.TblEmbarkation"
sql = sql & " WHERE     (TypeVacation = 1) AND (VacationPaied IS NULL OR  VacationPaied = 0)"
sql = sql & "  AND (Emp_ID = " & Emp_id & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
For i = 1 To rs2.RecordCount
StrIDes = StrIDes & "," & IIf(IsNull(rs2("id").value), 0, rs2("id").value)
NoDay = NoDay + IIf(IsNull(rs2("MoveVacBalance").value), 0, rs2("MoveVacBalance").value)
rs2.MoveNext
Next i
End If
IDes = StrIDes
NoVaction = NoDay
End Sub
 Public Function GetTrnasectionID(Optional MainTransaction_ID As Double, Optional Transaction_Type As Integer)
Dim StrIDes As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "select * from TblTransctionIDES where MainTransaction_ID=" & MainTransaction_ID & "and Transaction_Type=" & Transaction_Type & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
StrIDes = "0,0"
If rs2.RecordCount > 0 Then
rs2.MoveFirst
For i = 1 To rs2.RecordCount
StrIDes = StrIDes & "," & IIf(IsNull(rs2("Transaction_ID").value), 0, rs2("Transaction_ID").value)
rs2.MoveNext
Next i
End If
GetTrnasectionID = StrIDes
End Function
Public Sub SaveTrnasectionID(Optional MainTransaction_ID As Double, Optional Transaction_ID As Long, Optional Transaction_Type As Integer)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " select * from TblTransctionIDES where 1=-1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
rs2.AddNew
rs2("MainTransaction_ID").value = MainTransaction_ID
rs2("Transaction_ID").value = Transaction_ID
rs2("Transaction_Type").value = Transaction_Type
rs2.update
End Sub
Public Function CheckSettingsLikeContract() As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select CommContract from TblVacationSettings where CommContract=1 "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckSettingsLikeContract = True
Else
CheckSettingsLikeContract = False
End If
End Function

 Public Function GetAccountCodeHiding() As String
Dim My_SQL As String
Dim FlgBign As Boolean
Dim Account_Code5 As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
  My_SQL = " SELECT    AccountCode"
  My_SQL = My_SQL & " From AccountSetting"
  My_SQL = My_SQL & " where     TreeAccount = 1"
rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
My_SQL = " ( SELECT     Account_Code"
My_SQL = My_SQL & " From dbo.Accounts"
FlgBign = False
For i = 1 To rs2.RecordCount
Account_Code5 = IIf(IsNull(rs2("AccountCode").value), "", rs2("AccountCode").value)
If Account_Code5 <> "" Then
If FlgBign = False Then
FlgBign = True
My_SQL = My_SQL & " where   (Account_Code LIKE N'" & Account_Code5 & "%') and Account_Code<>'" & Account_Code5 & "' "
Else
My_SQL = My_SQL & " or  (Account_Code LIKE N'" & Account_Code5 & "%')and Account_Code<>'" & Account_Code5 & "' "
End If
End If

rs2.MoveNext
Next i
Else
GetAccountCodeHiding = ""
Exit Function
End If
My_SQL = My_SQL & ")"
GetAccountCodeHiding = My_SQL

GetAccountCodeHiding = " and ACCOUNTS.Account_Code not in " & My_SQL
End Function

  Public Function GetValueAddedAccount(Optional RecDate As Date, Optional ByRef Account_CodeDept As String, Optional ByRef Account_CodeCridit As String, Optional Trans_Account As Integer = 0, Optional TransType As Integer) As Boolean
If mdifrmmain.taxes = False Then
GetValueAddedAccount = True
ElseIf CheckAnyVAT(RecDate) = False Then
GetValueAddedAccount = True
Else
Dim sql As String
GetValueAddedAccount = False
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     AccDep ,AccCir"
sql = sql & " FROM       TblSettsReqLimK"
sql = sql + "  WHERE     (" & SQLDate(RecDate, True) & "BETWEEN RecordDate AND RecordDateTo) "
If Trans_Account = 1 Then
sql = sql + " and      ((AccOrTran = 1) OR  (AccOrTran IS NULL)) and TransType=" & TransType & " "
Else
sql = sql + " and      (AccOrTran = 0)"
End If

rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
Account_CodeDept = IIf(IsNull(rs2("AccDep").value), "", rs2("AccDep").value)
Account_CodeCridit = IIf(IsNull(rs2("AccCir").value), "", rs2("AccCir").value)
GetValueAddedAccount = True
Else
GetValueAddedAccount = False
Account_CodeCridit = ""
Account_CodeDept = ""
End If
End If
End Function

 Public Function CheckAnyVAT(Optional RecDate As Date) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "select * from TblSettsReqLimK where 1=1 "
sql = sql + " and      (" & SQLDate(RecDate, True) & "BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) or dbo.TblSettsReqLimK.RecordDateTo<=" & SQLDate(RecDate, True) & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckAnyVAT = True
Else
CheckAnyVAT = False
End If
End Function

Public Function ScreenAproved(Optional Transaction_ID As Double, Optional ScreenName As String) As Boolean
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
If SystemOptions.CaNUpdateApprovedDoc = True Then ScreenAproved = False: Exit Function
 If CheckAprroveScreen(ScreenName) = True Then
 If Transaction_ID = 0 Then
 sql = "Select * from ApprovalData where ScreenName='" & ScreenName
 Else
 sql = "Select * from ApprovalData where ScreenName='" & ScreenName & "' and Transaction_ID =" & Transaction_ID & ""
End If
 sql = sql & " and (NOT (ApprovDate IS NULL) or NOT (CancelApprove IS NULL) )"
 rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If rs2.RecordCount > 0 Then
 ScreenAproved = True
 Else
 ScreenAproved = False
 End If
 Else
 ScreenAproved = False
 End If
End Function
 Public Function MainCurrency() As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select ID from currency where basic=1 "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
MainCurrency = IIf(IsNull(rs2("ID").value), 1, rs2("ID").value)
Else
MainCurrency = 1
End If
End Function

    Public Sub getinsttPayedToContNote(Optional NoteID As Double = 0, Optional ByRef RentValuePayed As Double, Optional ByRef CommissionsPayed As Double, Optional ByRef InsurancePayed As Double, Optional ByRef WaterPayed As Double, Optional ByRef ElectricPayed As Double, Optional ByRef TelandNetPayed As Double, Optional ByRef TotalOldValue As Double, Optional Istallid As Double, Optional ByRef VATPayed As Double)
    On Error Resume Next

    Dim total As Single

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

 sql = "select (value) As total   ,(RentValuePayed) as RentValuePayed,(CommissionsPayed) as CommissionsPayed"
 sql = sql & "  ,(InsurancePayed) as InsurancePayed,(WaterPayed) as WaterPayed"
 sql = sql & "  ,(ElectricPayed) as ElectricPayed,(TelandNetPayed) as TelandNetPayed ,(OldValuePayed) as TotalOldValue ,(VATPayed) as VATPayed"
 sql = sql & "  from ContracttBillInstallmentsDone  where NoteID=" & NoteID & "  and Istallid=" & Istallid & ""

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount = 0 Then
        total = 0
          RentValuePayed = 0
  CommissionsPayed = 0
  InsurancePayed = 0
  WaterPayed = 0
  ElectricPayed = 0
  TelandNetPayed = 0
  TotalOldValue = 0
  VATPayed = 0
    Else

        total = IIf(IsNull(Rs3("total").value), 0, Rs3("total").value)
        RentValuePayed = IIf(IsNull(Rs3("RentValuePayed").value), 0, Rs3("RentValuePayed").value)
        CommissionsPayed = IIf(IsNull(Rs3("CommissionsPayed").value), 0, Rs3("CommissionsPayed").value)
        InsurancePayed = IIf(IsNull(Rs3("InsurancePayed").value), 0, Rs3("InsurancePayed").value)
        WaterPayed = IIf(IsNull(Rs3("WaterPayed").value), 0, Rs3("WaterPayed").value)
        ElectricPayed = IIf(IsNull(Rs3("ElectricPayed").value), 0, Rs3("ElectricPayed").value)
        TelandNetPayed = IIf(IsNull(Rs3("TelandNetPayed").value), 0, Rs3("TelandNetPayed").value)
        TotalOldValue = IIf(IsNull(Rs3("TotalOldValue").value), 0, Rs3("TotalOldValue").value)
        VATPayed = IIf(IsNull(Rs3("VATPayed").value), 0, Rs3("VATPayed").value)
    End If

    Rs3.Close

End Sub
   Public Function GetMosim(Optional Omra_Hajj As Integer) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     ID "
sql = sql & " FROM         dbo.TblCompaniesGroup"
sql = sql & " where CurrYear=1 "
If Omra_Hajj = 0 Then
sql = sql & " and Omra_Hajj=0 "
Else
sql = sql & " and Omra_Hajj=1 "
End If
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMosim = IIf(IsNull(rs2("ID").value), 0, rs2("ID").value)
Else
GetMosim = 0
End If
End Function
Public Function GetIDOrder(Optional NoteSerial1 As Double, Optional SeasonsID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " select ID from tblbookingrequest where NoteSerial1=" & NoteSerial1 & " and SeasonsID=" & SeasonsID & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetIDOrder = IIf(IsNull(rs2("ID").value), 0, rs2("ID").value)
Else
GetIDOrder = 0
End If
End Function

Public Function GetCustomerVAT(CusID As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = " SELECT   * from TblCustemers  "

    sql = sql & " Where (CusID = " & CusID & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
                        GetCustomerVAT = IIf(IsNull(rs("VATNO").value), "", rs("VATNO").value)
    Else
    GetCustomerVAT = ""
    End If

End Function
  Public Function PercentgValueAddedAccount(Optional RecDate As Date, Optional Account_code As String, Optional BranchID As Double, Optional ByRef ForcedFlg As Integer) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimKDet.PercentD ,dbo.TblSettsReqLimKDet.ForcedFlg"
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE     (dbo.TblSettsReqLimKDet.Typ = 9) AND (dbo.TblSettsReqLimKDet.BranchID = " & BranchID & ") AND (dbo.TblSettsReqLimKDet.Account_Code = '" & Account_code & "')  "
sql = sql + " and      (" & SQLDate(RecDate, True) & "BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) and TblSettsReqLimK.AccOrTran=0 "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
PercentgValueAddedAccount = IIf(IsNull(rs2("PercentD").value), 0, rs2("PercentD").value)
    If Not IsNull(rs2("ForcedFlg").value) Then
        If (rs2("ForcedFlg").value) = True Then
            ForcedFlg = 1
        Else
            ForcedFlg = 0
        End If
    Else
        ForcedFlg = 0
    End If
Else
    PercentgValueAddedAccount = 0
    ForcedFlg = 0
End If
End Function

 Public Function PercentgValueAdded(Optional RecDate As Date, Optional StoreID As Long, Optional ItemID As Double, Optional Transe As Integer) As Double
Dim Percent As Double

If UCase(frmname) = "FRMSALEBILL" Then
    If chkTaxExempt.value = vbChecked Then
         VatGrid.rows = 1
         Percent = 0
         PercentgValueAdded = 0
         Exit Function
    End If
End If

Percent = 0
If CheckItemFreeVAT(RecDate, StoreID, ItemID, Transe) = True Then
PercentgValueAdded = -1
Else
Percent = PercentgValueAddedAll(RecDate, StoreID, ItemID, Transe)
If Percent > 0 Then
PercentgValueAdded = Percent
Else
Percent = PercentgValueAddedGroup(RecDate, StoreID, ItemID, Transe)
If Percent > 0 Then
PercentgValueAdded = Percent
Else
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblSettsReqLimKDet.PercentD"
sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
sql = sql & " WHERE   TblSettsReqLimK.SelectType=2 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) and   (dbo.TblSettsReqLimKDet.Typ = 0 AND (dbo.TblSettsReqLimKDet.ItemID = " & ItemID & ")) "
sql = sql + " AND (dbo.TblSettsReqLimKDet.StoreID = " & StoreID & ")  "
sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
PercentgValueAdded = IIf(IsNull(rs2("PercentD").value), 0, rs2("PercentD").value)
Else
PercentgValueAdded = 0
End If
End If
End If
End If
End Function
  Public Function GetServerdate(ServerDate As Date, ServerTime As Date)
    Dim StrTemp As String
    Dim RsTemp As ADODB.Recordset

        StrTemp = "select Getdate() as ServerDate , RIGHT(CONVERT(VARCHAR, GETDATE(), 100),7) as ServerTime"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrTemp, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not IsNull(RsTemp("ServerDate").value) Then
             ServerDate = Format(RsTemp("ServerDate").value, "yyyy/M/d")
             ServerTime = RsTemp("ServerTime").value

            End If


        End If

        RsTemp.Close
        Set RsTemp = Nothing


End Function
Public Sub GetID_CodeSqureProject(Optional ByRef SquareCode As String = "", Optional ByRef ID As Double = 0)
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     ID, SquareCode"
sql = sql & " From dbo.TblProjecInvestment"
If SquareCode <> "" Then
sql = sql & " WHERE     (SquareCode = N'" & SquareCode & "')"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
ID = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
Else
ID = 0
End If
End If
If ID <> 0 Then
sql = sql & " WHERE     (ID = " & ID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
SquareCode = IIf(IsNull(Rs3("SquareCode").value), "", Rs3("SquareCode").value)
Else
SquareCode = ""
End If
End If
End Sub
Public Function CheckPayment(Optional NoteID As Double) As Boolean
Dim StrSQL As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=4    "
    StrSQL = StrSQL & "and   (NOT (Status IS NULL)) and NoteID=" & NoteID & ""
    Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.RecordCount > 0 Then
    CheckPayment = True
    Else
    CheckPayment = False
    End If
End Function

Public Function CheckAprroveScreen(Optional ScreenName As String) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     ScreenName"
sql = sql & " From dbo.TblApprovalDef"
sql = sql & " WHERE     (ScreenName = N'" & ScreenName & "')"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckAprroveScreen = True
Else
CheckAprroveScreen = False
End If
End Function
Public Sub UpdateItemsDefaultUnit()
Dim sql As String
Dim ItemID As Double
Dim i As Integer
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select ItemID from TblItems "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
For i = 1 To Rs3.RecordCount
ItemID = IIf(IsNull(Rs3("ItemID").value), 0, Rs3("ItemID").value)
Cn.Execute "update TblItemsUnits  set DefaultUnit=1 where ItemID=" & ItemID & " and UnitFactor= " & GetMaxUnitFactor(ItemID) & ""
Rs3.MoveNext
Next i
End If
End Sub
Public Function GetMaxUnitFactor(Optional ItemID As Double) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     MAX(UnitFactor) AS MaxUnitFactor"
sql = sql & " From dbo.TblItemsUnits"
'Sql = Sql & " GROUP BY ItemID"
sql = sql & " where      (ItemID = " & ItemID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetMaxUnitFactor = IIf(IsNull(Rs3("MaxUnitFactor").value), 0, Rs3("MaxUnitFactor").value)
Else
GetMaxUnitFactor = 0
End If
End Function

Public Function Calcul30orRminder(Optional ID As Integer) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     id, culc30orRminder"
sql = sql & " From dbo.MOFRAD"
sql = sql & " Where (culc30orRminder = 1) and id=" & ID & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Calcul30orRminder = True
Else
Calcul30orRminder = False
End If
End Function
Public Function CheckSettingsVacType() As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select Typ from TblVacationSettings where Typ=1 "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckSettingsVacType = True
Else
CheckSettingsVacType = False
End If
End Function
Public Function GetSettingsVacPeriod() As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select NoMonth from TblVacationSettings "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetSettingsVacPeriod = IIf(IsNull(Rs3("NoMonth").value), 0, Rs3("NoMonth").value)
Else
GetSettingsVacPeriod = 0
End If
End Function

  Public Function GetSettingsVacDate(Optional RecDate As Date, Optional ByRef ID As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset

sql = "select * from TblVacationSettingsDet  "
Set Rs3 = New ADODB.Recordset
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Set Rs3 = New ADODB.Recordset
sql = "Select * from TblVacationSettingsDet "
sql = sql & " where FrmDate<= " & SQLDate(RecDate, True) & " "
sql = sql & " and  ToDate >= " & SQLDate(RecDate, True) & " "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetSettingsVacDate = True
ID = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
Else
ID = 0
GetSettingsVacDate = False
End If
Else
GetSettingsVacDate = True
ID = 1
End If
End Function
Public Function GetSettingsVacDateAllow(Optional RecDate As Date, Optional ID As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset

sql = "select * from TblVacationSettingsDet  "
Set Rs3 = New ADODB.Recordset
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Set Rs3 = New ADODB.Recordset
sql = "Select * from TblVacationSettingsDet "
sql = sql & " where AlowDate >= " & SQLDate(RecDate, True) & " and ID=" & ID & " "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetSettingsVacDateAllow = True
Else
GetSettingsVacDateAllow = False
End If
Else
GetSettingsVacDateAllow = True
End If
End Function


Public Function URLEncode2(ByVal str As String) As String
    Dim intLen As Integer
    Dim X As Integer
    Dim curChar As Long
    Dim newStr As String

    intLen = Len(str)
    newStr = ""

    For X = 1 To intLen
        curChar = Asc(mId$(str, X, 1))

        If (curChar < 48 Or curChar > 57) And (curChar < 65 Or curChar > 90) And (curChar < 97 Or curChar > 122) Then
            newStr = newStr & "%" & Hex(curChar)
        Else
            newStr = newStr & CHR(curChar)
        End If

    Next X

    URLEncode2 = newStr
End Function

Public Function URLEncode( _
    ByVal URL As String, _
    Optional ByVal SpacePlus As Boolean = True) As String

    Dim cchEscaped As Long
    Dim HRESULT As Long

    If Len(URL) > INTERNET_MAX_URL_LENGTH Then
        Err.Raise &H8004D700, "URLUtility.URLEncode", _
                  "URL parameter too long"
    End If

    cchEscaped = Len(URL) * 1.5
    URLEncode = String$(cchEscaped, 0)
    HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
    If HRESULT = E_POINTER Then
        URLEncode = String$(cchEscaped, 0)
        HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
    End If

    If HRESULT <> S_OK Then
        Err.Raise Err.LastDllError, "URLUtility.URLEncode", _
                  "System error"
    End If

    URLEncode = left$(URLEncode, cchEscaped)
    If SpacePlus Then
        URLEncode = Replace$(URLEncode, "+", "%2B")
        URLEncode = Replace$(URLEncode, " ", "+")
    End If
End Function


Public Function ConvertToUnicode(st As String) As String
    Dim chrArray(0 To 149) As String
    Dim unicodeArray(0 To 149) As String

10  chrArray(0) = "°"
20  unicodeArray(0) = "060D"
30  chrArray(1) = "∫"
40  unicodeArray(1) = "061B"
50  chrArray(2) = "ø"
60  unicodeArray(2) = "061F"
70  chrArray(3) = "¡"
80  unicodeArray(3) = "0621"
90  chrArray(4) = "¬"
100     unicodeArray(4) = "0622"
110     chrArray(5) = "√"
120     unicodeArray(5) = "0623"
130     chrArray(6) = "ƒ"
140     unicodeArray(6) = "0624"
150     chrArray(7) = "≈"
160     unicodeArray(7) = "0625"
170     chrArray(8) = "∆"
180     unicodeArray(8) = "0626"
190     chrArray(9) = "«"
200     unicodeArray(9) = "0627"
210     chrArray(10) = "»"
220     unicodeArray(10) = "0628"
230     chrArray(11) = "…"
240     unicodeArray(11) = "0629"
250     chrArray(12) = " "
260     unicodeArray(12) = "062A"
270     chrArray(13) = "À"
280     unicodeArray(13) = "062B"
290     chrArray(14) = "Ã"
300     unicodeArray(14) = "062C"
310     chrArray(15) = "Õ"
320     unicodeArray(15) = "062D"
330     chrArray(16) = "Œ"
340     unicodeArray(16) = "062E"
350     chrArray(17) = "œ"
360     unicodeArray(17) = "062F"
370     chrArray(18) = "–"
380     unicodeArray(18) = "0630"
390     chrArray(19) = "—"
400     unicodeArray(19) = "0631"
410     chrArray(20) = "“"
420     unicodeArray(20) = "0632"
430     chrArray(21) = "”"
440     unicodeArray(21) = "0633"
450     chrArray(22) = "‘"
460     unicodeArray(22) = "0634"
470     chrArray(23) = "’"
480     unicodeArray(23) = "0635"
490     chrArray(24) = "÷"
500     unicodeArray(24) = "0636"
510     chrArray(25) = "ÿ"
520     unicodeArray(25) = "0637"
530     chrArray(26) = "Ÿ"
540     unicodeArray(26) = "0638"
550     chrArray(27) = "⁄"
560     unicodeArray(27) = "0639"
570     chrArray(28) = "€"
580     unicodeArray(28) = "063A"
590     chrArray(29) = "›"
600     unicodeArray(29) = "0641"
610     chrArray(30) = "ﬁ"
620     unicodeArray(30) = "0642"
630     chrArray(31) = "ﬂ"
640     unicodeArray(31) = "0643"
650     chrArray(32) = "·"
660     unicodeArray(32) = "0644"
670     chrArray(33) = "„"
680     unicodeArray(33) = "0645"
690     chrArray(34) = "‰"
700     unicodeArray(34) = "0646"
710     chrArray(35) = "Â"
720     unicodeArray(35) = "0647"
730     chrArray(36) = "Ê"
740     unicodeArray(36) = "0648"
750     chrArray(37) = "Ï"
760     unicodeArray(37) = "0649"
770     chrArray(38) = "Ì"
780     unicodeArray(38) = "064A"
790     chrArray(39) = "‹"
800     unicodeArray(39) = "0640"
810     chrArray(40) = ""
820     unicodeArray(40) = "064B"
830     chrArray(41) = "Ò"
840     unicodeArray(41) = "064C"
850     chrArray(42) = "Ú"
860     unicodeArray(42) = "064D"
870     chrArray(43) = "Û"
880     unicodeArray(43) = "064E"
890     chrArray(44) = "ı"
900     unicodeArray(44) = "064F"
910     chrArray(45) = "ˆ"
920     unicodeArray(45) = "0650"
930     chrArray(46) = "¯"
940     unicodeArray(46) = "0651"
950     chrArray(47) = "˙"
960     unicodeArray(47) = "0652"
970     chrArray(48) = "!"
980     unicodeArray(48) = "0021"
990     chrArray(49) = """"
1000    unicodeArray(49) = "0022"
1010    chrArray(50) = "#"
1020    unicodeArray(50) = "0023"
1030    chrArray(51) = "$"
1040    unicodeArray(51) = "0024"
1050    chrArray(52) = "%"
1060    unicodeArray(52) = "0025"
1070    chrArray(53) = "&"
1080    unicodeArray(53) = "0026"
1090    chrArray(54) = "'"
1100    unicodeArray(54) = "0027"
1110    chrArray(55) = "("
1120    unicodeArray(55) = "0028"
1130    chrArray(56) = ")"
1140    unicodeArray(56) = "0029"
1150    chrArray(57) = "*"
1160    unicodeArray(57) = "002A"
1170    chrArray(58) = "+"
1180    unicodeArray(58) = "002B"
1190    chrArray(59) = ","
1200    unicodeArray(59) = "002C"
1210    chrArray(60) = "-"
1220    unicodeArray(60) = "002D"
1230    chrArray(61) = "."
1240    unicodeArray(61) = "002E"
1250    chrArray(62) = "/"
1260    unicodeArray(62) = "002F"
1270    chrArray(63) = "0"
1280    unicodeArray(63) = "0030"
1290    chrArray(64) = "1"
1300    unicodeArray(64) = "0031"
1310    chrArray(65) = "2"
1320    unicodeArray(65) = "0032"
1330    chrArray(66) = "3"
1340    unicodeArray(66) = "0033"
1350    chrArray(67) = "4"
1360    unicodeArray(67) = "0034"
1370    chrArray(68) = "5"
1380    unicodeArray(68) = "0035"
1390    chrArray(69) = "6"
1400    unicodeArray(69) = "0036"
1410    chrArray(70) = "7"
1420    unicodeArray(70) = "0037"
1430    chrArray(71) = "8"
1440    unicodeArray(71) = "0038"
1450    chrArray(72) = "9"
1460    unicodeArray(72) = "0039"
1470    chrArray(73) = ":"
1480    unicodeArray(73) = "003A"
1490    chrArray(74) = ""
1500    unicodeArray(74) = "003B"
1510    chrArray(75) = "<"
1520    unicodeArray(75) = "003C"
1530    chrArray(76) = "="
1540    unicodeArray(76) = "003D"
1550    chrArray(77) = ">"
1560    unicodeArray(77) = "003E"
1570    chrArray(78) = "?"
1580    unicodeArray(78) = "003F"
1590    chrArray(79) = "@"
1600    unicodeArray(79) = "0040"
1610    chrArray(80) = "A"
1620    unicodeArray(80) = "0041"
1630    chrArray(81) = "B"
1640    unicodeArray(81) = "0042"
1650    chrArray(82) = "C"
1660    unicodeArray(82) = "0043"
1670    chrArray(83) = "D"
1680    unicodeArray(83) = "0044"
1690    chrArray(84) = "E"
1700    unicodeArray(84) = "0045"
1710    chrArray(85) = "F"
1720    unicodeArray(85) = "0046"
1730    chrArray(86) = "G"
1740    unicodeArray(86) = "0047"
1750    chrArray(87) = "H"
1760    unicodeArray(87) = "0048"
1770    chrArray(88) = "I"
1780    unicodeArray(88) = "0049"
1790    chrArray(89) = "J"
1800    unicodeArray(89) = "004A"
1810    chrArray(90) = "K"
1820    unicodeArray(90) = "004B"
1830    chrArray(91) = "L"
1840    unicodeArray(91) = "004C"
1850    chrArray(92) = "M"
1860    unicodeArray(92) = "004D"
1870    chrArray(93) = "N"
1880    unicodeArray(93) = "004E"
1890    chrArray(94) = "O"
1900    unicodeArray(94) = "004F"
1910    chrArray(95) = "P"
1920    unicodeArray(95) = "0050"
1930    chrArray(96) = "Q"
1940    unicodeArray(96) = "0051"
1950    chrArray(97) = "R"
1960    unicodeArray(97) = "0052"
1970    chrArray(98) = "S"
1980    unicodeArray(98) = "0053"
1990    chrArray(99) = "T"
2000    unicodeArray(99) = "0054"
2010    chrArray(100) = "U"
2020    unicodeArray(100) = "0055"
2030    chrArray(101) = "V"
2040    unicodeArray(101) = "0056"
2050    chrArray(102) = "W"
2060    unicodeArray(102) = "0057"
2070    chrArray(103) = "X"
2080    unicodeArray(103) = "0058"
2090    chrArray(104) = "Y"
2100    unicodeArray(104) = "0059"
2110    chrArray(105) = "Z"
2120    unicodeArray(105) = "005A"
2130    chrArray(106) = "[" '"("
2140    unicodeArray(106) = "005B"
2150    chrArray(107) = Trim("\ ")
2160    unicodeArray(107) = "005C"
2170    chrArray(108) = "]" '")"
2180    unicodeArray(108) = "005D"
2190    chrArray(109) = "^"
2200    unicodeArray(109) = "005E"
2210    chrArray(110) = "_"
2220    unicodeArray(110) = "005F"
2230    chrArray(111) = "`"
2240    unicodeArray(111) = "0060"
2250    chrArray(112) = "a"
2260    unicodeArray(112) = "0061"
2270    chrArray(113) = "b"
2280    unicodeArray(113) = "0062"
2290    chrArray(114) = "c"
2300    unicodeArray(114) = "0063"
2310    chrArray(115) = "d"
2320    unicodeArray(115) = "0064"
2330    chrArray(116) = "e"
2340    unicodeArray(116) = "0065"
2350    chrArray(117) = "f"
2360    unicodeArray(117) = "0066"
2370    chrArray(118) = "g"
2380    unicodeArray(118) = "0067"
2390    chrArray(119) = "h"
2400    unicodeArray(119) = "0068"
2410    chrArray(120) = "i"
2420    unicodeArray(120) = "0069"
2430    chrArray(121) = "j"
2440    unicodeArray(121) = "006A"
2450    chrArray(122) = "k"
2460    unicodeArray(122) = "006B"
2470    chrArray(123) = "l"
2480    unicodeArray(123) = "006C"
2490    chrArray(124) = "m"
2500    unicodeArray(124) = "006D"
2510    chrArray(125) = "n"
2520    unicodeArray(125) = "006E"
2530    chrArray(126) = "o"
2540    unicodeArray(126) = "006F"
2550    chrArray(127) = "p"
2560    unicodeArray(127) = "0070"
2570    chrArray(128) = "q"
2580    unicodeArray(128) = "0071"
2590    chrArray(129) = "r"
2600    unicodeArray(129) = "0072"
2610    chrArray(130) = "s"
2620    unicodeArray(130) = "0073"
2630    chrArray(131) = "t"
2640    unicodeArray(131) = "0074"
2650    chrArray(132) = "u"
2660    unicodeArray(132) = "0075"
2670    chrArray(133) = "v"
2680    unicodeArray(133) = "0076"
2690    chrArray(134) = "w"
2700    unicodeArray(134) = "0077"
2710    chrArray(135) = "x"
2720    unicodeArray(135) = "0078"
2730    chrArray(136) = "y"
2740    unicodeArray(136) = "0079"
2750    chrArray(137) = "z"
2760    unicodeArray(137) = "007A"
2770    chrArray(138) = "{"
2780    unicodeArray(138) = "007B"
2790    chrArray(139) = "|"
2800    unicodeArray(139) = "007C"
2810    chrArray(140) = "}"
2820    unicodeArray(140) = "007D"
2830    chrArray(141) = "~"
2840    unicodeArray(141) = "007E"
2850    chrArray(142) = "©"
2860    unicodeArray(142) = "00A9"
2870    chrArray(143) = "Æ"
2880    unicodeArray(143) = "00AE"
2890    chrArray(144) = "˜"
2900    unicodeArray(144) = "00F7"
2910    chrArray(145) = "◊"
2920    unicodeArray(145) = "00F7"
2930    chrArray(146) = "ß"
2940    unicodeArray(146) = "00A7"
2950    chrArray(147) = " "
2960    unicodeArray(147) = "0020"
2970    chrArray(148) = CHR$(13)
2980    unicodeArray(148) = "000D"
2990    chrArray(149) = "\r"
3000    unicodeArray(149) = "000A"

        Dim strResult As String, i As Integer, c As Integer
3010    strResult = ""

3020    For i = 1 To Len(st)
3030        For c = 0 To 149

3040            If (chrArray(c) = mId(st, i, 1)) Then
3050                strResult = strResult & unicodeArray(c)
3060            End If

3070        Next c
3080    Next i

3090    ConvertToUnicode = strResult

End Function

 Function SendEmailForCustomer(CusID As Integer, subject1 As String, Msg As String, ByRef msgstatus As String)
'Dim subject As String
    'Dim msg As String
    Dim CompanyName As String
        Dim cOptions As ClsCompanyInfo
        'Dim msg As String
    Set cOptions = New ClsCompanyInfo
    Dim Email As String
    Dim customername As String
    customername = ""
Email = GetCustomerEmail(CusID, customername)
    CompanyName = cOptions.ArabCompanyName & CHR(13) & CurrentBranchName


      subject = " «·”«œ… / " & customername & " " & subject1

      Dim RetVal As String
   RetVal = SendMail(Email, _
        Trim$(subject), _
         "»ﬁŒ…", _
        Msg, _
        "txtServer", _
        25, _
         "txtUsername", _
        "txtPassword", _
       "", _
         False, True)
         msgstatus = IIf(RetVal = "ok", " „ «—”«· «·“Ì«—…", RetVal)

 End Function

Public Function get_TblPaymentTypet(ID As Long, _
                                 filed As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from TblPaymentType where PaymentID=" & ID

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount = 0 Then get_TblPaymentTypet = "": Exit Function
    If IsNull(Rs3(filed).value) Then get_TblPaymentTypet = "": Exit Function
    If Not IsNull(Rs3(filed).value) Then get_TblPaymentTypet = Rs3(filed).value: Exit Function
    Rs3.Close

End Function
Public Function ChekPayedSalary(Optional YearID As Integer, Optional MonthID As Integer, Optional BranchID As Integer) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     BranchId, MONTH(RecordDate) AS MonthID, YEAR(RecordDate) AS YearID"
sql = sql & " From dbo.emp_salary"
sql = sql & " Where (year(RecordDate) = " & YearID & ") And (Month(RecordDate) = " & MonthID & ")"
sql = sql & " AND BranchID=" & BranchID

Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
ChekPayedSalary = True
Else
ChekPayedSalary = False
End If
End Function
Public Function GetReustValue(Optional StoreID As Long, Optional ItemID As Long) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim AllowQty As Double
Dim UnitFacort As Double
Dim Qty As Double
sql = " SELECT     Xb.Qty, Xb.StoreID, Xb.ItemID, Xb.UnitFactor, Xb.UnitID, BX.QNty"
sql = sql & " FROM         (SELECT     dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.UnitFactor,"
sql = sql & "                                              dbo.TblSettsRequestLimitDet.unitid"
sql = sql & "                        FROM         dbo.TblSettsRequestLimitDet INNER JOIN"
sql = sql & "                                              dbo.Transaction_Details ON dbo.TblSettsRequestLimitDet.ItemID = dbo.Transaction_Details.Item_ID"
sql = sql & "                        Where (dbo.TblSettsRequestLimitDet.typ = 0)"
sql = sql & "                        GROUP BY dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.UnitFactor,"
sql = sql & "                                              dbo.TblSettsRequestLimitDet.UnitID) Xb INNER JOIN"
sql = sql & "                          (SELECT     SUM(dbo.TransactionTypes.StockEffect * dbo.Transaction_Details.Quantity) AS QNty, dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID"
sql = sql & "                             FROM         dbo.Transactions INNER JOIN"
sql = sql & "                                                   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
sql = sql & "                                                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
sql = sql & "                             GROUP BY Item_ID, StoreID) BX ON BX.Item_ID = Xb.ItemID  AND BX.StoreID = Xb.StoreID"
sql = sql & "  Where Xb.StoreID = " & StoreID & " And Xb.ItemID = " & ItemID & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
UnitFacort = IIf(IsNull(Rs3("UnitFactor").value), 0, Rs3("UnitFactor").value)
Qty = IIf(IsNull(Rs3("Qty").value), 0, Rs3("Qty").value)
AllowQty = IIf(IsNull(Rs3("QNty").value), 0, Rs3("QNty").value)
If UnitFacort > 0 Then
AllowQty = AllowQty / UnitFacort
End If
GetReustValue = AllowQty - Qty
Else
GetReustValue = 0
End If
End Function

Public Function DescUnitFact(Optional ItemID As Long, Optional UntID As Long) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "SELECT     TOP 100 PERCENT UnitFactor, UnitID"
sql = sql & " From dbo.TblItemsUnits"
sql = sql & " Where (ItemID = " & ItemID & ") And (unitid = " & UntID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
DescUnitFact = IIf(IsNull(rs2("UnitFactor").value), 1, rs2("UnitFactor").value)
Else
DescUnitFact = 0
End If
End Function
Public Function DescUnit(Optional ItemID As Long, Optional UnitID As Long) As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     TOP 100 PERCENT dbo.TblItemsUnits.FactorBySmallUnit, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
sql = sql & " FROM         dbo.TblItemsUnits INNER JOIN"
sql = sql & "                      dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
sql = sql & "  Where (dbo.TblItemsUnits.ItemID = " & ItemID & ") And (dbo.TblItemsUnits.DefaultUnit = 1)"
sql = sql & " ORDER BY dbo.TblItemsUnits.ItemID"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
DescUnit = DescUnitFact(ItemID, UnitID)
DescUnit = DescUnit & " "
If SystemOptions.UserInterface = ArabicInterface Then
DescUnit = DescUnit & IIf(IsNull(rs2("UnitName").value), "", rs2("UnitName").value)
Else
DescUnit = DescUnit & IIf(IsNull(rs2("UnitNamee").value), "", rs2("UnitNamee").value)
End If
Else
DescUnit = ""
End If
End Function
Public Sub GetCodeIDProject(Optional ByRef ID As Double = 0, Optional ByRef fullcode As String, Optional getMaterial_account As Integer = 0, Optional ByRef Material_account As String)
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = "Select id ,Fullcode ,Material_account from projects "
If ID <> 0 Then
sql = sql & " where id=" & ID & ""
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
fullcode = IIf(IsNull(Rs4("Fullcode").value), "", Rs4("Fullcode").value)
Else
fullcode = ""
End If
Else

sql = sql & " where Fullcode=N'" & fullcode & "'"


Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
ID = IIf(IsNull(Rs4("id").value), "", Rs4("id").value)
                If getMaterial_account <> 0 Then
                             Material_account = IIf(IsNull(Rs4("Material_account").value), "", Rs4("Material_account").value)
                End If
Else
ID = 0
End If
End If
End Sub
Public Function ChicIsLotNo(Optional ItemID As Long) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     ItemID, ChkLot"
sql = sql & " From dbo.TblItems"
sql = sql & " Where (ItemID = " & ItemID & ") And (ChkLot = 1)"
Rs3.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
ChicIsLotNo = True
Else
ChicIsLotNo = False
End If
End Function
'Public Function UpdateTransactionsCost(Transaction_IDs As String)
'If SystemOptions.AllowCostnNewShape = False Then Exit Function
'Dim sql As String
'Dim Rs3 As ADODB.Recordset
'Set Rs3 = New ADODB.Recordset
'             Dim OldQty As Double
'             Dim OldCost As Double
'              Dim NewQty As Double
'               Dim NewCost As Double
'               Dim StockEffect As Integer
'
'               Dim StoreID As Double
'Dim Item_ID As Double
'Dim Transaction_Date As Date
'Dim Transaction_ID As Double
'
''sql = "Select * from TbLSheft where TypHour=1 "
'sql = "SELECT   dbo.Transaction_Details.QtyBySmalltUnit  , dbo.Transactions.Transaction_Type, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, "
'sql = sql & "                       dbo.Transaction_Details.OldCost, dbo.Transaction_Details.OldQty, dbo.Transaction_Details.NewCost, dbo.Transaction_Details.NewQty,"
'                      sql = sql & "  dbo.Transactions.Transaction_ID , dbo.TransactionTypes.StockEffect , dbo.Transactions.StoreID, dbo.Transactions.Transaction_Date, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price"
'
'
'sql = sql & "  FROM         dbo.Transactions INNER JOIN"
'sql = sql & "                       dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
'sql = sql & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
'sql = sql & "  Where (dbo.Transactions.Transaction_ID in ( " & Transaction_IDs & " ))"
'
'
''WaelCost
'sql = sql & "  ORDER BY dbo.Transactions.Transaction_Date DESC ,dbo.TransactionTypes.StockEffect, dbo.Transactions.Transaction_ID DESC, dbo.Transaction_Details.ID DESC"
'
'
'Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'If Rs3.RecordCount > 0 Then
'For i = 1 To Rs3.RecordCount
' StockEffect = IIf(IsNull(Rs3("StockEffect").value), 0, Rs3("StockEffect").value)
' StoreID = IIf(IsNull(Rs3("StoreID").value), 0, Rs3("StoreID").value)
' Item_ID = IIf(IsNull(Rs3("Item_ID").value), 0, Rs3("Item_ID").value)
'  Transaction_Date = IIf(IsNull(Rs3("Transaction_Date").value), Date, Rs3("Transaction_Date").value)
'   Transaction_ID = IIf(IsNull(Rs3("Transaction_ID").value), 0, Rs3("Transaction_ID").value)
'
' getItemCostData Transaction_Date, Item_ID, StoreID, Transaction_ID, OldQty, OldCost, NewQty, NewCost
' Dim QtyBySmalltUnit As Double
'             If StockEffect = -1 Then
'              QtyBySmalltUnit = IIf(IsNull(Rs3("QtyBySmalltUnit").value), 1, Rs3("QtyBySmalltUnit").value)
'
'                  Rs3("OldQty").value = NewQty
'                   Rs3("OldCost").value = NewCost
'
'                  Rs3("NewQty").value = Rs3("OldQty").value - Rs3("Quantity").value
'                   Rs3("NewCost").value = Rs3("OldCost").value ' ((Rs3("OldQty").value * Rs3("OldCost").value) + (Rs3("Quantity").value * Rs3("Price").value)) / (Rs3("Quantity").value + Rs3("OldQty").value)
'
'
'            Rs3.update
'
'                ElseIf StockEffect = 1 Then 'input
'                 QtyBySmalltUnit = IIf(IsNull(Rs3("QtyBySmalltUnit").value), 1, Rs3("QtyBySmalltUnit").value)
'
'                     Rs3("OldQty").value = NewQty
'                       Rs3("OldCost").value = NewCost
'
'                      Rs3("NewQty").value = Rs3("Quantity").value + Rs3("OldQty").value
'                      If val(Rs3("Quantity").value + Rs3("OldQty").value) <> 0 Then
'                       Rs3("NewCost").value = ((Round(Rs3("OldQty").value, 4) * Round(Rs3("OldCost").value, 4)) + (Round(Rs3("Quantity").value, 4) * Round(Rs3("Price").value, 4))) / (Round(Rs3("Quantity").value, 4) + Round(Rs3("OldQty").value, 4))           'IIf(Rs3("Quantity").value + Rs3("OldQty").value <> 0, Rs3("Quantity").value + Rs3("OldQty").value, 0)
'                       Else
'                      Rs3("NewCost").value = 0
'                       End If
'                       Rs3.update
'
'                Else
'
'
'                End If
'Rs3.MoveNext
'Next i
'Else
'
'End If
'
'
'
'End Function
Public Function updateCopyNo(tablename As String, Filedname As String, transactionfiledname As String, Transaction_ID As Double)
'        updateCopyNo "Transactions", "CopyNO", "Transaction_ID", val(Me.XPTxtBillID.Text)

Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset

sql = ""
sql = "SELECT     " & Filedname & ""
sql = sql & " From dbo." & tablename & ""
sql = sql & " Where (" & transactionfiledname & " = " & Transaction_ID & ")"

Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
 lastCopyno = IIf(IsNull(Rs3(Filedname).value), 0, Rs3(Filedname).value)

 Cn.Execute "update  " & tablename & " set " & Filedname & "=" & Filedname & "+1 where Transaction_ID=" & Transaction_ID
 End If
End Function

Public Function NoHourInShift(Optional ByRef NoHour As Double, Optional EmpID As Double) As Boolean

Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
'sql = "Select * from TbLSheft where TypHour=1 "
sql = "SELECT     dbo.TbLSheft.NoHourManaula"
sql = sql & "  FROM         dbo.TbLSheft INNER JOIN"
sql = sql & "                        dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
sql = sql & "  Where (dbo.TbLSheft.TypHour = 1) And (dbo.TblShiftWorker.EmpID = " & EmpID & ")"

Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
NoHourInShift = True
NoHour = IIf(IsNull(Rs3("NoHourManaula").value), 0, Rs3("NoHourManaula").value)
Else
NoHour = 0
NoHourInShift = False
End If
End Function
  '***************************
   Public Function GetOrdersData(Optional ID As Double, Optional ByRef EmpName As String _
 , Optional ByRef EmpMbile As String, Optional ByRef OrdeNo As String, Optional NoteSerial1 As Double, Optional SeasonsID As Double) As Double
Dim Rs4 As ADODB.Recordset
Dim sql As String
Set Rs4 = New ADODB.Recordset
sql = "SELECT    * from tblbookingrequest where ID=" & ID & " and StusID=1"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
EmpName = IIf(IsNull(Rs4("EmpName").value), "", Rs4("EmpName").value)
EmpMbile = IIf(IsNull(Rs4("EmpMbile").value), "", Rs4("EmpMbile").value)
SeasonsID = IIf(IsNull(Rs4("SeasonsID").value), 0, Rs4("SeasonsID").value)
NoteSerial1 = IIf(IsNull(Rs4("NoteSerial1").value), 0, Rs4("NoteSerial1").value)
Else
NoteSerial1 = 0
SeasonsID = 0
EmpName = ""
EmpMbile = ""
End If
End Function

   Public Function CheckDelLocations(CustomerlocationID As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From Transactions  Where CustomerlocationID=" & CustomerlocationID
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckDelLocations = False
    Else
        CheckDelLocations = True
    End If

    rs.Close
    Set rs = Nothing
End Function

  Public Function GeTuserFullCode(UserID As Double) As String
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

If UserID = 1 Then
GeTuserFullCode = "0000"
Exit Function
End If
    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = " SELECT     dbo.TblUsers.UserID, dbo.TblEmployee.Fullcode"
sql = sql & "  FROM         dbo.TblUsers INNER JOIN"
sql = sql & "  dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
sql = sql & "   WHERE     (dbo.TblUsers.UserID = " & UserID & ")"
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Rs1.RecordCount > 0 Then
        GeTuserFullCode = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
    Else
        GeTuserFullCode = 0
    End If
    Rs1.Close
End Function

  Public Function TimeStamp(date1 As Date) As String
Dim StartDate As String
Dim EndTime As String
Dim startTime As String
Dim EndDate As String
Dim dblStart As Double
Dim dblEnd As Double
Dim DateTimeStart As Date
Dim DateTimeEnd As Date
Dim TotalHrs As String
StartDate = "1/1/1970"
startTime = "00:00:00"
EndDate = CStr(date1)
EndTime = CStr(Time)
DateTimeStart = FormatDateTime(StartDate & " " & startTime)
DateTimeEnd = FormatDateTime(EndDate & " " & EndTime)
TimeStamp = DateDiff("s", DateTimeStart, DateTimeEnd, vbUseSystemDayOfWeek, _
vbUseSystem)
End Function


  Public Function GetItemsData(ByRef ItemName As String, Optional ByRef ItemID As Double, Optional ByRef fullcode As String, Optional ByRef PartNo As String) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
If ItemName <> "" Then
sql = "Select * from  TblItems where ItemName ='" & ItemName & "'"
Else
sql = "Select * from  TblItems where code ='" & fullcode & "'"
End If

rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
ItemID = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
itemcode = IIf(IsNull(rs2("FullCode").value), "", rs2("FullCode").value)
PartNo = IIf(IsNull(rs2("PartNo").value), "", rs2("PartNo").value)
ItemName = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
Else
ItemID = 0
itemcode = ""
PartNo = ""
fullcode = ""
End If
End Function

Public Function GetStoreData(ByRef StoreName As String, Optional ByRef StoreID As Double, Optional ByRef BranchID As Double, Optional fullcode As String) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
If fullcode = "" Then
sql = "Select * from  TblStore where StoreName ='" & StoreName & "'"
Else
sql = "Select * from  TblStore where Code ='" & fullcode & "'"
End If

rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
StoreID = IIf(IsNull(rs2("storeid").value), 0, rs2("storeid").value)
BranchID = IIf(IsNull(rs2("BranchId").value), 0, rs2("BranchId").value)
StoreName = IIf(IsNull(rs2("storename").value), 0, rs2("storename").value)
Else
StoreName = ""
StoreID = 0
BranchID = 0
End If
End Function

 Public Function CheCkTriningRequest() As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from  TblTrainingRequest"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheCkTriningRequest = True
Else
CheCkTriningRequest = False
End If
End Function
 Public Function GetBranchnmeFromnotes(NoteID As Double, Optional ByRef branch_id As Double, Optional ByRef branch_name As String, Optional ByRef branch_namee As String, Optional Vat As Double)
    Dim sql As String
    Dim rs As New ADODB.Recordset


        sql = "SELECT   vat,  dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
sql = sql & "  FROM         dbo.Notes INNER JOIN"
sql = sql & "                        dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
sql = sql & "   Where (dbo.Notes.noteID = " & NoteID & ")"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then

        branch_id = IIf(IsNull(rs("branch_no").value), 0, rs("branch_no").value)
     branch_name = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
     branch_namee = IIf(IsNull(rs("branch_namee").value), 0, rs("branch_namee").value)
     Vat = IIf(IsNull(rs("vat").value), 0, rs("vat").value)

    Else
    Vat = 0
    branch_id = 0
    branch_name = ""
      branch_namee = ""

    End If
    rs.Close
End Function

    Public Function GetMixIdFormCode(MixCode As String) As Double
    Dim sql As String
    Dim rs As New ADODB.Recordset


        sql = "SELECT     ID From dbo.TblDefComItem WHERE     (MaxNo = '" & MixCode & "')"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then

        GetMixIdFormCode = IIf(IsNull(rs("ID").value), 0, rs("ID").value)

    Else
        GetMixIdFormCode = 0

    End If
    rs.Close
End Function
Function CheckRemainsetyforprojectopr(project_id1 As Double, pand_id As Double, oper_id As Double, Item_ID As Double, Transaction_ID As Double, TransQty As Double, txtmodflage As String, Optional ByRef OPRQTY As Double, Optional ByRef IssuedOprQty As Double) As Double

 Dim sql As String
    Dim rs As New ADODB.Recordset

Dim SQL1 As String
    Dim Rs1 As New ADODB.Recordset







    '„⁄—›… ﬂ„ÌÂ «·»‰œ
    SQL1 = "  SELECT      sum( dbo.TblMatrials.COUNT) OPRQTY   FROM     terms_operations"
SQL1 = SQL1 & " inner join TblMatrials  on dbo.TblMatrials.Opr = dbo.terms_operations.id"
SQL1 = SQL1 & "  Where 1 = 1   "
SQL1 = SQL1 & "   and  (dbo.terms_operations.Project_ID =" & project_id1 & ")"
SQL1 = SQL1 & "   and  (dbo.terms_operations.projectdes_id=" & pand_id & ")"

SQL1 = SQL1 & "    and dbo.terms_operations.OPRIDD =" & oper_id & ""
 SQL1 = SQL1 & "   and  dbo.TblMatrials.ItemID  =" & Item_ID & ""


Rs1.Open SQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Rs1.RecordCount > 0 Then

        OPRQTY = IIf(IsNull(Rs1("OPRQTY").value), 0, Rs1("OPRQTY").value)

    Else
        OPRQTY = 0

    End If
    Rs1.Close


sql = " SELECT      SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS IssuedOprQty"
sql = sql & "   FROM            dbo.Transactions INNER JOIN"
sql = sql & "                            dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
sql = sql & "                            dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
sql = sql & "   Where (dbo.Transaction_Details.project_id1 = " & project_id1 & ")"
sql = sql & "    AND (dbo.Transaction_Details.Pand_ID = " & pand_id & ")"
sql = sql & "    AND (dbo.Transaction_Details.Oper_ID = " & oper_id & ")"
sql = sql & "    AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
If txtmodflage = "E" Then
sql = sql & "    AND (dbo.Transactions.Transaction_ID <>  " & Transaction_ID & ")"
End If

sql = sql & "   AND (dbo.TransactionTypes.projectInclude = 1)"

rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then

        IssuedOprQty = IIf(IsNull(rs("IssuedOprQty").value), 0, rs("IssuedOprQty").value)

    Else
        IssuedOprQty = 0

    End If
    Dim remainqty As Double
   remainqty = OPRQTY + IssuedOprQty
  remainqty = remainqty - TransQty
  CheckRemainsetyforprojectopr = remainqty

    rs.Close
   '     rs1.Close
End Function


 Public Function ProjectItemsCheck(Optional projectId As Double, Optional ProjectDes_ID As Double, Optional OPRIDD As Double, Optional ItemID As Double) As Double
    Dim sql As String
    Dim rs As New ADODB.Recordset


    sql = "  SELECT       count( dbo.TblMatrials.ItemID) projectItems"
    sql = sql & "   FROM         dbo.TblMatrials RIGHT OUTER JOIN                       dbo.terms_operations ON dbo.TblMatrials.Opr = dbo.terms_operations.id LEFT OUTER JOIN                       dbo.TblItems ON dbo.TblMatrials.ItemID = dbo.TblItems.ItemID"
sql = sql & "    Where 1 = 1"
 sql = sql & "   and  (dbo.TblMatrials.ProjectID =" & projectId & ")"

 If ItemID <> 0 Then
sql = sql & "    and dbo.terms_operations.ProjectDes_ID =" & ProjectDes_ID
sql = sql & "   and  dbo.terms_operations.OPRIDD =" & OPRIDD
sql = sql & "   and  dbo.TblMatrials.ItemID  =" & ItemID


 End If
'--where Qtyissue<0
rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then

        ProjectItemsCheck = IIf(IsNull(rs("projectItems").value), 0, rs("projectItems").value)

    Else
        ProjectItemsCheck = 0

    End If
    rs.Close
End Function
  Public Function GetInstructorCode(Optional ByRef ID As Integer, Optional ByRef fullcode As String, Optional Type1 As Integer = 0)
    Dim sql As String
    Dim rs As New ADODB.Recordset

    If Type1 = 0 Then
        sql = "select * from TblInstructors where ID= " & ID
    ElseIf Type1 = 1 Then
    sql = "select * from TblInstructors where  FullCode ='" & fullcode & "'"
    End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then

        ID = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
        fullcode = IIf(IsNull(rs("FullCode").value), "", rs("FullCode").value)
    Else
        ID = 0
        fullcode = ""
    End If
    rs.Close
End Function



  Public Function GetInstudentGroupCode(Optional ByRef ID As Integer, Optional ByRef fullcode As String, Optional Type1 As Integer = 0, Optional ByRef BranchID As Integer)
    Dim sql As String
    Dim rs As New ADODB.Recordset

    If Type1 = 0 Then
        sql = "select * from TblStuGroup where ID= " & ID
    ElseIf Type1 = 1 Then
    sql = "select * from TblStuGroup where  Code ='" & fullcode & "'"
    End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then

        BranchID = IIf(IsNull(rs("BranchID").value), 0, rs("BranchID").value)
        ID = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
        fullcode = IIf(IsNull(rs("Code").value), "", rs("Code").value)
    Else
    BranchID = 0
        ID = 0
        fullcode = ""
    End If
    rs.Close
End Function

  Public Sub GetTriningStudentInformation(Optional ID As Double, Optional ByRef QualiID As Double, Optional ByRef SexID As Integer, Optional ByRef UQama As String _
, Optional ByRef phone As String, Optional ByRef Email As String, Optional ByRef Address As String, Optional ByRef DateBrithH As String, Optional ByRef DateBrith As Date, Optional ByRef Mobile As String, Optional ByRef BranchID As Integer = 0)
Dim sql As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
sql = " SELECT  *"
sql = sql & " From dbo.TblTrainingRequest"
sql = sql & " Where (TypeTrain = 1) And (ID = " & ID & ")"
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
QualiID = IIf(IsNull(Rs6("QualiID").value), 0, Rs6("QualiID").value)
SexID = IIf(IsNull(Rs6("SexID").value), -1, Rs6("SexID").value)
UQama = IIf(IsNull(Rs6("UQama").value), "", Rs6("UQama").value)
phone = IIf(IsNull(Rs6("Phone").value), "", Rs6("Phone").value)
Email = IIf(IsNull(Rs6("Email").value), "", Rs6("Email").value)
Mobile = IIf(IsNull(Rs6("Mobile").value), "", Rs6("Mobile").value)
DateBrithH = IIf(IsNull(Rs6("DateBrithH").value), "", Rs6("DateBrithH").value)
Address = IIf(IsNull(Rs6("Address").value), "", Rs6("Address").value)
DateBrith = IIf(IsNull(Rs6("DateBrith").value), Date, Rs6("DateBrith").value)
BranchID = IIf(IsNull(Rs6("BranchID").value), 0, Rs6("BranchID").value)
Else
BranchID = 0
Mobile = ""
Address = ""
QualiID = 0
SexID = -1
UQama = ""
phone = ""
Email = ""

End If
End Sub
 Public Sub GetContStudentInformation(Optional ContID As Double, Optional ByRef CompID As Double, Optional ByRef NoStud As Double)
Dim sql As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
sql = " SELECT  *"
sql = sql & " From dbo.TblContrStudent"
sql = sql & " Where (ContType = 1) And (ID = " & ContID & ")"
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
CompID = IIf(IsNull(Rs6("CompID").value), 0, Rs6("CompID").value)
NoStud = IIf(IsNull(Rs6("NoStud").value), 0, Rs6("NoStud").value)
Else
CompID = 0
NoStud = 0
End If
End Sub

Public Sub GetNominStudentInformation(Optional ContID As Double, Optional ByRef CompID As Double, Optional ByRef NoStud As Double, Optional ByRef ContNoID As Double)
Dim sql As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
sql = " SELECT  *"
sql = sql & " From dbo.TblStuCandidacy"
sql = sql & " Where (ID = " & ContID & ")"
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
CompID = IIf(IsNull(Rs6("CompID").value), 0, Rs6("CompID").value)
NoStud = IIf(IsNull(Rs6("NoStudCon").value), 0, Rs6("NoStudCon").value)
ContNoID = IIf(IsNull(Rs6("ContNoID").value), 0, Rs6("ContNoID").value)
Else
ContNoID = 0
CompID = 0
NoStud = 0
End If
End Sub
 Public Function GetStudentCode(Optional ByRef ID As Integer, Optional ByRef fullcode As String, Optional Type1 As Integer = 0, Optional ByRef UQama As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset

    If Type1 = 0 Then
        sql = "select * from TblStudent where ID= " & ID
    ElseIf Type1 = 1 Then
    sql = "select * from TblStudent where  FullCode ='" & fullcode & "'"
    ElseIf Type1 = 2 Then
    sql = "select * from TblStudent where  UQama ='" & UQama & "'"
    End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then

        ID = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
        fullcode = IIf(IsNull(rs("FullCode").value), "", rs("FullCode").value)
        UQama = IIf(IsNull(rs("UQama").value), "", rs("UQama").value)
    Else
    UQama = ""
        ID = 0
        fullcode = ""
    End If
    rs.Close
End Function


  Public Function GetInformationofStudent(Optional ID As Integer, Optional ByRef UQama As String, Optional ByRef StudentPhone As String, Optional ByRef DcbQualiID As Double = 0)
    Dim sql As String
    Dim Rs9 As New ADODB.Recordset
        sql = "select * from TblStudent where ID= " & ID
    Rs9.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Rs9.RecordCount > 0 Then
        UQama = IIf(IsNull(Rs9("UQama").value), "", Rs9("UQama").value)
        StudentPhone = IIf(IsNull(Rs9("StudentPhone").value), "", Rs9("StudentPhone").value)
        DcbQualiID = IIf(IsNull(Rs9("DcbQualiID").value), 0, Rs9("DcbQualiID").value)
    Else
    StudentPhone = ""
        UQama = ""
        DcbQualiID = 0
    End If
    Rs9.Close
End Function

Public Function GetCursInformation(Optional ID As Double = 0, Optional ByRef NoHour As Double, Optional Price As Double)
Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
sql = "Select * from TblStudentCurs where id=" & ID & ""
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
Price = IIf(IsNull(Rs5("Price").value), 0, Rs5("Price").value)
NoHour = IIf(IsNull(Rs5("NoHour").value), 0, Rs5("NoHour").value)
Else
NoHour = 0
Price = 0
End If
End Function
'*********************

Public Function ChekEmpInProject(Optional EmpID As Integer = 0, Optional MonthID As Integer = 0, Optional YearID As Integer = 0) As Boolean
Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
sql = " SELECT     dbo.opr_Employee.opr_type, dbo.opr_employee_details.pk_id, dbo.opr_employee_details.ContProjSalar, dbo.opr_employee_details.Emp_id, "
sql = sql & "                       dbo.opr_Employee.YEARS , dbo.opr_Employee.Months"
sql = sql & " FROM         dbo.opr_Employee LEFT OUTER JOIN"
sql = sql & "                      dbo.opr_employee_details ON dbo.opr_Employee.id = dbo.opr_employee_details.pk_id"
sql = sql & " WHERE     (dbo.opr_Employee.opr_type = 0) AND (dbo.opr_Employee.Years = " & YearID & ") AND (dbo.opr_Employee.Months = " & MonthID & ") AND (dbo.opr_employee_details.ContProjSalar = 2) AND"
sql = sql & "                      (dbo.opr_employee_details.Emp_id = " & EmpID & ")"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
ChekEmpInProject = True
Else
ChekEmpInProject = False
End If
End Function

 Public Sub RetriveOrderInformation(Optional ID As Double, Optional ByRef ProgrammID As Double, Optional VehicleNo As Double)
Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
sql = "Select * From tblbookingrequest  where id=" & ID & " and StusID=1"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, cmadcmdtext
If Rs5.RecordCount > 0 Then
ProgrammID = IIf(IsNull(Rs5("ProgrammID").value), 0, Rs5("ProgrammID").value)
VehicleNo = IIf(IsNull(Rs5("VehicleNo").value), 0, Rs5("VehicleNo").value)
Else
VehicleNo = 0
ProgrammID = 0
End If
End Sub

Public Sub GetProjectsBillInformation(Optional ID As Long, Optional ByRef project_no As String, Optional ByRef bill_to As Integer = 0, Optional ByRef branch_no As Integer _
, Optional ByRef total As Double, Optional ByRef Project_name As String, Optional ByRef ManualNO As String, Optional ByRef note_id As Double _
, Optional ByRef UserID As Long, Optional ByRef discount As Double, Optional ByRef advancedPayment As Double, Optional ByRef revenue_account As String, Optional ByRef Results As Double, Optional ByRef Remarks As String _
, Optional ByRef discount1value As Double, Optional ByRef discount2value As Double, Optional ByRef discount1ID As Integer, Optional ByRef discount2ID As Integer, Optional ByRef subContractorId As Long)
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim sql As String
sql = "SELECT     dbo.projects.Project_name AS Project_nameH, dbo.projects.Project_nameE, dbo.project_billl.*"
sql = sql & " FROM         dbo.project_billl LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.project_billl.project_no = dbo.projects.id"
sql = sql & " where dbo.project_billl.id =" & ID & ""
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
project_no = IIf(IsNull(Rs7("project_no").value), 0, Rs7("project_no").value)
bill_to = IIf(IsNull(Rs7("bill_to").value), 0, Rs7("bill_to").value)
branch_no = IIf(IsNull(Rs7("branch_no").value), 0, Rs7("branch_no").value)
total = IIf(IsNull(Rs7("total").value), 0, Rs7("total").value)
If SystemOptions.UserInterface = ArabicInterface Then
Project_name = IIf(IsNull(Rs7("Project_nameH").value), "", Rs7("Project_nameH").value)
Else
Project_name = IIf(IsNull(Rs7("Project_nameE").value), "", Rs7("Project_nameE").value)
End If
ManualNO = IIf(IsNull(Rs7("ManualNO").value), 0, Rs7("ManualNO").value)
note_id = IIf(IsNull(Rs7("note_id").value), 0, Rs7("note_id").value)
UserID = IIf(IsNull(Rs7("UserID").value), 0, Rs7("UserID").value)
discount = IIf(IsNull(Rs7("Discount").value), 0, Rs7("Discount").value)
advancedPayment = IIf(IsNull(Rs7("advancedPayment").value), 0, Rs7("advancedPayment").value)
revenue_account = IIf(IsNull(Rs7("revenue_account").value), "", Rs7("revenue_account").value)
Results = IIf(IsNull(Rs7("Results").value), 0, Rs7("Results").value)
Remarks = IIf(IsNull(Rs7("Remarks").value), "", Rs7("Remarks").value)
discount1value = IIf(IsNull(Rs7("discount1value").value), 0, Rs7("discount1value").value)
discount2value = IIf(IsNull(Rs7("discount2value").value), 0, Rs7("discount2value").value)
discount1ID = IIf(IsNull(Rs7("discount1ID").value), -1, Rs7("discount1ID").value)
discount2ID = IIf(IsNull(Rs7("discount2ID").value), -1, Rs7("discount2ID").value)
subContractorId = IIf(IsNull(Rs7("subContractorId").value), 0, Rs7("subContractorId").value)
Else
project_no = ""
bill_to = 0
branch_no = 0
total = 0
Project_name = ""
ManualNO = ""
note_id = 0
UserID = 0
discount = 0
advancedPayment = 0
revenue_account = ""
Results = 0
Remarks = ""
discount1value = 0
discount2value = 0
discount1ID = -1
discount2ID = -1
subContractorId = 0
End If
End Sub
Public Sub GetInfomationDividInvestment(Optional ID As Double, Optional ByRef Nourth As Double, Optional ByRef South As Double _
, Optional ByRef East As Double, Optional ByRef West As Double, Optional ByRef Area As Double)
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     ID, Nourth, South, East, West, Area"
sql = sql & " From dbo.TblDivInvestInformation"
sql = sql & " Where (ID = " & ID & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
Nourth = IIf(IsNull(Rs8("Nourth").value), 0, Rs8("Nourth").value)
South = IIf(IsNull(Rs8("South").value), 0, Rs8("South").value)
East = IIf(IsNull(Rs8("East").value), 0, Rs8("East").value)
West = IIf(IsNull(Rs8("West").value), 0, Rs8("West").value)
Area = IIf(IsNull(Rs8("Area").value), 0, Rs8("Area").value)
Else
West = 0
South = 0
Nourth = 0
East = 0
End If
End Sub


Function CheckSusAccounts() As Boolean
 Dim branch_name    As String
 Dim branch_namee    As String
 Dim SUM As Double
 Dim i As Integer
 Dim SUMDebit  As Double
 Dim SUMCrebit As Double

 CheckSusAccounts = False
 Dim rsBranch As New ADODB.Recordset
   My_SQL = "SELECT  * From TblBranchesData"

    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsBranch.RecordCount > 0 Then
        rsBranch.MoveFirst
    End If


 For Branch = 1 To rsBranch.RecordCount
BramchId = IIf(IsNull(rsBranch("branch_id").value), 0, rsBranch("branch_id").value)
branch_name = IIf(IsNull(rsBranch("branch_name").value), 0, rsBranch("branch_name").value)
branch_namee = IIf(IsNull(rsBranch("branch_namee").value), 0, rsBranch("branch_namee").value)
SUMDebit = 0
SUMCrebit = 0
    With FrmAccEditJournal.Fg_Journal

                       For i = .FixedRows To .rows - 1

                           If val(.TextMatrix(i, .ColIndex("BranchID"))) = BramchId Then
                               SUMDebit = SUMDebit + val(.TextMatrix(i, .ColIndex("DebitValue")))
                               SUMCrebit = SUMCrebit + val(.TextMatrix(i, .ColIndex("CreditValue")))
                           End If

                Next i
                     SUMDebit = Round(SUMDebit, SystemOptions.SysDefCurrencyForamt)
                 SUMCrebit = Round(SUMCrebit, SystemOptions.SysDefCurrencyForamt)


                    If val(SUMDebit) <> val(SUMCrebit) Then
                                               If SystemOptions.UserInterface = ArabicInterface Then
                                                       MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_name, vbCritical
                                               Else
                                                       MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_namee, vbCritical
                                               End If
                                                       CheckSusAccounts = False
                                               Exit Function

                               End If
    End With
    rsBranch.MoveNext
    Next Branch
    rsBranch.Close
    CheckSusAccounts = True


End Function

Function CheckSusAccounts1() As Boolean
 Dim branch_name    As String
 Dim branch_namee    As String
 Dim SUM As Double
 Dim i As Integer
 Dim SUMDebit  As Double
 Dim SUMCrebit As Double
 CheckSusAccounts1 = False
 Dim rsBranch As New ADODB.Recordset
   My_SQL = "SELECT  * From TblBranchesData"

    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsBranch.RecordCount > 0 Then
        rsBranch.MoveFirst
    End If


 For Branch = 1 To rsBranch.RecordCount
BramchId = IIf(IsNull(rsBranch("branch_id").value), 0, rsBranch("branch_id").value)
branch_name = IIf(IsNull(rsBranch("branch_name").value), 0, rsBranch("branch_name").value)
branch_namee = IIf(IsNull(rsBranch("branch_namee").value), 0, rsBranch("branch_namee").value)
SUMDebit = 0
SUMCrebit = 0
    With FrmAccEditJournal1.Fg_Journal

                       For i = .FixedRows To .rows - 1

                           If val(.TextMatrix(i, .ColIndex("BranchID"))) = BramchId Then
                               SUMDebit = SUMDebit + val(.TextMatrix(i, .ColIndex("DebitValue")))
                               SUMCrebit = SUMCrebit + val(.TextMatrix(i, .ColIndex("CreditValue")))
                           End If

                Next i
                     SUMDebit = Round(SUMDebit, SystemOptions.SysDefCurrencyForamt)
                 SUMCrebit = Round(SUMCrebit, SystemOptions.SysDefCurrencyForamt)


                    If Round(val(SUMDebit)) <> Round(val(SUMCrebit)) Then
                                               If SystemOptions.UserInterface = ArabicInterface Then
                                                       MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_name, vbCritical
                                               Else
                                                       MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_namee, vbCritical
                                               End If
                                                       CheckSusAccounts1 = False
                                               Exit Function

                               End If
    End With
    rsBranch.MoveNext
    Next Branch
    rsBranch.Close
    CheckSusAccounts1 = True


End Function
Function CheckSusAccounts3() As Boolean
 Dim branch_name    As String
 Dim branch_namee    As String
 Dim SUM As Double
 Dim i As Integer
 Dim SUMDebit  As Double
 Dim SUMCrebit As Double
 CheckSusAccounts3 = False
 Dim rsBranch As New ADODB.Recordset
   My_SQL = "SELECT  * From TblBranchesData"

    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsBranch.RecordCount > 0 Then
        rsBranch.MoveFirst
    End If


 For Branch = 1 To rsBranch.RecordCount
BramchId = IIf(IsNull(rsBranch("branch_id").value), 0, rsBranch("branch_id").value)
branch_name = IIf(IsNull(rsBranch("branch_name").value), 0, rsBranch("branch_name").value)
branch_namee = IIf(IsNull(rsBranch("branch_namee").value), 0, rsBranch("branch_namee").value)
SUMDebit = 0
SUMCrebit = 0
    With FrmAccEditJournal3.Fg_Journal

                       For i = .FixedRows To .rows - 1

                           If val(.TextMatrix(i, .ColIndex("BranchID"))) = BramchId Then
                               SUMDebit = SUMDebit + val(.TextMatrix(i, .ColIndex("DebitValue")))
                               SUMCrebit = SUMCrebit + val(.TextMatrix(i, .ColIndex("CreditValue")))
                           End If

                Next i
                     SUMDebit = Round(SUMDebit, SystemOptions.SysDefCurrencyForamt)
                 SUMCrebit = Round(SUMCrebit, SystemOptions.SysDefCurrencyForamt)


                    If val(SUMDebit) <> val(SUMCrebit) Then
                                               If SystemOptions.UserInterface = ArabicInterface Then
                                                       MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_name, vbCritical
                                               Else
                                                       MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_namee, vbCritical
                                               End If
                                                       CheckSusAccounts3 = False
                                               Exit Function

                               End If
    End With
    rsBranch.MoveNext
    Next Branch
    rsBranch.Close
    CheckSusAccounts3 = True


End Function

Function CheckSusAccounts4() As Boolean
 Dim branch_name    As String
 Dim branch_namee    As String
 Dim SUM As Double
 Dim i As Integer
 Dim SUMDebit  As Double
 Dim SUMCrebit As Double
 CheckSusAccounts4 = False
 Dim rsBranch As New ADODB.Recordset
 If SystemOptions.AllowUnbalncedByBranchAccount = True Then CheckSusAccounts4 = True: Exit Function
   My_SQL = "SELECT  * From TblBranchesData"


    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsBranch.RecordCount > 0 Then
        rsBranch.MoveFirst
    End If


 For Branch = 1 To rsBranch.RecordCount
BramchId = IIf(IsNull(rsBranch("branch_id").value), 0, rsBranch("branch_id").value)
branch_name = IIf(IsNull(rsBranch("branch_name").value), 0, rsBranch("branch_name").value)
branch_namee = IIf(IsNull(rsBranch("branch_namee").value), 0, rsBranch("branch_namee").value)
SUMDebit = 0
SUMCrebit = 0
    With FrmAccEditJournal4.Fg_Journal

                       For i = .FixedRows To .rows - 1

                           If val(.TextMatrix(i, .ColIndex("BranchID"))) = BramchId Then
                               SUMDebit = SUMDebit + val(.TextMatrix(i, .ColIndex("DebitValue")))
                               SUMCrebit = SUMCrebit + val(.TextMatrix(i, .ColIndex("CreditValue")))
                           End If

                   Next i
                  SUMDebit = Round(SUMDebit, SystemOptions.SysDefCurrencyForamt)
                 SUMCrebit = Round(SUMCrebit, SystemOptions.SysDefCurrencyForamt)

                    If val(SUMDebit) <> val(SUMCrebit) Then

                                               If SystemOptions.UserInterface = ArabicInterface Then
                                                       MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_name, vbCritical
                                               Else
                                                       MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_namee, vbCritical
                                               End If
                                                       CheckSusAccounts4 = False
                                               Exit Function

                               End If
    End With
    rsBranch.MoveNext
    Next Branch
    rsBranch.Close
    CheckSusAccounts4 = True


End Function

Public Function CheciIPOBySal_SharCount(Optional InvesID As Double = 0) As Boolean
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     TypeVaSh, InvesNo"
sql = sql & " From dbo.TblIPO"
sql = sql & " Where (TypeVaSh = 1) And (InvesNo = " & InvesID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheciIPOBySal_SharCount = True
Else
CheciIPOBySal_SharCount = False
End If
End Function

Function getStorenames(StoreID As Double, Optional StoreName As String, Optional storenamee As String)
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String

sql = "Select * from TblStore where StoreID=" & StoreID
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
        StoreName = IIf(IsNull(Rs8("StoreName").value), 0, Rs8("StoreName").value)
        storenamee = IIf(IsNull(Rs8("StoreNamee").value), 0, Rs8("StoreNamee").value)

Else
        StoreName = ""
        storenamee = ""

End If

End Function

 Public Sub GetLandInformation(Optional Land As Double = 0, Optional ByRef Area As Double)
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = "Select * from TblBuyLanReEst where ID=" & Land & ""
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
Area = IIf(IsNull(Rs8("Area").value), 0, Rs8("Area").value)
Else
Area = 0
End If
End Sub

Public Sub GetInvestInformation(Optional InvID As Double = 0, Optional ByRef InvesTotal As Double, Optional ByRef CountShare As Double, Optional ByRef ShareValue As Double)
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = "Select * from TblIPO where OrderInvse=" & InvID & ""
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
InvesTotal = IIf(IsNull(Rs8("InvesTotal").value), 0, Rs8("InvesTotal").value)
CountShare = IIf(IsNull(Rs8("CountShare").value), 0, Rs8("CountShare").value)
ShareValue = IIf(IsNull(Rs8("ShareValue").value), 0, Rs8("ShareValue").value)
Else
ShareValue = 0
InvesTotal = 0
CountShare = 0
End If

End Sub
'////////////////////////////////
Public Sub SavedTranInvest(Optional IDIPO As Double = 0, Optional BuyBilID As Double = 0, Optional des As String, Optional SharCount As Double = 0, Optional ShareValue As Double = 0 _
, Optional InvesID As Double = 0, Optional CusID As Double = 0, Optional Effict As Double = 0)
Dim StrSQL As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblTransactionInvest Where (1 = -1)"
    Rs5.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Rs5.AddNew
    Rs5("IDIPO").value = IDIPO
    Rs5("BuyBilID").value = BuyBilID
    Rs5("Des").value = des
    Rs5("SharCount").value = SharCount
    Rs5("ShareValue").value = ShareValue
    Rs5("InvesID").value = InvesID
    Rs5("CusID").value = CusID
    Rs5("Effict").value = Effict
Rs5.update
End Sub
''///////////////////////
   Public Function GetTotalSharOfCustomer(Optional CusID As Double = 0, Optional InvesID As Double = 0) As Double
    Dim Rs8 As ADODB.Recordset
    Set Rs8 = New ADODB.Recordset
    Dim sql As String
  sql = "   SELECT     SUM(SharCount * Effict) AS Totalshar, CusID, InvesID"
  sql = sql & "  From dbo.TblTransactionInvest"
  sql = sql & "  Where (CusID = " & CusID & ") And (InvesID = " & InvesID & ")"
  sql = sql & "  GROUP BY CusID, InvesID"
  Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs8.RecordCount > 0 Then
  GetTotalSharOfCustomer = IIf(IsNull(Rs8("Totalshar").value), 0, Rs8("Totalshar").value)
  Else
  GetTotalSharOfCustomer = 0
  End If
   End Function

Public Function GetFixedAssetAddAccount(FixedassetId As Double) As String
Dim GroupID As Integer
Dim FAADDAccount As String
Dim ParetnAccount As String
Dim GroupName As String
Dim Account_Code5 As String
    Dim StrSQL As String
    Dim X As String
    Dim Account_Code_dynamic As String
'GetAllDataAboutFixedAsset CInt(FixedassetId), , GroupID, , , , , , , , , , , , , , , , , , , , , , , , FAADDAccount, ParetnAccount, GroupName, GroupNamee
If FAADDAccount = "" Then



    If SystemOptions.AssetAccount = True Then
        X = ParetnAccount

      Account_Code5 = ModAccounts.AddNewAccount(X, " «÷«›«  " & GroupName, True, False, GroupNamee & "  Additions")

    Else
       Account_Code5 = ModAccounts.AddNewAccount(Account_Code_dynamic, " «÷«›«  " & GroupName, True, False, GroupNamee & "  Additions")
    End If


    StrSQL = "update FixedAssetsGroup  set  Account_Code5='" & Account_Code5 & "' where GroupID=" & GroupID
 Cn.Execute StrSQL
End If

GetFixedAssetAddAccount = Account_Code5
End Function

Public Function GetSalaryEmployee(Optional Emp_id As Integer = 0)
If Emp_id <> 0 Then
Dim sql As String
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
   sql = "select sum(DEV_Value1) as Total"
sql = sql & "  from("

sql = sql & " SELECT     dbo.EmpSalaryComponent.[Value] AS Total, dbo.mofrad.AddOrDiscount,"
sql = sql & " DEV_Value1=Case"
 sql = sql & " When AddOrDiscount=0   Then Value * 1"
 sql = sql & " Else  Value * -1"
 sql = sql & " End"
sql = sql & " FROM         dbo.mofrad INNER JOIN"
sql = sql & " dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type RIGHT OUTER JOIN"
sql = sql & " dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode"
'sql = sql & " Where (dbo.EmpSalaryComponent.Emp_id = 2)"
sql = sql & " WHERE     (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
     sql = sql & "   AND (dbo.mofrad.salary=1)"
     sql = sql & " )x"
Rs9.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
GetSalaryEmployee = IIf(IsNull(Rs9("Total").value), 0, Rs9("Total").value)
Else
GetSalaryEmployee = 0
End If
End If
End Function

Public Function CheckUnitContractMerg(unitno As Integer) As Boolean
Dim RsDetails As ADODB.Recordset
Dim StrSQL As String
    Set RsDetails = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.TblIqrMerg Where (UntID =" & unitno & ")"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   CheckUnitContractMerg = True
   Else
   CheckUnitContractMerg = False
   End If
End Function
Public Function ChekSanNumber(Optional branch_no As Integer = 0, Optional Sanad_No As Integer = 0) As Boolean
Dim str As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
ChekSanNumber = False
str = "SELECT     branch_no, sanad_no"
str = str & " From dbo.sanad_numbering"
str = str & " Where(Sanad_No = " & Sanad_No & ") And (branch_no = " & branch_no & ") and numbering_id <> 0"
Rs6.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
ChekSanNumber = True
Else
ChekSanNumber = False
End If
End Function

  Public Function Deletepost(ScreenName As String, tablename As String, FieldName As String, DepID As Integer, BranchID As Integer, Transaction_ID As Variant, NoteSerial As String)
Dim StrSQL As String
'StrSQL = "update " & tablename & " set  Posted=null ,PostedDate=null " & "where " & FieldName & "=" & Transaction_ID
'Cn.Execute StrSQL


StrSQL = "delete  ApprovalData  where Transaction_ID =" & Transaction_ID & "  and ScreenName='" & ScreenName & "'"
Cn.Execute StrSQL




 End Function
Public Function SendTopost(ScreenName As String, tablename As String, FieldName As String, DepID As Integer, BranchID As Integer, Transaction_ID As Variant, NoteSerial As String, Optional NoteID As Double = -1, Optional EmpDepartemenID As Integer, Optional OverProject As Double)

'user_id

Dim StrSQL As String
StrSQL = "update " & tablename & " set  Posted=" & user_id & ",PostedDate=" & SQLDate(Now, True) & "where " & FieldName & "=" & Transaction_ID
Cn.Execute StrSQL


 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
   'EmpDepartemenID = GetempDepartementidFromUserid(CInt(user_id))
    sql = "  select  dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID ,dbo.TbllevelWorker.EmpID1 , "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & ScreenName & "')"
sql = sql & " and     (dbo.TblApprovalDef.BranchId =" & BranchID & ")"
If EmpDepartemenID <> 0 Then
sql = sql & " and        dbo.TblApprovalDef.DepartmentID  =" & EmpDepartemenID
End If

sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

'Dim EmpDepartemenID As Integer
Dim UserID As Integer
Dim UserID1 As Integer
Dim UserID2 As Integer
Dim EmpID As Integer
         currentdate = Now

                        GetApprovalDepartement DepID, UserID, EmpID, BranchID, UserID1, UserID2
            Dim currcusor As Integer
            currcusor = 1
            If UserID <> 0 Then
           '***************************************
                                 RSApproval.AddNew
                        RSApproval("ScreenName").value = ScreenName
                        RSApproval("levelo").value = 0
                       RSApproval("EmpID").value = UserID
                       RSApproval("noteid").value = NoteID

                        RSApproval("levelorder").value = 0
                         RSApproval("currorder").value = 0
                          RSApproval("Transaction_ID").value = val(Transaction_ID)
                          RSApproval("overproject").value = val(OverProject)


                          RSApproval("NoteSerial").value = NoteSerial
                        RSApproval("Transaction_Date").value = Date

                          RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(ScreenName), currentdate)
                       RSApproval("SendTime").value = currentdate


                                RSApproval("Currcursor").value = currcusor
                                 RSApproval("FromUser").value = user_name


                        RSApproval.update
              End If



            If UserID1 <> 0 Then
           '***************************************
           currcusor = currcusor + 1
                                 RSApproval.AddNew
                                 RSApproval("overproject").value = val(OverProject)
                        RSApproval("ScreenName").value = ScreenName
                        RSApproval("levelo").value = 0
                       RSApproval("EmpID").value = UserID1
                        RSApproval("levelorder").value = 0
                         RSApproval("currorder").value = 0
                          RSApproval("Transaction_ID").value = Transaction_ID
                          RSApproval("NoteSerial").value = NoteSerial
                        RSApproval("Transaction_Date").value = Date

                          RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(ScreenName), currentdate)
                       RSApproval("SendTime").value = currentdate


                                RSApproval("Currcursor").value = currcusor
                                 RSApproval("FromUser").value = user_name

                        RSApproval("noteid").value = NoteID
                        RSApproval.update
              End If


           If UserID2 <> 0 Then
           '***************************************
           currcusor = currcusor + 1
                                 RSApproval.AddNew
                                 RSApproval("overproject").value = val(OverProject)
                        RSApproval("ScreenName").value = ScreenName
                        RSApproval("levelo").value = 0
                       RSApproval("EmpID").value = UserID2
                        RSApproval("levelorder").value = 0
                         RSApproval("currorder").value = 0
                          RSApproval("Transaction_ID").value = Transaction_ID
                          RSApproval("NoteSerial").value = NoteSerial
                        RSApproval("Transaction_Date").value = Date

                          RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(ScreenName), currentdate)
                       RSApproval("SendTime").value = currentdate


                                RSApproval("Currcursor").value = currcusor
                                 RSApproval("FromUser").value = user_name

                        RSApproval("noteid").value = NoteID
                        RSApproval.update
              End If
 Dim Flag As Integer
 Flag = 1
   Dim empID2 As Integer
    If Rs1.RecordCount > 0 Then




            For i = 1 To Rs1.RecordCount


           '****************************************
            empID2 = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
            If CheckWorkState(empID2) = True Then
            Else
            empID2 = IIf(IsNull(Rs1("empID1").value), 0, Rs1("empID1").value)
            End If
            If CheckWorkState(empID2) = True Then
              RSApproval.AddNew
              RSApproval("overproject").value = val(OverProject)
                RSApproval("ScreenName").value = ScreenName
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)

               RSApproval("EmpID").value = empID2
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = Transaction_ID
                  RSApproval("NoteSerial").value = NoteSerial
                RSApproval("Transaction_Date").value = Date

                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(ScreenName), currentdate)
               RSApproval("SendTime").value = currentdate

                 If Flag = 1 And UserID = 0 And UserID1 = 0 And UserID2 = 0 Then
                        RSApproval("Currcursor").value = 1
                       RSApproval("FromUser").value = user_name
                       Flag = 2
                End If
                RSApproval("noteid").value = NoteID
                RSApproval.update
              End If
                Rs1.MoveNext
            Next i

    End If

End Function




Public Function GetExchangReq(Optional ID As Double = 0, Optional ByRef YearID As Integer, Optional MonthID As Integer, Optional ByRef BranchID As Integer) As String
Dim sql As String
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
GetExchangReq = ""
If ID <> 0 Then
sql = " SELECT      *"
sql = sql & " From dbo.TblExchangeRequest"
sql = sql & " Where (id = " & ID & ")"
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
GetExchangReq = IIf(IsNull(Rs9("AllID").value), "", Rs9("AllID").value)
YearID = IIf(IsNull(Rs9("DurationID").value), 0, Rs9("DurationID").value)
MonthID = IIf(IsNull(Rs9("Month").value), 0, Rs9("Month").value)
BranchID = IIf(IsNull(Rs9("BranchID").value), 0, Rs9("BranchID").value)
Else
MonthID = 0
YearID = 0
GetExchangReq = ""
BranchID = 0
End If
End If
End Function

'Public FirstPeriodAll As Date
'Select Emp_ID,Transaction_ID from transactions  where  Transaction_Type=61
  Public Function GetTblBuyLandRealEstate(Optional ByRef ID As Integer, Optional ByRef fullcode As String, Optional Type1 As Integer = 0)
    Dim sql As String
    Dim rs As New ADODB.Recordset

    If Type1 = 0 Then
        sql = "select * from TblBuyLanReEst where ID= " & ID
    Else
    sql = "select * from TblBuyLanReEst where  FullCode ='" & fullcode & "'"
    End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then

        ID = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
        fullcode = IIf(IsNull(rs("FullCode").value), "", rs("FullCode").value)
    Else
        ID = 0
        fullcode = ""
    End If
    rs.Close
End Function


   Public Sub GetVocationEntitlements(ID As Integer, _
Optional BranchID As Integer, _
  Optional EmpID As Integer, _
                                Optional ByRef salary As Double, _
                                 Optional ByRef SalEntitOther As Double, _
                                 Optional ByRef other As Double, _
                                 Optional ByRef Advance As Double, _
                                 Optional ByRef ValueTickt As Double, _
                                 Optional ByRef SalaryVocation As Double, Optional ByRef InsuranceValue As Double, Optional PreSalary As Double)

     Dim StrSQL  As String
     Dim ch8 As Integer
     Dim ch6 As Boolean
     ch6 = False
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    StrSQL = "select * From tblVocationEntitlements     where id =" & ID & "  and not (NoteSerial is null) "

    If CheckAprroveScreen("FrmVocationEntitlements") = True Then
     StrSQL = StrSQL & " and approved =1"

End If



    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If (rs.RecordCount) > 0 Then
       BranchID = IIf(IsNull(rs("BranchID").value), Current_branch, (rs("BranchID").value))
      EmpID = IIf(IsNull(rs("EmpID").value), 0, (rs("EmpID").value))

       salary = Round(IIf(IsNull(rs("salary").value), 0, (rs("salary").value)), 2)
       SalEntitOther = Round(IIf(IsNull(rs("SalEntitOther").value), 0, (rs("SalEntitOther").value)), 2)
       other = Round(IIf(IsNull(rs("Other").value), 0, (rs("Other").value)), 2)
       Advance = Round(IIf(IsNull(rs("Advance").value), 0, (rs("Advance").value)), 2)
       ValueTickt = Round(IIf(IsNull(rs("ValueTickt").value), 0, (rs("ValueTickt").value)), 2)
       SalaryVocation = Round(IIf(IsNull(rs("SalaryVocation").value), 0, (rs("SalaryVocation").value)), 2)
       InsuranceValue = Round(IIf(IsNull(rs("InsuranceValue").value), 0, (rs("InsuranceValue").value)), 2)
       PreSalary = Round(IIf(IsNull(rs("PreSalary").value), 0, (rs("PreSalary").value)), 2)
       ch8 = Round(IIf(IsNull(rs("ch8").value), 0, (rs("ch8").value)), 2)
       ch6 = IIf(IsNull(rs("ch6").value), False, (rs("ch6").value))
       If ch6 = True Then
       salary = Round(IIf(IsNull(rs("salary").value), 0, (rs("salary").value)), 2) - Round(IIf(IsNull(rs("Decrease").value), 0, (rs("Decrease").value)), 2)
       End If
       If ch8 = 0 Then
       PreSalary = 0
       End If
         Else
         BranchID = 0
         EmpID = 0
         salary = 0
         SalEntitOther = 0
         other = 0
         ValueTickt = 0
         Advance = 0
    End If
    rs.Close
End Sub

 Public Sub GetEnd_Service(Optional ID As Double = 0, Optional ByRef BranchID As Integer, Optional ByRef EmpID As Double = 0, Optional ByRef total As Double = 0, Optional ByRef LastMonth As Double = 0, Optional ByRef Ticket As Double = 0, Optional ByRef Custom As Double = 0, Optional ByRef net As Double = 0, Optional ByRef TotalAdvance As Double = 0, Optional ByRef TxtVlueVaction As Double = 0, Optional ByRef TotalCash As Double, Optional ByRef LastTotal As Double = 0, Optional ByRef EndService As Double, Optional ByRef CusTiket As Double, Optional ByRef AddOther As Double, Optional ByRef DiffTekit As Double, Optional ByRef Discounts As Double, Optional ByRef TicktConract As Double, Optional ByRef DisSalary As Double)
If ID <> 0 Then
Dim sql As String
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
sql = "select * from End_of_service where id=" & ID & " and not (NoteSerial is null) "

 If CheckAprroveScreen("End_oF_service") = True Then
     StrWhere = StrWhere & " and approved =1"

End If


Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
EmpID = IIf(IsNull(Rs9("EmpID").value), 0, Rs9("EmpID").value)
total = Round(IIf(IsNull(Rs9("NetEnd").value), IIf(IsNull(Rs9("total").value), 0, Rs9("total").value), Rs9("NetEnd").value), 2)
LastMonth = Round(IIf(IsNull(Rs9("LastMonth").value), 0, Rs9("LastMonth").value), 2)
Ticket = Round(IIf(IsNull(Rs9("Ticket").value), 0, Rs9("Ticket").value), 2)
Custom = Round(IIf(IsNull(Rs9("Custom").value), 0, Rs9("Custom").value), 2)
net = Round(IIf(IsNull(Rs9("net").value), 0, Rs9("net").value), 2)
DiffTekit = Round(IIf(IsNull(Rs9("DiffTekit").value), 0, Rs9("DiffTekit").value), 2)
TotalAdvance = Round(IIf(IsNull(Rs9("TotalAdvance").value), 0, Rs9("TotalAdvance").value), 2)
TxtVlueVaction = Round(IIf(IsNull(Rs9("TxtVlueVaction").value), 0, Rs9("TxtVlueVaction").value), 2)
TotalCash = Round(IIf(IsNull(Rs9("TotalCash").value), 0, Rs9("TotalCash").value), 2)
LastTotal = Round(IIf(IsNull(Rs9("LastTotal").value), 0, Rs9("LastTotal").value), 2)
BranchID = Round(IIf(IsNull(Rs9("BranchID").value), 0, Rs9("BranchID").value), 2)
AddOther = Round(IIf(IsNull(Rs9("AddOther").value), 0, Rs9("AddOther").value), 2)
CusTiket = Round(IIf(IsNull(Rs9("CusTiket").value), 0, Rs9("CusTiket").value), 2)
EndService = Round(IIf(IsNull(Rs9("EndService").value), 0, Rs9("EndService").value), 2)
Discounts = Round(IIf(IsNull(Rs9("Discounts").value), 0, Rs9("Discounts").value), 2)
TicktConract = Round(IIf(IsNull(Rs9("TicktConract").value), 0, Rs9("TicktConract").value), 2)
DisSalary = Round(IIf(IsNull(Rs9("DisSalary").value), 0, Rs9("DisSalary").value), 2)
Else
TicktConract = 0
Discounts = 0
EndService = 0
CusTiket = 0
AddOther = 0
EmpID = 0
total = 0
LastMonth = 0
Ticket = 0
Custom = 0
net = 0
TotalAdvance = 0
TxtVlueVaction = 0
TotalCash = 0
LastTotal = 0
BranchID = 0
DiffTekit = 0
DisSalary = 0
End If
End If
End Sub

 Public Sub GetEnd_Servicex13082017(Optional ID As Double = 0, Optional ByRef BranchID As Integer, Optional ByRef EmpID As Double = 0, Optional ByRef total As Double = 0, Optional ByRef LastMonth As Double = 0, Optional ByRef Ticket As Double = 0, Optional ByRef Custom As Double = 0, Optional ByRef net As Double = 0, Optional ByRef TotalAdvance As Double = 0, Optional ByRef TxtVlueVaction As Double = 0, Optional ByRef TotalCash As Double, Optional ByRef LastTotal As Double = 0, Optional ByRef EndService As Double, Optional ByRef CusTiket As Double, Optional ByRef AddOther As Double, Optional ByRef DiffTekit As Double, Optional ByRef Discounts As Double, Optional ByRef TicktConract As Double)
If ID <> 0 Then
Dim sql As String
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
sql = "select * from End_of_service where id=" & ID & " "
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
EmpID = IIf(IsNull(Rs9("EmpID").value), 0, Rs9("EmpID").value)
total = Round(IIf(IsNull(Rs9("NetEnd").value), IIf(IsNull(Rs9("total").value), 0, Rs9("total").value), Rs9("NetEnd").value), 2)
LastMonth = Round(IIf(IsNull(Rs9("LastMonth").value), 0, Rs9("LastMonth").value))
Ticket = Round(IIf(IsNull(Rs9("Ticket").value), 0, Rs9("Ticket").value), 2)
Custom = Round(IIf(IsNull(Rs9("Custom").value), 0, Rs9("Custom").value), 2)
net = Round(IIf(IsNull(Rs9("net").value), 0, Rs9("net").value), 2)
DiffTekit = Round(IIf(IsNull(Rs9("DiffTekit").value), 0, Rs9("DiffTekit").value), 2)
TotalAdvance = Round(IIf(IsNull(Rs9("TotalAdvance").value), 0, Rs9("TotalAdvance").value), 2)
TxtVlueVaction = Round(IIf(IsNull(Rs9("TxtVlueVaction").value), 0, Rs9("TxtVlueVaction").value), 2)
TotalCash = Round(IIf(IsNull(Rs9("TotalCash").value), 0, Rs9("TotalCash").value), 2)
LastTotal = Round(IIf(IsNull(Rs9("LastTotal").value), 0, Rs9("LastTotal").value), 2)
BranchID = Round(IIf(IsNull(Rs9("BranchID").value), 0, Rs9("BranchID").value), 2)
AddOther = Round(IIf(IsNull(Rs9("AddOther").value), 0, Rs9("AddOther").value), 2)
CusTiket = Round(IIf(IsNull(Rs9("CusTiket").value), 0, Rs9("CusTiket").value), 2)
EndService = Round(IIf(IsNull(Rs9("EndService").value), 0, Rs9("EndService").value), 2)
Discounts = Round(IIf(IsNull(Rs9("Discounts").value), 0, Rs9("Discounts").value), 2)
TicktConract = Round(IIf(IsNull(Rs9("TicktConract").value), 0, Rs9("TicktConract").value), 2)
Else
TicktConract = 0
Discounts = 0
EndService = 0
CusTiket = 0
AddOther = 0
EmpID = 0
total = 0
LastMonth = 0
Ticket = 0
Custom = 0
net = 0
TotalAdvance = 0
TxtVlueVaction = 0
TotalCash = 0
LastTotal = 0
BranchID = 0
DiffTekit = 0
End If
End If
End Sub


 Public Function ChekTransferNo(ChqueNum As String, BankID As Double, NoteID As Double, ByRef NoteSerial1 As String) As Boolean
Dim Rs7 As ADODB.Recordset
Dim sql As String
ChekTransferNo = False
Set Rs7 = New ADODB.Recordset
sql = "SELECT     *"
sql = sql & " From notes"
sql = sql & "  WHERE     NoteCashingType=2"
sql = sql & "  and     NoteID<>" & NoteID
sql = sql & "  and     ChqueNum='" & ChqueNum & "'"
sql = sql & "  and     BankID=" & BankID


Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekTransferNo = True

NoteSerial1 = IIf(IsNull(Rs7("NoteSerial1").value), "", Rs7("NoteSerial1").value)
Else
ChekTransferNo = False
NoteSerial1 = ""
End If
End Function
Public Function GetIncoiceNoByOrder(order_no As String) As String
Dim Rs7 As ADODB.Recordset
Dim sql As String
'GetIncoiceNoByOrder = False
Set Rs7 = New ADODB.Recordset
sql = "SELECT     NoteSerial1"
sql = sql & "   From dbo.transactions"
sql = sql & "   Where (Transaction_Type = 21 And CBoBasedON = 2)"
sql = sql & "   AND (order_no = '" & order_no & "')"




Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then


GetIncoiceNoByOrder = IIf(IsNull(Rs7("NoteSerial1").value), "", Rs7("NoteSerial1").value)
Else
GetIncoiceNoByOrder = ""

End If
End Function

Public Function ChekClodePeriod(RecordDate As Date) As Boolean
Dim Rs7 As ADODB.Recordset
Dim sql As String
ChekClodePeriod = False
Set Rs7 = New ADODB.Recordset




sql = "SELECT     StartDate, EndDate"
sql = sql & " From dbo.TblAccountIntervals"
sql = sql & "  WHERE     (StartDate <=" & SQLDate(RecordDate, True) & " AND (EndDate >= " & SQLDate(RecordDate, True) & " ))and OpenState=1 "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekClodePeriod = True
Else
ChekClodePeriod = False
End If
End Function

Public Function ChekClodePeriodx(RecordDate As Date) As Boolean
Dim Rs7 As ADODB.Recordset
Dim sql As String
ChekClodePeriodx = False
Set Rs7 = New ADODB.Recordset
sql = "SELECT     StartDate, EndDate"
sql = sql & " From dbo.TblOpenClosPeriodDet1"
sql = sql & "  WHERE     (StartDate <=" & SQLDate(RecordDate, True) & " AND (EndDate >= " & SQLDate(RecordDate, True) & " ))"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekClodePeriodx = True
Else
ChekClodePeriodx = False
End If
End Function

 Public Function GetInsuranceAccount(Optional ByRef Acount_Code1 As String, _
  Optional ByRef Acount_Code2 As String, Optional ByRef CitizenVal1 As Double, _
  Optional ByRef ResidentVal1 As Double, Optional ByRef Acount_Code4 As String, _
  Optional ByRef Acount_Code3 As String, Optional ByRef CitizenVal2 As Double, _
  Optional ByRef ResidentVal2 As Double)
Dim My_SQL As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset

 My_SQL = " SELECT    * from  TblSocialInsurance"

 Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
 If Rs7.RecordCount > 0 Then

Acount_Code1 = IIf(IsNull(Rs7("Acount_Code1").value), "", Rs7("Acount_Code1").value)
 Acount_Code2 = IIf(IsNull(Rs7("Acount_Code2").value), "", Rs7("Acount_Code2").value)

 Acount_Code3 = IIf(IsNull(Rs7("Acount_Code3").value), "", Rs7("Acount_Code3").value)
 Acount_Code4 = IIf(IsNull(Rs7("Acount_Code4").value), "", Rs7("Acount_Code4").value)

  CitizenVal1 = IIf(IsNull(Rs7("CitizenVal1").value), 0, Rs7("CitizenVal1").value)
  ResidentVal1 = IIf(IsNull(Rs7("ResidentVal1").value), 0, Rs7("ResidentVal1").value)
  CitizenVal2 = IIf(IsNull(Rs7("CitizenVal2").value), 0, Rs7("CitizenVal2").value)
  ResidentVal2 = IIf(IsNull(Rs7("ResidentVal2").value), 0, Rs7("ResidentVal2"))

 Else
Acount_Code1 = ""
Acount_Code2 = ""
Acount_Code3 = ""
Acount_Code4 = ""
CitizenVal1 = 0
ResidentVal1 = 0
CitizenVal2 = 0
ResidentVal2 = 0

 End If



End Function

 Public Function GetCartData(card As String, Optional Name As String = "") As Integer
Dim My_SQL As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset


 My_SQL = " SELECT     name, namee, tel, card, discount, CartTYpe"
My_SQL = My_SQL & "   From dbo.TblCusCsh"
My_SQL = My_SQL & "   WHERE     (card =  '" & card & "')"



 Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
 If Rs7.RecordCount > 0 Then

 If SystemOptions.UserInterface = ArabicInterface Then
Name = IIf(IsNull(Rs7("name").value), "", Rs7("name").value)
 Else
 Name = IIf(IsNull(Rs7("namee").value), "", Rs7("namee").value)
 End If

 Else
Name = ""
 End If

End Function
 Public Function GetApprovalDepartement(DeparmentID As Integer, Optional ByRef UserID As Integer, Optional ByRef EmpID As Integer, Optional ByRef BranchID As Integer, Optional ByRef UserID1 As Integer, Optional ByRef UserID2 As Integer) As Integer
Dim My_SQL As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset


 My_SQL = " SELECT  dbo.TblEmpDepartments.UserId1,dbo.TblEmpDepartments.UserId2,   dbo.TblEmpDepartments.UserId, dbo.TblEmpDepartments.DeparmentID, dbo.TblUsers.Empid"
My_SQL = My_SQL & " FROM         dbo.TblEmpDepartments INNER JOIN"
My_SQL = My_SQL & "                       dbo.TblUsers ON dbo.TblEmpDepartments.UserId = dbo.TblUsers.UserID"
My_SQL = My_SQL & "  Where (dbo.TblEmpDepartments.DeparmentID = " & DeparmentID & ")"
'My_SQL = My_SQL & "  and (dbo.TblEmpDepartments.BranchId = " & BranchID & ")"



 Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
 If Rs7.RecordCount > 0 Then
 UserID = IIf(IsNull(Rs7("UserId").value), 0, Rs7("UserId").value)
 UserID1 = IIf(IsNull(Rs7("UserId1").value), 0, Rs7("UserId1").value)
 UserID2 = IIf(IsNull(Rs7("UserId2").value), 0, Rs7("UserId2").value)

 EmpID = IIf(IsNull(Rs7("Empid").value), 0, Rs7("Empid").value)
 Else
UserID = 0
UserID1 = 0
UserID2 = 0

 EmpID = 0
 End If

End Function

 Public Function GetEmpIdfromProduction(Transaction_ID As Integer) As Integer
Dim My_SQL As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset


 My_SQL = "SELECT     Emp_id From dbo.transactions WHERE     (Transaction_ID = " & Transaction_ID & ")"
 Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
 If Rs7.RecordCount > 0 Then
 GetEmpIdfromProduction = IIf(IsNull(Rs7("Emp_id").value), "", Rs7("Emp_id").value)
 Else
 GetEmpIdfromProduction = 0
 End If

End Function



Public Function GetWaiterForTable(ID As Integer) As Integer
Dim My_SQL As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset


 My_SQL = "SELECT     Emp_id From dbo.Stables WHERE     (id = " & ID & ")"
 Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
 If Rs7.RecordCount > 0 Then
 GetWaiterForTable = IIf(IsNull(Rs7("Emp_id").value), 0, Rs7("Emp_id").value)
 Else
 GetWaiterForTable = 0
 End If

End Function


Public Function getAccountSerial_Code(Optional filed As String, Optional FiledWher As String, Optional str As String) As String
Dim My_SQL As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If " & Filed &" <> "" Then
 My_SQL = "  select " & filed & " as Acoud from ACCOUNTS where " & FiledWher & "='" & str & "'"
 Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
 If Rs7.RecordCount > 0 Then
 getAccountSerial_Code = IIf(IsNull(Rs7("Acoud").value), "", Rs7("Acoud").value)
 Else
 getAccountSerial_Code = ""
 End If
 End If
End Function
Public Function CheckCartDiscount(value As Double) As Double
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    If value = 0 Then CheckCartDiscount = 0: Exit Function

    StrSQL = "SELECT     Lowsalary  , HighSalary, AdvValue"
StrSQL = StrSQL + "  From dbo.TblCustomerPoints"
StrSQL = StrSQL + " Where (Lowsalary <= " & value & ") And (HighSalary >= " & value & ")"


    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
         CheckCartDiscount = IIf(IsNull(rs("AdvValue").value), 0, rs("AdvValue").value)


    Else
         CheckCartDiscount = 0

    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function CheckSpecialOffer(overdate As Date, Optional ByRef brnchid As Double = -1, Optional ByRef Sales As Double = -1, _
  Optional ByRef GetFree As Double = -1, Optional ByRef discount As Double = -1, Optional ByRef FromPrice As Double = -1) As Boolean
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = " SELECT     dbo.TblItemShowDitailses.BrnchID, dbo.TblItemShowDitailses.Type, dbo.TblItemShows.StartSDate, dbo.TblItemShows.EndDate, dbo.TblItemShows.Sales, "
StrSQL = StrSQL + "                        dbo.TblItemShows.GetFree , dbo.TblItemShows.Discount, dbo.TblItemShows.FromPrice"
StrSQL = StrSQL + "  FROM         dbo.TblItemShowDitailses INNER JOIN"
StrSQL = StrSQL + "                        dbo.TblItemShows ON dbo.TblItemShowDitailses.ID2 = dbo.TblItemShows.ID"
StrSQL = StrSQL + "  Where (dbo.TblItemShowDitailses.type = 1) And (dbo.TblItemShowDitailses.BrnchID = " & brnchid & ") And (Not (dbo.TblItemShows.Sales Is Null))"

 StrSQL = StrSQL + "  and     (" & SQLDate(overdate, True) & "BETWEEN dbo.TblItemShows.StartSDate AND dbo.TblItemShows.EndDate) "
   StrSQL = StrSQL + "  AND (dbo.TblItemShows.TypePoliceP = 4) "
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
         Sales = IIf(IsNull(rs("Sales").value), -1, rs("Sales").value)
         GetFree = IIf(IsNull(rs("GetFree").value), -1, rs("GetFree").value)
         discount = IIf(IsNull(rs("Discount").value), -1, rs("Discount").value)
         FromPrice = IIf(IsNull(rs("FromPrice").value), -1, rs("FromPrice").value) ' 0 min   a max

         CheckSpecialOffer = True
    Else
         CheckSpecialOffer = False

    End If

    rs.Close
    Set rs = Nothing
End Function
Public Function CheckoverInbranch(Optional ID2 As Double = -1, Optional BranchID As Double = -1) As Boolean
    Dim rs As New ADODB.Recordset
   Dim StrSQL As String

StrSQL = " SELECT     ID2, BrnchID, Type"
 StrSQL = StrSQL + "    From dbo.TblItemShowDitailses"
 StrSQL = StrSQL + "    Where (ID2 = " & ID2 & ") And (brnchid = " & BranchID & ")"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then

         CheckoverInbranch = True
    Else
         CheckoverInbranch = False

    End If

End Function
Public Function CheckItem(ItemID As Long, _
                          overdate As Date, _
                          Optional ByRef typedisid As Double = -1, _
                          Optional ByRef discount As Double = -1, _
                          Optional TypePoliceP As Double = -1, _
                          Optional BranchID As Double = -1, _
                          Optional freeitemid As Long, _
                          Optional freeitemUnitid As Long, _
                          Optional ByRef Amount As Long, _
                          Optional freeitemQty As Double, _
                          Optional qtyPercentage As Double) As Boolean
    Dim rs          As New ADODB.Recordset
    Dim ID2         As Double
    Dim StrSQL      As String
    Dim CurrWeekday As Integer
    StrSQL = "SELECT  amount ,InfITemSho, "
    StrSQL = StrSQL + "     TblItemShows.Sa ,"
    StrSQL = StrSQL + "     TblItemShows.Su,"
    StrSQL = StrSQL + "     TblItemShows.Mo,"
    StrSQL = StrSQL + "     TblItemShows.Tu ,  "
    StrSQL = StrSQL + "     TblItemShows.We,"
    StrSQL = StrSQL + "     TblItemShows.Th,"
    StrSQL = StrSQL + "      TblItemShows.Fr,"
    StrSQL = StrSQL + "     dbo.TblItemShows.StartSDate , dbo.TblItemShows.EndDate, dbo.TblItemShowDitailses.ItemID, dbo.TblItemShowDitailses.discount, "
    StrSQL = StrSQL + "                       dbo.TblItemShowDitailses.uniteid , dbo.TblItemShowDitailses.typedisid, dbo.TblItemShows.id"
    StrSQL = StrSQL + " , dbo.TblItemShowDitailses.ID2 FROM         dbo.TblItemShows LEFT OUTER JOIN"
    StrSQL = StrSQL + "                       dbo.TblItemShowDitailses ON dbo.TblItemShows.ID = dbo.TblItemShowDitailses.ID2"
    StrSQL = StrSQL + "  WHERE     (" & SQLDate(overdate, True) & "BETWEEN dbo.TblItemShows.StartSDate AND dbo.TblItemShows.EndDate) AND (NOT (dbo.TblItemShowDitailses.ItemID IS NULL))"
    StrSQL = StrSQL + "  and TblItemShowDitailses.ItemID=" & ItemID
    If TypePoliceP = 4 Then
        StrSQL = StrSQL + "  and TblItemShows.TypePoliceP=" & TypePoliceP
    End If
    CurrWeekday = Weekday(overdate, vbUseSystemDayOfWeek)
  
    Dim strsql2 As String
    strsql2 = "Select * from TblItemShowDitailses where Type = 3"
    rs.Open strsql2, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        strsql2 = "Select * from TblItemShowDitailses where Type = 3 and BrnchID = " & PPointID
        rs.Close
        rs.Open strsql2, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If rs.EOF Then
            CheckItem = False
            Exit Function
            '    Else
            '        rs.Close
        End If
    End If
   
    rs.Close
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim Sa, Su, Mo, Tu, We, Th, Fr As Integer
    If Not (rs.BOF Or rs.EOF) Then

        Sa = IIf(IsNull(rs("Sa").value) Or rs("Sa").value = 0, 0, 1)
        Su = IIf(IsNull(rs("Su").value) Or rs("Su").value = 0, 0, 2)
        Mo = IIf(IsNull(rs("Mo").value) Or rs("Mo").value = 0, 0, 3)
        Tu = IIf(IsNull(rs("Tu").value) Or rs("Tu").value = 0, 0, 4)
        We = IIf(IsNull(rs("We").value) Or rs("We").value = 0, 0, 5)
        Th = IIf(IsNull(rs("Th").value) Or rs("Th").value = 0, 0, 6)
        Fr = IIf(IsNull(rs("Fr").value) Or rs("Fr").value = 0, 0, 7)
        If CurrWeekday = Sa _
           Or CurrWeekday = Su _
           Or CurrWeekday = Mo _
           Or CurrWeekday = Tu _
           Or CurrWeekday = We _
           Or CurrWeekday = Th _
           Or CurrWeekday = Fr _
           Or CurrWeekday = Sa _
           Then
            '***********************************************************************
            InfITemSho = IIf(IsNull(rs("InfITemSho").value), -1, rs("InfITemSho").value)
            typedisid = IIf(IsNull(rs("typedisid").value), -1, rs("typedisid").value)
            discount = IIf(IsNull(rs("discount").value), -1, rs("discount").value)
            AmountQTY = IIf(IsNull(rs("amount").value), 0, rs("amount").value)
            If AmountQTY = 0 Then
                AmountQTY = 1
            End If
            Amount = AmountQTY
            '     Dim qtyPercentage As Integer
            If InfITemSho <> "" Then
                VarSet = Split(InfITemSho, "#", , vbTextCompare)

                If VarSet(0) <> Empty Or VarSet(0) <> "" Then
                    freeitemid = VarSet(0)
                    freeitemUnitid = VarSet(1)

                    freeitemQty = VarSet(2)

                    qtyPercentage = freeitemQty / AmountQTY
                Else
                    qtyPercentage = 0
                End If

            End If
            
            ID2 = IIf(IsNull(rs("ID2").value), -1, rs("ID2").value)

            If CheckoverInbranch(ID2, BranchID) = True Then
                CheckItem = True
            Else
                CheckItem = False
            End If

            '***********************************************************
        Else
            CheckItem = False

        End If

    Else
        CheckItem = False

    End If

    rs.Close
    Set rs = Nothing
End Function
Public Function daysInMonth(rdate As Date) As Long
Dim yr As Long
Dim mnth As Long
yr = year(rdate)
mnth = Month(rdate)
  ' Return the number of days in the specified month.
  daysInMonth = day(DateSerial(yr, mnth + 1, 1) - 1)
 ' daysInMonth = day(DateSerial(yr, mnth + 1, 1))
End Function

Public Function get__Account(ID As Integer, _
                                  filed As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from TblDoCumentsTypes where id=" & ID

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount = 0 Then get__Account = "": Exit Function
    If IsNull(Rs3(filed).value) Then get__Account = "": Exit Function
    If Not IsNull(Rs3(filed).value) Then get__Account = Rs3(filed).value: Exit Function
    Rs3.Close

End Function

Public Function CheckItemSpecialOffer(ItemID As Long, overdate As Date, Optional ByRef typedisid As Double = -1, _
Optional ByRef Sales As Double = -1, Optional GetFree As Double = -1, _
Optional discount As Double = -1, Optional FromPrice As Double = -1, Optional TypePoliceP As Double = -1 _
, Optional BranchID As Double = -1) As Boolean
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ID2 As Double
    StrSQL = "SELECT     dbo.TblItemShows.StartSDate, dbo.TblItemShows.EndDate, dbo.TblItemShowDitailses.ItemID, dbo.TblItemShows.discount, "
 StrSQL = StrSQL + "                       dbo.TblItemShows.Sales , dbo.TblItemShows.GetFree, dbo.TblItemShows.FromPrice"
 StrSQL = StrSQL + "  , dbo.TblItemShowDitailses.ID2 FROM         dbo.TblItemShows LEFT OUTER JOIN"
 StrSQL = StrSQL + "                       dbo.TblItemShowDitailses ON dbo.TblItemShows.ID = dbo.TblItemShowDitailses.ID2"
 StrSQL = StrSQL + "  WHERE     (" & SQLDate(overdate, True) & "BETWEEN dbo.TblItemShows.StartSDate AND dbo.TblItemShows.EndDate) AND (NOT (dbo.TblItemShowDitailses.ItemID IS NULL))"
  StrSQL = StrSQL + "  and TblItemShowDitailses.ItemID=" & ItemID
  If TypePoliceP = 4 Then
  StrSQL = StrSQL + "  and TblItemShows.TypePoliceP=" & TypePoliceP
  End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
    ID2 = IIf(IsNull(rs("ID2").value), -1, rs("ID2").value)

         Sales = IIf(IsNull(rs("Sales").value), -1, rs("Sales").value)
         GetFree = IIf(IsNull(rs("GetFree").value), -1, rs("GetFree").value)
                  discount = IIf(IsNull(rs("Discount").value), -1, rs("Discount").value)
                  FromPrice = IIf(IsNull(rs("FromPrice").value), -1, rs("FromPrice").value)

           If CheckoverInbranch(ID2, BranchID) = True Then
                CheckItemSpecialOffer = True
         Else
                  CheckItemSpecialOffer = False
         End If


        ' CheckItemSpecialOffer = True
    Else
         CheckItemSpecialOffer = False

    End If

    rs.Close
    Set rs = Nothing
End Function



 Public Function GetbalanceBar(AccountCode As String) As String
Dim Balance As String
Dim balanceString As String
Dim account_name As String
 Dim Account_NameEng As String
Dim account_serial As String
Dim str As String

    WriteCustomerBalPublic AccountCode, Balance, balanceString, , , account_name, Account_NameEng, account_serial
  If SystemOptions.UserInterface = ArabicInterface Then
    str = "ﬂÊœ «·Õ”««» : " & account_serial & CHR(13)
    str = str & "«”„ «·Õ”««» : " & account_name & CHR(13)
    str = str & "—’Ìœ «·Õ”««» : " & balanceString & CHR(13)

Else

    str = "Account Code: " & account_serial & CHR(13)
    str = str & "Account Name: " & Account_NameEng & CHR(13)
    str = str & "Balance: " & balanceString & CHR(13)

End If
    GetbalanceBar = str
End Function

Public Function GetFixedIDFromCode(Optional code As String, Optional ByRef FixedID As Integer)

    Dim sql As String
    Dim rs As New ADODB.Recordset


        sql = "select * from FixedAssets where  code ='" & code & "'"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        FixedID = IIf(IsNull(rs("id").value), 0, rs("id").value)

    Else
        FixedID = 0
    End If

    rs.Close

End Function

Public Function Reload(frmIn As Form)
Unload frmIn
Load frmIn
frmIn.show
End Function


Function GETLINKSQL(StoreName As Integer, Optional myindex As Integer = 0, Optional updatestate As String, Optional ByRef groupcodes As String) As String


Dim GROUPSTR As String

If StoreName = 0 Then Exit Function

If updatestate = "" And myindex = 0 Then

 StrSQL = "Select * From TblItems  where 1=0"
GoTo ll
End If



If myindex = 0 Then
        If updatestate <> "E" And updatestate <> "N" Then

         StrSQL = "Select * From TblItems  where IsArchive=0  "
        GoTo ll
        End If

End If

GROUPSTR = " (SELECT     dbo.TblLink_Item_To_Store_Details3.GroupID"
GROUPSTR = GROUPSTR + " FROM         dbo.TblLink_Item_To_StoreH INNER JOIN"
GROUPSTR = GROUPSTR + " dbo.TblLink_Item_To_Store_Details1 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details1.Ind INNER JOIN"
GROUPSTR = GROUPSTR + " dbo.TblLink_Item_To_Store_Details2 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details2.Ind INNER JOIN"
GROUPSTR = GROUPSTR + " dbo.TblLink_Item_To_Store_Details3 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details3.Ind"
GROUPSTR = GROUPSTR + " Where (dbo.TblLink_Item_To_Store_Details1.StoreId = " & val(StoreName) & ")"
GROUPSTR = GROUPSTR + " GROUP BY dbo.TblLink_Item_To_Store_Details3.GroupID)"

If myindex = 0 Then
getallgroupsdata GROUPSTR, groupcodes, updatestate
Strforitems = groupcodes

Else

End If




If myindex = 0 Then
StrSQL = " SELECT    distinct ItemID, ItemCode, ItemName, GroupID, HaveSerial, LastUpdate, UserID, PurchasePrice, SallingPrice, RequestLimit, CustomerPrice, DealerPrice, HaveGuarantee, "

StrSQL = StrSQL + " GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, ItemCase, ItemMaking, ItemMakingNew, code, Branch_NO, Fullcode,"
StrSQL = StrSQL + " prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, itemSerials, BinLocation, minvalueqty, MaxValueqty, FreeQty, barCodeNO, CatlogNO, FactoryNO, TemplateID,"
StrSQL = StrSQL + " ItemMaxDiscount, OverHead, Wight, Content, Dippre, Source, Typenew,ItemWithOutVAT"

ElseIf myindex = 1 Then
StrSQL = "SELECT   distinct  ItemID, barCodeNO"
ElseIf myindex = 2 Then
If SystemOptions.UserInterface = ArabicInterface Then
StrSQL = "SELECT   distinct  ItemID, ItemName"
Else
StrSQL = "SELECT   distinct  ItemID, ItemNamee"
End If
End If


                 StrSQL = StrSQL + " From dbo.TblItems"
                 StrSQL = StrSQL + " where  IsArchive =0 and GroupID in ("


                 StrSQL = StrSQL + " select GroupID from fullgroups () )"



            '    StrSQL = StrSQL + " or itemid in("
            '    StrSQL = StrSQL + " SELECT     dbo.TblLink_Item_To_Store_Details2.ItemID"
            '    StrSQL = StrSQL + " FROM         dbo.TblLink_Item_To_StoreH INNER JOIN"
            '    StrSQL = StrSQL + " dbo.TblLink_Item_To_Store_Details1 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details1.Ind INNER JOIN"
            '    StrSQL = StrSQL + " dbo.TblLink_Item_To_Store_Details2 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details2.Ind INNER JOIN"
            '    StrSQL = StrSQL + " dbo.TblLink_Item_To_Store_Details3 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details3.Ind"
            '    StrSQL = StrSQL + " Where (dbo.TblLink_Item_To_Store_Details1.StoreId = " & val(STORENAME) & ")"
            '    StrSQL = StrSQL + " GROUP BY dbo.TblLink_Item_To_Store_Details2.ItemID"
            '    StrSQL = StrSQL + " )"
ll:
GETLINKSQL = StrSQL

End Function


Function GETLINKSQLByActivity(XXXX As Integer, Optional myindex As Integer = 0, Optional updatestate As String, Optional ByRef groupcodes As String) As String


Dim GROUPSTR As String

If user_id = 0 Then Exit Function





 GROUPSTR = " SELECT     dbo.Groups.GroupID"
GROUPSTR = GROUPSTR + " FROM         dbo.Groups "
GROUPSTR = GROUPSTR + " Where     dbo.Groups.ActivityTypeId in (  "
 GROUPSTR = GROUPSTR & "  SELECT     dbo.TblBranchesData.ActivityTypeId"
GROUPSTR = GROUPSTR & "   FROM         dbo.TblUsersBranches INNER JOIN"
GROUPSTR = GROUPSTR & "                        dbo.TblBranchesData ON dbo.TblUsersBranches.BranchID = dbo.TblBranchesData.branch_id"
GROUPSTR = GROUPSTR & "   Where (dbo.TblUsersBranches.UserID = " & user_id & ")"
GROUPSTR = GROUPSTR & " ) "



If 0 = 0 Then
getallgroupsdata GROUPSTR, groupcodes, "xx"
Strforitems = groupcodes
End If




If myindex = 0 Then 'grid
StrSQL = " SELECT    distinct ItemID, ItemCode, ItemName, GroupID, HaveSerial, LastUpdate, UserID, PurchasePrice, SallingPrice, RequestLimit, CustomerPrice, DealerPrice, HaveGuarantee, "

StrSQL = StrSQL + " GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, ItemCase, ItemMaking, ItemMakingNew, code, Branch_NO, Fullcode,"
StrSQL = StrSQL + " prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, itemSerials, BinLocation, minvalueqty, MaxValueqty, FreeQty, barCodeNO, CatlogNO, FactoryNO, TemplateID,"
StrSQL = StrSQL + " ItemMaxDiscount, OverHead, Wight, Content, Dippre, Source, Typenew,ItemWithOutVAT"

ElseIf myindex = 1 Then 'item code combo
StrSQL = "SELECT   distinct  ItemID, barCodeNO"
ElseIf myindex = 2 Then
If SystemOptions.UserInterface = ArabicInterface Then 'ItemName  code combo
StrSQL = "SELECT   distinct  ItemID, ItemName"
Else
StrSQL = "SELECT   distinct  ItemID, ItemNamee"
End If
End If


                 StrSQL = StrSQL + " From dbo.TblItems"
                 StrSQL = StrSQL + " where  IsArchive =0 and GroupID in ("


                 StrSQL = StrSQL + " select GroupID from fullgroups () )"



GETLINKSQLByActivity = StrSQL

End Function

Function getallgroupsdata(Optional strIngroups As String = "", Optional ByRef groupcodes As String, Optional ByRef updateStatus As String)
Dim sql As String
'Dim groupcodes As String
 On Error Resume Next
 GoTo ll
 'newwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
      Dim StrSQL  As String
     Dim rs As ADODB.Recordset
     Dim fullcode As String
    Set rs = New ADODB.Recordset
    StrSQL = "select     Fullcode  FROM         dbo.Groups WHERE     groupID IN (" & strIngroups & ")"


  groupcodes = "SELECT     Fullcode From dbo.TblItems WHERE     (Fullcode LIKE N'0') "

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    'rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.RecordCount) > 0 Then
       rs.MoveFirst
       'OR (Fullcode LIKE N'1001%')

       For i = 0 To rs.RecordCount
       fullcode = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
       groupcodes = groupcodes & " OR (Fullcode LIKE N'" & fullcode & "')"
       rs.MoveNext
       Next i
   groupcodesPublic = groupcodes
     '  groupCodes

    End If

    rs.Close

'newwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
ll:
 'Exit Function
 If updateStatus = "R" Or updateStatus = "" Then
  Exit Function
 End If

sql = " drop  FUNCTION fullgroups"
Cn.Execute sql

 sql = "CREATE FUNCTION fullgroups ()"
 sql = sql & " RETURNS @xTable TABLE"
sql = sql & " ("
sql = sql & " groupid INT,"
sql = sql & " parentidd INT,"
sql = sql & " Iteration INT"
sql = sql & " )"
 sql = sql & " AS"
  sql = sql & " Begin"
  sql = sql & " DECLARE @rowsAdded INT"
  sql = sql & " DECLARE @Iteration INT;"
  sql = sql & " DECLARE @MaxRecursion INT ;   "
  sql = sql & " set @Iteration=1;"
  sql = sql & " set @MaxRecursion=10000;"
sql = sql & "  INSERT @xTable"

sql = sql & " SELECT groupid, parentid, @Iteration"
sql = sql & " From Groups"
sql = sql & " WHERE  groupid in  (  " & strIngroups & " )"
sql = sql & " SET @rowsAdded=@@rowcount ;"
sql = sql & " WHILE @rowsAdded > 0 AND @Iteration <= @MaxRecursion BEGIN"
 sql = sql & "   INSERT @xTable"
 sql = sql & "   SELECT e.groupid, parentid, @Iteration + 1"
sql = sql & " FROM groups e"
sql = sql & "    INNER JOIN @xTable r ON e.parentid = r.groupid"
sql = sql & " WHERE   parentid <> e.groupid AND r.Iteration = @Iteration;"
sql = sql & " SET @rowsAdded=@@rowcount;"
sql = sql & "    SET @Iteration = @Iteration + 1;"
sql = sql & " End"
sql = sql & " Return"
sql = sql & " End"



db_createOrUpdateFuctionSQL "fullgroups", sql


End Function

 Function CreateRecusiveGroup(strIngroups As String, UserID As Double)
Dim sql As String
On Error Resume Next
sql = " drop  FUNCTION fullgroups" & UserID
Cn.Execute sql

 sql = "CREATE FUNCTION fullgroups" & UserID & " ()"
 sql = sql & " RETURNS @xTable TABLE"
sql = sql & " ("
sql = sql & " groupid INT,"
sql = sql & " parentidd INT,"
sql = sql & " Iteration INT"
sql = sql & " )"
 sql = sql & " AS"
  sql = sql & " Begin"
  sql = sql & " DECLARE @rowsAdded INT"
  sql = sql & " DECLARE @Iteration INT;"
  sql = sql & " DECLARE @MaxRecursion INT ;   "
  sql = sql & " set @Iteration=1;"
  sql = sql & " set @MaxRecursion=10000;"
sql = sql & "  INSERT @xTable"

sql = sql & " SELECT groupid, parentid, @Iteration"
sql = sql & " From Groups"
sql = sql & " WHERE  groupid in  (  " & strIngroups & " )"
sql = sql & " SET @rowsAdded=@@rowcount ;"
sql = sql & " WHILE @rowsAdded > 0 AND @Iteration <= @MaxRecursion BEGIN"
 sql = sql & "   INSERT @xTable"
 sql = sql & "   SELECT e.groupid, parentid, @Iteration + 1"
sql = sql & " FROM groups e"
sql = sql & "    INNER JOIN @xTable r ON e.parentid = r.groupid"
sql = sql & " WHERE   parentid <> e.groupid AND r.Iteration = @Iteration;"
sql = sql & " SET @rowsAdded=@@rowcount;"
sql = sql & "    SET @Iteration = @Iteration + 1;"
sql = sql & " End"
sql = sql & " Return"
sql = sql & " End"



 db_createOrUpdateFuctionSQL "fullgroups" & UserID, sql


 End Function
Public Function checkRentAccount(Account_code As String) As Boolean
    Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
 StrSQL = "  SELECT       * "
StrSQL = StrSQL & " From dbo.ExpensesType"
StrSQL = StrSQL & " WHERE  Transportation=1 and   Account_Code='" & Account_code & "'"

checkRentAccount = False

  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If (rs.RecordCount) > 0 Then

             checkRentAccount = True
             Else
             checkRentAccount = False

    End If



 End Function
 Public Function checkmanyApproval(frmname As String)
    Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
 StrSQL = "  SELECT     TOP 100 PERCENT ScreenName, ApprovName, ApprovNamee"
StrSQL = StrSQL & " From dbo.TblApprovalDef"
StrSQL = StrSQL & " WHERE     (ScreenName = N'" & frmname & "')"

checkmanyApproval = False

  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If (rs.RecordCount) > 0 Then

                If (rs.RecordCount) > 1 Then
                    checkmanyApproval = True

                End If

    End If



 End Function
Public Function FillApprovedTableNew(ScreenName As String, Transaction_ID As Double, NoteSerial1 As String, ID As Integer)
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.id =" & ID & ")"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            currentdate = Now
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = ScreenName
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = Transaction_ID
                  RSApproval("NoteSerial").value = NoteSerial1
                RSApproval("Transaction_Date").value = Date

                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(ScreenName), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If

                RSApproval.update
                Rs1.MoveNext
            Next i

    End If



End Function
Public Function GetItemPriceByWitdth(Item_ID As Long, Width As Double, Optional ByVal LngUnitID As Long = 0) As Double
   'Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
 StrSQL = "SELECT     dbo.Fn_GetPriceItem(" & Item_ID & ", " & Width & ") AS WidthPrice  "
StrSQL = StrSQL & " From dbo.TblItems"
StrSQL = StrSQL & "  Where (IsPriceIsPerview =1 and ItemID = " & Item_ID & ")"



  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If (rs.RecordCount) > 0 Then
GetItemPriceByWitdth = IIf(IsNull(rs("WidthPrice").value), 0, (rs("WidthPrice").value))
Else
GetItemPriceByWitdth = 0
    End If
    If GetItemPriceByWitdth = 0 Then GetItemPriceByWitdth = GetItemPrice(Item_ID, , LngUnitID)
End Function



Public Function checkdataexist(StrSQL As String) As Boolean
   'Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

checkdataexist = False

  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If (rs.RecordCount) > 0 Then
checkdataexist = True
    End If
End Function









 Function getpricebycustomrContract(customerid As Double, UnitID As Long, ItemID As Long, Optional vendor As Integer = 0, Optional ByVal mSalesMan As Integer, Optional mCashCust As String = "", Optional Transaction_Date As Date) As Double
    Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim total As Double
    If vendor = 0 Then
 StrSQL = "SELECT     dbo.TblCustomerContractDetails.Price - dbo.TblCustomerContractDetails.Discount AS net"
StrSQL = StrSQL & "  FROM         dbo.TblCustomerContract INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblCustomerContractDetails ON dbo.TblCustomerContract.TblCustomerContractD = dbo.TblCustomerContractDetails.TblCustomerContractD"
StrSQL = StrSQL & "  Where (dbo.TblCustomerContract.CustomerId = " & customerid & ")  And (dbo.TblCustomerContractDetails.ItemID = " & ItemID & ")"
StrSQL = StrSQL & " and " & SQLDate(Transaction_Date, True) & "   BETWEEN FromDate and Todate"
     If customerid = 2 Then '   Õ«·Â «·⁄„Ì· «·‰ﬁœÌ
    StrSQL = StrSQL & "  And ( dbo.TblCustomerContract.CashCustomerName  =  '" & mCashCust & "') "
    End If

    If SystemOptions.IsCustSalesManCashRelated Then 'Õ«·Â «·„‰œÊ»
    StrSQL = StrSQL & "  And (  dbo.TblCustomerContract.Emp_ID  = " & mSalesMan & ")"
     End If



Else


 StrSQL = " SELECT     ISNULL(dbo.TblVendorContractDetails.Price, 0) - ISNULL(dbo.TblVendorContractDetails.Discount, 0) AS net"
StrSQL = StrSQL & " FROM         dbo.TblVendorContract INNER JOIN"
StrSQL = StrSQL & " dbo.TblVendorContractDetails ON dbo.TblVendorContract.TblVendorContractD = dbo.TblVendorContractDetails.TblVendorContractD"
StrSQL = StrSQL & "  Where (dbo.TblVendorContractDetails.ItemID = " & ItemID & ") And (dbo.TblVendorContract.VendorID = " & customerid & ")" 'And (dbo.TblVendorContractDetails.unitid = " & unitid & ")

'StrSQL = StrSQL & "  Where (dbo.TblCustomerContract.VendorId = " & customerid & ") And (dbo.TblVendorContract.UnitID = " & unitid & ") And (dbo.TblVendorContract.ItemID = " & ItemID & ")"


End If








  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If (rs.RecordCount) > 0 Then

       total = IIf(IsNull(rs("net").value), 0, (rs("net").value))

   getpricebycustomrContract = total
 Else
 total = 0
 getpricebycustomrContract = 0
    End If


 End Function



 Function CheckChildforgroup(tablename As String, GroupIDFild As String, ParentIDFiles As String, GroupIDValue As Integer) As Boolean
    Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim total As Double
 StrSQL = "SELECT     COUNT(" & GroupIDFild & ") AS total"
StrSQL = StrSQL & "  From " & tablename & ""
StrSQL = StrSQL & "   WHERE     (" & ParentIDFiles & " = " & GroupIDValue & ")"
  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If (rs.RecordCount) > 0 Then


       total = IIf(IsNull(rs("total").value), 0, (rs("total").value))
  If total > 0 Then
   CheckChildforgroup = True
  Else
   CheckChildforgroup = False
  End If

         Else
         CheckChildforgroup = False
    End If


 End Function

 Function CheckCustomerSaleType(CusID As Double) As Double
    Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim total As Double
 StrSQL = "SELECT    * "
StrSQL = StrSQL & "  From  TblCustemers"
StrSQL = StrSQL & "   WHERE CusID =" & CusID
  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If (rs.RecordCount) > 0 Then


       CheckCustomerSaleType = IIf(IsNull(rs("SaleType").value), 0, (rs("SaleType").value))

    End If


 End Function
 Function GetProjectID(NoteID As Double)
     Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    StrSQL = "SELECT     project_id From dbo.Notes Where (NoteID = " & NoteID & ")"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    'rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.RecordCount) > 0 Then


       GetProjectID = IIf(IsNull(rs("project_id").value), 0, (rs("project_id").value))

         Else
         GetProjectID = 0
    End If

    rs.Close

 End Function
 Function updateopeningbalanceNewFromsqlTrialBalance2( _
        Optional FromDate As Date, _
        Optional ToDate As Date, _
        Optional continous As Boolean = False, _
        Optional ActivityId As Integer = 0, _
        Optional BranchID As Integer = 0, _
        Optional Account_code As String = "", _
        Optional updatetype As Integer = 0, _
        Optional composite As Boolean = False, _
        Optional lastlevel As Boolean = False _
    )

    'x1
    '0 balance Sheet
    '1 trial balances

    Dim openingbalacedate As Date
    Dim FromDate1 As Date
    Dim StrSQL As String

    ' getOpeningBalancedate P_DTPickerAccFrom , DTPickerAccTo.value, returnedfromdate, returnedTOdate, year(P_DTPickerAccFrom ), openingbalacedate
    getOpeningBalancedate , , , , year(ToDate), openingbalacedate, continous

    If openingbalacedate = FromDate Then

        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account)"
        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,0,last_account)"
        StrSQL = StrSQL & " ,CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,1,last_account)"

        ' GetBalanceCreditORdepitByActivity(@fromdate datetime,@Todate datetime ,@accountcode as varchar(255),@Credit_Or_Debit as integer,@Activity_Id as integer )
        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account)"
            StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & ActivityId & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & ActivityId & ",last_account)"
        End If

        If BranchID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account)"
            StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & BranchID & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & BranchID & ",last_account)"
        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            openingbalanceDes = "—’Ìœ «›  «ÕÌ   ›Ì" & openingbalacedate
        Else
            openingbalanceDes = "Opening Balance In " & openingbalacedate
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            openingbalanceDes = "—’Ìœ Õ Ï    " & FromDate - 1
        Else
            openingbalanceDes = " Balance Untill " & FromDate - 1
        End If

        FromDate1 = FromDate - 1
        StrSQL = " update ACCOUNTS"

        StrSQL = StrSQL & " set  balance= dbo.GetBalance('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & ")"
        StrSQL = StrSQL & " ,opening_balance= isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)" & _
                          " + isnull(dbo.GetBalance('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code ,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),0)"
        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,0,last_account)"
        StrSQL = StrSQL & " ,CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,1,last_account)"

        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account)"
            StrSQL = StrSQL & " ,opening_balance= isnull(dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)" & _
                              " + isnull(dbo.GetBalanceByActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0)"
            StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & ActivityId & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & ActivityId & ",last_account)"
        End If

        If BranchID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account)"
            StrSQL = StrSQL & " ,opening_balance= isnull(dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)" & _
                              " + isnull(dbo.GetBalanceByBranch('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0)"
            StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & BranchID & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & BranchID & ",last_account)"
        End If

    End If

    '==========================
    ' ‘—Êÿ WHERE ··Ã“¡ «·√Ê·
    '==========================
    If updatetype = 1 Then  ' „Ì“«‰
        StrSQL = StrSQL & " WHERE (last_account = 1)  and  (AccountTypes = 1 or AccountTypes = 0)"

    ElseIf updatetype = 5 Then  ' „Ì“«‰ „” ÊÌ« 

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & " WHERE Account_Code in (" & Account_code & ")"
            If lastlevel = False Then
                StrSQL = StrSQL & " AND (last_account = 0)"
            End If
        Else
            If lastlevel = False Then
                StrSQL = StrSQL & " WHERE (last_account = 0)"
            End If
        End If

    ElseIf updatetype = 2 Then  ' Õ”«» «” «–

        If getAccountTypes(Account_code) <> 1 Then ' ·Ê ﬂ«‰ Õ”«» „’—Ê› «Ê «Ì—«œ
            GoTo Part2
        End If

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  WHERE  Parent_Account_Code ='" & Account_code & "'"
        End If

    ElseIf updatetype = 3 Or updatetype = 4 Then '„‘—Ê⁄     '  ﬂ‘› Õ”«» „ÊŸ›

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  WHERE  Account_Code in (" & Account_code & ")"
        End If

    Else
        'StrSQL = StrSQL & " WHERE     (last_account = 1) "
    End If

    Cn.CommandTimeout = 10000
    Cn.Execute StrSQL

    'part2****************************************************************************
    If getAccountTypes(Account_code) = 1 Then ' ·Ê ﬂ«‰ Õ”«»   „Ì“«‰Ì…
        Exit Function
    End If

Part2:
    openingbalacedate = GetOpeningBalanceDateForType2(FromDate)

    If SystemOptions.UserInterface = ArabicInterface Then
        openingbalanceDes = "—’Ìœ Õ Ï    " & FromDate - 1
    Else
        openingbalanceDes = " Balance Untill " & FromDate - 1
    End If

    FromDate1 = FromDate - 1
    StrSQL = " update ACCOUNTS"

    StrSQL = StrSQL & " set  balance= dbo.GetBalance('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & ")"
    StrSQL = StrSQL & " ,opening_balance= isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)" & _
                      " + isnull(dbo.GetBalance('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code ,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),0)"
    StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,0,last_account)"
    StrSQL = StrSQL & " ,CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,1,last_account)"

    If ActivityId <> 0 Then
        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account)"
        StrSQL = StrSQL & " ,opening_balance= isnull(dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)" & _
                          " + isnull(dbo.GetBalanceByActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0)"
        StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & ActivityId & ",last_account)"
        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & ActivityId & ",last_account)"
    End If

    If BranchID <> 0 Then

        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account)"
        StrSQL = StrSQL & " ,opening_balance= isnull(dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)" & _
                          " + isnull(dbo.GetBalanceByBranch('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0)"
        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & BranchID & ",last_account)"
        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & BranchID & ",last_account)"

    End If

    '==========================
    ' ‘—Êÿ WHERE ··Ã“¡ «·À«‰Ì
    '==========================
    If updatetype = 1 Then  ' „Ì“«‰
        StrSQL = StrSQL & " WHERE     (last_account = 1)  and  (AccountTypes = 2) "

    ElseIf updatetype = 5 Then 'ÿ„” ÊÌ«    ' „Ì“«‰

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  WHERE  Account_Code in (" & Account_code & ") AND (AccountTypes = 2)"
        Else
            StrSQL = StrSQL & "  WHERE  (AccountTypes = 2)"
        End If

        If lastlevel = False Then
            StrSQL = StrSQL & " AND (last_account = 0)"
        End If

    ElseIf updatetype = 2 Then  ' Õ”«» «” «–

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  WHERE  Parent_Account_Code ='" & Account_code & "'"
        End If

    ElseIf updatetype = 3 Or updatetype = 4 Then ' «Ê „‘—Ê⁄    '  ﬂ‘› Õ”«» „ÊŸ›

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  WHERE    (AccountTypes = 2)  and Account_Code in (" & Account_code & ")"
        End If

    Else

        StrSQL = StrSQL & " WHERE     (last_account = 1) "
    End If

    Cn.CommandTimeout = 10000
    Cn.Execute StrSQL

End Function

Function updateopeningbalanceNewFromsqlTrialBalanceOld(Optional FromDate As Date, Optional ToDate As Date, Optional continous As Boolean = False, Optional ActivityId As Integer = 0, Optional BranchID As Integer = 0, Optional Account_code As String = "", Optional updatetype As Integer = 0, Optional composite As Boolean, Optional lastlevel As Boolean = False)
'x1
    '0 balance Sheet
    '1 trial balances
    Dim openingbalacedate As Date
    ' getOpeningBalancedate P_DTPickerAccFrom , DTPickerAccTo.value, returnedfromdate, returnedTOdate, year(P_DTPickerAccFrom ), openingbalacedate
    getOpeningBalancedate , , , , year(ToDate), openingbalacedate, continous

    Dim StrSQL As String

    If openingbalacedate = FromDate Then

        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account)"
        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,0,last_account)"
        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,1,last_account)"

        ' GetBalanceCreditORdepitByActivity(@fromdate datetime,@Todate datetime ,@accountcode as varchar(255),@Credit_Or_Debit as integer,@Activity_Id as integer )
        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account)"
            StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & ActivityId & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & ActivityId & ",last_account)"

        End If

        If BranchID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account)"
            '      strsql = strsql & " balance= dbo.GetBalanceByBranch('" & SQLDate(fromdate) & "','" & SQLDate(todate) & "'," & BranchId & ", Account_code)"
            StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & BranchID & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & BranchID & ",last_account)"

        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            openingbalanceDes = "—’Ìœ «›  «ÕÌ   ›Ì" & openingbalacedate
        Else
            openingbalanceDes = "Opening Balance In " & openingbalacedate
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            openingbalanceDes = "—’Ìœ Õ Ï    " & FromDate - 1
        Else
            openingbalanceDes = " Balance Untill " & FromDate - 1
        End If

        Dim FromDate1 As Date
        FromDate1 = FromDate - 1
        StrSQL = " update ACCOUNTS"

        StrSQL = StrSQL & " set  balance= dbo.GetBalance('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)  +   isnull(dbo.GetBalance('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code ,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & ")),0) "
        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,0,last_account)"
        StrSQL = StrSQL & " ,CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,1,last_account)"

        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
            StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & ActivityId & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & ActivityId & ",last_account)"

        End If

        If BranchID <> 0 Then

            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByBranch('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0) "
            StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & BranchID & ",last_account)"
            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & BranchID & ",last_account)"

        End If

    End If

    If updatetype = 1 Then  ' „Ì“«‰
        StrSQL = StrSQL & " WHERE     (last_account = 1)  and  (AccountTypes = 1 or AccountTypes = 0)"

    ElseIf updatetype = 5 Then  ' „Ì“«‰ „” ÊÌ« 

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_Code in (" & Account_code & ")"

        End If

                If lastlevel = False Then
                    StrSQL = StrSQL & " WHERE     (last_account = 0) "
                End If

    ElseIf updatetype = 2 Then  ' Õ”«» «” «–

        If getAccountTypes(Account_code) <> 1 Then ' ·Ê ﬂ«‰ Õ”«» „’—Ê› «Ê «Ì—«œ
            GoTo Part2
        End If

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Parent_Account_Code ='" & Account_code & "'"

        End If

    ElseIf updatetype = 3 Or updatetype = 4 Then '„‘—Ê⁄     '  ﬂ‘› Õ”«» „ÊŸ›

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_Code in (" & Account_code & ")"

        End If

    Else

        'StrSQL = StrSQL & " WHERE     (last_account = 1) "
    End If

    'StrSQL = StrSQL & " WHERE     (last_account = 1) "

    Cn.CommandTimeout = 10000

    Cn.Execute StrSQL
    'DoEvents

    'part2****************************************************************************
    If getAccountTypes(Account_code) = 1 Then ' ·Ê ﬂ«‰ Õ”«»   „Ì“«‰Ì…
        Exit Function
    End If

Part2:
    openingbalacedate = GetOpeningBalanceDateForType2(FromDate)

    If SystemOptions.UserInterface = ArabicInterface Then
        openingbalanceDes = "—’Ìœ Õ Ï    " & FromDate - 1
    Else
        openingbalanceDes = " Balance Untill " & FromDate - 1
    End If

    FromDate1 = FromDate - 1
    StrSQL = " update ACCOUNTS"

    StrSQL = StrSQL & " set  balance= dbo.GetBalance('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
    StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)  +   isnull(dbo.GetBalance('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code ,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & ")),0) "
    StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,0,last_account)"
    StrSQL = StrSQL & " ,CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,1,last_account)"

    If ActivityId <> 0 Then
        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
        StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & ActivityId & ",last_account)"
        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & ActivityId & ",last_account)"

    End If

    If BranchID <> 0 Then

        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByBranch('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0) "
        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,0," & BranchID & ",last_account)"
        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "',Account_code,1," & BranchID & ",last_account)"

    End If

    If updatetype = 1 Then  ' „Ì“«‰
        StrSQL = StrSQL & " WHERE     (last_account = 1)  and  (AccountTypes = 2) "
    ElseIf updatetype = 5 Then 'ÿ„” ÊÌ«    ' „Ì“«‰
                If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_Code in (" & Account_code & ")"

        End If
        If lastlevel = False Then
        StrSQL = StrSQL & " WHERE     (last_account = 0)  and  (AccountTypes = 2)"
        End If

    ElseIf updatetype = 2 Then  ' Õ”«» «” «–

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Parent_Account_Code ='" & Account_code & "'"

        End If

    ElseIf updatetype = 3 Or updatetype = 4 Then ' «Ê „‘—Ê⁄    '  ﬂ‘› Õ”«» „ÊŸ›

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where    (AccountTypes = 2)  and Account_Code in (" & Account_code & ")"

        End If

    Else

        StrSQL = StrSQL & " WHERE     (last_account = 1) "
    End If

    'StrSQL = StrSQL & " WHERE     (last_account = 1) "

    Cn.CommandTimeout = 10000

    Cn.Execute StrSQL
    'DoEvents

End Function
Public Sub GetVocationEntitlementsx(ID As Integer, _
Optional BranchID As Integer, _
  Optional EmpID As Integer, _
                                Optional ByRef salary As Double, _
                                 Optional ByRef SalEntitOther As Double, _
                                 Optional ByRef other As Double, _
                                 Optional ByRef Advance As Double, _
                                 Optional ByRef ValueTickt As Double, _
                                 Optional ByRef SalaryVocation As Double, Optional ByRef InsuranceValue As Double)

     Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    StrSQL = "select * From tblVocationEntitlements     where id =" & ID & " and PayedPayment is null "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If (rs.RecordCount) > 0 Then
       BranchID = IIf(IsNull(rs("BranchID").value), Current_branch, (rs("BranchID").value))
      EmpID = IIf(IsNull(rs("EmpID").value), 0, (rs("EmpID").value))

       salary = IIf(IsNull(rs("salary").value), 0, (rs("salary").value))
       SalEntitOther = IIf(IsNull(rs("SalEntitOther").value), 0, (rs("SalEntitOther").value))
       other = IIf(IsNull(rs("Other").value), 0, (rs("Other").value))
       Advance = IIf(IsNull(rs("Advance").value), 0, (rs("Advance").value))
    ValueTickt = IIf(IsNull(rs("ValueTickt").value), 0, (rs("ValueTickt").value))
               SalaryVocation = IIf(IsNull(rs("SalaryVocation").value), 0, (rs("SalaryVocation").value))
               InsuranceValue = IIf(IsNull(rs("InsuranceValue").value), 0, (rs("InsuranceValue").value))
         Else
         BranchID = 0
         EmpID = 0
         salary = 0
         SalEntitOther = 0
         other = 0
         ValueTickt = 0
         Advance = 0
    End If
    rs.Close
End Sub



      Public Sub OrderExchange(Serial1 As String, _
                                Optional ByRef Type1 As Integer, _
                                Optional ByRef txtperson As String, Optional ByRef des As String, Optional ByRef Price As Double, Optional ByRef EmpID As Integer, Optional ByRef basedOn As Integer, Optional ByRef orderNo As String, Optional ByRef Transaction_ID As Long, Optional ByRef CusID As Double, Optional ByRef FromType As Integer, Optional ByRef Account_code As String, Optional CurrcyID As Integer, Optional Rate As Double, Optional valuee As Double, Optional salary_or_advance As Integer)

     Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblExchange     where NoteSerial1 ='" & Serial1 & "'"
    If SystemOptions.MonyeIssueVchrNoMust = True Then
    StrSQL = StrSQL & "   and price >0 "
    End If

     If CheckAprroveScreen("FrmTypeExchange") = True Then

  StrSQL = StrSQL & "  and Approved = 1"
  End If


    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    'rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.RecordCount) > 0 Then

       Account_code = IIf(IsNull(rs("Account_Code").value), -1, (rs("Account_Code").value))
       FromType = IIf(IsNull(rs("FromType").value), -1, (rs("FromType").value))
       CusID = IIf(IsNull(rs("CusID").value), 0, (rs("CusID").value))
      basedOn = IIf(IsNull(rs("basedOn").value), 0, (rs("basedOn").value))
         Type1 = IIf(IsNull(rs("Type").value), 0, (rs("Type").value))
     salary_or_advance = IIf(IsNull(rs("salary_or_advance").value), 0, (rs("salary_or_advance").value))

       EmpID = IIf(IsNull(rs("EmpID").value), 0, (rs("EmpID").value))




       Price = IIf(IsNull(rs("Price").value), 0, (rs("Price").value))
       des = IIf(IsNull(rs("Des").value), "", (rs("Des").value))
       txtperson = IIf(IsNull(rs("ToPerson").value), "", (rs("ToPerson").value))
       orderNo = IIf(IsNull(rs("orderNo").value), 0, (rs("orderNo").value))
       Transaction_ID = IIf(IsNull(rs("Transaction_ID").value), 0, (rs("Transaction_ID").value))
       CurrcyID = IIf(IsNull(rs("CurrcyID").value), MainCurrency, (rs("CurrcyID").value))
       valuee = IIf(IsNull(rs("PriceE").value), 0, (rs("PriceE").value))
       Rate = IIf(IsNull(rs("Rate").value), 1, (rs("Rate").value))


               Else
               Price = -1
    End If

    rs.Close

End Sub

Public Sub OrderExchangeold(Serial1 As String, _
                                Optional ByRef Type1 As Integer, _
                                Optional ByRef txtperson As String, Optional ByRef des As String, Optional ByRef Price As Double, Optional ByRef EmpID As Integer, Optional ByRef basedOn As Integer, Optional ByRef orderNo As String)

     Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblExchange     where NoteSerial1 ='" & Serial1 & "'"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    'rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.RecordCount) > 0 Then

      basedOn = IIf(IsNull(rs("basedOn").value), 0, (rs("basedOn").value))
         Type1 = IIf(IsNull(rs("Type").value), 0, (rs("Type").value))
       EmpID = IIf(IsNull(rs("EmpID").value), 0, (rs("EmpID").value))
       Price = IIf(IsNull(rs("Price").value), 0, (rs("Price").value))
       des = IIf(IsNull(rs("des").value), "", (rs("des").value))
       txtperson = IIf(IsNull(rs("ToPerson").value), "", (rs("ToPerson").value))
               orderNo = IIf(IsNull(rs("orderNo").value), 0, (rs("orderNo").value))
               Else
               Price = -1
    End If

    rs.Close

End Sub

 Public Function GetACCOUNTSCode(LngItemID As String, Optional ID As Integer = 0) As String
    Dim StrSQL As String
    Dim rs As ADODB.Recordset

    If LngItemID <> "" Then
   If ID = 1 Then
       StrSQL = "Select Account_Serial  From ACCOUNTS Where Account_Code='" & LngItemID & "'"
     Else
     StrSQL = "Select Account_Code  From ACCOUNTS Where Account_Serial='" & LngItemID & "'"
     End If
        Set rs = New ADODB.Recordset

        If Cn.State = adStateClosed Then
            open_my_connection
        End If

        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
        If ID = 1 Then
            GetACCOUNTSCode = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
            Else
            GetACCOUNTSCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
         End If
        Else
        End If

        rs.Close
        Set rs = Nothing
    End If

End Function
Public Sub printCopounBarcode(m_PrintTarget As PrintTarget, Optional Serial1 As Double)

    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo
  '  Name = Name & ".rpt"

    If Dir(App.path & "\REPORTS\REPORTS NEW\" & "BarCodeCopoun.rpt") = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        MsgBox "«· ﬁ—Ì— €Ì— „ÊÃÊœ"
        Exit Sub
    End If

        MySQL = " "


  MySQL = " SELECT     dbo.TblCoupons.ID, dbo.TblCoupons.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCoupons.RecordDate, "
  MySQL = MySQL & "                     dbo.TblCoupons.Remarks, dbo.TblCoupons.FromDate, dbo.TblCoupons.ToDate, dbo.TblCoupons.FromDate2, dbo.TblCoupons.ToDate2, dbo.TblCoupons.RdTyp,"
  MySQL = MySQL & "                    dbo.TblCoupons.Num, dbo.TblCoupons.Vlue, dbo.TblCouponsDet.Remarks AS RemarksDet, dbo.TblCouponsDet.TypTrans, dbo.TblCouponsDet.Num AS NumDet,"
  MySQL = MySQL & "                    dbo.TblCouponsDet.Vlue AS VlueDet, dbo.TblCouponsDet.FromVlue, dbo.TblCouponsDet.TOVlue, dbo.TblCouponsDet.BillNo, dbo.TblCouponsDet.ReturnBillNo,"
  MySQL = MySQL & "                    dbo.TblCouponsDet.Transaction_ID, dbo.TblCouponsDet.RetTransaction_ID, dbo.TblCouponsDet.NewBillNo, dbo.TblCouponsDet.NewTransaction_ID,"
  MySQL = MySQL & "                    dbo.TblCouponsDet.discount"
  MySQL = MySQL & "     FROM         dbo.TblCoupons LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCouponsDet ON dbo.TblCoupons.ID = dbo.TblCouponsDet.CoupID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.TblCoupons.BranchID = dbo.TblBranchesData.branch_id"
  MySQL = MySQL & "     Where ( TypTrans =1 and dbo.TblCoupons.ID =" & Serial1 & ")"
  MySQL = MySQL & "ORDER BY dbo.TblCouponsDet.Vlue"




    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass



        Set xReport = xApp.OpenReport(App.path & "\REPORTS\REPORTS NEW\" & "BarCodeCopoun.rpt")
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName



    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title

    Set CViewer = New ClsReportViewer
hide_logo = True
    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, App.path & "\REPORTS\REPORTS NEW\" & "BarCodeCopoun.rpt"

    Set xApp = Nothing
    Set xReport = Nothing
    Screen.MousePointer = vbDefault
    hide_logo = False
End Sub
 Public Sub printCodeBarcode(m_PrintTarget As PrintTarget, Optional Name As String, Optional lblindex As Integer, Optional UnitName As String = "")

    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo
    Name = Name & ".rpt"

    If Dir(App.path & "\Reports\Inventory\" & Name) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        MsgBox "«· ﬁ—Ì— €Ì— „ÊÃÊœ"
        Exit Sub
    End If

        MySQL = " "
If lblindex = 1 Then
MySQL = " SELECT  DealerPrice, code128,TblPrintBarCode.ProductionDate,   dbo.TblItems.ItemComment  ,  dbo.TblItems.TotalCalories, dbo.TblItems.shortName,   dbo.TblItems.PrintedName,    dbo.TblItems.ItemCode, dbo.TblItems.ItemNamee, dbo.TblPrintBarCode.Code, dbo.TblPrintBarCode.PartNo, dbo.TblPrintBarCode.Cost, dbo.TblPrintBarCode.Name, dbo.TblPrintBarCode.Color, "
MySQL = MySQL & "                          dbo.TblPrintBarCode.size, dbo.TblPrintBarCode.class, dbo.TblPrintBarCode.CodeAnalisys, dbo.TblPrintBarCode.ExpiryDate, dbo.TblPrintBarCode.LotNO, dbo.TblPrintBarCode.VatYou, dbo.TblPrintBarCode.VAT,"
MySQL = MySQL & "                          dbo.TblPrintBarCode.Total"
MySQL = MySQL & "  FROM            dbo.TblPrintBarCode LEFT OUTER JOIN"
MySQL = MySQL & "                          dbo.TblItems ON dbo.TblPrintBarCode.Item_ID = dbo.TblItems.ItemID"

Else
 MySQL = " SELECT  DealerPrice,   code128,  dbo.TblItems.ItemComment  ,  dbo.TblItems.TotalCalories, dbo.TblItems.shortName,   dbo.TblItems.PrintedName,   dbo.TblPrintBarCode.Code, dbo.TblPrintBarCode.PartNo, dbo.TblPrintBarCode.Cost, dbo.TblPrintBarCode.Name, dbo.TblPrintBarCode.Color, dbo.TblPrintBarCode.size, dbo.TblPrintBarCode.class, "
 MySQL = MySQL & "                        dbo.ItemsDetails.ItemDetailedCode , dbo.ItemsDetails.ItemID, dbo.TblItems.ItemNamee, dbo.TblPrintBarCode.VatYou, dbo.TblPrintBarCode.VAT, dbo.TblPrintBarCode.Total"
 MySQL = MySQL & "          FROM            dbo.TblItems RIGHT OUTER JOIN"
 MySQL = MySQL & "                        dbo.ItemsDetails ON dbo.TblItems.ItemID = dbo.ItemsDetails.ItemId RIGHT OUTER JOIN"
 MySQL = MySQL & "                        dbo.TblPrintBarCode ON dbo.ItemsDetails.ParrtNoCode = dbo.TblPrintBarCode.Code"
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

'    If SystemOptions.UserInterface = EnglishInterface Then

'    Else

        Set xReport = xApp.OpenReport(App.path & "\Reports\Inventory\" & Name)
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName

'    End If
xReport.ParameterFields(4).AddCurrentValue UnitName
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title

    Set CViewer = New ClsReportViewer
hide_logo = True
    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, App.path & "\Reports\Inventory\" & Name, , MySQL

    Set xApp = Nothing
    Set xReport = Nothing
    Screen.MousePointer = vbDefault
    hide_logo = False
End Sub


Public Function get_Customer_information(ID As Integer, Optional ByRef Mobile As String)
   Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
sql = sql & " select * from TblCustemers "

sql = sql & " WHERE     (CusID = " & ID & ")"

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount > 0 Then

        Mobile = IIf(IsNull(Rs3("Cus_mobile").value), "", Rs3("Cus_mobile").value)
      End If
End Function

Public Function SavecustomerData(cuPhone As String, cuname As String)
Dim customername As String
If cuPhone = "" Then Exit Function

customername = GetCashCustomernamebyphone(cuPhone)
Dim ID As String

If customername = "" Then
ID = new_id("TblCusCsh", "id", "")

  add_record_to_table "TblCusCsh", "id,name,namee,tel", ID & ",'" & cuname & "','" & cuname & "','" & cuPhone & "'", "id", val(ID)


End If

End Function
Public Function GetCommisionPercentages(typeid As Integer, EmpID, Optional ByRef Rent As Double, Optional ByRef InternalComm As Double, Optional ByRef ExternalComm As Double, Optional ByRef Revenue As Double)
Dim RsDetails As ADODB.Recordset
Dim StrSQL As String
Dim GroupID As Integer
    Set RsDetails = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.TBLSalesRepData Where (EmpID =" & EmpID & ")"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   GroupID = IIf(IsNull(RsDetails("GroupID").value), 0, RsDetails("GroupID").value)

   Else
   GroupID = 0
   End If

   RsDetails.Close
   Set RsDetails = Nothing
       Set RsDetails = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.TBLSalesRepGroups Where (typeid =" & typeid & ")"
         If GroupID <> 0 Then
         StrSQL = StrSQL & " and id=" & GroupID
         End If

   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   Rent = IIf(IsNull(RsDetails("Rent").value), 0, RsDetails("Rent").value)
   InternalComm = IIf(IsNull(RsDetails("InternalComm").value), 0, RsDetails("InternalComm").value)
   ExternalComm = IIf(IsNull(RsDetails("ExternalComm").value), 0, RsDetails("ExternalComm").value)
   Revenue = IIf(IsNull(RsDetails("Revenue").value), 0, RsDetails("Revenue").value)
  typeid = IIf(IsNull(RsDetails("TypeiD").value), 0, RsDetails("TypeiD").value)
   End If


   If typeid > 2 Then
      RsDetails.Close
   Set RsDetails = Nothing
       Set RsDetails = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.TBLSalesRepGroups Where (typeid =" & typeid & ")"
         If GroupID <> 0 Then
         StrSQL = StrSQL & " and id=" & GroupID
         End If

   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   Rent = IIf(IsNull(RsDetails("Rent").value), 0, RsDetails("Rent").value)
   InternalComm = IIf(IsNull(RsDetails("InternalComm").value), 0, RsDetails("InternalComm").value)
   ExternalComm = IIf(IsNull(RsDetails("ExternalComm").value), 0, RsDetails("ExternalComm").value)
   Revenue = IIf(IsNull(RsDetails("Revenue").value), 0, RsDetails("Revenue").value)
  typeid = IIf(IsNull(RsDetails("TypeiD").value), 0, RsDetails("TypeiD").value)
   End If

   End If


End Function
Public Function checkDepositeRent(ID As Integer, ddate As Date) As String
Dim RsDetails As ADODB.Recordset
Dim StrSQL As String
    Set RsDetails = New ADODB.Recordset
         StrSQL = " SELECT     dbo.Notes.NoteDate, dbo.Notes.NoteSerial1, dbo.Notes.AllowDate, dbo.Notes.AllowDateH, dbo.Notes.renterName, dbo.Notes.akarid, dbo.TblAqar.aqarNo, "
StrSQL = StrSQL & "   dbo.TblAqar.aqarname, dbo.TblAkarUnit.name AS unittype, dbo.TblAkarUnit.namee AS unittypee, dbo.TblAqarDetai.unitno, dbo.TblAqarDetai.Id AS unitid"
StrSQL = StrSQL & "  FROM         dbo.Notes INNER JOIN"
StrSQL = StrSQL & "    dbo.TblAqar ON dbo.Notes.akarid = dbo.TblAqar.Aqarid INNER JOIN"
StrSQL = StrSQL & "     dbo.TblAkarUnit ON dbo.Notes.unittype = dbo.TblAkarUnit.id INNER JOIN"
StrSQL = StrSQL & "    dbo.TblAqarDetai ON dbo.Notes.UnitNo = dbo.TblAqarDetai.Id"
StrSQL = StrSQL & "   WHERE     (dbo.Notes.NoteDate <= " & SQLDate(ddate, True) & "  ) AND (dbo.TblAqarDetai.Id = " & ID & ")"



   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   checkDepositeRent = "ÌÊÃœ ⁄—»Ê‰ ⁄·Ï Â–… «·ÊÕœ… »√”„  " & IIf(IsNull(RsDetails("renterName").value), "", RsDetails("renterName").value) & "  Ì‰ ÂÌ ›Ì " & IIf(IsNull(RsDetails("AllowDate").value), "", RsDetails("AllowDate").value) & "    «·„Ê«›ﬁ " & IIf(IsNull(RsDetails("AllowDateh").value), "", RsDetails("AllowDateh").value)
   Else
   checkDepositeRent = ""
   End If
End Function

Public Function checkEmpDiscount(EmpID As Integer, value As Double, discount As Double) As Boolean
Dim RsDetails As ADODB.Recordset
Dim StrSQL As String
Dim discountvalue As Double
checkEmpDiscount = False
    Set RsDetails = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TBLSalesRepData  Where (EmpID =" & EmpID & ")"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   discountvalue = IIf(IsNull(RsDetails("DiscountValue").value), 0, RsDetails("DiscountValue").value)
   If value * discountvalue / 100 >= discount Then
   checkEmpDiscount = True
   Else
   checkEmpDiscount = False
   End If


   Else
   checkEmpDiscount = True
   End If
End Function
Public Function getLastLevel() As Integer
Dim RsDetails As ADODB.Recordset
Dim StrSQL As String

    Set RsDetails = New ADODB.Recordset
         StrSQL = " SELECT     MAX([Level]) AS lastlevel FROM         dbo.AccountsLevelsDetails"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   getLastLevel = IIf(IsNull(RsDetails("lastlevel").value), 0, RsDetails("lastlevel").value)
   Else
   getLastLevel = 0
   End If
End Function
Public Function checkContractTransactions(ContNo As Double) As Boolean
Dim RsDetails As ADODB.Recordset
Dim StrSQL As String
checkContractTransactions = False
    Set RsDetails = New ADODB.Recordset
         StrSQL = "SELECT     *   from dbo.Notes Where (ContNo =" & ContNo & ") and dbo.Notes.CashingType=8 "
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   checkContractTransactions = True
   Else
   checkContractTransactions = False
   End If
End Function


Public Function checkOutContract(ContNo As Integer) As Integer
Dim RsDetails As ADODB.Recordset
Dim StrSQL As String
    Set RsDetails = New ADODB.Recordset
         StrSQL = "SELECT     OutContract   from dbo.TblContract Where (ContNo =" & ContNo & ")"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   checkOutContract = IIf(IsNull(RsDetails("OutContract").value), 0, RsDetails("OutContract").value)
   Else
   checkOutContract = 0
   End If
End Function
Public Function CheckUnitContract(unitno As Integer) As Boolean
Dim RsDetails As ADODB.Recordset
Dim StrSQL As String
    Set RsDetails = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.TblContract Where (UnitNo =" & unitno & ")"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   CheckUnitContract = True
   Else
   CheckUnitContract = False
   End If
End Function

Public Function CheckUnitContractxxx(unitno As Integer) As Boolean
Dim RsDetails As ADODB.Recordset
Dim StrSQL As String
    Set RsDetails = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.TblContract Where (UnitNo =" & unitno & ")"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails.RecordCount > 0 Then
   CheckUnitContractxxx = True
   End If
End Function

Public Function AlarmsDates()
Dim AskOption As Boolean
Dim Askinterval As String
Dim Askcount As Integer
    On Error Resume Next
        Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Set rs = New ADODB.Recordset
 Dim rentInstallmentdate As Date
  AskOption = GetSetting(StrAppRegPath, "View_Type", "RentInstallments", True)
    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_RentInstallments", "")
    Askcount = GetSetting(StrAppRegPath, "Setting", "Count_RentInstallments", 0)

    If AskOption = True And Askinterval <> "" Then
    rentInstallmentdate = DateAdd((Askinterval), 1 * Askcount, Date)

    End If
My_SQL = " SELECT     dbo.TblContractInstallments.*, dbo.TblCustemers.CusName AS CusName, dbo.TblCustemers.Cus_mobile AS Cus_mobile, dbo.TblCustemers.CusNamee AS Expr3, "
My_SQL = My_SQL & "                      dbo.TblContract.NoteSerial1 AS NoteSerial11, dbo.TblContract.CusID AS Expr5, dbo.TblContract.StrDate AS Expr6, dbo.TblContractInstallments.Installdate AS Expr7,"
My_SQL = My_SQL & "                      dbo.TblContractInstallments.InstalldateH AS Expr9, dbo.TblContractInstallments.InstallNo AS Expr10, dbo.TblContractInstallments.Commissions AS Expr11,"
My_SQL = My_SQL & "                      dbo.TblContractInstallments.RentValue AS RentValue_1, dbo.TblContractInstallments.Insurance AS Insurance_1, dbo.TblContract.ContNo AS Expr8,"
My_SQL = My_SQL & "                      { fn IFNULL(dbo.ContracttBillInstallmentsDone.[Value], 0) } AS Allpayed, { fn IFNULL(dbo.TblContractInstallments.installValue, 0)"
My_SQL = My_SQL & "                      } - { fn IFNULL(dbo.ContracttBillInstallmentsDone.[Value], 0) } AS newremains, dbo.TblAqar.aqarNo AS IaqarNo, dbo.TblAqar.aqarname AS Iaqarname,"
My_SQL = My_SQL & "                      dbo.TblAkarUnit.name AS Unitname, dbo.TblAkarUnit.namee AS Unitnamee, dbo.TblAqarDetai.unitno AS unitnoNam, dbo.TblContract.Phone AS Phone"
My_SQL = My_SQL & " FROM         dbo.TblContractInstallments INNER JOIN"
My_SQL = My_SQL & "                      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo INNER JOIN"
My_SQL = My_SQL & "                      dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.ContracttBillInstallmentsDone ON dbo.TblContract.ContNo = dbo.ContracttBillInstallmentsDone.istallid"
My_SQL = My_SQL & " WHERE   (dbo.TblContractInstallments.Status = 0 OR dbo.TblContractInstallments.Status IS NULL)"

'My_SQL = My_SQL & " WHERE   (dbo.TblContractInstallments.Status = 0 OR dbo.TblContractInstallments.Status IS NULL)"



       My_SQL = My_SQL + " and (Installdate <=" & SQLDate(rentInstallmentdate, True) & ")"



    My_SQL = My_SQL + "   AND (dbo.TblContract.Branch_NO = " & Current_branch & ")"


My_SQL = My_SQL + "   order by Installdate "
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
If rs.RecordCount > 0 Then

RSRentAlarm.show
RSRentAlarm.FillGrid My_SQL
End If



'StrSQL = " SELECT     dbo.TblCardAuthorizationReform.ID,dbo.TblCardAuthorizationReform.SendSMS, dbo.TblCardAuthorizationReform.RecordDate, dbo.TblCardAuthorizationReform.ClientName,"
'StrSQL = StrSQL & "      dbo.TblCardAuthorizationReform.Telephone, dbo.TblCardAuthorizationReform.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
'StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.CarModelID, dbo.TblCarModels.CarID, dbo.TblCarModels.ModelE, dbo.TblCarModels.Model,"
'StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.PlateNo, dbo.TblCardAuthorizationReform.ColorID, dbo.TblColor.name AS namecolor, dbo.TblColor.namee AS nameecolor,"
'StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.YearFact, dbo.TblCardAuthorizationReform.OrderStatus, dbo.TblCardAuthorizationReform.Accept,"
'StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.EndDate, dbo.TblCardAuthorizationReform.Month_Day, dbo.TblCardAuthorizationReform.Granty,"
'StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.DateEndG, dbo.TblCardAuthorizationReform.DateStartG, dbo.TblCardAuthorizationReform.CarMeter,"
'StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.LongGranty, dbo.TblCardAuthorizationReform.PayFirst, dbo.TblCardAuthorizationReform.AmountAccept,"
'StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.Complaint, dbo.TblCardAuthorizationReform.Noteinitial, dbo.TblCardAuthorizationReform.Shaseh,"
'StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.NotAccept, dbo.TblCardAuthorizationReform.RecordeTime, dbo.TblCardAuthorizationReform.typerequest,"
'StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.FitterID, dbo.TblUsers.UserName, dbo.TblCardAuthorizationReform.mobile, dbo.TblCardAuthorizationReform.Cash,"
'StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.Accoun, dbo.TblCardAuthorizationReform.credit, dbo.TblCardAuthorizationReform.fax, dbo.TblCardAuthorizationReform.box,"
'StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.email, dbo.TblCardAuthorizationReform.address, dbo.TblCardAuthorizationReform.boxzip, dbo.TblCardAuthorizationReform.codereg,"
'StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.typereg, dbo.TblCardAuthorizationReform.codedoor, dbo.TblCardAuthorizationReform.DateEnter,"
'StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.DateExit, dbo.TblCardAuthorizationReform.persons, dbo.TblCardAuthorizationReform.Companies,"
'StrSQL = StrSQL & "     dbo.TblCardAuthorizationReform.driver, dbo.TblCardAuthorizationReform.DateAcutExite, dbo.TblCardAuthorizationReform.DateExptExit,"
'StrSQL = StrSQL & "    dbo.TblCardAuthorizationReform.TimeAcutExite , dbo.TblCardAuthorizationReform.TimeExptExit, dbo.TblCardAuthorizationReform.ResonUnderWait, dbo.TblCardAuthorizationReform.Payed"
'StrSQL = StrSQL & " FROM    dbo.TblCardAuthorizationReform LEFT OUTER JOIN "
'StrSQL = StrSQL & "  dbo.TblUsers ON dbo.TblCardAuthorizationReform.FitterID = dbo.TblUsers.UserID LEFT OUTER JOIN"
' StrSQL = StrSQL & " dbo.TBLCarTypes ON dbo.TblCardAuthorizationReform.CarTypeID = dbo.TBLCarTypes.id FULL OUTER JOIN"
'StrSQL = StrSQL & " dbo.TblColor ON dbo.TblCardAuthorizationReform.ColorID = dbo.TblColor.Id FULL OUTER JOIN"
'StrSQL = StrSQL & "  dbo.TblCarModels ON dbo.TblCardAuthorizationReform.CarModelID = dbo.TblCarModels.Id"
'StrSQL = StrSQL & " Where  (dbo.TblCardAuthorizationReform.OrderStatus <=10)"
'
'
'StrSQL = StrSQL & " and (TblCardAuthorizationReform.RecordDate <=  DATEadd( mm,-1,GETDATE()) and TblCardAuthorizationReform.RecordDate >=  DATEadd( mm,-2,GETDATE()) and TblCardAuthorizationReform.orderStatus < 2 )  "
'rs.Close
'    rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'If rs.RecordCount > 0 Then
'
'FrmCarReporonlin2.show
''FrmCarReporonlin2.FillGrid My_SQL
'End If




End Function
Function dhLastDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the last day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhLastDayInMonth = DateSerial(year(dtmDate), _
     Month(dtmDate) + 1, 0)
End Function
Public Function dhFirstDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the first day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhFirstDayInMonth = DateSerial(year(dtmDate), _
     Month(dtmDate), 1)
End Function
Public Function GheckLinkItem(Item_ID As Long, _
                                        ByRef StoreID As Integer) As Boolean

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double


        sql = "SELECT     ItemID, StoreID"
sql = sql & " from dbo.TblLink_Item_To_Store_Details2"
sql = sql & "  Where (StoreID = " & StoreID & ") And (ItemID = " & Item_ID & ")"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
       GheckLinkItem = True

    Else
       GheckLinkItem = False
    End If

    rs.Close

End Function






Public Function ShowAttachments(TxtSerial1 As String, txtopeation_type As String, Optional ByVal mmIDD As String = "")
    If mmIDD = "" Then mmIDD = TxtSerial1
    If TxtSerial1 = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
         
            MsgBox "·«»œ „‰ «Õ Ì«— ”‰œ «Ê·«": Exit Function
        Else
            MsgBox "Select Voucher Firstly": Exit Function
        End If
 
    End If
    Dim mfrm As Form
    If SystemOptions.IsBlue Then
        Set mfrm = New imaged2
    Else
        Set mfrm = New imaged
    End If
    Set mfrm = New imaged
    
    imaged.SUBJECT_NO = (TxtSerial1)
    'imaged.mIDD = mmIDD
    Unload imaged
    imaged.show

    If SystemOptions.UserInterface = EnglishInterface Then

        imaged.Label9.Caption = "Voucher #"
        imaged.Caption = "Voucher Attachment"
         imaged.Label6.Caption = "Voucher #"
    Else

        imaged.Label9.Caption = "„—›ﬁ«  «·”‰œ    —ﬁ„"
        imaged.Caption = "„—›ﬁ«  «·”‰œ  "
       imaged.Label6.Caption = "—ﬁ„  «·”‰œ"

    End If
    
    imaged.SUBJECT_NO = (TxtSerial1)
  imaged.txtopeation_type = txtopeation_type
'imaged.mIDD = mmIDD
    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
    
 
 Dim Position As Integer

Position = InStr(1, TxtSerial1, "-")

If Position > 0 Then
Dim str1 As String
Dim str2 As String
Dim ConnectionStr As String
str1 = mId(TxtSerial1, 1, Position - 1)
str2 = mId(TxtSerial1, Position + 1, Len(TxtSerial1))
ConnectionStr = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type
ConnectionStr = ConnectionStr & "' and ( subject_no='" & (str1) & "'" & " or  subject_no='" & (TxtSerial1) & "') "
  imaged.Adodc1.RecordSource = ConnectionStr
  
Else
imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
End If



    imaged.Adodc1.Refresh




    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If
    
    If Position > 0 Then
ConnectionStr = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type
ConnectionStr = ConnectionStr & "' and ( subject_no='" & (str1) & "'" & " or  subject_no='" & (TxtSerial1) & "') "
  imaged.Adodc4.RecordSource = ConnectionStr

 
imaged.Adodc4.RecordSource = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
End If

imaged.Adodc4.RecordSource = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
DoEvents
imaged.Adodc4.Refresh
    imaged.Adodc4.Refresh
    imaged.DataGrid2.Refresh
    
End Function

Public Function ShowAttachments33(TxtSerial1 As String, txtopeation_type As String, Optional ByVal mmIDD As String = "")
    If mmIDD = "" Then mmIDD = TxtSerial1
    If TxtSerial1 = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then

            MsgBox "·«»œ „‰ «Õ Ì«— ”‰œ «Ê·«": Exit Function
        Else
            MsgBox "Select Voucher Firstly": Exit Function
        End If

    End If
    Dim mfrm As Form
    If SystemOptions.IsBlue Then
        Set mfrm = New imaged2
    Else
        Set mfrm = New imaged
    End If


    mfrm.SUBJECT_NO = (TxtSerial1)
    mfrm.mIDD = mmIDD
    Unload mfrm
    mfrm.show

    If SystemOptions.UserInterface = EnglishInterface Then

        mfrm.Label9.Caption = "Voucher #"
        mfrm.Caption = "Voucher Attachment"
         mfrm.Label6.Caption = "Voucher #"
    Else

        mfrm.Label9.Caption = "„—›ﬁ«  «·”‰œ    —ﬁ„"
        mfrm.Caption = "„—›ﬁ«  «·”‰œ  "
       mfrm.Label6.Caption = "—ﬁ„  «·”‰œ"

    End If

    mfrm.SUBJECT_NO = (TxtSerial1)
  mfrm.txtopeation_type = txtopeation_type
mfrm.mIDD = mmIDD
    mfrm.Adodc1.CommandType = adCmdText
    mfrm.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"


 Dim Position As Integer

Position = InStr(1, TxtSerial1, "-")

If Position > 0 Then
Dim str1 As String
Dim str2 As String
Dim ConnectionStr As String
str1 = mId(TxtSerial1, 1, Position - 1)
str2 = mId(TxtSerial1, Position + 1, Len(TxtSerial1))
ConnectionStr = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type
ConnectionStr = ConnectionStr & "' and ( subject_no='" & (str1) & "'" & " or  subject_no='" & (TxtSerial1) & "') "
  mfrm.Adodc1.RecordSource = ConnectionStr

Else
mfrm.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
End If



    'mfrm.Adodc1.Refresh




    If mfrm.Adodc1.Recordset.RecordCount > 0 Then

        mfrm.DBPix201.Visible = True
    Else
        mfrm.DBPix201.Visible = False
    End If

    If Position > 0 Then
ConnectionStr = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type
ConnectionStr = ConnectionStr & "' and ( subject_no='" & (str1) & "'" & " or  subject_no='" & (TxtSerial1) & "') "
  mfrm.Adodc4.RecordSource = ConnectionStr

Else
mfrm.Adodc4.RecordSource = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
End If
DoEvents
mfrm.Adodc4.Refresh
    mfrm.Adodc4.Refresh
    mfrm.DataGrid2.Refresh

End Function
Public Sub Translatefrm(Frm As Form)
    Set frmTranslations.Frm = Frm

    frmTranslations.show 1
End Sub


Public Function GetiItemsNewDetails(Optional uniteid As Integer = 0, Optional sizeid As Integer = 0, Optional ColorID As Integer = 0, Optional ClassId As Integer = 0, _
Optional ByRef UnitName As String, Optional ByRef sizename As String, Optional ByRef colorname As String, Optional ByRef classname As String)
  Dim sql As String
    Dim rs As New ADODB.Recordset

  sql = "select UnitID,UnitName,UnitNamee from TblUnites"

sql = sql & " WHERE     (UnitID = " & uniteid & ")"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
             If SystemOptions.UserInterface = ArabicInterface Then
                 UnitName = IIf(IsNull(rs("unitname").value), "", rs("unitname").value)
            Else
                UnitName = IIf(IsNull(rs("UnitNamee").value), "", rs("UnitNamee").value)
            End If

 Else
        UnitName = ""
    End If
    rs.Close


  sql = "select  *   from TblItemsSizes "

sql = sql & " WHERE     (SizeId = " & sizeid & ")"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
             If SystemOptions.UserInterface = ArabicInterface Then
                 sizename = IIf(IsNull(rs("sizename").value), "", rs("sizename").value)
            Else
                sizename = IIf(IsNull(rs("sizename").value), "", rs("sizename").value)
            End If

 Else
        sizename = ""
    End If
    rs.Close



  sql = "select  *   from TblItemsColors "

sql = sql & " WHERE     (ColorID = " & ColorID & ")"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
             If SystemOptions.UserInterface = ArabicInterface Then
                 colorname = IIf(IsNull(rs("ColorName").value), "", rs("ColorName").value)
            Else
                colorname = IIf(IsNull(rs("ColorName").value), "", rs("ColorName").value)
            End If

 Else
        colorname = ""
    End If
    rs.Close



  sql = "select  *   from TblItemsclasses "

sql = sql & " WHERE     (SizeId = " & ClassId & ")"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
             If SystemOptions.UserInterface = ArabicInterface Then
                 classname = IIf(IsNull(rs("SizeName").value), "", rs("SizeName").value)
            Else
                classname = IIf(IsNull(rs("SizeName").value), "", rs("SizeName").value)
            End If

 Else
        classname = ""
    End If
    rs.Close




End Function

Public Function GetGoldData(TTypeId As Integer, typeid As Integer, uniteid As Integer, _
Optional ByRef UnitName As String, Optional ByRef ttypename As String, Optional ByRef typename As String)
  Dim sql As String
    Dim rs As New ADODB.Recordset

  sql = "select UnitID,UnitName,UnitNamee from TblUnites"

sql = sql & " WHERE     (UnitID = " & uniteid & ")"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
             If SystemOptions.UserInterface = ArabicInterface Then
                 UnitName = IIf(IsNull(rs("unitname").value), "", rs("unitname").value)
            Else
                UnitName = IIf(IsNull(rs("UnitNamee").value), "", rs("UnitNamee").value)
            End If

 Else
        UnitName = ""
    End If
    rs.Close



  sql = "select   id,name,nameE  from TblGType "

sql = sql & " WHERE     (id = " & TTypeId & ")"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
             If SystemOptions.UserInterface = ArabicInterface Then
                 ttypename = IIf(IsNull(rs("name").value), "", rs("name").value)
            Else
                ttypename = IIf(IsNull(rs("nameE").value), "", rs("nameE").value)
            End If

 Else
        ttypename = ""
    End If
    rs.Close




End Function
Public Function loadmyModule()
On Error Resume Next
      Dim StrSQL As String
          Dim rs As New ADODB.Recordset
Dim ID As Integer
Dim Pid As Double
Dim code As Double
Dim i As Integer
       'menumenu

mdifrmmain.CeramicEstimation.Visible = False
  '    mdifrmmain.Reports.Visible = False
            mdifrmmain.AgeingMAster.Visible = False
      mdifrmmain.AssetsMngBase.Visible = False
      mdifrmmain.rsInvestment.Visible = False
      'mdifrmmain.planningMnu.Visible = False
      mdifrmmain.POSTRansactiosG.Visible = False
      mdifrmmain.SalesIns.Visible = False
      mdifrmmain.shipmentMnu.Visible = False
      mdifrmmain.ProductionPlan.Visible = False
      mdifrmmain.MnuElevators.Visible = False
        mdifrmmain.taxes.Visible = False
              mdifrmmain.hajMnu.Visible = False
                    mdifrmmain.TransporterMain.Visible = False
                          mdifrmmain.CarMaintenance.Visible = False
                                mdifrmmain.Strategy.Visible = False
                      mdifrmmain.MnuMaintnance.Visible = False



mdifrmmain.StudentMenue.Visible = False
      mdifrmmain.dev.Visible = False
       mdifrmmain.mangDep.Visible = False
       mdifrmmain.BankOp.Visible = False
       mdifrmmain.MnuElevators.Visible = False
       mdifrmmain.SalesIns.Visible = False
         mdifrmmain.rsInvestment.Visible = False
            mdifrmmain.MnuAccounts.Visible = False
            mdifrmmain.Currency.Visible = False
           mdifrmmain.LIFEINDICATORMNU.Visible = False
           mdifrmmain.COLLECTIONS.Visible = False
           mdifrmmain.Container.Visible = False
           mdifrmmain.RealEstateMarketing.Visible = False



mdifrmmain.FinAnalysis.Visible = False
 mdifrmmain.MNUFixedAssets.Visible = False
mdifrmmain.mnuEmployee.Visible = False
mdifrmmain.StockControl.Visible = False
    mdifrmmain.Purchase.Visible = False

 mdifrmmain.MarketingMnu.Visible = False
      mdifrmmain.hajMnu.Visible = False
            mdifrmmain.Sales.Visible = False
      mdifrmmain.shipmentMnu.Visible = False
  mdifrmmain.POSTRansactiosG.Visible = False
mdifrmmain.prdo.Visible = False
    mdifrmmain.ProductionPlan.Visible = False
mdifrmmain.MnuProjects.Visible = False
  mdifrmmain.TransporterMain.Visible = False
     mdifrmmain.CarMaintenance.Visible = False
mdifrmmain.MnuMaintnance.Visible = False
  mdifrmmain.Strategy.Visible = False
  mdifrmmain.Archiving.Visible = False
  mdifrmmain.LegalIssue.Visible = False
   mdifrmmain.Tailor.Visible = False
   mdifrmmain.rentcar.Visible = False
   mdifrmmain.Beauty.Visible = False

   If SystemOptions.SpecialVersion = True Then

mdifrmmain.AssetsMngReport(0).Visible = False
mdifrmmain.AssetsMngReport(8).Visible = False
      mdifrmmain.AssetsMngReport(14).Visible = False
      mdifrmmain.AssetsMng(2).Visible = False

   End If


   mdifrmmain.eye.Visible = False
   mdifrmmain.gobus.Visible = False
 'mdifrmmain.m2.Visible = False
'mdifrmmain.ArrowsBase.Visible = False
   mdifrmmain.AssetsMngBase.Visible = False
 mdifrmmain.Reports.Visible = False
  mdifrmmain.Tools.Visible = False
mdifrmmain.Basicdata.Visible = False
   mdifrmmain.dev.Visible = False
 'mdifrmmain.planningMnu.Visible = False
mdifrmmain.tech.Visible = True



'mdifrmmain.MnueHouseMain.Visible = False
'mdifrmmain.FarmerMnue.Visible = False
'mdifrmmain.GoldMenu.Visible = False
mdifrmmain.mangDep.Visible = False

mdifrmmain.xyz.Visible = False
mdifrmmain.Farm.Visible = False


'SystemOptions.Ecnomy = True

If SystemOptions.Ecnomy = True Then

With mdifrmmain

.MnuAccounts.Visible = True
.Currency.Visible = True
.MNUFixedAssets.Visible = True
.mnuEmployee.Visible = True
.StockControl.Visible = True
.Purchase.Visible = True
.Sales.Visible = True
.Help.Visible = True
.MnuToolsSetPrinters(0).Visible = True
.Basicdata.Visible = True
.Tools.Visible = True
.Reports.Visible = True


.StockControlBasicSub(4).Visible = False
.StockControlBasicSub(5).Visible = False
.StockControlBasicSub(6).Visible = False
.StockControlBasicSub(7).Visible = False
.StockControlBasicSub(8).Visible = False
.PurchaseBasic(3).Visible = False
.PurchaseBasic(4).Visible = False
.PurchaseBasic(4).Visible = False

.Expenses(0).Visible = False
.Expenses(1).Visible = False

.ExpensesSub(0).Visible = False
.ExpensesSub(1).Visible = False
.Cashing(1).Visible = False
.MnuBoxDrawing.Visible = False
.MNUFixedAssets.Visible = False
'.xxxxx(6).Visible = False
'.emptyMnu.Visible = False
.mnuEmployeeBasic(2).Visible = False
.mnuEmployeeBasic(3).Visible = False
.mnuEmployeeBasic(4).Visible = False
.mnuEmployeeBasic(5).Visible = False
.Vscstionsssub(0).Visible = False
.Vscstionsssub(1).Visible = False
.Vscstionsssub(2).Visible = False
.Vscstionsssub(5).Visible = False

.mnuEmployeeBasic(8).Visible = False

.StockControlBasicSub(12).Visible = False
.TradingTransaction(1).Visible = False
.TradingTransactionSub1(0).Visible = False
.TradingTransaction(7).Visible = False
.TradingTransaction(9).Visible = False
.TradingTransaction(10).Visible = False
.PurchaseBasic(1).Visible = False
.PurchaseTransactionssubs(0).Visible = False
.PurchaseTransactionssubs(2).Visible = False
.PurchaseTransactionssubs1(0).Visible = False
.PurchaseTransactions(1).Visible = False
.PurchaseTransactions(2).Visible = False

.SalesBasicSubsub(0).Visible = False
.SalesBasicSubsub(2).Visible = False
.SalesBasicSub(2).Visible = False
.SalesBasicSub(4).Visible = False
.SalesBasicSub(5).Visible = False
.SalesBasicSub(6).Visible = False
.SalesBasicSub(9).Visible = False
.SalesBasicSub(10).Visible = False
.SalesTransactionssubss00(0).Visible = False
.SalesTransactionssubss00(2).Visible = True
.SalesTransactionssubss000(0).Visible = False
.SalesTransactions(4).Visible = False
.SalesTransactions(5).Visible = False
.SalesTransactions(6).Visible = False
.SalesTransactions(8).Visible = False
.SalesTransactions(11).Visible = False
.SalesTransactions(12).Visible = False



 GoTo Lite


End With
End If

code = 10111982

         StrSQL = "SELECT *  From Pmanger "

    '    StrSQL = "SELECT branch_id,branch_namee From TblBranchesData"
      rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
       For i = 1 To rs.RecordCount
                    ID = IIf(IsNull(rs("id").value), "", rs("id").value)
                 Pid = IIf(IsNull(rs("Pid").value), "", rs("Pid").value)

             If ID = 1 Then
                 If Pid = i * i + code Then
                 mdifrmmain.MnuAccounts.Visible = True
                 Else
                  mdifrmmain.MnuAccounts.Visible = False
                 End If
         End If

         If ID = 2 Then
                If Pid = i * i + code Then
                 mdifrmmain.Currency.Visible = True
                 Else
                  mdifrmmain.Currency.Visible = False
                 End If
          End If

          If ID = 3 Then
                         If Pid = i * i + code Then
                 mdifrmmain.FinAnalysis.Visible = True
                 Else
                  mdifrmmain.FinAnalysis.Visible = False
                 End If
            End If

            If ID = 4 Then
               If Pid = i * i + code Then
                 mdifrmmain.MNUFixedAssets.Visible = True
                 Else
                  mdifrmmain.MNUFixedAssets.Visible = False
                 End If
                End If

                 If ID = 5 Then
                          If Pid = i * i + code Then
                 mdifrmmain.mnuEmployee.Visible = True
                 Else
                  mdifrmmain.mnuEmployee.Visible = False
                 End If
                 End If

                 If ID = 6 Then
                If Pid = i * i + code Then
                 mdifrmmain.StockControl.Visible = True
                 Else
                  mdifrmmain.StockControl.Visible = False
                 End If
                 End If

                 If ID = 7 Then
              If Pid = i * i + code Then
                 mdifrmmain.Purchase.Visible = True
                 Else
                  mdifrmmain.Purchase.Visible = False
                 End If
            End If


                  If ID = 8 Then
                  If Pid = i * i + code Then
                 mdifrmmain.MarketingMnu.Visible = True
                 Else
                  mdifrmmain.MarketingMnu.Visible = False
                 End If
                 End If

            If ID = 9 Then
              If Pid = i * i + code Then
                 mdifrmmain.Sales.Visible = True
                 Else
                  mdifrmmain.Sales.Visible = False
                 End If
             End If

             If ID = 10 Then
                      If Pid = i * i + code Then
                 mdifrmmain.shipmentMnu.Visible = True
                 Else
                  mdifrmmain.shipmentMnu.Visible = False
                 End If
          End If
             If ID = 11 Then
                            If Pid = i * i + code Then
                 mdifrmmain.POSTRansactiosG.Visible = True
                 Else
                  mdifrmmain.POSTRansactiosG.Visible = False
                 End If
                 End If

               If ID = 12 Then
                                    If Pid = i * i + code Then
                 mdifrmmain.prdo.Visible = True
                 Else
                  mdifrmmain.prdo.Visible = False
                 End If

              End If
              If ID = 13 Then
                                         If Pid = i * i + code Then
                 mdifrmmain.ProductionPlan.Visible = True
                 Else
                  mdifrmmain.ProductionPlan.Visible = False
                 End If
               End If

              If ID = 14 Then
                                                   If Pid = i * i + code Then
                 mdifrmmain.MnuProjects.Visible = True
                 Else
                  mdifrmmain.MnuProjects.Visible = False
                 End If
               End If
                 If ID = 15 Then
                   If Pid = i * i + code Then
                 mdifrmmain.TransporterMain.Visible = True
                 Else
                  mdifrmmain.TransporterMain.Visible = False
                 End If

               End If
               If ID = 16 Then
                                 If Pid = i * i + code Then
                 mdifrmmain.CarMaintenance.Visible = True
                 Else
                  mdifrmmain.CarMaintenance.Visible = False
                 End If
                End If

              If ID = 17 Then
                 If Pid = i * i + code Then
                 mdifrmmain.MnuMaintnance.Visible = True
                 Else
                  mdifrmmain.MnuMaintnance.Visible = False
                 End If
               End If

                 If ID = 18 Then
                 If Pid = i * i + code Then
                 mdifrmmain.Strategy.Visible = True
                 Else
                  mdifrmmain.Strategy.Visible = False
                 End If
                 End If
               If ID = 19 Then
                                 If Pid = i * i + code Then
                                mdifrmmain.Archiving.Visible = True
                                Else
                                 mdifrmmain.Archiving.Visible = False
                                End If
            End If

                 If ID = 20 Then
                             If Pid = i * i + code Then
                 mdifrmmain.StudentMenue.Visible = True
                 Else
                  mdifrmmain.StudentMenue.Visible = False
                 End If
                 End If
                 If ID = 21 Then
                                 If Pid = i * i + code Then
                ' mdifrmmain.ArrowsBase.Visible = True
                 Else
                  'mdifrmmain.ArrowsBase.Visible = False
                 End If
            End If

            If ID = 22 Then
                                            If Pid = i * i + code Then
                 mdifrmmain.AssetsMngBase.Visible = True
                 Else
                  mdifrmmain.AssetsMngBase.Visible = False
                 End If
           End If
           If ID = 23 Then
                    If Pid = i * i + code Then
                 mdifrmmain.Reports.Visible = True
                 Else
                  mdifrmmain.Reports.Visible = False
                 End If

            End If

            If ID = 24 Then
                   If Pid = i * i + code Then
                 mdifrmmain.Tools.Visible = True
                 Else

                  mdifrmmain.Tools.Visible = False
                 End If

                End If



           If ID = 25 Then
           If Pid = i * i + code Then
                 mdifrmmain.Basicdata.Visible = True
                 Else
                  mdifrmmain.Basicdata.Visible = False
                 End If
             End If

                '               If id = 26 And Pid = I * I + code Then
                ' mdifrmmain.Tech.Visible = True
                ' Else
                '  mdifrmmain.Tech.Visible = False
                ' End If

            If ID = 27 Then
            If Pid = i * i + code Then
                 mdifrmmain.dev.Visible = True
                 Else
                  mdifrmmain.dev.Visible = False
                 End If

            End If




 '                If ID = 28 Then
 '        If Pid = i * i + code Then
 '           '     mdifrmmain.MnueHouseMain.Visible = True
 '                Else
 '                 'mdifrmmain.MnueHouseMain.Visible = False
            '     End If
 '       End If

                        If ID = 28 Then
         If Pid = i * i + code Then
              mdifrmmain.Container.Visible = True
                 Else
                mdifrmmain.Container.Visible = False
                 End If
        End If

                 If ID = 29 Then
         If Pid = i * i + code Then
              mdifrmmain.COLLECTIONS.Visible = True
                 Else
                mdifrmmain.COLLECTIONS.Visible = False
                 End If
        End If
       'End If

                       If ID = 30 Then
         If Pid = i * i + code Then
                mdifrmmain.CeramicEstimation.Visible = True
                 Else
                  mdifrmmain.CeramicEstimation.Visible = False
                 End If
        End If


                               If ID = 31 Then
         If Pid = i * i + code Then
                mdifrmmain.RealEstateMarketing.Visible = True
                 Else
                 mdifrmmain.RealEstateMarketing.Visible = False
                 End If
        End If



                                    If ID = 32 Then
         If Pid = i * i + code Then
                 mdifrmmain.BankOp.Visible = True
                 Else
                  mdifrmmain.BankOp.Visible = False
                 End If
        End If




                                    If ID = 33 Then
         If Pid = i * i + code Then
                 mdifrmmain.mangDep.Visible = True
                 Else
                  mdifrmmain.mangDep.Visible = False
                 End If
        End If


                                            If ID = 34 Then
         If Pid = i * i + code Then
                 mdifrmmain.rsInvestment.Visible = True
                 Else
                  mdifrmmain.rsInvestment.Visible = False
                 End If
        End If


        If ID = 35 Then
                        If Pid = i * i + code Then
                                mdifrmmain.SalesIns.Visible = True
                                Else
                                 mdifrmmain.SalesIns.Visible = False
                                End If
        End If

             If ID = 36 Then
                        If Pid = i * i + code Then
                                mdifrmmain.MnuElevators.Visible = True
                                Else
                                 mdifrmmain.MnuElevators.Visible = False
                                End If
        End If

             If ID = 37 Then
                        If Pid = i * i + code Then
                                mdifrmmain.hajMnu.Visible = True
                                Else
                                 mdifrmmain.hajMnu.Visible = False
                                End If
        End If

     If ID = 38 Then
                        If Pid = i * i + code Then
                                mdifrmmain.LIFEINDICATORMNU.Visible = True
                                Else
                                 mdifrmmain.LIFEINDICATORMNU.Visible = False
                                End If
        End If


     If ID = 39 Then
                        If Pid = i * i + code Then
                                mdifrmmain.AgeingMAster.Visible = True
                                Else
                                 mdifrmmain.AgeingMAster.Visible = False
                                End If
        End If


     If ID = 40 Then
                        If Pid = i * i + code Then
                                   mdifrmmain.taxes.Visible = True
                                Else
                                   mdifrmmain.taxes.Visible = False
                                End If
        End If


     If ID = 41 Then
                        If Pid = i * i + code Then
                                mdifrmmain.LegalIssue.Visible = True
                                Else
                                   mdifrmmain.LegalIssue.Visible = False
                                End If
        End If


              If ID = 42 Then
                             If Pid = i * i + code Then
                                mdifrmmain.Tailor.Visible = True
                                Else
                                   mdifrmmain.Tailor.Visible = False
                                End If
        End If


              If ID = 43 Then
                             If Pid = i * i + code Then
                                mdifrmmain.rentcar.Visible = True
                                Else
                                   mdifrmmain.rentcar.Visible = False
                                End If
        End If


                     If ID = 44 Then
                             If Pid = i * i + code Then
                                mdifrmmain.Beauty.Visible = True
                                Else
                                   mdifrmmain.Beauty.Visible = False
                                End If
        End If


                     If ID = 45 Then
                             If Pid = i * i + code Then
                                mdifrmmain.eye.Visible = True
                                Else
                                   mdifrmmain.eye.Visible = False
                                End If
        End If



    If ID = 46 Then
                             If Pid = i * i + code Then
                                mdifrmmain.gobus.Visible = True
                                Else
                                   mdifrmmain.gobus.Visible = False
                                End If
        End If

                If ID = 47 Then
                             If Pid = i * i + code Then
                                mdifrmmain.xyz.Visible = True
                                Else
                                   mdifrmmain.xyz.Visible = False
                                End If
        End If


                               If ID = 48 Then
                             If Pid = i * i + code Then
                                mdifrmmain.Farm.Visible = True
                                Else
                                   mdifrmmain.Farm.Visible = False
                                End If
        End If



        'mdifrmmain.rentcar.Visible =False


  'LIFEINDICATORMNU

ll:


               rs.MoveNext


         Next i
    End If

    mdifrmmain.tech.Visible = True

rs.Close
Lite:
End Function

Public Function GetBrancheName(branch_id As Integer) As String
      Dim StrSQL As String
          Dim rs As New ADODB.Recordset

         StrSQL = "SELECT *  From TblBranchesData where branch_id=" & branch_id

    '    StrSQL = "SELECT branch_id,branch_namee From TblBranchesData"
      rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetBrancheName = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)

 Else
 GetBrancheName = ""
    End If




End Function






 Public Function GetIqarCode(Optional EmpCode As String, _
                                      Optional ByRef Emp_id As Variant, _
                                      Optional Emp_id1 As Double = 0, _
                                      Optional ByRef EmpCode1 As String, Optional ByRef ownerid As Variant)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    If Emp_id1 <> 0 Then
        sql = "select * from TblAqar where Aqarid= " & Emp_id1
    Else

        sql = "select * from TblAqar where  aqarNo ='" & EmpCode & "'"
    End If

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_id = IIf(IsNull(rs("Aqarid").value), 0, rs("Aqarid").value)
        EmpCode1 = IIf(IsNull(rs("aqarNo").value), 0, rs("aqarNo").value)
   ownerid = IIf(IsNull(rs("ownerid").value), 0, rs("ownerid").value)

    Else
        Emp_id = 0
    End If

    rs.Close

End Function

    Public Function GetIqarUnitData(ID As Long, Optional ByRef unitno As String, Optional ByRef meterPrice As Double, Optional ByRef Length As Double, Optional ByRef customerid As Integer, Optional ByRef rentType As Integer _
, Optional ByRef roomscount As Double, Optional ByRef LoungeCount As Double, Optional ByRef WCcount As Double, Optional ByRef account As Double, Optional ByRef kithchencount As Double, Optional ByRef ElectAccount As String, Optional MiniRentValue As Double, Optional ByRef Typed As Integer) As String

     Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

If ID = 0 Then Typed = 1: Exit Function
 sql = "SELECT     dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblAqarDetai.Id, dbo.TblAqarDetai.Aqarid, dbo.TblAqarDetai.length, dbo.TblAqarDetai.unitdesc, "
 sql = sql & "                    dbo.TblAqarDetai.Typed, dbo.TblAqarDetai.unitno, dbo.TblAqarDetai.unittype, dbo.TblAqarDetai.rentType, dbo.TblAqarDetai.meterPrice, dbo.TblAqarDetai.roomscount,"
sql = sql & "                      dbo.TblAqarDetai.WCcount, dbo.TblAqarDetai.kithchencount, dbo.TblAqarDetai.RentValue, dbo.TblAqarDetai.customerid, dbo.TblAqarDetai.haveFurniture,"
 sql = sql & "                     dbo.TblAqarDetai.namerentType, dbo.TblAqarDetai.Floor, dbo.TblAqarDetai.LoungeCount, dbo.TblAqarDetai.ACCount, dbo.TblAqarDetai.ACCountspleat,"
sql = sql & "                      dbo.TblAqarDetai.UnitElectric , dbo.TblAqarDetai.electric, dbo.TblAqarDetai.Water, dbo.TblAqarDetai.Services, dbo.TblAqarDetai.status , dbo.TblAqarDetai.MiniRentValue"
sql = sql & " FROM         dbo.TblAqarDetai LEFT OUTER JOIN"
sql = sql & "                      dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id"
sql = sql & "  WHERE     (dbo.TblAqarDetai.Id = " & ID & ")"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
     Typed = IIf(IsNull(rs("Typed").value), 1, rs("Typed").value) - 1
     If Typed = -1 Then Typed = 1
        unitno = IIf(IsNull(rs("unitno").value), "", rs("unitno").value)
                                    If SystemOptions.UserInterface = ArabicInterface Then
                                             UnittypeName = IIf(IsNull(rs("name").value), "", rs("name").value)
                                     Else
                                      UnittypeName = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                                     End If
        Length = IIf(IsNull(rs("length").value), 0, val(rs("length").value))
   rentType = IIf(IsNull(rs("rentType").value), 0, rs("rentType").value)
                    If rentType = 0 Then

                                 If SystemOptions.UserInterface = ArabicInterface Then
                                 rentTypeName = " «·ﬁÌ„… «·«ÌÃ«—Ì…"
                                 Else
                                 rentTypeName = "By Unit"
                                 End If
                    Else
                          If SystemOptions.UserInterface = ArabicInterface Then
                                 rentTypeName = " »«·„ — "
                                 Else
                                 rentTypeName = "By Meter"
                                 End If
                    End If
    MiniRentValue = IIf(IsNull(rs("MiniRentValue").value), 0, rs("MiniRentValue").value)
   ElectAccount = IIf(IsNull(rs("UnitElectric").value), "", rs("UnitElectric").value)
   meterPrice = IIf(IsNull(rs("meterPrice").value), 0, rs("meterPrice").value)
   roomscount = IIf(IsNull(rs("roomscount").value), 0, rs("roomscount").value)
   LoungeCount = IIf(IsNull(rs("LoungeCount").value), 0, rs("LoungeCount").value)
   WCcount = IIf(IsNull(rs("WCcount").value), 0, rs("WCcount").value)
   account = IIf(IsNull(rs("ACCount").value), 0, rs("ACCount").value)
   kithchencount = IIf(IsNull(rs("kithchencount").value), 0, rs("kithchencount").value)
   Length = IIf(IsNull(rs("length").value), 0, val(rs("length").value))

   GetIqarUnitData = ""
        If SystemOptions.UserInterface = ArabicInterface Then

  If Length <> 0 Then
GetIqarUnitData = "«·„”«Õ… " & Length & "  „ —" & vbNewLine
  End If

    If roomscount <> 0 Then
      GetIqarUnitData = GetIqarUnitData & roomscount & "  €—›… " & vbNewLine
  End If

    If LoungeCount <> 0 Then
  GetIqarUnitData = GetIqarUnitData & LoungeCount & "’«·Â" & vbNewLine
  End If

      If WCcount <> 0 Then


     GetIqarUnitData = GetIqarUnitData & WCcount & "Õ„«„" & vbNewLine
  End If

      If account <> 0 Then
    GetIqarUnitData = GetIqarUnitData & account & „ﬂ»› & vbNewLine

  End If

    Else

    End If
               End If
    rs.Close

 End Function



   Public Function GetTblCustemersCode(Optional EmpCode As String, _
                                      Optional ByRef Emp_id As Variant, _
                                      Optional Emp_id1 As Variant = 0, _
                                      Optional ByRef EmpCode1 As String, Optional Type1 As Integer = 1, Optional BranchID As Integer = 0)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    If Emp_id1 <> 0 Then
        sql = "select * from TblCustemers where CusID= " & Emp_id1
    Else
    sql = "select * from TblCustemers where  Fullcode ='" & EmpCode & "'"
 If Type1 = 2 Then
    sql = sql & " AND Type  = 2 "
   ElseIf Type1 = 1 Then
    sql = sql & " AND Type  = 1 "
  ElseIf Type1 = 56 Then
    sql = sql & " AND Type  = 56 "
ElseIf Type1 = 57 Then
    sql = sql & " AND Type  = 57 "
        End If
    End If


        If SystemOptions.usertype <> UserAdminAll Then
       '     StrSQL = StrSQL & " and   ( BranchId=0 or BranchId=" & Current_branch & ")  "
                       sql = sql & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"

        End If
             If BranchID <> 0 Then
         '   StrSQL = StrSQL & " and   BranchId=" & BranchID
                sql = sql & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"


        End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        Emp_id = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
        EmpCode1 = IIf(IsNull(rs("Fullcode").value), 0, rs("Fullcode").value)

    Else
        Emp_id = 0
    End If

    rs.Close

End Function





  Public Sub RetrivePoNo(Optional order_no As String = "", Optional ByRef PONo As String, Optional ByRef oorderdate As Date, Optional ByRef CBoBasedON As Integer)

    Dim StrSQL As String

    Dim rs As ADODB.Recordset

    'On Error GoTo ErrTrap
    StrSQL = "Select * from transactions  where    NoteSerial1='" & order_no & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
      PONo = IIf(IsNull(rs("PONo").value), "", rs("PONo").value)
        CBoBasedON = val(IIf(IsNull(rs("CBoBasedON").value), 0, rs("CBoBasedON").value))
      oorderdate = IIf(IsNull(rs("oorderdate").value), Date, rs("oorderdate").value)
      Else
      CBoBasedON = 0
       End If
End Sub

Public Function CheckNoteAdvancedPayments(NoteID As Double, Optional ByRef CusID As Long) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset

  sql = "SELECT   *"
sql = sql & " from dbo.notes"
sql = sql & " WHERE     (NoteID = " & NoteID & ")"

 Dim NCashingType As Integer
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        NCashingType = IIf(IsNull(rs("NCashingType").value), 0, rs("NCashingType").value)
        CusID = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
        If NCashingType = 3 Then
               CheckNoteAdvancedPayments = True
               Exit Function
               Else
               CheckNoteAdvancedPayments = False
        End If

 Else
        CheckNoteAdvancedPayments = False
        CusID = 0
    End If

End Function


Public Function GETNationality(ID As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset

  sql = "SELECT     NCODE, id"
sql = sql & " from dbo.Nationality"
sql = sql & " WHERE     (id = " & ID & ")"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GETNationality = IIf(IsNull(rs("NCODE").value), "", rs("NCODE").value)
 Else
        GETNationality = ""
    End If

End Function



Public Function GETlASTiSSUEDATE(Emp_id As Integer) As Date
    Dim sql As String
    Dim rs As New ADODB.Recordset

  sql = "SELECT     MAX(todate) AS MaxDate from dbo.TblEmpHolidaysDetails WHERE     (Emp_ID = " & Emp_id & ")"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
 GETlASTiSSUEDATE = IIf(IsNull(rs("MaxDate").value), Date, rs("MaxDate").value)
 Else
 GETlASTiSSUEDATE = Date
    End If

End Function
Public Function calcenaddate(StartDate As Date, interval As Integer, intervalvalindex As Integer) As Date
Dim intervalchar As String
If intervalvalindex = 0 Then
        intervalchar = "M"
 ElseIf intervalvalindex = 1 Then
      intervalchar = "YYYY"
ElseIf intervalvalindex = 2 Then
      intervalchar = "D"



Else
           intervalchar = "YYYY"
End If

calcenaddate = DateAdd(intervalchar, interval, StartDate)


End Function






Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, Notevalue As Double, DebitAccount As String, CreditAcc As String, des As String, NoteDate As Date, Optional debitvatacc As String, Optional Creditvatacc As String, Optional VATValue As Double)


    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer

 Dim StrSQL As String

         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords

 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰




    my_branch = BranchID


        StrTempAccountCode = DebitAccount


            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = "”‰œ   " & des & "   " & TxtNoteSerial1V
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If



If debitvatacc <> "" And VATValue > 0 Then


       StrTempAccountCode = debitvatacc
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, VATValue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If


End If



       StrTempAccountCode = CreditAcc
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue + VATValue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If




If Creditvatacc <> "" And VATValue > 0 Then


       StrTempAccountCode = Creditvatacc
   LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, VATValue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If


End If


ErrTrap:
End Function

Public Function showsforms(Index As Integer)
 Select Case Index
Case 0
Load FrmCarAuthontication
  FrmCarAuthontication.show

Case 2
         If checkApility("FrmOut") = False Then
                Exit Function
            End If


            FrmOut.show
            FrmOut.TxtTicketNO.Visible = True
            FrmOut.lbl(32).Visible = True
 Case 3
Load FrmBillCarMaintExtra
FrmBillCarMaintExtra.show
Case 4
Load FrmCarReporonlin
FrmCarReporonlin.show
Case 5
Load FrmCarReportsRequerNo
FrmCarReportsRequerNo.show
 Case 6
Load FrmBillComputerChek
FrmBillComputerChek.show
Case 7
 Load FrmOrderOpen
 FrmOrderOpen.show
 Case 8
 Load FrmCarReporonlin2
 FrmCarReporonlin2.show
Case 9

If SystemOptions.ShowBillCommisions = 0 Then Exit Function
Load FrmCommisRece
FrmCommisRece.show

Case 10

Load FrmCustemers
FrmCustemers.show
Case 11
If SystemOptions.ShowBillCommisions = 0 Then Exit Function


Load FrmCommisReport
 FrmCommisReport.show
Case 12
Load FrmCustemers
FrmCustemers.show
End Select

 End Function


Public Function SetPrinter2(PrnName As String)
    Dim Prn As Printer
    If Printers.count > 0 Then
        For Each Prn In Printers
            If Prn.DeviceName = PrnName Then
                Set Printer = Prn
                Exit For
            End If
        Next Prn
    End If
End Function

  Public Function createCustomer(CusName As String, CusNamee As String, Optional BranchID As Integer = 0, Optional ByRef CusID As Double, Optional ByVal Cus_mobile As String = "", Optional ByRef mCode As String = "") As Integer

   Dim RsTemp As New ADODB.Recordset
  Dim currentcode As String
Dim s As String, mPreFix As String


                StrSQL = "Select * From TblCustemers where CusName='" & CusNamee & "'"
                StrSQL = StrSQL & " or CusNamee='" & CusNamee & "'"


      RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then


                  createCustomer = 0
                  CusID = IIf(IsNull(RsTemp("CusID").value), 0, RsTemp("CusID").value)
                      Exit Function
                    End If




Dim ParentAccount As String
Dim parent_account As String
       Dim Account_Code_dynamic As String
            Account_Code_dynamic = get_account_code_branch(8, 1)

            If Account_Code_dynamic = "NO branch" Then
                MsgBox " ·« ÌÊÃœ —»ÿ Õ”«»«  ", vbCritical
              createCustomer = -1
              Exit Function
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "  ·« ÌÊÃœ —»ÿ Õ”«»«  ", vbCritical
       createCustomer = -1
              Exit Function
                End If
            End If
parent_account = Account_Code_dynamic




     CusID = CStr(new_id("TblCustemers", "CusID", "", True))



Dim Account_code As String
 Dim Account_code1 As String
 Dim Account_code2 As String

        ParentAccount = ""
        Account_code1 = ""
        Account_code2 = ""

          If SystemOptions.CustomerhavethreeAccounts = False Then
        Account_code = ModAccounts.AddNewAccount(parent_account, CusName, True, False, CusNamee)


'        parent_account


          Else

                    If SystemOptions.CustomerhavethreeAccounts = True Then
                        ParentAccount = ModAccounts.AddNewAccount(parent_account, CusName, False, False, CusNamee)
                        'rs("ParentAccount").value = ParentAccount

                        Account_code = ModAccounts.AddNewAccount(ParentAccount, CusName, True, False, CusNamee)
                        Account_code1 = ModAccounts.AddNewAccount(ParentAccount, CusName & "   ‘Ìﬂ«   Õ  «· Õ’Ì· ", True, False, CusNamee & "  Under Collection Cheque  ")
                        Account_code2 = ModAccounts.AddNewAccount(ParentAccount, CusName & "   „œ›Ê⁄«  „ﬁœ„…  ", True, False, CusName & " Advanced Payments")

                    Else
                        Account_code = ModAccounts.AddNewAccount(Account_Code_dynamic, CusName, True, False, CusNamee)
                      '  rs("ParentAccount").value = Null

                    End If

        End If


            If CStr(CusID) <> "" Then
                s = " SELECT Top 1  FIELD_no,prifix From Coding WHERE  FIELD_no = 4 and IsNull(prifix,'') <> ''"
                Dim rsDummy As New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    mPreFix = rsDummy!prifix & ""
                End If
                currentcode = get_coding(Current_branch, "TblCustemers", 4, mPreFix)
                mCode = mPreFix & CusID
            End If

        StrSQL = " insert into TblCustemers  (type,CusID,CusName,CusNameE,code,branchID,parent_account,Account_Code,Account_Code1,Account_Code2,ParentAccount,Cus_mobile,FullCode) "
        StrSQL = StrSQL & " VALUES (1," & CusID & ",'" & CusName & "' , '" & CusNamee & "' ,  '" & mCode & "'," & BranchID & ",'" & parent_account & "','" & Account_code & "','" & Account_code1 & "','" & Account_code2 & "','" & ParentAccount & "','" & Cus_mobile & "' ,'" & mCode & "')"

  Cn.Execute StrSQL

End Function







Public Function AutoSel(Cmb As ComboBox, KeyCode As Integer)

    Debug.Print KeyCode

    If KeyCode = vbEnter Then Exit Function
    If KeyCode = 8 Then Exit Function    'Backspace
    If KeyCode = 37 Then Exit Function  'left key
    If KeyCode = 38 Then Exit Function 'up arrow key
    If KeyCode = 39 Then Exit Function  'right key
    If KeyCode = 40 Then Exit Function  'down arrow key
    If KeyCode = 46 Then Exit Function  'delete key
    If KeyCode = 33 Then Exit Function  'page up key
    If KeyCode = 34 Then Exit Function  'page down key
    If KeyCode = 35 Then Exit Function  'end key
    If KeyCode = 36 Then Exit Function  'home key


    Dim Text As String
    Text = Cmb.Text

    Dim i As Long
    Dim Temp As String


    For i = 0 To Cmb.ListCount - 1
        Temp = left(Cmb.List(i), Len(Text))
        If LCase(Temp) = LCase(Text) Then
            Cmb.Text = Cmb.List(i)
            Cmb.ListIndex = i
            Cmb.SelStart = Len(Text)
            Cmb.SelLength = Len(Cmb.List(i))
            'Cmb.SetFocus
        End If
    Next

End Function
Public Function GetWeekdayName(DayNO As Integer) As String

    If SystemOptions.UserInterface = ArabicInterface Then

        Select Case DayNO

            Case 1
                GetWeekdayName = "«·”» "

            Case 2
                GetWeekdayName = "«·«Õœ"

            Case 3
                GetWeekdayName = "«·«À‰Ì‰"

            Case 4
                GetWeekdayName = "«·À·«À«¡"

            Case 5
                GetWeekdayName = "«·«—»⁄«¡"

            Case 6
                GetWeekdayName = "«·Œ„Ì”"

            Case 7
                GetWeekdayName = "«·Ã„⁄Â"

        End Select

    Else

        Select Case DayNO

            Case 1
                GetWeekdayName = "Saturday"

            Case 2
                GetWeekdayName = "Sunday"

            Case 3
                GetWeekdayName = "Monday"

            Case 4
                GetWeekdayName = "Tuesday"

            Case 5
                GetWeekdayName = "Wednesday"

            Case 6
                GetWeekdayName = "Thursday"

            Case 7
                GetWeekdayName = "Friday"

        End Select

    End If

End Function

Function MoveUpDown(ByRef List As ListBox, upDown As Integer)
 Dim currentpos As Integer
Dim currentname As String
 Dim BEFOREPOSTION  As Integer
Dim BEFORENAME As String

 Dim AfterPOSTION  As Integer
Dim AfterNAME As String

If upDown = 0 Then 'up
currentpos = List.ListIndex
currentname = List.List(currentpos)

If currentpos = 0 Then Exit Function

BEFOREPOSTION = List.ListIndex - 1
BEFORENAME = List.List(BEFOREPOSTION)


List.List(BEFOREPOSTION) = currentname

List.List(currentpos) = BEFORENAME
List.ListIndex = BEFOREPOSTION
Else

currentpos = List.ListIndex
currentname = List.List(currentpos)

If currentpos = List.ListCount - 1 Then Exit Function

AfterPOSTION = List.ListIndex + 1
AfterNAME = List.List(AfterPOSTION)


List.List(AfterPOSTION) = currentname

List.List(currentpos) = AfterNAME

List.ListIndex = AfterPOSTION


End If

End Function




Public Function saveApprovalData(Transactionid As Double, _
                                 Transaction_Type As Double, _
                                 NoteSerial As Double, _
                                 frmname As String)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer

    sql = "SELECT     dbo.TblApprovalDefDetails.PlainMessageID, dbo.TbLLevels.Name, dbo.TbLLevels.Namee, dbo.TbllevelWorker.EmpID"
    sql = sql & "  FROM         dbo.TbllevelWorker INNER JOIN"
    sql = sql & "  dbo.TbLLevels ON dbo.TbllevelWorker.LevelID = dbo.TbLLevels.LevelID INNER JOIN"
    sql = sql & "  dbo.TblApprovalDef INNER JOIN"
    sql = sql & "  dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID ON"
    sql = sql & "  dbo.TbLLevels.LevelID = dbo.TblApprovalDefDetails.PlainMessageID"
    sql = sql & "  WHERE     (dbo.TblApprovalDef.ScreenName = N'" & frmname & "')"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Cn.Execute "delete TblTransactionsApproval where Transaction_Type=" & Transaction_Type & " and Transactionid=" & Transactionid

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount

            If Not (IsNull(rs("PlainMessageID").value)) Then
                sql = "insert into TblTransactionsApproval (Transaction_Type,NoteSerial,Level,Transactionid,CurrUserID,UserID)  "
                sql = sql & "Values (" & Transaction_Type & "," & NoteSerial & "," & rs("PlainMessageID").value & "," & Transactionid & "," & user_id & "," & rs("empID").value & ")"
                Cn.Execute sql
            End If

            rs.MoveNext
        Next i

    End If

End Function


Public Function getInfoMessage1(ID As Integer, _
                                               Optional ByRef Name As String, _
                                               Optional ByRef speed As Double, _
                                               Optional ByRef fontsize As Double, _
                                               Optional ByRef fontcolor As Double, _
                                               Optional ByRef backcolor As Double, _
                                               Optional ByRef show As Boolean)


 On Error Resume Next

    Dim SQL1 As String
    Dim Rs1 As New ADODB.Recordset


     SQL1 = " SELECT   * from InfoSettings1   "
      SQL1 = SQL1 + "  WHERE     (" & SQLDate(Date, True) & " BETWEEN dbo.InfoSettings1.StartDate AND dbo.InfoSettings1.EndDate)  "


     '        Sql1 = Sql1 + " where  (startdate >=" & SQLDate(Date, True) & ""



     '   Sql1 = Sql1 + " and enddate <=" & SQLDate(Date, True) & ""
     '    Sql1 = Sql1 + "and CAST(StartTime As Time) >= CAST(CURDATE()() As Time) "
     '  Sql1 = Sql1 + "and  CAST(enddate As Time) <= CAST(CURDATE()() As Time) "
     Rs1.Open SQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            If 1 = 1 Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                     Name = IIf(IsNull(Rs1("name").value), "", Rs1("name").value)
                Else
                   Name = IIf(IsNull(Rs1("namee").value), "", Rs1("namee").value)
           End If

                speed = IIf(IsNull(Rs1("speed").value), 50, Rs1("speed").value)
                fontsize = IIf(IsNull(Rs1("fontsize").value), 12, Rs1("fontsize").value)
                fontcolor = IIf(IsNull(Rs1("fontcolor").value), 255, Rs1("fontcolor").value)
                backcolor = IIf(IsNull(Rs1("backcolor").value), 0, Rs1("backcolor").value)
               show = True
                WebForm.info1Timer.interval = speed
              Else
              show = False
              End If
    Else
      show = False
    End If

End Function
Public Function getInfoMessage(ID As Integer, _
                                               Optional ByRef Name As String, _
                                               Optional ByRef speed As Double, _
                                               Optional ByRef fontsize As Double, _
                                               Optional ByRef fontcolor As Double, _
                                               Optional ByRef backcolor As Double, _
                                               Optional ByRef show As Boolean)


 On Error Resume Next

    Dim SQL1 As String
    Dim Rs1 As New ADODB.Recordset
     SQL1 = " SELECT   * from InfoSettings1   "

             SQL1 = SQL1 + " where  (startdate >=" & SQLDate(Date, True) & ""



        SQL1 = SQL1 + " and enddate <=" & SQLDate(Date, True) & ""
         SQL1 = SQL1 + "and CAST(StartTime As Time) >= CAST(CURDATE()() As Time) "
       SQL1 = SQL1 + "and  CAST(enddate As Time) <= CAST(CURDATE()() As Time) "



'   rs1.Open sql1, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs1.RecordCount > 0 Then



'   End If


    Dim sql As String
    Dim rs As New ADODB.Recordset
     sql = " SELECT   * from InfoSettings  "
   rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
            If rs("Show").value = True Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                     Name = IIf(IsNull(rs("name").value), "", rs("name").value)
                Else
                   Name = IIf(IsNull(rs("namee").value), "", rs("namee").value)
           End If

                speed = IIf(IsNull(rs("speed").value), 50, rs("speed").value)
                fontsize = IIf(IsNull(rs("fontsize").value), 12, rs("fontsize").value)
                fontcolor = IIf(IsNull(rs("fontcolor").value), 255, rs("fontcolor").value)
                backcolor = IIf(IsNull(rs("backcolor").value), 0, rs("backcolor").value)
               show = True
                WebForm.Timer2.interval = speed
              Else
              show = False
              End If
    Else
      show = False
    End If

End Function


Public Function AddTofaforites(Optional formname As String, Optional Displayname As String _
, Optional Displaynamee As String)
'On Error Resume Next
Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Noofmenues As Double
    'Dim TimeCateg As Double
 Dim str As String
 If formname = "" Then Exit Function
    sql = "SELECT     COUNT(id) AS Noofmenues"
sql = sql & " from dbo.TblMyMenue"
sql = sql & " WHERE      userid=" & user_id

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
                Noofmenues = IIf(IsNull(rs("Noofmenues").value), 0, rs("Noofmenues").value)

    Else
                Noofmenues = 0
    End If

   rs.Close
'CHECK FORM NAME NOT EXIST BEFORE
    sql = "SELECT     *  "
sql = sql & " from dbo.TblMyMenue"
sql = sql & " WHERE      userid=" & user_id & " and  formname='" & formname & "'"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText


    If rs.RecordCount > 0 Then
                               If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« Ì„ﬂ‰ «÷«›… Â–… «·‘«‘… „ÊÃÊœ… »«·›⁄· ›Ì «·„›÷·«  ", vbInformation
                     Else
                     MsgBox "can't Add to Favorites It's Already Exist ", vbInfor, mation
                     End If

Exit Function

     End If



   If Noofmenues <= 30 Then


                        str = "insert into  TblMyMenue   (  USERID,formname,Displayname,Displaynamee) "
                         str = str & "values( " & user_id & ",'" & formname & "','" & Displayname & "','" & Displaynamee & "'  )"
                         Cn.Execute str
   Else
                       If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« Ì„ﬂ‰ «÷«›… «Œ—Ì «·Ï «·„›÷·«  ", vbInformation
                     Else
                     MsgBox "can't Add to Favorites ", vbInformation
                     End If


   End If
                               If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „  «·«÷«›… ··„›÷·Â ", vbInformation
                     Else
                     MsgBox "Added to Favorites Success ", vbInformation
                     End If

Call mdifrmmain.showFavoritesMenue
End Function

Public Function CheckLastApprovLevel(Optional ScreenName As String, _
Optional Transaction_ID As Double = 0, Optional NoteID As Double = 0) As Double
     Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim NoOfMinute As Double
    'Dim TimeCateg As Double

    sql = "SELECT     COUNT(id) AS NotApproved"
sql = sql & " from dbo.ApprovalData"
sql = sql & " WHERE    empid <>0  and   (ScreenName = N'" & ScreenName & "')  AND (ApprovDate IS NULL)"
If Transaction_ID <> 0 Then
sql = sql & " AND (Transaction_ID = " & Transaction_ID & ")   "
End If

 If NoteID <> 0 Then
sql = sql & " AND (NoteID = " & NoteID & ")   "
End If

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        CheckLastApprovLevel = IIf(IsNull(rs("NotApproved").value), 0, rs("NotApproved").value)

    Else
    CheckLastApprovLevel = 0
    End If


End Function

Public Function GetTimeforTransaction(Optional ScreenName As String, _
Optional ByRef TimeCateg As Double) As Double

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim NoOfMinute As Double
    'Dim TimeCateg As Double

    sql = "SELECT     ScreenName, timeCount, TimeCateg"
sql = sql & " From dbo.TblApprovalDef"
sql = sql & " WHERE     (ScreenName = N'" & ScreenName & "') "



    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        TimeCateg = IIf(IsNull(rs("TimeCateg").value), 0, rs("TimeCateg").value)
        NoOfMinute = IIf(IsNull(rs("timeCount").value), 0, rs("timeCount").value)
            If TimeCateg = 0 Then
                     NoOfMinute = NoOfMinute * 1
            ElseIf TimeCateg = 1 Then
                     NoOfMinute = NoOfMinute * 60
            ElseIf TimeCateg = 2 Then
                   NoOfMinute = NoOfMinute * 60 * 24
            End If


    Else
    NoOfMinute = 0
    End If
   GetTimeforTransaction = NoOfMinute

End Function


Public Function GetlastPurchasedata(Transaction_Type As Double, Item_ID As Double _
                                               , FromDate As Date, ToDate As Date, Optional ByRef LastPurchaseDate As String _
                                          , Optional ByRef LastPrice As Double, Optional ByRef lastQty As Double)

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = " SELECT     TOP 100 PERCENT MAX(dbo.Transactions.Transaction_Date) AS lastPurchaseDate, dbo.Transaction_Details.showPrice AS lastPrice, "
 sql = sql & " dbo.Transaction_Details.ShowQty AS lastQty"
 sql = sql & "  FROM         dbo.Transactions INNER JOIN"
sql = sql & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
sql = sql & "   Where (dbo.Transactions.Transaction_Type = " & Transaction_Type & ") And (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
sql = sql & "   GROUP BY dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ShowQty"
sql = sql & "   HAVING      (MAX(dbo.Transactions.Transaction_Date) >= " & SQLDate(FromDate, True) & " AND MAX(dbo.Transactions.Transaction_Date)"
sql = sql & "    <= " & SQLDate(ToDate, True) & ")"
sql = sql & "   ORDER BY MAX(dbo.Transactions.Transaction_Date) DESC"





    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        LastPurchaseDate = IIf(IsNull(rs("lastPurchaseDate").value), "", rs("lastPurchaseDate").value)
 LastPrice = Round(IIf(IsNull(rs("lastPrice").value), 0, rs("lastPrice").value), 2)
 lastQty = Round(IIf(IsNull(rs("lastQty").value), 0, rs("lastQty").value), 2)


    Else
     LastPrice = 0
 lastQty = 0
 LastPurchaseDate = ""

    End If
rs.Close
Set rs = Nothing
End Function
Public Function checkmanyStores(Optional ByRef str As String = "") As Boolean

    Dim sql As String
    Dim rs As New ADODB.Recordset
If SystemOptions.UserInterface = ArabicInterface Then
  sql = "SELECT     dbo.TblUsersStores.StoreID, dbo.TblStore.StoreName "
Else
sql = "SELECT     dbo.TblUsersStores.StoreID, dbo.TblStore.StoreNamee "
End If

sql = sql & "  FROM         dbo.TblUsersStores LEFT OUTER JOIN"
sql = sql & "  dbo.TblStore ON dbo.TblUsersStores.StoreID = dbo.TblStore.StoreID"
If user_id <> 1 Then
sql = sql & "    Where (dbo.TblUsersStores.userid = " & user_id & ")"
Else
   checkmanyStores = False
   Exit Function
  End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
str = sql
        checkmanyStores = True


    Else
   checkmanyStores = False
    End If

End Function


Public Function checkmanyBranches(Optional ByRef str As String = "") As Boolean

    Dim sql As String
    Dim rs As New ADODB.Recordset
If SystemOptions.UserInterface = ArabicInterface Then
  sql = "SELECT     dbo.TblUsersBranches.BranchID, dbo.TblBranchesData.branch_name "
Else
sql = "SELECT     dbo.TblUsersBranches.BranchID, dbo.TblBranchesData.branch_namee "
End If

sql = sql & "   FROM         dbo.TblUsersBranches INNER JOIN"
sql = sql & "   dbo.TblBranchesData ON dbo.TblUsersBranches.BranchID = dbo.TblBranchesData.branch_id"
If user_id <> 1 Then
sql = sql & "    Where (dbo.TblUsersBranches.UserID = " & user_id & ")"
Else
   checkmanyBranches = False
   Exit Function
  End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
str = sql
        checkmanyBranches = True


    Else
   checkmanyBranches = False
    End If

End Function
Public Function GetYearlyAverage(Transaction_Type As Double, _
        Item_ID As Double, FromDate As Date, ToDate As Date, Optional ByRef GetYearlyAverage1 As Double)

    Dim sql As String
    Dim rs As New ADODB.Recordset

 '   sql = " SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.ShowQty) / COUNT(DISTINCT MONTH(dbo.Transactions.Transaction_Date)) AS AverageMonthly "
'sql = sql & "   FROM         dbo.Transactions INNER JOIN"
'sql = sql & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'sql = sql & "    WHERE     (dbo.Transactions.Transaction_Type = " & Transaction_Type & ") AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND (dbo.Transactions.Transaction_Date >=  " & SQLDate(fromdate, True) & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True) & ")"

  sql = " SELECT    SUM(dbo.Transaction_Details.ShowQty) / 1 as YearlyAverage"
sql = sql & "    FROM         dbo.Transactions INNER JOIN"
sql = sql & "     dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
sql = sql & "   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
sql = sql & "   WHERE     (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND (dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True) & ") AND (dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & ")"
sql = sql & "    AND (dbo.TransactionTypes.StockEffect = - 1)"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetYearlyAverage1 = Round(IIf(IsNull(rs("YearlyAverage").value), 0, rs("YearlyAverage").value), 2)


    Else
   GetYearlyAverage1 = 0
    End If

End Function



Public Function GetMonthlyAverage(Transaction_Type As Double, _
        Item_ID As Double, FromDate As Date, ToDate As Date, Optional ByRef AverageMonthly As Double)

    Dim sql As String
    Dim rs As New ADODB.Recordset

 '   sql = " SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.ShowQty) / COUNT(DISTINCT MONTH(dbo.Transactions.Transaction_Date)) AS AverageMonthly "
'sql = sql & "   FROM         dbo.Transactions INNER JOIN"
'sql = sql & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'sql = sql & "    WHERE     (dbo.Transactions.Transaction_Type = " & Transaction_Type & ") AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND (dbo.Transactions.Transaction_Date >=  " & SQLDate(fromdate, True) & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True) & ")"

  sql = " SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.ShowQty) / COUNT(DISTINCT MONTH(dbo.Transactions.Transaction_Date)) AS AverageMonthly"
sql = sql & "    FROM         dbo.Transactions INNER JOIN"
sql = sql & "     dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
sql = sql & "   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
sql = sql & "   WHERE     (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND (dbo.Transactions.Transaction_Date >= " & SQLDate(FromDate, True) & ") AND (dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & ")"
sql = sql & "    AND (dbo.TransactionTypes.StockEffect = - 1)"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        AverageMonthly = Round(IIf(IsNull(rs("AverageMonthly").value), 0, rs("AverageMonthly").value), 2)


    Else
   AverageMonthly = 0
    End If

End Function



Public Function GetempDepartementidFromUserid(UserID As Double) As Double

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = " SELECT     DepartmentID"
sql = sql & "  From dbo.TblEmployee"
sql = sql & "   Where (Emp_id = " & GetempidFromUserid(UserID) & ")"


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetempDepartementidFromUserid = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)


    Else
   GetempDepartementidFromUserid = 0
    End If

End Function

Public Function GetempidFromUserid(UserID As Double) As Double

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = "  select * from TblUsers where UserID =" & UserID


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetempidFromUserid = IIf(IsNull(rs("Empid").value), 0, rs("Empid").value)


    Else
   GetempidFromUserid = 0
    End If

End Function

Public Function GetCurrentApprovalForTransactions(Transaction_ID As Double, _
                                               Optional ScreenName As String) As Double

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = " SELECT    MIN(id) AS MinId"
 sql = sql & " from dbo.ApprovalData"
 sql = sql & " WHERE  empid <>0 and     (ApprovDate IS NULL and CancelApprove IS NULL) AND (Transaction_ID = " & Transaction_ID & ") AND (ScreenName = N'" & ScreenName & "')"
 sql = sql & " ORDER BY MIN(id)  "


    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetCurrentApprovalForTransactions = IIf(IsNull(rs("MinId").value), 0, rs("MinId").value)


    Else
   GetCurrentApprovalForTransactions = 0
    End If

End Function
Public Function getClassInformations(ID As Integer, _
                                               Optional ByRef Name As String, _
   Optional ByRef DiscountPercentage As Double, _
   Optional ByRef PerfectPercentage As Double _
   , Optional ByRef Account_code As String)

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = " SELECT   * from TblItemsclasses  "

    sql = sql & " Where (SizeId = " & ID & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Name = IIf(IsNull(rs("SizeName").value), "", rs("SizeName").value)
   Account_code = IIf(IsNull(rs("Account_code").value), "", rs("Account_code").value)
        DiscountPercentage = IIf(IsNull(rs("DiscountPercentage").value), 0, rs("DiscountPercentage").value)
    PerfectPercentage = IIf(IsNull(rs("PerfectPercentage").value), 0, rs("PerfectPercentage").value)

    Else
        Name = ""

        DiscountPercentage = 0
    End If
   getClassInformations = DiscountPercentage
End Function


Public Function getMaintenancetypeInformations(ID As Integer, _
                                               Optional ByRef Name As String, _
                                               Optional ByRef km As String, _
                                               Optional ByRef Remarks As String, _
                                               Optional ByRef alarmBfore As Double)

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = " SELECT   * from MaintenanceTypes  "

    sql = sql & " Where (id = " & ID & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Name = IIf(IsNull(rs("name").value), "", rs("name").value)
        km = IIf(IsNull(rs("km").value), "", rs("km").value)
        Remarks = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
        alarmBfore = IIf(IsNull(rs("alarmBfore").value), 0, rs("alarmBfore").value)

    Else
        Name = ""
        km = ""
        Remarks = ""
        alarmBfore = 0
    End If

End Function

Public Function CreateLogo(xRport As CRAXDRT.Report, _
                           Optional BranchID As Double = 0, _
                           Optional ByVal StrN As String = "") As Boolean
    Dim rs          As ADODB.Recordset
    Dim BolShowLogo As Boolean
    Dim xLogo       As CRAXDRT.OLEObject
    Dim StrFileName As String
    Dim MsgErr      As String
    Dim StrSQL      As String
    On Error GoTo hErr

    Set rs = New ADODB.Recordset
    If SystemOptions.WorkWithBranchLogo = False Then
        rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    Else

        If BranchID = 0 Then
            StrSQL = "SELECT     *  from TblBranchesData Where (branch_id = " & Current_branch & ")"
        Else
            StrSQL = "SELECT     *  from TblBranchesData Where (branch_id = " & BranchID & ")"

        End If
        rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    End If

    If rs.BOF Or rs.EOF Then
        CreateLogo = False
        Exit Function
    End If

    BolShowLogo = IIf(IsNull(rs("ShowLogoInReports").value), 0, rs("ShowLogoInReports").value)

    If BolShowLogo = True And hide_logo = False Then
        If SystemOptions.WorkWithBranchLogo = False Then
            LoadPictureFromDB Nothing, rs, "CompanyLogo", StrFileName
        Else
            LoadPictureFromDB Nothing, rs, "branchLogo", StrFileName

        End If

        If StrN <> "" Then StrFileName = StrN

        Set xLogo = xRport.Areas(1).Sections(1).AddPictureObject(StrFileName, 100, 100)
        xLogo.Width = SystemOptions.logowidth
        xLogo.Height = SystemOptions.logoHeight
        xLogo.backcolor = vbWhite
        xLogo.BorderColor = 255
        xLogo.CloseAtPageBreak = True
        xLogo.HyperlinkText = "BYTE"
        xLogo.HyperlinkType = crHyperlinkWebsite
        xRport.Areas(1).Sections(1).SuppressIfBlank = True
        xRport.Areas(1).Sections(1).Height = xLogo.Height + 250
        CreateLogo = True
    Else
        CreateLogo = False
    End If

    rs.Close
    Set rs = Nothing
    Exit Function
hErr:
    MsgErr = "Œÿ« ›Ï "
    MsgErr = MsgErr & CHR(13) & "CreateLogo"
    MsgErr = MsgErr & CHR(13) & Err.Description
    MsgErr = MsgErr & CHR(13) & Err.Number
    MsgErr = MsgErr & CHR(13) & Err.Source
    WriteInLogFile MsgErr
    CreateLogo = False
End Function

Public Function getitemAgeingData(FromDate As Date, _
                                  ToDate As Date, _
                                  Optional GroupID As Integer = 0, _
                                  Optional Item_ID As Integer)

    Dim late_interval As Integer
    Dim ItemID As Long
    Dim column_location As Integer
    Dim i As Integer
    Dim sql As String
    Dim Rs3 As New ADODB.Recordset

    Dim QtyBalance As Double
    Dim BalanceValue As Double
    Dim AvgCost As Double

    Cn.Execute "DELETE FROM TblTempItemAging"

    sql = ""
    sql = sql & " ;WITH LastMove AS ( "
    sql = sql & "     SELECT "
    sql = sql & "         TD.Item_ID, "
    sql = sql & "         MAX(T.Transaction_Date) AS LastDate "
    sql = sql & "     FROM dbo.Transaction_Details TD "
    sql = sql & "     INNER JOIN dbo.Transactions T "
    sql = sql & "         ON TD.Transaction_ID = T.Transaction_ID "
    sql = sql & "     INNER JOIN dbo.TblItems I "
    sql = sql & "         ON TD.Item_ID = I.ItemID "
    sql = sql & "     WHERE T.Transaction_Type = 21 "
    sql = sql & "       AND T.Transaction_Date >= " & SQLDate(FromDate, True)
    sql = sql & "       AND T.Transaction_Date <= " & SQLDate(ToDate, True)

    If GroupID <> 0 Then
        sql = sql & " AND I.GroupID = " & GroupID
    End If

    If Item_ID <> 0 Then
        sql = sql & " AND TD.Item_ID = " & Item_ID
    End If

    sql = sql & "     GROUP BY TD.Item_ID "
    sql = sql & " ), StockAgg AS ( "
    sql = sql & "     SELECT "
    sql = sql & "         TD.Item_ID, "
    sql = sql & "         SUM(ISNULL(TD.Quantity,0) * ISNULL(TT.StockEffect,0)) AS QtyBalance, "
    sql = sql & "         SUM(ROUND(ISNULL(TD.Quantity,0) * ISNULL(TD.Price,0) * ISNULL(TT.StockEffect,0), 2)) AS BalanceValue "
    sql = sql & "     FROM dbo.Transaction_Details TD "
    sql = sql & "     INNER JOIN dbo.Transactions T "
    sql = sql & "         ON TD.Transaction_ID = T.Transaction_ID "
    sql = sql & "     INNER JOIN dbo.TransactionTypes TT "
    sql = sql & "         ON T.Transaction_Type = TT.Transaction_Type "
    sql = sql & "     INNER JOIN dbo.TblItems I "
    sql = sql & "         ON TD.Item_ID = I.ItemID "
    sql = sql & "     WHERE ISNULL(TT.StockEffect,0) <> 0 "
    sql = sql & "       AND T.Transaction_Date <= " & SQLDate(ToDate, True)

    If GroupID <> 0 Then
        sql = sql & " AND I.GroupID = " & GroupID
    End If

    If Item_ID <> 0 Then
        sql = sql & " AND TD.Item_ID = " & Item_ID
    End If

    sql = sql & "     GROUP BY TD.Item_ID "
    sql = sql & " ) "
    sql = sql & " SELECT "
    sql = sql & "     LM.Item_ID, "
    sql = sql & "     LM.LastDate, "
    sql = sql & "     DATEDIFF(DAY, LM.LastDate, " & SQLDate(ToDate, True) & ") AS DIFFerents, "
    sql = sql & "     ISNULL(SA.QtyBalance,0) AS QtyBalance, "
    sql = sql & "     ISNULL(SA.BalanceValue,0) AS BalanceValue, "
    sql = sql & "     CASE "
    sql = sql & "         WHEN ISNULL(SA.QtyBalance,0) = 0 THEN 0 "
    sql = sql & "         ELSE ISNULL(SA.BalanceValue,0) / SA.QtyBalance "
    sql = sql & "     END AS AvgCost "
    sql = sql & " FROM LastMove LM "
    sql = sql & " LEFT JOIN StockAgg SA "
    sql = sql & "     ON LM.Item_ID = SA.Item_ID "
    sql = sql & " ORDER BY LM.Item_ID "

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount = 0 Then
        Rs3.Close
        Exit Function
    End If

    Rs3.MoveFirst

    For i = 1 To Rs3.RecordCount

        late_interval = IIf(IsNull(Rs3.Fields("DIFFerents").value), 0, Rs3.Fields("DIFFerents").value)
        ItemID = IIf(IsNull(Rs3.Fields("Item_ID").value), 0, Rs3.Fields("Item_ID").value)
        QtyBalance = IIf(IsNull(Rs3.Fields("QtyBalance").value), 0, Rs3.Fields("QtyBalance").value)
        BalanceValue = IIf(IsNull(Rs3.Fields("BalanceValue").value), 0, Rs3.Fields("BalanceValue").value)
        AvgCost = IIf(IsNull(Rs3.Fields("AvgCost").value), 0, Rs3.Fields("AvgCost").value)

        column_location = get_late_location2(late_interval)

        Cn.Execute "INSERT INTO TblTempItemAging (ItemID, LateID, QtyBalance, BalanceValue, AvgCost) VALUES (" & _
                   ItemID & "," & _
                   column_location & "," & _
                   Replace(CStr(QtyBalance), ",", ".") & "," & _
                   Replace(CStr(BalanceValue), ",", ".") & "," & _
                   Replace(CStr(AvgCost), ",", ".") & ")"

        Rs3.MoveNext
    Next i

    Rs3.Close

End Function


Public Function getitemAgeingDataOld(FromDate As Date, _
                                  ToDate As Date, _
                                  Optional GroupID As Integer = 0, _
                                  Optional Item_ID As Integer)
    Dim NameOfAgeType As String

    Dim late_interval As Integer
    Dim ItemID As Long
    Dim Dean_age As Integer

    Dim column_location As Integer
    Dim column_COLOR As String
    Dim customerid As Integer
    Dim i As Integer
    Dim sql As String
    Dim DefaultSalesPersonId As Integer
    Dim Rs3 As New ADODB.Recordset

    sql = "SELECT     TOP 100 PERCENT MAX(dbo.Transactions.Transaction_Date) AS LastDate, dbo.Transaction_Details.Item_ID"
    sql = sql & " ,  DATEDIFF(day,MAX(dbo.Transactions.Transaction_Date)"
    sql = sql & " , " & SQLDate(ToDate, True) & ") as DIFFerents"
    sql = sql & " FROM         dbo.Transaction_Details INNER JOIN"
    sql = sql & "   dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    sql = sql & "  INNER JOIN dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    sql = sql & " WHERE     (dbo.Transactions.Transaction_Type = 21)"
    sql = sql & " AND (dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & " )"
    sql = sql & " AND (dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & " )"

    If GroupID <> 0 Then
        sql = sql & " AND (dbo.TblItems.GroupID = " & GroupID & ")"
    End If

    If Item_ID <> 0 Then
        sql = sql & " AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
    End If

    sql = sql & " GROUP BY dbo.Transaction_Details.Item_ID"
    sql = sql & " ORDER BY dbo.Transaction_Details.Item_ID"

    Dim str As String
    Dim Note_Value As Double
    str = "delete TblTempItemAging"

    Cn.Execute str

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount = 0 Then Exit Function

    If Rs3.RecordCount > 0 Then

        Rs3.MoveFirst

        For i = 1 To Rs3.RecordCount

            late_interval = Rs3.Fields("DIFFerents").value
            ItemID = Rs3.Fields("Item_ID").value
            column_location = get_late_location2(late_interval)
            '     column_COLOR = get_late_COLOR(column_location, NameOfAgeType)

            add_record_to_table "TblTempItemAging", " ItemID,LateID ", ItemID & " ," & column_location, "ItemID", 0

            Rs3.MoveNext
        Next i

    End If

    Rs3.Close

    Dim StrSQL As String

End Function

Public Function GetNetsalaryVouchers(NoteType As Integer, _
                                     FromDate As Date, _
                                     ToDate As Date) As Double
    Dim StrSQL  As String
    Dim DepitValue As Double
    Dim CreditValue As Double

    Dim Account_Code_dynamic7 As String '–„„ «·„ÊŸ›Ì‰
    Account_Code_dynamic7 = get_account_code_branch(7, my_branch)

    If Account_Code_dynamic7 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic7 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    –„„ «·„ÊŸ›Ì‰          ", vbCritical

            Exit Function
        End If
    End If

    Dim Account_Code_dynamic29 As String '«·«ÃÊ— «·„” Õﬁ… «·„ÊŸ›Ì‰
    Account_Code_dynamic29 = get_account_code_branch(29, my_branch)

    If Account_Code_dynamic29 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Exit Function
    Else

        If Account_Code_dynamic29 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    «·«ÃÊ—  «·„” Õﬁ… «·„ÊŸ›Ì‰          ", vbCritical

            Exit Function
        End If
    End If

    StrSQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    StrSQL = StrSQL & " FROM         dbo.Notes INNER JOIN"
    StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    StrSQL = StrSQL & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
    StrSQL = StrSQL & " WHERE     (dbo.Notes.NoteType = " & NoteType & ") AND (dbo.Notes.NoteDate >=" & SQLDate(FromDate, True) & " ) AND (dbo.Notes.NoteDate <= " & SQLDate(ToDate, True) & ")"
    'StrSQL = StrSQL & " AND (branch_no = " & Val(P_dcBranch) & ")"
    StrSQL = StrSQL & "  AND (dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic7 & "' AND dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic29 & "')"
    'StrSQL = StrSQL & "  AND (branch_no = " & val(P_dcBranch) & ")"

    StrSQL = StrSQL & " GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    StrSQL = StrSQL & " HAVING      (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"

    Dim RsUnitData As New ADODB.Recordset

    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        DepitValue = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
    Else
        DepitValue = 0

    End If

    RsUnitData.Close

    StrSQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    StrSQL = StrSQL & " FROM         dbo.Notes INNER JOIN"
    StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    StrSQL = StrSQL & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
    StrSQL = StrSQL & " WHERE     (dbo.Notes.NoteType = " & NoteType & ") AND (dbo.Notes.NoteDate >=" & SQLDate(FromDate, True) & " ) AND (dbo.Notes.NoteDate <= " & SQLDate(ToDate, True) & ")"
    'StrSQL = StrSQL & " AND (branch_no = " & Val(P_dcBranch) & ")"
    StrSQL = StrSQL & "  AND (dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic7 & "' AND dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic29 & "')"
    'StrSQL = StrSQL & "  AND (branch_no = " & val(P_dcBranch) & ")"

    StrSQL = StrSQL & " GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
    StrSQL = StrSQL & " HAVING      (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1)"

    Dim RsUnitData1 As New ADODB.Recordset

    RsUnitData1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData1.RecordCount) > 0 Then

        CreditValue = IIf(IsNull(RsUnitData1("Total").value), 0, (RsUnitData1("Total").value))
    Else
        CreditValue = 0

    End If

    RsUnitData1.Close

    GetNetsalaryVouchers = Abs(DepitValue - CreditValue)

End Function

Public Function CostForMaintenance(TicktNO As String) As Double

    Dim StrSQL  As String
    StrSQL = "SELECT     SUM(dbo.Transaction_Details.Price) AS Cost"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.TicketNO = N'" & TicktNO & "')"

    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        CostForMaintenance = IIf(IsNull(RsUnitData("Cost").value), 0, (RsUnitData("Cost").value))

    Else
        CostForMaintenance = 0

    End If

    RsUnitData.Close

End Function
Public Function CheckCustomerID(CustGID As Double, _
                                Optional ByRef Custcode As String, _
                                Optional ByRef CustName As String, Optional ByRef block As Boolean = False, Optional ByRef reson As String) As Boolean

    Dim StrSQL  As String
    StrSQL = "SELECT    *  FROM      TblCustemers where CustGID=" & CustGID

    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        CheckCustomerID = True
        Custcode = IIf(IsNull(RsUnitData("Fullcode").value), "", (RsUnitData("Fullcode").value))

        If SystemOptions.UserInterface = ArabicInterface Then
            CustName = IIf(IsNull(RsUnitData("CusName").value), 0, (RsUnitData("CusName").value))
        Else
            CustName = IIf(IsNull(RsUnitData("CusNamee").value), 0, (RsUnitData("CusNamee").value))
        End If
        If RsUnitData("locked").value = True Then
        block = True
        reson = IIf(IsNull(RsUnitData("Remark2").value), "", (RsUnitData("Remark2").value))
        End If

    Else
        CheckCustomerID = False

    End If

    RsUnitData.Close

End Function

Public Function CheckCustomerIDold(CustGID As Double, _
                                Optional ByRef Custcode As String, _
                                Optional ByRef CustName As String) As Boolean

    Dim StrSQL  As String
    StrSQL = "SELECT    *  FROM      TblCustemers where CustGID=" & CustGID

    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        CheckCustomerIDold = True
        Custcode = IIf(IsNull(RsUnitData("Fullcode").value), "", (RsUnitData("Fullcode").value))

        If SystemOptions.UserInterface = ArabicInterface Then
            CustName = IIf(IsNull(RsUnitData("CusName").value), 0, (RsUnitData("CusName").value))
        Else
            CustName = IIf(IsNull(RsUnitData("CusNamee").value), 0, (RsUnitData("CusNamee").value))
        End If

    Else
        CheckCustomerIDold = False

    End If

    RsUnitData.Close

End Function

Public Function GetScreenDescription(ScreenName As String) As String
    Dim StrSQL  As String
    StrSQL = "SELECT    *  FROM      TblWorkFollow where ScreenName='" & ScreenName & "'"

    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        GetScreenDescription = IIf(IsNull(RsUnitData("Remark").value), 0, (RsUnitData("Remark").value))
    Else
        GetScreenDescription = ""

    End If

    RsUnitData.Close

End Function
Public Function GetFirstBox() As Integer
    Dim StrSQL  As String
    StrSQL = "SELECT    * from TblBoxesData "
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                    GetFirstBox = IIf(IsNull(RsUnitData("BoxID").value), 0, (RsUnitData("BoxID").value))
    Else
        GetFirstBox = 0

    End If

    RsUnitData.Close

End Function

Public Function GetSalesValue(FromDate As Date, _
                              ToDate As Date, _
                              ItemType As Integer) As Double
    Dim StrSQL  As String
    StrSQL = "SELECT     SUM(Total) AS Totalvalue, SUM(TotalQty) AS totalQty"
    StrSQL = StrSQL & " FROM         dbo.QryItemsSalesTotal(21, DEFAULT, DEFAULT, " & SQLDate(FromDate, True) & ", " & SQLDate(ToDate, True) & "," & ItemType & ") QryItemsSalesTotal"
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                    GetSalesValue = IIf(IsNull(RsUnitData("Totalvalue").value), 0, (RsUnitData("Totalvalue").value))
    Else
        GetSalesValue = 0

    End If

    RsUnitData.Close

End Function

Public Function GetTransactionsEData(Transaction_ID As Integer, _
                                     Optional ByRef TransactionEnglishName As String, _
                                     Optional ByRef CusNamee As String, _
                                     Optional ByRef storenamee As String)
    Dim StrSQL  As String
    StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Type, dbo.TransactionTypes.TransactionEnglishName, dbo.TransactionTypes.TransactionTypeName, "
    StrSQL = StrSQL & "   dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID"
    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_ID = " & Transaction_ID & ")"
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        TransactionEnglishName = IIf(IsNull(RsUnitData("TransactionEnglishName").value), 0, (RsUnitData("TransactionEnglishName").value))
        CusNamee = IIf(IsNull(RsUnitData("CusNamee").value), 0, (RsUnitData("CusNamee").value))
        storenamee = IIf(IsNull(RsUnitData("StoreNamee").value), 0, (RsUnitData("StoreNamee").value))

    Else
        TransactionEnglishName = ""
        CusNamee = ""
        storenamee = ""

    End If

    RsUnitData.Close

End Function

Public Function GetISSueVoucherForProductionValue(FromDate As Date, _
                                                  ToDate As Date, _
                                                  ItemType As Integer) As Double
    Dim StrSQL  As String
    StrSQL = "SELECT     SUM(Total) AS Totalvalue, SUM(TotalQty) AS totalQty"
    StrSQL = StrSQL & " FROM         dbo.QryItemsSalesTotal(27, DEFAULT, DEFAULT, " & SQLDate(FromDate, True) & ", " & SQLDate(ToDate, True) & "," & ItemType & ") QryItemsSalesTotal"
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        GetISSueVoucherForProductionValue = IIf(IsNull(RsUnitData("Totalvalue").value), 0, (RsUnitData("Totalvalue").value))
    Else
        GetISSueVoucherForProductionValue = 0

    End If

    RsUnitData.Close

End Function

Public Function GetExpensestotal(FromDate As Date, _
                                 ToDate As Date) As Double
    Dim StrSQL  As String

    ' „ «· ⁄œÌ· ·«Õ ”«» «·„’—Ê› «·„Œ›÷ ›Ì 18 12 2012
    StrSQL = "  SELECT     SUM ("
    StrSQL = StrSQL & "  Case"
    StrSQL = StrSQL & "    When Credit_Or_Debit=0 Then Value*1"
    StrSQL = StrSQL & " When Credit_Or_Debit=1 Then Value*-1"
    StrSQL = StrSQL & " Else  0"
    StrSQL = StrSQL & " End"
    StrSQL = StrSQL & " ) AS Total"
    StrSQL = StrSQL & " FROM         dbo.ACCOUNTS INNER JOIN"
    StrSQL = StrSQL & " dbo.ExpensesType ON dbo.ACCOUNTS.Account_Code = dbo.ExpensesType.Account_Code INNER JOIN"
    StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
    StrSQL = StrSQL & " WHERE     (dbo.ExpensesType.TypicalProduction = 1)  AND"
    StrSQL = StrSQL & "       RecordDate >= " & SQLDate(FromDate, True)
    StrSQL = StrSQL & "  AND RecordDate <= " & SQLDate(ToDate, True)
    'StrSQL = StrSQL & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = " & val(P_dcBranch) & ")"
    'StrSQL = StrSQL & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = 1)"
    Debug.Print StrSQL

    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        GetExpensestotal = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
    Else
        GetExpensestotal = 0

    End If

    RsUnitData.Close
End Function

Public Function gettotal(NoteType As Integer, _
                         FromDate As Date, _
                         ToDate As Date, _
                         Optional AllocationType As Integer = -1, Optional branch_no As Integer = 1) As Double
    Dim StrSQL  As String

    StrSQL = "  SELECT     SUM(Note_Value) AS Total from dbo.Notes"

    StrSQL = StrSQL & " WHERE      NoteDate >= " & SQLDate(FromDate, True)
    StrSQL = StrSQL & "  AND NoteDate <= " & SQLDate(ToDate, True)
    StrSQL = StrSQL & " AND (NoteType = " & NoteType & ")"
     'StrSQL = StrSQL & " AND (branch_no = " & val(P_dcBranch) & ")"
         StrSQL = StrSQL & " AND (branch_no =" & branch_no & ")"
    If AllocationType <> -1 Then
        StrSQL = StrSQL & " AND  AllocationType=" & AllocationType
    End If

    Dim RsUnitData As New ADODB.Recordset

    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        gettotal = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
    Else
        gettotal = 0

    End If

    RsUnitData.Close
End Function
Public Function gettransactiontotal(Transaction_ID As Long) As Double


    Dim StrSQL  As String

    StrSQL = "   SELECT     QryTransactionsTotal.TransNet + Transactions.vat as TransNet, dbo.Transactions.Transaction_ID"
StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & "                       dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID"
StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_ID = " & Transaction_ID & ")"


    Dim RsUnitData As New ADODB.Recordset

    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        gettransactiontotal = IIf(IsNull(RsUnitData("TransNet").value), 0, (RsUnitData("TransNet").value))
    Else
        gettransactiontotal = 0

    End If

    RsUnitData.Close

End Function

Public Function GetSalesCost(FromDate As Date, _
                             ToDate As Date) As Double
                             ' ﬂ·›… ”‰œ«  «·’—› «·„Œ“‰Ì
    Dim StrSQL  As String
           StrSQL = "  SELECT     SUM(dbo.Transaction_Details.SHOWQTY * dbo.Transaction_Details.SHOWPrice) AS TotalCost"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    StrSQL = StrSQL & " WHERE     (dbo.Transactions.Transaction_Type = 19) and    (dbo.Transactions.Transaction_Date  >= " & SQLDate(FromDate, True)
    StrSQL = StrSQL & "  AND (dbo.Transactions.Transaction_Date  <= " & SQLDate(ToDate, True)
    StrSQL = StrSQL & " ))"
    'StrSQL = StrSQL & " AND (dbo.Transaction_Details.BranchId  = " & val(P_dcBranch) & ") and Doctype is null"
    StrSQL = StrSQL & "  and  ( Doctype is null  or Doctype in(SELECT     id FROM         dbo.TblDoCumentsTypes  WHERE     (WorkWithProducction = 1))   )  "

    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        GetSalesCost = IIf(IsNull(RsUnitData("TotalCost").value), 0, (RsUnitData("TotalCost").value))
    Else
        GetSalesCost = 0

    End If

    RsUnitData.Close
End Function

'Â–… «·œ«·Â ··›’· »Ì‰ ›« Ê—… „«·Ì… Ê›Ê« Ì— «·«’Ê·
Public Function GetFinInvoiceType(NoteID As Double) As Double
    Dim StrSQL  As String


    StrSQL = "   SELECT     bill_type From dbo.notes_all Where (noteid = " & NoteID & ")"


    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then

        GetFinInvoiceType = IIf(IsNull(RsUnitData("bill_type").value), 0, (RsUnitData("bill_type").value))
    Else
        GetFinInvoiceType = 0

    End If

    RsUnitData.Close
End Function

Public Function GetEmployeeSalaryProject(Emp_id As Integer, whrstr As String, Optional MonthID As Integer = 0, Optional YearID As Integer = 0) As Double
  Dim sql As String
  Dim Rs3 As ADODB.Recordset
  Set Rs3 = New ADODB.Recordset
  sql = " SELECT     dbo.ProJectMofrdSalar.EmpID, SUM(dbo.ProJectMofrdSalar.Total) AS SumValuee"
  sql = sql & "   FROM         dbo.mofrad RIGHT OUTER JOIN"
  sql = sql & "                    dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type RIGHT OUTER JOIN"
  sql = sql & "                    dbo.ProJectMofrdSalar ON dbo.mofrdat.mofrad_code = dbo.ProJectMofrdSalar.MofrdID"
  sql = sql & " Where (dbo.mofrdat.mofrad_type  IN (" & whrstr & ")) And (dbo.ProJectMofrdSalar.YearID = " & YearID & ") And (dbo.ProJectMofrdSalar.MonthID = " & MonthID & ")"
  sql = sql & " GROUP BY dbo.ProJectMofrdSalar.EmpID"
  sql = sql & " HAVING      (dbo.ProJectMofrdSalar.EmpID = " & Emp_id & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetEmployeeSalaryProject = Abs(IIf(IsNull(Rs3("SumValuee").value), 0, Rs3("SumValuee").value))
Else
GetEmployeeSalaryProject = 0
End If
End Function


Public Function getEmployeeCashAssest(EmpID As Integer)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim Balance As Double
    Dim Account_code As String
    Balance = 0

    sql = "SELECT *    from TblBoxesData  WHERE     empid = " & EmpID & " and Type =1 "

    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adcmtext

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            Account_code = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            Balance = Balance + get_balanceFromGl(Account_code, , , True)
            'Balance = Balance + get_balanceFromGlNew(Account_code, , , , , , , , , , True)
            

            rs.MoveNext
        Next i

    End If




    getEmployeeCashAssest = Balance
End Function


Public Function GetActiveInvestmenAccound(Optional InveID As Double = 0) As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = "Select Account_Code6 from TblActivateInvestment where InviseNo=" & InveID & ""
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GetActiveInvestmenAccound = IIf(IsNull(Rs8("Account_Code6").value), "", Rs8("Account_Code6").value)
Else
GetActiveInvestmenAccound = ""
End If
End Function
Public Function get_employee_information(ID As Integer, Optional ByRef date1 As Date _
, Optional ByRef DepartmentID As Double, Optional ByRef SpecificationID As Double _
, Optional ByRef JobTypeID As Double, Optional ByRef gradeID As Double _
, Optional ByRef Account_code2 As String, Optional ByRef Account_code As String _
, Optional ByRef endContractPerMonth As Double, Optional ByRef Nationality As String _
 , Optional ByRef mangerid As Integer, Optional ByRef swapedempid As Integer _
 , Optional ByRef GroupID As Integer, Optional ByRef NumPasp As String _
 , Optional ByRef NumEkama As String, Optional ByRef placeEkama As String _
  , Optional ByRef pasplace As String, Optional ByRef DateEndekamaH As String _
 , Optional ByRef DateEndPasp As Date, Optional ByRef BignDateWork As Date _
   , Optional ByRef LastDate As Date, Optional ByRef JobTypeName As String, Optional ByRef Contract_period1 As Integer _
  , Optional ByRef Contract_periodno1 As Integer, Optional ByRef visano As String, Optional ByRef dcjopstatus As Integer _
  , Optional ByRef JobTypeIDIqama As Integer, Optional ByRef DateMoveNo As Date, Optional ByRef DateExpoekama As String, Optional ByRef Mobile As String, Optional ByRef BlnceVocat As Integer = 0, Optional ByRef Emp_Phone As String, Optional ByRef Contract_date1 As Date, Optional ByRef RegionID As Integer = 0, Optional ByRef due_period As Integer, Optional ByRef Due_period_no As Integer, Optional ByRef Holiday_period_no As Integer, Optional ByRef Holiday_period As Integer _
  , Optional BranchID As Integer, Optional DriverLicenseendH As String, Optional DriverLicense As String, Optional ByRef lastHolidaydate As Date, Optional ByRef lastHolidaydateH As String, Optional ADDtype_Contract As Integer)

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
sql = " SELECT     dbo.TblEmpSpecifications.SpecificationName AS Expr1, dbo.TblEmpDepartments.DepartmentName AS Expr2, dbo.TblEmpDepartments.DepartmentNamee AS Expr3,"
sql = sql & "                      dbo.TblEmpJobsTypes.JobTypeName AS Expr4, dbo.TblEmpJobsTypes.JobTypeNamee AS Expr5, dbo.TblEmpGrades.namee AS grdename,"
sql = sql & "                      dbo.TblEmpGrades.name AS grdenamee, dbo.TblEmployee.*, dbo.TblEmployee.JobTypeID3 AS JobTypeID3Iq, TblEmpJobsTypes_1.JobTypeName AS jobnameiqama,"
sql = sql & "                      TblEmpJobsTypes_1.JobTypeNamee AS jobnameiqamaE, dbo.Contract.Contract_period_no , dbo.Contract.Contract_period "
sql = sql & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
sql = sql & "                      dbo.Contract ON dbo.TblEmployee.Emp_ID = dbo.Contract.Emp_id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_1 ON dbo.TblEmployee.JobTypeID3 = TblEmpJobsTypes_1.JobTypeID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpGrades ON dbo.TblEmployee.gradeID = dbo.TblEmpGrades.gradeid LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpSpecifications ON dbo.TblEmployee.SpecificationID = dbo.TblEmpSpecifications.SpecificationID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID"

sql = sql & " WHERE     (dbo.TblEmployee.Emp_ID = " & ID & ")"

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount > 0 Then

        date1 = IIf(Not IsDate(Rs3("BignDateWork").value), Date, Rs3("BignDateWork").value)
 DepartmentID = IIf(Not IsNull(Rs3("DepartmentID").value), Rs3("DepartmentID").value, 0)

 SpecificationID = IIf(Not IsNull(Rs3("SpecificationID").value), Rs3("SpecificationID").value, 0)
 JobTypeID = IIf(Not IsNull(Rs3("JobTypeID").value), Rs3("JobTypeID").value, 0)
 JobTypeIDIqama = IIf(Not IsNull(Rs3("JobTypeID3").value), Rs3("JobTypeID3").value, 0)
 gradeID = IIf(Not IsNull(Rs3("gradeID").value), Rs3("gradeID").value, 0)
 Account_code2 = IIf(Not IsNull(Rs3("Account_Code2").value), Rs3("Account_Code2").value, "")
 Account_code = IIf(Not IsNull(Rs3("Account_code").value), Rs3("Account_code").value, "")
 Nationality = IIf(Not IsNull(Rs3("Nationality").value), Rs3("Nationality").value, "")
 mangerid = IIf(Not IsNull(Rs3("mangerid").value), Rs3("mangerid").value, 0)
 swapedempid = IIf(Not IsNull(Rs3("swapedempid").value), Rs3("swapedempid").value, 0)
 GroupID = IIf(Not IsNull(Rs3("GroupID").value), Rs3("GroupID").value, 0)
 NumEkama = IIf(Not IsNull(Rs3("NumEkama").value), Rs3("NumEkama").value, "")
 NumPasp = IIf(Not IsNull(Rs3("NumPasp").value), Rs3("NumPasp").value, "")
 JobTypeName = IIf(Not IsNull(Rs3("Expr4").value), Rs3("Expr4").value, "")
 Contract_period1 = IIf(Not IsNull(Rs3("Contract_period").value), Rs3("Contract_period").value, -1)
Contract_periodno1 = IIf(Not IsNull(Rs3("Contract_period_no").value), Rs3("Contract_period_no").value, 0)
visano = IIf(Not IsNull(Rs3("VisaNo").value), Rs3("VisaNo").value, "")
dcjopstatus = IIf(Not IsNull(Rs3("jopstatusid").value), Rs3("jopstatusid").value, 0)
DateMoveNo = IIf(IsNull(Rs3("DateMoveno").value), Date, Rs3("DateMoveno"))
DateEndekamaH = IIf(Not IsNull(Rs3("DateEndekamah").value), Rs3("DateEndekamah").value, ToHijriDate(Date))
DateExpoekama = IIf(Not IsNull(Rs3("DateExpoekamah").value), Rs3("DateExpoekamah").value, ToHijriDate(Date))
DateEndPasp = IIf(Not IsNull(Rs3("DateEndPasp").value), Rs3("DateEndPasp").value, Date)
pasplace = IIf(Not IsNull(Rs3("pasplace").value), Rs3("pasplace").value, "")
placeEkama = IIf(Not IsNull(Rs3("placeEkama").value), Rs3("placeEkama").value, "")
RegionID = IIf(Not IsNull(Rs3("RegionID").value), Rs3("RegionID").value, 0)
Emp_Phone = IIf(Not IsNull(Rs3("Emp_Phone").value), Rs3("Emp_Phone").value, "")
BignDateWork = IIf(Not IsNull(Rs3("BignDateWork").value), Rs3("BignDateWork").value, Date)
LastDate = IIf(Not IsNull(Rs3("LastDate").value), Rs3("LastDate").value, Date)
Mobile = IIf(Not IsNull(Rs3("Emp_mobile").value), Rs3("Emp_mobile").value, "")
BlnceVocat = IIf(Not IsNull(Rs3("BlnceVocat").value), Rs3("BlnceVocat").value, 0)
BranchID = IIf(Not IsNull(Rs3("BranchId").value), Rs3("BranchId").value, 0)

DriverLicense = IIf(Not IsNull(Rs3("DriverLicense").value), Rs3("DriverLicense").value, "")
DriverLicenseendH = IIf(Not IsNull(Rs3("DriverLicenseendH").value), Rs3("DriverLicenseendH").value, ToHijriDate(Date))
lastHolidaydate = IIf(Not IsNull(Rs3("lastHolidaydate").value), Rs3("lastHolidaydate").value, Date)
lastHolidaydateH = IIf(Not IsNull(Rs3("lastHolidaydateH").value), Rs3("lastHolidaydateH").value, ToHijriDate(Date))
'GroupID

    Else
 BranchID = 0
        date1 = Date
 DepartmentID = 0
 SpecificationID = 0
 JobTypeID = 0
 gradeID = 0
 Nationality = ""
 mangerid = 0
 swapedempid = 0
 GroupID = 0
 lastHolidaydate = Date
 lastHolidaydateH = ToHijriDate(Date)
    End If

    Rs3.Close
Dim Contract_period_no As Double
Dim Contract_period  As Double

Dim Contract_date As Date
sql = "  select * from Contract WHERE     (dbo.Contract.Emp_ID = " & ID & ")"

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.RecordCount > 0 Then
     Contract_period_no = IIf(Not IsNull(Rs3("Contract_period_no").value), Rs3("Contract_period_no").value, 0)
  Contract_period = IIf(Not IsNull(Rs3("Contract_period").value), Rs3("Contract_period").value, 0)
    Contract_date = IIf(Not IsNull(Rs3("Contract_date").value), Rs3("Contract_date").value, Date)
    Contract_date1 = IIf(Not IsNull(Rs3("Contract_date").value), Rs3("Contract_date").value, Date)
  endContractPerMonth = DateDiff("M", Contract_date, Date)
        Due_period_no = IIf(Not IsNull(Rs3("Due_period_no").value), Rs3("Due_period_no").value, -1)
due_period = IIf(Not IsNull(Rs3("due_period").value), Rs3("due_period").value, 0)
  endContractPerMonth = DateDiff("M", Contract_date, Date)
       Holiday_period_no = IIf(IsNull(Rs3("Holiday_period_no").value), 0, Rs3("Holiday_period_no").value)
 ADDtype_Contract = IIf(IsNull(Rs3("ADDtype_Contract").value), 0, Rs3("ADDtype_Contract").value)
    If IsNull(Rs3("Holiday_period").value) Then
        Holiday_period = 0
    Else
       Holiday_period = Rs3("Holiday_period").value
    End If

If Contract_period = 1 Then

Contract_period_no = Contract_period_no * 12
End If

endContractPerMonth = Contract_period_no - endContractPerMonth

    Else

    End If

End Function



Public Function OpenRecordSet(ByVal SqlStatment As String, _
                              OpenType As CursorTypeEnum, _
                              LockType As LockTypeEnum, _
                              Optional IsLocal As Boolean = False, _
                              Optional TrayCachFrist As Boolean = False, _
                              Optional CursorLocation As Integer = -1, _
                              Optional UseFunctionLockType As Boolean = False) As ADODB.Recordset
    AduitLastSQL = SqlStatment
    'Samy Commnt this line
'    If StringStartsWith(LTrim$(UCase(SqlStatment)), "EXEC ") Then
'        Set OpenRecordSet = OpenRecordSetForSP(SqlStatment, OpenType, LockType, IsLocal, TrayCachFrist, CursorLocation)
'        Exit Function
'    End If
    '********************
    BuildAppInfo
    '*********************
'    Do While IsTestConnection
'        DoEvents
'    Loop

RE:
    On Error GoTo eh

    '    DatabaseName = "Topsystems111"
    If IsLocal Then
        '        StrConn1 = "Driver={SQL Server};Packet Size=32768;Server=" & ServerNameLocal & _
                 '           ";Uid=" & UserNameLocal & ";Pwd=" & PASSWORDLocal & _
                 '           ";Database=" & DatabaseNameLocal & ";App=" & MyAPPUserInfo

        StrConn1 = "Provider=SQLNCLI11.1;Data Source=" & ServerNameLocal & _
                   ";User ID=" & UserNameLocal & ";Password=" & PASSWORDLocal & _
                   ";Initial Catalog=" & DatabaseNameLocal & ";DataTypeCompatibility=80;Application Name=" & MyAPPUserInfo
        '-----------------------------------------------
        If DBLocal.State <> adStateOpen Then
            If Not DBLocalIsCustomStringConnection Then
                DBLocal.ConnectionTimeout = 5000
                DBLocal.CommandTimeout = 5000
                DBLocal.IsolationLevel = adXactReadUncommitted
            Else
                StrConn1 = DBLocalCustomStringConnection
                StrConn1 = SetConnectionSection(StrConn1, "Data Source", ServerNameLocal)
                StrConn1 = SetConnectionSection(StrConn1, "Initial Catalog", DatabaseNameLocal)
                StrConn1 = SetConnectionSection(StrConn1, "User ID", UserNameLocal)
                StrConn1 = SetConnectionSection(StrConn1, "Password", PASSWORDLocal)
            End If
            '--------------------
            DBLocal.Open StrConn1
            '--------------------
        End If
        '--------------------------------------
        DBLocal.CommandTimeout = 5000
        '--------------------------------------
        Set OpenRecordSet = New ADODB.Recordset
        If CursorLocation <> -1 Then
            OpenRecordSet.CursorLocation = CursorLocation
        End If
        OpenRecordSet.Open SqlStatment, DBLocal, OpenType, LockType, adCmdText

    Else
        '        StrConn = "Driver={SQL Server};Packet Size=32768;Server=" & ServerName & _
                 '           ";Uid=" & UserName & ";Pwd=" & password & _
                 '           ";Database=" & DatabaseName & ";App=" & MyAPPUserInfo
        'Provider=SQLOLEDB.1;Password=makkahttd;Persist Security Info=True;User ID=sa;Data Source=196.186.0.205\SQL2008
        StrConn = "Provider=SQLNCLI11.1;Data Source=" & ServerName & _
                  ";User ID=" & UserName & ";Password=" & Password & _
                  ";Initial Catalog=" & DatabaseName & ";DataTypeCompatibility=80;Application Name=" & MyAPPUserInfo
        ' -----------------------------------------------
        '·⁄·«Ã «Œ ·«› «·œ« « »Ì“ ›Ì Õ«·… «·DLL
        ' -----------------------------------------------
        If isDebugMode() Then
            If db.State = adStateOpen Then
                If UCase(ServerName) <> UCase(ServerNameINI) Then
                    Set tt = db.Execute("SELECT SERVERPROPERTY(N'MachineName')AS MachineName, CONNECTIONPROPERTY('local_net_address') AS IPAddress,SERVERPROPERTY('InstanceName') AS InstanceName;")
                    MyMachineName = StrConv(StrConv(tt!MachineName, vbUnicode), vbFromUnicode)    ' TT!MachineName & ""
                    MyIpAddress = CStr(tt!IPAddress) & ""    '  StrConv(tt!IPAddress, vbUnicode)    'CStr(TT!IPAddress) & ""
                    '  MyIPAddress = Replace(CStr(MyIPAddress), " ", "")
                    MyInstanceName = StrConv(StrConv(tt!InstanceName, vbUnicode), vbFromUnicode)    'CStr(TT!InstanceName) & ""
                    If MyInstanceName <> "" Then
                        MyMachineName = MyMachineName & "\" & MyInstanceName
                        MyIpAddress = MyIpAddress & "\" & MyInstanceName
                    End If
                    If UCase(ServerName) <> UCase(MyMachineName) And UCase(ServerName) <> UCase(MyIpAddress) Then
                        'this run only once In same dll when click on form
                        db.Close
                    End If
                End If
            End If
        End If
        '-----------------------------------------------
        If db.State <> adStateOpen Then
            If Not DBIsCustomStringConnection Then
                db.ConnectionTimeout = 50
                db.CommandTimeout = 10000

            Else
                StrConn1 = DBCustomStringConnection
                StrConn1 = SetConnectionSection(StrConn1, "Data Source", ServerName)
                StrConn1 = SetConnectionSection(StrConn1, "Initial Catalog", DatabaseName)
                StrConn1 = SetConnectionSection(StrConn1, "User ID", UserName)
                StrConn1 = SetConnectionSection(StrConn1, "Password", Password)
                StrConn = StrConn1
                RptConn = StrConn1
            End If
            '---------------------
            db.Open StrConn
            '---------------------
        End If
        '--------------------------------------
        db.CommandTimeout = 5000
        '--------------------------------------

        Set OpenRecordSet = New ADODB.Recordset

        If CursorLocation <> -1 Then
            OpenRecordSet.CursorLocation = CursorLocation
        End If



        Set OpenRecordSet.ActiveConnection = db

        OpenRecordSet.Properties("Preserve On Commit").value = True
        OpenRecordSet.Properties("Preserve On Abort").value = True
        '*********************Samy************************************
        If Not UseFunctionLockType Then
            If LockType = adLockOptimistic Or LockType = adLockPessimistic Then
                LockType = adLockOptimistic
            End If
        End If
        '**********************************************************

        OpenRecordSet.Open SqlStatment, , OpenType, LockType, adCmdText
        '*********************
        'AddToCollection OpenRecordSet
        '***********************
        '*****************************************************************************************
        '****************Tray to caching Database that are readonly**********by Khalid************
        '*****************************************************************************************
        If TrayCachFrist And (OpenType = adOpenStatic) And (LockType = adLockReadOnly) Then
            TempfileName = CreateTempFileName()
            OpenRecordSet.save TempfileName, adPersistXML
            OpenRecordSet.Close
            OpenRecordSet.Open TempfileName, "Provider=mspersist"
        End If
        '*****************************************************************************************
        '*****************************************************************************************
        '*****************************************************************************************
        If InStr(1, SqlStatment, "ActiveUsers ", vbTextCompare) = 0 And _
           InStr(1, SqlStatment, "FormDesign ", vbTextCompare) = 0 And _
           InStr(1, SqlStatment, "MenuRights ", vbTextCompare) = 0 And _
           InStr(1, SqlStatment, "MenuShortCuts ", vbTextCompare) = 0 And _
           InStr(1, SqlStatment, "EmployeeNoticeBoard ", vbTextCompare) = 0 And _
           InStr(1, SqlStatment, "Translations ", vbTextCompare) = 0 And _
           InStr(1, SqlStatment, "Translations2 ", vbTextCompare) = 0 _
           Then
            If QState Then QValue = QValue + vbNewLine + SqlStatment Else QValue = ""
            '       If QValueP <> 0 Then CopyMemory QValueP, ByVal QValue, LenB(QValue)
        End If
        '*******************
    End If
    '**********************
    ' „ «Ìﬁ«›Â ⁄‘«‰ «·»·Ã »·«Ï
    '    If Not checkedSqlBefor Then
    '        SqlServerVersionCheck
    '    End If
    '***********************
    Exit Function
eh:
    MErr = Err.Number
    If MErr = -2147467259 Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Samy
        If Not isDebugMode Then
            If Not NotShowTestConnection Then
                FrmTestConnection.IsLocal = IsLocal
                IsTestConnection = True
                FrmTestConnection.show 1    ', FrmMain
                If Not Tested Then
                    '                Unload FrmTestConnection
                    IsTestConnection = False
                    Exit Function
                Else
                    '                Unload FrmTestConnection
                    IsTestConnection = False
                    GoTo RE
                End If

            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
        If Err.Number = -2147217871 Then
            MsgBox MyErrorHandler(Err.Number) & ":ORS", , 5
        ElseIf Err.Number = -2147217900 Then
            Resume Next
        Else
            MsgBox MyErrorHandler(Err.Number) & ":ORS"
        End If
    End If

End Function
'**************************************

'**************************************



Public Function MyErrorHandler(ErrNo As Long) As String
    Mmsg = ""
    Select Case ErrNo

    Case 0
        MyErrorHandler = ""
        Exit Function

    Case -2147217864

        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " „ ≈Ã—«¡  ⁄œÌ·«  ⁄·Ï Â–Â «·‘«‘Â „‰ ÃÂ«“ ¬Œ—- „‰ ›÷·ﬂ «⁄œ  Õ„Ì· «·Õ—ﬂÂ À„ Õ«Ê· „—Â «Œ—Ï" & " - Optimistic concurrency erorr "
        Else
            Mmsg = "This Form is editing from another computer- realod and Try again " & " - Optimistic concurrency erorr "
        End If

    Case -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "«·ÃÂ«“ «·Œ«œ„ «·—∆Ì”Ì „€·ﬁ √Ê €Ì— „ÊÃÊœ ⁄·Ï Â–Â «·‘»ﬂ…" & " - " & ErrNo
        Else
            Mmsg = "Couldn't close file , Because it's already in use by another person or process" & " - " & ErrNo
        End If
    Case -2147352567
        'If SystemOptions.UserInterface = ArabicInterface Then
        '    mMsg = "ÌÃ»  Œ’Ì’ «·ÿ«»⁄«  „‰ ≈œ«—… «·‰Ÿ«„" & " - " & ErrNo
        'Else
        '    mMsg = "Select Correct Report Printer Device" & " - " & ErrNo
        'End If
    Case 3155, 3022, -2147217873, -2147217900    ' insert fail
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " ·«Ì„ﬂ‰ «÷«›… Â–« «·”Ã· ° Â–Â «·»Ì«‰«   „  ”ÃÌ·Â« „‰ ﬁ»·" & " - " & ErrNo
        Else
            Mmsg = "You Can not Add this Record , May be there is Dublicated values" & " - " & ErrNo
        End If
    Case 3200    ' Change Or Delete Failed
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " ·«Ì„ﬂ‰ «·€«¡ √Ê  ⁄œÌ· Â–« «·”Ã·  »”»» ÊÃÊœ »Ì«‰«  √Œ—Ï „— »ÿ… »Â ÊÌÃ» «·€«¡Â« √Ê·«" & " - " & ErrNo
        Else
            Mmsg = "You Can not Delete Or Modify this Record , Because There Is Some Data Depends On It " & " - " & ErrNo
        End If
    Case 3157, 3046, 3202, 3218    ' Update Fail
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " Â‰«ﬂ ›‘· ›Ï  Œ“Ì‰ «· ⁄œÌ·«  ° ﬁœ ÌﬂÊ‰ «·”Ã· „ﬁ›· »Ê«”ÿ… „” Œœ„ ¬Œ—° Õ«Ê· „—… √Œ—Ï" & " - " & ErrNo
        Else
            Mmsg = "Update Failed , May be The record is locked by another User , Try Again " & " - " & ErrNo
        End If
    Case 3186, 3187, 3188
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "”Ã· „€·ﬁ »Ê«”ÿ… „” Œœ„ ¬Œ—" & " - " & ErrNo
        Else
            Mmsg = "Current Record locked by Another user" & " - " & ErrNo
        End If
    Case 3167
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " „ «·€«¡ Â–« «·”Ã· »«·›⁄· " & " - " & ErrNo
        Else
            Mmsg = "Record Already Deleted" & " - " & ErrNo
        End If
    Case 3314
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "„‰ ›÷·ﬂ √ﬂ„· «·»Ì«‰«  ﬁ»· «· Œ“Ì‰" & " - " & ErrNo
        Else
            Mmsg = "Please Complete the data before saving" & " - " & ErrNo
        End If
    Case 3262, 3211, 3212    ' Locked by another user and wait
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "·« Ì„ﬂ‰ ≈€·«ﬁ «·„·› »”»» ÊÃÊœ „” Œœ„ ¬Œ— ÌﬁÊ„ »≈” Œœ«„Â √Ê ﬁ«„ »≈€·«ﬁÂ" & " - " & ErrNo
        Else
            Mmsg = "Couldn't close file , Because it's already in use by another person or process" & " - " & ErrNo
        End If
    Case 3197    ' Couldn't repaire this files
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "√ﬂÀ— „‰ „” Œœ„ Õ«Ê·Ê«  €ÌÌ— ‰›” «·»Ì«‰«  ›Ï ‰›” «·Êﬁ " & " - " & ErrNo
        Else
            Mmsg = "Another Users are attempting to change the same data at the same time" & " - " & ErrNo
        End If
    Case 3056    ' Couldn't repaire this files
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "·« Ì„ﬂ‰  ’·ÌÕ «·„·›«  «·„” Œœ„…" & " - " & ErrNo
        Else
            Mmsg = "Couldn't repaire this files" & " - " & ErrNo
        End If
    Case 3014, 3037    ' Can't open any more files
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "·« Ì„ﬂ‰ › Õ „·›«  √Œ—Ï" & " - " & ErrNo
        Else
            Mmsg = "Can't open any more files" & " - " & ErrNo
        End If
    Case 3356, 3260, 3261, 3189, 3008, 3164, 3006    ' Table or Database Locked
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "«·„·› „€·ﬁ »Ê«”ÿ… „” Œœ„ ¬Œ—" & " - " & ErrNo
        Else
            Mmsg = "The File is Locked by Another User" & " - " & ErrNo
        End If
    Case 3201    ' Add Or Edit Fail
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " ·«Ì„ﬂ‰ «÷«›… Â–« «·”Ã· √Ê «· ⁄œÌ· ›ÌÂ ° ·√‰Â „— »ÿ »„·› ·„ Ì „ «·≈÷«›… √Ê «· ⁄œÌ· ›ÌÂ Õ Ï «·¬‰" & " - " & ErrNo
        Else
            Mmsg = "You Can not Add this Record or Change it , Because it's Linked to a File that has not been Added or Changed till Now" & " - " & ErrNo
        End If
    Case -2147217887
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "Œÿ√ €Ì— „⁄—Ê› ° Õ«Ê·  ‰›Ì– ‰›” «·⁄„·Ì… „—… √Œ—Ï" & " - " & ErrNo
        Else
            Mmsg = "Undefined Error , Try again : " & ErrNo
        End If
    Case 3704
        '********By Khalid
        On Error Resume Next
        db.Close
        Exit Function
    Case -1000000001
        '«—Ê—— » «⁄ «·«Ê Ê ﬂÊ„»Ì·Ì  ›ﬂﬂ „‰Â
        MyErrorHandler = ""
        Exit Function
    End Select
    '*************************
    If Err.Number = vbObjectError + 1000 Then
        If Not ArabicInterface Then
            mText = Trim(Mmsg)
            If Trim(mText) <> "" Then
                Cond = "Arabic = N'" & Trim(mText) & "'"

                s = "Select * from Translations where " & Cond
                Set Translations = OpenRecordSet(s, adOpenStatic, adLockReadOnly)
                '------------------------
                If Not Translations.EOF Then
                    Mmsg = IIf(Trim(Translations!English & "") <> "", Trim(Translations!English & ""), Mmsg)
                End If
            End If
        End If

        Mmsg = Mmsg & vbNewLine & Err.Description
    Else
        Mmsg = Mmsg & vbNewLine & Err.Description & " : " & Err.Number
    End If
    '*************************
    If ErrNo <> -2147217864 Then  '  Ã«Â· «—Ê—— «·ﬂÊ‰ﬂ—‰”Ï  ‘Ìﬂ
        If db.Errors.count > 0 Then
            ss = ""
            Dim adoErr As ADODB.Error
            j = 1
            On Error GoTo EEE
            For Each adoErr In db.Errors
                If adoErr.Number <> 0 Then
                    If j = 1 Then ss = vbNewLine & "-------SQL Errors-------"
                    ss = ss & vbNewLine & "Error (" & j & ")=" & adoErr.Description & " : " & adoErr.Number & ":" & adoErr.SQLState & ":" & adoErr.NativeError
                    j = j + 1
                End If
            Next adoErr
EEE:
            ' for this rand error Not enough storage is available to process this command.
            If Err.Number = 48 Then
                Set adoErr = db.Errors(0)
                ss = ss & vbNewLine & "Error (48)=" & adoErr.Description & " : " & adoErr.Number & ":" & adoErr.SQLState & ":" & adoErr.NativeError
            End If
            On Error GoTo 0
            Mmsg = Mmsg & vbNewLine & ss
        End If
    End If
    '*************************
    'If Trim(mMsg) <> "()(0)" Then MyErrorHandler = mMsg Else MyErrorHandler = ""
    MyErrorHandler = Mmsg & ":" & Erl
    IsAboutError = True

End Function



Public Sub BuildAppInfo()
    IsGetExeDate = True
    ExeVer = App.Major & "." & App.Minor & "." & App.Revision
    'œÏ Â ” Œœ„ ›Ï «·—”«Ì· «··Ï Â —ÊÕ ··ÌÊ“—“ ⁄‘«‰ ‰⁄—› „‰ «·ÌÊ“— «·«Ê‰ ·«Ì‰
    MyAPPUserInfo = "" 'IIf(CurrentUser & "" = "", "•", CurrentUser) & "|" & App.EXEName & "|" & ExeVer & "|" & CustomerCode & "|" & GetEXEDateForConnectio

End Sub



Public Function SetConnectionSection(ByVal ConnStr As String, _
                                     ByVal SectionName As String, _
                                     ByVal SectionValue As String) As String
'StrConn1=SetConnectionSection(StrConn1,"Data Source","")
'StrConn1=SetConnectionSection(StrConn1,"Initial Catalog","")
'StrConn1=SetConnectionSection(StrConn1,"User ID","")
'StrConn1=SetConnectionSection(StrConn1,"Password","")
    Dim isFound As Boolean
    Dim vv
    '"Provider=SQLNCLI.1;Password=pass;Persist Security Info=True;User ID=sa;Initial Catalog=Database;Data Source=Server"
    vv = Split(ConnStr, ";")
    l = Len(UCase(SectionName) & "=")
    For i = 0 To UBound(vv)
        If left(UCase(vv(i)), l) = left(UCase(SectionName) & "=", l) Then
            vv(i) = SectionName & "=" & SectionValue
            isFound = True
            Exit For
        End If
    Next
    If isFound Then
        Dim ss As String
        For i = 0 To UBound(vv)
            ss = ss & vv(i) & ";"
        Next
        SetConnectionSection = ss
    Else
        SetConnectionSection = ConnStr
    End If
End Function




Public Function isDebugMode() As Boolean
    isDebugMode = isDebugModeC
End Function





Private Function CreateTempFileName(Optional ByVal Prefix As String) As String
    Dim TempFile As String  ' receives name of temporary file
    Dim slength As Long   ' receives length of string returned for the path
    Dim lastfour As Long  ' receives hex value of the randomly assigned ????
    If TempPath = "" Then
        ' Get Windows's temporary file path
        TempPath = Space(255)  ' initialize the buffer to receive the path
        slength = GetTempPath(255, TempPath)  ' read the path name
        TempPath = left(TempPath, slength) & "BusinessDimensions\"    ' extract data from the variable
        CreateDir TempPath
        On Error Resume Next
        Kill TempPath & "*.*"
    End If

    ' Get a uniquely assigned random file
    '    TempFile = Space(255)  ' initialize buffer to receive the filename
    '    If Prefix = "" Then Prefix = "TopSys" 'Format(Now, "YYYYMMDDHHNNSS")
    '    lastfour = GetTempFileName(TempPath, Prefix, 0, TempFile)       ' get a unique temporary file name
    '    ' (Note that the file is also created for you in this case.)
    '    TempFile = Left(TempFile, InStr(TempFile, vbNullChar) - 1)   ' extract data from the variable
    TempFile = TempPath & Prefix & Format(Now, "YYYYMMDDHHNNSS") & Int(Rnd(100) * 1000) & ".xml"
    On Error Resume Next
    Kill TempFile
    CreateTempFileName = TempFile
End Function



Public Sub CreateDir(StrPath As String)
    On Error Resume Next
    Dim ArrFolders As Variant
    ArrFolders = Split(StrPath, "\")
    Dim i As Long
    Dim CurPath As String: CurPath = ArrFolders(0)
    MkDir CurPath

    For i = 1 To UBound(ArrFolders)
        CurPath = CurPath & "\" & ArrFolders(i)
        MkDir CurPath
    Next i
    On Error GoTo 0

    If Len(Dir(StrPath, vbDirectory)) = 0 Then
        Err.Raise vbObjectError, , "Can't create dir" & vbCrLf & StrPath & vbCrLf & ":(((("
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Public Sub TranslateForm(Frm As Object, ByVal Arabic As Boolean)
    On Error Resume Next
    '------------------------
    'If StrConn = "" Then Exit Sub    'for only one time in first instalation
    '------------------------
 '   If Arabic Or UCase(frm.Name) = "FRMMSGBOX" Then Exit Sub    ' „⁄—» „‰ Êﬁ  «· ’„Ì„
    '--------------------------------
    'If Arabic And frm.RightToLeft Then Exit Sub ' „⁄—» »«·›⁄·
    'If Not Arabic And Not frm.RightToLeft Then Exit Sub
    '-----------------------------------
    Dim rsDummy As New ADODB.Recordset
    Load Frm
    '**********************************
    Dim Ctr   As Control
    Dim mText As String
    '------------------------
  '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
 '   frm.RightToLeft = Arabic
 '   RTLTree frm
 '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
    '------------------------
    ' Field Captions & Visibility
    '------------------------
    If Trim(Frm.Name) = "frmReports" Then
        Msg = "·« Ì„ﬂ‰  —Ã„Â «· ﬁ—Ì— „‰ Â‰« .... «€·ﬁ «· ﬁ—Ì— Êﬁ„ »«⁄«œÂ › ÕÂ „⁄ «Œ Ì«— «· —Ã„Â «·„ÿ·Ê»Â"

        Msg2 = "Can't Translate Report from here .. reopn the report with choice languge  "
'        MyMsgbox IIf(ArabicInterface, Msg, Msg2)

        '        TranslateReport Rpt, Arabic
        '        frmReports.CV1.Refresh
        Exit Sub
    End If
    '------------------------
    mArabicCaption = ""
    mEnglishCaption = ""
    '------------------------
    mText = Trim(Frm.Caption)
    mArabicCaption = mText
    If Trim(mText) <> "" Then
        Frm.Caption = IIf(Arabic, mArabicCaption, mEnglishCaption) & TranslateText(mText, Arabic)
    End If
    '------------------------
Dim mIndexArr As Long
    For Each Ctr In Frm.Controls

        If (TypeOf Ctr Is Label) _
           Or (TypeOf Ctr Is CheckBox) _
           Or (TypeOf Ctr Is XtremeSuiteControls.CheckBox) _
           Or (TypeOf Ctr Is OptionButton) _
           Or (TypeOf Ctr Is RadioButton) _
           Or (TypeOf Ctr Is ISButton) _
           Or (TypeOf Ctr Is frame) _
           Or (TypeOf Ctr Is CommandButton) _
           Or (TypeOf Ctr Is XtremeSuiteControls.PushButton) _
           Or (TypeOf Ctr Is XtremeSuiteControls.PushButton) Then
            '------------------------
            mIndexArr = FindIndex(Frm, Ctr)
            If mIndexArr = 69 Then
            xx = xx
            End If

'            s = "SELECT * FROM Translations WHERE ControlName = N'" & Trim(Ctr.Name) & "' AND ControlIndex = N'" & IIf(mIndexArr = -99, "", mIndexArr) & "' AND FormName = N'" & Trim(Frm.Name) & "'"
'            Set rsDummy = New ADODB.Recordset
'            rsDummy.Open s, Cn, adOpenStatic
'            If Not rsDummy.EOF Then
'                If Trim(rsDummy!English & "") <> "" Then
'                    xx = xx
'                End If
'                mText = IIf(Trim(rsDummy!English & "") <> "", Trim(rsDummy!English & ""), Trim(Ctr.Caption))
'                If rsDummy!IsVisible Then
'                    Ctr.Visible = False
'
'                End If
            mText = Trim(Ctr.Caption)
             '   Ctr.Caption = mText
      
       
            Ctr.Caption = TranslateText(mText, False)
            '---------------------------
''            If TypeOf Ctr Is XPFrame30 Then
''                Ctr.Alignment = IIf(Arabic, 2, 0)
     '       Else

     '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
                'If Ctr.Alignment <> 2 Then Ctr.Alignment = IIf(Arabic, 1, 0)
                '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
     '       End If

        ElseIf typename(Ctr) = "CtlCostCenters" Or typename(Ctr) = "CtlBalancesPeriods" Then
            '--------------------------------
            TranslateForm Ctr, Arabic
        ElseIf TypeOf Ctr Is C1Tab Then
            '--------------------------------
            mtab = Ctr.Tab
          '  RTLTree Ctr    ' WillChange The Direction of the Tab Control
            '------------
            For j = 0 To Ctr.Tabs - 1
                mText = Trim(Ctr.TabCaption(j))
                Ctr.TabCaption(j) = TranslateText(mText, False)
            Next
            '----------------
            If mtab = 0 Then
                Ctr.Tab = 1
            Else
                Ctr.Tab = 0
            End If
            Ctr.Tab = mtab    ' To Redraw the Tab Contents
        ElseIf TypeOf Ctr Is VSFlexGrid Or StartsWithKeywords(Ctr.Name) Then
            '--------------------------------
            '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
            For r = 0 To Ctr.FixedRows - 1
                For j = 0 To Ctr.Cols - 1
                    mText = Trim(Ctr.TextMatrix(r, j))
                    Ctr.TextMatrix(r, j) = TranslateText(mText, False)
                Next
            Next

            For r = 0 To Ctr.rows - 1
                For j = 0 To Ctr.Cols - 1
                    If Not (r > 0 And j > 0) Then
                        mText = Trim(Ctr.TextMatrix(r, j))

                        If Not IsNumeric(mText) Then
                            Ctr.TextMatrix(r, j) = TranslateText(mText, False)
                        End If
                    End If
                Next
            Next
'⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
        ElseIf TypeOf Ctr Is TextBox Then
            '--------------------------------
            '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
            'If Ctr.Alignment <> 2 Then Ctr.Alignment = IIf(Arabic, 1, 0)
            '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â


            If Ctr.DataField <> "" Then
                s = "SELECT * FROM Translations WHERE ControlName = N'" & Trim(Ctr.DataField) & "' AND ControlIndex = N'" & Trim(Ctr.DataMember) & "' AND FormName = N'" & Trim(Frm.Name) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenStatic
                 If Not rsDummy.EOF Then

                     If rsDummy!IsVisible Then
                         Ctr.Visible = False

                     End If
                 End If
            End If


        ElseIf TypeOf Ctr Is PictureBox Then
            '--------------------------------

           '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
            ' «·Ì„Ì‰ Ì ÕÊ· ≈·Ï Ì”«— Ê«·⁄ﬂ”
'            If Ctr.Align = 3 Then
'                Ctr.Align = 4
'            ElseIf Ctr.Align = 4 Then
'                Ctr.Align = 3
'            End If
'⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
            '------------------------
        ElseIf TypeOf Ctr Is MSChart Then    ' SaMi 1/11/2009
            '--------------------------------
            For i = 1 To 4
                Ctr.Column = i
                mText1 = Trim(Ctr.ColumnLabel)
                Ctr.ColumnLabel = TranslateText(mText, False)
            Next
            '****************************************************
            mText2 = Trim(Ctr.RowLabel)
            Ctr.RowLabel = TranslateText(mText2, False)
            '************************************************
        End If
        '-------------------------------------
        ' Change Control Direction

        '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'        If Not TypeOf Ctr.Container Is SSTab Then    ' Problem In SSTab as a Container
'            If TypeOf Ctr.Container Is Form Then
'                Ctr.Left = Ctr.Container.ScaleWidth - Ctr.Left - Ctr.Width  '- IIf(Arabic, 0, 30)
'            Else
'                Ctr.Left = Ctr.Container.Width - Ctr.Left - Ctr.Width    '- IIf(Arabic, 0, 30)
'            End If
'        End If
'        '--------------------------------
'        Ctr.RightToLeft = Arabic    ' Some Controls Does not Support This Property, and some is Read Only
'⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
NextControl:
    Next
    '--------------------------------
    'ChangeControlsDirection Frm
    ' ********************************
    Exit Sub
eh:
    MsgBox MyErrorHandler(Err)
    Resume Next
End Sub



Public Function TranslateText(ByVal mText As String, _
                              Optional ByVal Arabic As Boolean = False) As String
    If mText = "" Then Exit Function
    '-------------------------------
    TranslateText = mText
    '-------------------------------
    If Arabic Then
        Cond = "English=N'" & Trim(mText) & "'"
    Else
        Cond = "Arabic=N'" & Trim(mText) & "'"
    End If

    s = "Select * from Translations where " & Cond
    Dim Translations As New ADODB.Recordset
    Translations.Open s, Cn, adOpenStatic, adLockReadOnly
    '------------------------
    If Not Translations.EOF Then
        If Arabic Then
            TranslateText = Trim(Translations!Arabic & "")
        Else
            TranslateText = IIf(Trim(Translations!English & "") <> "", Trim(Translations!English & ""), Trim(Translations!Arabic & ""))
        End If
    End If
End Function





Public Sub RTLTree(TV As Control, Optional RTL As Integer = 0)

    Dim TreeStyle As Long
    TreeStyle = GetWindowLong(TV.hWnd, GWL_EXSTYLE)
    If RTL > 0 Then
        If (TreeStyle And &H400000) = &H400000 And RTL = 1 Then Exit Sub
        If (TreeStyle And &H400000) = 0 And RTL = 2 Then Exit Sub
    End If
    SetWindowLong TV.hWnd, GWL_EXSTYLE, TreeStyle Xor &H400000
    'SetWindowLong TV.hWnd, GWL_EXSTYLE, TreeStyle Xor &H2& Xor ES_MULTILINE
    'x = SendMessage(TV.hWnd, EM_SETALIGN, ES_RIGHT + ES_MULTILINE, 0)
End Sub



Private Function FindIndex(ByRef F As Form, ByRef ctl As Control) As Integer
    Dim ctlTest As Control
    For Each ctlTest In F.Controls
        If (ctlTest.Name = ctl.Name) And (Not (ctlTest Is ctl)) Then
            'if the object is the same name but is not the same object we can assume it is a control array
            FindIndex = ctl.Index
            Exit Function
        End If
    Next
    'if we get here then no controls on the form have the same name so can't be a control array
    FindIndex = -99
End Function



Public Function GetBOFFromNatioanlID(MyNumber As Variant, MyTest As Byte) As Date


Dim MyProvinces As Variant

Dim r As Integer

Dim yy As String

Dim Ty As String * 1

Dim d As String * 2, m As String * 2, Y As String * 2, X As String * 2, xx As String * 2

'==============================================

'==============================================

GetBOFFromNatioanlID = Date

On Error GoTo 1

If Len(Trim(MyNumber)) = 0 Then

    GoTo 1

End If


If Not IsNumeric(MyNumber) Or Len(MyNumber) <> 14 Then

   ' GetBOFFromNatioanlID = "Error_MyNumber"

    GoTo 1

End If


If MyTest = 1 Then

    d = mId(MyNumber, 6, 2)

    m = mId(MyNumber, 4, 2)

    Y = mId(MyNumber, 2, 2)

    Ty = left(MyNumber, 1)


    Select Case Ty

        Case "2": yy = Y

        Case "3": yy = "20" & Y

        Case Else: yy = ""

    End Select

    If yy <> "" Then GetBOFFromNatioanlID = DateSerial(yy, m, d)


ElseIf MyTest = 2 Then

    If left(right(MyNumber, 2), 1) Mod 2 = 1 Then _
    yy = "–ﬂ—" Else yy = "«‰ÀÏ"

    GetBOFFromNatioanlID = yy


ElseIf MyTest = 3 Then

    X = mId(MyNumber, 8, 2)

    For r = LBound(MyProvinces) To UBound(MyProvinces)

        xx = MyProvinces(r)

        If X = xx Then

            GetBOFFromNatioanlID = right(MyProvinces(r), Len(MyProvinces(r)) - 3)

            Exit For

        End If

    Next

End If

1:

End Function






''Public workWithBarcode As Boolean
'
'Dim sql                   As String
'Public PPointID           As Integer
'Public CurrentCashireID   As Integer
'Public X2600              As Date
'Public groupcodesPublic   As String
'Public Strforitems        As String
'Public StrforitemsCodes   As String
'Public Strforitemsnames   As String
'Public groupcodesAll      As String
'Public firstrun           As Boolean
'Public Report_Folder      As String
'Public myGrid             As VSFlexGrid
'Public P_DTPickerAccFrom  As Date
'Public P_DTPickerAccTo    As Date
'Public P_DCActivity       As Integer
'Public P_DCRegionID       As Integer
'Public P_dcBranch         As Integer
'Public onLineMOde         As Boolean
'Public onlineservername   As String
'Public onlineDataBasename As String
'Public onlinusername      As String
'Public onlinepassword     As String
'Public onlinebackground   As String
'
'Public TempPath           As String
'
'
'Private Const LOCALE_SSHORTDATE = &H1F
'Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
'Private Declare Function SetLocaleInfo _
'                Lib "kernel32" _
'                Alias "SetLocaleInfoA" (ByVal Locale As Long, _
'                                        ByVal LCType As Long, _
'                                        ByVal lpLCData As String) As Boolean
'Private Declare Function GetLocaleInfo _
'                Lib "kernel32" _
'                Alias "GetLocaleInfoA" (ByVal Locale As Long, _
'                                        ByVal LCType As Long, _
'                                        ByVal lpLCData As String, _
'                                        ByVal cchData As Long) As Long
'Private Declare Function GetTempPath _
'                Lib "kernel32" _
'                Alias "GetTempPathA" (ByVal nBufferLength As Long, _
'                                      ByVal lpBuffer As String) As Long
'Dim lLocal  As Long
'Dim Length  As Long
'Dim lLocal2 As Long
'Dim buf     As String * 1024
'
'Dim length2 As Long
'
'Dim buf2    As String * 1024
'
'Dim a
'
'Private Const E_POINTER               As Long = &H80004003
'Private Const S_OK                    As Long = 0
'Private Const INTERNET_MAX_URL_LENGTH As Long = 2048
'Private Const URL_ESCAPE_PERCENT      As Long = &H1000&
'
'Private Declare Function UrlEscape _
'                Lib "shlwapi" _
'                Alias "UrlEscapeA" (ByVal pszUrl As String, _
'                                    ByVal pszEscaped As String, _
'                                    ByRef pcchEscaped As Long, _
'                                    ByVal dwFlags As Long) As Long
'
'Private Declare Function UrlUnescape _
'                Lib "shlwapi" _
'                Alias "UrlUnescapeA" (ByVal pszUrl As String, _
'                                      ByVal pszUnescaped As String, _
'                                      ByRef pcchUnescaped As Long, _
'                                      ByVal dwFlags As Long) As Long
'
'Function printAnyreport(Optional sql As String, _
'                        Optional Reportname As String, _
'                        Optional StrReportTitle As String)
'
'    'Set rs = New ADODB.Recordset
'    'rs.Open SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    Dim MySQL       As String
'    Dim RsData      As New ADODB.Recordset
'    Dim xApp        As New CRAXDRT.Application
'    Dim xReport     As CRAXDRT.Report
'    Dim CViewer     As ClsReportViewer
'
'    Dim StrFileName As String
'    Dim Msg         As String
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'
'        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "" & Reportname & ".rpt"
'    Else
'        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "" & Reportname & "e.rpt"
'
'    End If
'
'    If Dir(StrFileName) = "" Then
'        MsgBox " not found reports " & StrFileName
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    Set RsData = New ADODB.Recordset
'    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    Screen.MousePointer = vbArrowHourglass
'    Set xReport = xApp.OpenReport(StrFileName)
'    xReport.Database.SetDataSource RsData
'
'    Dim cCompanyInfo As New ClsCompanyInfo
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
'
'        StrReportTitle = "" '& StrAccountName
'
'    Else
'
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
'
'        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
'        StrReportTitle = ""
'
'    End If
'
'    Dim total As String
'    Dim totl  As Double
'
'    xReport.reporttitle = StrReportTitle
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.title
'    xReport.ReportAuthor = App.title
'    Set CViewer = New ClsReportViewer
'    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql
'
'    RsData.Close
'    Set RsData = Nothing
'    Screen.MousePointer = vbDefault
'
'End Function
'
'
'Public Function ChkDateFormat() As Boolean
'    If SystemOptions.CheckDateFormatCorrect = False Then
'        ChkDateFormat = True
'        Exit Function
'    End If
'
'    lLocal = GetSystemDefaultLCID()
'    Length = GetLocaleInfo(3073, LOCALE_SSHORTDATE, buf, Len(buf))
'    ChkDateFormat = True
'
'    a = left$(buf, Length - 1)
'    If SetLocaleInfo(3073, LOCALE_SSHORTDATE, "dd/MM/yyyy") = False Then
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            MsgBox "dd/mm/yyyy  ÌÃ» Ÿ»ÿ  ‰”Ìﬁ «· «—ÌŒ"
'        Else
'            MsgBox "  Date Formate Must Changed To : dd/mm/yyyy       "
'        End If
'        ChkDateFormat = False
'        Exit Function
'    End If
'
'    length2 = GetLocaleInfo(3073, 32, buf2, Len(buf2))
'    a = left$(buf2, length2 - 1)
'    If SetLocaleInfo(3073, 32, "dd MMMM, yyyy") = False Then
'
'        'MsgBox "ÌÃ» Ÿ»ÿ  ‰”Ìﬁ «· «—ÌŒ"
'        If SystemOptions.UserInterface = ArabicInterface Then
'            MsgBox "dd/mm/yyyy  ÌÃ» Ÿ»ÿ  ‰”Ìﬁ «· «—ÌŒ"
'        Else
'            MsgBox "  Date Formate Must Changed To : dd/mm/yyyy       "
'        End If
'
'        ChkDateFormat = False
'        Exit Function
'
'    End If
'
'End Function
'
'Function checkonlinedate() As Boolean
'    On Error Resume Next
'    Dim FileName As String
'    FileName = App.path & "\OnLineServer.txt"
'    If Dir(FileName, vbNormal) = "" Then checkonlinedate = False
'    Exit Function
'
'    Open FileName For Input As #1
'
'    Do Until EOF(1)
'        Line Input #1, a
'
'        If a <> "" Then
'            VarSet = Split(a, "*", , vbTextCompare)
'
'            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
'
'                onlineservername = (VarSet(0))
'                onlineDataBasename = (VarSet(1))
'                onlinusername = (VarSet(2))
'                onlinepassword = (VarSet(3))
'                onlinebackground = (VarSet(4))
'
'                onLineMOde = True
'
'                checkonlinedate = True
'                Exit Function
'            End If
'        End If
'
'    Loop
'
'    Close #1
'    checkonlinedate = False
'End Function
'
'Function print_report3_HyperLink(Optional innerStrAccountCode As String, _
'                                 Optional innerStrAccountnAME As String)
'    On Error Resume Next
'    On Error GoTo ErrTrap
'    Dim sql                                                  As String
'    Dim RsData                                               As New ADODB.Recordset
'    Dim xApp                                                 As New CRAXDRT.Application
'    Dim xReport                                              As CRAXDRT.Report
'    Dim CViewer                                              As ClsReportViewer
'    Dim StrReportTitle                                       As String
'    Dim StrFileName                                          As String
'    Dim Msg                                                  As String
'    Dim AccountTypes                                         As Integer
'
'    Dim OpeningBalancebeformdateMinus1                       As Double
'    Dim OpeningBalancebeformStartCurrentyearTOFromDAteminus1 As Double
'    Dim NewOpinning                                          As Double
'    Dim OpeningBalance                                       As Double
'    Dim ProfitBalance                                        As Double
'    Dim Rs1                                                  As ADODB.Recordset
'    Set Rs1 = New ADODB.Recordset
'
'    Dim i                  As Integer
'    Dim BranchID           As String
'    Dim HideZeroBalance    As Integer
'    Dim openingBalanceDate As Date
'    Dim FromdateMinus1     As Date
'    Dim StartCurrentDate   As Date
'    Dim BrcnActivety       As String
'    FromdateMinus1 = DateAdd("d", -1, P_DTPickerAccFrom)
'    getFirstPeriodDateInthisYear2 openingBalanceDate
'    getFirstPeriodDateInthisYear StartCurrentDate
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        HideZeroBalance = MsgBox("Â·  —Ìœ «Œ›«¡ Õ”«»«  ’›—ÌÂ ‰⁄„ «„ ·« ", vbInformation + vbYesNoCancel)
'    Else
'        HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
'    End If
'
'    If HideZeroBalance = 2 Then
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'    Dim BranshesReg As String
'
'    If val(P_DCRegionID) <> 0 Then
'        BranshesReg = BranchRegion(CDbl(P_DCRegionID))
'    End If
'    If val(P_DCActivity) <> 0 Then
'        BrcnActivety = BrcnhActivityType(CDbl(P_DCActivity))
'    End If
'
'    updateprofitAccount val(P_DCActivity), val(P_dcBranch), P_DTPickerAccTo, BranshesReg
'
'    sql = " SELECT    ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name, Account_NameEng , debitBalance ="
'    sql = sql & "                         (SELECT     SUM(DEV_Value1)"
'    sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
'    sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
'    sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d"
'    sql = sql & "                                              WHERE      (d.Credit_Or_Debit = 0 AND d.RecordDate >= " & SQLDate(P_DTPickerAccFrom, True) & " AND d.RecordDate <= " & SQLDate(P_DTPickerAccTo, True) & ") AND d.Account_Code = A.Account_Code  and(d.Posted IS NULL)"
'    If val(P_DCActivity) <> 0 Then
'        sql = sql & " and d.branch_id in (" & BrcnActivety & ")"
'    End If
'
'    If val(P_DCRegionID) <> 0 Then
'        sql = sql & " and d.branch_id in (" & BranshesReg & ")"
'    End If
'
'    If val(P_dcBranch) <> 0 Then
'        sql = sql & " and d.branch_id =" & val(P_dcBranch) & ""
'    End If
'    sql = sql & "  ) x),"
'    sql = sql & "                    CreditBalance ="
'    sql = sql & "                        (SELECT     SUM(DEV_Value2)"
'    sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
'    sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
'    sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d1"
'    sql = sql & "                                                   WHERE     (d1.Credit_Or_Debit = 1 AND d1.RecordDate >= " & SQLDate(P_DTPickerAccFrom, True) & "  AND d1.RecordDate <= " & SQLDate(P_DTPickerAccTo, True) & ") AND d1.Account_Code = A.Account_Code and(d1.Posted IS NULL)"
'    If val(P_DCActivity) <> 0 Then
'        sql = sql & " and d1.branch_id in (" & BrcnActivety & ")"
'    End If
'    If val(P_DCRegionID) <> 0 Then
'        sql = sql & " and d1.branch_id in (" & BranshesReg & ")"
'    End If
'    If val(P_dcBranch) <> 0 Then
'        sql = sql & " and d1.branch_id =" & val(P_dcBranch) & ""
'    End If
'    sql = sql & " ) x),"
'    sql = sql & "                     OpeningBalance ="
'    sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
'    sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
'    sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
'    sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 AS do"
'    sql = sql & "                                                   WHERE     (  do.Account_Code = A.Account_Code and(do.Posted IS NULL)"
'    If val(P_DCActivity) <> 0 Then
'        sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
'    End If
'    If val(P_DCRegionID) <> 0 Then
'        sql = sql & " and do.branch_id in (" & BranshesReg & ")"
'    End If
'    If val(P_dcBranch) <> 0 Then
'        sql = sql & " and do.branch_id =" & val(P_dcBranch) & ""
'    End If
'    sql = sql & "  )) x),"
'    sql = sql & "    OpeningBalancebeformdateMinus1 ="
'    sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
'    sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
'    sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
'    sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
'    sql = sql & "                                                   WHERE     ( do.RecordDate >=" & SQLDate(openingBalanceDate, True) & " and   do.RecordDate <= " & SQLDate(FromdateMinus1, True) & ") AND do.Account_Code = A.Account_Code and(do.Posted IS NULL)"
'    If val(P_DCActivity) <> 0 Then
'        sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
'    End If
'    If val(P_DCRegionID) <> 0 Then
'        sql = sql & " and do.branch_id in (" & BranshesReg & ")"
'    End If
'    If val(P_dcBranch) <> 0 Then
'        sql = sql & " and do.branch_id =" & val(P_dcBranch) & ""
'    End If
'    sql = sql & " ) x),"
'    sql = sql & "                    OpeningBalancebeformStartCurrentyearTOFromDAteminus1 ="
'    sql = sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
'    sql = sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
'    sql = sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
'    sql = sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
'    sql = sql & "                                                   WHERE     (do.RecordDate >= " & SQLDate(StartCurrentDate, True) & " AND do.RecordDate < " & SQLDate(P_DTPickerAccFrom, True) & ") AND do.Account_Code = A.Account_Code and(do.Posted IS NULL) "
'    If val(P_DCActivity) <> 0 Then
'        sql = sql & " and do.branch_id in (" & BrcnActivety & ")"
'    End If
'    If val(P_DCRegionID) <> 0 Then
'        sql = sql & " and do.branch_id in (" & BranshesReg & ")"
'    End If
'    If val(P_dcBranch) <> 0 Then
'        sql = sql & " and do.branch_id =" & val(P_dcBranch) & ""
'    End If
'    sql = sql & " ) x)"
'    sql = sql & " FROM         ACCOUNTS A"
'    sql = sql & " WHERE     A.last_account = 1   "
'
'    sql = sql & " and (A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS"
'    sql = sql & "    Where 1 = 1"
'    StrAccountCode = (innerStrAccountCode)
'    If mId(StrAccountCode, Len(StrAccountCode), 1) = "G" Then
'        StrAccountCode = mId(StrAccountCode, 1, Len(StrAccountCode) - 1)
'
'    End If
'
'    If StrAccountCode <> "" Then
'
'        sql = sql & " and A.Account_Code like'" & StrAccountCode & "a%'"
'    End If
'
'    If val(P_DCActivity) <> 0 Then
'        sql = sql & " and branch_id in (" & BrcnActivety & ")"
'    End If
'    If val(P_DCRegionID) <> 0 Then
'        sql = sql & " and branch_id in (" & BranshesReg & ")"
'    End If
'    If val(P_dcBranch) <> 0 Then
'        sql = sql & " and branch_id =" & val(P_dcBranch) & ""
'    End If
'    sql = sql & "   )"
'    sql = sql & " or A.Account_Code in(select Account_Code from  DOUBLE_ENTREY_VOUCHERS1"
'    sql = sql & "    Where 1 = 1"
'
'    If StrAccountCode <> "" Then
'        sql = sql & " and A.Account_Code like'" & StrAccountCode & "a%'"
'    End If
'
'    If val(P_DCActivity) <> 0 Then
'        sql = sql & " and branch_id in (" & BrcnActivety & ")"
'    End If
'    If val(P_DCRegionID) <> 0 Then
'        sql = sql & " and branch_id in (" & BranshesReg & ")"
'    End If
'    If val(P_dcBranch) <> 0 Then
'        sql = sql & " and branch_id =" & val(P_dcBranch) & ""
'    End If
'    sql = sql & "   ))"
'
'    sql = sql & "order by Account_Serial "
'    If SystemOptions.UserInterface = ArabicInterface Then
'        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewSahyper.rpt"
'    Else
'        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TrialBalanceNewSaEhyper.rpt"
'    End If
'    If Dir(StrFileName) = "" Then
'        'GetMsgs 139, vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'    Set RsData = New ADODB.Recordset
'    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If RsData.BOF Or RsData.EOF Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        Else
'            Msg = "No Data"
'        End If
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'    Dim desc As String
'    Screen.MousePointer = vbArrowHourglass
'    Set xReport = xApp.OpenReport(StrFileName)
'    xReport.Database.SetDataSource RsData
'    Dim cCompanyInfo As New ClsCompanyInfo
'    If SystemOptions.UserInterface = ArabicInterface Then
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
'        StrReportTitle = "" '& StrAccountName
'    Else
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
'        StrReportTitle = ""
'    End If
'    desc = ""
'    If val(P_DCActivity) <> 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            desc = desc & "«·‰‘«ÿ " & ": " & P_DCActivity & Chr(13)
'        Else
'            desc = desc & "Region" & ": " & P_DCActivity & Chr(13)
'        End If
'    End If
'
'    If val(P_DCRegionID) <> 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            desc = desc & "··„‰ÿﬁ…" & ": " & P_DCRegionID & Chr(13)
'        Else
'            desc = desc & "Activity" & ": " & P_DCRegionID & Chr(13)
'        End If
'    End If
'    If val(P_dcBranch) <> 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            desc = desc & "··›—⁄" & ": " & P_dcBranch & Chr(13)
'        Else
'            desc = desc & "Branch" & ": " & P_dcBranch & Chr(13)
'        End If
'    End If
'    xReport.ParameterFields(3).AddCurrentValue user_name
'    If HideZeroBalance = 6 Then
'        xReport.ParameterFields(6).AddCurrentValue 1
'    Else
'        xReport.ParameterFields(6).AddCurrentValue 0
'    End If
'    If Not IsNull(P_DTPickerAccFrom) Then
'        'xReport.ParameterFields(4).AddCurrentValue "" + CStr(P_DTPickerAccFrom)
'    End If
'    If Not IsNull(DTPickerAccTo) Then
'        ' xReport.ParameterFields(5).AddCurrentValue ToDate(P_DTPickerAccTo)
'    End If
'    '  xReport.ParameterFields(7).AddCurrentValue desc
'    xReport.reporttitle = "  Õ·Ì· «·Õ”«» " & Chr(13) & innerStrAccountnAME & Chr(13) & "   «·› —… „‰   " & P_DTPickerAccFrom & Chr(13) & "   «·Ì " & P_DTPickerAccTo & Chr(13) & desc
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.title
'    xReport.ReportAuthor = App.title
'    Set CViewer = New ClsReportViewer
'    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql
'    RsData.Close
'    Set RsData = Nothing
'    Screen.MousePointer = vbDefault
'ErrTrap:
'End Function
'
'Public Function updateAccountsmanully(Ayear As Double)
'    Dim StrSQL       As String
'    Dim Account_Code As String
'
'    StrSQL = " update ACCOUNTS   "
'
'    StrSQL = StrSQL & " SET opening_balance= "
'
'    StrSQL = StrSQL & " ("
'    StrSQL = StrSQL & " SELECT       dbo.TblBalanceSheetDetails.AValue"
'    StrSQL = StrSQL & " FROM            dbo.TblBalanceSheetHeader INNER JOIN"
'    StrSQL = StrSQL & "                                                   dbo.TblBalanceSheetDetails ON dbo.TblBalanceSheetHeader.BalanceSheetHeaderid = dbo.TblBalanceSheetDetails.BalanceSheetHeaderid"
'    StrSQL = StrSQL & " Where ACCOUNTS.Account_Code = TblBalanceSheetDetails.Account_Code and  (dbo.TblBalanceSheetHeader.DYear = " & Ayear & ")"
'    StrSQL = StrSQL & "  )"
'    StrSQL = StrSQL & "   where  ACCOUNTS.Account_Code in ("
'    StrSQL = StrSQL & "   SELECT       dbo.TblBalanceSheetDetails.Account_Code"
'    StrSQL = StrSQL & "   FROM            dbo.TblBalanceSheetHeader INNER JOIN                                                   dbo.TblBalanceSheetDetails"
'    StrSQL = StrSQL & "   ON dbo.TblBalanceSheetHeader.BalanceSheetHeaderid = dbo.TblBalanceSheetDetails.BalanceSheetHeaderid"
'    ' StrSQL = StrSQL & "  Where (dbo.TblBalanceSheetHeader.DYear = " & Ayear & "  and Avalue<>0 )"
'    StrSQL = StrSQL & "  Where (dbo.TblBalanceSheetHeader.DYear = " & Ayear & "     )"
'
'    StrSQL = StrSQL & "  )"
'
'    Cn.Execute StrSQL
'
'End Function
'
'Public Function GetCarsREbenueAcountCode(Optional ID As Double) As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim ownerid As Long
'    Dim sql     As String
'    sql = "select AccountPaym from TblCarsData where id=" & ID & " "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetCarsREbenueAcountCode = IIf(IsNull(rs2("AccountPaym").value), "", rs2("AccountPaym").value)
'        'ownerid = IIf(IsNull(rs2("ownerid").value), "", rs2("ownerid").value)
'        '    If GetAqarAcountCode = "" Then
'        '           GetAqarAcountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", ownerid)
'        '    End If
'    Else
'
'        GetCarsREbenueAcountCode = ""
'    End If
'
'End Function
'Public Function GetCarsREbenueAcountCode2(Optional ID As Double) As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim ownerid As Long
'    Dim sql     As String
'    sql = "select DCOwner from TblCarsData where id=" & ID & " "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        'rs2.Close
'
'        GetCarsREbenueAcountCode2 = IIf(IsNull(rs2("DCOwner").value), "", rs2("DCOwner").value)
'        'ownerid = IIf(IsNull(rs2("ownerid").value), "", rs2("ownerid").value)
'        '    If GetAqarAcountCode = "" Then
'        '           GetAqarAcountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", ownerid)
'        '    End If
'    Else
'
'        GetCarsREbenueAcountCode2 = ""
'    End If
'
'End Function
'
'Public Function checkManulanoisExist(Optional Transaction_Type As Double, _
'                                     Optional Transaction_ID As Double, _
'                                     Optional CusID As Double, _
'                                     Optional ManualNO As String, _
'                                     Optional ByRef NoteSerial1 As String) As Boolean
'
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim ownerid As Long
'    Dim sql     As String
'    sql = "SELECT     Transaction_ID, Transaction_Type, CusID, ManualNO, NoteSerial1"
'    sql = sql & "  From dbo.transactions"
'    sql = sql & "  WHERE     (Transaction_Type = " & Transaction_Type & ") "
'    sql = sql & "   AND (Transaction_ID <> " & Transaction_ID & ")"
'    sql = sql & "   AND (CusID = " & CusID & ")"
'    sql = sql & "   AND (ManualNO = '" & ManualNO & "')"
'
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        checkManulanoisExist = True
'        NoteSerial1 = IIf(IsNull(rs2("NoteSerial1").value), "", rs2("NoteSerial1").value)
'    Else
'
'        NoteSerial1 = ""
'        checkManulanoisExist = False
'    End If
'
'End Function
'
'Public Function GetCarsFixedAssetID(Optional ID As Double) As Double
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim ownerid As Long
'    Dim sql     As String
'    sql = " SELECT     fixedAssetid From dbo.TblCarsData  Where ID =" & ID & " "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetCarsFixedAssetID = IIf(IsNull(rs2("fixedAssetid").value), 0, rs2("fixedAssetid").value)
'
'    Else
'
'        GetCarsFixedAssetID = 0
'    End If
'
'End Function
'
'Public Function GetAqarAcountCode(Optional ID As Double) As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim ownerid As Long
'    Dim sql     As String
'    sql = "select * from TblAqar where Aqarid=" & ID & " "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetAqarAcountCode = IIf(IsNull(rs2("AccounCode").value), "", rs2("AccounCode").value)
'        'ownerid = IIf(IsNull(rs2("ownerid").value), "", rs2("ownerid").value)
'        '    If GetAqarAcountCode = "" Then
'        '           GetAqarAcountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", ownerid)
'        '    End If
'    Else
'
'        GetAqarAcountCode = ""
'    End If
'
'End Function
'
'Public Function GetCurrencyCode(Optional ID As Double, Optional filed As String) As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim sql As String
'    sql = " SELECT        " & filed & " AS RetuensF"
'    sql = sql & " From dbo.currency"
'    sql = sql & " WHERE        (id = " & ID & ") "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetCurrencyCode = IIf(IsNull(rs2("RetuensF").value), 0, rs2("RetuensF").value)
'    Else
'        GetCurrencyCode = 0
'    End If
'End Function
'
'Public Function GetValueFiter(Optional ID As Double, Optional filed As String) As Double
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim sql As String
'    sql = " SELECT        SUM(" & filed & ") AS Value"
'    sql = sql & " From dbo.TblFiterWaiverDet2"
'    sql = sql & " WHERE        (MasterID = " & ID & ") "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetValueFiter = IIf(IsNull(rs2("Value").value), 0, rs2("Value").value)
'    Else
'        GetValueFiter = 0
'    End If
'End Function
'Public Function GetValueFiterHeader(Optional ID As Double, _
'                                    Optional filed As String) As Double
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim sql As String
'    sql = " SELECT        SUM(" & filed & ") AS Value"
'    sql = sql & " From dbo.TblFiterWaiver"
'    sql = sql & " WHERE        (ID = " & ID & ") "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetValueFiterHeader = IIf(IsNull(rs2("Value").value), 0, rs2("Value").value)
'    Else
'        GetValueFiterHeader = 0
'    End If
'End Function
'Public Function CheckAkarPayments(NoteID As Double) As Boolean
'    Dim s  As String
'    Dim rs As New ADODB.Recordset
'
'    s = "select * From Notes where NoteType=5   "
'
'    s = s & "   and not (  (akarid is null )  and   (IqarID2 is null )  and   (NoteOrBonID is null ) )  "
'
'    s = s & " and NoteID= " & NoteID
'
'    rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If Not rs.EOF Then
'
'        CheckAkarPayments = True
'    Else
'        CheckAkarPayments = False
'    End If
'
'    rs.Close
'End Function
'
'Public Function CheckAkarCashes(NoteID As Double) As Boolean
'    Dim s  As String
'    Dim rs As New ADODB.Recordset
'
'    s = "select * From Notes where NoteType=4    "
'
'    s = s & " and CashingType >= 7"
'
'    s = s & " and NoteID= " & NoteID
'
'    rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If Not rs.EOF Then
'
'        CheckAkarCashes = True
'    Else
'        CheckAkarCashes = False
'    End If
'
'    rs.Close
'End Function
'
'Public Function CheckUserNotPermAccounts(UserID As Double, _
'                                         AccountCode As String) As Boolean
'    Dim s  As String
'    Dim rs As New ADODB.Recordset
'
'    s = "select * From tblUserPermAccounts where UserId=" & UserID & "  and AccountCode='" & AccountCode & "'"
'
'    rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'
'        CheckUserNotPermAccounts = True
'    Else
'        CheckUserNotPermAccounts = False
'    End If
'
'    rs.Close
'End Function
'
'Public Function CheckAkarExpenses(NoteID As Double) As Boolean
'    Dim s  As String
'    Dim rs As New ADODB.Recordset
'
'    s = "select * From notes_all where notetype=3"
'    s = s & " and  not (ToPriodDateH is null)"
'    s = s & " and NoteID= " & NoteID
'
'    rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If Not rs.EOF Then
'
'        CheckAkarExpenses = True
'    Else
'        CheckAkarExpenses = False
'    End If
'
'    rs.Close
'End Function
'
'Public Function GetDepAccByEmp(ByVal mEmp As Integer, _
'                               Optional ByVal mAccountIndex As Integer = 1) As String
'    Dim s  As String
'    Dim rs As New ADODB.Recordset
'
'    s = "SELECT TblEmpDepartments.Account_Code" & mAccountIndex & " as AccountName"
'
'    s = s & " From TblEmpDepartments"
'    s = s & " LEFT OUTER JOIN TblEmployee AS te"
'    s = s & " ON  te.DepartmentID = TblEmpDepartments.DeparmentID"
'    s = s & " Where te.Emp_ID = " & val(mEmp)
'
'    rs.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If Not rs.EOF Then
'        'GetDepAccByEmp = (rs!AccountName & "")
'
'        If IsNull(rs!AccountName & "") Or Trim(rs!AccountName & "") = "" Then
'            GetDepAccByEmp = "NO account"
'            Exit Function
'        End If
'        If Not IsNull(rs!AccountName) Then
'            If CheckAccountToJE(rs!AccountName & "") = True Then
'                GetDepAccByEmp = Trim(rs!AccountName) & ""
'                Exit Function
'            Else
'                GetDepAccByEmp = "NO account"
'                Exit Function
'            End If
'
'        End If
'    End If
'    rs.Close
'End Function
'
'Public Function CheckWORKINposvATsCREEN() As Boolean
'    If SystemOptions.GeneralVoucherCreateSalesGE = True Then
'        CheckWORKINposvATsCREEN = True
'        Exit Function
'    End If
'
'    Dim rs     As ADODB.Recordset
'    Dim StrSQL As String
'
'    StrSQL = "SELECT    * FROM TblReCalVATPO "
'
'    Set rs = New ADODB.Recordset
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        CheckWORKINposvATsCREEN = True
'    Else
'        CheckWORKINposvATsCREEN = False
'    End If
'
'    rs.Close
'    Set rs = Nothing
'End Function
'
'Public Function CheckintenalRequstQty(Item_ID As Double, order_no As String) As Double
'
'    Dim rs     As ADODB.Recordset
'    Dim StrSQL As String
'
'    StrSQL = "SELECT     SUM(dbo.Transaction_Details.ShowQty) AS totalMoving"
'    StrSQL = StrSQL + "FROM         dbo.Transaction_Details INNER JOIN"
'    StrSQL = StrSQL + "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
'    StrSQL = StrSQL + "WHERE     (dbo.Transactions.Transaction_Type = 10) AND (dbo.Transactions.BillBasedOn = 1) AND (dbo.Transactions.order_no = '" & order_no & "' and Item_ID=" & Item_ID & ")"
'
'    Set rs = New ADODB.Recordset
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (rs.BOF Or rs.EOF) Then
'        CheckintenalRequstQty = IIf(IsNull(rs("totalMoving").value), 0, rs("totalMoving").value)
'    Else
'        CheckintenalRequstQty = 0
'    End If
'
'    rs.Close
'    Set rs = Nothing
'End Function
'
'Public Function GetPaymentBank() As Long
'
'    Dim rs     As ADODB.Recordset
'    Dim StrSQL As String
'
'    StrSQL = "Select *  From  TblPaymentType "
'    StrSQL = StrSQL + " Where bankid<>0"
'    Set rs = New ADODB.Recordset
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (rs.BOF Or rs.EOF) Then
'        GetPaymentBank = IIf(IsNull(rs("bankid").value), 0, rs("bankid").value)
'    Else
'        GetPaymentBank = 0
'    End If
'
'    rs.Close
'    Set rs = Nothing
'End Function
'
'Public Function GetTotalSales(Optional Transaction_Type As Integer, _
'                              Optional Fromdate As Variant, _
'                              Optional todate As Variant, _
'                              Optional BranshesActiv As String, _
'                              Optional BrnchIDes As String, _
'                              Optional BranchID As Integer = 0) As Double
'    Dim StrSQL As String
'    Dim rs2    As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    StrSQL = " SELECT      sum(dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice) AS totalsale"
'    StrSQL = StrSQL & "   FROM         dbo.Transactions INNER JOIN"
'    StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'    StrSQL = StrSQL & "  WHERE     (dbo.Transactions.Transaction_Type = " & Transaction_Type & ")"
'    If BrnchIDes <> "-1" And BrnchIDes <> "" Then
'        StrSQL = StrSQL & "   and dbo.Transactions.BranchId in (" & BrnchIDes & ")"
'    End If
'    If BranshesActiv <> "-1" And BranshesActiv <> "" Then
'        StrSQL = StrSQL & "   and dbo.Transactions.BranchId in (" & BranshesActiv & ")"
'    End If
'
'    If BranchID <> 0 Then
'        StrSQL = StrSQL & "   and dbo.Transactions.BranchId= " & BranchID
'    End If
'
'    If Not IsNull(Fromdate) Then
'        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date >=" & SQLDate(CDate(Fromdate), True) & ""
'    End If
'
'    If Not IsNull(todate) Then
'        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date <=" & SQLDate(CDate(todate), True) & ""
'    End If
'    rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetTotalSales = IIf(IsNull(rs2("totalsale").value), 0, rs2("totalsale").value)
'    Else
'        GetTotalSales = 0
'    End If
'End Function
'Public Function get_StoreBYPurchasePerson(PurchasePersonid As Double) As Double
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim sql As String
'    sql = "Select StoreID from TblStore where PurchasePersonid=" & PurchasePersonid
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount = 0 Then
'        get_StoreBYPurchasePerson = 0
'        Exit Function
'    End If
'    If IsNull(Rs3("StoreID").value) Then
'        get_StoreBYPurchasePerson = 0
'        Exit Function
'    End If
'    If Not IsNull(Rs3("StoreID").value) Then
'        get_StoreBYPurchasePerson = Rs3("StoreID").value
'        Exit Function
'    End If
'    Rs3.Close
'
'End Function
'
'Public Function GetTblProcessDEF(ProcessDEFID As Long, _
'                                 Optional ByRef ProcessName As String, _
'                                 Optional ByRef ProcessNameE As String, _
'                                 Optional ByRef UnitID As Integer)
'    Dim rs  As ADODB.Recordset
'    Dim Rs1 As ADODB.Recordset
'    On Error Resume Next
'    Dim sql As String
'    Dim str As String
'    Set Rs1 = New ADODB.Recordset
'    sql = "SELECT * from TblProcessDEF where TblProcessDEFID=" & ProcessDEFID
'    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Rs1.RecordCount > 0 Then
'
'        ProcessName = IIf(IsNull(Rs1("ProcessName").value), "", Rs1("ProcessName").value)
'        ProcessNameE = IIf(IsNull(Rs1("ProcessNameE").value), "", Rs1("ProcessNameE").value)
'        UnitID = IIf(IsNull(Rs1("UnitID").value), "", Rs1("UnitID").value)
'
'        If ProcessDEFID = 1 Then
'            BranchID = 1
'            StoreID = 1
'            BoxID = 1
'            BankID = 1
'        End If
'
'    Else
'
'        If UserID = 1 Then
'            BranchID = 1
'            StoreID = 1
'            BoxID = 1
'            BankID = 1
'
'        Else
'            BranchID = 0
'            StoreID = 0
'            BoxID = 0
'            BankID = 0
'            EmpID = 0
'
'        End If
'
'    End If
'    'If checkmanyBranches("") = True Then usertype = 0
'    'If checkmanyStores("") = True Then usertype = 0
'    Rs1.Close
'End Function
'
'Public Function PercentgValueAddedAllToBarcode(Optional RecDate As Date, _
'                                               Optional ItemID As Double, _
'                                               Optional Transe As Integer) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt "
'    sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'    sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'    sql = sql & " WHERE     (TblSettsReqLimK.SelectType=0 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) )"
'    sql = sql + " AND (  dbo.TblSettsReqLimKDet.typ = 1)  "
'    sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
'    sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        PercentgValueAddedAllToBarcode = IIf(IsNull(rs2("MultiPercentTxt").value), 0, rs2("MultiPercentTxt").value)
'    Else
'        PercentgValueAddedAllToBarcode = 0
'    End If
'End Function
'Public Function PercentgValueAddedBarcode(Optional RecDate As Date, _
'                                          Optional ItemID As Double, _
'                                          Optional Transe As Integer) As Double
'    Dim Percent As Double
'    Percent = 0
'    If CheckItemFreeVATToBarcode(RecDate, ItemID, Transe) = True Then
'        PercentgValueAddedBarcode = -1
'    Else
'        Percent = PercentgValueAddedAllToBarcode(RecDate, ItemID, Transe)
'        If Percent > 0 Then
'            PercentgValueAddedBarcode = Percent
'        Else
'            Percent = PercentgValueAddedGroupToBarcode(RecDate, ItemID, Transe)
'            If Percent > 0 Then
'                PercentgValueAddedBarcode = Percent
'            Else
'                Dim sql As String
'                Dim rs2 As ADODB.Recordset
'                Set rs2 = New ADODB.Recordset
'                sql = " SELECT     dbo.TblSettsReqLimKDet.PercentD"
'                sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'                sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'                sql = sql & " WHERE   TblSettsReqLimK.SelectType=2 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) and   (dbo.TblSettsReqLimKDet.Typ = 0 AND (dbo.TblSettsReqLimKDet.ItemID = " & ItemID & ")) "
'
'                sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
'                sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
'                rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'                If rs2.RecordCount > 0 Then
'                    PercentgValueAddedBarcode = IIf(IsNull(rs2("PercentD").value), 0, rs2("PercentD").value)
'                Else
'                    PercentgValueAddedBarcode = 0
'                End If
'            End If
'        End If
'    End If
'End Function
'
'Public Function CheckItemFreeVATToBarcode(Optional RecDate As Date, _
'                                          Optional ItemID As Double, _
'                                          Optional Transe As Integer) As Boolean
'    Dim sql As String
'    If PercentgValueAddedGroupFreeToBarcode(RecDate, ItemID, Transe) = True Then
'        CheckItemFreeVATToBarcode = True
'    Else
'        Dim rs2 As ADODB.Recordset
'        Set rs2 = New ADODB.Recordset
'        sql = " SELECT     dbo.TblSettsReqLimKDet.PercentD"
'        sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'        sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'        sql = sql & " WHERE     (dbo.TblSettsReqLimKDet.Typ = 5 AND (dbo.TblSettsReqLimKDet.ItemID = " & ItemID & ")) "
'        sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
'        sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
'        rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If rs2.RecordCount > 0 Then
'            CheckItemFreeVATToBarcode = True
'        Else
'            CheckItemFreeVATToBarcode = False
'        End If
'    End If
'End Function
'Public Function PercentgValueAddedGroupFreeToBarcode(Optional RecDate As Date, _
'                                                     Optional ItemID As Double, _
'                                                     Optional Transe As Integer) As Boolean
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt"
'    sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'    sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'    sql = sql & " WHERE     (TblSettsReqLimK.SelectType=1 and (TblSettsReqLimK.ItemStust=0 ) and  " & ItemID & " in ( SELECT     dbo.TblItems.ItemID"
'    sql = sql + " FROM         dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
'    sql = sql + "                      dbo.TblItems ON dbo.TblSettsReqLimKDet.GroupID = dbo.TblItems.GroupID"
'    sql = sql + " Where (dbo.TblSettsReqLimKDet.typ = 2)  And (dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID)))"
'    sql = sql + " AND ( dbo.TblSettsReqLimKDet.Typ = 1)  "
'    sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
'    sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        PercentgValueAddedGroupFreeToBarcode = True
'    Else
'        PercentgValueAddedGroupFreeToBarcode = False
'    End If
'End Function
'Public Function Get_movingreciveTransaction_ID(Optional ByRef ReturnID As Double) As Double
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim sql As String
'    sql = "SELECT     Transaction_ID"
'    sql = sql & " From dbo.transactions"
'    sql = sql & " Where (Transaction_Type = 11) And (ReturnID = " & ReturnID & ")"
'
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        Get_movingreciveTransaction_ID = IIf(IsNull(rs2("Transaction_ID").value), 0, rs2("Transaction_ID").value)
'
'    Else
'        Get_movingreciveTransaction_ID = 0
'    End If
'End Function
'
'Public Sub Get_TradingContractinfo(Optional ByRef TradingContractID As Double, _
'                                   Optional ByRef TContractCustID As Double, _
'                                   Optional Typed As Integer = 0)
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim sql As String
'    sql = "select ID,TContract_CustID  from Tbl_TradingContract"
'
'    sql = sql & " "
'    If Typed = 0 Then
'        sql = sql & " where id=" & TradingContractID & ""
'    Else
'        sql = sql & " where TContract_CustID='" & TContractCustID & "'"
'    End If
'    sql = sql & " And IsNull(IsCanceld,0) <> 1"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        TradingContractID = IIf(IsNull(rs2("id").value), 0, rs2("id").value)
'        TContractCustID = IIf(IsNull(rs2("TContract_CustID").value), "", rs2("TContract_CustID").value)
'    Else
'        TContractCustID = 0
'        TradingContractID = 0
'    End If
'End Sub
'
'Public Function CheckCusCredit2(LngCusID As Long, _
'                                SngOutValue As Single, _
'                                IntCheckType As Integer, _
'                                Optional Transaction_ID As Double, _
'                                Optional ByRef MsgRe As String, _
'                                Optional Typd As Integer = 0, _
'                                Optional TransDate As Date, _
'                                Optional IssueDate As Date) As Boolean
'
'    Dim rs                   As ADODB.Recordset
'    Dim StrSQL               As String
'    Dim SngCreditLiimt       As Single
'    Dim SngCreditLimitCredit As Single
'    Dim SngCusAccount        As Single
'    Dim Msg                  As String
'    Dim StrTemp              As String
'    Dim IntRes               As Integer
'    Dim DepitInterval        As Integer
'    Dim DepitIntervalID      As Integer
'    Dim NoDay                As Integer
'    'On Local Error GoTo ErrTra
'
'    StrSQL = "Select DepitIntervalID, DepitInterval,Account_Code,CreditLimit,CreditLimitCredit From TblCustemers Where CusID=" & LngCusID & ""
'    Set rs = New ADODB.Recordset
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (rs.BOF Or rs.EOF) Then
'        SngCreditLiimt = IIf(IsNull(rs("CreditLimit").value), 0, rs("CreditLimit").value)
'        SngCreditLimitCredit = IIf(IsNull(rs("CreditLimitCredit").value), 0, rs("CreditLimitCredit").value)
'        DepitInterval = IIf(IsNull(rs("DepitInterval").value), 0, rs("DepitInterval").value)
'        DepitIntervalID = IIf(IsNull(rs("DepitIntervalID").value), 0, rs("DepitIntervalID").value)
'    Else
'        CheckCusCredit2 = False
'        Exit Function
'    End If
'
'    If IntCheckType = 0 Then
'
'        '«·ﬂ‘› ⁄·Ï «‰ „œÌÊ‰Ì… «·⁄„Ì· ·‰  “Ìœ ⁄‰ «·Õœ «·„Õœœ ·Â
'        If SngCreditLiimt = 0 Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Msg = "Ì—ÃÏ  ”ÃÌ· »Ì«‰«  Õœ «·«∆ „«‰ ·Â– «·⁄„Ì·"
'            Else
'                Msg = "Please enter data of credit limit"
'            End If
'            'NO CreditLimit For this customer
'            MsgRe = Msg
'            CheckCusCredit2 = False
'
'            Exit Function
'        Else
'
'            '------------------------------------------------
'            '»⁄œ «·√” ⁄·«„ ⁄‰ —’Ìœ «·⁄„Ì·
'            '*******new code********************************************
'            Dim Account_Code As String
'            Dim FirstPeriod  As Date
'            getFirstPeriodDateInthisYear FirstPeriod
'
'            Account_Code = GetMyAccountCode("TblCustemers", "CusID", LngCusID)  '
'            SngCusAccount = GetActualAccountBalance(Account_Code, 0, FirstPeriod, Date)
'            SngCusAccount = SngCusAccount - GetSumOfGeForOneAccount(Account_Code, Transaction_ID, 0)
'            If DepitIntervalID = 1 Then
'                DepitInterval = DepitInterval * 30
'            ElseIf DepitIntervalID = 2 Then
'                DepitInterval = DepitInterval * 365
'            End If
'            NoDay = DateDiff("d", IssueDate, TransDate)
'            NoDay = Abs(NoDay)
'            '***************************************************\
'
'            If SngCusAccount >= 0 Then  '„œÌ‰
'                If (Abs(SngCusAccount) + SngOutValue) > SngCreditLiimt Or (NoDay > DepitInterval) Then
'                    If SystemOptions.UserInterface = ArabicInterface Then
'                        Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
'                        Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê›   ŒÿÏ  «·√∆ „«‰ «·Œ«’ »«·⁄„Ì·...!!!"
'                    Else
'                        Msg = "This process can not be allowed...!!!"
'                        Msg = Msg & Chr(13) & "will exceed the credit limit...!!!"
'                    End If
'
'                    If (Abs(SngCusAccount) + SngOutValue) > SngCreditLiimt Then
'                        ' Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
'                        '  Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê›   ŒÿÏ Õœ «·√∆ „«‰ «·Œ«’ »«·⁄„Ì·...!!!"
'                        Msg = Msg & Chr(13) & "------------------------------------------------"
'                        If SystemOptions.UserInterface = ArabicInterface Then
'                            Msg = Msg & Chr(13) & "Õœ ≈∆ „«‰ «·⁄„Ì· : " & SngCreditLiimt & " " & WriteNo(CStr(SngCreditLiimt), 0)
'                            Msg = Msg & Chr(13) & "«·—’Ìœ «·Õ«·Ï ··⁄„Ì·  ﬁ»· Â–Â «·Õ—ﬂ…: "
'                        Else
'                            Msg = Msg & Chr(13) & "credit limit of customer : " & SngCreditLiimt & " " & WriteNo(CStr(SngCreditLiimt), 0)
'                            Msg = Msg & Chr(13) & "current balance before this is process: "
'                        End If
'                        If SystemOptions.UserInterface = ArabicInterface Then
'                            If SngCusAccount > 0 Then
'                                StrTemp = Abs(SngCusAccount) & "(„œÌ‰)"
'                            ElseIf SngCusAccount < 0 Then
'                                StrTemp = Abs(SngCusAccount) & "(œ«∆‰)"
'                            Else
'                                StrTemp = "(Œ«·’)"
'                            End If
'                        Else
'                            If SngCusAccount > 0 Then
'                                StrTemp = Abs(SngCusAccount) & "(debt)"
'                            ElseIf SngCusAccount < 0 Then
'                                StrTemp = Abs(SngCusAccount) & "(credit)"
'                            Else
'                                StrTemp = "(Zero)"
'                            End If
'                        End If
'
'                        Msg = Msg & StrTemp
'                        If SystemOptions.UserInterface = ArabicInterface Then
'                            Msg = Msg & Chr(13) & "«·„»·€ «·„—«œ  ”ÃÌ·Â ⁄·Ï «·⁄„Ì· : " & SngOutValue
'                        Else
'                            Msg = Msg & Chr(13) & "The amount to be recorded on the customer : " & SngOutValue
'                        End If
'                        ' Msg = Msg & Chr(13) & ""
'                    End If
'                    '//////////////////
'                    If (NoDay > DepitInterval) Then
'                        ' Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
'                        If SystemOptions.UserInterface = EnglishInterface Then
'                            ' Msg = Msg & Chr(13) & "------------------------------------------------"
'                            Msg = Msg & Chr(13) & "credit period  : " & " " & DepitInterval
'                            ' Msg = Msg & StrTemp
'                            Msg = Msg & Chr(13) & "The period to be recorded on the customer  : " & NoDay
'                            '   Msg = Msg & Chr(13) & ""
'                        Else
'                            '   Msg = Msg & Chr(13) & "------------------------------------------------"
'                            Msg = Msg & Chr(13) & "„œ… ≈∆ „«‰ «·⁄„Ì· : " & " " & DepitInterval
'                            '  Msg = Msg & StrTemp
'                            Msg = Msg & Chr(13) & "«·„œ… «·„—«œ  ”ÃÌ·Â«  ··⁄„Ì· : " & NoDay
'
'                        End If
'                    End If
'                    Msg = Msg & Chr(13) & ""
'                    '/////////////
'
'                    If SystemOptions.SendToAprovedSalesBill = False And Typd = 1 Then
'                    Else
'                        MsgRe = Msg
'                        CheckCusCredit2 = False
'                        Exit Function
'                    End If
'
'                End If
'
'                '------------------------------------------------
'            End If
'
'        End If
'
'    ElseIf IntCheckType = 1 Then
'
'        '«·ﬂ‘› ⁄·Ï «‰ œ«∆‰Ì… «·⁄„Ì· ·‰  “Ìœ ⁄‰ «·Õœ «·„Õœœ ·Â
'        If SngCreditLimitCredit = 0 Then
'            'NO CreditLimit For this customer
'            CheckCusCredit2 = True
'            Exit Function
'        Else
'            'Set Rs = New ADODB.Recordset
'            SngCusAccount = GetCustomerAccount(LngCusID, True)
'
'            '------------------------------------------------
'            '»⁄œ «·√” ⁄·«„ ⁄‰ —’Ìœ «·⁄„Ì·
'            If SngCusAccount >= 0 Then '„œÌ‰
'                If (Abs(SngCusAccount) + SngOutValue) > SngCreditLimitCredit Then
'                    If SystemOptions.UserInterface = ArabicInterface Then
'                        Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
'                        Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê›   ŒÿÏ Õœ «·√∆ „«‰ («·œ«∆‰) «·Œ«’ »«·⁄„Ì· ...!!!"
'                        Msg = Msg & Chr(13) & "------------------------------------------------"
'                        Msg = Msg & Chr(13) & "Õœ ≈∆ „«‰ «·⁄„Ì· : " & SngCreditLimitCredit & " " & WriteNo(CStr(SngCreditLimitCredit), 0)
'                        Msg = Msg & Chr(13) & "«·—’Ìœ «·Õ«·Ï ··⁄„Ì· : "
'
'                        If SngCusAccount < 0 Then
'                            StrTemp = Abs(SngCusAccount) & "(œ«∆‰)"
'                        Else
'                            StrTemp = "(Œ«·’)"
'                        End If
'
'                        Msg = Msg & StrTemp
'                        Msg = Msg & Chr(13) & "«·„»·€ «·„—«œ  ”ÃÌ·Â ⁄·Ï «·⁄„Ì· : " & SngOutValue
'                    Else
'                        ' Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
'                        ' Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê›   ŒÿÏ Õœ «·√∆ „«‰ («·œ«∆‰) «·Œ«’ »«·⁄„Ì· ...!!!"
'                        Msg = "This process can not be allowed...!!!"
'                        Msg = Msg & Chr(13) & "will exceed the credit limit...!!!"
'
'                        Msg = Msg & Chr(13) & "credit limit of customer : " & SngCreditLiimt & " " & WriteNo(CStr(SngCreditLiimt), 0)
'                        Msg = Msg & Chr(13) & "current balance before this is process: "
'
'                        Msg = Msg & Chr(13) & "------------------------------------------------"
'                        Msg = Msg & Chr(13) & "credit limit of customer : " & SngCreditLimitCredit & " " & WriteNo(CStr(SngCreditLimitCredit), 0)
'                        Msg = Msg & Chr(13) & "current balance : "
'
'                        If SngCusAccount < 0 Then
'                            StrTemp = Abs(SngCusAccount) & "(credit)"
'                        Else
'                            StrTemp = "(zero)"
'                        End If
'
'                        Msg = Msg & StrTemp
'                        Msg = Msg & Chr(13) & "The amount to be recorded on the customer : " & SngOutValue
'
'                    End If
'                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                    CheckCusCredit2 = False
'                    Exit Function
'                End If
'            End If
'
'            '------------------------------------------------
'        End If
'    End If
'
'    CheckCusCredit2 = True
'    Exit Function
'ErrTrap:
'    CheckCusCredit2 = False
'End Function
'
'Public Sub GetProjectInf(Optional ByRef ProjectID As Double, _
'                         Optional ByRef ProjectCode As String, _
'                         Optional Typed As Integer = 0)
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim sql As String
'    sql = "select id,Fullcode  from projects"
'    If Typed = 0 Then
'        sql = sql & " where id=" & ProjectID & ""
'    Else
'        sql = sql & " where Fullcode='" & ProjectCode & "'"
'    End If
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        ProjectID = IIf(IsNull(rs2("id").value), 0, rs2("id").value)
'        ProjectCode = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
'    Else
'        ProjectCode = ""
'        ProjectID = 0
'    End If
'End Sub
'
'Public Function BrcnhActivityType(Optional ActivityTypeId As Double) As String
'    Dim i         As Integer
'    Dim BrnchIDes As String
'    BrnchIDes = "-1"
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim sql As String
'    sql = "  SELECT     branch_id"
'    sql = sql & " From dbo.TblBranchesData"
'    sql = sql & " Where (ActivityTypeId = " & ActivityTypeId & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        rs2.MoveFirst
'        For i = 1 To rs2.RecordCount
'            BrnchIDes = BrnchIDes & "," & IIf(IsNull(rs2("branch_id").value), -1, rs2("branch_id").value)
'            rs2.MoveNext
'        Next i
'    End If
'    BrcnhActivityType = BrnchIDes
'End Function
'
'Public Function BranchRegion(Optional RegionID As Double) As String
'    Dim i         As Integer
'    Dim BrnchIDes As String
'    BrnchIDes = "-1"
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim sql As String
'    sql = "  SELECT     branch_id"
'    sql = sql & " From dbo.TblBranchesData"
'    sql = sql & " Where (RegionID = " & RegionID & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        rs2.MoveFirst
'        For i = 1 To rs2.RecordCount
'            BrnchIDes = BrnchIDes & "," & IIf(IsNull(rs2("branch_id").value), -1, rs2("branch_id").value)
'            rs2.MoveNext
'        Next i
'    End If
'    BranchRegion = BrnchIDes
'End Function
'
'Public Function PercentgValueAddedAccounProject(Optional RecDate As Date, _
'                                                Optional ByRef flg As Integer, _
'                                                Optional BranchID As Double, _
'                                                Optional ByRef ForcedFlg As Integer) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     dbo.TblSettsReqLimK.PercentH "
'    sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'    sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'    sql = sql & " WHERE   (dbo.TblSettsReqLimK.ProjAccount=1 ) and (dbo.TblSettsReqLimKDet.Typ = 9) AND (dbo.TblSettsReqLimKDet.BranchID = " & BranchID & ")   "
'    sql = sql + " and      (" & SQLDate(RecDate, True) & "BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) and TblSettsReqLimK.AccOrTran=0 "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        flg = 1
'        PercentgValueAddedAccounProject = IIf(IsNull(rs2("PercentH").value), 0, rs2("PercentH").value)
'        ForcedFlg = 0
'    Else
'        PercentgValueAddedAccounProject = 0
'        flg = 0
'        ForcedFlg = 0
'    End If
'End Function
'
'Public Function CheckProjectAccountDept(Optional ByRef Account_Code As String) As Integer
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     expanses_account, Material_account, Salary_account, legal, AccountUnderImp, AcountGood"
'    sql = sql & " From dbo.Projects"
'    sql = sql & " WHERE    (AcountGood = N'" & Account_Code & "') or  (expanses_account = N'" & Account_Code & "') or (Salary_account = N'" & Account_Code & "') or (legal = N'" & Account_Code & "')  or (Material_account = N'" & Account_Code & "')or (AccountUnderImp = N'" & Account_Code & "')"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        CheckProjectAccountDept = 0
'    Else
'        CheckProjectAccountDept = CheckProjectAccountCredit(Account_Code)
'    End If
'End Function
'Public Function CheckProjectAccountCredit(Optional ByRef Account_Code As String) As Integer
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     REVENUE_account"
'    sql = sql & " From dbo.Projects"
'    sql = sql & " WHERE     (REVENUE_account = N'" & Account_Code & "')"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        CheckProjectAccountCredit = 1
'    Else
'        CheckProjectAccountCredit = -1
'    End If
'End Function
'
'Public Function GetIssuedQty(order_no As String, _
'                             Optional Transaction_ID As Double, _
'                             Optional StoreId2 As Double, _
'                             Optional Item_ID As Double, _
'                             Optional OldID As Double) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = "SELECT     SUM(dbo.Transaction_Details.ShowQty) AS ISSUEDQTY"
'    sql = sql & "  FROM         dbo.Transactions INNER JOIN"
'    sql = sql & "                        dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'    sql = sql & "  WHERE     (dbo.Transactions.Transaction_Type = 19) AND (dbo.Transactions.order_no = '" & order_no & "') AND (dbo.Transactions.BillBasedOn = 2) AND"
'    sql = sql & "                        (dbo.Transactions.Transaction_ID <> " & Transaction_ID & ")"
'    sql = sql & "        and                 (dbo.Transactions.StoreID = " & StoreId2 & ")"
'    If Item_ID <> 0 Then
'        sql = sql & "        and                 (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
'    End If
'    If OldID <> 0 Then
'        sql = sql & "        and                 (dbo.Transaction_Details.OldID = " & OldID & ")"
'    End If
'
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetIssuedQty = IIf(IsNull(rs2("ISSUEDQTY").value), 0, rs2("ISSUEDQTY").value)
'    Else
'        GetIssuedQty = 0
'    End If
'
'End Function
'Public Function GetCusIDByCarID(Optional ID As Double) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = "Select * from TblVendorCars where ID =" & ID & ""
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetCusIDByCarID = IIf(IsNull(rs2("CustomerID").value), 0, rs2("CustomerID").value)
'    Else
'        GetCusIDByCarID = 0
'    End If
'End Function
'
'Public Sub GetAccountTypeTrans(Optional ID As Double, _
'                               Optional ByRef AccountRevenue As String, _
'                               Optional ByRef AccountExpense As String)
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = "Select * from TblTypesTransport where ID =" & ID & ""
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        AccountRevenue = IIf(IsNull(rs2("AccountRevenue").value), "", rs2("AccountRevenue").value)
'        AccountExpense = IIf(IsNull(rs2("AccountExpense").value), "", rs2("AccountExpense").value)
'    Else
'        AccountExpense = ""
'        AccountRevenue = ""
'    End If
'End Sub
'
'Public Sub DeletInvoiceofCustomer(Optional CusID As Double, _
'                                  Optional Transaction_Date As Date)
'    Cn.Execute "delete from Transactions where Transaction_Date=" & SQLDate(Transaction_Date, True) & " and CusID=" & CusID & " and Transaction_Type=21"
'End Sub
'Public Function CheckCustomerTrans(Optional CusID As Double, _
'                                   Optional Transaction_Date As Date) As Boolean
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     CusID2"
'    sql = sql & " FROM         dbo.Transaction_Details where Transaction_ID in (select Transaction_ID from Transactions "
'    sql = sql & " WHERE     (Transaction_Type=21  and Transaction_Date = " & SQLDate(Transaction_Date, True) & ") ) and CusID2=" & CusID & " "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        CheckCustomerTrans = True
'    Else
'        CheckCustomerTrans = False
'    End If
'End Function
'Public Function GetMaxIDTransection(Optional Item_ID As Long, _
'                                    Optional UnitID As Long) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     MAX(dbo.Transaction_Details.ID) AS MaxID"
'    sql = sql & " FROM         dbo.Transaction_Details INNER JOIN"
'    sql = sql & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
'    sql = sql & "                      dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
'    sql = sql & " Where (dbo.Transaction_Details.Item_ID = " & Item_ID & ") And (dbo.Transaction_Details.UnitID = " & UnitID & ") And (dbo.transactions.Transaction_Type = 21)"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetMaxIDTransection = IIf(IsNull(rs2("MaxID").value), 0, rs2("MaxID").value)
'    Else
'        GetMaxIDTransection = 0
'    End If
'End Function
'Public Function GetLastPrice(Optional Item_ID As Long, Optional UnitID As Long) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = "SELECT     showPrice"
'    sql = sql & " From dbo.Transaction_Details"
'    sql = sql & " WHERE     (ID = " & GetMaxIDTransection(Item_ID, UnitID) & ") "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetLastPrice = IIf(IsNull(rs2("showPrice").value), 0, Round(rs2("showPrice").value, 2))
'    Else
'        GetLastPrice = 0
'    End If
'End Function
'Public Sub GetItemsInformation(Optional Fullcode As String, _
'                               Optional ByRef ItemID As Double, _
'                               Optional ByRef Name As String)
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     ItemID, ItemName, ItemNamee"
'    sql = sql & " From dbo.TblItems"
'    sql = sql & " WHERE     (Fullcode = N'" & Fullcode & "') or (barCodeNO = N'" & Fullcode & "')"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        ItemID = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Name = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
'        Else
'            Name = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
'        End If
'    Else
'        Name = ""
'        ItemID = 0
'    End If
'End Sub
'Public Sub GetUnitID(Optional UnitName As String, Optional ByRef UnitID As Double)
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     UnitID, UnitName, UnitNamee"
'    sql = sql & " FROM         dbo.TblUnites "
'    sql = sql & " WHERE     (UnitNamee = N'" & UnitName & "') or (UnitName = N'" & UnitName & "')"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        UnitID = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
'    Else
'        UnitID = 0
'    End If
'End Sub
'
'Public Function GetRegVATNo(Optional branch_id As Integer) As String
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     VATNO"
'    sql = sql & " From dbo.TblBranchesData"
'    sql = sql & " Where (branch_id = " & branch_id & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetRegVATNo = IIf(IsNull(rs2("VATNO").value), "", rs2("VATNO").value)
'    Else
'        GetRegVATNo = ""
'    End If
'End Function
'Public Function checkmanyBoxes(Optional ByRef str As String = "") As Boolean
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'    If SystemOptions.UserInterface = ArabicInterface Then
'        sql = " SELECT     dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName "
'    Else
'        sql = " SELECT     dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxNameE "
'    End If
'    sql = sql & " FROM         dbo.TblUsersBoxes  LEFT OUTER JOIN "
'    sql = sql & "                     dbo.TblBoxesData ON dbo.TblUsersBoxes.BoxId = dbo.TblBoxesData.BoxID"
'    If user_id <> 1 Then
'        sql = sql & "    Where (dbo.TblUsersBoxes.userid = " & user_id & ")"
'    Else
'        checkmanyBoxes = False
'        Exit Function
'    End If
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        str = sql
'        checkmanyBoxes = True
'
'    Else
'        checkmanyBoxes = False
'    End If
'
'End Function
'Public Function CheckItemFreeVAT(Optional RecDate As Date, _
'                                 Optional StoreID As Long, _
'                                 Optional ItemID As Double, _
'                                 Optional Transe As Integer) As Boolean
'    Dim sql As String
'    If PercentgValueAddedGroupFree(RecDate, StoreID, ItemID, Transe) = True Then
'        CheckItemFreeVAT = True
'    Else
'        Dim rs2 As ADODB.Recordset
'        Set rs2 = New ADODB.Recordset
'        sql = " SELECT     dbo.TblSettsReqLimKDet.PercentD"
'        sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'        sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'        sql = sql & " WHERE     (dbo.TblSettsReqLimKDet.Typ = 5 AND (dbo.TblSettsReqLimKDet.ItemID = " & ItemID & ")) "
'        sql = sql + " AND (dbo.TblSettsReqLimKDet.StoreID = " & StoreID & ")  "
'        sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
'        sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
'        rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If rs2.RecordCount > 0 Then
'            CheckItemFreeVAT = True
'        Else
'            CheckItemFreeVAT = False
'        End If
'    End If
'End Function
'
'Public Function PercentgValueAddedAll(Optional RecDate As Date, _
'                                      Optional StoreID As Long, _
'                                      Optional ItemID As Double, _
'                                      Optional Transe As Integer) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt "
'    sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'    sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'    sql = sql & " WHERE     (TblSettsReqLimK.SelectType=0 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) )"
'    sql = sql + " AND (dbo.TblSettsReqLimKDet.StoreID = " & StoreID & " and dbo.TblSettsReqLimKDet.typ = 1)  "
'    sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
'    sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        PercentgValueAddedAll = IIf(IsNull(rs2("MultiPercentTxt").value), 0, rs2("MultiPercentTxt").value)
'    Else
'        PercentgValueAddedAll = 0
'    End If
'End Function
'Public Function PercentgValueAddedGroupFree(Optional RecDate As Date, _
'                                            Optional StoreID As Long, _
'                                            Optional ItemID As Double, _
'                                            Optional Transe As Integer) As Boolean
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt"
'    sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'    sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'    sql = sql & " WHERE     (TblSettsReqLimK.SelectType=1 and (TblSettsReqLimK.ItemStust=0 ) and  " & ItemID & " in ( SELECT     dbo.TblItems.ItemID"
'    sql = sql + " FROM         dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
'    sql = sql + "                      dbo.TblItems ON dbo.TblSettsReqLimKDet.GroupID = dbo.TblItems.GroupID"
'    sql = sql + " Where (dbo.TblSettsReqLimKDet.typ = 2)  And (dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID)))"
'    sql = sql + " AND (dbo.TblSettsReqLimKDet.StoreID = " & StoreID & " and dbo.TblSettsReqLimKDet.Typ = 1)  "
'    sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
'    sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        PercentgValueAddedGroupFree = True
'    Else
'        PercentgValueAddedGroupFree = False
'    End If
'End Function
'
'Public Function PercentgValueAddedGroup(Optional RecDate As Date, _
'                                        Optional StoreID As Long, _
'                                        Optional ItemID As Double, _
'                                        Optional Transe As Integer) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt"
'    sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'    sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'    sql = sql & " WHERE     (TblSettsReqLimK.SelectType=1 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) and  " & ItemID & " in ( SELECT     dbo.TblItems.ItemID"
'    sql = sql + " FROM         dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
'    sql = sql + "                      dbo.TblItems ON dbo.TblSettsReqLimKDet.GroupID = dbo.TblItems.GroupID"
'    sql = sql + " Where (dbo.TblSettsReqLimKDet.typ = 2)  And (dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID)))"
'    sql = sql + " AND (dbo.TblSettsReqLimKDet.StoreID = " & StoreID & " and dbo.TblSettsReqLimKDet.Typ = 1)  "
'    sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
'    sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        PercentgValueAddedGroup = IIf(IsNull(rs2("MultiPercentTxt").value), 0, rs2("MultiPercentTxt").value)
'    Else
'        PercentgValueAddedGroup = 0
'    End If
'End Function
'
'Public Function PercentgValueAddedGroupToBarcode(Optional RecDate As Date, _
'                                                 Optional ItemID As Double, _
'                                                 Optional Transe As Integer) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     dbo.TblSettsReqLimK.MultiPercentTxt"
'    sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'    sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'    sql = sql & " WHERE     (TblSettsReqLimK.SelectType=1 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) and  " & ItemID & " in ( SELECT     dbo.TblItems.ItemID"
'    sql = sql + " FROM         dbo.TblSettsReqLimKDet LEFT OUTER JOIN"
'    sql = sql + "                      dbo.TblItems ON dbo.TblSettsReqLimKDet.GroupID = dbo.TblItems.GroupID"
'    sql = sql + " Where (dbo.TblSettsReqLimKDet.typ = 2)  And (dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID)))"
'    sql = sql + " AND ( dbo.TblSettsReqLimKDet.Typ = 1)  "
'    sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
'    sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        PercentgValueAddedGroupToBarcode = IIf(IsNull(rs2("MultiPercentTxt").value), 0, rs2("MultiPercentTxt").value)
'    Else
'        PercentgValueAddedGroupToBarcode = 0
'    End If
'End Function
'
'Public Sub GetAssestMoveYearly(Optional RecDate As Date, _
'                               Optional ByRef YearMove As Double, _
'                               Optional ByRef YearNotMove As Double, _
'                               Optional ByRef MonthMove As Double, _
'                               Optional ByRef MonthNotMove As Double)
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " select * from TblSettsReqLimK "
'    sql = sql + " where       (" & SQLDate(RecDate, True) & "BETWEEN RecordDate AND RecordDateTo) "
'    sql = sql + " and  AccOrTran = 1 and TransType= 11"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        YearMove = IIf(IsNull(rs2("YearMove").value), 0, rs2("YearMove").value)
'        YearNotMove = IIf(IsNull(rs2("YearNotMove").value), 0, rs2("YearNotMove").value)
'        MonthMove = IIf(IsNull(rs2("MonthMove").value), 0, rs2("MonthMove").value)
'        MonthNotMove = IIf(IsNull(rs2("MonthNotMove").value), 0, rs2("MonthNotMove").value)
'    Else
'        MonthMove = 0
'        MonthNotMove = 0
'        YearMove = 0
'        YearNotMove = 0
'    End If
'End Sub
'Public Function GetCashCustomerPhoneByName(CashCustomerName As String) As String
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'    sql = "SELECT     CashCustomerName, CashCustomerPhone From dbo.Transactions  WHERE     (CashCustomerName = '" & CashCustomerName & "')"
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'        GetCashCustomerPhoneByName = IIf(IsNull(rs("CashCustomerPhone").value), "", rs("CashCustomerPhone").value)
'    Else
'        GetCashCustomerPhoneByName = ""
'    End If
'    rs.Close
'End Function
'Public Function GetItemUnitsId(Optional UnitName As String) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     UnitID, UnitName, UnitNamee"
'    sql = sql & " FROM         dbo.TblUnites where UnitName='" & UnitName & " '"
'    sql = sql & " or    UnitNamee='" & UnitName & " '"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetItemUnitsId = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
'    Else
'        GetItemUnitsId = 0
'    End If
'End Function
'
'Public Sub PercentgValueAddedAccount_Transec(Optional RecDate As Date, _
'                                             Optional TransType As Integer, _
'                                             Optional Dept_Credit As Integer, _
'                                             Optional ByRef AccountCode As String, _
'                                             Optional ByRef Percentage As Double)
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " select * from TblSettsReqLimK "
'    sql = sql + " where       (" & SQLDate(RecDate, True) & "BETWEEN RecordDate AND RecordDateTo) "
'    sql = sql + " and  AccOrTran = 1 and TransType= " & TransType & ""
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If rs2.RecordCount > 0 Then
'        Percentage = IIf(IsNull(rs2("PercentH").value), 0, rs2("PercentH").value)
'        If Dept_Credit = 0 Then
'            AccountCode = IIf(IsNull(rs2("AccDep").value), "", rs2("AccDep").value)
'        Else
'            AccountCode = IIf(IsNull(rs2("AccCir").value), "", rs2("AccCir").value)
'        End If
'    Else
'        AccountCode = ""
'        Percentage = 0
'    End If
'End Sub
'Public Function GetItemUnits(Optional UnitID As Double) As String
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     UnitID, UnitName, UnitNamee"
'    sql = sql & " FROM         dbo.TblUnites where UnitID=" & UnitID & " "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            GetItemUnits = IIf(IsNull(rs2("UnitName").value), "", rs2("UnitName").value)
'        Else
'            GetItemUnits = IIf(IsNull(rs2("UnitNamee").value), "", rs2("UnitNamee").value)
'        End If
'    Else
'        GetItemUnits = ""
'    End If
'End Function
'Public Function CheckCustomerCont(Optional customerid As Double, _
'                                  Optional RecDate As Date) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     TblCustomerContractD"
'    sql = sql & " From dbo.TblCustomerContract"
'    sql = sql & " WHERE     (Locked = 0 OR Locked IS NULL)"
'    sql = sql & " and CustomerId=" & customerid & ""
'    sql = sql & " and FromDate <=" & SQLDate(RecDate, True) & ""
'    sql = sql & " and Todate >=" & SQLDate(RecDate, True) & ""
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        CheckCustomerCont = IIf(IsNull(rs2("TblCustomerContractD").value), 0, rs2("TblCustomerContractD").value)
'    Else
'        CheckCustomerCont = False
'    End If
'End Function
'Public Function CheckWorkState(UserID As Integer) As Boolean
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    sql = " SELECT     dbo.TblUsers.UserID, dbo.TblEmployee.workstate"
'    sql = sql & " FROM         dbo.TblUsers LEFT OUTER JOIN"
'    sql = sql & "                      dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
'    sql = sql & " Where (dbo.TblUsers.UserID = " & UserID & ") And (dbo.TblEmployee.WorkState = 1)"
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        CheckWorkState = True
'    Else
'        CheckWorkState = False
'    End If
'End Function
'
'Public Function GetItemsTotalByStore(Optional Transaction_ID As Long, _
'                                     Optional StoreID As Integer) As Double
'    Dim DblTemp  As Double
'    Dim RowNum   As Long
'    Dim Msg      As String
'    Dim sql      As String
'    Dim linetotl As Double
'    Dim rs2      As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    On Local Error GoTo ErrTrap
'    sql = " SELECT     SUM(ShowQty * showPrice) AS Price"
'    sql = sql & "      From dbo.Transaction_Details"
'    sql = sql & "  WHERE     (StoreID2 = " & StoreID & ") AND (Transaction_ID = " & Transaction_ID & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        linetotl = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
'        If SystemOptions.PursgaseWithoutDecimal = True And GridTrans = PurchaseTransaction Then
'            DblTemp = DblTemp + Int(linetotl)
'        Else
'            DblTemp = DblTemp + Round(linetotl, SystemOptions.SysDefCurrencyForamt)
'        End If
'    Else
'        linetotl = 0
'    End If
'
'    GetItemsTotalByStore = linetotl
'    Exit Function
'ErrTrap:
'    Msg = "ERROR "
'    GetItemsTotalByStore = linetotl
'End Function
'
'Public Function CheckExpeIqar(Optional NoteID As Double) As Integer
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim StrSQL As String
'    StrSQL = "select * From notes_all where notetype=3"
'    StrSQL = StrSQL & " and  not (ToPriodDateH is null) and NoteID=" & NoteID & ""
'    rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        CheckExpeIqar = 1
'    Else
'        CheckExpeIqar = 0
'    End If
'End Function
'Public Function GetIDesUnpadiVacation(Optional Emp_id As Double) As String
'    Dim sql     As String
'    Dim StrIDes As String
'    Dim i       As Integer
'    Dim NoDay   As Double
'    NoDay = 0
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    StrIDes = "0,0"
'    sql = " SELECT     id, MoveVacBalance"
'    sql = sql & " From dbo.TblEmbarkation"
'    sql = sql & " WHERE     (TypeVacation = 1) AND (VacationPaied IS NULL OR  VacationPaied = 0)"
'    sql = sql & "  AND (Emp_ID = " & Emp_id & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        rs2.MoveFirst
'        For i = 1 To rs2.RecordCount
'            StrIDes = StrIDes & "," & IIf(IsNull(rs2("id").value), 0, rs2("id").value)
'            rs2.MoveNext
'        Next i
'    End If
'    GetIDesUnpadiVacation = StrIDes
'End Function
'Public Function GetNoDayUnpadiVacation2(Optional Emp_id As Double, _
'                                        Optional RdTypeVaction As Integer = 0) As Double
'    Dim sql   As String
'    Dim i     As Integer
'    Dim NoDay As Double
'    NoDay = 0
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'
'    sql = " SELECT     id, MoveVacBalance"
'    sql = sql & " From dbo.TblEmbarkation"
'    sql = sql & " WHERE     (TypeVacation = 1) AND (VacationPaied IS NULL OR  VacationPaied = 0)"
'    sql = sql & "  AND (Emp_ID = " & Emp_id & ")"
'    If RdTypeVaction = 1 Then
'        sql = sql & "  AND (RdTypeVaction = 1)"
'    Else
'        sql = sql & "  AND (RdTypeVaction = 0 or RdTypeVaction is null)"
'    End If
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        rs2.MoveFirst
'        For i = 1 To rs2.RecordCount
'            NoDay = NoDay + IIf(IsNull(rs2("MoveVacBalance").value), 0, rs2("MoveVacBalance").value)
'            rs2.MoveNext
'        Next i
'    End If
'    GetNoDayUnpadiVacation2 = NoDay
'End Function
'
'Public Function GetMaxIDVation(Optional EmpID As Double) As Double
'    If EmpID = 0 Then
'    GetMaxIDVation = 0
'    Exit Function
'    End If
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     MAX(ID) AS MaxID"
'    sql = sql & " From dbo.TblVocationEntitlements"
'    sql = sql & " Where (EmpID = " & EmpID & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetMaxIDVation = IIf(IsNull(rs2("MaxID").value), 0, rs2("MaxID").value)
'    Else
'        GetMaxIDVation = 0
'    End If
'End Function
'
'Public Function GetLastBalanceMonthVaction(Optional EmpID As Double, _
'                                           Optional ID As Double) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     LastBalanceMonth"
'    sql = sql & " From dbo.TblVocationEntitlements"
'    sql = sql & " WHERE     (ID = " & GetMaxIDVation(EmpID) & ") and ID<>" & ID & " "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetLastBalanceMonthVaction = IIf(IsNull(rs2("LastBalanceMonth").value), 0, rs2("LastBalanceMonth").value)
'    Else
'        GetLastBalanceMonthVaction = 0
'    End If
'End Function
'
'Public Function GetEmIDUnpaidVacation(Optional ID As Double) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     Emp_id"
'    sql = sql & " From dbo.TblEmpPassOver"
'    sql = sql & " Where (TypeTrans = 3) And (advanceID = " & ID & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetEmIDUnpaidVacation = IIf(IsNull(rs2("Emp_id").value), 0, rs2("Emp_id").value)
'    Else
'        GetEmIDUnpaidVacation = 0
'    End If
'End Function
'Public Sub GetNoDayUnpadiVacation(Optional Emp_id As Double, _
'                                  Optional ByRef IDes As String, _
'                                  Optional ByRef NoVaction As Double)
'    Dim sql     As String
'    Dim StrIDes As String
'    Dim i       As Integer
'    Dim NoDay   As Double
'    NoDay = 0
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    StrIDes = "0,0"
'    sql = " SELECT     id, MoveVacBalance"
'    sql = sql & " From dbo.TblEmbarkation"
'    sql = sql & " WHERE     (TypeVacation = 1) AND (VacationPaied IS NULL OR  VacationPaied = 0)"
'    sql = sql & "  AND (Emp_ID = " & Emp_id & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        rs2.MoveFirst
'        For i = 1 To rs2.RecordCount
'            StrIDes = StrIDes & "," & IIf(IsNull(rs2("id").value), 0, rs2("id").value)
'            NoDay = NoDay + IIf(IsNull(rs2("MoveVacBalance").value), 0, rs2("MoveVacBalance").value)
'            rs2.MoveNext
'        Next i
'    End If
'    IDes = StrIDes
'    NoVaction = NoDay
'End Sub
'Public Function GetTrnasectionID(Optional MainTransaction_ID As Double, _
'                                 Optional Transaction_Type As Integer)
'    Dim StrIDes As String
'    Dim sql     As String
'    Dim rs2     As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = "select * from TblTransctionIDES where MainTransaction_ID=" & MainTransaction_ID & "and Transaction_Type=" & Transaction_Type & " "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    StrIDes = "0,0"
'    If rs2.RecordCount > 0 Then
'        rs2.MoveFirst
'        For i = 1 To rs2.RecordCount
'            StrIDes = StrIDes & "," & IIf(IsNull(rs2("Transaction_ID").value), 0, rs2("Transaction_ID").value)
'            rs2.MoveNext
'        Next i
'    End If
'    GetTrnasectionID = StrIDes
'End Function
'Public Sub SaveTrnasectionID(Optional MainTransaction_ID As Double, _
'                             Optional Transaction_ID As Long, _
'                             Optional Transaction_Type As Integer)
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " select * from TblTransctionIDES where 1=-1"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    rs2.AddNew
'    rs2("MainTransaction_ID").value = MainTransaction_ID
'    rs2("Transaction_ID").value = Transaction_ID
'    rs2("Transaction_Type").value = Transaction_Type
'    rs2.update
'End Sub
'Public Function CheckSettingsLikeContract() As Boolean
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    sql = "Select CommContract from TblVacationSettings where CommContract=1 "
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        CheckSettingsLikeContract = True
'    Else
'        CheckSettingsLikeContract = False
'    End If
'End Function
'
'Public Function GetAccountCodeHiding() As String
'    Dim My_SQL        As String
'    Dim FlgBign       As Boolean
'    Dim Account_Code5 As String
'    Dim rs2           As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    My_SQL = " SELECT    AccountCode"
'    My_SQL = My_SQL & " From AccountSetting"
'    My_SQL = My_SQL & " where     TreeAccount = 1"
'    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        rs2.MoveFirst
'        My_SQL = " ( SELECT     Account_Code"
'        My_SQL = My_SQL & " From dbo.Accounts"
'        FlgBign = False
'        For i = 1 To rs2.RecordCount
'            Account_Code5 = IIf(IsNull(rs2("AccountCode").value), "", rs2("AccountCode").value)
'            If Account_Code5 <> "" Then
'                If FlgBign = False Then
'                    FlgBign = True
'                    My_SQL = My_SQL & " where   (Account_Code LIKE N'" & Account_Code5 & "%') and Account_Code<>'" & Account_Code5 & "' "
'                Else
'                    My_SQL = My_SQL & " or  (Account_Code LIKE N'" & Account_Code5 & "%')and Account_Code<>'" & Account_Code5 & "' "
'                End If
'            End If
'
'            rs2.MoveNext
'        Next i
'    Else
'        GetAccountCodeHiding = ""
'        Exit Function
'    End If
'    My_SQL = My_SQL & ")"
'    GetAccountCodeHiding = My_SQL
'
'    GetAccountCodeHiding = " and ACCOUNTS.Account_Code not in " & My_SQL
'End Function
'
'Public Function GetValueAddedAccount(Optional RecDate As Date, _
'                                     Optional ByRef Account_CodeDept As String, _
'                                     Optional ByRef Account_CodeCridit As String, _
'                                     Optional Trans_Account As Integer = 0, _
'                                     Optional TransType As Integer) As Boolean
'    If mdifrmmain.taxes = False Then
'        GetValueAddedAccount = True
'    ElseIf CheckAnyVAT(RecDate) = False Then
'        GetValueAddedAccount = True
'    Else
'        Dim sql As String
'        GetValueAddedAccount = False
'        Dim rs2 As ADODB.Recordset
'        Set rs2 = New ADODB.Recordset
'        sql = " SELECT     AccDep ,AccCir"
'        sql = sql & " FROM       TblSettsReqLimK"
'        sql = sql + "  WHERE     (" & SQLDate(RecDate, True) & "BETWEEN RecordDate AND RecordDateTo) "
'        If Trans_Account = 1 Then
'            sql = sql + " and      ((AccOrTran = 1) OR  (AccOrTran IS NULL)) and TransType=" & TransType & " "
'        Else
'            sql = sql + " and      (AccOrTran = 0)"
'        End If
'
'        rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If rs2.RecordCount > 0 Then
'            Account_CodeDept = IIf(IsNull(rs2("AccDep").value), "", rs2("AccDep").value)
'            Account_CodeCridit = IIf(IsNull(rs2("AccCir").value), "", rs2("AccCir").value)
'            GetValueAddedAccount = True
'        Else
'            GetValueAddedAccount = False
'            Account_CodeCridit = ""
'            Account_CodeDept = ""
'        End If
'    End If
'End Function
'
'Public Function CheckAnyVAT(Optional RecDate As Date) As Boolean
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = "select * from TblSettsReqLimK where 1=1 "
'    sql = sql + " and      (" & SQLDate(RecDate, True) & "BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) or dbo.TblSettsReqLimK.RecordDateTo<=" & SQLDate(RecDate, True) & " "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        CheckAnyVAT = True
'    Else
'        CheckAnyVAT = False
'    End If
'End Function
'
'Public Function ScreenAproved(Optional Transaction_ID As Double, _
'                              Optional ScreenName As String) As Boolean
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim sql As String
'    If SystemOptions.CaNUpdateApprovedDoc = True Then
'        ScreenAproved = False
'        Exit Function
'    End If
'    If CheckAprroveScreen(ScreenName) = True Then
'        If Transaction_ID = 0 Then
'            sql = "Select * from ApprovalData where ScreenName='" & ScreenName
'        Else
'            sql = "Select * from ApprovalData where ScreenName='" & ScreenName & "' and Transaction_ID =" & Transaction_ID & ""
'        End If
'        sql = sql & " and (NOT (ApprovDate IS NULL) or NOT (CancelApprove IS NULL) )"
'        rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If rs2.RecordCount > 0 Then
'            ScreenAproved = True
'        Else
'            ScreenAproved = False
'        End If
'    Else
'        ScreenAproved = False
'    End If
'End Function
'Public Function MainCurrency() As Integer
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = "Select ID from currency where basic=1 "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        MainCurrency = IIf(IsNull(rs2("ID").value), 1, rs2("ID").value)
'    Else
'        MainCurrency = 1
'    End If
'End Function
'
'Public Sub getinsttPayedToContNote(Optional NoteID As Double = 0, _
'                                   Optional ByRef RentValuePayed As Double, _
'                                   Optional ByRef CommissionsPayed As Double, _
'                                   Optional ByRef InsurancePayed As Double, _
'                                   Optional ByRef WaterPayed As Double, _
'                                   Optional ByRef ElectricPayed As Double, _
'                                   Optional ByRef TelandNetPayed As Double, _
'                                   Optional ByRef TotalOldValue As Double, _
'                                   Optional Istallid As Double, _
'                                   Optional ByRef VATPayed As Double)
'    On Error Resume Next
'
'    Dim total As Single
'
'    Dim Rs3   As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim sql As String
'
'    sql = "select (value) As total   ,(RentValuePayed) as RentValuePayed,(CommissionsPayed) as CommissionsPayed"
'    sql = sql & "  ,(InsurancePayed) as InsurancePayed,(WaterPayed) as WaterPayed"
'    sql = sql & "  ,(ElectricPayed) as ElectricPayed,(TelandNetPayed) as TelandNetPayed ,(OldValuePayed) as TotalOldValue ,(VATPayed) as VATPayed"
'    sql = sql & "  from ContracttBillInstallmentsDone  where NoteID=" & NoteID & "  and Istallid=" & Istallid & ""
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount = 0 Then
'        total = 0
'        RentValuePayed = 0
'        CommissionsPayed = 0
'        InsurancePayed = 0
'        WaterPayed = 0
'        ElectricPayed = 0
'        TelandNetPayed = 0
'        TotalOldValue = 0
'        VATPayed = 0
'    Else
'
'        total = IIf(IsNull(Rs3("total").value), 0, Rs3("total").value)
'        RentValuePayed = IIf(IsNull(Rs3("RentValuePayed").value), 0, Rs3("RentValuePayed").value)
'        CommissionsPayed = IIf(IsNull(Rs3("CommissionsPayed").value), 0, Rs3("CommissionsPayed").value)
'        InsurancePayed = IIf(IsNull(Rs3("InsurancePayed").value), 0, Rs3("InsurancePayed").value)
'        WaterPayed = IIf(IsNull(Rs3("WaterPayed").value), 0, Rs3("WaterPayed").value)
'        ElectricPayed = IIf(IsNull(Rs3("ElectricPayed").value), 0, Rs3("ElectricPayed").value)
'        TelandNetPayed = IIf(IsNull(Rs3("TelandNetPayed").value), 0, Rs3("TelandNetPayed").value)
'        TotalOldValue = IIf(IsNull(Rs3("TotalOldValue").value), 0, Rs3("TotalOldValue").value)
'        VATPayed = IIf(IsNull(Rs3("VATPayed").value), 0, Rs3("VATPayed").value)
'    End If
'
'    Rs3.Close
'
'End Sub
'Public Function GetMosim(Optional Omra_Hajj As Integer) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     ID "
'    sql = sql & " FROM         dbo.TblCompaniesGroup"
'    sql = sql & " where CurrYear=1 "
'    If Omra_Hajj = 0 Then
'        sql = sql & " and Omra_Hajj=0 "
'    Else
'        sql = sql & " and Omra_Hajj=1 "
'    End If
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetMosim = IIf(IsNull(rs2("ID").value), 0, rs2("ID").value)
'    Else
'        GetMosim = 0
'    End If
'End Function
'Public Function GetIDOrder(Optional NoteSerial1 As Double, _
'                           Optional SeasonsID As Double) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " select ID from tblbookingrequest where NoteSerial1=" & NoteSerial1 & " and SeasonsID=" & SeasonsID & ""
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        GetIDOrder = IIf(IsNull(rs2("ID").value), 0, rs2("ID").value)
'    Else
'        GetIDOrder = 0
'    End If
'End Function
'
'Public Function GetCustomerVAT(CusID As Integer) As String
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = " SELECT   * from TblCustemers  "
'
'    sql = sql & " Where (CusID = " & CusID & ")"
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        GetCustomerVAT = IIf(IsNull(rs("VATNO").value), "", rs("VATNO").value)
'    Else
'        GetCustomerVAT = ""
'    End If
'
'End Function
'Public Function PercentgValueAddedAccount(Optional RecDate As Date, _
'                                          Optional Account_Code As String, _
'                                          Optional BranchID As Double, _
'                                          Optional ByRef ForcedFlg As Integer) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     dbo.TblSettsReqLimKDet.PercentD ,dbo.TblSettsReqLimKDet.ForcedFlg"
'    sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'    sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'    sql = sql & " WHERE     (dbo.TblSettsReqLimKDet.Typ = 9) AND (dbo.TblSettsReqLimKDet.BranchID = " & BranchID & ") AND (dbo.TblSettsReqLimKDet.Account_Code = '" & Account_Code & "')  "
'    sql = sql + " and      (" & SQLDate(RecDate, True) & "BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) and TblSettsReqLimK.AccOrTran=0 "
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        PercentgValueAddedAccount = IIf(IsNull(rs2("PercentD").value), 0, rs2("PercentD").value)
'        If Not IsNull(rs2("ForcedFlg").value) Then
'            If (rs2("ForcedFlg").value) = True Then
'                ForcedFlg = 1
'            Else
'                ForcedFlg = 0
'            End If
'        Else
'            ForcedFlg = 0
'        End If
'    Else
'        PercentgValueAddedAccount = 0
'        ForcedFlg = 0
'    End If
'End Function
'
'Public Function PercentgValueAdded(Optional RecDate As Date, _
'                                   Optional StoreID As Long, _
'                                   Optional ItemID As Double, _
'                                   Optional Transe As Integer) As Double
'    Dim Percent As Double
'    Percent = 0
'    If CheckItemFreeVAT(RecDate, StoreID, ItemID, Transe) = True Then
'        PercentgValueAdded = -1
'    Else
'        Percent = PercentgValueAddedAll(RecDate, StoreID, ItemID, Transe)
'        If Percent > 0 Then
'            PercentgValueAdded = Percent
'        Else
'            Percent = PercentgValueAddedGroup(RecDate, StoreID, ItemID, Transe)
'            If Percent > 0 Then
'                PercentgValueAdded = Percent
'            Else
'                Dim sql As String
'                Dim rs2 As ADODB.Recordset
'                Set rs2 = New ADODB.Recordset
'                sql = " SELECT     dbo.TblSettsReqLimKDet.PercentD"
'                sql = sql & " FROM         dbo.TblSettsReqLimKDet RIGHT OUTER JOIN"
'                sql = sql & "                      dbo.TblSettsReqLimK ON dbo.TblSettsReqLimKDet.SetReqLID = dbo.TblSettsReqLimK.ID"
'                sql = sql & " WHERE   TblSettsReqLimK.SelectType=2 and (TblSettsReqLimK.ItemStust=1 or TblSettsReqLimK.ItemStust=2 ) and   (dbo.TblSettsReqLimKDet.Typ = 0 AND (dbo.TblSettsReqLimKDet.ItemID = " & ItemID & ")) "
'                sql = sql + " AND (dbo.TblSettsReqLimKDet.StoreID = " & StoreID & ")  "
'                sql = sql + "  and      (" & SQLDate(RecDate, True) & " BETWEEN dbo.TblSettsReqLimK.RecordDate AND dbo.TblSettsReqLimK.RecordDateTo) "
'                sql = sql & "  AND (dbo.TblSettsReqLimK.TransType = " & Transe & ")"
'                rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'                If rs2.RecordCount > 0 Then
'                    PercentgValueAdded = IIf(IsNull(rs2("PercentD").value), 0, rs2("PercentD").value)
'                Else
'                    PercentgValueAdded = 0
'                End If
'            End If
'        End If
'    End If
'End Function
'Public Function GetServerdate(ServerDate As Date, ServerTime As Date)
'    Dim StrTemp As String
'    Dim RsTemp  As ADODB.Recordset
'
'    StrTemp = "select Getdate() as ServerDate , RIGHT(CONVERT(VARCHAR, GETDATE(), 100),7) as ServerTime"
'    Set RsTemp = New ADODB.Recordset
'    RsTemp.Open StrTemp, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (RsTemp.BOF Or RsTemp.EOF) Then
'        If Not IsNull(RsTemp("ServerDate").value) Then
'            ServerDate = Format(RsTemp("ServerDate").value, "yyyy/M/d")
'            ServerTime = RsTemp("ServerTime").value
'
'        End If
'
'    End If
'
'    RsTemp.Close
'    Set RsTemp = Nothing
'
'End Function
'Public Sub GetID_CodeSqureProject(Optional ByRef SquareCode As String = "", _
'                                  Optional ByRef ID As Double = 0)
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    sql = " SELECT     ID, SquareCode"
'    sql = sql & " From dbo.TblProjecInvestment"
'    If SquareCode <> "" Then
'        sql = sql & " WHERE     (SquareCode = N'" & SquareCode & "')"
'        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Rs3.RecordCount > 0 Then
'            ID = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
'        Else
'            ID = 0
'        End If
'    End If
'    If ID <> 0 Then
'        sql = sql & " WHERE     (ID = " & ID & ")"
'        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Rs3.RecordCount > 0 Then
'            SquareCode = IIf(IsNull(Rs3("SquareCode").value), "", Rs3("SquareCode").value)
'        Else
'            SquareCode = ""
'        End If
'    End If
'End Sub
'Public Function CheckPayment(Optional NoteID As Double) As Boolean
'    Dim StrSQL As String
'    Dim Rs3    As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    StrSQL = "select * From Notes where NoteType=4    "
'    StrSQL = StrSQL & "and   (NOT (Status IS NULL)) and NoteID=" & NoteID & ""
'    Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        CheckPayment = True
'    Else
'        CheckPayment = False
'    End If
'End Function
'
'Public Function CheckAprroveScreen(Optional ScreenName As String) As Boolean
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     ScreenName"
'    sql = sql & " From dbo.TblApprovalDef"
'    sql = sql & " WHERE     (ScreenName = N'" & ScreenName & "')"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        CheckAprroveScreen = True
'    Else
'        CheckAprroveScreen = False
'    End If
'End Function
'Public Sub UpdateItemsDefaultUnit()
'    Dim sql    As String
'    Dim ItemID As Double
'    Dim i      As Integer
'    Dim Rs3    As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    sql = "Select ItemID from TblItems "
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        Rs3.MoveFirst
'        For i = 1 To Rs3.RecordCount
'            ItemID = IIf(IsNull(Rs3("ItemID").value), 0, Rs3("ItemID").value)
'            Cn.Execute "update TblItemsUnits  set DefaultUnit=1 where ItemID=" & ItemID & " and UnitFactor= " & GetMaxUnitFactor(ItemID) & ""
'            Rs3.MoveNext
'        Next i
'    End If
'End Sub
'Public Function GetMaxUnitFactor(Optional ItemID As Double) As Double
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    sql = " SELECT     MAX(UnitFactor) AS MaxUnitFactor"
'    sql = sql & " From dbo.TblItemsUnits"
'    'Sql = Sql & " GROUP BY ItemID"
'    sql = sql & " where      (ItemID = " & ItemID & ")"
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        GetMaxUnitFactor = IIf(IsNull(Rs3("MaxUnitFactor").value), 0, Rs3("MaxUnitFactor").value)
'    Else
'        GetMaxUnitFactor = 0
'    End If
'End Function
'
'Public Function Calcul30orRminder(Optional ID As Integer) As Boolean
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    sql = " SELECT     id, culc30orRminder"
'    sql = sql & " From dbo.MOFRAD"
'    sql = sql & " Where (culc30orRminder = 1) and id=" & ID & ""
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        Calcul30orRminder = True
'    Else
'        Calcul30orRminder = False
'    End If
'End Function
'Public Function CheckSettingsVacType() As Boolean
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    sql = "Select Typ from TblVacationSettings where Typ=1 "
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        CheckSettingsVacType = True
'    Else
'        CheckSettingsVacType = False
'    End If
'End Function
'Public Function GetSettingsVacPeriod() As Double
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    sql = "Select NoMonth from TblVacationSettings "
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        GetSettingsVacPeriod = IIf(IsNull(Rs3("NoMonth").value), 0, Rs3("NoMonth").value)
'    Else
'        GetSettingsVacPeriod = 0
'    End If
'End Function
'
'Public Function GetSettingsVacDate(Optional RecDate As Date, _
'                                   Optional ByRef ID As Double) As Boolean
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'
'    sql = "select * from TblVacationSettingsDet  "
'    Set Rs3 = New ADODB.Recordset
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        Set Rs3 = New ADODB.Recordset
'        sql = "Select * from TblVacationSettingsDet "
'        sql = sql & " where FrmDate<= " & SQLDate(RecDate, True) & " "
'        sql = sql & " and  ToDate >= " & SQLDate(RecDate, True) & " "
'        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Rs3.RecordCount > 0 Then
'            GetSettingsVacDate = True
'            ID = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
'        Else
'            ID = 0
'            GetSettingsVacDate = False
'        End If
'    Else
'        GetSettingsVacDate = True
'        ID = 1
'    End If
'End Function
'Public Function GetSettingsVacDateAllow(Optional RecDate As Date, _
'                                        Optional ID As Double) As Boolean
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'
'    sql = "select * from TblVacationSettingsDet  "
'    Set Rs3 = New ADODB.Recordset
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        Set Rs3 = New ADODB.Recordset
'        sql = "Select * from TblVacationSettingsDet "
'        sql = sql & " where AlowDate >= " & SQLDate(RecDate, True) & " and ID=" & ID & " "
'        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Rs3.RecordCount > 0 Then
'            GetSettingsVacDateAllow = True
'        Else
'            GetSettingsVacDateAllow = False
'        End If
'    Else
'        GetSettingsVacDateAllow = True
'    End If
'End Function
'
'Public Function URLEncode2(ByVal str As String) As String
'    Dim intLen  As Integer
'    Dim X       As Integer
'    Dim curChar As Long
'    Dim newStr  As String
'
'    intLen = Len(str)
'    newStr = ""
'
'    For X = 1 To intLen
'        curChar = Asc(mId$(str, X, 1))
'
'        If (curChar < 48 Or curChar > 57) And (curChar < 65 Or curChar > 90) And (curChar < 97 Or curChar > 122) Then
'            newStr = newStr & "%" & Hex(curChar)
'        Else
'            newStr = newStr & Chr(curChar)
'        End If
'
'    Next X
'
'    URLEncode2 = newStr
'End Function
'
'Public Function URLEncode(ByVal URL As String, _
'                          Optional ByVal SpacePlus As Boolean = True) As String
'
'    Dim cchEscaped As Long
'    Dim HRESULT    As Long
'
'    If Len(URL) > INTERNET_MAX_URL_LENGTH Then
'        Err.Raise &H8004D700, "URLUtility.URLEncode", _
'           "URL parameter too long"
'    End If
'
'    cchEscaped = Len(URL) * 1.5
'    URLEncode = String$(cchEscaped, 0)
'    HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
'    If HRESULT = E_POINTER Then
'        URLEncode = String$(cchEscaped, 0)
'        HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
'    End If
'
'    If HRESULT <> S_OK Then
'        Err.Raise Err.LastDllError, "URLUtility.URLEncode", _
'           "System error"
'    End If
'
'    URLEncode = left$(URLEncode, cchEscaped)
'    If SpacePlus Then
'        URLEncode = Replace$(URLEncode, "+", "%2B")
'        URLEncode = Replace$(URLEncode, " ", "+")
'    End If
'End Function
'
'Public Function ConvertToUnicode(st As String) As String
'    Dim chrArray(0 To 149)     As String
'    Dim unicodeArray(0 To 149) As String
'
'10  chrArray(0) = "°"
'20  unicodeArray(0) = "060D"
'30  chrArray(1) = "∫"
'40  unicodeArray(1) = "061B"
'50  chrArray(2) = "ø"
'60  unicodeArray(2) = "061F"
'70  chrArray(3) = "¡"
'80  unicodeArray(3) = "0621"
'90  chrArray(4) = "¬"
'100     unicodeArray(4) = "0622"
'110     chrArray(5) = "√"
'120     unicodeArray(5) = "0623"
'130     chrArray(6) = "ƒ"
'140     unicodeArray(6) = "0624"
'150     chrArray(7) = "≈"
'160     unicodeArray(7) = "0625"
'170     chrArray(8) = "∆"
'180     unicodeArray(8) = "0626"
'190     chrArray(9) = "«"
'200     unicodeArray(9) = "0627"
'210     chrArray(10) = "»"
'220     unicodeArray(10) = "0628"
'230     chrArray(11) = "…"
'240     unicodeArray(11) = "0629"
'250     chrArray(12) = " "
'260     unicodeArray(12) = "062A"
'270     chrArray(13) = "À"
'280     unicodeArray(13) = "062B"
'290     chrArray(14) = "Ã"
'300     unicodeArray(14) = "062C"
'310     chrArray(15) = "Õ"
'320     unicodeArray(15) = "062D"
'330     chrArray(16) = "Œ"
'340     unicodeArray(16) = "062E"
'350     chrArray(17) = "œ"
'360     unicodeArray(17) = "062F"
'370     chrArray(18) = "–"
'380     unicodeArray(18) = "0630"
'390     chrArray(19) = "—"
'400     unicodeArray(19) = "0631"
'410     chrArray(20) = "“"
'420     unicodeArray(20) = "0632"
'430     chrArray(21) = "”"
'440     unicodeArray(21) = "0633"
'450     chrArray(22) = "‘"
'460     unicodeArray(22) = "0634"
'470     chrArray(23) = "’"
'480     unicodeArray(23) = "0635"
'490     chrArray(24) = "÷"
'500     unicodeArray(24) = "0636"
'510     chrArray(25) = "ÿ"
'520     unicodeArray(25) = "0637"
'530     chrArray(26) = "Ÿ"
'540     unicodeArray(26) = "0638"
'550     chrArray(27) = "⁄"
'560     unicodeArray(27) = "0639"
'570     chrArray(28) = "€"
'580     unicodeArray(28) = "063A"
'590     chrArray(29) = "›"
'600     unicodeArray(29) = "0641"
'610     chrArray(30) = "ﬁ"
'620     unicodeArray(30) = "0642"
'630     chrArray(31) = "ﬂ"
'640     unicodeArray(31) = "0643"
'650     chrArray(32) = "·"
'660     unicodeArray(32) = "0644"
'670     chrArray(33) = "„"
'680     unicodeArray(33) = "0645"
'690     chrArray(34) = "‰"
'700     unicodeArray(34) = "0646"
'710     chrArray(35) = "Â"
'720     unicodeArray(35) = "0647"
'730     chrArray(36) = "Ê"
'740     unicodeArray(36) = "0648"
'750     chrArray(37) = "Ï"
'760     unicodeArray(37) = "0649"
'770     chrArray(38) = "Ì"
'780     unicodeArray(38) = "064A"
'790     chrArray(39) = "‹"
'800     unicodeArray(39) = "0640"
'810     chrArray(40) = ""
'820     unicodeArray(40) = "064B"
'830     chrArray(41) = "Ò"
'840     unicodeArray(41) = "064C"
'850     chrArray(42) = "Ú"
'860     unicodeArray(42) = "064D"
'870     chrArray(43) = "Û"
'880     unicodeArray(43) = "064E"
'890     chrArray(44) = "ı"
'900     unicodeArray(44) = "064F"
'910     chrArray(45) = "ˆ"
'920     unicodeArray(45) = "0650"
'930     chrArray(46) = "¯"
'940     unicodeArray(46) = "0651"
'950     chrArray(47) = "˙"
'960     unicodeArray(47) = "0652"
'970     chrArray(48) = "!"
'980     unicodeArray(48) = "0021"
'990     chrArray(49) = """"
'1000    unicodeArray(49) = "0022"
'1010    chrArray(50) = "#"
'1020    unicodeArray(50) = "0023"
'1030    chrArray(51) = "$"
'1040    unicodeArray(51) = "0024"
'1050    chrArray(52) = "%"
'1060    unicodeArray(52) = "0025"
'1070    chrArray(53) = "&"
'1080    unicodeArray(53) = "0026"
'1090    chrArray(54) = "'"
'1100    unicodeArray(54) = "0027"
'1110    chrArray(55) = "("
'1120    unicodeArray(55) = "0028"
'1130    chrArray(56) = ")"
'1140    unicodeArray(56) = "0029"
'1150    chrArray(57) = "*"
'1160    unicodeArray(57) = "002A"
'1170    chrArray(58) = "+"
'1180    unicodeArray(58) = "002B"
'1190    chrArray(59) = ","
'1200    unicodeArray(59) = "002C"
'1210    chrArray(60) = "-"
'1220    unicodeArray(60) = "002D"
'1230    chrArray(61) = "."
'1240    unicodeArray(61) = "002E"
'1250    chrArray(62) = "/"
'1260    unicodeArray(62) = "002F"
'1270    chrArray(63) = "0"
'1280    unicodeArray(63) = "0030"
'1290    chrArray(64) = "1"
'1300    unicodeArray(64) = "0031"
'1310    chrArray(65) = "2"
'1320    unicodeArray(65) = "0032"
'1330    chrArray(66) = "3"
'1340    unicodeArray(66) = "0033"
'1350    chrArray(67) = "4"
'1360    unicodeArray(67) = "0034"
'1370    chrArray(68) = "5"
'1380    unicodeArray(68) = "0035"
'1390    chrArray(69) = "6"
'1400    unicodeArray(69) = "0036"
'1410    chrArray(70) = "7"
'1420    unicodeArray(70) = "0037"
'1430    chrArray(71) = "8"
'1440    unicodeArray(71) = "0038"
'1450    chrArray(72) = "9"
'1460    unicodeArray(72) = "0039"
'1470    chrArray(73) = ":"
'1480    unicodeArray(73) = "003A"
'1490    chrArray(74) = ""
'1500    unicodeArray(74) = "003B"
'1510    chrArray(75) = "<"
'1520    unicodeArray(75) = "003C"
'1530    chrArray(76) = "="
'1540    unicodeArray(76) = "003D"
'1550    chrArray(77) = ">"
'1560    unicodeArray(77) = "003E"
'1570    chrArray(78) = "?"
'1580    unicodeArray(78) = "003F"
'1590    chrArray(79) = "@"
'1600    unicodeArray(79) = "0040"
'1610    chrArray(80) = "A"
'1620    unicodeArray(80) = "0041"
'1630    chrArray(81) = "B"
'1640    unicodeArray(81) = "0042"
'1650    chrArray(82) = "C"
'1660    unicodeArray(82) = "0043"
'1670    chrArray(83) = "D"
'1680    unicodeArray(83) = "0044"
'1690    chrArray(84) = "E"
'1700    unicodeArray(84) = "0045"
'1710    chrArray(85) = "F"
'1720    unicodeArray(85) = "0046"
'1730    chrArray(86) = "G"
'1740    unicodeArray(86) = "0047"
'1750    chrArray(87) = "H"
'1760    unicodeArray(87) = "0048"
'1770    chrArray(88) = "I"
'1780    unicodeArray(88) = "0049"
'1790    chrArray(89) = "J"
'1800    unicodeArray(89) = "004A"
'1810    chrArray(90) = "K"
'1820    unicodeArray(90) = "004B"
'1830    chrArray(91) = "L"
'1840    unicodeArray(91) = "004C"
'1850    chrArray(92) = "M"
'1860    unicodeArray(92) = "004D"
'1870    chrArray(93) = "N"
'1880    unicodeArray(93) = "004E"
'1890    chrArray(94) = "O"
'1900    unicodeArray(94) = "004F"
'1910    chrArray(95) = "P"
'1920    unicodeArray(95) = "0050"
'1930    chrArray(96) = "Q"
'1940    unicodeArray(96) = "0051"
'1950    chrArray(97) = "R"
'1960    unicodeArray(97) = "0052"
'1970    chrArray(98) = "S"
'1980    unicodeArray(98) = "0053"
'1990    chrArray(99) = "T"
'2000    unicodeArray(99) = "0054"
'2010    chrArray(100) = "U"
'2020    unicodeArray(100) = "0055"
'2030    chrArray(101) = "V"
'2040    unicodeArray(101) = "0056"
'2050    chrArray(102) = "W"
'2060    unicodeArray(102) = "0057"
'2070    chrArray(103) = "X"
'2080    unicodeArray(103) = "0058"
'2090    chrArray(104) = "Y"
'2100    unicodeArray(104) = "0059"
'2110    chrArray(105) = "Z"
'2120    unicodeArray(105) = "005A"
'2130    chrArray(106) = "[" '"("
'2140    unicodeArray(106) = "005B"
'2150    chrArray(107) = Trim("\ ")
'2160    unicodeArray(107) = "005C"
'2170    chrArray(108) = "]" '")"
'2180    unicodeArray(108) = "005D"
'2190    chrArray(109) = "^"
'2200    unicodeArray(109) = "005E"
'2210    chrArray(110) = "_"
'2220    unicodeArray(110) = "005F"
'2230    chrArray(111) = "`"
'2240    unicodeArray(111) = "0060"
'2250    chrArray(112) = "a"
'2260    unicodeArray(112) = "0061"
'2270    chrArray(113) = "b"
'2280    unicodeArray(113) = "0062"
'2290    chrArray(114) = "c"
'2300    unicodeArray(114) = "0063"
'2310    chrArray(115) = "d"
'2320    unicodeArray(115) = "0064"
'2330    chrArray(116) = "e"
'2340    unicodeArray(116) = "0065"
'2350    chrArray(117) = "f"
'2360    unicodeArray(117) = "0066"
'2370    chrArray(118) = "g"
'2380    unicodeArray(118) = "0067"
'2390    chrArray(119) = "h"
'2400    unicodeArray(119) = "0068"
'2410    chrArray(120) = "i"
'2420    unicodeArray(120) = "0069"
'2430    chrArray(121) = "j"
'2440    unicodeArray(121) = "006A"
'2450    chrArray(122) = "k"
'2460    unicodeArray(122) = "006B"
'2470    chrArray(123) = "l"
'2480    unicodeArray(123) = "006C"
'2490    chrArray(124) = "m"
'2500    unicodeArray(124) = "006D"
'2510    chrArray(125) = "n"
'2520    unicodeArray(125) = "006E"
'2530    chrArray(126) = "o"
'2540    unicodeArray(126) = "006F"
'2550    chrArray(127) = "p"
'2560    unicodeArray(127) = "0070"
'2570    chrArray(128) = "q"
'2580    unicodeArray(128) = "0071"
'2590    chrArray(129) = "r"
'2600    unicodeArray(129) = "0072"
'2610    chrArray(130) = "s"
'2620    unicodeArray(130) = "0073"
'2630    chrArray(131) = "t"
'2640    unicodeArray(131) = "0074"
'2650    chrArray(132) = "u"
'2660    unicodeArray(132) = "0075"
'2670    chrArray(133) = "v"
'2680    unicodeArray(133) = "0076"
'2690    chrArray(134) = "w"
'2700    unicodeArray(134) = "0077"
'2710    chrArray(135) = "x"
'2720    unicodeArray(135) = "0078"
'2730    chrArray(136) = "y"
'2740    unicodeArray(136) = "0079"
'2750    chrArray(137) = "z"
'2760    unicodeArray(137) = "007A"
'2770    chrArray(138) = "{"
'2780    unicodeArray(138) = "007B"
'2790    chrArray(139) = "|"
'2800    unicodeArray(139) = "007C"
'2810    chrArray(140) = "}"
'2820    unicodeArray(140) = "007D"
'2830    chrArray(141) = "~"
'2840    unicodeArray(141) = "007E"
'2850    chrArray(142) = "©"
'2860    unicodeArray(142) = "00A9"
'2870    chrArray(143) = "Æ"
'2880    unicodeArray(143) = "00AE"
'2890    chrArray(144) = "˜"
'2900    unicodeArray(144) = "00F7"
'2910    chrArray(145) = "◊"
'2920    unicodeArray(145) = "00F7"
'2930    chrArray(146) = "ß"
'2940    unicodeArray(146) = "00A7"
'2950    chrArray(147) = " "
'2960    unicodeArray(147) = "0020"
'2970    chrArray(148) = Chr$(13)
'2980    unicodeArray(148) = "000D"
'2990    chrArray(149) = "\r"
'3000    unicodeArray(149) = "000A"
'
'        Dim strResult As String, i As Integer, c As Integer
'3010    strResult = ""
'
'3020    For i = 1 To Len(st)
'3030        For c = 0 To 149
'
'3040            If (chrArray(c) = mId(st, i, 1)) Then
'3050                strResult = strResult & unicodeArray(c)
'3060            End If
'
'3070        Next c
'3080    Next i
'
'3090    ConvertToUnicode = strResult
'
'End Function
'
'Function SendEmailForCustomer(CusID As Integer, _
'                              subject1 As String, _
'                              Msg As String, _
'                              ByRef msgstatus As String)
'    'Dim subject As String
'    'Dim msg As String
'    Dim CompanyName As String
'    Dim cOptions    As ClsCompanyInfo
'    'Dim msg As String
'    Set cOptions = New ClsCompanyInfo
'    Dim Email        As String
'    Dim CustomerName As String
'    CustomerName = ""
'    Email = GetCustomerEmail(CusID, CustomerName)
'    CompanyName = cOptions.ArabCompanyName & Chr(13) & CurrentBranchName
'
'    subject = " «·”«œ… / " & CustomerName & " " & subject1
'
'    Dim RetVal As String
'    RetVal = SendMail(Email, _
'       Trim$(subject), _
'       "»ﬁŒ…", _
'       Msg, _
'       "txtServer", _
'       25, _
'       "txtUsername", _
'       "txtPassword", _
'       "", _
'       False, True)
'    msgstatus = IIf(RetVal = "ok", " „ «—”«· «·“Ì«—…", RetVal)
'
'End Function
'
'Public Function get_TblPaymentTypet(ID As Long, _
'   filed As String) As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim sql As String
'    sql = "Select  " & filed & " from TblPaymentType where PaymentID=" & ID
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount = 0 Then
'        get_TblPaymentTypet = ""
'        Exit Function
'    End If
'    If IsNull(Rs3(filed).value) Then
'        get_TblPaymentTypet = ""
'        Exit Function
'    End If
'    If Not IsNull(Rs3(filed).value) Then
'        get_TblPaymentTypet = Rs3(filed).value
'        Exit Function
'    End If
'    Rs3.Close
'
'End Function
'Public Function ChekPayedSalary(Optional YearID As Integer, _
'                                Optional MonthID As Integer, _
'                                Optional BranchID As Integer) As Boolean
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    sql = " SELECT     BranchId, MONTH(RecordDate) AS MonthID, YEAR(RecordDate) AS YearID"
'    sql = sql & " From dbo.emp_salary"
'    sql = sql & " Where (year(RecordDate) = " & YearID & ") And (Month(RecordDate) = " & MonthID & ")"
'    sql = sql & " AND BranchID=" & BranchID
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        ChekPayedSalary = True
'    Else
'        ChekPayedSalary = False
'    End If
'End Function
'Public Function GetReustValue(Optional StoreID As Long, Optional ItemID As Long) As Double
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim AllowQty   As Double
'    Dim UnitFacort As Double
'    Dim Qty        As Double
'    sql = " SELECT     Xb.Qty, Xb.StoreID, Xb.ItemID, Xb.UnitFactor, Xb.UnitID, BX.QNty"
'    sql = sql & " FROM         (SELECT     dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.UnitFactor,"
'    sql = sql & "                                              dbo.TblSettsRequestLimitDet.unitid"
'    sql = sql & "                        FROM         dbo.TblSettsRequestLimitDet INNER JOIN"
'    sql = sql & "                                              dbo.Transaction_Details ON dbo.TblSettsRequestLimitDet.ItemID = dbo.Transaction_Details.Item_ID"
'    sql = sql & "                        Where (dbo.TblSettsRequestLimitDet.typ = 0)"
'    sql = sql & "                        GROUP BY dbo.TblSettsRequestLimitDet.Qty, dbo.TblSettsRequestLimitDet.StoreID, dbo.TblSettsRequestLimitDet.ItemID, dbo.TblSettsRequestLimitDet.UnitFactor,"
'    sql = sql & "                                              dbo.TblSettsRequestLimitDet.UnitID) Xb INNER JOIN"
'    sql = sql & "                          (SELECT     SUM(dbo.TransactionTypes.StockEffect * dbo.Transaction_Details.Quantity) AS QNty, dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID"
'    sql = sql & "                             FROM         dbo.Transactions INNER JOIN"
'    sql = sql & "                                                   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
'    sql = sql & "                                                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
'    sql = sql & "                             GROUP BY Item_ID, StoreID) BX ON BX.Item_ID = Xb.ItemID  AND BX.StoreID = Xb.StoreID"
'    sql = sql & "  Where Xb.StoreID = " & StoreID & " And Xb.ItemID = " & ItemID & ""
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        UnitFacort = IIf(IsNull(Rs3("UnitFactor").value), 0, Rs3("UnitFactor").value)
'        Qty = IIf(IsNull(Rs3("Qty").value), 0, Rs3("Qty").value)
'        AllowQty = IIf(IsNull(Rs3("QNty").value), 0, Rs3("QNty").value)
'        If UnitFacort > 0 Then
'            AllowQty = AllowQty / UnitFacort
'        End If
'        GetReustValue = AllowQty - Qty
'    Else
'        GetReustValue = 0
'    End If
'End Function
'
'Public Function DescUnitFact(Optional ItemID As Long, Optional UntID As Long) As Double
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = "SELECT     TOP 100 PERCENT UnitFactor, UnitID"
'    sql = sql & " From dbo.TblItemsUnits"
'    sql = sql & " Where (ItemID = " & ItemID & ") And (unitid = " & UntID & ")"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        DescUnitFact = IIf(IsNull(rs2("UnitFactor").value), 1, rs2("UnitFactor").value)
'    Else
'        DescUnitFact = 0
'    End If
'End Function
'Public Function DescUnit(Optional ItemID As Long, Optional UnitID As Long) As String
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = " SELECT     TOP 100 PERCENT dbo.TblItemsUnits.FactorBySmallUnit, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
'    sql = sql & " FROM         dbo.TblItemsUnits INNER JOIN"
'    sql = sql & "                      dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
'    sql = sql & "  Where (dbo.TblItemsUnits.ItemID = " & ItemID & ") And (dbo.TblItemsUnits.DefaultUnit = 1)"
'    sql = sql & " ORDER BY dbo.TblItemsUnits.ItemID"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        DescUnit = DescUnitFact(ItemID, UnitID)
'        DescUnit = DescUnit & " "
'        If SystemOptions.UserInterface = ArabicInterface Then
'            DescUnit = DescUnit & IIf(IsNull(rs2("UnitName").value), "", rs2("UnitName").value)
'        Else
'            DescUnit = DescUnit & IIf(IsNull(rs2("UnitNamee").value), "", rs2("UnitNamee").value)
'        End If
'    Else
'        DescUnit = ""
'    End If
'End Function
'Public Sub GetCodeIDProject(Optional ByRef ID As Double = 0, _
'                            Optional ByRef Fullcode As String, _
'                            Optional getMaterial_account As Integer = 0, _
'                            Optional ByRef Material_account As String)
'    Dim sql As String
'    Dim Rs4 As ADODB.Recordset
'    Set Rs4 = New ADODB.Recordset
'    sql = "Select id ,Fullcode ,Material_account from projects "
'    If ID <> 0 Then
'        sql = sql & " where id=" & ID & ""
'        Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Rs4.RecordCount > 0 Then
'            Fullcode = IIf(IsNull(Rs4("Fullcode").value), "", Rs4("Fullcode").value)
'        Else
'            Fullcode = ""
'        End If
'    Else
'
'        sql = sql & " where Fullcode=N'" & Fullcode & "'"
'
'        Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Rs4.RecordCount > 0 Then
'            ID = IIf(IsNull(Rs4("id").value), "", Rs4("id").value)
'            If getMaterial_account <> 0 Then
'                Material_account = IIf(IsNull(Rs4("Material_account").value), "", Rs4("Material_account").value)
'            End If
'        Else
'            ID = 0
'        End If
'    End If
'End Sub
'Public Function ChicIsLotNo(Optional ItemID As Long) As Boolean
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    sql = " SELECT     ItemID, ChkLot"
'    sql = sql & " From dbo.TblItems"
'    sql = sql & " Where (ItemID = " & ItemID & ") And (ChkLot = 1)"
'    Rs3.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        ChicIsLotNo = True
'    Else
'        ChicIsLotNo = False
'    End If
'End Function
Public Function UpdateTransactionsCost(Transaction_IDs As String)
    If SystemOptions.AllowCostnNewShape = False Then
        Exit Function
    End If
    Dim sql As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim OldQty           As Double
    Dim OldCost          As Double
    Dim NewQty           As Double
    Dim NewCost          As Double
    Dim StockEffect      As Integer

    Dim StoreID          As Double
    Dim Item_ID          As Double
    Dim Transaction_Date As Date
    Dim Transaction_ID   As Double
Dim LngUnitID As Long

    'sql = "Select * from TbLSheft where TypHour=1 "
    sql = "SELECT   dbo.Transaction_Details.QtyBySmalltUnit  , dbo.Transactions.Transaction_Type, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, "
    sql = sql & "                       dbo.Transaction_Details.OldCost, dbo.Transaction_Details.OldQty, dbo.Transaction_Details.NewCost, dbo.Transaction_Details.NewQty,"
    sql = sql & "  dbo.Transactions.Transaction_ID , dbo.TransactionTypes.StockEffect , dbo.Transactions.StoreID, dbo.Transactions.Transaction_Date, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price"

    sql = sql & "  FROM         dbo.Transactions INNER JOIN"
    sql = sql & "                       dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    sql = sql & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    sql = sql & "  Where (dbo.Transactions.Transaction_ID in ( " & Transaction_IDs & " ))"

    'WaelCost
    sql = sql & "  ORDER BY dbo.Transactions.Transaction_Date DESC ,dbo.TransactionTypes.StockEffect, dbo.Transactions.Transaction_ID DESC, dbo.Transaction_Details.ID DESC"

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount > 0 Then
        For i = 1 To Rs3.RecordCount
            StockEffect = IIf(IsNull(Rs3("StockEffect").value), 0, Rs3("StockEffect").value)
            StoreID = IIf(IsNull(Rs3("StoreID").value), 0, Rs3("StoreID").value)
            Item_ID = IIf(IsNull(Rs3("Item_ID").value), 0, Rs3("Item_ID").value)
            Transaction_Date = IIf(IsNull(Rs3("Transaction_Date").value), Date, Rs3("Transaction_Date").value)
            Transaction_ID = IIf(IsNull(Rs3("Transaction_ID").value), 0, Rs3("Transaction_ID").value)

            getItemCostData Transaction_Date, Item_ID, StoreID, Transaction_ID, OldQty, OldCost, NewQty, NewCost
            Dim QtyBySmalltUnit As Double
            If StockEffect = -1 Then
                QtyBySmalltUnit = IIf(IsNull(Rs3("QtyBySmalltUnit").value), 1, Rs3("QtyBySmalltUnit").value)

                Rs3("OldQty").value = NewQty
                Rs3("OldCost").value = NewCost

                Rs3("NewQty").value = Rs3("OldQty").value - Rs3("Quantity").value
                Rs3("NewCost").value = Rs3("OldCost").value ' ((Rs3("OldQty").value * Rs3("OldCost").value) + (Rs3("Quantity").value * Rs3("Price").value)) / (Rs3("Quantity").value + Rs3("OldQty").value)

                Rs3.update

            ElseIf StockEffect = 1 Then 'input
                QtyBySmalltUnit = IIf(IsNull(Rs3("QtyBySmalltUnit").value), 1, Rs3("QtyBySmalltUnit").value)

                Rs3("OldQty").value = NewQty
                Rs3("OldCost").value = NewCost

                Rs3("NewQty").value = Rs3("Quantity").value + Rs3("OldQty").value
                If val(Rs3("Quantity").value + Rs3("OldQty").value) <> 0 Then
                    Rs3("NewCost").value = ((Round(Rs3("OldQty").value, 4) * Round(Rs3("OldCost").value, 4)) + (Round(Rs3("Quantity").value, 4) * Round(Rs3("Price").value, 4))) / (Round(Rs3("Quantity").value, 4) + Round(Rs3("OldQty").value, 4))           'IIf(Rs3("Quantity").value + Rs3("OldQty").value <> 0, Rs3("Quantity").value + Rs3("OldQty").value, 0)
                Else
                    Rs3("NewCost").value = 0
                End If
                Rs3.update

            Else

            End If
            Rs3.MoveNext
        Next i
    Else

    End If

End Function
'Public Function updateCopyNo(tablename As String, _
'                             Filedname As String, _
'                             transactionfiledname As String, _
'                             Transaction_ID As Double)
'    '        updateCopyNo "Transactions", "CopyNO", "Transaction_ID", val(Me.XPTxtBillID.Text)
'
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'
'    sql = ""
'    sql = "SELECT     " & Filedname & ""
'    sql = sql & " From dbo." & tablename & ""
'    sql = sql & " Where (" & transactionfiledname & " = " & Transaction_ID & ")"
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        lastCopyno = IIf(IsNull(Rs3(Filedname).value), 0, Rs3(Filedname).value)
'
'        Cn.Execute "update  " & tablename & " set " & Filedname & "=" & Filedname & "+1 where Transaction_ID=" & Transaction_ID
'    End If
'End Function
'
Public Function getItemCostData(Transaction_Date As Date, _
                                Item_ID As Double, _
                                StoreID As Double, _
                                Optional Transaction_ID As Double = -1, _
                                Optional ByRef OldQty As Double, _
                                Optional ByRef OldCost As Double, _
                                Optional ByRef NewQty As Double, _
                                Optional ByRef NewCost As Double, _
                                Optional ByVal IsFromGetCostPrice As Boolean = False, _
                                Optional ByVal LngUnitID As Long = 0, _
                                Optional ByRef UnitFactor As Double = 1, _
                                Optional ByRef SecOrder As Integer = 1)
    'Transaction_ID = -1
    Dim sql             As String
    Dim QtyBySmalltUnit As Double
 '   Dim UnitFactor As Double
 '   Dim SecOrder As Integer
    Dim Rs3             As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rsItemCost As New ADODB.Recordset
    Dim s          As String
    
    'sql = "Select * from TbLSheft where TypHour=1 "
    sql = "SELECT     TOP 1 dbo.Transaction_Details.QtyBySmalltUnit, dbo.Transactions.Transaction_Type,TblItemsUnits.UnitFactor,TblItemsUnits.SecOrder, dbo.TransactionTypes.StockEffect, dbo.Transactions.Transaction_Date, dbo.Transaction_Details.Item_ID, "
    sql = sql & " dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transaction_Details.OldQty,Transaction_Details.UnitID, dbo.Transaction_Details.OldCost, dbo.Transaction_Details.NewQty,"
    sql = sql & "                       dbo.Transaction_Details.NewCost , dbo.Transactions.Transaction_ID"
    sql = sql & " FROM         dbo.Transactions INNER JOIN"
    sql = sql & "                       dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    sql = sql & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    sql = sql & "                       Inner join TblItemsUnits On TblItemsUnits.ItemId = Transaction_Details.Item_ID and TblItemsUnits.UnitId =Transaction_Details.UnitID "
    sql = sql & "  WHERE     (dbo.TransactionTypes.StockEffect <>0) "
    sql = sql & "  AND (dbo.Transactions.Transaction_Date <= " & SQLDate(Transaction_Date, True) & ")"
    sql = sql & "  AND               (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
    If LngUnitID <> 0 Then
    '    sql = sql & "        and (Transaction_Details.UnitID = " & LngUnitID & ")"
    End If

    If StoreID = 0 Then
    Else
        'sql = sql & " AND (dbo.Transactions.StoreID = " & StoreId & ") "
    End If
    If Transaction_ID > 0 Then
        sql = sql & "  AND (dbo.Transactions.Transaction_ID <>  " & Transaction_ID & ")"
    End If

    sql = sql & "  ORDER BY dbo.Transactions.Transaction_Date DESC , dbo.Transactions.Transaction_ID DESC, dbo.Transaction_Details.ID DESC"

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.RecordCount > 0 Then
        QtyBySmalltUnit = IIf(IsNull(Rs3("QtyBySmalltUnit").value), 1, Rs3("QtyBySmalltUnit").value)
        OldQty = IIf(IsNull(Rs3("oldQty").value), 0, Rs3("oldQty").value)
        OldCost = IIf(IsNull(Rs3("oldCost").value), 0, Rs3("oldCost").value)
        NewQty = IIf(IsNull(Rs3("NewQty").value), 0, Rs3("NewQty").value)
        NewCost = IIf(IsNull(Rs3("NewCost").value), 0, Rs3("NewCost").value)
        
        UnitFactor = IIf(IsNull(Rs3("UnitFactor").value), 0, Rs3("UnitFactor").value)
        SecOrder = IIf(IsNull(Rs3("SecOrder").value), 0, Rs3("SecOrder").value)
        
          
    
        'Wael Cost Salim
        'Õ«·… «· ﬂ·›… »«·”«·» ‰ ÌÃ… «·”Õ» ⁄·Ï «·„ﬂ‘Ê› ‰√ Ï »«Œ— ”⁄— ‘—«¡
        If NewCost <= 0 Then
            If NewCost = 0 And Transaction_ID = -950 Then Exit Function
            NewCost = getcostbuylastinvoice(CDbl(Item_ID), CDate(Transaction_Date), LngUnitID, UnitFactor, SecOrder)
            If NewCost <= 0 Then
                If Not IsFromGetCostPrice Then
                    NewCost = ModItemCostPrice.GetCostItemPrice(CLng(Item_ID), 0, "", , SystemOptions.SysMainStockCostMethod, , , CDate(Transaction_Date), val(Transaction_ID), val(Rs3("UnitID").value & ""), val(StoreID))
                End If
            End If
            If NewCost <= 0 Then
                If LngUnitID <> 0 Then
                    s = " Select UnitPurPrice from TblItemsUnits where ItemID = " & val(Item_ID) & "  and UnitId = " & val(LngUnitID)
                Else
                 
                End If
                   s = " Select UnitPurPrice from TblItemsUnits where ItemID = " & val(Item_ID) & "  and UnitId = " & val(Rs3("UnitID").value & "")
                Set rsItemCost = New ADODB.Recordset
                rsItemCost.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsItemCost.EOF Then
                    NewCost = val(rsItemCost!UnitPurPrice & "")
                End If
            End If
            sql = " Update Transaction_Details Set IsLastPurPrice = 1 where Transaction_ID =  " & Transaction_ID & " and Item_ID = " & Item_ID
            Cn.Execute sql
        End If
    Else

        If NewCost <= 0 Then
            
            NewCost = getcostbuylastinvoice(CDbl(Item_ID), CDate(Transaction_Date), LngUnitID, UnitFactor, SecOrder)
            If NewCost <= 0 Then
                If Not IsFromGetCostPrice Then
                    NewCost = ModItemCostPrice.GetCostItemPrice(CLng(Item_ID), 0, "", , SystemOptions.SysMainStockCostMethod, , , CDate(Transaction_Date), val(Transaction_ID), val(LngUnitID), val(StoreID))
                End If
            End If
            If NewCost <= 0 Then
                s = " Select UnitPurPrice from TblItemsUnits where ItemID = " & val(Item_ID) & "  and UnitId = " & val(LngUnitID)
                Set rsItemCost = New ADODB.Recordset
                rsItemCost.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsItemCost.EOF Then
                    NewCost = val(rsItemCost!UnitPurPrice & "")
                End If
            End If
            sql = " Update Transaction_Details Set IsLastPurPrice = 1 where Transaction_ID =  " & Transaction_ID & " and Item_ID = " & Item_ID
            Cn.Execute sql
        End If
        'OldQty = 0
        'OldCost = 0
        'NewQty = 0
        'NewCost = 0
    End If

End Function
'Public Function NoHourInShift(Optional ByRef NoHour As Double, _
'                              Optional EmpID As Double) As Boolean
'
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    'sql = "Select * from TbLSheft where TypHour=1 "
'    sql = "SELECT     dbo.TbLSheft.NoHourManaula"
'    sql = sql & "  FROM         dbo.TbLSheft INNER JOIN"
'    sql = sql & "                        dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
'    sql = sql & "  Where (dbo.TbLSheft.TypHour = 1) And (dbo.TblShiftWorker.EmpID = " & EmpID & ")"
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        NoHourInShift = True
'        NoHour = IIf(IsNull(Rs3("NoHourManaula").value), 0, Rs3("NoHourManaula").value)
'    Else
'        NoHour = 0
'        NoHourInShift = False
'    End If
'End Function
''***************************
'Public Function GetOrdersData(Optional ID As Double, _
'                              Optional ByRef EmpName As String, _
'                              Optional ByRef EmpMbile As String, _
'                              Optional ByRef OrdeNo As String, _
'                              Optional NoteSerial1 As Double, _
'                              Optional SeasonsID As Double) As Double
'    Dim Rs4 As ADODB.Recordset
'    Dim sql As String
'    Set Rs4 = New ADODB.Recordset
'    sql = "SELECT    * from tblbookingrequest where ID=" & ID & " and StusID=1"
'    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs4.RecordCount > 0 Then
'        EmpName = IIf(IsNull(Rs4("EmpName").value), "", Rs4("EmpName").value)
'        EmpMbile = IIf(IsNull(Rs4("EmpMbile").value), "", Rs4("EmpMbile").value)
'        SeasonsID = IIf(IsNull(Rs4("SeasonsID").value), 0, Rs4("SeasonsID").value)
'        NoteSerial1 = IIf(IsNull(Rs4("NoteSerial1").value), 0, Rs4("NoteSerial1").value)
'    Else
'        NoteSerial1 = 0
'        SeasonsID = 0
'        EmpName = ""
'        EmpMbile = ""
'    End If
'End Function
'
'Public Function CheckDelLocations(CustomerlocationID As Long) As Boolean
'    Dim rs     As ADODB.Recordset
'    Dim StrSQL As String
'    StrSQL = "Select * From Transactions  Where CustomerlocationID=" & CustomerlocationID
'    Set rs = New ADODB.Recordset
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (rs.BOF Or rs.EOF) Then
'        CheckDelLocations = False
'    Else
'        CheckDelLocations = True
'    End If
'
'    rs.Close
'    Set rs = Nothing
'End Function
'
'Public Function GeTuserFullCode(UserID As Double) As String
'    Dim rs  As ADODB.Recordset
'    Dim Rs1 As ADODB.Recordset
'
'    If UserID = 1 Then
'        GeTuserFullCode = "0000"
'        Exit Function
'    End If
'    Dim sql As String
'    Dim str As String
'    Set Rs1 = New ADODB.Recordset
'    sql = " SELECT     dbo.TblUsers.UserID, dbo.TblEmployee.Fullcode"
'    sql = sql & "  FROM         dbo.TblUsers INNER JOIN"
'    sql = sql & "  dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
'    sql = sql & "   WHERE     (dbo.TblUsers.UserID = " & UserID & ")"
'    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If Rs1.RecordCount > 0 Then
'        GeTuserFullCode = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
'    Else
'        GeTuserFullCode = 0
'    End If
'    Rs1.Close
'End Function
'
'Public Function TimeStamp(date1 As Date) As String
'    Dim StartDate     As String
'    Dim EndTime       As String
'    Dim startTime     As String
'    Dim EndDate       As String
'    Dim dblStart      As Double
'    Dim dblEnd        As Double
'    Dim DateTimeStart As Date
'    Dim DateTimeEnd   As Date
'    Dim TotalHrs      As String
'    StartDate = "1/1/1970"
'    startTime = "00:00:00"
'    EndDate = CStr(date1)
'    EndTime = CStr(Time)
'    DateTimeStart = FormatDateTime(StartDate & " " & startTime)
'    DateTimeEnd = FormatDateTime(EndDate & " " & EndTime)
'    TimeStamp = DateDiff("s", DateTimeStart, DateTimeEnd, vbUseSystemDayOfWeek, _
'       vbUseSystem)
'End Function
'
'Public Function GetItemsData(ByRef ItemName As String, _
'                             Optional ByRef ItemID As Double, _
'                             Optional ByRef Fullcode As String, _
'                             Optional ByRef PartNo As String) As Boolean
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    If ItemName <> "" Then
'        sql = "Select * from  TblItems where ItemName ='" & ItemName & "'"
'    Else
'        sql = "Select * from  TblItems where code ='" & Fullcode & "'"
'    End If
'
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        ItemID = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
'        itemcode = IIf(IsNull(rs2("FullCode").value), "", rs2("FullCode").value)
'        PartNo = IIf(IsNull(rs2("PartNo").value), "", rs2("PartNo").value)
'        ItemName = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
'    Else
'        ItemID = 0
'        itemcode = ""
'        PartNo = ""
'        Fullcode = ""
'    End If
'End Function
'
'Public Function GetStoreData(ByRef StoreName As String, _
'                             Optional ByRef StoreID As Double, _
'                             Optional ByRef BranchID As Double, _
'                             Optional Fullcode As String) As Boolean
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    If Fullcode = "" Then
'        sql = "Select * from  TblStore where StoreName ='" & StoreName & "'"
'    Else
'        sql = "Select * from  TblStore where Code ='" & Fullcode & "'"
'    End If
'
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        StoreID = IIf(IsNull(rs2("storeid").value), 0, rs2("storeid").value)
'        BranchID = IIf(IsNull(rs2("BranchId").value), 0, rs2("BranchId").value)
'        StoreName = IIf(IsNull(rs2("storename").value), 0, rs2("storename").value)
'    Else
'        StoreName = ""
'        StoreID = 0
'        BranchID = 0
'    End If
'End Function
'
'Public Function CheCkTriningRequest() As Boolean
'    Dim sql As String
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    sql = "Select * from  TblTrainingRequest"
'    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rs2.RecordCount > 0 Then
'        CheCkTriningRequest = True
'    Else
'        CheCkTriningRequest = False
'    End If
'End Function
'Public Function GetBranchnmeFromnotes(NoteID As Double, _
'                                      Optional ByRef branch_id As Double, _
'                                      Optional ByRef branch_name As String, _
'                                      Optional ByRef branch_namee As String, _
'                                      Optional Vat As Double)
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = "SELECT   vat,  dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
'    sql = sql & "  FROM         dbo.Notes INNER JOIN"
'    sql = sql & "                        dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
'    sql = sql & "   Where (dbo.Notes.noteID = " & NoteID & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'
'        branch_id = IIf(IsNull(rs("branch_no").value), 0, rs("branch_no").value)
'        branch_name = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
'        branch_namee = IIf(IsNull(rs("branch_namee").value), 0, rs("branch_namee").value)
'        Vat = IIf(IsNull(rs("vat").value), 0, rs("vat").value)
'
'    Else
'        Vat = 0
'        branch_id = 0
'        branch_name = ""
'        branch_namee = ""
'
'    End If
'    rs.Close
'End Function
'
'Public Function GetMixIdFormCode(MixCode As String) As Double
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = "SELECT     ID From dbo.TblDefComItem WHERE     (MaxNo = '" & MixCode & "')"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'
'        GetMixIdFormCode = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
'
'    Else
'        GetMixIdFormCode = 0
'
'    End If
'    rs.Close
'End Function
'Function CheckRemainsetyforprojectopr(project_id1 As Double, _
'                                      pand_id As Double, _
'                                      oper_id As Double, _
'                                      Item_ID As Double, _
'                                      Transaction_ID As Double, _
'                                      TransQty As Double, _
'                                      txtmodflage As String, _
'                                      Optional ByRef OPRQTY As Double, _
'                                      Optional ByRef IssuedOprQty As Double) As Double
'
'    Dim sql  As String
'    Dim rs   As New ADODB.Recordset
'
'    Dim SQL1 As String
'    Dim Rs1  As New ADODB.Recordset
'
'    '„⁄—›… ﬂ„ÌÂ «·»‰œ
'    SQL1 = "  SELECT      sum( dbo.TblMatrials.COUNT) OPRQTY   FROM     terms_operations"
'    SQL1 = SQL1 & " inner join TblMatrials  on dbo.TblMatrials.Opr = dbo.terms_operations.id"
'    SQL1 = SQL1 & "  Where 1 = 1   "
'    SQL1 = SQL1 & "   and  (dbo.terms_operations.Project_ID =" & project_id1 & ")"
'    SQL1 = SQL1 & "   and  (dbo.terms_operations.projectdes_id=" & pand_id & ")"
'
'    SQL1 = SQL1 & "    and dbo.terms_operations.OPRIDD =" & oper_id & ""
'    SQL1 = SQL1 & "   and  dbo.TblMatrials.ItemID  =" & Item_ID & ""
'
'    Rs1.Open SQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If Rs1.RecordCount > 0 Then
'
'        OPRQTY = IIf(IsNull(Rs1("OPRQTY").value), 0, Rs1("OPRQTY").value)
'
'    Else
'        OPRQTY = 0
'
'    End If
'    Rs1.Close
'
'    sql = " SELECT      SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS IssuedOprQty"
'    sql = sql & "   FROM            dbo.Transactions INNER JOIN"
'    sql = sql & "                            dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
'    sql = sql & "                            dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'    sql = sql & "   Where (dbo.Transaction_Details.project_id1 = " & project_id1 & ")"
'    sql = sql & "    AND (dbo.Transaction_Details.Pand_ID = " & pand_id & ")"
'    sql = sql & "    AND (dbo.Transaction_Details.Oper_ID = " & oper_id & ")"
'    sql = sql & "    AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
'    If txtmodflage = "E" Then
'        sql = sql & "    AND (dbo.Transactions.Transaction_ID <>  " & Transaction_ID & ")"
'    End If
'
'    sql = sql & "   AND (dbo.TransactionTypes.projectInclude = 1)"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'
'        IssuedOprQty = IIf(IsNull(rs("IssuedOprQty").value), 0, rs("IssuedOprQty").value)
'
'    Else
'        IssuedOprQty = 0
'
'    End If
'    Dim remainqty As Double
'    remainqty = OPRQTY + IssuedOprQty
'    remainqty = remainqty - TransQty
'    CheckRemainsetyforprojectopr = remainqty
'
'    rs.Close
'    '     rs1.Close
'End Function
'
'Public Function ProjectItemsCheck(Optional ProjectID As Double, _
'                                  Optional ProjectDes_ID As Double, _
'                                  Optional OPRIDD As Double, _
'                                  Optional ItemID As Double) As Double
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = "  SELECT       count( dbo.TblMatrials.ItemID) projectItems"
'    sql = sql & "   FROM         dbo.TblMatrials RIGHT OUTER JOIN                       dbo.terms_operations ON dbo.TblMatrials.Opr = dbo.terms_operations.id LEFT OUTER JOIN                       dbo.TblItems ON dbo.TblMatrials.ItemID = dbo.TblItems.ItemID"
'    sql = sql & "    Where 1 = 1"
'    sql = sql & "   and  (dbo.TblMatrials.ProjectID =" & ProjectID & ")"
'
'    If ItemID <> 0 Then
'        sql = sql & "    and dbo.terms_operations.ProjectDes_ID =" & ProjectDes_ID
'        sql = sql & "   and  dbo.terms_operations.OPRIDD =" & OPRIDD
'        sql = sql & "   and  dbo.TblMatrials.ItemID  =" & ItemID
'
'    End If
'    '--where Qtyissue<0
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'
'        ProjectItemsCheck = IIf(IsNull(rs("projectItems").value), 0, rs("projectItems").value)
'
'    Else
'        ProjectItemsCheck = 0
'
'    End If
'    rs.Close
'End Function
'Public Function GetInstructorCode(Optional ByRef ID As Integer, _
'                                  Optional ByRef Fullcode As String, _
'                                  Optional Type1 As Integer = 0)
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    If Type1 = 0 Then
'        sql = "select * from TblInstructors where ID= " & ID
'    ElseIf Type1 = 1 Then
'        sql = "select * from TblInstructors where  FullCode ='" & Fullcode & "'"
'    End If
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'
'        ID = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
'        Fullcode = IIf(IsNull(rs("FullCode").value), "", rs("FullCode").value)
'    Else
'        ID = 0
'        Fullcode = ""
'    End If
'    rs.Close
'End Function
'
'Public Function GetInstudentGroupCode(Optional ByRef ID As Integer, _
'                                      Optional ByRef Fullcode As String, _
'                                      Optional Type1 As Integer = 0, _
'                                      Optional ByRef BranchID As Integer)
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    If Type1 = 0 Then
'        sql = "select * from TblStuGroup where ID= " & ID
'    ElseIf Type1 = 1 Then
'        sql = "select * from TblStuGroup where  Code ='" & Fullcode & "'"
'    End If
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'
'        BranchID = IIf(IsNull(rs("BranchID").value), 0, rs("BranchID").value)
'        ID = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
'        Fullcode = IIf(IsNull(rs("Code").value), "", rs("Code").value)
'    Else
'        BranchID = 0
'        ID = 0
'        Fullcode = ""
'    End If
'    rs.Close
'End Function
'
'Public Sub GetTriningStudentInformation(Optional ID As Double, _
'                                        Optional ByRef QualiID As Double, _
'                                        Optional ByRef SexID As Integer, _
'                                        Optional ByRef UQama As String, _
'                                        Optional ByRef phone As String, _
'                                        Optional ByRef Email As String, _
'                                        Optional ByRef Address As String, _
'                                        Optional ByRef DateBrithH As String, _
'                                        Optional ByRef DateBrith As Date, _
'                                        Optional ByRef Mobile As String, _
'                                        Optional ByRef BranchID As Integer = 0)
'    Dim sql As String
'    Dim Rs6 As ADODB.Recordset
'    Set Rs6 = New ADODB.Recordset
'    sql = " SELECT  *"
'    sql = sql & " From dbo.TblTrainingRequest"
'    sql = sql & " Where (TypeTrain = 1) And (ID = " & ID & ")"
'    Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs6.RecordCount > 0 Then
'        QualiID = IIf(IsNull(Rs6("QualiID").value), 0, Rs6("QualiID").value)
'        SexID = IIf(IsNull(Rs6("SexID").value), -1, Rs6("SexID").value)
'        UQama = IIf(IsNull(Rs6("UQama").value), "", Rs6("UQama").value)
'        phone = IIf(IsNull(Rs6("Phone").value), "", Rs6("Phone").value)
'        Email = IIf(IsNull(Rs6("Email").value), "", Rs6("Email").value)
'        Mobile = IIf(IsNull(Rs6("Mobile").value), "", Rs6("Mobile").value)
'        DateBrithH = IIf(IsNull(Rs6("DateBrithH").value), "", Rs6("DateBrithH").value)
'        Address = IIf(IsNull(Rs6("Address").value), "", Rs6("Address").value)
'        DateBrith = IIf(IsNull(Rs6("DateBrith").value), Date, Rs6("DateBrith").value)
'        BranchID = IIf(IsNull(Rs6("BranchID").value), 0, Rs6("BranchID").value)
'    Else
'        BranchID = 0
'        Mobile = ""
'        Address = ""
'        QualiID = 0
'        SexID = -1
'        UQama = ""
'        phone = ""
'        Email = ""
'
'    End If
'End Sub
'Public Sub GetContStudentInformation(Optional ContID As Double, _
'                                     Optional ByRef CompID As Double, _
'                                     Optional ByRef NoStud As Double)
'    Dim sql As String
'    Dim Rs6 As ADODB.Recordset
'    Set Rs6 = New ADODB.Recordset
'    sql = " SELECT  *"
'    sql = sql & " From dbo.TblContrStudent"
'    sql = sql & " Where (ContType = 1) And (ID = " & ContID & ")"
'    Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs6.RecordCount > 0 Then
'        CompID = IIf(IsNull(Rs6("CompID").value), 0, Rs6("CompID").value)
'        NoStud = IIf(IsNull(Rs6("NoStud").value), 0, Rs6("NoStud").value)
'    Else
'        CompID = 0
'        NoStud = 0
'    End If
'End Sub
'
'Public Sub GetNominStudentInformation(Optional ContID As Double, _
'                                      Optional ByRef CompID As Double, _
'                                      Optional ByRef NoStud As Double, _
'                                      Optional ByRef ContNoID As Double)
'    Dim sql As String
'    Dim Rs6 As ADODB.Recordset
'    Set Rs6 = New ADODB.Recordset
'    sql = " SELECT  *"
'    sql = sql & " From dbo.TblStuCandidacy"
'    sql = sql & " Where (ID = " & ContID & ")"
'    Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs6.RecordCount > 0 Then
'        CompID = IIf(IsNull(Rs6("CompID").value), 0, Rs6("CompID").value)
'        NoStud = IIf(IsNull(Rs6("NoStudCon").value), 0, Rs6("NoStudCon").value)
'        ContNoID = IIf(IsNull(Rs6("ContNoID").value), 0, Rs6("ContNoID").value)
'    Else
'        ContNoID = 0
'        CompID = 0
'        NoStud = 0
'    End If
'End Sub
'Public Function GetStudentCode(Optional ByRef ID As Integer, _
'                               Optional ByRef Fullcode As String, _
'                               Optional Type1 As Integer = 0, _
'                               Optional ByRef UQama As String)
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    If Type1 = 0 Then
'        sql = "select * from TblStudent where ID= " & ID
'    ElseIf Type1 = 1 Then
'        sql = "select * from TblStudent where  FullCode ='" & Fullcode & "'"
'    ElseIf Type1 = 2 Then
'        sql = "select * from TblStudent where  UQama ='" & UQama & "'"
'    End If
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'
'        ID = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
'        Fullcode = IIf(IsNull(rs("FullCode").value), "", rs("FullCode").value)
'        UQama = IIf(IsNull(rs("UQama").value), "", rs("UQama").value)
'    Else
'        UQama = ""
'        ID = 0
'        Fullcode = ""
'    End If
'    rs.Close
'End Function
'
'Public Function GetInformationofStudent(Optional ID As Integer, _
'                                        Optional ByRef UQama As String, _
'                                        Optional ByRef StudentPhone As String, _
'                                        Optional ByRef DcbQualiID As Double = 0)
'    Dim sql As String
'    Dim Rs9 As New ADODB.Recordset
'    sql = "select * from TblStudent where ID= " & ID
'    Rs9.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If Rs9.RecordCount > 0 Then
'        UQama = IIf(IsNull(Rs9("UQama").value), "", Rs9("UQama").value)
'        StudentPhone = IIf(IsNull(Rs9("StudentPhone").value), "", Rs9("StudentPhone").value)
'        DcbQualiID = IIf(IsNull(Rs9("DcbQualiID").value), 0, Rs9("DcbQualiID").value)
'    Else
'        StudentPhone = ""
'        UQama = ""
'        DcbQualiID = 0
'    End If
'    Rs9.Close
'End Function
'
'Public Function GetCursInformation(Optional ID As Double = 0, _
'                                   Optional ByRef NoHour As Double, _
'                                   Optional Price As Double)
'    Dim sql As String
'    Dim Rs5 As ADODB.Recordset
'    Set Rs5 = New ADODB.Recordset
'    sql = "Select * from TblStudentCurs where id=" & ID & ""
'    Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs5.RecordCount > 0 Then
'        Price = IIf(IsNull(Rs5("Price").value), 0, Rs5("Price").value)
'        NoHour = IIf(IsNull(Rs5("NoHour").value), 0, Rs5("NoHour").value)
'    Else
'        NoHour = 0
'        Price = 0
'    End If
'End Function
''*********************
'
'Public Function ChekEmpInProject(Optional EmpID As Integer = 0, _
'                                 Optional MonthID As Integer = 0, _
'                                 Optional YearID As Integer = 0) As Boolean
'    Dim sql As String
'    Dim Rs5 As ADODB.Recordset
'    Set Rs5 = New ADODB.Recordset
'    sql = " SELECT     dbo.opr_Employee.opr_type, dbo.opr_employee_details.pk_id, dbo.opr_employee_details.ContProjSalar, dbo.opr_employee_details.Emp_id, "
'    sql = sql & "                       dbo.opr_Employee.YEARS , dbo.opr_Employee.Months"
'    sql = sql & " FROM         dbo.opr_Employee LEFT OUTER JOIN"
'    sql = sql & "                      dbo.opr_employee_details ON dbo.opr_Employee.id = dbo.opr_employee_details.pk_id"
'    sql = sql & " WHERE     (dbo.opr_Employee.opr_type = 0) AND (dbo.opr_Employee.Years = " & YearID & ") AND (dbo.opr_Employee.Months = " & MonthID & ") AND (dbo.opr_employee_details.ContProjSalar = 2) AND"
'    sql = sql & "                      (dbo.opr_employee_details.Emp_id = " & EmpID & ")"
'    Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs5.RecordCount > 0 Then
'        ChekEmpInProject = True
'    Else
'        ChekEmpInProject = False
'    End If
'End Function
'
'Public Sub RetriveOrderInformation(Optional ID As Double, _
'                                   Optional ByRef ProgrammID As Double, _
'                                   Optional VehicleNo As Double)
'    Dim sql As String
'    Dim Rs5 As ADODB.Recordset
'    Set Rs5 = New ADODB.Recordset
'    sql = "Select * From tblbookingrequest  where id=" & ID & " and StusID=1"
'    Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, cmadcmdtext
'    If Rs5.RecordCount > 0 Then
'        ProgrammID = IIf(IsNull(Rs5("ProgrammID").value), 0, Rs5("ProgrammID").value)
'        VehicleNo = IIf(IsNull(Rs5("VehicleNo").value), 0, Rs5("VehicleNo").value)
'    Else
'        VehicleNo = 0
'        ProgrammID = 0
'    End If
'End Sub
'
'Public Sub GetProjectsBillInformation(Optional ID As Long, _
'                                      Optional ByRef project_no As String, _
'                                      Optional ByRef bill_to As Integer = 0, _
'                                      Optional ByRef branch_no As Integer, _
'                                      Optional ByRef total As Double, _
'                                      Optional ByRef Project_name As String, _
'                                      Optional ByRef ManualNO As String, _
'                                      Optional ByRef note_id As Double, _
'                                      Optional ByRef UserID As Long, _
'                                      Optional ByRef discount As Double, _
'                                      Optional ByRef advancedPayment As Double, _
'                                      Optional ByRef revenue_account As String, _
'                                      Optional ByRef Results As Double, _
'                                      Optional ByRef Remarks As String, _
'                                      Optional ByRef discount1value As Double, _
'                                      Optional ByRef discount2value As Double, _
'                                      Optional ByRef discount1ID As Integer, _
'                                      Optional ByRef discount2ID As Integer, _
'                                      Optional ByRef subContractorId As Long)
'    Dim Rs7 As ADODB.Recordset
'    Set Rs7 = New ADODB.Recordset
'    Dim sql As String
'    sql = "SELECT     dbo.projects.Project_name AS Project_nameH, dbo.projects.Project_nameE, dbo.project_billl.*"
'    sql = sql & " FROM         dbo.project_billl LEFT OUTER JOIN"
'    sql = sql & "                      dbo.projects ON dbo.project_billl.project_no = dbo.projects.id"
'    sql = sql & " where dbo.project_billl.id =" & ID & ""
'    Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs7.RecordCount > 0 Then
'        project_no = IIf(IsNull(Rs7("project_no").value), 0, Rs7("project_no").value)
'        bill_to = IIf(IsNull(Rs7("bill_to").value), 0, Rs7("bill_to").value)
'        branch_no = IIf(IsNull(Rs7("branch_no").value), 0, Rs7("branch_no").value)
'        total = IIf(IsNull(Rs7("total").value), 0, Rs7("total").value)
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Project_name = IIf(IsNull(Rs7("Project_nameH").value), "", Rs7("Project_nameH").value)
'        Else
'            Project_name = IIf(IsNull(Rs7("Project_nameE").value), "", Rs7("Project_nameE").value)
'        End If
'        ManualNO = IIf(IsNull(Rs7("ManualNO").value), 0, Rs7("ManualNO").value)
'        note_id = IIf(IsNull(Rs7("note_id").value), 0, Rs7("note_id").value)
'        UserID = IIf(IsNull(Rs7("UserID").value), 0, Rs7("UserID").value)
'        discount = IIf(IsNull(Rs7("Discount").value), 0, Rs7("Discount").value)
'        advancedPayment = IIf(IsNull(Rs7("advancedPayment").value), 0, Rs7("advancedPayment").value)
'        revenue_account = IIf(IsNull(Rs7("revenue_account").value), "", Rs7("revenue_account").value)
'        Results = IIf(IsNull(Rs7("Results").value), 0, Rs7("Results").value)
'        Remarks = IIf(IsNull(Rs7("Remarks").value), "", Rs7("Remarks").value)
'        discount1value = IIf(IsNull(Rs7("discount1value").value), 0, Rs7("discount1value").value)
'        discount2value = IIf(IsNull(Rs7("discount2value").value), 0, Rs7("discount2value").value)
'        discount1ID = IIf(IsNull(Rs7("discount1ID").value), -1, Rs7("discount1ID").value)
'        discount2ID = IIf(IsNull(Rs7("discount2ID").value), -1, Rs7("discount2ID").value)
'        subContractorId = IIf(IsNull(Rs7("subContractorId").value), 0, Rs7("subContractorId").value)
'    Else
'        project_no = ""
'        bill_to = 0
'        branch_no = 0
'        total = 0
'        Project_name = ""
'        ManualNO = ""
'        note_id = 0
'        UserID = 0
'        discount = 0
'        advancedPayment = 0
'        revenue_account = ""
'        Results = 0
'        Remarks = ""
'        discount1value = 0
'        discount2value = 0
'        discount1ID = -1
'        discount2ID = -1
'        subContractorId = 0
'    End If
'End Sub
'Public Sub GetInfomationDividInvestment(Optional ID As Double, _
'                                        Optional ByRef Nourth As Double, _
'                                        Optional ByRef South As Double, _
'                                        Optional ByRef East As Double, _
'                                        Optional ByRef West As Double, _
'                                        Optional ByRef Area As Double)
'    Dim Rs8 As ADODB.Recordset
'    Set Rs8 = New ADODB.Recordset
'    Dim sql As String
'    sql = " SELECT     ID, Nourth, South, East, West, Area"
'    sql = sql & " From dbo.TblDivInvestInformation"
'    sql = sql & " Where (ID = " & ID & ")"
'    Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs8.RecordCount > 0 Then
'        Nourth = IIf(IsNull(Rs8("Nourth").value), 0, Rs8("Nourth").value)
'        South = IIf(IsNull(Rs8("South").value), 0, Rs8("South").value)
'        East = IIf(IsNull(Rs8("East").value), 0, Rs8("East").value)
'        West = IIf(IsNull(Rs8("West").value), 0, Rs8("West").value)
'        Area = IIf(IsNull(Rs8("Area").value), 0, Rs8("Area").value)
'    Else
'        West = 0
'        South = 0
'        Nourth = 0
'        East = 0
'    End If
'End Sub
'
'Function CheckSusAccounts() As Boolean
'    Dim branch_name  As String
'    Dim branch_namee As String
'    Dim SUM          As Double
'    Dim i            As Integer
'    Dim SUMDebit     As Double
'    Dim SUMCrebit    As Double
'
'    CheckSusAccounts = False
'    Dim rsBranch As New ADODB.Recordset
'    My_SQL = "SELECT  * From TblBranchesData"
'
'    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rsBranch.RecordCount > 0 Then
'        rsBranch.MoveFirst
'    End If
'
'    For Branch = 1 To rsBranch.RecordCount
'        BramchId = IIf(IsNull(rsBranch("branch_id").value), 0, rsBranch("branch_id").value)
'        branch_name = IIf(IsNull(rsBranch("branch_name").value), 0, rsBranch("branch_name").value)
'        branch_namee = IIf(IsNull(rsBranch("branch_namee").value), 0, rsBranch("branch_namee").value)
'        SUMDebit = 0
'        SUMCrebit = 0
'        With FrmAccEditJournal.Fg_Journal
'
'            For i = .FixedRows To .Rows - 1
'
'                If val(.TextMatrix(i, .ColIndex("BranchID"))) = BramchId Then
'                    SUMDebit = SUMDebit + val(.TextMatrix(i, .ColIndex("DebitValue")))
'                    SUMCrebit = SUMCrebit + val(.TextMatrix(i, .ColIndex("CreditValue")))
'                End If
'
'            Next i
'            SUMDebit = Round(SUMDebit, SystemOptions.SysDefCurrencyForamt)
'            SUMCrebit = Round(SUMCrebit, SystemOptions.SysDefCurrencyForamt)
'
'            If val(SUMDebit) <> val(SUMCrebit) Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_name, vbCritical
'                Else
'                    MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_namee, vbCritical
'                End If
'                CheckSusAccounts = False
'                Exit Function
'
'            End If
'        End With
'        rsBranch.MoveNext
'    Next Branch
'    rsBranch.Close
'    CheckSusAccounts = True
'
'End Function
'
'Function CheckSusAccounts1() As Boolean
'    Dim branch_name  As String
'    Dim branch_namee As String
'    Dim SUM          As Double
'    Dim i            As Integer
'    Dim SUMDebit     As Double
'    Dim SUMCrebit    As Double
'    CheckSusAccounts1 = False
'    Dim rsBranch As New ADODB.Recordset
'    My_SQL = "SELECT  * From TblBranchesData"
'
'    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rsBranch.RecordCount > 0 Then
'        rsBranch.MoveFirst
'    End If
'
'    For Branch = 1 To rsBranch.RecordCount
'        BramchId = IIf(IsNull(rsBranch("branch_id").value), 0, rsBranch("branch_id").value)
'        branch_name = IIf(IsNull(rsBranch("branch_name").value), 0, rsBranch("branch_name").value)
'        branch_namee = IIf(IsNull(rsBranch("branch_namee").value), 0, rsBranch("branch_namee").value)
'        SUMDebit = 0
'        SUMCrebit = 0
'        With FrmAccEditJournal1.Fg_Journal
'
'            For i = .FixedRows To .Rows - 1
'
'                If val(.TextMatrix(i, .ColIndex("BranchID"))) = BramchId Then
'                    SUMDebit = SUMDebit + val(.TextMatrix(i, .ColIndex("DebitValue")))
'                    SUMCrebit = SUMCrebit + val(.TextMatrix(i, .ColIndex("CreditValue")))
'                End If
'
'            Next i
'            SUMDebit = Round(SUMDebit, SystemOptions.SysDefCurrencyForamt)
'            SUMCrebit = Round(SUMCrebit, SystemOptions.SysDefCurrencyForamt)
'
'            If val(SUMDebit) <> val(SUMCrebit) Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_name, vbCritical
'                Else
'                    MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_namee, vbCritical
'                End If
'                CheckSusAccounts1 = False
'                Exit Function
'
'            End If
'        End With
'        rsBranch.MoveNext
'    Next Branch
'    rsBranch.Close
'    CheckSusAccounts1 = True
'
'End Function
'Function CheckSusAccounts3() As Boolean
'    Dim branch_name  As String
'    Dim branch_namee As String
'    Dim SUM          As Double
'    Dim i            As Integer
'    Dim SUMDebit     As Double
'    Dim SUMCrebit    As Double
'    CheckSusAccounts3 = False
'    Dim rsBranch As New ADODB.Recordset
'    My_SQL = "SELECT  * From TblBranchesData"
'
'    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rsBranch.RecordCount > 0 Then
'        rsBranch.MoveFirst
'    End If
'
'    For Branch = 1 To rsBranch.RecordCount
'        BramchId = IIf(IsNull(rsBranch("branch_id").value), 0, rsBranch("branch_id").value)
'        branch_name = IIf(IsNull(rsBranch("branch_name").value), 0, rsBranch("branch_name").value)
'        branch_namee = IIf(IsNull(rsBranch("branch_namee").value), 0, rsBranch("branch_namee").value)
'        SUMDebit = 0
'        SUMCrebit = 0
'        With FrmAccEditJournal3.Fg_Journal
'
'            For i = .FixedRows To .Rows - 1
'
'                If val(.TextMatrix(i, .ColIndex("BranchID"))) = BramchId Then
'                    SUMDebit = SUMDebit + val(.TextMatrix(i, .ColIndex("DebitValue")))
'                    SUMCrebit = SUMCrebit + val(.TextMatrix(i, .ColIndex("CreditValue")))
'                End If
'
'            Next i
'            SUMDebit = Round(SUMDebit, SystemOptions.SysDefCurrencyForamt)
'            SUMCrebit = Round(SUMCrebit, SystemOptions.SysDefCurrencyForamt)
'
'            If val(SUMDebit) <> val(SUMCrebit) Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_name, vbCritical
'                Else
'                    MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_namee, vbCritical
'                End If
'                CheckSusAccounts3 = False
'                Exit Function
'
'            End If
'        End With
'        rsBranch.MoveNext
'    Next Branch
'    rsBranch.Close
'    CheckSusAccounts3 = True
'
'End Function
'
'Function CheckSusAccounts4() As Boolean
'    Dim branch_name  As String
'    Dim branch_namee As String
'    Dim SUM          As Double
'    Dim i            As Integer
'    Dim SUMDebit     As Double
'    Dim SUMCrebit    As Double
'    CheckSusAccounts4 = False
'    Dim rsBranch As New ADODB.Recordset
'    If SystemOptions.AllowUnbalncedByBranchAccount = True Then
'        CheckSusAccounts4 = True
'        Exit Function
'    End If
'    My_SQL = "SELECT  * From TblBranchesData"
'
'    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rsBranch.RecordCount > 0 Then
'        rsBranch.MoveFirst
'    End If
'
'    For Branch = 1 To rsBranch.RecordCount
'        BramchId = IIf(IsNull(rsBranch("branch_id").value), 0, rsBranch("branch_id").value)
'        branch_name = IIf(IsNull(rsBranch("branch_name").value), 0, rsBranch("branch_name").value)
'        branch_namee = IIf(IsNull(rsBranch("branch_namee").value), 0, rsBranch("branch_namee").value)
'        SUMDebit = 0
'        SUMCrebit = 0
'        With FrmAccEditJournal4.Fg_Journal
'
'            For i = .FixedRows To .Rows - 1
'
'                If val(.TextMatrix(i, .ColIndex("BranchID"))) = BramchId Then
'                    SUMDebit = SUMDebit + val(.TextMatrix(i, .ColIndex("DebitValue")))
'                    SUMCrebit = SUMCrebit + val(.TextMatrix(i, .ColIndex("CreditValue")))
'                End If
'
'            Next i
'            SUMDebit = Round(SUMDebit, SystemOptions.SysDefCurrencyForamt)
'            SUMCrebit = Round(SUMCrebit, SystemOptions.SysDefCurrencyForamt)
'
'            If val(SUMDebit) <> val(SUMCrebit) Then
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_name, vbCritical
'                Else
'                    MsgBox "«·ﬁÌœ €Ì— „ “‰ »«·‰”»… ·›—⁄ " & branch_namee, vbCritical
'                End If
'                CheckSusAccounts4 = False
'                Exit Function
'
'            End If
'        End With
'        rsBranch.MoveNext
'    Next Branch
'    rsBranch.Close
'    CheckSusAccounts4 = True
'
'End Function
'
'Public Function CheciIPOBySal_SharCount(Optional InvesID As Double = 0) As Boolean
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim sql As String
'    sql = " SELECT     TypeVaSh, InvesNo"
'    sql = sql & " From dbo.TblIPO"
'    sql = sql & " Where (TypeVaSh = 1) And (InvesNo = " & InvesID & ")"
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        CheciIPOBySal_SharCount = True
'    Else
'        CheciIPOBySal_SharCount = False
'    End If
'End Function
'
'Function getStorenames(StoreID As Double, _
'                       Optional StoreName As String, _
'                       Optional storenamee As String)
'    Dim Rs8 As ADODB.Recordset
'    Set Rs8 = New ADODB.Recordset
'    Dim sql As String
'
'    sql = "Select * from TblStore where StoreID=" & StoreID
'    Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs8.RecordCount > 0 Then
'        StoreName = IIf(IsNull(Rs8("StoreName").value), 0, Rs8("StoreName").value)
'        storenamee = IIf(IsNull(Rs8("StoreNamee").value), 0, Rs8("StoreNamee").value)
'
'    Else
'        StoreName = ""
'        storenamee = ""
'
'    End If
'
'End Function
'
'Public Sub GetLandInformation(Optional Land As Double = 0, Optional ByRef Area As Double)
'    Dim Rs8 As ADODB.Recordset
'    Set Rs8 = New ADODB.Recordset
'    Dim sql As String
'    sql = "Select * from TblBuyLanReEst where ID=" & Land & ""
'    Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs8.RecordCount > 0 Then
'        Area = IIf(IsNull(Rs8("Area").value), 0, Rs8("Area").value)
'    Else
'        Area = 0
'    End If
'End Sub
'
'Public Sub GetInvestInformation(Optional InvID As Double = 0, _
'                                Optional ByRef InvesTotal As Double, _
'                                Optional ByRef CountShare As Double, _
'                                Optional ByRef ShareValue As Double)
'    Dim Rs8 As ADODB.Recordset
'    Set Rs8 = New ADODB.Recordset
'    Dim sql As String
'    sql = "Select * from TblIPO where OrderInvse=" & InvID & ""
'    Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs8.RecordCount > 0 Then
'        InvesTotal = IIf(IsNull(Rs8("InvesTotal").value), 0, Rs8("InvesTotal").value)
'        CountShare = IIf(IsNull(Rs8("CountShare").value), 0, Rs8("CountShare").value)
'        ShareValue = IIf(IsNull(Rs8("ShareValue").value), 0, Rs8("ShareValue").value)
'    Else
'        ShareValue = 0
'        InvesTotal = 0
'        CountShare = 0
'    End If
'
'End Sub
''////////////////////////////////
'Public Sub SavedTranInvest(Optional IDIPO As Double = 0, _
'                           Optional BuyBilID As Double = 0, _
'                           Optional des As String, _
'                           Optional SharCount As Double = 0, _
'                           Optional ShareValue As Double = 0, _
'                           Optional InvesID As Double = 0, _
'                           Optional CusID As Double = 0, _
'                           Optional Effict As Double = 0)
'    Dim StrSQL As String
'    Dim Rs5    As ADODB.Recordset
'    Set Rs5 = New ADODB.Recordset
'    StrSQL = "SELECT  *  from TblTransactionInvest Where (1 = -1)"
'    Rs5.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    Rs5.AddNew
'    Rs5("IDIPO").value = IDIPO
'    Rs5("BuyBilID").value = BuyBilID
'    Rs5("Des").value = des
'    Rs5("SharCount").value = SharCount
'    Rs5("ShareValue").value = ShareValue
'    Rs5("InvesID").value = InvesID
'    Rs5("CusID").value = CusID
'    Rs5("Effict").value = Effict
'    Rs5.update
'End Sub
'''///////////////////////
'Public Function GetTotalSharOfCustomer(Optional CusID As Double = 0, _
'                                       Optional InvesID As Double = 0) As Double
'    Dim Rs8 As ADODB.Recordset
'    Set Rs8 = New ADODB.Recordset
'    Dim sql As String
'    sql = "   SELECT     SUM(SharCount * Effict) AS Totalshar, CusID, InvesID"
'    sql = sql & "  From dbo.TblTransactionInvest"
'    sql = sql & "  Where (CusID = " & CusID & ") And (InvesID = " & InvesID & ")"
'    sql = sql & "  GROUP BY CusID, InvesID"
'    Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs8.RecordCount > 0 Then
'        GetTotalSharOfCustomer = IIf(IsNull(Rs8("Totalshar").value), 0, Rs8("Totalshar").value)
'    Else
'        GetTotalSharOfCustomer = 0
'    End If
'End Function
'
'Public Function GetFixedAssetAddAccount(FixedassetId As Double) As String
'    Dim GroupID              As Integer
'    Dim FAADDAccount         As String
'    Dim ParetnAccount        As String
'    Dim GroupName            As String
'    Dim Account_Code5        As String
'    Dim StrSQL               As String
'    Dim X                    As String
'    Dim Account_Code_dynamic As String
'    'GetAllDataAboutFixedAsset CInt(FixedassetId), , GroupID, , , , , , , , , , , , , , , , , , , , , , , , FAADDAccount, ParetnAccount, GroupName, GroupNamee
'    If FAADDAccount = "" Then
'
'        If SystemOptions.AssetAccount = True Then
'            X = ParetnAccount
'
'            Account_Code5 = ModAccounts.AddNewAccount(X, " «÷«›«  " & GroupName, True, False, GroupNamee & "  Additions")
'
'        Else
'            Account_Code5 = ModAccounts.AddNewAccount(Account_Code_dynamic, " «÷«›«  " & GroupName, True, False, GroupNamee & "  Additions")
'        End If
'
'        StrSQL = "update FixedAssetsGroup  set  Account_Code5='" & Account_Code5 & "' where GroupID=" & GroupID
'        Cn.Execute StrSQL
'    End If
'
'    GetFixedAssetAddAccount = Account_Code5
'End Function
'
'Public Function GetSalaryEmployee(Optional Emp_id As Integer = 0)
'    If Emp_id <> 0 Then
'        Dim sql As String
'        Dim Rs9 As ADODB.Recordset
'        Set Rs9 = New ADODB.Recordset
'        sql = "select sum(DEV_Value1) as Total"
'        sql = sql & "  from("
'
'        sql = sql & " SELECT     dbo.EmpSalaryComponent.[Value] AS Total, dbo.mofrad.AddOrDiscount,"
'        sql = sql & " DEV_Value1=Case"
'        sql = sql & " When AddOrDiscount=0   Then Value * 1"
'        sql = sql & " Else  Value * -1"
'        sql = sql & " End"
'        sql = sql & " FROM         dbo.mofrad INNER JOIN"
'        sql = sql & " dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type RIGHT OUTER JOIN"
'        sql = sql & " dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode"
'        'sql = sql & " Where (dbo.EmpSalaryComponent.Emp_id = 2)"
'        sql = sql & " WHERE     (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
'        sql = sql & "   AND (dbo.mofrad.salary=1)"
'        sql = sql & " )x"
'        Rs9.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'        If Rs9.RecordCount > 0 Then
'            GetSalaryEmployee = IIf(IsNull(Rs9("Total").value), 0, Rs9("Total").value)
'        Else
'            GetSalaryEmployee = 0
'        End If
'    End If
'End Function
'
'Public Function CheckUnitContractMerg(unitno As Integer) As Boolean
'    Dim RsDetails As ADODB.Recordset
'    Dim StrSQL    As String
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = "SELECT     *  from dbo.TblIqrMerg Where (UntID =" & unitno & ")"
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    If RsDetails.RecordCount > 0 Then
'        CheckUnitContractMerg = True
'    Else
'        CheckUnitContractMerg = False
'    End If
'End Function
'Public Function ChekSanNumber(Optional branch_no As Integer = 0, _
'                              Optional Sanad_No As Integer = 0) As Boolean
'    Dim str As String
'    Dim Rs6 As ADODB.Recordset
'    Set Rs6 = New ADODB.Recordset
'    ChekSanNumber = False
'    str = "SELECT     branch_no, sanad_no"
'    str = str & " From dbo.sanad_numbering"
'    str = str & " Where(Sanad_No = " & Sanad_No & ") And (branch_no = " & branch_no & ") and numbering_id <> 0"
'    Rs6.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs6.RecordCount > 0 Then
'        ChekSanNumber = True
'    Else
'        ChekSanNumber = False
'    End If
'End Function
'
'Public Function Deletepost(ScreenName As String, _
'                           tablename As String, _
'                           FieldName As String, _
'                           DepID As Integer, _
'                           BranchID As Integer, _
'                           Transaction_ID As Variant, _
'                           NoteSerial As String)
'    Dim StrSQL As String
'    'StrSQL = "update " & tablename & " set  Posted=null ,PostedDate=null " & "where " & FieldName & "=" & Transaction_ID
'    'Cn.Execute StrSQL
'
'    StrSQL = "delete  ApprovalData  where Transaction_ID =" & Transaction_ID & "  and ScreenName='" & ScreenName & "'"
'    Cn.Execute StrSQL
'
'End Function
'Public Function SendTopost(ScreenName As String, _
'                           tablename As String, _
'                           FieldName As String, _
'                           DepID As Integer, _
'                           BranchID As Integer, _
'                           Transaction_ID As Variant, _
'                           NoteSerial As String, _
'                           Optional NoteID As Double = -1, _
'                           Optional EmpDepartemenID As Integer, _
'                           Optional OverProject As Double)
'
'    'user_id
'
'    Dim StrSQL As String
'    StrSQL = "update " & tablename & " set  Posted=" & user_id & ",PostedDate=" & SQLDate(Now, True) & "where " & FieldName & "=" & Transaction_ID
'    Cn.Execute StrSQL
'
'    Dim RSApproval As New ADODB.Recordset
'    Set RSApproval = New ADODB.Recordset
'    Dim currentdate As Date
'    RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'    Dim sql As String
'    Dim Rs1 As New ADODB.Recordset
'    Dim i   As Integer
'    'EmpDepartemenID = GetempDepartementidFromUserid(CInt(user_id))
'    sql = "  select  dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID ,dbo.TbllevelWorker.EmpID1 , "
'    sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
'    sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
'    sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
'    sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
'    sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & ScreenName & "')"
'    sql = sql & " and     (dbo.TblApprovalDef.BranchId =" & BranchID & ")"
'    If EmpDepartemenID <> 0 Then
'        sql = sql & " and        dbo.TblApprovalDef.DepartmentID  =" & EmpDepartemenID
'    End If
'
'    sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "
'
'    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    'Dim EmpDepartemenID As Integer
'    Dim UserID  As Integer
'    Dim UserID1 As Integer
'    Dim UserID2 As Integer
'    Dim EmpID   As Integer
'    currentdate = Now
'
'    GetApprovalDepartement DepID, UserID, EmpID, BranchID, UserID1, UserID2
'    Dim currcusor As Integer
'    currcusor = 1
'    If UserID <> 0 Then
'        '***************************************
'        RSApproval.AddNew
'        RSApproval("ScreenName").value = ScreenName
'        RSApproval("levelo").value = 0
'        RSApproval("EmpID").value = UserID
'        RSApproval("noteid").value = NoteID
'
'        RSApproval("levelorder").value = 0
'        RSApproval("currorder").value = 0
'        RSApproval("Transaction_ID").value = val(Transaction_ID)
'        RSApproval("overproject").value = val(OverProject)
'
'        RSApproval("NoteSerial").value = NoteSerial
'        RSApproval("Transaction_Date").value = Date
'
'        RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(ScreenName), currentdate)
'        RSApproval("SendTime").value = currentdate
'
'        RSApproval("Currcursor").value = currcusor
'        RSApproval("FromUser").value = user_name
'
'        RSApproval.update
'    End If
'
'    If UserID1 <> 0 Then
'        '***************************************
'        currcusor = currcusor + 1
'        RSApproval.AddNew
'        RSApproval("overproject").value = val(OverProject)
'        RSApproval("ScreenName").value = ScreenName
'        RSApproval("levelo").value = 0
'        RSApproval("EmpID").value = UserID1
'        RSApproval("levelorder").value = 0
'        RSApproval("currorder").value = 0
'        RSApproval("Transaction_ID").value = Transaction_ID
'        RSApproval("NoteSerial").value = NoteSerial
'        RSApproval("Transaction_Date").value = Date
'
'        RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(ScreenName), currentdate)
'        RSApproval("SendTime").value = currentdate
'
'        RSApproval("Currcursor").value = currcusor
'        RSApproval("FromUser").value = user_name
'
'        RSApproval("noteid").value = NoteID
'        RSApproval.update
'    End If
'
'    If UserID2 <> 0 Then
'        '***************************************
'        currcusor = currcusor + 1
'        RSApproval.AddNew
'        RSApproval("overproject").value = val(OverProject)
'        RSApproval("ScreenName").value = ScreenName
'        RSApproval("levelo").value = 0
'        RSApproval("EmpID").value = UserID2
'        RSApproval("levelorder").value = 0
'        RSApproval("currorder").value = 0
'        RSApproval("Transaction_ID").value = Transaction_ID
'        RSApproval("NoteSerial").value = NoteSerial
'        RSApproval("Transaction_Date").value = Date
'
'        RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(ScreenName), currentdate)
'        RSApproval("SendTime").value = currentdate
'
'        RSApproval("Currcursor").value = currcusor
'        RSApproval("FromUser").value = user_name
'
'        RSApproval("noteid").value = NoteID
'        RSApproval.update
'    End If
'    Dim Flag As Integer
'    Flag = 1
'    Dim empID2 As Integer
'    If Rs1.RecordCount > 0 Then
'
'        For i = 1 To Rs1.RecordCount
'
'            '****************************************
'            empID2 = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
'            If CheckWorkState(empID2) = True Then
'            Else
'                empID2 = IIf(IsNull(Rs1("empID1").value), 0, Rs1("empID1").value)
'            End If
'            If CheckWorkState(empID2) = True Then
'                RSApproval.AddNew
'                RSApproval("overproject").value = val(OverProject)
'                RSApproval("ScreenName").value = ScreenName
'                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
'
'                RSApproval("EmpID").value = empID2
'                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
'                RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
'                RSApproval("Transaction_ID").value = Transaction_ID
'                RSApproval("NoteSerial").value = NoteSerial
'                RSApproval("Transaction_Date").value = Date
'
'                RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(ScreenName), currentdate)
'                RSApproval("SendTime").value = currentdate
'
'                If Flag = 1 And UserID = 0 And UserID1 = 0 And UserID2 = 0 Then
'                    RSApproval("Currcursor").value = 1
'                    RSApproval("FromUser").value = user_name
'                    Flag = 2
'                End If
'                RSApproval("noteid").value = NoteID
'                RSApproval.update
'            End If
'            Rs1.MoveNext
'        Next i
'
'    End If
'
'End Function
'
'Public Function GetExchangReq(Optional ID As Double = 0, _
'                              Optional ByRef YearID As Integer, _
'                              Optional MonthID As Integer, _
'                              Optional ByRef BranchID As Integer) As String
'    Dim sql As String
'    Dim Rs9 As ADODB.Recordset
'    Set Rs9 = New ADODB.Recordset
'    GetExchangReq = ""
'    If ID <> 0 Then
'        sql = " SELECT      *"
'        sql = sql & " From dbo.TblExchangeRequest"
'        sql = sql & " Where (id = " & ID & ")"
'        Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Rs9.RecordCount > 0 Then
'            GetExchangReq = IIf(IsNull(Rs9("AllID").value), "", Rs9("AllID").value)
'            YearID = IIf(IsNull(Rs9("DurationID").value), 0, Rs9("DurationID").value)
'            MonthID = IIf(IsNull(Rs9("Month").value), 0, Rs9("Month").value)
'            BranchID = IIf(IsNull(Rs9("BranchID").value), 0, Rs9("BranchID").value)
'        Else
'            MonthID = 0
'            YearID = 0
'            GetExchangReq = ""
'            BranchID = 0
'        End If
'    End If
'End Function
'
''Public FirstPeriodAll As Date
''Select Emp_ID,Transaction_ID from transactions  where  Transaction_Type=61
'Public Function GetTblBuyLandRealEstate(Optional ByRef ID As Integer, _
'                                        Optional ByRef Fullcode As String, _
'                                        Optional Type1 As Integer = 0)
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    If Type1 = 0 Then
'        sql = "select * from TblBuyLanReEst where ID= " & ID
'    Else
'        sql = "select * from TblBuyLanReEst where  FullCode ='" & Fullcode & "'"
'    End If
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'
'        ID = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
'        Fullcode = IIf(IsNull(rs("FullCode").value), "", rs("FullCode").value)
'    Else
'        ID = 0
'        Fullcode = ""
'    End If
'    rs.Close
'End Function
'
'Public Sub GetVocationEntitlements(ID As Integer, _
'                                   Optional BranchID As Integer, _
'                                   Optional EmpID As Integer, _
'                                   Optional ByRef salary As Double, _
'                                   Optional ByRef SalEntitOther As Double, _
'                                   Optional ByRef other As Double, _
'                                   Optional ByRef Advance As Double, _
'                                   Optional ByRef ValueTickt As Double, _
'                                   Optional ByRef SalaryVocation As Double, _
'                                   Optional ByRef InsuranceValue As Double, _
'                                   Optional PreSalary As Double)
'
'    Dim StrSQL As String
'    Dim ch8    As Integer
'    Dim ch6    As Boolean
'    ch6 = False
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    StrSQL = "select * From tblVocationEntitlements     where id =" & ID & "  and not (NoteSerial is null) "
'
'    If CheckAprroveScreen("FrmVocationEntitlements") = True Then
'        StrSQL = StrSQL & " and approved =1"
'
'    End If
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If (rs.RecordCount) > 0 Then
'        BranchID = IIf(IsNull(rs("BranchID").value), Current_branch, (rs("BranchID").value))
'        EmpID = IIf(IsNull(rs("EmpID").value), 0, (rs("EmpID").value))
'
'        salary = Round(IIf(IsNull(rs("salary").value), 0, (rs("salary").value)), 2)
'        SalEntitOther = Round(IIf(IsNull(rs("SalEntitOther").value), 0, (rs("SalEntitOther").value)), 2)
'        other = Round(IIf(IsNull(rs("Other").value), 0, (rs("Other").value)), 2)
'        Advance = Round(IIf(IsNull(rs("Advance").value), 0, (rs("Advance").value)), 2)
'        ValueTickt = Round(IIf(IsNull(rs("ValueTickt").value), 0, (rs("ValueTickt").value)), 2)
'        SalaryVocation = Round(IIf(IsNull(rs("SalaryVocation").value), 0, (rs("SalaryVocation").value)), 2)
'        InsuranceValue = Round(IIf(IsNull(rs("InsuranceValue").value), 0, (rs("InsuranceValue").value)), 2)
'        PreSalary = Round(IIf(IsNull(rs("PreSalary").value), 0, (rs("PreSalary").value)), 2)
'        ch8 = Round(IIf(IsNull(rs("ch8").value), 0, (rs("ch8").value)), 2)
'        ch6 = IIf(IsNull(rs("ch6").value), False, (rs("ch6").value))
'        If ch6 = True Then
'            salary = Round(IIf(IsNull(rs("salary").value), 0, (rs("salary").value)), 2) - Round(IIf(IsNull(rs("Decrease").value), 0, (rs("Decrease").value)), 2)
'        End If
'        If ch8 = 0 Then
'            PreSalary = 0
'        End If
'    Else
'        BranchID = 0
'        EmpID = 0
'        salary = 0
'        SalEntitOther = 0
'        other = 0
'        ValueTickt = 0
'        Advance = 0
'    End If
'    rs.Close
'End Sub
'
'Public Sub GetEnd_Service(Optional ID As Double = 0, _
'                          Optional ByRef BranchID As Integer, _
'                          Optional ByRef EmpID As Double = 0, _
'                          Optional ByRef total As Double = 0, _
'                          Optional ByRef LastMonth As Double = 0, _
'                          Optional ByRef Ticket As Double = 0, _
'                          Optional ByRef Custom As Double = 0, _
'                          Optional ByRef net As Double = 0, _
'                          Optional ByRef TotalAdvance As Double = 0, _
'                          Optional ByRef TxtVlueVaction As Double = 0, _
'                          Optional ByRef TotalCash As Double, _
'                          Optional ByRef LastTotal As Double = 0, _
'                          Optional ByRef EndService As Double, _
'                          Optional ByRef CusTiket As Double, _
'                          Optional ByRef AddOther As Double, _
'                          Optional ByRef DiffTekit As Double, _
'                          Optional ByRef Discounts As Double, _
'                          Optional ByRef TicktConract As Double, _
'                          Optional ByRef DisSalary As Double)
'    If ID <> 0 Then
'        Dim sql As String
'        Dim Rs9 As ADODB.Recordset
'        Set Rs9 = New ADODB.Recordset
'        sql = "select * from End_of_service where id=" & ID & " and not (NoteSerial is null) "
'
'        If CheckAprroveScreen("End_oF_service") = True Then
'            StrWhere = StrWhere & " and approved =1"
'
'        End If
'
'        Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Rs9.RecordCount > 0 Then
'            EmpID = IIf(IsNull(Rs9("EmpID").value), 0, Rs9("EmpID").value)
'            total = Round(IIf(IsNull(Rs9("NetEnd").value), IIf(IsNull(Rs9("total").value), 0, Rs9("total").value), Rs9("NetEnd").value), 2)
'            LastMonth = Round(IIf(IsNull(Rs9("LastMonth").value), 0, Rs9("LastMonth").value), 2)
'            Ticket = Round(IIf(IsNull(Rs9("Ticket").value), 0, Rs9("Ticket").value), 2)
'            Custom = Round(IIf(IsNull(Rs9("Custom").value), 0, Rs9("Custom").value), 2)
'            net = Round(IIf(IsNull(Rs9("net").value), 0, Rs9("net").value), 2)
'            DiffTekit = Round(IIf(IsNull(Rs9("DiffTekit").value), 0, Rs9("DiffTekit").value), 2)
'            TotalAdvance = Round(IIf(IsNull(Rs9("TotalAdvance").value), 0, Rs9("TotalAdvance").value), 2)
'            TxtVlueVaction = Round(IIf(IsNull(Rs9("TxtVlueVaction").value), 0, Rs9("TxtVlueVaction").value), 2)
'            TotalCash = Round(IIf(IsNull(Rs9("TotalCash").value), 0, Rs9("TotalCash").value), 2)
'            LastTotal = Round(IIf(IsNull(Rs9("LastTotal").value), 0, Rs9("LastTotal").value), 2)
'            BranchID = Round(IIf(IsNull(Rs9("BranchID").value), 0, Rs9("BranchID").value), 2)
'            AddOther = Round(IIf(IsNull(Rs9("AddOther").value), 0, Rs9("AddOther").value), 2)
'            CusTiket = Round(IIf(IsNull(Rs9("CusTiket").value), 0, Rs9("CusTiket").value), 2)
'            EndService = Round(IIf(IsNull(Rs9("EndService").value), 0, Rs9("EndService").value), 2)
'            Discounts = Round(IIf(IsNull(Rs9("Discounts").value), 0, Rs9("Discounts").value), 2)
'            TicktConract = Round(IIf(IsNull(Rs9("TicktConract").value), 0, Rs9("TicktConract").value), 2)
'            DisSalary = Round(IIf(IsNull(Rs9("DisSalary").value), 0, Rs9("DisSalary").value), 2)
'        Else
'            TicktConract = 0
'            Discounts = 0
'            EndService = 0
'            CusTiket = 0
'            AddOther = 0
'            EmpID = 0
'            total = 0
'            LastMonth = 0
'            Ticket = 0
'            Custom = 0
'            net = 0
'            TotalAdvance = 0
'            TxtVlueVaction = 0
'            TotalCash = 0
'            LastTotal = 0
'            BranchID = 0
'            DiffTekit = 0
'            DisSalary = 0
'        End If
'    End If
'End Sub
'
'Public Sub GetEnd_Servicex13082017(Optional ID As Double = 0, _
'                                   Optional ByRef BranchID As Integer, _
'                                   Optional ByRef EmpID As Double = 0, _
'                                   Optional ByRef total As Double = 0, _
'                                   Optional ByRef LastMonth As Double = 0, _
'                                   Optional ByRef Ticket As Double = 0, _
'                                   Optional ByRef Custom As Double = 0, _
'                                   Optional ByRef net As Double = 0, _
'                                   Optional ByRef TotalAdvance As Double = 0, _
'                                   Optional ByRef TxtVlueVaction As Double = 0, _
'                                   Optional ByRef TotalCash As Double, _
'                                   Optional ByRef LastTotal As Double = 0, _
'                                   Optional ByRef EndService As Double, _
'                                   Optional ByRef CusTiket As Double, _
'                                   Optional ByRef AddOther As Double, _
'                                   Optional ByRef DiffTekit As Double, _
'                                   Optional ByRef Discounts As Double, _
'                                   Optional ByRef TicktConract As Double)
'    If ID <> 0 Then
'        Dim sql As String
'        Dim Rs9 As ADODB.Recordset
'        Set Rs9 = New ADODB.Recordset
'        sql = "select * from End_of_service where id=" & ID & " "
'        Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Rs9.RecordCount > 0 Then
'            EmpID = IIf(IsNull(Rs9("EmpID").value), 0, Rs9("EmpID").value)
'            total = Round(IIf(IsNull(Rs9("NetEnd").value), IIf(IsNull(Rs9("total").value), 0, Rs9("total").value), Rs9("NetEnd").value), 2)
'            LastMonth = Round(IIf(IsNull(Rs9("LastMonth").value), 0, Rs9("LastMonth").value))
'            Ticket = Round(IIf(IsNull(Rs9("Ticket").value), 0, Rs9("Ticket").value), 2)
'            Custom = Round(IIf(IsNull(Rs9("Custom").value), 0, Rs9("Custom").value), 2)
'            net = Round(IIf(IsNull(Rs9("net").value), 0, Rs9("net").value), 2)
'            DiffTekit = Round(IIf(IsNull(Rs9("DiffTekit").value), 0, Rs9("DiffTekit").value), 2)
'            TotalAdvance = Round(IIf(IsNull(Rs9("TotalAdvance").value), 0, Rs9("TotalAdvance").value), 2)
'            TxtVlueVaction = Round(IIf(IsNull(Rs9("TxtVlueVaction").value), 0, Rs9("TxtVlueVaction").value), 2)
'            TotalCash = Round(IIf(IsNull(Rs9("TotalCash").value), 0, Rs9("TotalCash").value), 2)
'            LastTotal = Round(IIf(IsNull(Rs9("LastTotal").value), 0, Rs9("LastTotal").value), 2)
'            BranchID = Round(IIf(IsNull(Rs9("BranchID").value), 0, Rs9("BranchID").value), 2)
'            AddOther = Round(IIf(IsNull(Rs9("AddOther").value), 0, Rs9("AddOther").value), 2)
'            CusTiket = Round(IIf(IsNull(Rs9("CusTiket").value), 0, Rs9("CusTiket").value), 2)
'            EndService = Round(IIf(IsNull(Rs9("EndService").value), 0, Rs9("EndService").value), 2)
'            Discounts = Round(IIf(IsNull(Rs9("Discounts").value), 0, Rs9("Discounts").value), 2)
'            TicktConract = Round(IIf(IsNull(Rs9("TicktConract").value), 0, Rs9("TicktConract").value), 2)
'        Else
'            TicktConract = 0
'            Discounts = 0
'            EndService = 0
'            CusTiket = 0
'            AddOther = 0
'            EmpID = 0
'            total = 0
'            LastMonth = 0
'            Ticket = 0
'            Custom = 0
'            net = 0
'            TotalAdvance = 0
'            TxtVlueVaction = 0
'            TotalCash = 0
'            LastTotal = 0
'            BranchID = 0
'            DiffTekit = 0
'        End If
'    End If
'End Sub
'
'Public Function ChekTransferNo(ChqueNum As String, _
'                               BankID As Double, _
'                               NoteID As Double, _
'                               ByRef NoteSerial1 As String) As Boolean
'    Dim Rs7 As ADODB.Recordset
'    Dim sql As String
'    ChekTransferNo = False
'    Set Rs7 = New ADODB.Recordset
'    sql = "SELECT     *"
'    sql = sql & " From notes"
'    sql = sql & "  WHERE     NoteCashingType=2"
'    sql = sql & "  and     NoteID<>" & NoteID
'    sql = sql & "  and     ChqueNum='" & ChqueNum & "'"
'    sql = sql & "  and     BankID=" & BankID
'
'    Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs7.RecordCount > 0 Then
'        ChekTransferNo = True
'
'        NoteSerial1 = IIf(IsNull(Rs7("NoteSerial1").value), "", Rs7("NoteSerial1").value)
'    Else
'        ChekTransferNo = False
'        NoteSerial1 = ""
'    End If
'End Function
'Public Function GetIncoiceNoByOrder(order_no As String) As String
'    Dim Rs7 As ADODB.Recordset
'    Dim sql As String
'    'GetIncoiceNoByOrder = False
'    Set Rs7 = New ADODB.Recordset
'    sql = "SELECT     NoteSerial1"
'    sql = sql & "   From dbo.transactions"
'    sql = sql & "   Where (Transaction_Type = 21 And CBoBasedON = 2)"
'    sql = sql & "   AND (order_no = '" & order_no & "')"
'
'    Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs7.RecordCount > 0 Then
'
'        GetIncoiceNoByOrder = IIf(IsNull(Rs7("NoteSerial1").value), "", Rs7("NoteSerial1").value)
'    Else
'        GetIncoiceNoByOrder = ""
'
'    End If
'End Function
'
'Public Function ChekClodePeriod(RecordDate As Date) As Boolean
'    Dim Rs7 As ADODB.Recordset
'    Dim sql As String
'    ChekClodePeriod = False
'    Set Rs7 = New ADODB.Recordset
'
'    sql = "SELECT     StartDate, EndDate"
'    sql = sql & " From dbo.TblAccountIntervals"
'    sql = sql & "  WHERE     (StartDate <=" & SQLDate(RecordDate, True) & " AND (EndDate >= " & SQLDate(RecordDate, True) & " ))and OpenState=1 "
'    Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs7.RecordCount > 0 Then
'        ChekClodePeriod = True
'    Else
'        ChekClodePeriod = False
'    End If
'End Function
'
'Public Function ChekClodePeriodx(RecordDate As Date) As Boolean
'    Dim Rs7 As ADODB.Recordset
'    Dim sql As String
'    ChekClodePeriodx = False
'    Set Rs7 = New ADODB.Recordset
'    sql = "SELECT     StartDate, EndDate"
'    sql = sql & " From dbo.TblOpenClosPeriodDet1"
'    sql = sql & "  WHERE     (StartDate <=" & SQLDate(RecordDate, True) & " AND (EndDate >= " & SQLDate(RecordDate, True) & " ))"
'    Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs7.RecordCount > 0 Then
'        ChekClodePeriodx = True
'    Else
'        ChekClodePeriodx = False
'    End If
'End Function
'
'Public Function GetInsuranceAccount(Optional ByRef Acount_Code1 As String, _
'                                    Optional ByRef Acount_Code2 As String, _
'                                    Optional ByRef CitizenVal1 As Double, _
'                                    Optional ByRef ResidentVal1 As Double, _
'                                    Optional ByRef Acount_Code4 As String, _
'                                    Optional ByRef Acount_Code3 As String, _
'                                    Optional ByRef CitizenVal2 As Double, _
'                                    Optional ByRef ResidentVal2 As Double)
'    Dim My_SQL As String
'    Dim Rs7    As ADODB.Recordset
'    Set Rs7 = New ADODB.Recordset
'
'    My_SQL = " SELECT    * from  TblSocialInsurance"
'
'    Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
'    If Rs7.RecordCount > 0 Then
'
'        Acount_Code1 = IIf(IsNull(Rs7("Acount_Code1").value), "", Rs7("Acount_Code1").value)
'        Acount_Code2 = IIf(IsNull(Rs7("Acount_Code2").value), "", Rs7("Acount_Code2").value)
'
'        Acount_Code3 = IIf(IsNull(Rs7("Acount_Code3").value), "", Rs7("Acount_Code3").value)
'        Acount_Code4 = IIf(IsNull(Rs7("Acount_Code4").value), "", Rs7("Acount_Code4").value)
'
'        CitizenVal1 = IIf(IsNull(Rs7("CitizenVal1").value), 0, Rs7("CitizenVal1").value)
'        ResidentVal1 = IIf(IsNull(Rs7("ResidentVal1").value), 0, Rs7("ResidentVal1").value)
'        CitizenVal2 = IIf(IsNull(Rs7("CitizenVal2").value), 0, Rs7("CitizenVal2").value)
'        ResidentVal2 = IIf(IsNull(Rs7("ResidentVal2").value), 0, Rs7("ResidentVal2"))
'
'    Else
'        Acount_Code1 = ""
'        Acount_Code2 = ""
'        Acount_Code3 = ""
'        Acount_Code4 = ""
'        CitizenVal1 = 0
'        ResidentVal1 = 0
'        CitizenVal2 = 0
'        ResidentVal2 = 0
'
'    End If
'
'End Function
'
'Public Function GetCartData(card As String, Optional Name As String = "") As Integer
'    Dim My_SQL As String
'    Dim Rs7    As ADODB.Recordset
'    Set Rs7 = New ADODB.Recordset
'
'    My_SQL = " SELECT     name, namee, tel, card, discount, CartTYpe"
'    My_SQL = My_SQL & "   From dbo.TblCusCsh"
'    My_SQL = My_SQL & "   WHERE     (card =  '" & card & "')"
'
'    Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
'    If Rs7.RecordCount > 0 Then
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Name = IIf(IsNull(Rs7("name").value), "", Rs7("name").value)
'        Else
'            Name = IIf(IsNull(Rs7("namee").value), "", Rs7("namee").value)
'        End If
'
'    Else
'        Name = ""
'    End If
'
'End Function
'Public Function GetApprovalDepartement(DeparmentID As Integer, _
'                                       Optional ByRef UserID As Integer, _
'                                       Optional ByRef EmpID As Integer, _
'                                       Optional ByRef BranchID As Integer, _
'                                       Optional ByRef UserID1 As Integer, _
'                                       Optional ByRef UserID2 As Integer) As Integer
'    Dim My_SQL As String
'    Dim Rs7    As ADODB.Recordset
'    Set Rs7 = New ADODB.Recordset
'
'    My_SQL = " SELECT  dbo.TblEmpDepartments.UserId1,dbo.TblEmpDepartments.UserId2,   dbo.TblEmpDepartments.UserId, dbo.TblEmpDepartments.DeparmentID, dbo.TblUsers.Empid"
'    My_SQL = My_SQL & " FROM         dbo.TblEmpDepartments INNER JOIN"
'    My_SQL = My_SQL & "                       dbo.TblUsers ON dbo.TblEmpDepartments.UserId = dbo.TblUsers.UserID"
'    My_SQL = My_SQL & "  Where (dbo.TblEmpDepartments.DeparmentID = " & DeparmentID & ")"
'    'My_SQL = My_SQL & "  and (dbo.TblEmpDepartments.BranchId = " & BranchID & ")"
'
'    Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
'    If Rs7.RecordCount > 0 Then
'        UserID = IIf(IsNull(Rs7("UserId").value), 0, Rs7("UserId").value)
'        UserID1 = IIf(IsNull(Rs7("UserId1").value), 0, Rs7("UserId1").value)
'        UserID2 = IIf(IsNull(Rs7("UserId2").value), 0, Rs7("UserId2").value)
'
'        EmpID = IIf(IsNull(Rs7("Empid").value), 0, Rs7("Empid").value)
'    Else
'        UserID = 0
'        UserID1 = 0
'        UserID2 = 0
'
'        EmpID = 0
'    End If
'
'End Function
'
'Public Function GetEmpIdfromProduction(Transaction_ID As Integer) As Integer
'    Dim My_SQL As String
'    Dim Rs7    As ADODB.Recordset
'    Set Rs7 = New ADODB.Recordset
'
'    My_SQL = "SELECT     Emp_id From dbo.transactions WHERE     (Transaction_ID = " & Transaction_ID & ")"
'    Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
'    If Rs7.RecordCount > 0 Then
'        GetEmpIdfromProduction = IIf(IsNull(Rs7("Emp_id").value), "", Rs7("Emp_id").value)
'    Else
'        GetEmpIdfromProduction = 0
'    End If
'
'End Function
'
'Public Function GetWaiterForTable(ID As Integer) As Integer
'    Dim My_SQL As String
'    Dim Rs7    As ADODB.Recordset
'    Set Rs7 = New ADODB.Recordset
'
'    My_SQL = "SELECT     Emp_id From dbo.Stables WHERE     (id = " & ID & ")"
'    Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
'    If Rs7.RecordCount > 0 Then
'        GetWaiterForTable = IIf(IsNull(Rs7("Emp_id").value), 0, Rs7("Emp_id").value)
'    Else
'        GetWaiterForTable = 0
'    End If
'
'End Function
'
'Public Function getAccountSerial_Code(Optional filed As String, _
'                                      Optional FiledWher As String, _
'                                      Optional str As String) As String
'    Dim My_SQL As String
'    Dim Rs7    As ADODB.Recordset
'    Set Rs7 = New ADODB.Recordset
'    If " & Filed &" <> "" Then
'        My_SQL = "  select " & filed & " as Acoud from ACCOUNTS where " & FiledWher & "='" & str & "'"
'        Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
'        If Rs7.RecordCount > 0 Then
'            getAccountSerial_Code = IIf(IsNull(Rs7("Acoud").value), "", Rs7("Acoud").value)
'        Else
'            getAccountSerial_Code = ""
'        End If
'    End If
'End Function
'Public Function CheckCartDiscount(value As Double) As Double
'    Dim rs     As New ADODB.Recordset
'    Dim StrSQL As String
'
'    If value = 0 Then
'        CheckCartDiscount = 0
'        Exit Function
'    End If
'    StrSQL = "SELECT     Lowsalary  , HighSalary, AdvValue"
'    StrSQL = StrSQL + "  From dbo.TblCustomerPoints"
'    StrSQL = StrSQL + " Where (Lowsalary <= " & value & ") And (HighSalary >= " & value & ")"
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (rs.BOF Or rs.EOF) Then
'        CheckCartDiscount = IIf(IsNull(rs("AdvValue").value), 0, rs("AdvValue").value)
'
'    Else
'        CheckCartDiscount = 0
'
'    End If
'
'    rs.Close
'    Set rs = Nothing
'End Function
'
'Public Function CheckSpecialOffer(overdate As Date, _
'                                  Optional ByRef brnchid As Double = -1, _
'                                  Optional ByRef Sales As Double = -1, _
'                                  Optional ByRef GetFree As Double = -1, _
'                                  Optional ByRef discount As Double = -1, _
'                                  Optional ByRef FromPrice As Double = -1) As Boolean
'    Dim rs     As New ADODB.Recordset
'    Dim StrSQL As String
'    StrSQL = " SELECT     dbo.TblItemShowDitailses.BrnchID, dbo.TblItemShowDitailses.Type, dbo.TblItemShows.StartSDate, dbo.TblItemShows.EndDate, dbo.TblItemShows.Sales, "
'    StrSQL = StrSQL + "                        dbo.TblItemShows.GetFree , dbo.TblItemShows.Discount, dbo.TblItemShows.FromPrice"
'    StrSQL = StrSQL + "  FROM         dbo.TblItemShowDitailses INNER JOIN"
'    StrSQL = StrSQL + "                        dbo.TblItemShows ON dbo.TblItemShowDitailses.ID2 = dbo.TblItemShows.ID"
'    StrSQL = StrSQL + "  Where (dbo.TblItemShowDitailses.type = 1) And (dbo.TblItemShowDitailses.BrnchID = " & brnchid & ") And (Not (dbo.TblItemShows.Sales Is Null))"
'
'    StrSQL = StrSQL + "  and     (" & SQLDate(overdate, True) & "BETWEEN dbo.TblItemShows.StartSDate AND dbo.TblItemShows.EndDate) "
'    StrSQL = StrSQL + "  AND (dbo.TblItemShows.TypePoliceP = 4) "
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (rs.BOF Or rs.EOF) Then
'        Sales = IIf(IsNull(rs("Sales").value), -1, rs("Sales").value)
'        GetFree = IIf(IsNull(rs("GetFree").value), -1, rs("GetFree").value)
'        discount = IIf(IsNull(rs("Discount").value), -1, rs("Discount").value)
'        FromPrice = IIf(IsNull(rs("FromPrice").value), -1, rs("FromPrice").value) ' 0 min   a max
'
'        CheckSpecialOffer = True
'    Else
'        CheckSpecialOffer = False
'
'    End If
'
'    rs.Close
'    Set rs = Nothing
'End Function
'Public Function CheckoverInbranch(Optional ID2 As Double = -1, _
'                                  Optional BranchID As Double = -1) As Boolean
'    Dim rs     As New ADODB.Recordset
'    Dim StrSQL As String
'
'    StrSQL = " SELECT     ID2, BrnchID, Type"
'    StrSQL = StrSQL + "    From dbo.TblItemShowDitailses"
'    StrSQL = StrSQL + "    Where (ID2 = " & ID2 & ") And (brnchid = " & BranchID & ")"
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (rs.BOF Or rs.EOF) Then
'
'        CheckoverInbranch = True
'    Else
'        CheckoverInbranch = False
'
'    End If
'
'End Function
'Public Function CheckItem(ItemID As Long, _
'                          overdate As Date, _
'                          Optional ByRef typedisid As Double = -1, _
'                          Optional ByRef discount As Double = -1, _
'                          Optional TypePoliceP As Double = -1, _
'                          Optional BranchID As Double = -1, _
'                          Optional freeitemid As Long, _
'                          Optional freeitemUnitid As Long, _
'                          Optional ByRef Amount As Long, _
'                          Optional freeitemQty As Double, _
'                          Optional qtyPercentage As Double) As Boolean
'    Dim rs          As New ADODB.Recordset
'    Dim ID2         As Double
'    Dim StrSQL      As String
'    Dim CurrWeekday As Integer
'    StrSQL = "SELECT  amount ,InfITemSho, "
'    StrSQL = StrSQL + "     TblItemShows.Sa ,"
'    StrSQL = StrSQL + "     TblItemShows.Su,"
'    StrSQL = StrSQL + "     TblItemShows.Mo,"
'    StrSQL = StrSQL + "     TblItemShows.Tu ,  "
'    StrSQL = StrSQL + "     TblItemShows.We,"
'    StrSQL = StrSQL + "     TblItemShows.Th,"
'    StrSQL = StrSQL + "      TblItemShows.Fr,"
'    StrSQL = StrSQL + "     dbo.TblItemShows.StartSDate , dbo.TblItemShows.EndDate, dbo.TblItemShowDitailses.ItemID, dbo.TblItemShowDitailses.discount, "
'    StrSQL = StrSQL + "                       dbo.TblItemShowDitailses.uniteid , dbo.TblItemShowDitailses.typedisid, dbo.TblItemShows.id"
'    StrSQL = StrSQL + " , dbo.TblItemShowDitailses.ID2 FROM         dbo.TblItemShows LEFT OUTER JOIN"
'    StrSQL = StrSQL + "                       dbo.TblItemShowDitailses ON dbo.TblItemShows.ID = dbo.TblItemShowDitailses.ID2"
'    StrSQL = StrSQL + "  WHERE     (" & SQLDate(overdate, True) & "BETWEEN dbo.TblItemShows.StartSDate AND dbo.TblItemShows.EndDate) AND (NOT (dbo.TblItemShowDitailses.ItemID IS NULL))"
'    StrSQL = StrSQL + "  and TblItemShowDitailses.ItemID=" & ItemID
'    If TypePoliceP = 4 Then
'        StrSQL = StrSQL + "  and TblItemShows.TypePoliceP=" & TypePoliceP
'    End If
'    CurrWeekday = Weekday(overdate, vbUseSystemDayOfWeek)
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    Dim Sa, Su, Mo, Tu, We, Th, Fr As Integer
'    If Not (rs.BOF Or rs.EOF) Then
'
'        Sa = IIf(IsNull(rs("Sa").value) Or rs("Sa").value = 0, 0, 1)
'        Su = IIf(IsNull(rs("Su").value) Or rs("Su").value = 0, 0, 2)
'        Mo = IIf(IsNull(rs("Mo").value) Or rs("Mo").value = 0, 0, 3)
'        Tu = IIf(IsNull(rs("Tu").value) Or rs("Tu").value = 0, 0, 4)
'        We = IIf(IsNull(rs("We").value) Or rs("We").value = 0, 0, 5)
'        Th = IIf(IsNull(rs("Th").value) Or rs("Th").value = 0, 0, 6)
'        Fr = IIf(IsNull(rs("Fr").value) Or rs("Fr").value = 0, 0, 7)
'        If CurrWeekday = Sa _
'           Or CurrWeekday = Su _
'           Or CurrWeekday = Mo _
'           Or CurrWeekday = Tu _
'           Or CurrWeekday = We _
'           Or CurrWeekday = Th _
'           Or CurrWeekday = Fr _
'           Or CurrWeekday = Sa _
'           Then
'            '***********************************************************************
'            InfITemSho = IIf(IsNull(rs("InfITemSho").value), -1, rs("InfITemSho").value)
'            typedisid = IIf(IsNull(rs("typedisid").value), -1, rs("typedisid").value)
'            discount = IIf(IsNull(rs("discount").value), -1, rs("discount").value)
'            AmountQTY = IIf(IsNull(rs("amount").value), 0, rs("amount").value)
'            If AmountQTY = 0 Then
'                AmountQTY = 1
'            End If
'            Amount = AmountQTY
'            '     Dim qtyPercentage As Integer
'            If InfITemSho <> "" Then
'                VarSet = Split(InfITemSho, "#", , vbTextCompare)
'
'                If VarSet(0) <> Empty Or VarSet(0) <> "" Then
'                    freeitemid = VarSet(0)
'                    freeitemUnitid = VarSet(1)
'
'                    freeitemQty = VarSet(2)
'
'                    qtyPercentage = freeitemQty / AmountQTY
'                Else
'                    qtyPercentage = 0
'                End If
'
'            End If
'
'            ID2 = IIf(IsNull(rs("ID2").value), -1, rs("ID2").value)
'
'            If CheckoverInbranch(ID2, BranchID) = True Then
'                CheckItem = True
'            Else
'                CheckItem = False
'            End If
'
'            '***********************************************************
'        Else
'            CheckItem = False
'
'        End If
'
'    Else
'        CheckItem = False
'
'    End If
'
'    rs.Close
'    Set rs = Nothing
'End Function
'Public Function DaysInMonth(rdate As Date) As Long
'    Dim yr   As Long
'    Dim mnth As Long
'    yr = year(rdate)
'    mnth = Month(rdate)
'    ' Return the number of days in the specified month.
'    DaysInMonth = day(DateSerial(yr, mnth + 1, 1) - 1)
'End Function
'
'Public Function get__Account(ID As Integer, _
'   filed As String) As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim sql As String
'    sql = "Select  " & filed & " from TblDoCumentsTypes where id=" & ID
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount = 0 Then
'        get__Account = ""
'        Exit Function
'    End If
'    If IsNull(Rs3(filed).value) Then
'        get__Account = ""
'        Exit Function
'    End If
'    If Not IsNull(Rs3(filed).value) Then
'        get__Account = Rs3(filed).value
'        Exit Function
'    End If
'    Rs3.Close
'
'End Function
'
'Public Function CheckItemSpecialOffer(ItemID As Long, _
'                                      overdate As Date, _
'                                      Optional ByRef typedisid As Double = -1, _
'                                      Optional ByRef Sales As Double = -1, _
'                                      Optional GetFree As Double = -1, _
'                                      Optional discount As Double = -1, _
'                                      Optional FromPrice As Double = -1, _
'                                      Optional TypePoliceP As Double = -1, _
'                                      Optional BranchID As Double = -1) As Boolean
'    Dim rs     As New ADODB.Recordset
'    Dim StrSQL As String
'    Dim ID2    As Double
'    StrSQL = "SELECT     dbo.TblItemShows.StartSDate, dbo.TblItemShows.EndDate, dbo.TblItemShowDitailses.ItemID, dbo.TblItemShows.discount, "
'    StrSQL = StrSQL + "                       dbo.TblItemShows.Sales , dbo.TblItemShows.GetFree, dbo.TblItemShows.FromPrice"
'    StrSQL = StrSQL + "  , dbo.TblItemShowDitailses.ID2 FROM         dbo.TblItemShows LEFT OUTER JOIN"
'    StrSQL = StrSQL + "                       dbo.TblItemShowDitailses ON dbo.TblItemShows.ID = dbo.TblItemShowDitailses.ID2"
'    StrSQL = StrSQL + "  WHERE     (" & SQLDate(overdate, True) & "BETWEEN dbo.TblItemShows.StartSDate AND dbo.TblItemShows.EndDate) AND (NOT (dbo.TblItemShowDitailses.ItemID IS NULL))"
'    StrSQL = StrSQL + "  and TblItemShowDitailses.ItemID=" & ItemID
'    If TypePoliceP = 4 Then
'        StrSQL = StrSQL + "  and TblItemShows.TypePoliceP=" & TypePoliceP
'    End If
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (rs.BOF Or rs.EOF) Then
'        ID2 = IIf(IsNull(rs("ID2").value), -1, rs("ID2").value)
'
'        Sales = IIf(IsNull(rs("Sales").value), -1, rs("Sales").value)
'        GetFree = IIf(IsNull(rs("GetFree").value), -1, rs("GetFree").value)
'        discount = IIf(IsNull(rs("Discount").value), -1, rs("Discount").value)
'        FromPrice = IIf(IsNull(rs("FromPrice").value), -1, rs("FromPrice").value)
'
'        If CheckoverInbranch(ID2, BranchID) = True Then
'            CheckItemSpecialOffer = True
'        Else
'            CheckItemSpecialOffer = False
'        End If
'
'        ' CheckItemSpecialOffer = True
'    Else
'        CheckItemSpecialOffer = False
'
'    End If
'
'    rs.Close
'    Set rs = Nothing
'End Function
'
'Public Function GetbalanceBar(AccountCode As String) As String
'    Dim Balance         As String
'    Dim balanceString   As String
'    Dim account_name    As String
'    Dim Account_NameEng As String
'    Dim account_serial  As String
'    Dim str             As String
'
'    WriteCustomerBalPublic AccountCode, Balance, balanceString, , , account_name, Account_NameEng, account_serial
'    If SystemOptions.UserInterface = ArabicInterface Then
'        str = "ﬂÊœ «·Õ”««» : " & account_serial & Chr(13)
'        str = str & "«”„ «·Õ”««» : " & account_name & Chr(13)
'        str = str & "—’Ìœ «·Õ”««» : " & balanceString & Chr(13)
'
'    Else
'
'        str = "Account Code: " & account_serial & Chr(13)
'        str = str & "Account Name: " & Account_NameEng & Chr(13)
'        str = str & "Balance: " & balanceString & Chr(13)
'
'    End If
'    GetbalanceBar = str
'End Function
'
'Public Function GetFixedIDFromCode(Optional code As String, _
'                                   Optional ByRef FixedID As Integer)
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = "select * from FixedAssets where  code ='" & code & "'"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        FixedID = IIf(IsNull(rs("id").value), 0, rs("id").value)
'
'    Else
'        FixedID = 0
'    End If
'
'    rs.Close
'
'End Function
'
'Public Function Reload(frmIn As Form)
'    Unload frmIn
'    Load frmIn
'    frmIn.show
'End Function
'
'Function GETLINKSQL(StoreName As Integer, _
'                    Optional myindex As Integer = 0, _
'                    Optional updatestate As String, _
'                    Optional ByRef groupcodes As String) As String
'
'    Dim GROUPSTR As String
'
'    If StoreName = 0 Then
'        Exit Function
'    End If
'
'    If updatestate = "" And myindex = 0 Then
'
'        StrSQL = "Select * From TblItems  where 1=0"
'        GoTo ll
'    End If
'
'    If myindex = 0 Then
'        If updatestate <> "E" And updatestate <> "N" Then
'
'            StrSQL = "Select * From TblItems  where IsArchive=0  "
'            GoTo ll
'        End If
'
'    End If
'
'    GROUPSTR = " (SELECT     dbo.TblLink_Item_To_Store_Details3.GroupID"
'    GROUPSTR = GROUPSTR + " FROM         dbo.TblLink_Item_To_StoreH INNER JOIN"
'    GROUPSTR = GROUPSTR + " dbo.TblLink_Item_To_Store_Details1 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details1.Ind INNER JOIN"
'    GROUPSTR = GROUPSTR + " dbo.TblLink_Item_To_Store_Details2 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details2.Ind INNER JOIN"
'    GROUPSTR = GROUPSTR + " dbo.TblLink_Item_To_Store_Details3 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details3.Ind"
'    GROUPSTR = GROUPSTR + " Where (dbo.TblLink_Item_To_Store_Details1.StoreId = " & val(StoreName) & ")"
'    GROUPSTR = GROUPSTR + " GROUP BY dbo.TblLink_Item_To_Store_Details3.GroupID)"
'
'    If myindex = 0 Then
'        getallgroupsdata GROUPSTR, groupcodes, updatestate
'        Strforitems = groupcodes
'
'    Else
'
'    End If
'
'    If myindex = 0 Then
'        StrSQL = " SELECT    distinct ItemID, ItemCode, ItemName, GroupID, HaveSerial, LastUpdate, UserID, PurchasePrice, SallingPrice, RequestLimit, CustomerPrice, DealerPrice, HaveGuarantee, "
'
'        StrSQL = StrSQL + " GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, ItemCase, ItemMaking, ItemMakingNew, code, Branch_NO, Fullcode,"
'        StrSQL = StrSQL + " prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, itemSerials, BinLocation, minvalueqty, MaxValueqty, FreeQty, barCodeNO, CatlogNO, FactoryNO, TemplateID,"
'        StrSQL = StrSQL + " ItemMaxDiscount, OverHead, Wight, Content, Dippre, Source, Typenew"
'
'    ElseIf myindex = 1 Then
'        StrSQL = "SELECT   distinct  ItemID, barCodeNO"
'    ElseIf myindex = 2 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            StrSQL = "SELECT   distinct  ItemID, ItemName"
'        Else
'            StrSQL = "SELECT   distinct  ItemID, ItemNamee"
'        End If
'    End If
'
'    StrSQL = StrSQL + " From dbo.TblItems"
'    StrSQL = StrSQL + " where  IsArchive =0 and GroupID in ("
'
'    StrSQL = StrSQL + " select GroupID from fullgroups () )"
'
'    '    StrSQL = StrSQL + " or itemid in("
'    '    StrSQL = StrSQL + " SELECT     dbo.TblLink_Item_To_Store_Details2.ItemID"
'    '    StrSQL = StrSQL + " FROM         dbo.TblLink_Item_To_StoreH INNER JOIN"
'    '    StrSQL = StrSQL + " dbo.TblLink_Item_To_Store_Details1 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details1.Ind INNER JOIN"
'    '    StrSQL = StrSQL + " dbo.TblLink_Item_To_Store_Details2 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details2.Ind INNER JOIN"
'    '    StrSQL = StrSQL + " dbo.TblLink_Item_To_Store_Details3 ON dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details3.Ind"
'    '    StrSQL = StrSQL + " Where (dbo.TblLink_Item_To_Store_Details1.StoreId = " & val(STORENAME) & ")"
'    '    StrSQL = StrSQL + " GROUP BY dbo.TblLink_Item_To_Store_Details2.ItemID"
'    '    StrSQL = StrSQL + " )"
'll:
'    GETLINKSQL = StrSQL
'
'End Function
'
'Function GETLINKSQLByActivity(XXXX As Integer, _
'                              Optional myindex As Integer = 0, _
'                              Optional updatestate As String, _
'                              Optional ByRef groupcodes As String) As String
'
'    Dim GROUPSTR As String
'
'    If user_id = 0 Then
'        Exit Function
'    End If
'
'    GROUPSTR = " SELECT     dbo.Groups.GroupID"
'    GROUPSTR = GROUPSTR + " FROM         dbo.Groups "
'    GROUPSTR = GROUPSTR + " Where     dbo.Groups.ActivityTypeId in (  "
'    GROUPSTR = GROUPSTR & "  SELECT     dbo.TblBranchesData.ActivityTypeId"
'    GROUPSTR = GROUPSTR & "   FROM         dbo.TblUsersBranches INNER JOIN"
'    GROUPSTR = GROUPSTR & "                        dbo.TblBranchesData ON dbo.TblUsersBranches.BranchID = dbo.TblBranchesData.branch_id"
'    GROUPSTR = GROUPSTR & "   Where (dbo.TblUsersBranches.UserID = " & user_id & ")"
'    GROUPSTR = GROUPSTR & " ) "
'
'    If 0 = 0 Then
'        getallgroupsdata GROUPSTR, groupcodes, "xx"
'        Strforitems = groupcodes
'    End If
'
'    If myindex = 0 Then 'grid
'        StrSQL = " SELECT    distinct ItemID, ItemCode, ItemName, GroupID, HaveSerial, LastUpdate, UserID, PurchasePrice, SallingPrice, RequestLimit, CustomerPrice, DealerPrice, HaveGuarantee, "
'
'        StrSQL = StrSQL + " GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, ItemCase, ItemMaking, ItemMakingNew, code, Branch_NO, Fullcode,"
'        StrSQL = StrSQL + " prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, itemSerials, BinLocation, minvalueqty, MaxValueqty, FreeQty, barCodeNO, CatlogNO, FactoryNO, TemplateID,"
'        StrSQL = StrSQL + " ItemMaxDiscount, OverHead, Wight, Content, Dippre, Source, Typenew"
'
'    ElseIf myindex = 1 Then 'item code combo
'        StrSQL = "SELECT   distinct  ItemID, barCodeNO"
'    ElseIf myindex = 2 Then
'        If SystemOptions.UserInterface = ArabicInterface Then 'ItemName  code combo
'            StrSQL = "SELECT   distinct  ItemID, ItemName"
'        Else
'            StrSQL = "SELECT   distinct  ItemID, ItemNamee"
'        End If
'    End If
'
'    StrSQL = StrSQL + " From dbo.TblItems"
'    StrSQL = StrSQL + " where  IsArchive =0 and GroupID in ("
'
'    StrSQL = StrSQL + " select GroupID from fullgroups () )"
'
'    GETLINKSQLByActivity = StrSQL
'
'End Function
'
'Function getallgroupsdata(Optional strIngroups As String = "", _
'                          Optional ByRef groupcodes As String, _
'                          Optional ByRef updateStatus As String)
'    Dim sql As String
'    'Dim groupcodes As String
'    On Error Resume Next
'    GoTo ll
'    'newwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
'    Dim StrSQL   As String
'    Dim rs       As ADODB.Recordset
'    Dim Fullcode As String
'    Set rs = New ADODB.Recordset
'    StrSQL = "select     Fullcode  FROM         dbo.Groups WHERE     groupID IN (" & strIngroups & ")"
'
'    groupcodes = "SELECT     Fullcode From dbo.TblItems WHERE     (Fullcode LIKE N'0') "
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    'rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (rs.RecordCount) > 0 Then
'        rs.MoveFirst
'        'OR (Fullcode LIKE N'1001%')
'
'        For i = 0 To rs.RecordCount
'            Fullcode = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
'            groupcodes = groupcodes & " OR (Fullcode LIKE N'" & Fullcode & "')"
'            rs.MoveNext
'        Next i
'        groupcodesPublic = groupcodes
'        '  groupCodes
'
'    End If
'
'    rs.Close
'
'    'newwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
'll:
'    'Exit Function
'    If updateStatus = "R" Or updateStatus = "" Then
'        Exit Function
'    End If
'
'    sql = " drop  FUNCTION fullgroups"
'    Cn.Execute sql
'
'    sql = "CREATE FUNCTION fullgroups ()"
'    sql = sql & " RETURNS @xTable TABLE"
'    sql = sql & " ("
'    sql = sql & " groupid INT,"
'    sql = sql & " parentidd INT,"
'    sql = sql & " Iteration INT"
'    sql = sql & " )"
'    sql = sql & " AS"
'    sql = sql & " Begin"
'    sql = sql & " DECLARE @rowsAdded INT"
'    sql = sql & " DECLARE @Iteration INT;"
'    sql = sql & " DECLARE @MaxRecursion INT ;   "
'    sql = sql & " set @Iteration=1;"
'    sql = sql & " set @MaxRecursion=10000;"
'    sql = sql & "  INSERT @xTable"
'
'    sql = sql & " SELECT groupid, parentid, @Iteration"
'    sql = sql & " From Groups"
'    sql = sql & " WHERE  groupid in  (  " & strIngroups & " )"
'    sql = sql & " SET @rowsAdded=@@rowcount ;"
'    sql = sql & " WHILE @rowsAdded > 0 AND @Iteration <= @MaxRecursion BEGIN"
'    sql = sql & "   INSERT @xTable"
'    sql = sql & "   SELECT e.groupid, parentid, @Iteration + 1"
'    sql = sql & " FROM groups e"
'    sql = sql & "    INNER JOIN @xTable r ON e.parentid = r.groupid"
'    sql = sql & " WHERE   parentid <> e.groupid AND r.Iteration = @Iteration;"
'    sql = sql & " SET @rowsAdded=@@rowcount;"
'    sql = sql & "    SET @Iteration = @Iteration + 1;"
'    sql = sql & " End"
'    sql = sql & " Return"
'    sql = sql & " End"
'
'    db_createOrUpdateFuctionSQL "fullgroups", sql
'
'End Function
'
'Function CreateRecusiveGroup(strIngroups As String, UserID As Double)
'    Dim sql As String
'    On Error Resume Next
'    sql = " drop  FUNCTION fullgroups" & UserID
'    Cn.Execute sql
'
'    sql = "CREATE FUNCTION fullgroups" & UserID & " ()"
'    sql = sql & " RETURNS @xTable TABLE"
'    sql = sql & " ("
'    sql = sql & " groupid INT,"
'    sql = sql & " parentidd INT,"
'    sql = sql & " Iteration INT"
'    sql = sql & " )"
'    sql = sql & " AS"
'    sql = sql & " Begin"
'    sql = sql & " DECLARE @rowsAdded INT"
'    sql = sql & " DECLARE @Iteration INT;"
'    sql = sql & " DECLARE @MaxRecursion INT ;   "
'    sql = sql & " set @Iteration=1;"
'    sql = sql & " set @MaxRecursion=10000;"
'    sql = sql & "  INSERT @xTable"
'
'    sql = sql & " SELECT groupid, parentid, @Iteration"
'    sql = sql & " From Groups"
'    sql = sql & " WHERE  groupid in  (  " & strIngroups & " )"
'    sql = sql & " SET @rowsAdded=@@rowcount ;"
'    sql = sql & " WHILE @rowsAdded > 0 AND @Iteration <= @MaxRecursion BEGIN"
'    sql = sql & "   INSERT @xTable"
'    sql = sql & "   SELECT e.groupid, parentid, @Iteration + 1"
'    sql = sql & " FROM groups e"
'    sql = sql & "    INNER JOIN @xTable r ON e.parentid = r.groupid"
'    sql = sql & " WHERE   parentid <> e.groupid AND r.Iteration = @Iteration;"
'    sql = sql & " SET @rowsAdded=@@rowcount;"
'    sql = sql & "    SET @Iteration = @Iteration + 1;"
'    sql = sql & " End"
'    sql = sql & " Return"
'    sql = sql & " End"
'
'    db_createOrUpdateFuctionSQL "fullgroups" & UserID, sql
'
'End Function
'Public Function checkRentAccount(Account_Code As String) As Boolean
'    Dim StrSQL As String
'    Dim rs     As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    StrSQL = "  SELECT       * "
'    StrSQL = StrSQL & " From dbo.ExpensesType"
'    StrSQL = StrSQL & " WHERE  Transportation=1 and   Account_Code='" & Account_Code & "'"
'
'    checkRentAccount = False
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If (rs.RecordCount) > 0 Then
'
'        checkRentAccount = True
'    Else
'        checkRentAccount = False
'
'    End If
'
'End Function
'Public Function checkmanyApproval(frmname As String)
'    Dim StrSQL As String
'    Dim rs     As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    StrSQL = "  SELECT     TOP 100 PERCENT ScreenName, ApprovName, ApprovNamee"
'    StrSQL = StrSQL & " From dbo.TblApprovalDef"
'    StrSQL = StrSQL & " WHERE     (ScreenName = N'" & frmname & "')"
'
'    checkmanyApproval = False
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If (rs.RecordCount) > 0 Then
'
'        If (rs.RecordCount) > 1 Then
'            checkmanyApproval = True
'
'        End If
'
'    End If
'
'End Function
'Public Function FillApprovedTableNew(ScreenName As String, _
'                                     Transaction_ID As Double, _
'                                     NoteSerial1 As String, _
'                                     ID As Integer)
'    Dim RSApproval As New ADODB.Recordset
'    Set RSApproval = New ADODB.Recordset
'    Dim currentdate As Date
'    RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'    Dim sql As String
'    Dim Rs1 As New ADODB.Recordset
'    Dim i   As Integer
'    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
'    sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
'    sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
'    sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
'    sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
'    sql = sql & " WHERE     (dbo.TblApprovalDef.id =" & ID & ")"
'    sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "
'
'    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Rs1.RecordCount > 0 Then
'        currentdate = Now
'        For i = 1 To Rs1.RecordCount
'            RSApproval.AddNew
'            RSApproval("ScreenName").value = ScreenName
'            RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
'            RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
'            RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
'            RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
'            RSApproval("Transaction_ID").value = Transaction_ID
'            RSApproval("NoteSerial").value = NoteSerial1
'            RSApproval("Transaction_Date").value = Date
'
'            RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(ScreenName), currentdate)
'            RSApproval("SendTime").value = currentdate
'
'            If i = 1 Then
'                RSApproval("Currcursor").value = 1
'                RSApproval("FromUser").value = user_name
'            End If
'
'            RSApproval.update
'            Rs1.MoveNext
'        Next i
'
'    End If
'
'End Function
'Public Function GetItemPriceByWitdth(Item_ID As Long, _
'                                     Width As Double, _
'                                     Optional ByVal LngUnitID As Long = 0) As Double
'    'Dim StrSQL  As String
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    StrSQL = "SELECT     dbo.Fn_GetPriceItem(" & Item_ID & ", " & Width & ") AS WidthPrice  "
'    StrSQL = StrSQL & " From dbo.TblItems"
'    StrSQL = StrSQL & "  Where (IsPriceIsPerview =1 and ItemID = " & Item_ID & ")"
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If (rs.RecordCount) > 0 Then
'        GetItemPriceByWitdth = IIf(IsNull(rs("WidthPrice").value), 0, (rs("WidthPrice").value))
'    Else
'        GetItemPriceByWitdth = 0
'    End If
'    If GetItemPriceByWitdth = 0 Then GetItemPriceByWitdth = GetItemPrice(Item_ID, , LngUnitID)
'End Function
'
'Public Function checkdataexist(StrSQL As String) As Boolean
'    'Dim StrSQL  As String
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'
'    checkdataexist = False
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If (rs.RecordCount) > 0 Then
'        checkdataexist = True
'    End If
'End Function
'
'Function getpricebycustomrContract(customerid As Double, _
'                                   UnitID As Long, _
'                                   ItemID As Long, _
'                                   Optional vendor As Integer = 0, _
'                                   Optional ByVal mSalesMan As Integer, _
'                                   Optional mCashCust As String = "", _
'                                   Optional Transaction_Date As Date) As Double
'    Dim StrSQL As String
'    Dim rs     As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim total As Double
'    If vendor = 0 Then
'        StrSQL = "SELECT     dbo.TblCustomerContractDetails.Price - dbo.TblCustomerContractDetails.Discount AS net"
'        StrSQL = StrSQL & "  FROM         dbo.TblCustomerContract INNER JOIN"
'        StrSQL = StrSQL & "                        dbo.TblCustomerContractDetails ON dbo.TblCustomerContract.TblCustomerContractD = dbo.TblCustomerContractDetails.TblCustomerContractD"
'        StrSQL = StrSQL & "  Where (dbo.TblCustomerContract.CustomerId = " & customerid & ")  And (dbo.TblCustomerContractDetails.ItemID = " & ItemID & ")"
'        StrSQL = StrSQL & " and " & SQLDate(Transaction_Date, True) & "   BETWEEN FromDate and Todate"
'        If customerid = 2 Then '   Õ«·Â «·⁄„Ì· «·‰ﬁœÌ
'            StrSQL = StrSQL & "  And ( dbo.TblCustomerContract.CashCustomerName  =  '" & mCashCust & "') "
'        End If
'
'        If SystemOptions.IsCustSalesManCashRelated Then 'Õ«·Â «·„‰œÊ»
'            StrSQL = StrSQL & "  And (  dbo.TblCustomerContract.Emp_ID  = " & mSalesMan & ")"
'        End If
'
'    Else
'
'        StrSQL = " SELECT     ISNULL(dbo.TblVendorContractDetails.Price, 0) - ISNULL(dbo.TblVendorContractDetails.Discount, 0) AS net"
'        StrSQL = StrSQL & " FROM         dbo.TblVendorContract INNER JOIN"
'        StrSQL = StrSQL & " dbo.TblVendorContractDetails ON dbo.TblVendorContract.TblVendorContractD = dbo.TblVendorContractDetails.TblVendorContractD"
'        StrSQL = StrSQL & "  Where (dbo.TblVendorContractDetails.ItemID = " & ItemID & ") And (dbo.TblVendorContract.VendorID = " & customerid & ")" 'And (dbo.TblVendorContractDetails.unitid = " & unitid & ")
'
'        'StrSQL = StrSQL & "  Where (dbo.TblCustomerContract.VendorId = " & customerid & ") And (dbo.TblVendorContract.UnitID = " & unitid & ") And (dbo.TblVendorContract.ItemID = " & ItemID & ")"
'
'    End If
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If (rs.RecordCount) > 0 Then
'
'        total = IIf(IsNull(rs("net").value), 0, (rs("net").value))
'
'        getpricebycustomrContract = total
'    Else
'        total = 0
'        getpricebycustomrContract = 0
'    End If
'
'End Function
'
'Function CheckChildforgroup(tablename As String, _
'                            GroupIDFild As String, _
'                            ParentIDFiles As String, _
'                            GroupIDValue As Integer) As Boolean
'    Dim StrSQL As String
'    Dim rs     As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim total As Double
'    StrSQL = "SELECT     COUNT(" & GroupIDFild & ") AS total"
'    StrSQL = StrSQL & "  From " & tablename & ""
'    StrSQL = StrSQL & "   WHERE     (" & ParentIDFiles & " = " & GroupIDValue & ")"
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If (rs.RecordCount) > 0 Then
'
'        total = IIf(IsNull(rs("total").value), 0, (rs("total").value))
'        If total > 0 Then
'            CheckChildforgroup = True
'        Else
'            CheckChildforgroup = False
'        End If
'
'    Else
'        CheckChildforgroup = False
'    End If
'
'End Function
'
'Function CheckCustomerSaleType(CusID As Double) As Double
'    Dim StrSQL As String
'    Dim rs     As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim total As Double
'    StrSQL = "SELECT    * "
'    StrSQL = StrSQL & "  From  TblCustemers"
'    StrSQL = StrSQL & "   WHERE CusID =" & CusID
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If (rs.RecordCount) > 0 Then
'
'        CheckCustomerSaleType = IIf(IsNull(rs("SaleType").value), 0, (rs("SaleType").value))
'
'    End If
'
'End Function
'Function GetProjectID(NoteID As Double)
'    Dim StrSQL As String
'    Dim rs     As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    StrSQL = "SELECT     project_id From dbo.Notes Where (NoteID = " & NoteID & ")"
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    'rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (rs.RecordCount) > 0 Then
'
'        GetProjectID = IIf(IsNull(rs("project_id").value), 0, (rs("project_id").value))
'
'    Else
'        GetProjectID = 0
'    End If
'
'    rs.Close
'
'End Function
'Function updateopeningbalanceNewFromsqlTrialBalance2(Optional Fromdate As Date, _
'                                                     Optional todate As Date, _
'                                                     Optional continous As Boolean = False, _
'                                                     Optional ActivityId As Integer = 0, _
'                                                     Optional BranchID As Integer = 0, _
'                                                     Optional Account_Code As String = "", _
'                                                     Optional updatetype As Integer = 0, _
'                                                     Optional composite As Boolean, _
'                                                     Optional lastlevel As Boolean = False)
'    'x1
'    '0 balance Sheet
'    '1 trial balances
'    Dim openingbalacedate As Date
'    ' getOpeningBalancedate P_DTPickerAccFrom , DTPickerAccTo.value, returnedfromdate, returnedTOdate, year(P_DTPickerAccFrom ), openingbalacedate
'    getOpeningBalancedate , , , , year(todate), openingbalacedate, continous
'
'    Dim StrSQL As String
'
'    If openingbalacedate = Fromdate Then
'
'        StrSQL = " update ACCOUNTS"
'        StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account)"
'        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', Account_code,0,last_account)"
'        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', Account_code,1,last_account)"
'
'        ' GetBalanceCreditORdepitByActivity(@fromdate datetime,@Todate datetime ,@accountcode as varchar(255),@Credit_Or_Debit as integer,@Activity_Id as integer )
'        If ActivityId <> 0 Then
'            StrSQL = " update ACCOUNTS"
'            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account)"
'            StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,0," & ActivityId & ",last_account)"
'            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,1," & ActivityId & ",last_account)"
'
'        End If
'
'        If BranchID <> 0 Then
'            StrSQL = " update ACCOUNTS"
'            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account)"
'            '      strsql = strsql & " balance= dbo.GetBalanceByBranch('" & SQLDate(fromdate) & "','" & SQLDate(todate) & "'," & BranchId & ", Account_code)"
'            StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,0," & BranchID & ",last_account)"
'            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,1," & BranchID & ",last_account)"
'
'        End If
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            openingbalanceDes = "—’Ìœ «›  «ÕÌ   ›Ì" & openingbalacedate
'        Else
'            openingbalanceDes = "Opening Balance In " & openingbalacedate
'        End If
'
'    Else
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            openingbalanceDes = "—’Ìœ Õ Ï    " & Fromdate - 1
'        Else
'            openingbalanceDes = " Balance Untill " & Fromdate - 1
'        End If
'
'        Dim FromDate1 As Date
'        FromDate1 = Fromdate - 1
'        StrSQL = " update ACCOUNTS"
'
'        StrSQL = StrSQL & " set  balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', Account_code,last_account),"
'        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)  +   isnull(dbo.GetBalance('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code ,last_account),0) "
'        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', Account_code,0,last_account)"
'        StrSQL = StrSQL & " ,CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', Account_code,1,last_account)"
'
'        If ActivityId <> 0 Then
'            StrSQL = " update ACCOUNTS"
'            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "'," & ActivityId & ", Account_code,last_account),"
'            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
'            StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,0," & ActivityId & ",last_account)"
'            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,1," & ActivityId & ",last_account)"
'
'        End If
'
'        If BranchID <> 0 Then
'
'            StrSQL = " update ACCOUNTS"
'            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "'," & BranchID & ", Account_code,last_account),"
'            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByBranch('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0) "
'            StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,0," & BranchID & ",last_account)"
'            StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,1," & BranchID & ",last_account)"
'
'        End If
'
'    End If
'
'    If updatetype = 1 Then  ' „Ì“«‰
'        StrSQL = StrSQL & " WHERE     (last_account = 1)  and  (AccountTypes = 1 or AccountTypes = 0)"
'
'    ElseIf updatetype = 5 Then  ' „Ì“«‰ „” ÊÌ« 
'
'        If Trim(Account_Code) <> "" Then
'            StrSQL = StrSQL & "  where  Account_Code in (" & Account_Code & ")"
'
'        End If
'
'        If lastlevel = False Then
'            StrSQL = StrSQL & " WHERE     (last_account = 0) "
'        End If
'
'    ElseIf updatetype = 2 Then  ' Õ”«» «” «–
'
'        If getAccountTypes(Account_Code) <> 1 Then ' ·Ê ﬂ«‰ Õ”«» „’—Ê› «Ê «Ì—«œ
'            GoTo Part2
'        End If
'
'        If Trim(Account_Code) <> "" Then
'            StrSQL = StrSQL & "  where  Parent_Account_Code ='" & Account_Code & "'"
'
'        End If
'
'    ElseIf updatetype = 3 Or updatetype = 4 Then '„‘—Ê⁄     '  ﬂ‘› Õ”«» „ÊŸ›
'
'        If Trim(Account_Code) <> "" Then
'            StrSQL = StrSQL & "  where  Account_Code in (" & Account_Code & ")"
'
'        End If
'
'    Else
'
'        'StrSQL = StrSQL & " WHERE     (last_account = 1) "
'    End If
'
'    'StrSQL = StrSQL & " WHERE     (last_account = 1) "
'
'    Cn.CommandTimeout = 10000
'
'    Cn.Execute StrSQL
'    'DoEvents
'
'    'part2****************************************************************************
'    If getAccountTypes(Account_Code) = 1 Then ' ·Ê ﬂ«‰ Õ”«»   „Ì“«‰Ì…
'        Exit Function
'    End If
'
'Part2:
'    openingbalacedate = GetOpeningBalanceDateForType2(Fromdate)
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        openingbalanceDes = "—’Ìœ Õ Ï    " & Fromdate - 1
'    Else
'        openingbalanceDes = " Balance Untill " & Fromdate - 1
'    End If
'
'    FromDate1 = Fromdate - 1
'    StrSQL = " update ACCOUNTS"
'
'    StrSQL = StrSQL & " set  balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', Account_code,last_account),"
'    StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)  +   isnull(dbo.GetBalance('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code ,last_account),0) "
'    StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', Account_code,0,last_account)"
'    StrSQL = StrSQL & " ,CreditBalance= dbo.GetBalanceCreditORdepit('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "', Account_code,1,last_account)"
'
'    If ActivityId <> 0 Then
'        StrSQL = " update ACCOUNTS"
'        StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "'," & ActivityId & ", Account_code,last_account),"
'        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
'        StrSQL = StrSQL & ", DepitBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,0," & ActivityId & ",last_account)"
'        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByActivity('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,1," & ActivityId & ",last_account)"
'
'    End If
'
'    If BranchID <> 0 Then
'
'        StrSQL = " update ACCOUNTS"
'        StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "'," & BranchID & ", Account_code,last_account),"
'        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByBranch('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0) "
'        StrSQL = StrSQL & " ,DepitBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,0," & BranchID & ",last_account)"
'        StrSQL = StrSQL & ", CreditBalance= dbo.GetBalanceCreditORdepitByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(todate) & "',Account_code,1," & BranchID & ",last_account)"
'
'    End If
'
'    If updatetype = 1 Then  ' „Ì“«‰
'        StrSQL = StrSQL & " WHERE     (last_account = 1)  and  (AccountTypes = 2) "
'    ElseIf updatetype = 5 Then 'ÿ„” ÊÌ«    ' „Ì“«‰
'        If Trim(Account_Code) <> "" Then
'            StrSQL = StrSQL & "  where  Account_Code in (" & Account_Code & ")"
'
'        End If
'        If lastlevel = False Then
'            StrSQL = StrSQL & " WHERE     (last_account = 0)  and  (AccountTypes = 2)"
'        End If
'
'    ElseIf updatetype = 2 Then  ' Õ”«» «” «–
'
'        If Trim(Account_Code) <> "" Then
'            StrSQL = StrSQL & "  where  Parent_Account_Code ='" & Account_Code & "'"
'
'        End If
'
'    ElseIf updatetype = 3 Or updatetype = 4 Then ' «Ê „‘—Ê⁄    '  ﬂ‘› Õ”«» „ÊŸ›
'
'        If Trim(Account_Code) <> "" Then
'            StrSQL = StrSQL & "  where    (AccountTypes = 2)  and Account_Code in (" & Account_Code & ")"
'
'        End If
'
'    Else
'
'        StrSQL = StrSQL & " WHERE     (last_account = 1) "
'    End If
'
'    'StrSQL = StrSQL & " WHERE     (last_account = 1) "
'
'    Cn.CommandTimeout = 10000
'
'    Cn.Execute StrSQL
'    'DoEvents
'
'End Function
'Public Sub GetVocationEntitlementsx(ID As Integer, _
'                                    Optional BranchID As Integer, _
'                                    Optional EmpID As Integer, _
'                                    Optional ByRef salary As Double, _
'                                    Optional ByRef SalEntitOther As Double, _
'                                    Optional ByRef other As Double, _
'                                    Optional ByRef Advance As Double, _
'                                    Optional ByRef ValueTickt As Double, _
'                                    Optional ByRef SalaryVocation As Double, _
'                                    Optional ByRef InsuranceValue As Double)
'
'    Dim StrSQL As String
'    Dim rs     As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    StrSQL = "select * From tblVocationEntitlements     where id =" & ID & " and PayedPayment is null "
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If (rs.RecordCount) > 0 Then
'        BranchID = IIf(IsNull(rs("BranchID").value), Current_branch, (rs("BranchID").value))
'        EmpID = IIf(IsNull(rs("EmpID").value), 0, (rs("EmpID").value))
'
'        salary = IIf(IsNull(rs("salary").value), 0, (rs("salary").value))
'        SalEntitOther = IIf(IsNull(rs("SalEntitOther").value), 0, (rs("SalEntitOther").value))
'        other = IIf(IsNull(rs("Other").value), 0, (rs("Other").value))
'        Advance = IIf(IsNull(rs("Advance").value), 0, (rs("Advance").value))
'        ValueTickt = IIf(IsNull(rs("ValueTickt").value), 0, (rs("ValueTickt").value))
'        SalaryVocation = IIf(IsNull(rs("SalaryVocation").value), 0, (rs("SalaryVocation").value))
'        InsuranceValue = IIf(IsNull(rs("InsuranceValue").value), 0, (rs("InsuranceValue").value))
'    Else
'        BranchID = 0
'        EmpID = 0
'        salary = 0
'        SalEntitOther = 0
'        other = 0
'        ValueTickt = 0
'        Advance = 0
'    End If
'    rs.Close
'End Sub
'
'Public Sub OrderExchange(Serial1 As String, _
'                         Optional ByRef Type1 As Integer, _
'                         Optional ByRef txtperson As String, _
'                         Optional ByRef des As String, _
'                         Optional ByRef Price As Double, _
'                         Optional ByRef EmpID As Integer, _
'                         Optional ByRef basedOn As Integer, _
'                         Optional ByRef orderNo As String, _
'                         Optional ByRef Transaction_ID As Integer, _
'                         Optional ByRef CusID As Double, _
'                         Optional ByRef FromType As Integer, _
'                         Optional ByRef Account_Code As String, _
'                         Optional CurrcyID As Integer, _
'                         Optional Rate As Double, _
'                         Optional valuee As Double, _
'                         Optional salary_or_advance As Integer)
'
'    Dim StrSQL As String
'    Dim rs     As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    StrSQL = "select * From TblExchange     where NoteSerial1 ='" & Serial1 & "'"
'    If SystemOptions.MonyeIssueVchrNoMust = True Then
'        StrSQL = StrSQL & "   and price >0 "
'    End If
'
'    If CheckAprroveScreen("FrmTypeExchange") = True Then
'
'        StrSQL = StrSQL & "  and Approved = 1"
'    End If
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    'rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (rs.RecordCount) > 0 Then
'
'        Account_Code = IIf(IsNull(rs("Account_Code").value), -1, (rs("Account_Code").value))
'        FromType = IIf(IsNull(rs("FromType").value), -1, (rs("FromType").value))
'        CusID = IIf(IsNull(rs("CusID").value), 0, (rs("CusID").value))
'        basedOn = IIf(IsNull(rs("basedOn").value), 0, (rs("basedOn").value))
'        Type1 = IIf(IsNull(rs("Type").value), 0, (rs("Type").value))
'        salary_or_advance = IIf(IsNull(rs("salary_or_advance").value), 0, (rs("salary_or_advance").value))
'
'        EmpID = IIf(IsNull(rs("EmpID").value), 0, (rs("EmpID").value))
'
'        Price = IIf(IsNull(rs("Price").value), 0, (rs("Price").value))
'        des = IIf(IsNull(rs("Des").value), "", (rs("Des").value))
'        txtperson = IIf(IsNull(rs("ToPerson").value), "", (rs("ToPerson").value))
'        orderNo = IIf(IsNull(rs("orderNo").value), 0, (rs("orderNo").value))
'        Transaction_ID = IIf(IsNull(rs("Transaction_ID").value), 0, (rs("Transaction_ID").value))
'        CurrcyID = IIf(IsNull(rs("CurrcyID").value), MainCurrency, (rs("CurrcyID").value))
'        valuee = IIf(IsNull(rs("PriceE").value), 0, (rs("PriceE").value))
'        Rate = IIf(IsNull(rs("Rate").value), 1, (rs("Rate").value))
'
'    Else
'        Price = -1
'    End If
'
'    rs.Close
'
'End Sub
'
'Public Sub OrderExchangeold(Serial1 As String, _
'                            Optional ByRef Type1 As Integer, _
'                            Optional ByRef txtperson As String, _
'                            Optional ByRef des As String, _
'                            Optional ByRef Price As Double, _
'                            Optional ByRef EmpID As Integer, _
'                            Optional ByRef basedOn As Integer, _
'                            Optional ByRef orderNo As String)
'
'    Dim StrSQL As String
'    Dim rs     As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    StrSQL = "select * From TblExchange     where NoteSerial1 ='" & Serial1 & "'"
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    'rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (rs.RecordCount) > 0 Then
'
'        basedOn = IIf(IsNull(rs("basedOn").value), 0, (rs("basedOn").value))
'        Type1 = IIf(IsNull(rs("Type").value), 0, (rs("Type").value))
'        EmpID = IIf(IsNull(rs("EmpID").value), 0, (rs("EmpID").value))
'        Price = IIf(IsNull(rs("Price").value), 0, (rs("Price").value))
'        des = IIf(IsNull(rs("des").value), "", (rs("des").value))
'        txtperson = IIf(IsNull(rs("ToPerson").value), "", (rs("ToPerson").value))
'        orderNo = IIf(IsNull(rs("orderNo").value), 0, (rs("orderNo").value))
'    Else
'        Price = -1
'    End If
'
'    rs.Close
'
'End Sub
'
'Public Function GetACCOUNTSCode(LngItemID As String, Optional ID As Integer = 0) As String
'    Dim StrSQL As String
'    Dim rs     As ADODB.Recordset
'
'    If LngItemID <> "" Then
'        If ID = 1 Then
'            StrSQL = "Select Account_Serial  From ACCOUNTS Where Account_Code='" & LngItemID & "'"
'        Else
'            StrSQL = "Select Account_Code  From ACCOUNTS Where Account_Serial='" & LngItemID & "'"
'        End If
'        Set rs = New ADODB.Recordset
'
'        If Cn.State = adStateClosed Then
'            open_my_connection
'        End If
'
'        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'        If Not (rs.BOF Or rs.EOF) Then
'            If ID = 1 Then
'                GetACCOUNTSCode = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
'            Else
'                GetACCOUNTSCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
'            End If
'        Else
'        End If
'
'        rs.Close
'        Set rs = Nothing
'    End If
'
'End Function
'Public Sub printCopounBarcode(m_PrintTarget As PrintTarget, Optional Serial1 As Double)
'
'    Dim MySQL        As String
'    Dim RsData       As New ADODB.Recordset
'    Dim xApp         As New CRAXDRT.Application
'    Dim xReport      As CRAXDRT.Report
'    Dim CViewer      As ClsReportViewer
'    Dim cCompanyInfo As ClsCompanyInfo
'    '  Name = Name & ".rpt"
'
'    If Dir(App.path & "\REPORTS\REPORTS NEW\" & "BarCodeCopoun.rpt") = "" Then
'        'GetMsgs 139, vbExclamation
'        Screen.MousePointer = vbDefault
'        MsgBox "«· ﬁ—Ì— €Ì— „ÊÃÊœ"
'        Exit Sub
'    End If
'
'    MySQL = " "
'
'    MySQL = " SELECT     dbo.TblCoupons.ID, dbo.TblCoupons.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCoupons.RecordDate, "
'    MySQL = MySQL & "                     dbo.TblCoupons.Remarks, dbo.TblCoupons.FromDate, dbo.TblCoupons.ToDate, dbo.TblCoupons.FromDate2, dbo.TblCoupons.ToDate2, dbo.TblCoupons.RdTyp,"
'    MySQL = MySQL & "                    dbo.TblCoupons.Num, dbo.TblCoupons.Vlue, dbo.TblCouponsDet.Remarks AS RemarksDet, dbo.TblCouponsDet.TypTrans, dbo.TblCouponsDet.Num AS NumDet,"
'    MySQL = MySQL & "                    dbo.TblCouponsDet.Vlue AS VlueDet, dbo.TblCouponsDet.FromVlue, dbo.TblCouponsDet.TOVlue, dbo.TblCouponsDet.BillNo, dbo.TblCouponsDet.ReturnBillNo,"
'    MySQL = MySQL & "                    dbo.TblCouponsDet.Transaction_ID, dbo.TblCouponsDet.RetTransaction_ID, dbo.TblCouponsDet.NewBillNo, dbo.TblCouponsDet.NewTransaction_ID,"
'    MySQL = MySQL & "                    dbo.TblCouponsDet.discount"
'    MySQL = MySQL & "     FROM         dbo.TblCoupons LEFT OUTER JOIN"
'    MySQL = MySQL & "                    dbo.TblCouponsDet ON dbo.TblCoupons.ID = dbo.TblCouponsDet.CoupID LEFT OUTER JOIN"
'    MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.TblCoupons.BranchID = dbo.TblBranchesData.branch_id"
'    MySQL = MySQL & "     Where ( TypTrans =1 and dbo.TblCoupons.ID =" & Serial1 & ")"
'    MySQL = MySQL & "ORDER BY dbo.TblCouponsDet.Vlue"
'
'    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'
'    Screen.MousePointer = vbArrowHourglass
'
'    Set xReport = xApp.OpenReport(App.path & "\REPORTS\REPORTS NEW\" & "BarCodeCopoun.rpt")
'    xReport.Database.SetDataSource RsData
'    Set cCompanyInfo = New ClsCompanyInfo
'    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
'
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.title
'    xReport.ReportAuthor = App.title
'
'    Set CViewer = New ClsReportViewer
'    hide_logo = True
'    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, App.path & "\REPORTS\REPORTS NEW\" & "BarCodeCopoun.rpt"
'
'    Set xApp = Nothing
'    Set xReport = Nothing
'    Screen.MousePointer = vbDefault
'    hide_logo = False
'End Sub
'Public Sub printCodeBarcode(m_PrintTarget As PrintTarget, _
'                            Optional Name As String, _
'                            Optional lblindex As Integer, _
'                            Optional UnitName As String = "")
'
'    Dim MySQL        As String
'    Dim RsData       As New ADODB.Recordset
'    Dim xApp         As New CRAXDRT.Application
'    Dim xReport      As CRAXDRT.Report
'    Dim CViewer      As ClsReportViewer
'    Dim cCompanyInfo As ClsCompanyInfo
'    Name = Name & ".rpt"
'
'    If Dir(App.path & "\Reports\Inventory\" & Name) = "" Then
'        'GetMsgs 139, vbExclamation
'        Screen.MousePointer = vbDefault
'        MsgBox "«· ﬁ—Ì— €Ì— „ÊÃÊœ"
'        Exit Sub
'    End If
'
'    MySQL = " "
'    If lblindex = 1 Then
'        MySQL = " SELECT  DealerPrice, code128,TblPrintBarCode.ProductionDate,   dbo.TblItems.ItemComment  ,  dbo.TblItems.TotalCalories, dbo.TblItems.shortName,   dbo.TblItems.PrintedName,    dbo.TblItems.ItemCode, dbo.TblItems.ItemNamee, dbo.TblPrintBarCode.Code, dbo.TblPrintBarCode.PartNo, dbo.TblPrintBarCode.Cost, dbo.TblPrintBarCode.Name, dbo.TblPrintBarCode.Color, "
'        MySQL = MySQL & "                          dbo.TblPrintBarCode.size, dbo.TblPrintBarCode.class, dbo.TblPrintBarCode.CodeAnalisys, dbo.TblPrintBarCode.ExpiryDate, dbo.TblPrintBarCode.LotNO, dbo.TblPrintBarCode.VatYou, dbo.TblPrintBarCode.VAT,"
'        MySQL = MySQL & "                          dbo.TblPrintBarCode.Total"
'        MySQL = MySQL & "  FROM            dbo.TblPrintBarCode LEFT OUTER JOIN"
'        MySQL = MySQL & "                          dbo.TblItems ON dbo.TblPrintBarCode.Item_ID = dbo.TblItems.ItemID"
'
'    Else
'        MySQL = " SELECT  DealerPrice,   code128,  dbo.TblItems.ItemComment  ,  dbo.TblItems.TotalCalories, dbo.TblItems.shortName,   dbo.TblItems.PrintedName,   dbo.TblPrintBarCode.Code, dbo.TblPrintBarCode.PartNo, dbo.TblPrintBarCode.Cost, dbo.TblPrintBarCode.Name, dbo.TblPrintBarCode.Color, dbo.TblPrintBarCode.size, dbo.TblPrintBarCode.class, "
'        MySQL = MySQL & "                        dbo.ItemsDetails.ItemDetailedCode , dbo.ItemsDetails.ItemID, dbo.TblItems.ItemNamee, dbo.TblPrintBarCode.VatYou, dbo.TblPrintBarCode.VAT, dbo.TblPrintBarCode.Total"
'        MySQL = MySQL & "          FROM            dbo.TblItems RIGHT OUTER JOIN"
'        MySQL = MySQL & "                        dbo.ItemsDetails ON dbo.TblItems.ItemID = dbo.ItemsDetails.ItemId RIGHT OUTER JOIN"
'        MySQL = MySQL & "                        dbo.TblPrintBarCode ON dbo.ItemsDetails.ParrtNoCode = dbo.TblPrintBarCode.Code"
'    End If
'
'    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'
'    Screen.MousePointer = vbArrowHourglass
'
'    '    If SystemOptions.UserInterface = EnglishInterface Then
'
'    '    Else
'
'    Set xReport = xApp.OpenReport(App.path & "\Reports\Inventory\" & Name)
'    xReport.Database.SetDataSource RsData
'    Set cCompanyInfo = New ClsCompanyInfo
'    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
'
'    '    End If
'    xReport.ParameterFields(4).AddCurrentValue UnitName
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.title
'    xReport.ReportAuthor = App.title
'
'    Set CViewer = New ClsReportViewer
'    hide_logo = True
'    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, App.path & "\Reports\Inventory\" & Name, , MySQL
'
'    Set xApp = Nothing
'    Set xReport = Nothing
'    Screen.MousePointer = vbDefault
'    hide_logo = False
'End Sub
'
'Public Function get_Customer_information(ID As Integer, Optional ByRef Mobile As String)
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim sql As String
'    sql = sql & " select * from TblCustemers "
'
'    sql = sql & " WHERE     (CusID = " & ID & ")"
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount > 0 Then
'
'        Mobile = IIf(IsNull(Rs3("Cus_mobile").value), "", Rs3("Cus_mobile").value)
'    End If
'End Function
'
'Public Function SavecustomerData(cuPhone As String, cuname As String)
'    Dim CustomerName As String
'    If cuPhone = "" Then
'        Exit Function
'    End If
'
'    CustomerName = GetCashCustomernamebyphone(cuPhone)
'    Dim ID As String
'
'    If CustomerName = "" Then
'        ID = new_id("TblCusCsh", "id", "")
'
'        add_record_to_table "TblCusCsh", "id,name,namee,tel", ID & ",'" & cuname & "','" & cuname & "','" & cuPhone & "'", "id", val(ID)
'
'    End If
'
'End Function
'Public Function GetCommisionPercentages(typeid As Integer, _
'                                        EmpID, _
'                                        Optional ByRef Rent As Double, _
'                                        Optional ByRef InternalComm As Double, _
'                                        Optional ByRef ExternalComm As Double, _
'                                        Optional ByRef Revenue As Double)
'    Dim RsDetails As ADODB.Recordset
'    Dim StrSQL    As String
'    Dim GroupID   As Integer
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = "SELECT     *  from dbo.TBLSalesRepData Where (EmpID =" & EmpID & ")"
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    If RsDetails.RecordCount > 0 Then
'        GroupID = IIf(IsNull(RsDetails("GroupID").value), 0, RsDetails("GroupID").value)
'
'    Else
'        GroupID = 0
'    End If
'
'    RsDetails.Close
'    Set RsDetails = Nothing
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = "SELECT     *  from dbo.TBLSalesRepGroups Where (typeid =" & typeid & ")"
'    If GroupID <> 0 Then
'        StrSQL = StrSQL & " and id=" & GroupID
'    End If
'
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    If RsDetails.RecordCount > 0 Then
'        Rent = IIf(IsNull(RsDetails("Rent").value), 0, RsDetails("Rent").value)
'        InternalComm = IIf(IsNull(RsDetails("InternalComm").value), 0, RsDetails("InternalComm").value)
'        ExternalComm = IIf(IsNull(RsDetails("ExternalComm").value), 0, RsDetails("ExternalComm").value)
'        Revenue = IIf(IsNull(RsDetails("Revenue").value), 0, RsDetails("Revenue").value)
'        typeid = IIf(IsNull(RsDetails("TypeiD").value), 0, RsDetails("TypeiD").value)
'    End If
'
'    If typeid > 2 Then
'        RsDetails.Close
'        Set RsDetails = Nothing
'        Set RsDetails = New ADODB.Recordset
'        StrSQL = "SELECT     *  from dbo.TBLSalesRepGroups Where (typeid =" & typeid & ")"
'        If GroupID <> 0 Then
'            StrSQL = StrSQL & " and id=" & GroupID
'        End If
'
'        RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'        If RsDetails.RecordCount > 0 Then
'            Rent = IIf(IsNull(RsDetails("Rent").value), 0, RsDetails("Rent").value)
'            InternalComm = IIf(IsNull(RsDetails("InternalComm").value), 0, RsDetails("InternalComm").value)
'            ExternalComm = IIf(IsNull(RsDetails("ExternalComm").value), 0, RsDetails("ExternalComm").value)
'            Revenue = IIf(IsNull(RsDetails("Revenue").value), 0, RsDetails("Revenue").value)
'            typeid = IIf(IsNull(RsDetails("TypeiD").value), 0, RsDetails("TypeiD").value)
'        End If
'
'    End If
'
'End Function
'Public Function checkDepositeRent(ID As Integer, ddate As Date) As String
'    Dim RsDetails As ADODB.Recordset
'    Dim StrSQL    As String
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = " SELECT     dbo.Notes.NoteDate, dbo.Notes.NoteSerial1, dbo.Notes.AllowDate, dbo.Notes.AllowDateH, dbo.Notes.renterName, dbo.Notes.akarid, dbo.TblAqar.aqarNo, "
'    StrSQL = StrSQL & "   dbo.TblAqar.aqarname, dbo.TblAkarUnit.name AS unittype, dbo.TblAkarUnit.namee AS unittypee, dbo.TblAqarDetai.unitno, dbo.TblAqarDetai.Id AS unitid"
'    StrSQL = StrSQL & "  FROM         dbo.Notes INNER JOIN"
'    StrSQL = StrSQL & "    dbo.TblAqar ON dbo.Notes.akarid = dbo.TblAqar.Aqarid INNER JOIN"
'    StrSQL = StrSQL & "     dbo.TblAkarUnit ON dbo.Notes.unittype = dbo.TblAkarUnit.id INNER JOIN"
'    StrSQL = StrSQL & "    dbo.TblAqarDetai ON dbo.Notes.UnitNo = dbo.TblAqarDetai.Id"
'    StrSQL = StrSQL & "   WHERE     (dbo.Notes.NoteDate <= " & SQLDate(ddate, True) & "  ) AND (dbo.TblAqarDetai.Id = " & ID & ")"
'
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    If RsDetails.RecordCount > 0 Then
'        checkDepositeRent = "ÌÊÃœ ⁄—»Ê‰ ⁄·Ï Â–… «·ÊÕœ… »√”„  " & IIf(IsNull(RsDetails("renterName").value), "", RsDetails("renterName").value) & "  Ì‰ ÂÌ ›Ì " & IIf(IsNull(RsDetails("AllowDate").value), "", RsDetails("AllowDate").value) & "    «·„Ê«›ﬁ " & IIf(IsNull(RsDetails("AllowDateh").value), "", RsDetails("AllowDateh").value)
'    Else
'        checkDepositeRent = ""
'    End If
'End Function
'
'Public Function checkEmpDiscount(EmpID As Integer, _
'                                 value As Double, _
'                                 discount As Double) As Boolean
'    Dim RsDetails     As ADODB.Recordset
'    Dim StrSQL        As String
'    Dim discountvalue As Double
'    checkEmpDiscount = False
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = "SELECT     *  from  TBLSalesRepData  Where (EmpID =" & EmpID & ")"
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    If RsDetails.RecordCount > 0 Then
'        discountvalue = IIf(IsNull(RsDetails("DiscountValue").value), 0, RsDetails("DiscountValue").value)
'        If value * discountvalue / 100 >= discount Then
'            checkEmpDiscount = True
'        Else
'            checkEmpDiscount = False
'        End If
'
'    Else
'        checkEmpDiscount = True
'    End If
'End Function
'Public Function getLastLevel() As Integer
'    Dim RsDetails As ADODB.Recordset
'    Dim StrSQL    As String
'
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = " SELECT     MAX([Level]) AS lastlevel FROM         dbo.AccountsLevelsDetails"
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    If RsDetails.RecordCount > 0 Then
'        getLastLevel = IIf(IsNull(RsDetails("lastlevel").value), 0, RsDetails("lastlevel").value)
'    Else
'        getLastLevel = 0
'    End If
'End Function
'Public Function checkContractTransactions(ContNo As Double) As Boolean
'    Dim RsDetails As ADODB.Recordset
'    Dim StrSQL    As String
'    checkContractTransactions = False
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = "SELECT     *   from dbo.Notes Where (ContNo =" & ContNo & ") and dbo.Notes.CashingType=8 "
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    If RsDetails.RecordCount > 0 Then
'        checkContractTransactions = True
'    Else
'        checkContractTransactions = False
'    End If
'End Function
'
'Public Function checkOutContract(ContNo As Integer) As Integer
'    Dim RsDetails As ADODB.Recordset
'    Dim StrSQL    As String
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = "SELECT     OutContract   from dbo.TblContract Where (ContNo =" & ContNo & ")"
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    If RsDetails.RecordCount > 0 Then
'        checkOutContract = IIf(IsNull(RsDetails("OutContract").value), 0, RsDetails("OutContract").value)
'    Else
'        checkOutContract = 0
'    End If
'End Function
'Public Function CheckUnitContract(unitno As Integer) As Boolean
'    Dim RsDetails As ADODB.Recordset
'    Dim StrSQL    As String
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = "SELECT     *  from dbo.TblContract Where (UnitNo =" & unitno & ")"
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    If RsDetails.RecordCount > 0 Then
'        CheckUnitContract = True
'    Else
'        CheckUnitContract = False
'    End If
'End Function
'
'Public Function CheckUnitContractxxx(unitno As Integer) As Boolean
'    Dim RsDetails As ADODB.Recordset
'    Dim StrSQL    As String
'    Set RsDetails = New ADODB.Recordset
'    StrSQL = "SELECT     *  from dbo.TblContract Where (UnitNo =" & unitno & ")"
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'    If RsDetails.RecordCount > 0 Then
'        CheckUnitContractxxx = True
'    End If
'End Function
'
'Public Function AlarmsDates()
'    Dim AskOption   As Boolean
'    Dim Askinterval As String
'    Dim Askcount    As Integer
'    On Error Resume Next
'    Dim i      As Integer
'    Dim rs     As ADODB.Recordset
'    Dim My_SQL As String
'    Set rs = New ADODB.Recordset
'    Dim rentInstallmentdate As Date
'    AskOption = GetSetting(StrAppRegPath, "View_Type", "RentInstallments", True)
'    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_RentInstallments", "")
'    Askcount = GetSetting(StrAppRegPath, "Setting", "Count_RentInstallments", 0)
'
'    If AskOption = True And Askinterval <> "" Then
'        rentInstallmentdate = DateAdd((Askinterval), 1 * Askcount, Date)
'
'    End If
'    My_SQL = " SELECT     dbo.TblContractInstallments.*, dbo.TblCustemers.CusName AS CusName, dbo.TblCustemers.Cus_mobile AS Cus_mobile, dbo.TblCustemers.CusNamee AS Expr3, "
'    My_SQL = My_SQL & "                      dbo.TblContract.NoteSerial1 AS NoteSerial11, dbo.TblContract.CusID AS Expr5, dbo.TblContract.StrDate AS Expr6, dbo.TblContractInstallments.Installdate AS Expr7,"
'    My_SQL = My_SQL & "                      dbo.TblContractInstallments.InstalldateH AS Expr9, dbo.TblContractInstallments.InstallNo AS Expr10, dbo.TblContractInstallments.Commissions AS Expr11,"
'    My_SQL = My_SQL & "                      dbo.TblContractInstallments.RentValue AS RentValue_1, dbo.TblContractInstallments.Insurance AS Insurance_1, dbo.TblContract.ContNo AS Expr8,"
'    My_SQL = My_SQL & "                      { fn IFNULL(dbo.ContracttBillInstallmentsDone.[Value], 0) } AS Allpayed, { fn IFNULL(dbo.TblContractInstallments.installValue, 0)"
'    My_SQL = My_SQL & "                      } - { fn IFNULL(dbo.ContracttBillInstallmentsDone.[Value], 0) } AS newremains, dbo.TblAqar.aqarNo AS IaqarNo, dbo.TblAqar.aqarname AS Iaqarname,"
'    My_SQL = My_SQL & "                      dbo.TblAkarUnit.name AS Unitname, dbo.TblAkarUnit.namee AS Unitnamee, dbo.TblAqarDetai.unitno AS unitnoNam, dbo.TblContract.Phone AS Phone"
'    My_SQL = My_SQL & " FROM         dbo.TblContractInstallments INNER JOIN"
'    My_SQL = My_SQL & "                      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo INNER JOIN"
'    My_SQL = My_SQL & "                      dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'    My_SQL = My_SQL & "                      dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
'    My_SQL = My_SQL & "                      dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
'    My_SQL = My_SQL & "                      dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
'    My_SQL = My_SQL & "                      dbo.ContracttBillInstallmentsDone ON dbo.TblContract.ContNo = dbo.ContracttBillInstallmentsDone.istallid"
'    My_SQL = My_SQL & " WHERE   (dbo.TblContractInstallments.Status = 0 OR dbo.TblContractInstallments.Status IS NULL)"
'
'    'My_SQL = My_SQL & " WHERE   (dbo.TblContractInstallments.Status = 0 OR dbo.TblContractInstallments.Status IS NULL)"
'
'    My_SQL = My_SQL + " and (Installdate <=" & SQLDate(rentInstallmentdate, True) & ")"
'
'    My_SQL = My_SQL + "   AND (dbo.TblContract.Branch_NO = " & Current_branch & ")"
'
'    My_SQL = My_SQL + "   order by Installdate "
'    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'
'        RSRentAlarm.show
'        RSRentAlarm.FillGrid My_SQL
'    End If
'
'End Function
'Function dhLastDayInMonth(Optional dtmDate As Date = 0) As Date
'    ' Return the last day in the specified month.
'    If dtmDate = 0 Then
'        ' Did the caller pass in a date? If not, use
'        ' the current date.
'        dtmDate = Date
'    End If
'    dhLastDayInMonth = DateSerial(year(dtmDate), _
'       Month(dtmDate) + 1, 0)
'End Function
'Public Function dhFirstDayInMonth(Optional dtmDate As Date = 0) As Date
'    ' Return the first day in the specified month.
'    If dtmDate = 0 Then
'        ' Did the caller pass in a date? If not, use
'        ' the current date.
'        dtmDate = Date
'    End If
'    dhFirstDayInMonth = DateSerial(year(dtmDate), _
'       Month(dtmDate), 1)
'End Function
'Public Function GheckLinkItem(Item_ID As Long, _
'   ByRef StoreID As Integer) As Boolean
'
'    Dim sql     As String
'    Dim rs      As New ADODB.Recordset
'    Dim Balance As Double
'
'    sql = "SELECT     ItemID, StoreID"
'    sql = sql & " from dbo.TblLink_Item_To_Store_Details2"
'    sql = sql & "  Where (StoreID = " & StoreID & ") And (ItemID = " & Item_ID & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        GheckLinkItem = True
'
'    Else
'        GheckLinkItem = False
'    End If
'
'    rs.Close
'
'End Function
'
'Public Function ShowAttachments(TxtSerial1 As String, txtopeation_type As String, Optional ByVal mmIDD As String = "")
'    If mmIDD = "" Then mmIDD = TxtSerial1
'    If TxtSerial1 = "" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'
'            MsgBox "·«»œ „‰ «Õ Ì«— ”‰œ «Ê·«": Exit Function
'        Else
'            MsgBox "Select Voucher Firstly": Exit Function
'        End If
'
'    End If
'    Dim mfrm As Form
'    If SystemOptions.IsBlue Then
'        Set mfrm = New imaged2
'    Else
'        Set mfrm = New imaged
'    End If
'    Set mfrm = New imaged
'
'    imaged.SUBJECT_NO = (TxtSerial1)
'    'imaged.mIDD = mmIDD
'    Unload imaged
'    imaged.show
'
'    If SystemOptions.UserInterface = EnglishInterface Then
'
'        imaged.Label9.Caption = "Voucher #"
'        imaged.Caption = "Voucher Attachment"
'         imaged.Label6.Caption = "Voucher #"
'    Else
'
'        imaged.Label9.Caption = "„—›ﬁ«  «·”‰œ    —ﬁ„"
'        imaged.Caption = "„—›ﬁ«  «·”‰œ  "
'       imaged.Label6.Caption = "—ﬁ„  «·”‰œ"
'
'    End If
'
'    imaged.SUBJECT_NO = (TxtSerial1)
'  imaged.txtopeation_type = txtopeation_type
''imaged.mIDD = mmIDD
'    imaged.Adodc1.CommandType = adCmdText
'    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
'
'
' Dim Position As Integer
'
'Position = InStr(1, TxtSerial1, "-")
'
'If Position > 0 Then
'Dim str1 As String
'Dim str2 As String
'Dim ConnectionStr As String
'str1 = mId(TxtSerial1, 1, Position - 1)
'str2 = mId(TxtSerial1, Position + 1, Len(TxtSerial1))
'ConnectionStr = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type
'ConnectionStr = ConnectionStr & "' and ( subject_no='" & (str1) & "'" & " or  subject_no='" & (TxtSerial1) & "') "
'  imaged.Adodc1.RecordSource = ConnectionStr
'
'Else
'imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
'End If
'
'
'
'    imaged.Adodc1.Refresh
'
'
'
'
'    If imaged.Adodc1.Recordset.RecordCount > 0 Then
'
'        imaged.DBPix201.Visible = True
'    Else
'        imaged.DBPix201.Visible = False
'    End If
'
'    If Position > 0 Then
'ConnectionStr = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type
'ConnectionStr = ConnectionStr & "' and ( subject_no='" & (str1) & "'" & " or  subject_no='" & (TxtSerial1) & "') "
'  imaged.Adodc4.RecordSource = ConnectionStr
'
'
'imaged.Adodc4.RecordSource = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
'End If
'
'imaged.Adodc4.RecordSource = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
'DoEvents
'imaged.Adodc4.Refresh
'    imaged.Adodc4.Refresh
'    imaged.DataGrid2.Refresh
'
'End Function
'Public Function ShowAttachments6(TxtSerial1 As String, _
'                                txtopeation_type As String, _
'                                Optional ByVal mmIDD As String = "")
'    If mmIDD = "" Then mmIDD = TxtSerial1
'    If TxtSerial1 = "" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'
'            MsgBox "·«»œ „‰ «Õ Ì«— ”‰œ «Ê·«"
'            Exit Function
'        Else
'            MsgBox "Select Voucher Firstly"
'            Exit Function
'        End If
'
'    End If
'    Dim mfrm As Form
'    If SystemOptions.IsBlue Then
'        Set mfrm = New imaged2
'    Else
'        Set mfrm = New imaged
'    End If
'
'    mfrm.SUBJECT_NO = (TxtSerial1)
'    mfrm.mIDD = mmIDD
'    Unload mfrm
'    '  mfrm.show
'    mfrm.SUBJECT_NO = (TxtSerial1)
'    mfrm.mIDD = mmIDD
'    ' imaged.SUBJECT_NO = (TxtSerial1)
'    ' imaged.mIDD = mfrm.mIDD
'    'imaged.txtopeation_type = txtopeation_type
'    '  imaged.show
'    If SystemOptions.UserInterface = EnglishInterface Then
'
'        mfrm.Label9.Caption = "Voucher #"
'        mfrm.Caption = "Voucher Attachment"
'        mfrm.Label6.Caption = "Voucher #"
'    Else
'
'        mfrm.Label9.Caption = "„—›ﬁ«  «·”‰œ    —ﬁ„"
'        mfrm.Caption = "„—›ﬁ«  «·”‰œ  "
'        mfrm.Label6.Caption = "—ﬁ„  «·”‰œ"
'
'    End If
'    mfrm.Adodc1.ConnectionString = Cn.ConnectionString
'    mfrm.Adodc2.ConnectionString = Cn.ConnectionString
'    mfrm.Adodc3.ConnectionString = Cn.ConnectionString
'    mfrm.Adodc4.ConnectionString = Cn.ConnectionString
'    mfrm.SUBJECT_NO = (TxtSerial1)
'    mfrm.txtopeation_type = txtopeation_type
'    mfrm.mIDD = mmIDD
'    mfrm.Adodc1.CommandType = adCmdText
'    mfrm.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
'    mfrm.Adodc4.RecordSource = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
'
'    Dim Position As Integer
'
'    Position = InStr(1, TxtSerial1, "-")
'
'    If Position > 0 Then
'        Dim str1          As String
'        Dim str2          As String
'        Dim ConnectionStr As String
'        str1 = mId(TxtSerial1, 1, Position - 1)
'        str2 = mId(TxtSerial1, Position + 1, Len(TxtSerial1))
'        ConnectionStr = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type
'        ConnectionStr = ConnectionStr & "' and ( subject_no='" & (str1) & "'" & " or  subject_no='" & (TxtSerial1) & "') "
'        mfrm.Adodc1.RecordSource = ConnectionStr
'
'    Else
'        mfrm.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
'    End If
'
'    'mfrm.Adodc1.Refresh
'mfrm.Adodc1.Refresh
''mfrm.Adodc2.Refresh
''mfrm.Adodc3.Refresh
'mfrm.Adodc4.Refresh
'    If mfrm.Adodc1.Recordset.RecordCount > 0 Then
'
'        mfrm.DBPix201.Visible = True
'    Else
'        mfrm.DBPix201.Visible = False
'    End If
'
'    If Position > 0 Then
'        ConnectionStr = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type
'        ConnectionStr = ConnectionStr & "' and ( subject_no='" & (str1) & "'" & " or  subject_no='" & (TxtSerial1) & "') "
'        mfrm.Adodc4.RecordSource = ConnectionStr
'
'    Else
'        mfrm.Adodc4.RecordSource = "SELECT * FROM Subject_doc WHERE operation_type = '" & txtopeation_type & "' and subject_no='" & (TxtSerial1) & "'"
'    End If
'    '  DoEvents
'    mfrm.Adodc1.Refresh
'    mfrm.Adodc4.Refresh
'    mfrm.DataGrid2.Refresh
'    mfrm.show 1
'    mfrm.Hide
'End Function
'Public Sub Translatefrm(Frm As Form)
'    Set frmTranslations.Frm = Frm
'
'    frmTranslations.show 1
'End Sub
'
'Public Function GetiItemsNewDetails(Optional uniteid As Integer = 0, _
'                                    Optional sizeid As Integer = 0, _
'                                    Optional ColorID As Integer = 0, _
'                                    Optional ClassId As Integer = 0, _
'                                    Optional ByRef UnitName As String, _
'                                    Optional ByRef sizename As String, _
'                                    Optional ByRef colorname As String, _
'                                    Optional ByRef classname As String)
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = "select UnitID,UnitName,UnitNamee from TblUnites"
'
'    sql = sql & " WHERE     (UnitID = " & uniteid & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            UnitName = IIf(IsNull(rs("unitname").value), "", rs("unitname").value)
'        Else
'            UnitName = IIf(IsNull(rs("UnitNamee").value), "", rs("UnitNamee").value)
'        End If
'
'    Else
'        UnitName = ""
'    End If
'    rs.Close
'
'    sql = "select  *   from TblItemsSizes "
'
'    sql = sql & " WHERE     (SizeId = " & sizeid & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            sizename = IIf(IsNull(rs("sizename").value), "", rs("sizename").value)
'        Else
'            sizename = IIf(IsNull(rs("sizename").value), "", rs("sizename").value)
'        End If
'
'    Else
'        sizename = ""
'    End If
'    rs.Close
'
'    sql = "select  *   from TblItemsColors "
'
'    sql = sql & " WHERE     (ColorID = " & ColorID & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            colorname = IIf(IsNull(rs("ColorName").value), "", rs("ColorName").value)
'        Else
'            colorname = IIf(IsNull(rs("ColorName").value), "", rs("ColorName").value)
'        End If
'
'    Else
'        colorname = ""
'    End If
'    rs.Close
'
'    sql = "select  *   from TblItemsclasses "
'
'    sql = sql & " WHERE     (SizeId = " & ClassId & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            classname = IIf(IsNull(rs("SizeName").value), "", rs("SizeName").value)
'        Else
'            classname = IIf(IsNull(rs("SizeName").value), "", rs("SizeName").value)
'        End If
'
'    Else
'        classname = ""
'    End If
'    rs.Close
'
'End Function
'
'Public Function GetGoldData(TTypeId As Integer, _
'                            typeid As Integer, _
'                            uniteid As Integer, _
'                            Optional ByRef UnitName As String, _
'                            Optional ByRef ttypename As String, _
'                            Optional ByRef typename As String)
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = "select UnitID,UnitName,UnitNamee from TblUnites"
'
'    sql = sql & " WHERE     (UnitID = " & uniteid & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            UnitName = IIf(IsNull(rs("unitname").value), "", rs("unitname").value)
'        Else
'            UnitName = IIf(IsNull(rs("UnitNamee").value), "", rs("UnitNamee").value)
'        End If
'
'    Else
'        UnitName = ""
'    End If
'    rs.Close
'
'    sql = "select   id,name,nameE  from TblGType "
'
'    sql = sql & " WHERE     (id = " & TTypeId & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            ttypename = IIf(IsNull(rs("name").value), "", rs("name").value)
'        Else
'            ttypename = IIf(IsNull(rs("nameE").value), "", rs("nameE").value)
'        End If
'
'    Else
'        ttypename = ""
'    End If
'    rs.Close
'
'End Function
'Public Function loadmyModule()
'    On Error Resume Next
'    Dim StrSQL As String
'    Dim rs     As New ADODB.Recordset
'    Dim ID     As Integer
'    Dim Pid    As Double
'    Dim code   As Double
'    Dim i      As Integer
'    'menumenu
'
'    mdifrmmain.CeramicEstimation.Visible = False
'    '    mdifrmmain.Reports.Visible = False
'    mdifrmmain.AgeingMAster.Visible = False
'    mdifrmmain.AssetsMngBase.Visible = False
'    mdifrmmain.rsInvestment.Visible = False
'    'mdifrmmain.planningMnu.Visible = False
'    mdifrmmain.POSTRansactiosG.Visible = False
'    mdifrmmain.SalesIns.Visible = False
'    mdifrmmain.shipmentMnu.Visible = False
'    mdifrmmain.ProductionPlan.Visible = False
'    mdifrmmain.MnuElevators.Visible = False
'    mdifrmmain.taxes.Visible = False
'    mdifrmmain.hajMnu.Visible = False
'    mdifrmmain.TransporterMain.Visible = False
'    mdifrmmain.CarMaintenance.Visible = False
'    mdifrmmain.Strategy.Visible = False
'    mdifrmmain.MnuMaintnance.Visible = False
'
'    mdifrmmain.StudentMenue.Visible = False
'    mdifrmmain.dev.Visible = False
'    mdifrmmain.mangDep.Visible = False
'    mdifrmmain.BankOp.Visible = False
'    mdifrmmain.MnuElevators.Visible = False
'    mdifrmmain.SalesIns.Visible = False
'    mdifrmmain.rsInvestment.Visible = False
'    mdifrmmain.MnuAccounts.Visible = False
'    mdifrmmain.Currency.Visible = False
'    mdifrmmain.LIFEINDICATORMNU.Visible = False
'    mdifrmmain.COLLECTIONS.Visible = False
'    mdifrmmain.Container.Visible = False
'    mdifrmmain.RealEstateMarketing.Visible = False
'
'    mdifrmmain.FinAnalysis.Visible = False
'    mdifrmmain.MNUFixedAssets.Visible = False
'    mdifrmmain.mnuEmployee.Visible = False
'    mdifrmmain.StockControl.Visible = False
'    mdifrmmain.Purchase.Visible = False
'
'    mdifrmmain.MarketingMnu.Visible = False
'    mdifrmmain.hajMnu.Visible = False
'    mdifrmmain.Sales.Visible = False
'    mdifrmmain.shipmentMnu.Visible = False
'    mdifrmmain.POSTRansactiosG.Visible = False
'    mdifrmmain.prdo.Visible = False
'    mdifrmmain.ProductionPlan.Visible = False
'    mdifrmmain.MnuProjects.Visible = False
'    mdifrmmain.TransporterMain.Visible = False
'    mdifrmmain.CarMaintenance.Visible = False
'    mdifrmmain.MnuMaintnance.Visible = False
'    mdifrmmain.Strategy.Visible = False
'    mdifrmmain.Archiving.Visible = False
'    mdifrmmain.LegalIssue.Visible = False
'    mdifrmmain.Tailor.Visible = False
'    mdifrmmain.rentcar.Visible = False
'    mdifrmmain.Beauty.Visible = False
'
'    If SystemOptions.SpecialVersion = True Then
'
'        mdifrmmain.AssetsMngReport(0).Visible = False
'        mdifrmmain.AssetsMngReport(8).Visible = False
'        mdifrmmain.AssetsMngReport(14).Visible = False
'        mdifrmmain.AssetsMng(2).Visible = False
'
'    End If
'
'    mdifrmmain.eye.Visible = False
'    mdifrmmain.gobus.Visible = False
'    'mdifrmmain.m2.Visible = False
'    'mdifrmmain.ArrowsBase.Visible = False
'    mdifrmmain.AssetsMngBase.Visible = False
'    mdifrmmain.Reports.Visible = False
'    mdifrmmain.Tools.Visible = False
'    mdifrmmain.Basicdata.Visible = False
'    mdifrmmain.dev.Visible = False
'    'mdifrmmain.planningMnu.Visible = False
'    mdifrmmain.tech.Visible = True
'
'    'mdifrmmain.MnueHouseMain.Visible = False
'    'mdifrmmain.FarmerMnue.Visible = False
'    'mdifrmmain.GoldMenu.Visible = False
'    mdifrmmain.mangDep.Visible = False
'
'    mdifrmmain.xyz.Visible = False
'    mdifrmmain.Farm.Visible = False
'
'    'SystemOptions.Ecnomy = True
'
'    If SystemOptions.Ecnomy = True Then
'
'        With mdifrmmain
'
'            .MnuAccounts.Visible = True
'            .Currency.Visible = True
'            .MNUFixedAssets.Visible = True
'            .mnuEmployee.Visible = True
'            .StockControl.Visible = True
'            .Purchase.Visible = True
'            .Sales.Visible = True
'            .Help.Visible = True
'            .MnuToolsSetPrinters(0).Visible = True
'            .Basicdata.Visible = True
'            .Tools.Visible = True
'            .Reports.Visible = True
'
'            .StockControlBasicSub(4).Visible = False
'            .StockControlBasicSub(5).Visible = False
'            .StockControlBasicSub(6).Visible = False
'            .StockControlBasicSub(7).Visible = False
'            .StockControlBasicSub(8).Visible = False
'            .PurchaseBasic(3).Visible = False
'            .PurchaseBasic(4).Visible = False
'            .PurchaseBasic(4).Visible = False
'
'            .Expenses(0).Visible = False
'            .Expenses(1).Visible = False
'
'            .ExpensesSub(0).Visible = False
'            .ExpensesSub(1).Visible = False
'            .Cashing(1).Visible = False
'            .MnuBoxDrawing.Visible = False
'            .MNUFixedAssets.Visible = False
'            '.xxxxx(6).Visible = False
'            '.emptyMnu.Visible = False
'            .mnuEmployeeBasic(2).Visible = False
'            .mnuEmployeeBasic(3).Visible = False
'            .mnuEmployeeBasic(4).Visible = False
'            .mnuEmployeeBasic(5).Visible = False
'            .Vscstionsssub(0).Visible = False
'            .Vscstionsssub(1).Visible = False
'            .Vscstionsssub(2).Visible = False
'            .Vscstionsssub(5).Visible = False
'
'            .mnuEmployeeBasic(8).Visible = False
'
'            .StockControlBasicSub(12).Visible = False
'            .TradingTransaction(1).Visible = False
'            .TradingTransactionSub1(0).Visible = False
'            .TradingTransaction(7).Visible = False
'            .TradingTransaction(9).Visible = False
'            .TradingTransaction(10).Visible = False
'            .PurchaseBasic(1).Visible = False
'            .PurchaseTransactionssubs(0).Visible = False
'            .PurchaseTransactionssubs(2).Visible = False
'            .PurchaseTransactionssubs1(0).Visible = False
'            .PurchaseTransactions(1).Visible = False
'            .PurchaseTransactions(2).Visible = False
'
'            .SalesBasicSubsub(0).Visible = False
'            .SalesBasicSubsub(2).Visible = False
'            .SalesBasicSub(2).Visible = False
'            .SalesBasicSub(4).Visible = False
'            .SalesBasicSub(5).Visible = False
'            .SalesBasicSub(6).Visible = False
'            .SalesBasicSub(9).Visible = False
'            .SalesBasicSub(10).Visible = False
'            .SalesTransactionssubss00(0).Visible = False
'            .SalesTransactionssubss00(2).Visible = True
'            .SalesTransactionssubss000(0).Visible = False
'            .SalesTransactions(4).Visible = False
'            .SalesTransactions(5).Visible = False
'            .SalesTransactions(6).Visible = False
'            .SalesTransactions(8).Visible = False
'            .SalesTransactions(11).Visible = False
'
'            GoTo Lite
'
'        End With
'    End If
'
'    code = 10111982
'
'    StrSQL = "SELECT *  From Pmanger "
'
'    '    StrSQL = "SELECT branch_id,branch_namee From TblBranchesData"
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        For i = 1 To rs.RecordCount
'            ID = IIf(IsNull(rs("id").value), "", rs("id").value)
'            Pid = IIf(IsNull(rs("Pid").value), "", rs("Pid").value)
'
'            If ID = 1 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.MnuAccounts.Visible = True
'                Else
'                    mdifrmmain.MnuAccounts.Visible = False
'                End If
'            End If
'
'            If ID = 2 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Currency.Visible = True
'                Else
'                    mdifrmmain.Currency.Visible = False
'                End If
'            End If
'
'            If ID = 3 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.FinAnalysis.Visible = True
'                Else
'                    mdifrmmain.FinAnalysis.Visible = False
'                End If
'            End If
'
'            If ID = 4 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.MNUFixedAssets.Visible = True
'                Else
'                    mdifrmmain.MNUFixedAssets.Visible = False
'                End If
'            End If
'
'            If ID = 5 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.mnuEmployee.Visible = True
'                Else
'                    mdifrmmain.mnuEmployee.Visible = False
'                End If
'            End If
'
'            If ID = 6 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.StockControl.Visible = True
'                Else
'                    mdifrmmain.StockControl.Visible = False
'                End If
'            End If
'
'            If ID = 7 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Purchase.Visible = True
'                Else
'                    mdifrmmain.Purchase.Visible = False
'                End If
'            End If
'
'            If ID = 8 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.MarketingMnu.Visible = True
'                Else
'                    mdifrmmain.MarketingMnu.Visible = False
'                End If
'            End If
'
'            If ID = 9 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Sales.Visible = True
'                Else
'                    mdifrmmain.Sales.Visible = False
'                End If
'            End If
'
'            If ID = 10 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.shipmentMnu.Visible = True
'                Else
'                    mdifrmmain.shipmentMnu.Visible = False
'                End If
'            End If
'            If ID = 11 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.POSTRansactiosG.Visible = True
'                Else
'                    mdifrmmain.POSTRansactiosG.Visible = False
'                End If
'            End If
'
'            If ID = 12 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.prdo.Visible = True
'                Else
'                    mdifrmmain.prdo.Visible = False
'                End If
'
'            End If
'            If ID = 13 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.ProductionPlan.Visible = True
'                Else
'                    mdifrmmain.ProductionPlan.Visible = False
'                End If
'            End If
'
'            If ID = 14 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.MnuProjects.Visible = True
'                Else
'                    mdifrmmain.MnuProjects.Visible = False
'                End If
'            End If
'            If ID = 15 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.TransporterMain.Visible = True
'                Else
'                    mdifrmmain.TransporterMain.Visible = False
'                End If
'
'            End If
'            If ID = 16 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.CarMaintenance.Visible = True
'                Else
'                    mdifrmmain.CarMaintenance.Visible = False
'                End If
'            End If
'
'            If ID = 17 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.MnuMaintnance.Visible = True
'                Else
'                    mdifrmmain.MnuMaintnance.Visible = False
'                End If
'            End If
'
'            If ID = 18 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Strategy.Visible = True
'                Else
'                    mdifrmmain.Strategy.Visible = False
'                End If
'            End If
'            If ID = 19 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Archiving.Visible = True
'                Else
'                    mdifrmmain.Archiving.Visible = False
'                End If
'            End If
'
'            If ID = 20 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.StudentMenue.Visible = True
'                Else
'                    mdifrmmain.StudentMenue.Visible = False
'                End If
'            End If
'            If ID = 21 Then
'                If Pid = i * i + code Then
'                    ' mdifrmmain.ArrowsBase.Visible = True
'                Else
'                    'mdifrmmain.ArrowsBase.Visible = False
'                End If
'            End If
'
'            If ID = 22 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.AssetsMngBase.Visible = True
'                Else
'                    mdifrmmain.AssetsMngBase.Visible = False
'                End If
'            End If
'            If ID = 23 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Reports.Visible = True
'                Else
'                    mdifrmmain.Reports.Visible = False
'                End If
'
'            End If
'
'            If ID = 24 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Tools.Visible = True
'                Else
'
'                    mdifrmmain.Tools.Visible = False
'                End If
'
'            End If
'
'            If ID = 25 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Basicdata.Visible = True
'                Else
'                    mdifrmmain.Basicdata.Visible = False
'                End If
'            End If
'
'            '               If id = 26 And Pid = I * I + code Then
'            ' mdifrmmain.Tech.Visible = True
'            ' Else
'            '  mdifrmmain.Tech.Visible = False
'            ' End If
'
'            If ID = 27 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.dev.Visible = True
'                Else
'                    mdifrmmain.dev.Visible = False
'                End If
'
'            End If
'
'            '                If ID = 28 Then
'            '        If Pid = i * i + code Then
'            '           '     mdifrmmain.MnueHouseMain.Visible = True
'            '                Else
'            '                 'mdifrmmain.MnueHouseMain.Visible = False
'            '     End If
'            '       End If
'
'            If ID = 28 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Container.Visible = True
'                Else
'                    mdifrmmain.Container.Visible = False
'                End If
'            End If
'
'            If ID = 29 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.COLLECTIONS.Visible = True
'                Else
'                    mdifrmmain.COLLECTIONS.Visible = False
'                End If
'            End If
'            'End If
'
'            If ID = 30 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.CeramicEstimation.Visible = True
'                Else
'                    mdifrmmain.CeramicEstimation.Visible = False
'                End If
'            End If
'
'            If ID = 31 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.RealEstateMarketing.Visible = True
'                Else
'                    mdifrmmain.RealEstateMarketing.Visible = False
'                End If
'            End If
'
'            If ID = 32 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.BankOp.Visible = True
'                Else
'                    mdifrmmain.BankOp.Visible = False
'                End If
'            End If
'
'            If ID = 33 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.mangDep.Visible = True
'                Else
'                    mdifrmmain.mangDep.Visible = False
'                End If
'            End If
'
'            If ID = 34 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.rsInvestment.Visible = True
'                Else
'                    mdifrmmain.rsInvestment.Visible = False
'                End If
'            End If
'
'            If ID = 35 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.SalesIns.Visible = True
'                Else
'                    mdifrmmain.SalesIns.Visible = False
'                End If
'            End If
'
'            If ID = 36 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.MnuElevators.Visible = True
'                Else
'                    mdifrmmain.MnuElevators.Visible = False
'                End If
'            End If
'
'            If ID = 37 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.hajMnu.Visible = True
'                Else
'                    mdifrmmain.hajMnu.Visible = False
'                End If
'            End If
'
'            If ID = 38 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.LIFEINDICATORMNU.Visible = True
'                Else
'                    mdifrmmain.LIFEINDICATORMNU.Visible = False
'                End If
'            End If
'
'            If ID = 39 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.AgeingMAster.Visible = True
'                Else
'                    mdifrmmain.AgeingMAster.Visible = False
'                End If
'            End If
'
'            If ID = 40 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.taxes.Visible = True
'                Else
'                    mdifrmmain.taxes.Visible = False
'                End If
'            End If
'
'            If ID = 41 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.LegalIssue.Visible = True
'                Else
'                    mdifrmmain.LegalIssue.Visible = False
'                End If
'            End If
'
'            If ID = 42 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Tailor.Visible = True
'                Else
'                    mdifrmmain.Tailor.Visible = False
'                End If
'            End If
'
'            If ID = 43 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.rentcar.Visible = True
'                Else
'                    mdifrmmain.rentcar.Visible = False
'                End If
'            End If
'
'            If ID = 44 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Beauty.Visible = True
'                Else
'                    mdifrmmain.Beauty.Visible = False
'                End If
'            End If
'
'            If ID = 45 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.eye.Visible = True
'                Else
'                    mdifrmmain.eye.Visible = False
'                End If
'            End If
'
'            If ID = 46 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.gobus.Visible = True
'                Else
'                    mdifrmmain.gobus.Visible = False
'                End If
'            End If
'
'            If ID = 47 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.xyz.Visible = True
'                Else
'                    mdifrmmain.xyz.Visible = False
'                End If
'            End If
'
'            If ID = 48 Then
'                If Pid = i * i + code Then
'                    mdifrmmain.Farm.Visible = True
'                Else
'                    mdifrmmain.Farm.Visible = False
'                End If
'            End If
'
'            'mdifrmmain.rentcar.Visible =False
'
'            'LIFEINDICATORMNU
'
'll:
'
'            rs.MoveNext
'
'        Next i
'    End If
'
'    mdifrmmain.tech.Visible = True
'
'    rs.Close
'Lite:
'End Function
'
'Public Function GetBrancheName(branch_id As Integer) As String
'    Dim StrSQL As String
'    Dim rs     As New ADODB.Recordset
'
'    StrSQL = "SELECT *  From TblBranchesData where branch_id=" & branch_id
'
'    '    StrSQL = "SELECT branch_id,branch_namee From TblBranchesData"
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        GetBrancheName = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
'
'    Else
'        GetBrancheName = ""
'    End If
'
'End Function
'
'Public Function GetIqarCode(Optional EmpCode As String, _
'                            Optional ByRef Emp_id As Variant, _
'                            Optional Emp_id1 As Double = 0, _
'                            Optional ByRef EmpCode1 As String, _
'                            Optional ByRef ownerid As Variant)
'
'    Dim sql     As String
'    Dim rs      As New ADODB.Recordset
'    Dim Balance As Double
'
'    If Emp_id1 <> 0 Then
'        sql = "select * from TblAqar where Aqarid= " & Emp_id1
'    Else
'
'        sql = "select * from TblAqar where  aqarNo ='" & EmpCode & "'"
'    End If
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        Emp_id = IIf(IsNull(rs("Aqarid").value), 0, rs("Aqarid").value)
'        EmpCode1 = IIf(IsNull(rs("aqarNo").value), 0, rs("aqarNo").value)
'        ownerid = IIf(IsNull(rs("ownerid").value), 0, rs("ownerid").value)
'
'    Else
'        Emp_id = 0
'    End If
'
'    rs.Close
'
'End Function
'
'Public Function GetIqarUnitData(ID As Integer, _
'                                Optional ByRef unitno As String, _
'                                Optional ByRef meterPrice As Double, _
'                                Optional ByRef Length As Double, _
'                                Optional ByRef customerid As Integer, _
'                                Optional ByRef rentType As Integer, _
'                                Optional ByRef roomscount As Double, _
'                                Optional ByRef LoungeCount As Double, _
'                                Optional ByRef WCcount As Double, _
'                                Optional ByRef account As Double, _
'                                Optional ByRef kithchencount As Double, _
'                                Optional ByRef ElectAccount As String, _
'                                Optional MiniRentValue As Double, _
'                                Optional ByRef Typed As Integer) As String
'
'    Dim sql     As String
'    Dim rs      As New ADODB.Recordset
'    Dim Balance As Double
'
'    sql = "SELECT     dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblAqarDetai.Id, dbo.TblAqarDetai.Aqarid, dbo.TblAqarDetai.length, dbo.TblAqarDetai.unitdesc, "
'    sql = sql & "                    dbo.TblAqarDetai.Typed, dbo.TblAqarDetai.unitno, dbo.TblAqarDetai.unittype, dbo.TblAqarDetai.rentType, dbo.TblAqarDetai.meterPrice, dbo.TblAqarDetai.roomscount,"
'    sql = sql & "                      dbo.TblAqarDetai.WCcount, dbo.TblAqarDetai.kithchencount, dbo.TblAqarDetai.RentValue, dbo.TblAqarDetai.customerid, dbo.TblAqarDetai.haveFurniture,"
'    sql = sql & "                     dbo.TblAqarDetai.namerentType, dbo.TblAqarDetai.Floor, dbo.TblAqarDetai.LoungeCount, dbo.TblAqarDetai.ACCount, dbo.TblAqarDetai.ACCountspleat,"
'    sql = sql & "                      dbo.TblAqarDetai.UnitElectric , dbo.TblAqarDetai.electric, dbo.TblAqarDetai.Water, dbo.TblAqarDetai.Services, dbo.TblAqarDetai.status , dbo.TblAqarDetai.MiniRentValue"
'    sql = sql & " FROM         dbo.TblAqarDetai LEFT OUTER JOIN"
'    sql = sql & "                      dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id"
'    sql = sql & "  WHERE     (dbo.TblAqarDetai.Id = " & ID & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        Typed = IIf(IsNull(rs("Typed").value), 1, rs("Typed").value) - 1
'        unitno = IIf(IsNull(rs("unitno").value), "", rs("unitno").value)
'        If SystemOptions.UserInterface = ArabicInterface Then
'            UnittypeName = IIf(IsNull(rs("name").value), "", rs("name").value)
'        Else
'            UnittypeName = IIf(IsNull(rs("namee").value), "", rs("namee").value)
'        End If
'        Length = IIf(IsNull(rs("length").value), 0, val(rs("length").value))
'        rentType = IIf(IsNull(rs("rentType").value), 0, rs("rentType").value)
'        If rentType = 0 Then
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'                rentTypeName = " «·ﬁÌ„… «·«ÌÃ«—Ì…"
'            Else
'                rentTypeName = "By Unit"
'            End If
'        Else
'            If SystemOptions.UserInterface = ArabicInterface Then
'                rentTypeName = " »«·„ — "
'            Else
'                rentTypeName = "By Meter"
'            End If
'        End If
'        MiniRentValue = IIf(IsNull(rs("MiniRentValue").value), 0, rs("MiniRentValue").value)
'        ElectAccount = IIf(IsNull(rs("UnitElectric").value), "", rs("UnitElectric").value)
'        meterPrice = IIf(IsNull(rs("meterPrice").value), 0, rs("meterPrice").value)
'        roomscount = IIf(IsNull(rs("roomscount").value), 0, rs("roomscount").value)
'        LoungeCount = IIf(IsNull(rs("LoungeCount").value), 0, rs("LoungeCount").value)
'        WCcount = IIf(IsNull(rs("WCcount").value), 0, rs("WCcount").value)
'        account = IIf(IsNull(rs("ACCount").value), 0, rs("ACCount").value)
'        kithchencount = IIf(IsNull(rs("kithchencount").value), 0, rs("kithchencount").value)
'        Length = IIf(IsNull(rs("length").value), 0, val(rs("length").value))
'
'        GetIqarUnitData = ""
'        If SystemOptions.UserInterface = ArabicInterface Then
'
'            If Length <> 0 Then
'                GetIqarUnitData = "«·„”«Õ… " & Length & "  „ —" & vbNewLine
'            End If
'
'            If roomscount <> 0 Then
'                GetIqarUnitData = GetIqarUnitData & roomscount & "  €—›… " & vbNewLine
'            End If
'
'            If LoungeCount <> 0 Then
'                GetIqarUnitData = GetIqarUnitData & LoungeCount & "’«·Â" & vbNewLine
'            End If
'
'            If WCcount <> 0 Then
'
'                GetIqarUnitData = GetIqarUnitData & WCcount & "Õ„«„" & vbNewLine
'            End If
'
'            If account <> 0 Then
'                GetIqarUnitData = GetIqarUnitData & account & „ﬂ»› & vbNewLine
'
'            End If
'
'        Else
'
'        End If
'    End If
'    rs.Close
'
'End Function
'
'Public Function GetTblCustemersCode(Optional EmpCode As String, _
'                                    Optional ByRef Emp_id As Variant, _
'                                    Optional Emp_id1 As Variant = 0, _
'                                    Optional ByRef EmpCode1 As String, _
'                                    Optional Type1 As Integer = 1, _
'                                    Optional BranchID As Integer = 0)
'
'    Dim sql     As String
'    Dim rs      As New ADODB.Recordset
'    Dim Balance As Double
'
'    If Emp_id1 <> 0 Then
'        sql = "select * from TblCustemers where CusID= " & Emp_id1
'    Else
'        sql = "select * from TblCustemers where  Fullcode ='" & EmpCode & "'"
'        If Type1 = 2 Then
'            sql = sql & " AND Type  = 2 "
'        ElseIf Type1 = 1 Then
'            sql = sql & " AND Type  = 1 "
'        ElseIf Type1 = 56 Then
'            sql = sql & " AND Type  = 56 "
'        ElseIf Type1 = 57 Then
'            sql = sql & " AND Type  = 57 "
'        End If
'    End If
'
'    If SystemOptions.usertype <> UserAdminAll Then
'        '     StrSQL = StrSQL & " and   ( BranchId=0 or BranchId=" & Current_branch & ")  "
'        sql = sql & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"
'
'    End If
'    If BranchID <> 0 Then
'        '   StrSQL = StrSQL & " and   BranchId=" & BranchID
'        sql = sql & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"
'
'    End If
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'
'        Emp_id = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
'        EmpCode1 = IIf(IsNull(rs("Fullcode").value), 0, rs("Fullcode").value)
'
'    Else
'        Emp_id = 0
'    End If
'
'    rs.Close
'
'End Function
'
'Public Sub RetrivePoNo(Optional order_no As String = "", _
'                       Optional ByRef PONo As String, _
'                       Optional ByRef oorderdate As Date, _
'                       Optional ByRef CBoBasedON As Integer)
'
'    Dim StrSQL As String
'
'    Dim rs     As ADODB.Recordset
'
'    'On Error GoTo ErrTrap
'    StrSQL = "Select * from transactions  where    NoteSerial1='" & order_no & "'"
'    Set rs = New ADODB.Recordset
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'        PONo = IIf(IsNull(rs("PONo").value), "", rs("PONo").value)
'        CBoBasedON = val(IIf(IsNull(rs("CBoBasedON").value), 0, rs("CBoBasedON").value))
'        oorderdate = IIf(IsNull(rs("oorderdate").value), Date, rs("oorderdate").value)
'    Else
'        CBoBasedON = 0
'    End If
'End Sub
'
'Public Function CheckNoteAdvancedPayments(NoteID As Double, _
'                                          Optional ByRef CusID As Long) As Boolean
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = "SELECT   *"
'    sql = sql & " from dbo.notes"
'    sql = sql & " WHERE     (NoteID = " & NoteID & ")"
'
'    Dim NCashingType As Integer
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        NCashingType = IIf(IsNull(rs("NCashingType").value), 0, rs("NCashingType").value)
'        CusID = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
'        If NCashingType = 3 Then
'            CheckNoteAdvancedPayments = True
'            Exit Function
'        Else
'            CheckNoteAdvancedPayments = False
'        End If
'
'    Else
'        CheckNoteAdvancedPayments = False
'        CusID = 0
'    End If
'
'End Function
'
'Public Function GETNationality(ID As Integer) As String
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = "SELECT     NCODE, id"
'    sql = sql & " from dbo.Nationality"
'    sql = sql & " WHERE     (id = " & ID & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        GETNationality = IIf(IsNull(rs("NCODE").value), "", rs("NCODE").value)
'    Else
'        GETNationality = ""
'    End If
'
'End Function
'
'Public Function GETlASTiSSUEDATE(Emp_id As Integer) As Date
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = "SELECT     MAX(todate) AS MaxDate from dbo.TblEmpHolidaysDetails WHERE     (Emp_ID = " & Emp_id & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        GETlASTiSSUEDATE = IIf(IsNull(rs("MaxDate").value), Date, rs("MaxDate").value)
'    Else
'        GETlASTiSSUEDATE = Date
'    End If
'
'End Function
'Public Function calcenaddate(StartDate As Date, _
'                             interval As Integer, _
'                             intervalvalindex As Integer) As Date
'    Dim intervalchar As String
'    If intervalvalindex = 0 Then
'        intervalchar = "M"
'    ElseIf intervalvalindex = 1 Then
'        intervalchar = "YYYY"
'    ElseIf intervalvalindex = 2 Then
'        intervalchar = "D"
'
'    Else
'        intervalchar = "YYYY"
'    End If
'
'    calcenaddate = DateAdd(intervalchar, interval, StartDate)
'
'End Function
'
'Public Function CREATE_VOUCHER_GE(general_noteid As Long, _
'                                  BranchID As Integer, _
'                                  UserID As Long, _
'                                  Notevalue As Double, _
'                                  DebitAccount As String, _
'                                  CreditAcc As String, _
'                                  des As String, _
'                                  NoteDate As Date, _
'                                  Optional debitvatacc As String, _
'                                  Optional Creditvatacc As String, _
'                                  Optional VATValue As Double)
'
'    Dim LngDevID             As Long
'    Dim LngDevNO             As Integer
'    Dim StrTempAccountCode   As String
'    Dim StrTempDes           As String
'    Dim SngTemp              As Variant
'    Dim Account_Code_dynamic As String
'    Dim i                    As Integer
'
'    Dim StrSQL               As String
'
'    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
'    Cn.Execute StrSQL, , adExecuteNoRecords
'
'    LngDevNO = 0
'
'    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'    '«·ÿ—› «·„Ì‰
'
'    my_branch = BranchID
'
'    StrTempAccountCode = DebitAccount
'
'    ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
'    StrTempDes = "”‰œ   " & des & "   " & TxtNoteSerial1V
'    LngDevNO = LngDevNO + 1
'
'    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
'        GoTo ErrTrap
'    End If
'
'    If debitvatacc <> "" And VATValue > 0 Then
'
'        StrTempAccountCode = debitvatacc
'        LngDevNO = LngDevNO + 1
'        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, VATValue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
'            GoTo ErrTrap
'        End If
'
'    End If
'
'    StrTempAccountCode = CreditAcc
'    LngDevNO = LngDevNO + 1
'    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue + VATValue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
'        GoTo ErrTrap
'    End If
'
'    If Creditvatacc <> "" And VATValue > 0 Then
'
'        StrTempAccountCode = Creditvatacc
'        LngDevNO = LngDevNO + 1
'        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, VATValue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
'            GoTo ErrTrap
'        End If
'
'    End If
'
'ErrTrap:
'End Function
'
'Public Function showsforms(Index As Integer)
'    Select Case Index
'        Case 0
'            Load FrmCarAuthontication
'            FrmCarAuthontication.show
'
'        Case 2
'            If checkApility("FrmOut") = False Then
'                Exit Function
'            End If
'
'            FrmOut.show
'            FrmOut.TxtTicketNO.Visible = True
'            FrmOut.lbl(32).Visible = True
'        Case 3
'            Load FrmBillCarMaintExtra
'            FrmBillCarMaintExtra.show
'        Case 4
'            Load FrmCarReporonlin
'            FrmCarReporonlin.show
'        Case 5
'            Load FrmCarReportsRequerNo
'            FrmCarReportsRequerNo.show
'        Case 6
'            Load FrmBillComputerChek
'            FrmBillComputerChek.show
'        Case 7
'            Load FrmOrderOpen
'            FrmOrderOpen.show
'        Case 8
'            Load FrmCarReporonlin2
'            FrmCarReporonlin2.show
'        Case 9
'
'            If SystemOptions.ShowBillCommisions = 0 Then
'                Exit Function
'            End If
'            Load FrmCommisRece
'            FrmCommisRece.show
'
'        Case 10
'
'            Load FrmCustemers
'            FrmCustemers.show
'        Case 11
'            If SystemOptions.ShowBillCommisions = 0 Then
'                Exit Function
'            End If
'
'            Load FrmCommisReport
'            FrmCommisReport.show
'        Case 12
'            Load FrmCustemers
'            FrmCustemers.show
'    End Select
'
'End Function
'
'Public Function SetPrinter2(PrnName As String)
'    Dim Prn As Printer
'    If Printers.count > 0 Then
'        For Each Prn In Printers
'            If Prn.DeviceName = PrnName Then
'                Set Printer = Prn
'                Exit For
'            End If
'        Next Prn
'    End If
'End Function
'
'Public Function createCustomer(CusName As String, _
'                               CusNamee As String, _
'                               Optional BranchID As Integer = 0, _
'                               Optional ByRef CusID As Double, _
'                               Optional ByVal Cus_mobile As String = "", _
'                               Optional ByRef mCode As String = "") As Integer
'
'    Dim RsTemp      As New ADODB.Recordset
'    Dim currentcode As String
'    Dim s           As String, mPreFix As String
'
'    StrSQL = "Select * From TblCustemers where CusName='" & CusNamee & "'"
'    StrSQL = StrSQL & " or CusNamee='" & CusNamee & "'"
'
'    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If RsTemp.RecordCount > 0 Then
'
'        createCustomer = 0
'        CusID = IIf(IsNull(RsTemp("CusID").value), 0, RsTemp("CusID").value)
'        Exit Function
'    End If
'
'    Dim ParentAccount        As String
'    Dim parent_account       As String
'    Dim Account_Code_dynamic As String
'    Account_Code_dynamic = get_account_code_branch(8, 1)
'
'    If Account_Code_dynamic = "NO branch" Then
'        MsgBox " ·« ÌÊÃœ —»ÿ Õ”«»«  ", vbCritical
'        createCustomer = -1
'        Exit Function
'    Else
'
'        If Account_Code_dynamic = "NO account" Then
'            MsgBox "  ·« ÌÊÃœ —»ÿ Õ”«»«  ", vbCritical
'            createCustomer = -1
'            Exit Function
'        End If
'    End If
'    parent_account = Account_Code_dynamic
'
'    CusID = CStr(new_id("TblCustemers", "CusID", "", True))
'
'    Dim Account_Code  As String
'    Dim Account_code1 As String
'    Dim Account_code2 As String
'
'    ParentAccount = ""
'    Account_code1 = ""
'    Account_code2 = ""
'
'    If SystemOptions.CustomerhavethreeAccounts = False Then
'        Account_Code = ModAccounts.AddNewAccount(parent_account, CusName, True, False, CusNamee)
'
'        '        parent_account
'
'    Else
'
'        If SystemOptions.CustomerhavethreeAccounts = True Then
'            ParentAccount = ModAccounts.AddNewAccount(parent_account, CusName, False, False, CusNamee)
'            'rs("ParentAccount").value = ParentAccount
'
'            Account_Code = ModAccounts.AddNewAccount(ParentAccount, CusName, True, False, CusNamee)
'            Account_code1 = ModAccounts.AddNewAccount(ParentAccount, CusName & "   ‘Ìﬂ«   Õ  «· Õ’Ì· ", True, False, CusNamee & "  Under Collection Cheque  ")
'            Account_code2 = ModAccounts.AddNewAccount(ParentAccount, CusName & "   „œ›Ê⁄«  „ﬁœ„…  ", True, False, CusName & " Advanced Payments")
'
'        Else
'            Account_Code = ModAccounts.AddNewAccount(Account_Code_dynamic, CusName, True, False, CusNamee)
'            '  rs("ParentAccount").value = Null
'
'        End If
'
'    End If
'
'    If CStr(CusID) <> "" Then
'        s = " SELECT Top 1  FIELD_no,prifix From Coding WHERE  FIELD_no = 4 and IsNull(prifix,'') <> ''"
'        Dim rsDummy As New ADODB.Recordset
'        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
'        If Not rsDummy.EOF Then
'            mPreFix = rsDummy!prifix & ""
'        End If
'        currentcode = get_coding(Current_branch, "TblCustemers", 4, mPreFix)
'        mCode = mPreFix & CusID
'    End If
'
'    StrSQL = " insert into TblCustemers  (type,CusID,CusName,CusNameE,code,branchID,parent_account,Account_Code,Account_Code1,Account_Code2,ParentAccount,Cus_mobile,FullCode) "
'    StrSQL = StrSQL & " VALUES (1," & CusID & ",'" & CusName & "' , '" & CusNamee & "' ,  '" & mCode & "'," & BranchID & ",'" & parent_account & "','" & Account_Code & "','" & Account_code1 & "','" & Account_code2 & "','" & ParentAccount & "','" & Cus_mobile & "' ,'" & mCode & "')"
'
'    Cn.Execute StrSQL
'
'End Function
'
'Public Function AutoSel(Cmb As ComboBox, KeyCode As Integer)
'
'    Debug.Print KeyCode
'
'    If KeyCode = vbEnter Then
'        Exit Function
'    End If
'    If KeyCode = 8 Then
'        Exit Function    'Backspace
'    End If
'    If KeyCode = 37 Then
'        Exit Function  'left key
'    End If
'    If KeyCode = 38 Then
'        Exit Function 'up arrow key
'    End If
'    If KeyCode = 39 Then
'        Exit Function  'right key
'    End If
'    If KeyCode = 40 Then
'        Exit Function  'down arrow key
'    End If
'    If KeyCode = 46 Then
'        Exit Function  'delete key
'    End If
'    If KeyCode = 33 Then
'        Exit Function  'page up key
'    End If
'    If KeyCode = 34 Then
'        Exit Function  'page down key
'    End If
'    If KeyCode = 35 Then
'        Exit Function  'end key
'    End If
'    If KeyCode = 36 Then
'        Exit Function  'home key
'    End If
'
'    Dim Text As String
'    Text = Cmb.Text
'
'    Dim i    As Long
'    Dim temp As String
'
'    For i = 0 To Cmb.ListCount - 1
'        temp = left(Cmb.List(i), Len(Text))
'        If LCase(temp) = LCase(Text) Then
'            Cmb.Text = Cmb.List(i)
'            Cmb.ListIndex = i
'            Cmb.SelStart = Len(Text)
'            Cmb.SelLength = Len(Cmb.List(i))
'            'Cmb.SetFocus
'        End If
'    Next
'
'End Function
'Public Function GetWeekdayName(DayNO As Integer) As String
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'
'        Select Case DayNO
'
'            Case 1
'                GetWeekdayName = "«·”» "
'
'            Case 2
'                GetWeekdayName = "«·«Õœ"
'
'            Case 3
'                GetWeekdayName = "«·«À‰Ì‰"
'
'            Case 4
'                GetWeekdayName = "«·À·«À«¡"
'
'            Case 5
'                GetWeekdayName = "«·«—»⁄«¡"
'
'            Case 6
'                GetWeekdayName = "«·Œ„Ì”"
'
'            Case 7
'                GetWeekdayName = "«·Ã„⁄Â"
'
'        End Select
'
'    Else
'
'        Select Case DayNO
'
'            Case 1
'                GetWeekdayName = "Saturday"
'
'            Case 2
'                GetWeekdayName = "Sunday"
'
'            Case 3
'                GetWeekdayName = "Monday"
'
'            Case 4
'                GetWeekdayName = "Tuesday"
'
'            Case 5
'                GetWeekdayName = "Wednesday"
'
'            Case 6
'                GetWeekdayName = "Thursday"
'
'            Case 7
'                GetWeekdayName = "Friday"
'
'        End Select
'
'    End If
'
'End Function
'
'Function MoveUpDown(ByRef List As ListBox, upDown As Integer)
'    Dim currentpos    As Integer
'    Dim currentname   As String
'    Dim BEFOREPOSTION As Integer
'    Dim BEFORENAME    As String
'
'    Dim AfterPOSTION  As Integer
'    Dim AfterNAME     As String
'
'    If upDown = 0 Then 'up
'        currentpos = List.ListIndex
'        currentname = List.List(currentpos)
'
'        If currentpos = 0 Then
'            Exit Function
'        End If
'
'        BEFOREPOSTION = List.ListIndex - 1
'        BEFORENAME = List.List(BEFOREPOSTION)
'
'        List.List(BEFOREPOSTION) = currentname
'
'        List.List(currentpos) = BEFORENAME
'        List.ListIndex = BEFOREPOSTION
'    Else
'
'        currentpos = List.ListIndex
'        currentname = List.List(currentpos)
'
'        If currentpos = List.ListCount - 1 Then
'            Exit Function
'        End If
'
'        AfterPOSTION = List.ListIndex + 1
'        AfterNAME = List.List(AfterPOSTION)
'
'        List.List(AfterPOSTION) = currentname
'
'        List.List(currentpos) = AfterNAME
'
'        List.ListIndex = AfterPOSTION
'
'    End If
'
'End Function
'
'Public Function saveApprovalData(Transactionid As Double, _
'                                 Transaction_Type As Double, _
'                                 NoteSerial As Double, _
'                                 frmname As String)
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'    Dim i   As Integer
'
'    sql = "SELECT     dbo.TblApprovalDefDetails.PlainMessageID, dbo.TbLLevels.Name, dbo.TbLLevels.Namee, dbo.TbllevelWorker.EmpID"
'    sql = sql & "  FROM         dbo.TbllevelWorker INNER JOIN"
'    sql = sql & "  dbo.TbLLevels ON dbo.TbllevelWorker.LevelID = dbo.TbLLevels.LevelID INNER JOIN"
'    sql = sql & "  dbo.TblApprovalDef INNER JOIN"
'    sql = sql & "  dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID ON"
'    sql = sql & "  dbo.TbLLevels.LevelID = dbo.TblApprovalDefDetails.PlainMessageID"
'    sql = sql & "  WHERE     (dbo.TblApprovalDef.ScreenName = N'" & frmname & "')"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    Cn.Execute "delete TblTransactionsApproval where Transaction_Type=" & Transaction_Type & " and Transactionid=" & Transactionid
'
'    If rs.RecordCount > 0 Then
'
'        For i = 1 To rs.RecordCount
'
'            If Not (IsNull(rs("PlainMessageID").value)) Then
'                sql = "insert into TblTransactionsApproval (Transaction_Type,NoteSerial,Level,Transactionid,CurrUserID,UserID)  "
'                sql = sql & "Values (" & Transaction_Type & "," & NoteSerial & "," & rs("PlainMessageID").value & "," & Transactionid & "," & user_id & "," & rs("empID").value & ")"
'                Cn.Execute sql
'            End If
'
'            rs.MoveNext
'        Next i
'
'    End If
'
'End Function
'
'Public Function getInfoMessage1(ID As Integer, _
'                                Optional ByRef Name As String, _
'                                Optional ByRef speed As Double, _
'                                Optional ByRef fontsize As Double, _
'                                Optional ByRef fontcolor As Double, _
'                                Optional ByRef backcolor As Double, _
'                                Optional ByRef show As Boolean)
'
'    On Error Resume Next
'
'    Dim SQL1 As String
'    Dim Rs1  As New ADODB.Recordset
'
'    SQL1 = " SELECT   * from InfoSettings1   "
'    SQL1 = SQL1 + "  WHERE     (" & SQLDate(Date, True) & " BETWEEN dbo.InfoSettings1.StartDate AND dbo.InfoSettings1.EndDate)  "
'
'    '        Sql1 = Sql1 + " where  (startdate >=" & SQLDate(Date, True) & ""
'
'    '   Sql1 = Sql1 + " and enddate <=" & SQLDate(Date, True) & ""
'    '    Sql1 = Sql1 + "and CAST(StartTime As Time) >= CAST(CURDATE()() As Time) "
'    '  Sql1 = Sql1 + "and  CAST(enddate As Time) <= CAST(CURDATE()() As Time) "
'    Rs1.Open SQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Rs1.RecordCount > 0 Then
'        If 1 = 1 Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Name = IIf(IsNull(Rs1("name").value), "", Rs1("name").value)
'            Else
'                Name = IIf(IsNull(Rs1("namee").value), "", Rs1("namee").value)
'            End If
'
'            speed = IIf(IsNull(Rs1("speed").value), 50, Rs1("speed").value)
'            fontsize = IIf(IsNull(Rs1("fontsize").value), 12, Rs1("fontsize").value)
'            fontcolor = IIf(IsNull(Rs1("fontcolor").value), 255, Rs1("fontcolor").value)
'            backcolor = IIf(IsNull(Rs1("backcolor").value), 0, Rs1("backcolor").value)
'            show = True
'            WebForm.info1Timer.interval = speed
'        Else
'            show = False
'        End If
'    Else
'        show = False
'    End If
'
'End Function
'Public Function getInfoMessage(ID As Integer, _
'                               Optional ByRef Name As String, _
'                               Optional ByRef speed As Double, _
'                               Optional ByRef fontsize As Double, _
'                               Optional ByRef fontcolor As Double, _
'                               Optional ByRef backcolor As Double, _
'                               Optional ByRef show As Boolean)
'
'    On Error Resume Next
'
'    Dim SQL1 As String
'    Dim Rs1  As New ADODB.Recordset
'    SQL1 = " SELECT   * from InfoSettings1   "
'
'    SQL1 = SQL1 + " where  (startdate >=" & SQLDate(Date, True) & ""
'
'    SQL1 = SQL1 + " and enddate <=" & SQLDate(Date, True) & ""
'    SQL1 = SQL1 + "and CAST(StartTime As Time) >= CAST(CURDATE()() As Time) "
'    SQL1 = SQL1 + "and  CAST(enddate As Time) <= CAST(CURDATE()() As Time) "
'
'    '   rs1.Open sql1, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    '    If rs1.RecordCount > 0 Then
'
'    '   End If
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'    sql = " SELECT   * from InfoSettings  "
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.RecordCount > 0 Then
'        If rs("Show").value = True Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Name = IIf(IsNull(rs("name").value), "", rs("name").value)
'            Else
'                Name = IIf(IsNull(rs("namee").value), "", rs("namee").value)
'            End If
'
'            speed = IIf(IsNull(rs("speed").value), 50, rs("speed").value)
'            fontsize = IIf(IsNull(rs("fontsize").value), 12, rs("fontsize").value)
'            fontcolor = IIf(IsNull(rs("fontcolor").value), 255, rs("fontcolor").value)
'            backcolor = IIf(IsNull(rs("backcolor").value), 0, rs("backcolor").value)
'            show = True
'            WebForm.Timer2.interval = speed
'        Else
'            show = False
'        End If
'    Else
'        show = False
'    End If
'
'End Function
'
'Public Function AddTofaforites(Optional formname As String, _
'                               Optional Displayname As String, _
'                               Optional Displaynamee As String)
'    'On Error Resume Next
'    Dim sql        As String
'    Dim rs         As New ADODB.Recordset
'    Dim Noofmenues As Double
'    'Dim TimeCateg As Double
'    Dim str        As String
'    If formname = "" Then
'        Exit Function
'    End If
'    sql = "SELECT     COUNT(id) AS Noofmenues"
'    sql = sql & " from dbo.TblMyMenue"
'    sql = sql & " WHERE      userid=" & user_id
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        Noofmenues = IIf(IsNull(rs("Noofmenues").value), 0, rs("Noofmenues").value)
'
'    Else
'        Noofmenues = 0
'    End If
'
'    rs.Close
'    'CHECK FORM NAME NOT EXIST BEFORE
'    sql = "SELECT     *  "
'    sql = sql & " from dbo.TblMyMenue"
'    sql = sql & " WHERE      userid=" & user_id & " and  formname='" & formname & "'"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            MsgBox "·« Ì„ﬂ‰ «÷«›… Â–… «·‘«‘… „ÊÃÊœ… »«·›⁄· ›Ì «·„›÷·«  ", vbInformation
'        Else
'            MsgBox "can't Add to Favorites It's Already Exist ", vbInfor, mation
'        End If
'
'        Exit Function
'
'    End If
'
'    If Noofmenues <= 30 Then
'
'        str = "insert into  TblMyMenue   (  USERID,formname,Displayname,Displaynamee) "
'        str = str & "values( " & user_id & ",'" & formname & "','" & Displayname & "','" & Displaynamee & "'  )"
'        Cn.Execute str
'    Else
'        If SystemOptions.UserInterface = ArabicInterface Then
'            MsgBox "·« Ì„ﬂ‰ «÷«›… «Œ—Ì «·Ï «·„›÷·«  ", vbInformation
'        Else
'            MsgBox "can't Add to Favorites ", vbInformation
'        End If
'
'    End If
'    If SystemOptions.UserInterface = ArabicInterface Then
'        MsgBox " „  «·«÷«›… ··„›÷·Â ", vbInformation
'    Else
'        MsgBox "Added to Favorites Success ", vbInformation
'    End If
'
'    Call mdifrmmain.showFavoritesMenue
'End Function
'
'Public Function CheckLastApprovLevel(Optional ScreenName As String, _
'                                     Optional Transaction_ID As Double = 0, _
'                                     Optional NoteID As Double = 0) As Double
'    Dim sql        As String
'    Dim rs         As New ADODB.Recordset
'    Dim NoOfMinute As Double
'    'Dim TimeCateg As Double
'
'    sql = "SELECT     COUNT(id) AS NotApproved"
'    sql = sql & " from dbo.ApprovalData"
'    sql = sql & " WHERE    empid <>0  and   (ScreenName = N'" & ScreenName & "')  AND (ApprovDate IS NULL)"
'    If Transaction_ID <> 0 Then
'        sql = sql & " AND (Transaction_ID = " & Transaction_ID & ")   "
'    End If
'
'    If NoteID <> 0 Then
'        sql = sql & " AND (NoteID = " & NoteID & ")   "
'    End If
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        CheckLastApprovLevel = IIf(IsNull(rs("NotApproved").value), 0, rs("NotApproved").value)
'
'    Else
'        CheckLastApprovLevel = 0
'    End If
'
'End Function
'
'Public Function GetTimeforTransaction(Optional ScreenName As String, _
'                                      Optional ByRef TimeCateg As Double) As Double
'
'    Dim sql        As String
'    Dim rs         As New ADODB.Recordset
'    Dim NoOfMinute As Double
'    'Dim TimeCateg As Double
'
'    sql = "SELECT     ScreenName, timeCount, TimeCateg"
'    sql = sql & " From dbo.TblApprovalDef"
'    sql = sql & " WHERE     (ScreenName = N'" & ScreenName & "') "
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        TimeCateg = IIf(IsNull(rs("TimeCateg").value), 0, rs("TimeCateg").value)
'        NoOfMinute = IIf(IsNull(rs("timeCount").value), 0, rs("timeCount").value)
'        If TimeCateg = 0 Then
'            NoOfMinute = NoOfMinute * 1
'        ElseIf TimeCateg = 1 Then
'            NoOfMinute = NoOfMinute * 60
'        ElseIf TimeCateg = 2 Then
'            NoOfMinute = NoOfMinute * 60 * 24
'        End If
'
'    Else
'        NoOfMinute = 0
'    End If
'    GetTimeforTransaction = NoOfMinute
'
'End Function
'
'Public Function GetlastPurchasedata(Transaction_Type As Double, _
'                                    Item_ID As Double, _
'                                    Fromdate As Date, _
'                                    todate As Date, _
'                                    Optional ByRef LastPurchaseDate As String, _
'                                    Optional ByRef LastPrice As Double, _
'                                    Optional ByRef lastQty As Double)
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = " SELECT     TOP 100 PERCENT MAX(dbo.Transactions.Transaction_Date) AS lastPurchaseDate, dbo.Transaction_Details.showPrice AS lastPrice, "
'    sql = sql & " dbo.Transaction_Details.ShowQty AS lastQty"
'    sql = sql & "  FROM         dbo.Transactions INNER JOIN"
'    sql = sql & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'    sql = sql & "   Where (dbo.Transactions.Transaction_Type = " & Transaction_Type & ") And (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
'    sql = sql & "   GROUP BY dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ShowQty"
'    sql = sql & "   HAVING      (MAX(dbo.Transactions.Transaction_Date) >= " & SQLDate(Fromdate, True) & " AND MAX(dbo.Transactions.Transaction_Date)"
'    sql = sql & "    <= " & SQLDate(todate, True) & ")"
'    sql = sql & "   ORDER BY MAX(dbo.Transactions.Transaction_Date) DESC"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        LastPurchaseDate = IIf(IsNull(rs("lastPurchaseDate").value), "", rs("lastPurchaseDate").value)
'        LastPrice = Round(IIf(IsNull(rs("lastPrice").value), 0, rs("lastPrice").value), 2)
'        lastQty = Round(IIf(IsNull(rs("lastQty").value), 0, rs("lastQty").value), 2)
'
'    Else
'        LastPrice = 0
'        lastQty = 0
'        LastPurchaseDate = ""
'
'    End If
'    rs.Close
'    Set rs = Nothing
'End Function
'Public Function checkmanyStores(Optional ByRef str As String = "") As Boolean
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'    If SystemOptions.UserInterface = ArabicInterface Then
'        sql = "SELECT     dbo.TblUsersStores.StoreID, dbo.TblStore.StoreName "
'    Else
'        sql = "SELECT     dbo.TblUsersStores.StoreID, dbo.TblStore.StoreNamee "
'    End If
'
'    sql = sql & "  FROM         dbo.TblUsersStores LEFT OUTER JOIN"
'    sql = sql & "  dbo.TblStore ON dbo.TblUsersStores.StoreID = dbo.TblStore.StoreID"
'    If user_id <> 1 Then
'        sql = sql & "    Where (dbo.TblUsersStores.userid = " & user_id & ")"
'    Else
'        checkmanyStores = False
'        Exit Function
'    End If
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        str = sql
'        checkmanyStores = True
'
'    Else
'        checkmanyStores = False
'    End If
'
'End Function
'
'Public Function checkmanyBranches(Optional ByRef str As String = "") As Boolean
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'    If SystemOptions.UserInterface = ArabicInterface Then
'        sql = "SELECT     dbo.TblUsersBranches.BranchID, dbo.TblBranchesData.branch_name "
'    Else
'        sql = "SELECT     dbo.TblUsersBranches.BranchID, dbo.TblBranchesData.branch_namee "
'    End If
'
'    sql = sql & "   FROM         dbo.TblUsersBranches INNER JOIN"
'    sql = sql & "   dbo.TblBranchesData ON dbo.TblUsersBranches.BranchID = dbo.TblBranchesData.branch_id"
'    If user_id <> 1 Then
'        sql = sql & "    Where (dbo.TblUsersBranches.UserID = " & user_id & ")"
'    Else
'        checkmanyBranches = False
'        Exit Function
'    End If
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        str = sql
'        checkmanyBranches = True
'
'    Else
'        checkmanyBranches = False
'    End If
'
'End Function
'Public Function GetYearlyAverage(Transaction_Type As Double, _
'                                 Item_ID As Double, _
'                                 Fromdate As Date, _
'                                 todate As Date, _
'                                 Optional ByRef GetYearlyAverage1 As Double)
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    '   sql = " SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.ShowQty) / COUNT(DISTINCT MONTH(dbo.Transactions.Transaction_Date)) AS AverageMonthly "
'    'sql = sql & "   FROM         dbo.Transactions INNER JOIN"
'    'sql = sql & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'    'sql = sql & "    WHERE     (dbo.Transactions.Transaction_Type = " & Transaction_Type & ") AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND (dbo.Transactions.Transaction_Date >=  " & SQLDate(fromdate, True) & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True) & ")"
'
'    sql = " SELECT    SUM(dbo.Transaction_Details.ShowQty) / 1 as YearlyAverage"
'    sql = sql & "    FROM         dbo.Transactions INNER JOIN"
'    sql = sql & "     dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
'    sql = sql & "   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
'    sql = sql & "   WHERE     (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND (dbo.Transactions.Transaction_Date >= " & SQLDate(Fromdate, True) & ") AND (dbo.Transactions.Transaction_Date <=" & SQLDate(todate, True) & ")"
'    sql = sql & "    AND (dbo.TransactionTypes.StockEffect = - 1)"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        GetYearlyAverage1 = Round(IIf(IsNull(rs("YearlyAverage").value), 0, rs("YearlyAverage").value), 2)
'
'    Else
'        GetYearlyAverage1 = 0
'    End If
'
'End Function
'
'Public Function GetMonthlyAverage(Transaction_Type As Double, _
'                                  Item_ID As Double, _
'                                  Fromdate As Date, _
'                                  todate As Date, _
'                                  Optional ByRef AverageMonthly As Double)
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    '   sql = " SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.ShowQty) / COUNT(DISTINCT MONTH(dbo.Transactions.Transaction_Date)) AS AverageMonthly "
'    'sql = sql & "   FROM         dbo.Transactions INNER JOIN"
'    'sql = sql & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'    'sql = sql & "    WHERE     (dbo.Transactions.Transaction_Type = " & Transaction_Type & ") AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND (dbo.Transactions.Transaction_Date >=  " & SQLDate(fromdate, True) & "  AND dbo.Transactions.Transaction_Date <= " & SQLDate(todate, True) & ")"
'
'    sql = " SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.ShowQty) / COUNT(DISTINCT MONTH(dbo.Transactions.Transaction_Date)) AS AverageMonthly"
'    sql = sql & "    FROM         dbo.Transactions INNER JOIN"
'    sql = sql & "     dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
'    sql = sql & "   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
'    sql = sql & "   WHERE     (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND (dbo.Transactions.Transaction_Date >= " & SQLDate(Fromdate, True) & ") AND (dbo.Transactions.Transaction_Date <=" & SQLDate(todate, True) & ")"
'    sql = sql & "    AND (dbo.TransactionTypes.StockEffect = - 1)"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        AverageMonthly = Round(IIf(IsNull(rs("AverageMonthly").value), 0, rs("AverageMonthly").value), 2)
'
'    Else
'        AverageMonthly = 0
'    End If
'
'End Function
'
'Public Function GetempDepartementidFromUserid(UserID As Double) As Double
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = " SELECT     DepartmentID"
'    sql = sql & "  From dbo.TblEmployee"
'    sql = sql & "   Where (Emp_id = " & GetempidFromUserid(UserID) & ")"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        GetempDepartementidFromUserid = IIf(IsNull(rs("DepartmentID").value), 0, rs("DepartmentID").value)
'
'    Else
'        GetempDepartementidFromUserid = 0
'    End If
'
'End Function
'
'Public Function GetempidFromUserid(UserID As Double) As Double
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = "  select * from TblUsers where UserID =" & UserID
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        GetempidFromUserid = IIf(IsNull(rs("Empid").value), 0, rs("Empid").value)
'
'    Else
'        GetempidFromUserid = 0
'    End If
'
'End Function
'
'Public Function GetCurrentApprovalForTransactions(Transaction_ID As Double, _
'                                                  Optional ScreenName As String) As Double
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = " SELECT    MIN(id) AS MinId"
'    sql = sql & " from dbo.ApprovalData"
'    sql = sql & " WHERE  empid <>0 and     (ApprovDate IS NULL and CancelApprove IS NULL) AND (Transaction_ID = " & Transaction_ID & ") AND (ScreenName = N'" & ScreenName & "')"
'    sql = sql & " ORDER BY MIN(id)  "
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        GetCurrentApprovalForTransactions = IIf(IsNull(rs("MinId").value), 0, rs("MinId").value)
'
'    Else
'        GetCurrentApprovalForTransactions = 0
'    End If
'
'End Function
'Public Function getClassInformations(ID As Integer, _
'                                     Optional ByRef Name As String, _
'                                     Optional ByRef DiscountPercentage As Double, _
'                                     Optional ByRef PerfectPercentage As Double, _
'                                     Optional ByRef Account_Code As String)
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = " SELECT   * from TblItemsclasses  "
'
'    sql = sql & " Where (SizeId = " & ID & ")"
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        Name = IIf(IsNull(rs("SizeName").value), "", rs("SizeName").value)
'        Account_Code = IIf(IsNull(rs("Account_code").value), "", rs("Account_code").value)
'        DiscountPercentage = IIf(IsNull(rs("DiscountPercentage").value), 0, rs("DiscountPercentage").value)
'        PerfectPercentage = IIf(IsNull(rs("PerfectPercentage").value), 0, rs("PerfectPercentage").value)
'
'    Else
'        Name = ""
'
'        DiscountPercentage = 0
'    End If
'    getClassInformations = DiscountPercentage
'End Function
'
'Public Function getMaintenancetypeInformations(ID As Integer, _
'                                               Optional ByRef Name As String, _
'                                               Optional ByRef km As String, _
'                                               Optional ByRef Remarks As String, _
'                                               Optional ByRef alarmBfore As Double)
'
'    Dim sql As String
'    Dim rs  As New ADODB.Recordset
'
'    sql = " SELECT   * from MaintenanceTypes  "
'
'    sql = sql & " Where (id = " & ID & ")"
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        Name = IIf(IsNull(rs("name").value), "", rs("name").value)
'        km = IIf(IsNull(rs("km").value), "", rs("km").value)
'        Remarks = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
'        alarmBfore = IIf(IsNull(rs("alarmBfore").value), 0, rs("alarmBfore").value)
'
'    Else
'        Name = ""
'        km = ""
'        Remarks = ""
'        alarmBfore = 0
'    End If
'
'End Function
'
'Public Function CreateLogo(xRport As CRAXDRT.Report, _
'                           Optional BranchID As Double = 0, _
'                           Optional ByVal StrN As String = "") As Boolean
'    Dim rs          As ADODB.Recordset
'    Dim BolShowLogo As Boolean
'    Dim xLogo       As CRAXDRT.OLEObject
'    Dim StrFileName As String
'    Dim MsgErr      As String
'    Dim StrSQL      As String
'    On Error GoTo hErr
'
'    Set rs = New ADODB.Recordset
'    If SystemOptions.WorkWithBranchLogo = False Then
'        rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
'    Else
'
'        If BranchID = 0 Then
'            StrSQL = "SELECT     *  from TblBranchesData Where (branch_id = " & Current_branch & ")"
'        Else
'            StrSQL = "SELECT     *  from TblBranchesData Where (branch_id = " & BranchID & ")"
'
'        End If
'        rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'
'    End If
'
'    If rs.BOF Or rs.EOF Then
'        CreateLogo = False
'        Exit Function
'    End If
'
'    BolShowLogo = IIf(IsNull(rs("ShowLogoInReports").value), 0, rs("ShowLogoInReports").value)
'
'    If BolShowLogo = True And hide_logo = False Then
'        If SystemOptions.WorkWithBranchLogo = False Then
'            LoadPictureFromDB Nothing, rs, "CompanyLogo", StrFileName
'        Else
'            LoadPictureFromDB Nothing, rs, "branchLogo", StrFileName
'
'        End If
'
'        If StrN <> "" Then StrFileName = StrN
'
'        Set xLogo = xRport.Areas(1).Sections(1).AddPictureObject(StrFileName, 100, 100)
'        xLogo.Width = SystemOptions.logowidth
'        xLogo.Height = SystemOptions.logoHeight
'        xLogo.backcolor = vbWhite
'        xLogo.BorderColor = 255
'        xLogo.CloseAtPageBreak = True
'        xLogo.HyperlinkText = "BYTE"
'        xLogo.HyperlinkType = crHyperlinkWebsite
'        xRport.Areas(1).Sections(1).SuppressIfBlank = True
'        xRport.Areas(1).Sections(1).Height = xLogo.Height + 250
'        CreateLogo = True
'    Else
'        CreateLogo = False
'    End If
'
'    rs.Close
'    Set rs = Nothing
'    Exit Function
'hErr:
'    MsgErr = "Œÿ« ›Ï "
'    MsgErr = MsgErr & Chr(13) & "CreateLogo"
'    MsgErr = MsgErr & Chr(13) & Err.Description
'    MsgErr = MsgErr & Chr(13) & Err.Number
'    MsgErr = MsgErr & Chr(13) & Err.Source
'    WriteInLogFile MsgErr
'    CreateLogo = False
'End Function
'
'Public Function getitemAgeingData(Fromdate As Date, _
'                                  todate As Date, _
'                                  Optional GroupID As Integer = 0, _
'                                  Optional Item_ID As Integer)
'    Dim NameOfAgeType        As String
'
'    Dim late_interval        As Integer
'    Dim ItemID               As Long
'    Dim Dean_age             As Integer
'
'    Dim column_location      As Integer
'    Dim column_COLOR         As String
'    Dim customerid           As Integer
'    Dim i                    As Integer
'    Dim sql                  As String
'    Dim DefaultSalesPersonId As Integer
'    Dim Rs3                  As New ADODB.Recordset
'
'    sql = "SELECT     TOP 100 PERCENT MAX(dbo.Transactions.Transaction_Date) AS LastDate, dbo.Transaction_Details.Item_ID"
'    sql = sql & " ,  DATEDIFF(day,MAX(dbo.Transactions.Transaction_Date)"
'    sql = sql & " , " & SQLDate(todate, True) & ") as DIFFerents"
'    sql = sql & " FROM         dbo.Transaction_Details INNER JOIN"
'    sql = sql & "   dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
'    sql = sql & "  INNER JOIN dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
'    sql = sql & " WHERE     (dbo.Transactions.Transaction_Type = 21)"
'    sql = sql & " AND (dbo.Transactions.Transaction_Date >=" & SQLDate(Fromdate, True) & " )"
'    sql = sql & " AND (dbo.Transactions.Transaction_Date <=" & SQLDate(todate, True) & " )"
'
'    If GroupID <> 0 Then
'        sql = sql & " AND (dbo.TblItems.GroupID = " & GroupID & ")"
'    End If
'
'    If Item_ID <> 0 Then
'        sql = sql & " AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
'    End If
'
'    sql = sql & " GROUP BY dbo.Transaction_Details.Item_ID"
'    sql = sql & " ORDER BY dbo.Transaction_Details.Item_ID"
'
'    Dim str        As String
'    Dim Note_Value As Double
'    str = "delete TblTempItemAging"
'
'    Cn.Execute str
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount = 0 Then
'        Exit Function
'    End If
'
'    If Rs3.RecordCount > 0 Then
'
'        Rs3.MoveFirst
'
'        For i = 1 To Rs3.RecordCount
'
'            late_interval = Rs3.Fields("DIFFerents").value
'            ItemID = Rs3.Fields("Item_ID").value
'            column_location = get_late_location2(late_interval)
'            '     column_COLOR = get_late_COLOR(column_location, NameOfAgeType)
'
'            add_record_to_table "TblTempItemAging", " ItemID,LateID ", ItemID & " ," & column_location, "ItemID", 0
'
'            Rs3.MoveNext
'        Next i
'
'    End If
'
'    Rs3.Close
'
'    Dim StrSQL As String
'
'End Function
'
'Public Function GetNetsalaryVouchers(NoteType As Integer, _
'                                     Fromdate As Date, _
'                                     todate As Date) As Double
'    Dim StrSQL                As String
'    Dim DepitValue            As Double
'    Dim CreditValue           As Double
'
'    Dim Account_Code_dynamic7 As String '–„„ «·„ÊŸ›Ì‰
'    Account_Code_dynamic7 = get_account_code_branch(7, my_branch)
'
'    If Account_Code_dynamic7 = "NO branch" Then
'        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'        Exit Function
'    Else
'
'        If Account_Code_dynamic7 = "NO account" Then
'            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    –„„ «·„ÊŸ›Ì‰          ", vbCritical
'
'            Exit Function
'        End If
'    End If
'
'    Dim Account_Code_dynamic29 As String '«·«ÃÊ— «·„” Õﬁ… «·„ÊŸ›Ì‰
'    Account_Code_dynamic29 = get_account_code_branch(29, my_branch)
'
'    If Account_Code_dynamic29 = "NO branch" Then
'        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'        Exit Function
'    Else
'
'        If Account_Code_dynamic29 = "NO account" Then
'            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    «·«ÃÊ—  «·„” Õﬁ… «·„ÊŸ›Ì‰          ", vbCritical
'
'            Exit Function
'        End If
'    End If
'
'    StrSQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
'    StrSQL = StrSQL & " FROM         dbo.Notes INNER JOIN"
'    StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
'    StrSQL = StrSQL & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
'    StrSQL = StrSQL & " WHERE     (dbo.Notes.NoteType = " & NoteType & ") AND (dbo.Notes.NoteDate >=" & SQLDate(Fromdate, True) & " ) AND (dbo.Notes.NoteDate <= " & SQLDate(todate, True) & ")"
'    'StrSQL = StrSQL & " AND (branch_no = " & Val(P_dcBranch) & ")"
'    StrSQL = StrSQL & "  AND (dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic7 & "' AND dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic29 & "')"
'    'StrSQL = StrSQL & "  AND (branch_no = " & val(P_dcBranch) & ")"
'
'    StrSQL = StrSQL & " GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
'    StrSQL = StrSQL & " HAVING      (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0)"
'
'    Dim RsUnitData As New ADODB.Recordset
'
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        DepitValue = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
'    Else
'        DepitValue = 0
'
'    End If
'
'    RsUnitData.Close
'
'    StrSQL = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS Total, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
'    StrSQL = StrSQL & " FROM         dbo.Notes INNER JOIN"
'    StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
'    StrSQL = StrSQL & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
'    StrSQL = StrSQL & " WHERE     (dbo.Notes.NoteType = " & NoteType & ") AND (dbo.Notes.NoteDate >=" & SQLDate(Fromdate, True) & " ) AND (dbo.Notes.NoteDate <= " & SQLDate(todate, True) & ")"
'    'StrSQL = StrSQL & " AND (branch_no = " & Val(P_dcBranch) & ")"
'    StrSQL = StrSQL & "  AND (dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic7 & "' AND dbo.ACCOUNTS.Parent_Account_Code <> N'" & Account_Code_dynamic29 & "')"
'    'StrSQL = StrSQL & "  AND (branch_no = " & val(P_dcBranch) & ")"
'
'    StrSQL = StrSQL & " GROUP BY dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit"
'    StrSQL = StrSQL & " HAVING      (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1)"
'
'    Dim RsUnitData1 As New ADODB.Recordset
'
'    RsUnitData1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData1.RecordCount) > 0 Then
'
'        CreditValue = IIf(IsNull(RsUnitData1("Total").value), 0, (RsUnitData1("Total").value))
'    Else
'        CreditValue = 0
'
'    End If
'
'    RsUnitData1.Close
'
'    GetNetsalaryVouchers = Abs(DepitValue - CreditValue)
'
'End Function
'
'Public Function CostForMaintenance(TicktNO As String) As Double
'
'    Dim StrSQL As String
'    StrSQL = "SELECT     SUM(dbo.Transaction_Details.Price) AS Cost"
'    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
'    StrSQL = StrSQL & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.TicketNO = N'" & TicktNO & "')"
'
'    Dim RsUnitData As New ADODB.Recordset
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        CostForMaintenance = IIf(IsNull(RsUnitData("Cost").value), 0, (RsUnitData("Cost").value))
'
'    Else
'        CostForMaintenance = 0
'
'    End If
'
'    RsUnitData.Close
'
'End Function
'Public Function CheckCustomerID(CustGID As Double, _
'                                Optional ByRef Custcode As String, _
'                                Optional ByRef CustName As String, _
'                                Optional ByRef block As Boolean = False, _
'                                Optional ByRef reson As String) As Boolean
'
'    Dim StrSQL As String
'    StrSQL = "SELECT    *  FROM      TblCustemers where CustGID=" & CustGID
'
'    Dim RsUnitData As New ADODB.Recordset
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        CheckCustomerID = True
'        Custcode = IIf(IsNull(RsUnitData("Fullcode").value), "", (RsUnitData("Fullcode").value))
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            CustName = IIf(IsNull(RsUnitData("CusName").value), 0, (RsUnitData("CusName").value))
'        Else
'            CustName = IIf(IsNull(RsUnitData("CusNamee").value), 0, (RsUnitData("CusNamee").value))
'        End If
'        If RsUnitData("locked").value = True Then
'            block = True
'            reson = IIf(IsNull(RsUnitData("Remark2").value), "", (RsUnitData("Remark2").value))
'        End If
'
'    Else
'        CheckCustomerID = False
'
'    End If
'
'    RsUnitData.Close
'
'End Function
'
'Public Function CheckCustomerIDold(CustGID As Double, _
'                                   Optional ByRef Custcode As String, _
'                                   Optional ByRef CustName As String) As Boolean
'
'    Dim StrSQL As String
'    StrSQL = "SELECT    *  FROM      TblCustemers where CustGID=" & CustGID
'
'    Dim RsUnitData As New ADODB.Recordset
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        CheckCustomerIDold = True
'        Custcode = IIf(IsNull(RsUnitData("Fullcode").value), "", (RsUnitData("Fullcode").value))
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            CustName = IIf(IsNull(RsUnitData("CusName").value), 0, (RsUnitData("CusName").value))
'        Else
'            CustName = IIf(IsNull(RsUnitData("CusNamee").value), 0, (RsUnitData("CusNamee").value))
'        End If
'
'    Else
'        CheckCustomerIDold = False
'
'    End If
'
'    RsUnitData.Close
'
'End Function
'
'Public Function GetScreenDescription(ScreenName As String) As String
'    Dim StrSQL As String
'    StrSQL = "SELECT    *  FROM      TblWorkFollow where ScreenName='" & ScreenName & "'"
'
'    Dim RsUnitData As New ADODB.Recordset
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        GetScreenDescription = IIf(IsNull(RsUnitData("Remark").value), 0, (RsUnitData("Remark").value))
'    Else
'        GetScreenDescription = ""
'
'    End If
'
'    RsUnitData.Close
'
'End Function
'Public Function GetFirstBox() As Integer
'    Dim StrSQL As String
'    StrSQL = "SELECT    * from TblBoxesData "
'    Dim RsUnitData As New ADODB.Recordset
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'        GetFirstBox = IIf(IsNull(RsUnitData("BoxID").value), 0, (RsUnitData("BoxID").value))
'    Else
'        GetFirstBox = 0
'
'    End If
'
'    RsUnitData.Close
'
'End Function
'
'Public Function GetSalesValue(Fromdate As Date, _
'                              todate As Date, _
'                              ItemType As Integer) As Double
'    Dim StrSQL As String
'    StrSQL = "SELECT     SUM(Total) AS Totalvalue, SUM(TotalQty) AS totalQty"
'    StrSQL = StrSQL & " FROM         dbo.QryItemsSalesTotal(21, DEFAULT, DEFAULT, " & SQLDate(Fromdate, True) & ", " & SQLDate(todate, True) & "," & ItemType & ") QryItemsSalesTotal"
'    Dim RsUnitData As New ADODB.Recordset
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'        GetSalesValue = IIf(IsNull(RsUnitData("Totalvalue").value), 0, (RsUnitData("Totalvalue").value))
'    Else
'        GetSalesValue = 0
'
'    End If
'
'    RsUnitData.Close
'
'End Function
'
'Public Function GetTransactionsEData(Transaction_ID As Integer, _
'                                     Optional ByRef TransactionEnglishName As String, _
'                                     Optional ByRef CusNamee As String, _
'                                     Optional ByRef storenamee As String)
'    Dim StrSQL As String
'    StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Type, dbo.TransactionTypes.TransactionEnglishName, dbo.TransactionTypes.TransactionTypeName, "
'    StrSQL = StrSQL & "   dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee"
'    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
'    StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
'    StrSQL = StrSQL & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
'    StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID"
'    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_ID = " & Transaction_ID & ")"
'    Dim RsUnitData As New ADODB.Recordset
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        TransactionEnglishName = IIf(IsNull(RsUnitData("TransactionEnglishName").value), 0, (RsUnitData("TransactionEnglishName").value))
'        CusNamee = IIf(IsNull(RsUnitData("CusNamee").value), 0, (RsUnitData("CusNamee").value))
'        storenamee = IIf(IsNull(RsUnitData("StoreNamee").value), 0, (RsUnitData("StoreNamee").value))
'
'    Else
'        TransactionEnglishName = ""
'        CusNamee = ""
'        storenamee = ""
'
'    End If
'
'    RsUnitData.Close
'
'End Function
'
'Public Function GetISSueVoucherForProductionValue(Fromdate As Date, _
'                                                  todate As Date, _
'                                                  ItemType As Integer) As Double
'    Dim StrSQL As String
'    StrSQL = "SELECT     SUM(Total) AS Totalvalue, SUM(TotalQty) AS totalQty"
'    StrSQL = StrSQL & " FROM         dbo.QryItemsSalesTotal(27, DEFAULT, DEFAULT, " & SQLDate(Fromdate, True) & ", " & SQLDate(todate, True) & "," & ItemType & ") QryItemsSalesTotal"
'    Dim RsUnitData As New ADODB.Recordset
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        GetISSueVoucherForProductionValue = IIf(IsNull(RsUnitData("Totalvalue").value), 0, (RsUnitData("Totalvalue").value))
'    Else
'        GetISSueVoucherForProductionValue = 0
'
'    End If
'
'    RsUnitData.Close
'
'End Function
'
'Public Function GetExpensestotal(Fromdate As Date, _
'   todate As Date) As Double
'    Dim StrSQL As String
'
'    ' „ «· ⁄œÌ· ·«Õ ”«» «·„’—Ê› «·„Œ›÷ ›Ì 18 12 2012
'    StrSQL = "  SELECT     SUM ("
'    StrSQL = StrSQL & "  Case"
'    StrSQL = StrSQL & "    When Credit_Or_Debit=0 Then Value*1"
'    StrSQL = StrSQL & " When Credit_Or_Debit=1 Then Value*-1"
'    StrSQL = StrSQL & " Else  0"
'    StrSQL = StrSQL & " End"
'    StrSQL = StrSQL & " ) AS Total"
'    StrSQL = StrSQL & " FROM         dbo.ACCOUNTS INNER JOIN"
'    StrSQL = StrSQL & " dbo.ExpensesType ON dbo.ACCOUNTS.Account_Code = dbo.ExpensesType.Account_Code INNER JOIN"
'    StrSQL = StrSQL & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
'    StrSQL = StrSQL & " WHERE     (dbo.ExpensesType.TypicalProduction = 1)  AND"
'    StrSQL = StrSQL & "       RecordDate >= " & SQLDate(Fromdate, True)
'    StrSQL = StrSQL & "  AND RecordDate <= " & SQLDate(todate, True)
'    'StrSQL = StrSQL & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = " & val(P_dcBranch) & ")"
'    'StrSQL = StrSQL & " AND (DOUBLE_ENTREY_VOUCHERS.branch_id = 1)"
'    Debug.Print StrSQL
'
'    Dim RsUnitData As New ADODB.Recordset
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        GetExpensestotal = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
'    Else
'        GetExpensestotal = 0
'
'    End If
'
'    RsUnitData.Close
'End Function
'
'Public Function gettotal(NoteType As Integer, _
'                         Fromdate As Date, _
'                         todate As Date, _
'                         Optional AllocationType As Integer = -1, _
'                         Optional branch_no As Integer = 1) As Double
'    Dim StrSQL As String
'
'    StrSQL = "  SELECT     SUM(Note_Value) AS Total from dbo.Notes"
'
'    StrSQL = StrSQL & " WHERE      NoteDate >= " & SQLDate(Fromdate, True)
'    StrSQL = StrSQL & "  AND NoteDate <= " & SQLDate(todate, True)
'    StrSQL = StrSQL & " AND (NoteType = " & NoteType & ")"
'    'StrSQL = StrSQL & " AND (branch_no = " & val(P_dcBranch) & ")"
'    StrSQL = StrSQL & " AND (branch_no =" & branch_no & ")"
'    If AllocationType <> -1 Then
'        StrSQL = StrSQL & " AND  AllocationType=" & AllocationType
'    End If
'
'    Dim RsUnitData As New ADODB.Recordset
'
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        gettotal = IIf(IsNull(RsUnitData("Total").value), 0, (RsUnitData("Total").value))
'    Else
'        gettotal = 0
'
'    End If
'
'    RsUnitData.Close
'End Function
'Public Function gettransactiontotal(Transaction_ID As Integer) As Double
'
'    Dim StrSQL As String
'
'    StrSQL = "   SELECT     QryTransactionsTotal.TransNet, dbo.Transactions.Transaction_ID"
'    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
'    StrSQL = StrSQL & "                       dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID"
'    StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_ID = " & Transaction_ID & ")"
'
'    Dim RsUnitData As New ADODB.Recordset
'
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        gettransactiontotal = IIf(IsNull(RsUnitData("TransNet").value), 0, (RsUnitData("TransNet").value))
'    Else
'        gettransactiontotal = 0
'
'    End If
'
'    RsUnitData.Close
'
'End Function
'
'Public Function GetSalesCost(Fromdate As Date, _
'   todate As Date) As Double
'    ' ﬂ·›… ”‰œ«  «·’—› «·„Œ“‰Ì
'    Dim StrSQL As String
'    StrSQL = "  SELECT     SUM(dbo.Transaction_Details.SHOWQTY * dbo.Transaction_Details.SHOWPrice) AS TotalCost"
'    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
'    StrSQL = StrSQL & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'    StrSQL = StrSQL & " WHERE     (dbo.Transactions.Transaction_Type = 19) and    (dbo.Transactions.Transaction_Date  >= " & SQLDate(Fromdate, True)
'    StrSQL = StrSQL & "  AND (dbo.Transactions.Transaction_Date  <= " & SQLDate(todate, True)
'    StrSQL = StrSQL & " ))"
'    'StrSQL = StrSQL & " AND (dbo.Transaction_Details.BranchId  = " & val(P_dcBranch) & ") and Doctype is null"
'    StrSQL = StrSQL & "  and  ( Doctype is null  or Doctype in(SELECT     id FROM         dbo.TblDoCumentsTypes  WHERE     (WorkWithProducction = 1))   )  "
'
'    Dim RsUnitData As New ADODB.Recordset
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        GetSalesCost = IIf(IsNull(RsUnitData("TotalCost").value), 0, (RsUnitData("TotalCost").value))
'    Else
'        GetSalesCost = 0
'
'    End If
'
'    RsUnitData.Close
'End Function
'
''Â–… «·œ«·Â ··›’· »Ì‰ ›« Ê—… „«·Ì… Ê›Ê« Ì— «·«’Ê·
'Public Function GetFinInvoiceType(NoteID As Double) As Double
'    Dim StrSQL As String
'
'    StrSQL = "   SELECT     bill_type From dbo.notes_all Where (noteid = " & NoteID & ")"
'
'    Dim RsUnitData As New ADODB.Recordset
'    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If (RsUnitData.RecordCount) > 0 Then
'
'        GetFinInvoiceType = IIf(IsNull(RsUnitData("bill_type").value), 0, (RsUnitData("bill_type").value))
'    Else
'        GetFinInvoiceType = 0
'
'    End If
'
'    RsUnitData.Close
'End Function
'
'Public Function GetEmployeeSalaryProject(Emp_id As Integer, _
'                                         whrstr As String, _
'                                         Optional MonthID As Integer = 0, _
'                                         Optional YearID As Integer = 0) As Double
'    Dim sql As String
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    sql = " SELECT     dbo.ProJectMofrdSalar.EmpID, SUM(dbo.ProJectMofrdSalar.Total) AS SumValuee"
'    sql = sql & "   FROM         dbo.mofrad RIGHT OUTER JOIN"
'    sql = sql & "                    dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type RIGHT OUTER JOIN"
'    sql = sql & "                    dbo.ProJectMofrdSalar ON dbo.mofrdat.mofrad_code = dbo.ProJectMofrdSalar.MofrdID"
'    sql = sql & " Where (dbo.mofrdat.mofrad_type  IN (" & whrstr & ")) And (dbo.ProJectMofrdSalar.YearID = " & YearID & ") And (dbo.ProJectMofrdSalar.MonthID = " & MonthID & ")"
'    sql = sql & " GROUP BY dbo.ProJectMofrdSalar.EmpID"
'    sql = sql & " HAVING      (dbo.ProJectMofrdSalar.EmpID = " & Emp_id & ")"
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        GetEmployeeSalaryProject = Abs(IIf(IsNull(Rs3("SumValuee").value), 0, Rs3("SumValuee").value))
'    Else
'        GetEmployeeSalaryProject = 0
'    End If
'End Function
'
'Public Function getEmployeeCashAssest(EmpID As Integer)
'    Dim sql          As String
'    Dim rs           As ADODB.Recordset
'    Dim i            As Integer
'    Dim Balance      As Double
'    Dim Account_Code As String
'    Balance = 0
'
'    sql = "SELECT *    from TblBoxesData  WHERE     empid = " & EmpID & " and Type =1 "
'
'    Set rs = New ADODB.Recordset
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adcmtext
'
'    If rs.RecordCount > 0 Then
'
'        For i = 1 To rs.RecordCount
'            Account_Code = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
'            Balance = Balance + get_balanceFromGl(Account_Code)
'
'            rs.MoveNext
'        Next i
'
'    End If
'
'    getEmployeeCashAssest = Balance
'End Function
'
'Public Function GetActiveInvestmenAccound(Optional InveID As Double = 0) As String
'    Dim Rs8 As ADODB.Recordset
'    Set Rs8 = New ADODB.Recordset
'    Dim sql As String
'    sql = "Select Account_Code6 from TblActivateInvestment where InviseNo=" & InveID & ""
'    Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs8.RecordCount > 0 Then
'        GetActiveInvestmenAccound = IIf(IsNull(Rs8("Account_Code6").value), "", Rs8("Account_Code6").value)
'    Else
'        GetActiveInvestmenAccound = ""
'    End If
'End Function
'Public Function get_employee_information(ID As Integer, _
'                                         Optional ByRef date1 As Date, _
'                                         Optional ByRef DepartmentID As Double, _
'                                         Optional ByRef SpecificationID As Double, _
'                                         Optional ByRef JobTypeID As Double, _
'                                         Optional ByRef gradeID As Double, _
'                                         Optional ByRef Account_code2 As String, _
'                                         Optional ByRef Account_Code As String, _
'                                         Optional ByRef endContractPerMonth As Double, _
'                                         Optional ByRef Nationality As String, _
'                                         Optional ByRef mangerid As Integer, _
'                                         Optional ByRef swapedempid As Integer, _
'                                         Optional ByRef GroupID As Integer, _
'                                         Optional ByRef NumPasp As String, _
'                                         Optional ByRef NumEkama As String, _
'                                         Optional ByRef placeEkama As String, _
'                                         Optional ByRef pasplace As String, _
'                                         Optional ByRef DateEndekamaH As String, _
'                                         Optional ByRef DateEndPasp As Date, _
'                                         Optional ByRef BignDateWork As Date, _
'                                         Optional ByRef LastDate As Date, _
'                                         Optional ByRef JobTypeName As String, _
'                                         Optional ByRef Contract_period1 As Integer, _
'                                         Optional ByRef Contract_periodno1 As Integer, _
'                                         Optional ByRef visano As String, Optional ByRef dcjopstatus As Integer, Optional ByRef JobTypeIDIqama As Integer, Optional ByRef DateMoveNo As Date, Optional ByRef DateExpoekama As String, Optional ByRef Mobile As String, Optional ByRef BlnceVocat As Integer = 0, Optional ByRef Emp_Phone As String, Optional ByRef Contract_date1 As Date, Optional ByRef RegionID As Integer = 0, Optional ByRef due_period As Integer, Optional ByRef Due_period_no As Integer, Optional ByRef Holiday_period_no As Integer, Optional ByRef Holiday_period As Integer, Optional BranchID As Integer, Optional DriverLicenseendH As String, Optional DriverLicense As String, Optional ByRef lastHolidaydate As Date, Optional ByRef lastHolidaydateH As String, Optional ADDtype_Contract As Integer)
'
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim sql As String
'    sql = " SELECT     dbo.TblEmpSpecifications.SpecificationName AS Expr1, dbo.TblEmpDepartments.DepartmentName AS Expr2, dbo.TblEmpDepartments.DepartmentNamee AS Expr3,"
'    sql = sql & "                      dbo.TblEmpJobsTypes.JobTypeName AS Expr4, dbo.TblEmpJobsTypes.JobTypeNamee AS Expr5, dbo.TblEmpGrades.namee AS grdename,"
'    sql = sql & "                      dbo.TblEmpGrades.name AS grdenamee, dbo.TblEmployee.*, dbo.TblEmployee.JobTypeID3 AS JobTypeID3Iq, TblEmpJobsTypes_1.JobTypeName AS jobnameiqama,"
'    sql = sql & "                      TblEmpJobsTypes_1.JobTypeNamee AS jobnameiqamaE, dbo.Contract.Contract_period_no , dbo.Contract.Contract_period "
'    sql = sql & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
'    sql = sql & "                      dbo.Contract ON dbo.TblEmployee.Emp_ID = dbo.Contract.Emp_id LEFT OUTER JOIN"
'    sql = sql & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_1 ON dbo.TblEmployee.JobTypeID3 = TblEmpJobsTypes_1.JobTypeID LEFT OUTER JOIN"
'    sql = sql & "                      dbo.TblEmpGrades ON dbo.TblEmployee.gradeID = dbo.TblEmpGrades.gradeid LEFT OUTER JOIN"
'    sql = sql & "                      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
'    sql = sql & "                      dbo.TblEmpSpecifications ON dbo.TblEmployee.SpecificationID = dbo.TblEmpSpecifications.SpecificationID LEFT OUTER JOIN"
'    sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID"
'
'    sql = sql & " WHERE     (dbo.TblEmployee.Emp_ID = " & ID & ")"
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount > 0 Then
'
'        date1 = IIf(Not IsDate(Rs3("BignDateWork").value), Date, Rs3("BignDateWork").value)
'        DepartmentID = IIf(Not IsNull(Rs3("DepartmentID").value), Rs3("DepartmentID").value, 0)
'
'        SpecificationID = IIf(Not IsNull(Rs3("SpecificationID").value), Rs3("SpecificationID").value, 0)
'        JobTypeID = IIf(Not IsNull(Rs3("JobTypeID").value), Rs3("JobTypeID").value, 0)
'        JobTypeIDIqama = IIf(Not IsNull(Rs3("JobTypeID3").value), Rs3("JobTypeID3").value, 0)
'        gradeID = IIf(Not IsNull(Rs3("gradeID").value), Rs3("gradeID").value, 0)
'        Account_code2 = IIf(Not IsNull(Rs3("Account_Code2").value), Rs3("Account_Code2").value, "")
'        Account_Code = IIf(Not IsNull(Rs3("Account_code").value), Rs3("Account_code").value, "")
'        Nationality = IIf(Not IsNull(Rs3("Nationality").value), Rs3("Nationality").value, "")
'        mangerid = IIf(Not IsNull(Rs3("mangerid").value), Rs3("mangerid").value, 0)
'        swapedempid = IIf(Not IsNull(Rs3("swapedempid").value), Rs3("swapedempid").value, 0)
'        GroupID = IIf(Not IsNull(Rs3("GroupID").value), Rs3("GroupID").value, 0)
'        NumEkama = IIf(Not IsNull(Rs3("NumEkama").value), Rs3("NumEkama").value, "")
'        NumPasp = IIf(Not IsNull(Rs3("NumPasp").value), Rs3("NumPasp").value, "")
'        JobTypeName = IIf(Not IsNull(Rs3("Expr4").value), Rs3("Expr4").value, "")
'        Contract_period1 = IIf(Not IsNull(Rs3("Contract_period").value), Rs3("Contract_period").value, -1)
'        Contract_periodno1 = IIf(Not IsNull(Rs3("Contract_period_no").value), Rs3("Contract_period_no").value, 0)
'        visano = IIf(Not IsNull(Rs3("VisaNo").value), Rs3("VisaNo").value, "")
'        dcjopstatus = IIf(Not IsNull(Rs3("jopstatusid").value), Rs3("jopstatusid").value, 0)
'        DateMoveNo = IIf(IsNull(Rs3("DateMoveno").value), Date, Rs3("DateMoveno"))
'        DateEndekamaH = IIf(Not IsNull(Rs3("DateEndekamah").value), Rs3("DateEndekamah").value, ToHijriDate(Date))
'        DateExpoekama = IIf(Not IsNull(Rs3("DateExpoekamah").value), Rs3("DateExpoekamah").value, ToHijriDate(Date))
'        DateEndPasp = IIf(Not IsNull(Rs3("DateEndPasp").value), Rs3("DateEndPasp").value, Date)
'        pasplace = IIf(Not IsNull(Rs3("pasplace").value), Rs3("pasplace").value, "")
'        placeEkama = IIf(Not IsNull(Rs3("placeEkama").value), Rs3("placeEkama").value, "")
'        RegionID = IIf(Not IsNull(Rs3("RegionID").value), Rs3("RegionID").value, 0)
'        Emp_Phone = IIf(Not IsNull(Rs3("Emp_Phone").value), Rs3("Emp_Phone").value, "")
'        BignDateWork = IIf(Not IsNull(Rs3("BignDateWork").value), Rs3("BignDateWork").value, Date)
'        LastDate = IIf(Not IsNull(Rs3("LastDate").value), Rs3("LastDate").value, Date)
'        Mobile = IIf(Not IsNull(Rs3("Emp_mobile").value), Rs3("Emp_mobile").value, "")
'        BlnceVocat = IIf(Not IsNull(Rs3("BlnceVocat").value), Rs3("BlnceVocat").value, 0)
'        BranchID = IIf(Not IsNull(Rs3("BranchId").value), Rs3("BranchId").value, 0)
'
'        DriverLicense = IIf(Not IsNull(Rs3("DriverLicense").value), Rs3("DriverLicense").value, "")
'        DriverLicenseendH = IIf(Not IsNull(Rs3("DriverLicenseendH").value), Rs3("DriverLicenseendH").value, ToHijriDate(Date))
'        lastHolidaydate = IIf(Not IsNull(Rs3("lastHolidaydate").value), Rs3("lastHolidaydate").value, Date)
'        lastHolidaydateH = IIf(Not IsNull(Rs3("lastHolidaydateH").value), Rs3("lastHolidaydateH").value, ToHijriDate(Date))
'        'GroupID
'
'    Else
'        BranchID = 0
'        date1 = Date
'        DepartmentID = 0
'        SpecificationID = 0
'        JobTypeID = 0
'        gradeID = 0
'        Nationality = ""
'        mangerid = 0
'        swapedempid = 0
'        GroupID = 0
'        lastHolidaydate = Date
'        lastHolidaydateH = ToHijriDate(Date)
'    End If
'
'    Rs3.Close
'    Dim Contract_period_no As Double
'    Dim Contract_period    As Double
'
'    Dim Contract_date      As Date
'    sql = "  select * from Contract WHERE     (dbo.Contract.Emp_ID = " & ID & ")"
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs3.RecordCount > 0 Then
'        Contract_period_no = IIf(Not IsNull(Rs3("Contract_period_no").value), Rs3("Contract_period_no").value, 0)
'        Contract_period = IIf(Not IsNull(Rs3("Contract_period").value), Rs3("Contract_period").value, 0)
'        Contract_date = IIf(Not IsNull(Rs3("Contract_date").value), Rs3("Contract_date").value, Date)
'        Contract_date1 = IIf(Not IsNull(Rs3("Contract_date").value), Rs3("Contract_date").value, Date)
'        endContractPerMonth = DateDiff("M", Contract_date, Date)
'        Due_period_no = IIf(Not IsNull(Rs3("Due_period_no").value), Rs3("Due_period_no").value, -1)
'        due_period = IIf(Not IsNull(Rs3("due_period").value), Rs3("due_period").value, 0)
'        endContractPerMonth = DateDiff("M", Contract_date, Date)
'        Holiday_period_no = IIf(IsNull(Rs3("Holiday_period_no").value), 0, Rs3("Holiday_period_no").value)
'        ADDtype_Contract = IIf(IsNull(Rs3("ADDtype_Contract").value), 0, Rs3("ADDtype_Contract").value)
'        If IsNull(Rs3("Holiday_period").value) Then
'            Holiday_period = 0
'        Else
'            Holiday_period = Rs3("Holiday_period").value
'        End If
'
'        If Contract_period = 1 Then
'
'            Contract_period_no = Contract_period_no * 12
'        End If
'
'        endContractPerMonth = Contract_period_no - endContractPerMonth
'
'    Else
'
'    End If
'
'End Function
'
'Public Function OpenRecordSet(ByVal SqlStatment As String, _
'                              OpenType As CursorTypeEnum, _
'                              LockType As LockTypeEnum, _
'                              Optional IsLocal As Boolean = False, _
'                              Optional TrayCachFrist As Boolean = False, _
'                              Optional CursorLocation As Integer = -1, _
'                              Optional UseFunctionLockType As Boolean = False) As ADODB.Recordset
'    AduitLastSQL = SqlStatment
'    'Samy Commnt this line
'    '    If StringStartsWith(LTrim$(UCase(SqlStatment)), "EXEC ") Then
'    '        Set OpenRecordSet = OpenRecordSetForSP(SqlStatment, OpenType, LockType, IsLocal, TrayCachFrist, CursorLocation)
'    '        Exit Function
'    '    End If
'    '********************
'    BuildAppInfo
'    '*********************
'    '    Do While IsTestConnection
'    '        DoEvents
'    '    Loop
'
'RE:
'    On Error GoTo EH
'
'    '    DatabaseName = "Topsystems111"
'    If IsLocal Then
'        '        StrConn1 = "Driver={SQL Server};Packet Size=32768;Server=" & ServerNameLocal & _
'        '           ";Uid=" & UserNameLocal & ";Pwd=" & PASSWORDLocal & _
'        '           ";Database=" & DatabaseNameLocal & ";App=" & MyAPPUserInfo
'
'        StrConn1 = "Provider=SQLNCLI11.1;Data Source=" & ServerNameLocal & _
'           ";User ID=" & UserNameLocal & ";Password=" & PASSWORDLocal & _
'           ";Initial Catalog=" & DatabaseNameLocal & ";DataTypeCompatibility=80;Application Name=" & MyAPPUserInfo
'        '-----------------------------------------------
'        If DBLocal.State <> adStateOpen Then
'            If Not DBLocalIsCustomStringConnection Then
'                DBLocal.ConnectionTimeout = 5000
'                DBLocal.CommandTimeout = 5000
'                DBLocal.IsolationLevel = adXactReadUncommitted
'            Else
'                StrConn1 = DBLocalCustomStringConnection
'                StrConn1 = SetConnectionSection(StrConn1, "Data Source", ServerNameLocal)
'                StrConn1 = SetConnectionSection(StrConn1, "Initial Catalog", DatabaseNameLocal)
'                StrConn1 = SetConnectionSection(StrConn1, "User ID", UserNameLocal)
'                StrConn1 = SetConnectionSection(StrConn1, "Password", PASSWORDLocal)
'            End If
'            '--------------------
'            DBLocal.Open StrConn1
'            '--------------------
'        End If
'        '--------------------------------------
'        DBLocal.CommandTimeout = 5000
'        '--------------------------------------
'        Set OpenRecordSet = New ADODB.Recordset
'        If CursorLocation <> -1 Then
'            OpenRecordSet.CursorLocation = CursorLocation
'        End If
'        OpenRecordSet.Open SqlStatment, DBLocal, OpenType, LockType, adCmdText
'
'    Else
'        '        StrConn = "Driver={SQL Server};Packet Size=32768;Server=" & ServerName & _
'        '           ";Uid=" & UserName & ";Pwd=" & password & _
'        '           ";Database=" & DatabaseName & ";App=" & MyAPPUserInfo
'        'Provider=SQLOLEDB.1;Password=makkahttd;Persist Security Info=True;User ID=sa;Data Source=196.186.0.205\SQL2008
'        StrConn = "Provider=SQLNCLI11.1;Data Source=" & ServerName & _
'           ";User ID=" & UserName & ";Password=" & Password & _
'           ";Initial Catalog=" & DatabaseName & ";DataTypeCompatibility=80;Application Name=" & MyAPPUserInfo
'        ' -----------------------------------------------
'        '·⁄·«Ã «Œ ·«› «·œ« « »Ì“ ›Ì Õ«·… «·DLL
'        ' -----------------------------------------------
'        If isDebugMode() Then
'            If db.State = adStateOpen Then
'                If UCase(ServerName) <> UCase(ServerNameINI) Then
'                    Set tt = db.Execute("SELECT SERVERPROPERTY(N'MachineName')AS MachineName, CONNECTIONPROPERTY('local_net_address') AS IPAddress,SERVERPROPERTY('InstanceName') AS InstanceName;")
'                    MyMachineName = StrConv(StrConv(tt!MachineName, vbUnicode), vbFromUnicode)    ' TT!MachineName & ""
'                    MyIpAddress = CStr(tt!IPAddress) & ""    '  StrConv(tt!IPAddress, vbUnicode)    'CStr(TT!IPAddress) & ""
'                    '  MyIPAddress = Replace(CStr(MyIPAddress), " ", "")
'                    MyInstanceName = StrConv(StrConv(tt!InstanceName, vbUnicode), vbFromUnicode)    'CStr(TT!InstanceName) & ""
'                    If MyInstanceName <> "" Then
'                        MyMachineName = MyMachineName & "\" & MyInstanceName
'                        MyIpAddress = MyIpAddress & "\" & MyInstanceName
'                    End If
'                    If UCase(ServerName) <> UCase(MyMachineName) And UCase(ServerName) <> UCase(MyIpAddress) Then
'                        'this run only once In same dll when click on form
'                        db.Close
'                    End If
'                End If
'            End If
'        End If
'        '-----------------------------------------------
'        If db.State <> adStateOpen Then
'            If Not DBIsCustomStringConnection Then
'                db.ConnectionTimeout = 50
'                db.CommandTimeout = 10000
'
'            Else
'                StrConn1 = DBCustomStringConnection
'                StrConn1 = SetConnectionSection(StrConn1, "Data Source", ServerName)
'                StrConn1 = SetConnectionSection(StrConn1, "Initial Catalog", DatabaseName)
'                StrConn1 = SetConnectionSection(StrConn1, "User ID", UserName)
'                StrConn1 = SetConnectionSection(StrConn1, "Password", Password)
'                StrConn = StrConn1
'                RptConn = StrConn1
'            End If
'            '---------------------
'            db.Open StrConn
'            '---------------------
'        End If
'        '--------------------------------------
'        db.CommandTimeout = 5000
'        '--------------------------------------
'
'        Set OpenRecordSet = New ADODB.Recordset
'
'        If CursorLocation <> -1 Then
'            OpenRecordSet.CursorLocation = CursorLocation
'        End If
'
'        Set OpenRecordSet.ActiveConnection = db
'
'        OpenRecordSet.Properties("Preserve On Commit").value = True
'        OpenRecordSet.Properties("Preserve On Abort").value = True
'        '*********************Samy************************************
'        If Not UseFunctionLockType Then
'            If LockType = adLockOptimistic Or LockType = adLockPessimistic Then
'                LockType = adLockOptimistic
'            End If
'        End If
'        '**********************************************************
'
'        OpenRecordSet.Open SqlStatment, , OpenType, LockType, adCmdText
'        '*********************
'        'AddToCollection OpenRecordSet
'        '***********************
'        '*****************************************************************************************
'        '****************Tray to caching Database that are readonly**********by Khalid************
'        '*****************************************************************************************
'        If TrayCachFrist And (OpenType = adOpenStatic) And (LockType = adLockReadOnly) Then
'            TempfileName = CreateTempFileName()
'            OpenRecordSet.save TempfileName, adPersistXML
'            OpenRecordSet.Close
'            OpenRecordSet.Open TempfileName, "Provider=mspersist"
'        End If
'        '*****************************************************************************************
'        '*****************************************************************************************
'        '*****************************************************************************************
'        If InStr(1, SqlStatment, "ActiveUsers ", vbTextCompare) = 0 And _
'           InStr(1, SqlStatment, "FormDesign ", vbTextCompare) = 0 And _
'           InStr(1, SqlStatment, "MenuRights ", vbTextCompare) = 0 And _
'           InStr(1, SqlStatment, "MenuShortCuts ", vbTextCompare) = 0 And _
'           InStr(1, SqlStatment, "EmployeeNoticeBoard ", vbTextCompare) = 0 And _
'           InStr(1, SqlStatment, "Translations ", vbTextCompare) = 0 And _
'           InStr(1, SqlStatment, "Translations2 ", vbTextCompare) = 0 _
'           Then
'            If QState Then QValue = QValue + vbNewLine + SqlStatment Else QValue = ""
'            '       If QValueP <> 0 Then CopyMemory QValueP, ByVal QValue, LenB(QValue)
'        End If
'        '*******************
'    End If
'    '**********************
'    ' „ «Ìﬁ«›Â ⁄‘«‰ «·»·Ã »·«Ï
'    '    If Not checkedSqlBefor Then
'    '        SqlServerVersionCheck
'    '    End If
'    '***********************
'    Exit Function
'EH:
'    MErr = Err.Number
'    If MErr = -2147467259 Then
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        'Samy
'        If Not isDebugMode Then
'            If Not NotShowTestConnection Then
'                FrmTestConnection.IsLocal = IsLocal
'                IsTestConnection = True
'                FrmTestConnection.show 1    ', FrmMain
'                If Not Tested Then
'                    '                Unload FrmTestConnection
'                    IsTestConnection = False
'                    Exit Function
'                Else
'                    '                Unload FrmTestConnection
'                    IsTestConnection = False
'                    GoTo RE
'                End If
'
'            End If
'        End If
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Else
'        If Err.Number = -2147217871 Then
'            MsgBox MyErrorHandler(Err.Number) & ":ORS", , 5
'        ElseIf Err.Number = -2147217900 Then
'            Resume Next
'        Else
'            MsgBox MyErrorHandler(Err.Number) & ":ORS"
'        End If
'    End If
'
'End Function
''**************************************
'
''**************************************
'
'Public Function MyErrorHandler(ErrNo As Long) As String
'    Mmsg = ""
'    Select Case ErrNo
'
'        Case 0
'            MyErrorHandler = ""
'            Exit Function
'
'        Case -2147217864
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = " „ ≈Ã—«¡  ⁄œÌ·«  ⁄·Ï Â–Â «·‘«‘Â „‰ ÃÂ«“ ¬Œ—- „‰ ›÷·ﬂ «⁄œ  Õ„Ì· «·Õ—ﬂÂ À„ Õ«Ê· „—Â «Œ—Ï" & " - Optimistic concurrency erorr "
'            Else
'                Mmsg = "This Form is editing from another computer- realod and Try again " & " - Optimistic concurrency erorr "
'            End If
'
'        Case -2147467259
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = "«·ÃÂ«“ «·Œ«œ„ «·—∆Ì”Ì „€·ﬁ √Ê €Ì— „ÊÃÊœ ⁄·Ï Â–Â «·‘»ﬂ…" & " - " & ErrNo
'            Else
'                Mmsg = "Couldn't close file , Because it's already in use by another person or process" & " - " & ErrNo
'            End If
'        Case -2147352567
'            'If SystemOptions.UserInterface = ArabicInterface Then
'            '    mMsg = "ÌÃ»  Œ’Ì’ «·ÿ«»⁄«  „‰ ≈œ«—… «·‰Ÿ«„" & " - " & ErrNo
'            'Else
'            '    mMsg = "Select Correct Report Printer Device" & " - " & ErrNo
'            'End If
'        Case 3155, 3022, -2147217873, -2147217900    ' insert fail
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = " ·«Ì„ﬂ‰ «÷«›… Â–« «·”Ã· ° Â–Â «·»Ì«‰«   „  ”ÃÌ·Â« „‰ ﬁ»·" & " - " & ErrNo
'            Else
'                Mmsg = "You Can not Add this Record , May be there is Dublicated values" & " - " & ErrNo
'            End If
'        Case 3200    ' Change Or Delete Failed
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = " ·«Ì„ﬂ‰ «·€«¡ √Ê  ⁄œÌ· Â–« «·”Ã·  »”»» ÊÃÊœ »Ì«‰«  √Œ—Ï „— »ÿ… »Â ÊÌÃ» «·€«¡Â« √Ê·«" & " - " & ErrNo
'            Else
'                Mmsg = "You Can not Delete Or Modify this Record , Because There Is Some Data Depends On It " & " - " & ErrNo
'            End If
'        Case 3157, 3046, 3202, 3218    ' Update Fail
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = " Â‰«ﬂ ›‘· ›Ï  Œ“Ì‰ «· ⁄œÌ·«  ° ﬁœ ÌﬂÊ‰ «·”Ã· „ﬁ›· »Ê«”ÿ… „” Œœ„ ¬Œ—° Õ«Ê· „—… √Œ—Ï" & " - " & ErrNo
'            Else
'                Mmsg = "Update Failed , May be The record is locked by another User , Try Again " & " - " & ErrNo
'            End If
'        Case 3186, 3187, 3188
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = "”Ã· „€·ﬁ »Ê«”ÿ… „” Œœ„ ¬Œ—" & " - " & ErrNo
'            Else
'                Mmsg = "Current Record locked by Another user" & " - " & ErrNo
'            End If
'        Case 3167
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = " „ «·€«¡ Â–« «·”Ã· »«·›⁄· " & " - " & ErrNo
'            Else
'                Mmsg = "Record Already Deleted" & " - " & ErrNo
'            End If
'        Case 3314
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = "„‰ ›÷·ﬂ √ﬂ„· «·»Ì«‰«  ﬁ»· «· Œ“Ì‰" & " - " & ErrNo
'            Else
'                Mmsg = "Please Complete the data before saving" & " - " & ErrNo
'            End If
'        Case 3262, 3211, 3212    ' Locked by another user and wait
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = "·« Ì„ﬂ‰ ≈€·«ﬁ «·„·› »”»» ÊÃÊœ „” Œœ„ ¬Œ— ÌﬁÊ„ »≈” Œœ«„Â √Ê ﬁ«„ »≈€·«ﬁÂ" & " - " & ErrNo
'            Else
'                Mmsg = "Couldn't close file , Because it's already in use by another person or process" & " - " & ErrNo
'            End If
'        Case 3197    ' Couldn't repaire this files
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = "√ﬂÀ— „‰ „” Œœ„ Õ«Ê·Ê«  €ÌÌ— ‰›” «·»Ì«‰«  ›Ï ‰›” «·Êﬁ " & " - " & ErrNo
'            Else
'                Mmsg = "Another Users are attempting to change the same data at the same time" & " - " & ErrNo
'            End If
'        Case 3056    ' Couldn't repaire this files
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = "·« Ì„ﬂ‰  ’·ÌÕ «·„·›«  «·„” Œœ„…" & " - " & ErrNo
'            Else
'                Mmsg = "Couldn't repaire this files" & " - " & ErrNo
'            End If
'        Case 3014, 3037    ' Can't open any more files
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = "·« Ì„ﬂ‰ › Õ „·›«  √Œ—Ï" & " - " & ErrNo
'            Else
'                Mmsg = "Can't open any more files" & " - " & ErrNo
'            End If
'        Case 3356, 3260, 3261, 3189, 3008, 3164, 3006    ' Table or Database Locked
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = "«·„·› „€·ﬁ »Ê«”ÿ… „” Œœ„ ¬Œ—" & " - " & ErrNo
'            Else
'                Mmsg = "The File is Locked by Another User" & " - " & ErrNo
'            End If
'        Case 3201    ' Add Or Edit Fail
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = " ·«Ì„ﬂ‰ «÷«›… Â–« «·”Ã· √Ê «· ⁄œÌ· ›ÌÂ ° ·√‰Â „— »ÿ »„·› ·„ Ì „ «·≈÷«›… √Ê «· ⁄œÌ· ›ÌÂ Õ Ï «·¬‰" & " - " & ErrNo
'            Else
'                Mmsg = "You Can not Add this Record or Change it , Because it's Linked to a File that has not been Added or Changed till Now" & " - " & ErrNo
'            End If
'        Case -2147217887
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Mmsg = "Œÿ√ €Ì— „⁄—Ê› ° Õ«Ê·  ‰›Ì– ‰›” «·⁄„·Ì… „—… √Œ—Ï" & " - " & ErrNo
'            Else
'                Mmsg = "Undefined Error , Try again : " & ErrNo
'            End If
'        Case 3704
'            '********By Khalid
'            On Error Resume Next
'            db.Close
'            Exit Function
'        Case -1000000001
'            '«—Ê—— » «⁄ «·«Ê Ê ﬂÊ„»Ì·Ì  ›ﬂﬂ „‰Â
'            MyErrorHandler = ""
'            Exit Function
'    End Select
'    '*************************
'    If Err.Number = vbObjectError + 1000 Then
'        If Not ArabicInterface Then
'            mText = Trim(Mmsg)
'            If Trim(mText) <> "" Then
'                Cond = "Arabic = N'" & Trim(mText) & "'"
'
'                s = "Select * from Translations where " & Cond
'                Set Translations = OpenRecordSet(s, adOpenStatic, adLockReadOnly)
'                '------------------------
'                If Not Translations.EOF Then
'                    Mmsg = IIf(Trim(Translations!English & "") <> "", Trim(Translations!English & ""), Mmsg)
'                End If
'            End If
'        End If
'
'        Mmsg = Mmsg & vbNewLine & Err.Description
'    Else
'        Mmsg = Mmsg & vbNewLine & Err.Description & " : " & Err.Number
'    End If
'    '*************************
'    If ErrNo <> -2147217864 Then  '  Ã«Â· «—Ê—— «·ﬂÊ‰ﬂ—‰”Ï  ‘Ìﬂ
'        If db.Errors.count > 0 Then
'            ss = ""
'            Dim adoErr As ADODB.Error
'            j = 1
'            On Error GoTo EEE
'            For Each adoErr In db.Errors
'                If adoErr.Number <> 0 Then
'                    If j = 1 Then ss = vbNewLine & "-------SQL Errors-------"
'                    ss = ss & vbNewLine & "Error (" & j & ")=" & adoErr.Description & " : " & adoErr.Number & ":" & adoErr.SQLState & ":" & adoErr.NativeError
'                    j = j + 1
'                End If
'            Next adoErr
'EEE:
'            ' for this rand error Not enough storage is available to process this command.
'            If Err.Number = 48 Then
'                Set adoErr = db.Errors(0)
'                ss = ss & vbNewLine & "Error (48)=" & adoErr.Description & " : " & adoErr.Number & ":" & adoErr.SQLState & ":" & adoErr.NativeError
'            End If
'            On Error GoTo 0
'            Mmsg = Mmsg & vbNewLine & ss
'        End If
'    End If
'    '*************************
'    'If Trim(mMsg) <> "()(0)" Then MyErrorHandler = mMsg Else MyErrorHandler = ""
'    MyErrorHandler = Mmsg & ":" & Erl
'    IsAboutError = True
'
'End Function
'
'Public Sub BuildAppInfo()
'    IsGetExeDate = True
'    ExeVer = App.Major & "." & App.Minor & "." & App.Revision
'    'œÏ Â ” Œœ„ ›Ï «·—”«Ì· «··Ï Â —ÊÕ ··ÌÊ“—“ ⁄‘«‰ ‰⁄—› „‰ «·ÌÊ“— «·«Ê‰ ·«Ì‰
'    MyAPPUserInfo = "" 'IIf(CurrentUser & "" = "", "•", CurrentUser) & "|" & App.EXEName & "|" & ExeVer & "|" & CustomerCode & "|" & GetEXEDateForConnectio
'
'End Sub
'
'Public Function SetConnectionSection(ByVal ConnStr As String, _
'                                     ByVal SectionName As String, _
'                                     ByVal SectionValue As String) As String
'    'StrConn1=SetConnectionSection(StrConn1,"Data Source","")
'    'StrConn1=SetConnectionSection(StrConn1,"Initial Catalog","")
'    'StrConn1=SetConnectionSection(StrConn1,"User ID","")
'    'StrConn1=SetConnectionSection(StrConn1,"Password","")
'    Dim isFound As Boolean
'    Dim vv
'    '"Provider=SQLNCLI.1;Password=pass;Persist Security Info=True;User ID=sa;Initial Catalog=Database;Data Source=Server"
'    vv = Split(ConnStr, ";")
'    l = Len(UCase(SectionName) & "=")
'    For i = 0 To UBound(vv)
'        If left(UCase(vv(i)), l) = left(UCase(SectionName) & "=", l) Then
'            vv(i) = SectionName & "=" & SectionValue
'            isFound = True
'            Exit For
'        End If
'    Next
'    If isFound Then
'        Dim ss As String
'        For i = 0 To UBound(vv)
'            ss = ss & vv(i) & ";"
'        Next
'        SetConnectionSection = ss
'    Else
'        SetConnectionSection = ConnStr
'    End If
'End Function
'
'Public Function isDebugMode() As Boolean
'    isDebugMode = isDebugModeC
'End Function
'
'Private Function CreateTempFileName(Optional ByVal Prefix As String) As String
'    Dim TempFile As String  ' receives name of temporary file
'    Dim slength  As Long   ' receives length of string returned for the path
'    Dim lastfour As Long  ' receives hex value of the randomly assigned ????
'    If TempPath = "" Then
'        ' Get Windows's temporary file path
'        TempPath = Space(255)  ' initialize the buffer to receive the path
'        slength = GetTempPath(255, TempPath)  ' read the path name
'        TempPath = left(TempPath, slength) & "BusinessDimensions\"    ' extract data from the variable
'        CreateDir TempPath
'        On Error Resume Next
'        Kill TempPath & "*.*"
'    End If
'
'    ' Get a uniquely assigned random file
'    '    TempFile = Space(255)  ' initialize buffer to receive the filename
'    '    If Prefix = "" Then Prefix = "TopSys" 'Format(Now, "YYYYMMDDHHNNSS")
'    '    lastfour = GetTempFileName(TempPath, Prefix, 0, TempFile)       ' get a unique temporary file name
'    '    ' (Note that the file is also created for you in this case.)
'    '    TempFile = Left(TempFile, InStr(TempFile, vbNullChar) - 1)   ' extract data from the variable
'    TempFile = TempPath & Prefix & Format(Now, "YYYYMMDDHHNNSS") & Int(Rnd(100) * 1000) & ".xml"
'    On Error Resume Next
'    Kill TempFile
'    CreateTempFileName = TempFile
'End Function
'
'Public Sub CreateDir(StrPath As String)
'    On Error Resume Next
'    Dim ArrFolders As Variant
'    ArrFolders = Split(StrPath, "\")
'    Dim i       As Long
'    Dim CurPath As String: CurPath = ArrFolders(0)
'    MkDir CurPath
'
'    For i = 1 To UBound(ArrFolders)
'        CurPath = CurPath & "\" & ArrFolders(i)
'        MkDir CurPath
'    Next i
'    On Error GoTo 0
'
'    If Len(Dir(StrPath, vbDirectory)) = 0 Then
'        Err.Raise vbObjectError, , "Can't create dir" & vbCrLf & StrPath & vbCrLf & ":(((("
'    End If
'End Sub
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Public Sub TranslateForm(Frm As Object, ByVal Arabic As Boolean)
'    On Error Resume Next
'    '------------------------
'    'If StrConn = "" Then Exit Sub    'for only one time in first instalation
'    '------------------------
'    '   If Arabic Or UCase(frm.Name) = "FRMMSGBOX" Then Exit Sub    ' „⁄—» „‰ Êﬁ  «· ’„Ì„
'    '--------------------------------
'    'If Arabic And frm.RightToLeft Then Exit Sub ' „⁄—» »«·›⁄·
'    'If Not Arabic And Not frm.RightToLeft Then Exit Sub
'    '-----------------------------------
'    Dim rsDummy As New ADODB.Recordset
'    Load Frm
'    '**********************************
'    Dim Ctr   As Control
'    Dim mText As String
'    '------------------------
'    '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'    '   frm.RightToLeft = Arabic
'    '   RTLTree frm
'    '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'    '------------------------
'    ' Field Captions & Visibility
'    '------------------------
'    If Trim(Frm.Name) = "frmReports" Then
'        Msg = "·« Ì„ﬂ‰  —Ã„Â «· ﬁ—Ì— „‰ Â‰« .... «€·ﬁ «· ﬁ—Ì— Êﬁ„ »«⁄«œÂ › ÕÂ „⁄ «Œ Ì«— «· —Ã„Â «·„ÿ·Ê»Â"
'
'        Msg2 = "Can't Translate Report from here .. reopn the report with choice languge  "
'        '        MyMsgbox IIf(ArabicInterface, Msg, Msg2)
'
'        '        TranslateReport Rpt, Arabic
'        '        frmReports.CV1.Refresh
'        Exit Sub
'    End If
'    '------------------------
'    mArabicCaption = ""
'    mEnglishCaption = ""
'    '------------------------
'    mText = Trim(Frm.Caption)
'    mArabicCaption = mText
'    If Trim(mText) <> "" Then
'        Frm.Caption = IIf(Arabic, mArabicCaption, mEnglishCaption) & TranslateText(mText, Arabic)
'    End If
'    '------------------------
'    Dim mIndexArr As Long
'    For Each Ctr In Frm.Controls
'
'        If (TypeOf Ctr Is Label) _
'           Or (TypeOf Ctr Is CheckBox) _
'           Or (TypeOf Ctr Is XtremeSuiteControls.CheckBox) _
'           Or (TypeOf Ctr Is OptionButton) _
'           Or (TypeOf Ctr Is RadioButton) _
'           Or (TypeOf Ctr Is Frame) _
'           Or (TypeOf Ctr Is XtremeSuiteControls.PushButton) _
'           Or (TypeOf Ctr Is XtremeSuiteControls.PushButton) Then
'            '------------------------
'            mIndexArr = FindIndex(Frm, Ctr)
'            If mIndexArr = 69 Then
'                xx = xx
'            End If
'
'            s = "SELECT * FROM Translations WHERE ControlName = N'" & Trim(Ctr.Name) & "' AND ControlIndex = N'" & IIf(mIndexArr = -99, "", mIndexArr) & "' AND FormName = N'" & Trim(Frm.Name) & "'"
'            Set rsDummy = New ADODB.Recordset
'            rsDummy.Open s, Cn, adOpenStatic
'            If Not rsDummy.EOF Then
'                If Trim(rsDummy!English & "") <> "" Then
'                    xx = xx
'                End If
'                mText = IIf(Trim(rsDummy!English & "") <> "", Trim(rsDummy!English & ""), Trim(Ctr.Caption))
'                If rsDummy!IsVisible Then
'                    Ctr.Visible = False
'
'                End If
'                Ctr.Caption = mText
'            Else
'                mText = Trim(Ctr.Caption)
'            End If
'            rsDummy.Close
'            '    Ctr.Caption = TranslateText(mText, Arabic)
'            '---------------------------
'            ''            If TypeOf Ctr Is XPFrame30 Then
'            ''                Ctr.Alignment = IIf(Arabic, 2, 0)
'            '       Else
'
'            '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'            'If Ctr.Alignment <> 2 Then Ctr.Alignment = IIf(Arabic, 1, 0)
'            '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'            '       End If
'
'        ElseIf typename(Ctr) = "CtlCostCenters" Or typename(Ctr) = "CtlBalancesPeriods" Then
'            '--------------------------------
'            TranslateForm Ctr, Arabic
'        ElseIf TypeOf Ctr Is SSTab Then
'            '--------------------------------
'            '            mtab = Ctr.Tab
'            '            RTLTree Ctr    ' WillChange The Direction of the Tab Control
'            '            '------------
'            '            For j = 0 To Ctr.Tabs - 1
'            '                mText = Trim(Ctr.TabCaption(j))
'            '                Ctr.TabCaption(j) = TranslateText(mText, Arabic)
'            '            Next
'            '            '----------------
'            '            If mtab = 0 Then
'            '                Ctr.Tab = 1
'            '            Else
'            '                Ctr.Tab = 0
'            '            End If
'            '            Ctr.Tab = mtab    ' To Redraw the Tab Contents
'        ElseIf TypeOf Ctr Is VSFlexGrid Then
'            '--------------------------------
'            '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'            ''            For r = 0 To Ctr.FixedRows - 1
'            ''                For j = 0 To Ctr.Cols - 1
'            ''                    mText = Trim(Ctr.TextMatrix(r, j))
'            ''                    Ctr.TextMatrix(r, j) = TranslateText(mText, Arabic)
'            ''                Next
'            ''            Next
'            ''
'            ''            For r = 0 To Ctr.Rows - 1
'            ''                For j = 0 To Ctr.Cols - 1
'            ''                    If Not (r > 0 And j > 0) Then
'            ''                        mText = Trim(Ctr.TextMatrix(r, j))
'            ''
'            ''                        If Not IsNumeric(mText) Then
'            ''                            Ctr.TextMatrix(r, j) = TranslateText(mText, Arabic)
'            ''                        End If
'            ''                    End If
'            ''                Next
'            ''            Next
'            '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'        ElseIf TypeOf Ctr Is TextBox Then
'            '--------------------------------
'            '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'            'If Ctr.Alignment <> 2 Then Ctr.Alignment = IIf(Arabic, 1, 0)
'            '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'
'            If Ctr.DataField <> "" Then
'                s = "SELECT * FROM Translations WHERE ControlName = N'" & Trim(Ctr.DataField) & "' AND ControlIndex = N'" & Trim(Ctr.DataMember) & "' AND FormName = N'" & Trim(Frm.Name) & "'"
'                Set rsDummy = New ADODB.Recordset
'                rsDummy.Open s, Cn, adOpenStatic
'                If Not rsDummy.EOF Then
'
'                    If rsDummy!IsVisible Then
'                        Ctr.Visible = False
'
'                    End If
'                End If
'            End If
'
'        ElseIf TypeOf Ctr Is PictureBox Then
'            '--------------------------------
'
'            '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'            ' «·Ì„Ì‰ Ì ÕÊ· ≈·Ï Ì”«— Ê«·⁄ﬂ”
'            '            If Ctr.Align = 3 Then
'            '                Ctr.Align = 4
'            '            ElseIf Ctr.Align = 4 Then
'            '                Ctr.Align = 3
'            '            End If
'            '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'            '------------------------
'        ElseIf TypeOf Ctr Is MSChart Then    ' SaMi 1/11/2009
'            '--------------------------------
'            For i = 1 To 4
'                Ctr.Column = i
'                mText1 = Trim(Ctr.ColumnLabel)
'                Ctr.ColumnLabel = TranslateText(mText, Arabic)
'            Next
'            '****************************************************
'            mText2 = Trim(Ctr.RowLabel)
'            Ctr.RowLabel = TranslateText(mText2, Arabic)
'            '************************************************
'        End If
'        '-------------------------------------
'        ' Change Control Direction
'
'        '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'        '        If Not TypeOf Ctr.Container Is SSTab Then    ' Problem In SSTab as a Container
'        '            If TypeOf Ctr.Container Is Form Then
'        '                Ctr.Left = Ctr.Container.ScaleWidth - Ctr.Left - Ctr.Width  '- IIf(Arabic, 0, 30)
'        '            Else
'        '                Ctr.Left = Ctr.Container.Width - Ctr.Left - Ctr.Width    '- IIf(Arabic, 0, 30)
'        '            End If
'        '        End If
'        '        '--------------------------------
'        '        Ctr.RightToLeft = Arabic    ' Some Controls Does not Support This Property, and some is Read Only
'        '⁄‰œ «· —Ã„… ‰—Ã⁄ «·ﬂÊœ œÂ »«–‰ «··Â
'NextControl:
'    Next
'    '--------------------------------
'    'ChangeControlsDirection Frm
'    ' ********************************
'    Exit Sub
'EH:
'    MsgBox MyErrorHandler(Err)
'    Resume Next
'End Sub
'
'Public Function TranslateText(ByVal mText As String, _
'                              Optional ByVal Arabic As Boolean = False) As String
'    If mText = "" Then
'        Exit Function
'    End If
'    '-------------------------------
'    TranslateText = mText
'    '-------------------------------
'    If Arabic Then
'        Cond = "English=N'" & Trim(mText) & "'"
'    Else
'        Cond = "Arabic=N'" & Trim(mText) & "'"
'    End If
'
'    s = "Select * from Translations where " & Cond
'    Dim Translations As New ADODB.Recordset
'    Translations.Open s, Cn, adOpenStatic, adLockReadOnly
'    '------------------------
'    If Not Translations.EOF Then
'        If Arabic Then
'            TranslateText = Trim(Translations!Arabic & "")
'        Else
'            TranslateText = IIf(Trim(Translations!English & "") <> "", Trim(Translations!English & ""), Trim(Translations!Arabic & ""))
'        End If
'    End If
'End Function
'
'Public Sub RTLTree(TV As Control, Optional RTL As Integer = 0)
'
'    Dim TreeStyle As Long
'    TreeStyle = GetWindowLong(TV.hWnd, GWL_EXSTYLE)
'    If RTL > 0 Then
'        If (TreeStyle And &H400000) = &H400000 And RTL = 1 Then
'            Exit Sub
'        End If
'        If (TreeStyle And &H400000) = 0 And RTL = 2 Then
'            Exit Sub
'        End If
'    End If
'    SetWindowLong TV.hWnd, GWL_EXSTYLE, TreeStyle Xor &H400000
'    'SetWindowLong TV.hWnd, GWL_EXSTYLE, TreeStyle Xor &H2& Xor ES_MULTILINE
'    'x = SendMessage(TV.hWnd, EM_SETALIGN, ES_RIGHT + ES_MULTILINE, 0)
'End Sub
'
'Private Function FindIndex(ByRef F As Form, ByRef ctl As Control) As Integer
'    Dim ctlTest As Control
'    For Each ctlTest In F.Controls
'        If (ctlTest.Name = ctl.Name) And (Not (ctlTest Is ctl)) Then
'            'if the object is the same name but is not the same object we can assume it is a control array
'            FindIndex = ctl.Index
'            Exit Function
'        End If
'    Next
'    'if we get here then no controls on the form have the same name so can't be a control array
'    FindIndex = -99
'End Function
'
'Public Function GetBOFFromNatioanlID(MyNumber As Variant, MyTest As Byte) As Date
'
'    Dim MyProvinces As Variant
'
'    Dim r           As Integer
'
'    Dim yy          As String
'
'    Dim Ty          As String * 1
'
'    Dim d           As String * 2, m As String * 2, Y As String * 2, X As String * 2, xx As String * 2
'
'    '==============================================
'
'    '==============================================
'
'    GetBOFFromNatioanlID = Date
'
'    On Error GoTo 1
'
'    If Len(Trim(MyNumber)) = 0 Then
'
'        GoTo 1
'
'    End If
'
'    If Not IsNumeric(MyNumber) Or Len(MyNumber) <> 14 Then
'
'        ' GetBOFFromNatioanlID = "Error_MyNumber"
'
'        GoTo 1
'
'    End If
'
'    If MyTest = 1 Then
'
'        d = mId(MyNumber, 6, 2)
'
'        m = mId(MyNumber, 4, 2)
'
'        Y = mId(MyNumber, 2, 2)
'
'        Ty = left(MyNumber, 1)
'
'        Select Case Ty
'
'            Case "2": yy = Y
'
'            Case "3": yy = "20" & Y
'
'            Case Else: yy = ""
'
'        End Select
'
'        If yy <> "" Then GetBOFFromNatioanlID = DateSerial(yy, m, d)
'
'    ElseIf MyTest = 2 Then
'
'        If left(right(MyNumber, 2), 1) Mod 2 = 1 Then _
'           yy = "–ﬂ—" Else yy = "«‰ÀÏ"
'
'        GetBOFFromNatioanlID = yy
'
'    ElseIf MyTest = 3 Then
'
'        X = mId(MyNumber, 8, 2)
'
'        For r = LBound(MyProvinces) To UBound(MyProvinces)
'
'            xx = MyProvinces(r)
'
'            If X = xx Then
'
'                GetBOFFromNatioanlID = right(MyProvinces(r), Len(MyProvinces(r)) - 3)
'
'                Exit For
'
'            End If
'
'        Next
'
'    End If
'
'1:
'
'End Function
'

Public Function checkCustomerdata(CUSTID As Integer, Optional invocevalue As Double, Optional Invoicetype As Integer, Optional Dccurrency As String, Optional ByRef Export As Integer) As Boolean
If SystemOptions.ApplyEinvoice = False Then checkCustomerdata = True: Exit Function
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select * from TblCustemers where cusid=" & CUSTID
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
checkCustomerdata = True
 If CUSTID = 2 Then Exit Function
Dim VATNO, CustGID, BuildingNumber, StreetName, PostalZone, Id700, IdentificationCode As String
Dim CityID, GovernmentID As String
 Dim creditlocked As Integer
        If Dccurrency = "" Then Dccurrency = "SAR"
        If Len(Dccurrency) < 3 Then
         msgstr = msgstr & "ﬂÊœ «·⁄„·… ÌÃ» «‰ ÌﬂÊ‰ 3 Œ«‰«  ÿ»ﬁ« ·„ ÿ·»«  «·«Ì“Ê ": MsgBox msgstr, , vbCritical: checkCustomerdata = False: Exit Function
        
        End If
 

 If VATNO = "" Or VATNO = 0 Then VATNO = "N/A"
 
 If CUSTID = 2 And Invoicetype = 0 Then
 'msgstr = msgstr & "     ·«Ì„ﬂ‰ ⁄„·   ›« Ê—… ÷—Ì»Ì…  ·⁄„Ì· ‰ﬁœÌ ·⁄œ„ ÊÃÊœ »Ì«‰« … «·÷—Ì»Ì… «·”Ã·-«·—ﬁ„ «·÷—Ì»Ì-«·⁄‰Ê«‰ «·Êÿ‰Ì":  MsgBox msgstr, vbCritical:  checkCustomerdata = False: Exit Function
 End If
 
If Rs3.RecordCount > 0 Then
 
 VATNO = Round(val(Rs3!VATNO & ""))
 If VATNO = "" Or VATNO = 0 Then VATNO = "N/A"
 CustGID = IIf(IsNull(Rs3("CustGID").value), "", Rs3("CustGID").value)
 If CustGID = "" Or CustGID = 0 Then CustGID = "N/A"
 BuildingNumber = IIf(IsNull(Rs3("BuildingNumber").value), "", Rs3("BuildingNumber").value)
 StreetName = IIf(IsNull(Rs3("StreetName").value), "", Rs3("StreetName").value)
 CityID = IIf(IsNull(Rs3("CityID").value), 0, Rs3("CityID").value)
 GovernmentID = IIf(IsNull(Rs3("GovernmentID").value), "", Rs3("GovernmentID").value)
 PostalZone = IIf(IsNull(Rs3("PostalZone").value), "", Rs3("PostalZone").value)
 IdentificationCode = IIf(IsNull(Rs3("IdentificationCode").value), "", Rs3("IdentificationCode").value)
  creditlocked = IIf(IsNull(Rs3("creditlocked").value), 0, Rs3("creditlocked").value)
  Id700 = IIf(IsNull(Rs3("Id700").value), "", Rs3("Id700").value)
 'WAEL
  Export = IIf(IsNull(Rs3("export").value), 0, Rs3("export").value)
  
   If Invoicetype = 0 And (creditlocked = 1 Or CUSTID = 2) Then        '⁄„Ì· ‰ﬁœÌ ·«Ì Ì„ﬂ‰ ⁄„· ·Â ›« Ê—… ÷—Ì»Ì…
 msgstr = msgstr & "  ⁄„Ì· ‰ﬁœÌ ·«Ì Ì„ﬂ‰ ⁄„· ·Â ›« Ê—… ÷—Ì»Ì… ·«»œ «‰  ﬂÊ‰ ›« Ê—… „»”ÿ…":  MsgBox msgstr, vbCritical:  checkCustomerdata = False: Exit Function
 End If

 If creditlocked = 0 And invocevalue >= 1000 And Invoicetype = 1 And CUSTID <> 2 Then  '⁄„Ì· «Ã· ·«Ì„ﬂ‰ ⁄„· ·Â ›« Ê—… ÷—Ì»Ì… „»”ÿ… »«ﬂÀ— „‰ 1000 —Ì«·
 msgstr = msgstr & " ›« Ê—… B2B ·«Ì„ﬂ‰ ⁄„· ·Â ›« Ê—… ÷—Ì»Ì… „»”ÿ… »«ﬂÀ— „‰ 1000 —Ì«· ":  MsgBox msgstr, vbCritical:  checkCustomerdata = False: Exit Function
 End If
 
 If (Len(VATNO) <> 15 Or mId(VATNO, 15, 1) <> 3) And CUSTID <> 2 And Invoicetype = 0 And Id700 = "" And VATNO <> "N/A" Then
 msgstr = msgstr & " «·—ﬁ„ «·÷—Ì»Ì „ÿ·Ê» 15 Œ«‰…  ›Ì „·› «·⁄„Ì· ":  MsgBox msgstr, vbCritical:  checkCustomerdata = False: Exit Function
 
 End If

 If Len(CustGID) <> 10 And CUSTID <> 2 And Invoicetype = 0 And CustGID <> "N/A" Then
 msgstr = msgstr & " «·”Ã· «· Ã«—Ì «·÷—Ì»Ì „ÿ·Ê» 10 Œ«‰«   ›Ì „·› «·⁄„Ì· ": MsgBox msgstr, vbCritical:  checkCustomerdata = False: Exit Function
 End If
 
 
  If Len(BuildingNumber) <> 4 And CUSTID <> 2 And Invoicetype = 0 And CustGID <> "N/A" Then
 msgstr = msgstr & "   —ﬁ„ «·„Ì‰Ì   „ÿ·Ê» ›Ì „·› «·⁄„Ì· 4 Œ«‰«  ": MsgBox msgstr, vbCritical:  checkCustomerdata = False: Exit Function
 End If
 
  If StreetName = "" And CUSTID <> 2 And Invoicetype = 0 And CustGID <> "N/A" Then
 msgstr = msgstr & "     «·‘«—⁄ „ÿ·Ê» ›Ì „·› «·⁄„Ì· ": MsgBox msgstr, vbCritical:  checkCustomerdata = False: Exit Function
 End If
 
   If val(CityID) = 0 And CUSTID <> 2 And Invoicetype = 0 And CustGID <> "N/A" Then
 msgstr = msgstr & "     «·ÕÌ „ÿ·Ê» ›Ì „·› «·⁄„Ì· ": MsgBox msgstr, vbCritical:  checkCustomerdata = False: Exit Function
 End If
 
    If val(GovernmentID) = 0 And CUSTID <> 2 And Invoicetype = 0 And CustGID <> "N/A" Then
 msgstr = msgstr & "     «·„œÌ‰… „ÿ·Ê» ›Ì „·› «·⁄„Ì· ": MsgBox msgstr, vbCritical:  checkCustomerdata = False: Exit Function
 End If
  
     If IdentificationCode = "" And CUSTID <> 2 And Invoicetype = 0 And CustGID <> "N/A" Then
 msgstr = msgstr & "     —„“ «·œÊ·…  „ÿ·Ê» ›Ì „·› «·⁄„Ì· ": MsgBox msgstr, vbCritical:  checkCustomerdata = False: Exit Function
 End If
 
 
   If Len(PostalZone) <> 5 And CUSTID <> 2 And Invoicetype = 0 And CustGID <> "N/A" Then
 msgstr = msgstr & "   «·—„“ «·»—ÌœÌ     „ÿ·Ê» ›Ì „·› «·⁄„Ì· 5 Œ«‰«  ": MsgBox msgstr, vbCritical:  checkCustomerdata = False: Exit Function
 End If
 
End If
End Function



Public Function checkCustomerdata2(CUSTID As Integer, Optional invocevalue As Double, Optional Invoicetype As Integer, Optional Dccurrency As String) As Boolean
If SystemOptions.ApplyEinvoice = False Then checkCustomerdata2 = True: Exit Function
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select * from TblCustemers where cusid=" & CUSTID
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
checkCustomerdata2 = True
 
Dim VATNO, CustGID, BuildingNumber, StreetName, PostalZone, IdentificationCode As String
Dim CityID, GovernmentID As String
 Dim creditlocked As Integer
 
        If Len(Dccurrency) < 3 Then
         msgstr = msgstr & "ﬂÊœ «·⁄„·… ÌÃ» «‰ ÌﬂÊ‰ 3 Œ«‰«  ÿ»ﬁ« ·„ ÿ·»«  «·«Ì“Ê ": MsgBox msgstr, , vbCritical: checkCustomerdata2 = False: Exit Function
        
        End If
 

 
 
 If CUSTID = 2 And Invoicetype = 0 Then
 msgstr = msgstr & "     ·«Ì„ﬂ‰ ⁄„·   ›« Ê—… ÷—Ì»Ì…  ·⁄„Ì· ‰ﬁœÌ ·⁄œ„ ÊÃÊœ »Ì«‰« … «·÷—Ì»Ì… «·”Ã·-«·—ﬁ„ «·÷—Ì»Ì-«·⁄‰Ê«‰ «·Êÿ‰Ì":  MsgBox msgstr, vbCritical:  checkCustomerdata2 = False: Exit Function
 End If
 
If Rs3.RecordCount > 0 Then
 
 VATNO = IIf(IsNull(Rs3("VATNO").value), "", Rs3("VATNO").value)
 CustGID = IIf(IsNull(Rs3("CustGID").value), "", Rs3("CustGID").value)
 BuildingNumber = IIf(IsNull(Rs3("BuildingNumber").value), "", Rs3("BuildingNumber").value)
 StreetName = IIf(IsNull(Rs3("StreetName").value), "", Rs3("StreetName").value)
 CityID = IIf(IsNull(Rs3("CityID").value), 0, Rs3("CityID").value)
 GovernmentID = IIf(IsNull(Rs3("GovernmentID").value), "", Rs3("GovernmentID").value)
 PostalZone = IIf(IsNull(Rs3("PostalZone").value), "", Rs3("PostalZone").value)
 IdentificationCode = IIf(IsNull(Rs3("IdentificationCode").value), "", Rs3("IdentificationCode").value)
  creditlocked = IIf(IsNull(Rs3("creditlocked").value), 0, Rs3("creditlocked").value)
  
 
  

 If creditlocked = 0 And invocevalue >= 1000 And Invoicetype = 1 And CUSTID <> 2 Then  '⁄„Ì· «Ã· ·«Ì„ﬂ‰ ⁄„· ·Â ›« Ê—… ÷—Ì»Ì… „»”ÿ… »«ﬂÀ— „‰ 1000 —Ì«·
 msgstr = msgstr & " ›« Ê—… B2B ·«Ì„ﬂ‰ ⁄„· ·Â ›« Ê—… ÷—Ì»Ì… „»”ÿ… »«ﬂÀ— „‰ 1000 —Ì«· ":  MsgBox msgstr, vbCritical:  checkCustomerdata2 = False: Exit Function
 End If
 
 If (Len(VATNO) <> 15 Or mId(VATNO, 15, 1) <> 3) And CUSTID <> 2 And Invoicetype = 0 Then
 msgstr = msgstr & " «·—ﬁ„ «·÷—Ì»Ì „ÿ·Ê» 15 Œ«‰…  ›Ì „·› «·⁄„Ì· ":  MsgBox msgstr, vbCritical:  checkCustomerdata2 = False: Exit Function
 
 End If

 If Len(CustGID) <> 10 And CUSTID <> 2 And Invoicetype = 0 Then
 msgstr = msgstr & " «·”Ã· «· Ã«—Ì «·÷—Ì»Ì „ÿ·Ê» 10 Œ«‰«   ›Ì „·› «·⁄„Ì· ": MsgBox msgstr, vbCritical:  checkCustomerdata2 = False: Exit Function
 End If
 
 
  If Len(BuildingNumber) <> 4 And CUSTID <> 2 And Invoicetype = 0 Then
 msgstr = msgstr & "   —ﬁ„ «·„Ì‰Ì   „ÿ·Ê» ›Ì „·› «·⁄„Ì· 4 Œ«‰«  ": MsgBox msgstr, vbCritical:  checkCustomerdata2 = False: Exit Function
 End If
 
  If StreetName = "" And CUSTID <> 2 And Invoicetype = 0 Then
 msgstr = msgstr & "     «·‘«—⁄ „ÿ·Ê» ›Ì „·› «·⁄„Ì· ": MsgBox msgstr, vbCritical:  checkCustomerdata2 = False: Exit Function
 End If
 
   If val(CityID) = 0 And CUSTID <> 2 And Invoicetype = 0 Then
 msgstr = msgstr & "     «·ÕÌ „ÿ·Ê» ›Ì „·› «·⁄„Ì· ": MsgBox msgstr, vbCritical:  checkCustomerdata2 = False: Exit Function
 End If
 
    If val(GovernmentID) = 0 And CUSTID <> 2 And Invoicetype = 0 Then
 msgstr = msgstr & "     «·„œÌ‰… „ÿ·Ê» ›Ì „·› «·⁄„Ì· ": MsgBox msgstr, vbCritical:  checkCustomerdata2 = False: Exit Function
 End If
  
     If IdentificationCode = "" And CUSTID <> 2 And Invoicetype = 0 Then
 msgstr = msgstr & "     —„“ «·œÊ·…  „ÿ·Ê» ›Ì „·› «·⁄„Ì· ": MsgBox msgstr, vbCritical:  checkCustomerdata2 = False: Exit Function
 End If
 
 
   If Len(PostalZone) <> 5 And CUSTID <> 2 And Invoicetype = 0 Then
 msgstr = msgstr & "   «·—„“ «·»—ÌœÌ     „ÿ·Ê» ›Ì „·› «·⁄„Ì· 5 Œ«‰«  ": MsgBox msgstr, vbCritical:  checkCustomerdata2 = False: Exit Function
 End If
 
End If
End Function



Public Sub ExportToExcel(ByRef Frm As Form, _
                         ByRef G As Object, _
                         Optional Caption As String = "", _
                         Optional ByRef ObjExecel, Optional MainFormName As String = "", Optional mRtl As Integer = 0)

'***********************Khalid
'    If Not isDebugMode Then
'        If IsDisabledExcelShortCut Then
'            Exit Sub
'        End If
'    End If
    '****************************

    On Error GoTo eh
    Screen.MousePointer = vbHourglass
    '--- open new Excel File In memory ---
    Dim ExcelSheet
    Set ExcelSheet = CreateObject("excel.application")
    If Not IsObject(ExcelSheet) Then Exit Sub

    '--- Add new WorkBook ---
    '--- ByDefault contain 3 Woorksheet ---
    '    ExcelSheet.Workbooks.Add
    '
    'For i = 0 To G.Cols - 1
    'G.ColDataType(i) = flexDTStringC
    'Next
    '    '==== FIRST WORKSHEET ===================================
    '    '--- activate the second WorkSheet ---
    'FFF = "c:\~Temp" & Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xls"
    
    
'     For j = 0 To G2.Cols - 1
'        If Not G2.ColHidden(j) Then
'
'        End If
'    Next
    Dim fff
    fff = GetGridFileName(G, MainFormName)
    'FFF = "D:\ddd54dd.xls"
    'FG.saveGrid fff , flexFileExcel,  flexXLSaveFixedRows Or flexXLSaveFixedCols
    G.saveGrid fff, flexFileExcel, flexXLSaveFixedCells Or flexXLSaveRaw Or flexxl
       
       
       

Dim ExcelObj
  '  ExcelSheet.Workbooks.Open FFF
        Set ExcelObj = CreateObject("Excel.Application")
    '
  '  Set ExcelSheet = CreateObject("Excel.Sheet")
  
  
  
   Dim excelApp As Object
    Dim workbook As Object
    Dim FilePath As String
    Dim xlSheet As Object
    '  ÕœÌœ „”«— «·„·› «·–Ì  —Ìœ › ÕÂ
    
    
    ' ≈‰‘«¡ „ÀÌ· ÃœÌœ · ÿ»Ìﬁ Excel
    'et excelApp = CreateObject("Excel.Application")
    
    ' Ã⁄· Excel „—∆Ì ··„” Œœ„
    'excelApp.Visible = True
    'ExcelObj.Visible = True
    'Call UnblockFile(fff)
    ' › Õ «·„·›
    
    
    
    
    
    
    
'
'   Set workbook = ExcelObj.Workbooks.Open(fff, , , , , , , , , , , , , , True)
'   Set xlSheet = workbook.Sheets(1)
'
  ' xlSheet.UsedRange.Copy
  
  
 
  
  If ClipboardSetFile(CStr(fff)) Then
        MsgBox " „ ‰”Œ «·„·› ≈·Ï «·Õ«›Ÿ… »‰Ã«Õ!"
    Else
        MsgBox "›‘· ‰”Œ «·„·› ≈·Ï «·Õ«›Ÿ…."
    End If
   
   
  mFileName = fff
  'MsgBox fff
  Screen.MousePointer = vbDefault
  Exit Sub

    

eh:

 MsgBox "ÕœÀ Œÿ√ √À‰«¡ „Õ«Ê·… › Õ «·„·›: " & Err.Description, vbCritical, "Error"
    
    '  ‰ŸÌ› «·ﬂ«∆‰«  ›Ì Õ«·… «·Œÿ√
    If Not excelApp Is Nothing Then
        excelApp.Quit
        Set excelApp = Nothing
    End If
    
    
    Screen.MousePointer = vbDefault
    MsgBox "Â‰«ﬂ «Œÿ«¡ ° —»„« »”»» ⁄œ„ ÊÃÊœ «·«ﬂ”Ì· √Ê «·„·› „› ÊÕ „”»ﬁ«"

End Sub





Function ConvertDateToString(ByVal dt As Date) As String
    '  ÕÊÌ· «· «—ÌŒ ≈·Ï ”·”·… ‰’Ì… »’Ì€… "YYYY-MM-DD"
    ConvertDateToString = Format(dt, "yyyy-mm-dd")
End Function

Public Function GetMainForm(ByVal obj) As String
    Dim n As String
    On Error Resume Next
    n = obj.Container.Name

    If n = "" Then
        GetMainForm = obj.Name
    Else
        GetMainForm = GetMainForm(obj.Container)
    End If
End Function

Public Function GetGridFileName(ByVal G As Object, Optional MainFormName As String = "") As String
    Dim GlobalGridName As String
    Dim indexs As String
    Dim MainContainerName As String

    On Error Resume Next
    indexs = 0 'G.index
Dim dateString As String
dateString = ConvertDateToString(Now) ' ” Õ’· ⁄·Ï «· «—ÌŒ «·Õ«·Ì ﬂ‹ "YYYY-MM-DD"
If MainFormName <> "" Then
    dateString = ""
End If
    MainContainerName = GetMainForm(G.Container.Caption)
    GlobalGridName = MainContainerName & "\" & G.Name & indexs & MainFormName
    'GlobalGridName = "POS-P"
    GetGridFileName = App.path & GlobalGridName & "" & dateString & ".xls"
    'GetGridFileName = App.path & GlobalGridName & ".xls"

End Function


Function StartsWithKeywords(ByVal Name As String) As Boolean
    '  ÕÊÌ· «·‰’ ≈·Ï Õ«·… ’€Ì—… ·⁄„· „ﬁ«—‰… €Ì— Õ”«”… ·Õ«·… «·Õ—Ê›
    Dim lowerName As String
    lowerName = LCase(Name)
    
    ' «· Õﬁﬁ „‰ √‰ «·‰’ Ì»œ√ »√Ì „‰ «·ﬂ·„«  «·„Õœœ…
    If left(lowerName, 2) = "fg" Or left(lowerName, 4) = "grid" Or left(lowerName, 3) = "grd" Then
        StartsWithKeywords = True
    Else
        StartsWithKeywords = False
    End If
End Function


