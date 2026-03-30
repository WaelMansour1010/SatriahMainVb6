VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCustemerReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function CustemerData(CusID As Long, _
                             Optional customertype As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Titalname As String
    On Error GoTo ErrTrap

    If CusID = Null Then
        Exit Function
    End If

    'StrSQL = "select * From TblCustemers where CusID=" & CusID
    If Screen.ActiveForm.name = "FrmOtherCustomers" Then

        Titalname = "ÊÞÇÑíÑ ÈíÇäÇÊ ãÞÇæáí ÇáÈÇØä "
        
    End If

    If Screen.ActiveForm.name = "FrmCustemers" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Titalname = "ÊÞÇÑíÑ ÇáÚãáÇÁ "
        Else
            Titalname = " Customers Data "
        End If
    End If

    If customertype = 2 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            Titalname = "ÊÞÇÑíÑ ÇáãæÑÏíä "
        Else
            Titalname = " Supplier Data "
        End If
    End If

    StrSQL = "SELECT     dbo.ACCOUNTS.Account_Serial AS Expr1, dbo.TblCustemers.*"
    StrSQL = StrSQL & "  FROM         dbo.TblCustemers LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code"
    
    StrSQL = StrSQL & "  where CusID>2 and  Type=" & customertype

    If SystemOptions.UserInterface = ArabicInterface Then
        'StrSQL = StrSQL & "  ORDER BY  "
    Else

    End If

    Dim Reportpath As String

    If SystemOptions.UserInterface = ArabicInterface Then
        Reportpath = App.path & "\Reports\AllCustomerData.rpt"
    Else
        Reportpath = App.path & "\Reports\AllCustomerDatae.rpt"
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If Dir(Reportpath) = "" Then
        Msg = "ãáÝ ÇáÊÞÑíÑ ÛíÑ ãæÌæÏ..!!" & Chr(13)
        Msg = "ÈÑÌÇÁ ÇáÊÃßÏ ãä æÌæÏ åÐÇ ÇáãáÝ Ýì ãÓÇÑ ÇáÈÑäÇãÌ" & Chr(13)
        Msg = "\Reports\CustemerReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(Reportpath)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = Titalname
        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.Show
    FrmReport.TxtPath = Reportpath
    
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "ÚÝæÇð " & Chr(13) & "áÇíãßä ØÈÇÚÉ ÇáÊÞÑíÑ" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Sub ShowCustsBalances(IntReportType As Integer, _
                             m_PrintTarget As PrintTarget)

    'IntReportType=0 All
    'IntReportType=1 Customers
    'IntReportType=2 Suppliers

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

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\CustomersBalances.rpt"
    Else
        StrFileName = App.path & "\Reports\CustomersBalances_Eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "ãáÝ ÇáÊÞÑíÑ ÛíÑ ãæÌæÏ..!!" & Chr(13)
        Msg = Msg + "ÈÑÌÇÁ ÇáÊÃßÏ ãä æÌæÏ åÐÇ ÇáãáÝ Ýì ãÓÇÑ ÇáÈÑäÇãÌ" & Chr(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "Select TblCustemers.CusID, TblCustemers.CusName, TblCustemers.Type," & "TblCustemers.OpenBalance,TblCustemers.OpenBalanceType,TblCustemers.OpenBalanceDate," & "QryCustomersBalance.CustomerAccount "
        StrSQL = StrSQL + " FROM dbo.TblCustemers LEFT JOIN "
        StrSQL = StrSQL + " QryCustomersBalance ON TblCustemers.CusID =QryCustomersBalance.CusID "
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "Select TblCustemers.CusID, TblCustemers.CusName, TblCustemers.Type," & "TblCustemers.OpenBalance,TblCustemers.OpenBalanceType,TblCustemers.OpenBalanceDate," & "QryCustomersBalance.CustomerAccount "
        StrSQL = StrSQL + " FROM TblCustemers LEFT JOIN "
        StrSQL = StrSQL + " QryCustomersBalance ON TblCustemers.CusID =QryCustomersBalance.CusID "
    End If

    Set rs = New ADODB.Recordset

    If IntReportType = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrReportTitle = "ÊÞÑíÑ ÈÃÑÕÏÉ ÇáÚãáÇÁ æÇáãæÑÏíä"
        Else
            StrReportTitle = "Report By Customers And Suppliers Accounts"
        End If

        StrSQL = StrSQL + " Where (TblCustemers.Type=1 OR TblCustemers.Type=2)"
    ElseIf IntReportType = 1 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrReportTitle = "ÊÞÑíÑ ÈÃÑÕÏÉ ÇáÚãáÇÁ"
        Else
            StrReportTitle = "Report By Customers Accounts"
        End If

        StrSQL = StrSQL + " Where TblCustemers.Type=1"
    ElseIf IntReportType = 2 Then
        StrReportTitle = "ÊÞÑíÑ ÈÃÑÕÏÉ ÇáãæÑÏíä"
        StrSQL = StrSQL + " Where TblCustemers.Type=2"
    ElseIf IntReportType = 3 Then
        StrReportTitle = "ÊÞÑíÑ ÈÃÑÕÏÉ ÇáãÊÚáÞÇÊ"
        StrSQL = StrSQL + " Where TblCustemers.Type=3"
    ElseIf IntReportType = 4 Then
        StrReportTitle = "ÊÞÑíÑ ÈÃÑÕÏÉ (ÇáÚãáÇÁ æÇáãæÑÏíä æÇáãÊÚáÞÇÊ)"
        StrSQL = StrSQL + " Where (TblCustemers.Type=1 OR TblCustemers.Type=2 OR TblCustemers.Type=3)"
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        rs.Close
        Set rs = Nothing
        Exit Sub
    ElseIf rs.RecordCount = 0 Then
        Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ..!!"
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
        xReport.reporttitle = StrReportTitle
        Screen.MousePointer = vbDefault
    End If

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_PrintTarget, StrReportTitle, , , , StrFileName

    Exit Sub
ErrTrap:
    Msg = "ÚÝæÇð " & Chr(13) & "áÇíãßä ØÈÇÚÉ ÇáÊÞÑíÑ" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Sub

Public Function ClientAccount(LngCusID As Long, _
                              fromdate As Variant, _
                              todate As Variant, _
                              PrintTo As PrintTarget, _
                              IntStyle As Integer)

    'ßÔÝ ÍÓÇÈ Úãíá
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions  As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim StrSQL As String
    Dim SngCusBegainAccount As Single
    Dim StrFileName As String
    Dim cRportViewer As ClsReportViewer
    Dim StrReportTitle As String

    On Error GoTo ErrTrap

    If IntStyle = 0 Then
        StrFileName = App.path & "\Reports\Client_AccountStyle2.rpt"
    ElseIf IntStyle = 1 Then
        StrFileName = App.path & "\Reports\Client_Account.rpt"
    ElseIf IntStyle = 2 Then
        StrFileName = App.path & "\Reports\Client_AccountStyle3.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "ãáÝ ÇáÊÞÑíÑ ÛíÑ ãæÌæÏ..!!" & Chr(13)
        Msg = Msg + "ÈÑÌÇÁ ÇáÊÃßÏ ãä æÌæÏ åÐÇ ÇáãáÝ Ýì ãÓÇÑ ÇáÈÑäÇãÌ" & Chr(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    If IntStyle = 0 Or IntStyle = 1 Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "Select * From QryCustomerBalance Where CusID=" & LngCusID & ""

            If Not (IsNull(fromdate)) Then
                SngCusBegainAccount = GetCustomerAccount(LngCusID, True, fromdate)
                StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(CDate(fromdate), True) & ""
            Else
                SngCusBegainAccount = GetCustomerAccount(LngCusID, False)
            End If

            If Not (IsNull(todate)) Then
                StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(CDate(todate), True) & ""
            End If

        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT * FROM dbo.QryCustomerBalance(" & LngCusID & ") QryCustomerBalance Where (CusID =" & LngCusID & ")"

            If Not (IsNull(fromdate)) Then
                SngCusBegainAccount = GetCustomerAccount(LngCusID, True, fromdate)
                StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(CDate(fromdate), True) & ""
            Else
                SngCusBegainAccount = GetCustomerAccount(LngCusID, False)
            End If

            If Not (IsNull(todate)) Then
                StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(CDate(todate), True) & ""
            End If
        End If

    ElseIf IntStyle = 2 Then
        StrSQL = "SELECT     TOP 100 PERCENT QryCustomerBalance.NoteID, QryCustomerBalance.NoteDate," & "QryCustomerBalance.NoteSerial1, QryCustomerBalance.NoteType,QryCustomerBalance.NotesTypeName," & "QryCustomerBalance.Note_Value, QryCustomerBalance.DueDate, QryCustomerBalance.CusID," & "QryCustomerBalance.Transaction_ID, QryCustomerBalance.Transaction_Serial, QryCustomerBalance.Transaction_Date," & "QryCustomerBalance.TransactionTypeName,QryCustomerBalance.UserName, QryCustomerBalance.CreditOrDebit," & "QryCustomerBalance.CusName,QryCustomerBalance.Remark, QryCustomerBalance.TransTotal,dbo.Transaction_Details.Item_ID," & "dbo.TblItems.ItemCode, dbo.TblItems.ItemName,SUM(dbo.Transaction_Details.Quantity) AS SumQty," & "(dbo.Transaction_Details.Price) AS SumPrice "
        StrSQL = StrSQL + " FROM dbo.TblItems INNER JOIN dbo.Transaction_Details ON dbo.TblItems.ItemID =" & " dbo.Transaction_Details.Item_ID RIGHT OUTER JOIN dbo.QryCustomerBalance(" & LngCusID & ")" & "QryCustomerBalance ON dbo.Transaction_Details.Transaction_ID= QryCustomerBalance.Transaction_ID " & "WHERE     (QryCustomerBalance.CusID = " & LngCusID & ")"

        If Not (IsNull(fromdate)) Then
            SngCusBegainAccount = GetCustomerAccount(LngCusID, True, fromdate)
            StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(CDate(fromdate), True) & ""
        Else
            SngCusBegainAccount = GetCustomerAccount(LngCusID, False)
        End If

        If Not (IsNull(todate)) Then
            StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(CDate(todate), True) & ""
        End If

        StrSQL = StrSQL + "  GROUP BY QryCustomerBalance.NoteID, QryCustomerBalance.NoteDate, QryCustomerBalance.NoteSerial1," & "QryCustomerBalance.NoteType,QryCustomerBalance.NotesTypeName, QryCustomerBalance.Note_Value, QryCustomerBalance.DueDate," & "QryCustomerBalance.CusID,QryCustomerBalance.Transaction_ID,QryCustomerBalance.Transaction_Serial,QryCustomerBalance.Transaction_Date," & "QryCustomerBalance.TransactionTypeName, QryCustomerBalance.UserName, QryCustomerBalance.CreditOrDebit," & "QryCustomerBalance.CusName,QryCustomerBalance.Remark, QryCustomerBalance.TransTotal, dbo.Transaction_Details.Item_ID," & "dbo.TblItems.ItemCode,dbo.TblItems.ItemName,dbo.Transaction_Details.Price " & "ORDER BY QryCustomerBalance.NoteID, QryCustomerBalance.Transaction_ID"
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "áÇ ÊæÌÏ Ãí ÈíÇäÇÊ ááÚÑÖ"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        StrReportTitle = "ßÔÝ ÍÓÇÈ -" & rs("CusName").value & ""
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CDbl(SngCusBegainAccount)
        xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(fromdate), CDate("01/01/1000"), fromdate)
        xReport.ParameterFields(6).AddCurrentValue IIf(IsNull(todate), CDate("01/01/1000"), todate)
        xReport.reporttitle = StrReportTitle
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set cRportViewer = New ClsReportViewer
        cRportViewer.FireReport xReport, WindowTarget, xReport.reporttitle, , , , StrFileName
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "ÚÝæÇð " & Chr(13) & "áÇíãßä ØÈÇÚÉ ÇáÊÞÑíÑ" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Function ShowCusTransactions(LngCusID As Long, _
                                    StrCusName As String, _
                                    IntTransType As Integer, _
                                    Optional DateFrom As Variant = Null, _
                                    Optional DateTo As Variant = Null, _
                                    Optional BolShowNotes As Boolean = False)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cReportViewer As ClsReportViewer
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RsNotes As ADODB.Recordset
    Dim StrSQL  As String
    Dim StrTitle As String
    Dim DblCusAccount As Double
    Dim StrFileName As String
    Dim Msg As String

    On Error GoTo ErrTrap

    StrFileName = App.path & "\Reports\CustomerTransactions.rpt"
    '--------------------------------------
    'ãÚÑÝÉ ÞíãÉ ÇáÑÕíÏ ÇáÍÇáì ááÚãíá Çæ ÇáãæÑÏ
    '--------------------------------------
    DblCusAccount = GetCustomerAccount(LngCusID, True)
    StrSQL = "Select * From RptCusTransactions "
    StrSQL = StrSQL + " Where CusID=" & LngCusID & ""
    StrSQL = StrSQL + " AND Transaction_Type=" & IntTransType & ""

    If Not (IsNull(DateFrom)) Then
        StrSQL = StrSQL + " AND Transaction_Date >=" & SQLDate(CDate(DateFrom), True) & ""
    End If

    If Not (IsNull(DateTo)) Then
        StrSQL = StrSQL + " AND Transaction_Date <=" & SQLDate(CDate(DateTo), True) & ""
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "áÇ ÊæÌÏ ÈíÇäÇÊ ááÚÑÖ...!!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    If BolShowNotes = True Then
        StrTitle = "ÊÞÑíÑ ÈÇáãÈíÚÇÊ ÇáÃÌáÉ æÇáÊÍÕíáÇÊ ÇáÞÇÆãÉ ÚáíåÇ"
        Set RsNotes = New ADODB.Recordset
        MySQL = "Select * From QryPostedTransNotes Where NoteType=4 OR NoteType=9"
        RsNotes.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    End If

    If Dir(StrFileName) = "" Then
        Msg = "ãáÝ ÇáÊÞÑíÑ ÛíÑ ãæÌæÏ..!!" & Chr(13)
        Msg = Msg + "ÈÑÌÇÁ ÇáÊÃßÏ ãä æÌæÏ åÐÇ ÇáãáÝ Ýì ãÓÇÑ ÇáÈÑäÇãÌ" & Chr(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)

        If BolShowNotes = False Then
            xReport.Database.SetDataSource rs
        Else
            
            xReport.Database.tables(1).SetDataSource rs
            xReport.Database.tables(2).SetDataSource RsNotes
        End If
        
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        
        If DateFrom = Null Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If

        xReport.ParameterFields(6).AddCurrentValue DblCusAccount
        Screen.MousePointer = vbDefault
    End If

    If IntTransType = 1 Then
        StrTitle = "ÊÞÑíÑ ÈÝæÇÊíÑ ÇáãÔÊÑíÇÊ ( " & StrCusName & " )¡"
    ElseIf IntTransType = 2 Then
        StrTitle = "ÊÞÑíÑ ÈÝæÇÊíÑ ÇáãÈíÚÇÊ ( " & StrCusName & " )¡"
    ElseIf IntTransType = 5 Then
        StrTitle = "ÊÞÑíÑ ÈÝæÇÊíÑ  ãÑÊÌÚ ÇáãÔÊÑíÇÊ ( " & StrCusName & " )¡"
    ElseIf IntTransType = 9 Then
        StrTitle = "ÊÞÑíÑ ÈÝæÇÊíÑ  ãÑÊÌÚ ÇáãÈíÚÇÊ ( " & StrCusName & " )¡"
    End If

    Set cReportViewer = New ClsReportViewer
    xReport.reporttitle = StrTitle
    cReportViewer.FireReport xReport, WindowTarget, StrTitle, , , , StrFileName
    Set cReportViewer = Nothing
    Exit Function
ErrTrap:
    Msg = "ÚÝæÇð " & Chr(13) & "áÇíãßä ØÈÇÚÉ ÇáÊÞÑíÑ" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault

End Function
