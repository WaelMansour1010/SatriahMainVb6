Attribute VB_Name = "ModTaskPanel"
Option Explicit

Public Enum TaskPanelGroupsIDs
    TskGroupSales = 1
    TskGroupPurchase = 2
    TskGroupReSales = 3
    TskGroupRePurchase = 4
    TskGroupBoxes = 5
    TskGroupExpenses = 6
    TskGroupManCompsAsmply = 7
    TskGroupManSearch = 8
    TskGroupNotesRecivable = 9
    TskGroupNotesPayable = 10
End Enum

Public Enum TaskPnlTransGroupItemsIDs
    TskItemDayValTrans = 1
    TskItemWeekValTrans = 2
    TskItemMonthValTrans = 3
End Enum

Public Sub SetupGrid(FG As Object, _
                     Optional IntType As Integer = 0)
    Dim i As Integer
    Dim GrdBack  As New ClsBackGroundPic
    FG.WallPaper = GrdBack.Picture
    FG.ScrollBars = flexScrollBarBoth

    If IntType = 0 Then

        With FG
            .Font.name = "Tahoma"
            .Rows = .FixedRows + 5
            .Cols = 5
            .Editable = flexEDNone
            .ExtendLastCol = True
        
            .TextMatrix(0, 1) = "TransID"
            .ColKey(1) = "TransID"
            .ColHidden(1) = True

            If SystemOptions.UserInterface = ArabicInterface Then
                FG.RightToLeft = True
                .TextMatrix(0, 0) = "ã"
                .TextMatrix(0, 2) = "ÑÞã ÇáÝÇÊæÑÉ"
                .TextMatrix(0, 3) = "ÇÓã ÇáÚãíá"
                .TextMatrix(0, 4) = "ÞíãÉ ÇáÝÇÊæÑÉ"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                FG.RightToLeft = False
                .TextMatrix(0, 0) = "S"
                .TextMatrix(0, 2) = "Invoice Serial"
                .TextMatrix(0, 3) = "Customer Name"
                .TextMatrix(0, 4) = "Invoice Total"
            End If

        End With

    ElseIf IntType = 1 Then

        With FG
            .Font.name = "Tahoma"
            .Rows = .FixedRows + 5
            .Cols = 5
            .Editable = flexEDNone
            .ExtendLastCol = True
        
            .TextMatrix(0, 1) = "TableID"
            .ColKey(1) = "TableID"
            .ColHidden(1) = True

            If SystemOptions.UserInterface = ArabicInterface Then
                FG.RightToLeft = True
                .TextMatrix(0, 0) = "ã"
                .TextMatrix(0, 2) = "ÑÞã ÇáÝÇÊæÑÉ"
                .TextMatrix(0, 3) = "ÇÓã ÇáÚãíá"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                FG.RightToLeft = False
                .TextMatrix(0, 0) = "S"
                .TextMatrix(0, 2) = "Invoice Serial"
                .TextMatrix(0, 3) = "Customer Name"
            End If

        End With

    End If

    If SystemOptions.UserInterface = ArabicInterface Then

        For i = 0 To FG.Cols - 1
            FG.ColAlignment(i) = flexAlignRightCenter
            FG.FixedAlignment(i) = flexAlignRightCenter
        Next i

    Else

        For i = 0 To FG.Cols - 1
            FG.ColAlignment(i) = flexAlignLeftCenter
            FG.FixedAlignment(i) = flexAlignLeftCenter
        Next i

    End If

    FG.AutoSize 0, FG.Cols - 1, False
End Sub

Public Sub GetTransactionGroup(IntTransType As Integer, _
                               xGroup As TaskPanelGroup, _
                               FG As VSFlex8Ctl.vsFlexGrid, _
                               Optional xChart As Object = Nothing)

    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim DblTemp As Double
    Dim DblCount As Double
    Dim StrTemp As String
    Dim DblCashValue As Double
    Dim DblDueValue As Double
    Dim i  As Integer
    Dim StrGroupTitle As String
    Dim Msg As String

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        '    Exit Sub
    End If

    On Error GoTo hErr

    If IntTransType = 2 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrGroupTitle = "ÇáãÈíÚÇÊ"
        Else
            StrGroupTitle = "Sales"
        End If

    ElseIf IntTransType = 1 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrGroupTitle = "ÇáãÔÊÑíÇÊ"
        Else
            StrGroupTitle = "Purchases"
        End If

    ElseIf IntTransType = 9 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrGroupTitle = "ãÑÊÌÚ ÇáãÈíÚÇÊ"
        Else
            StrGroupTitle = "Retrun Sales"
        End If

    ElseIf IntTransType = 5 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrGroupTitle = "ãÑÊÌÚ ÇáãÔÊÑíÇÊ"
        Else
            StrGroupTitle = "Retrun Purchases"
        End If
    End If

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "Select Count(XX) as CountRes, Sum(XX)as Res,PaymentType" & " From ( "
        StrSQL = StrSQL + " Select dbo.QryOneTransactionTotal(Transaction_ID)as XX "
        StrSQL = StrSQL + ",Transactions.PaymentType From Transactions "
        StrSQL = StrSQL + " Where ( Transactions.Transaction_Type=" & IntTransType
        StrSQL = StrSQL + " AND Month(Transactions.Transaction_Date)=" & Month(Date) & ""
        StrSQL = StrSQL + " AND Day(Transactions.Transaction_Date)=" & Day(Date) & ")"
        StrSQL = StrSQL + " AND Year(Transactions.Transaction_Date)=" & year(Date) & ""
        StrSQL = StrSQL + " )XTable "
        StrSQL = StrSQL + " Group By PaymentType "
        StrSQL = StrSQL + " Order By PaymentType "
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then

        DoEvents
        StrSQL = "Select Count(XX) as CountRes, Sum(XX) as Res,PaymentType " & " From ( "
        StrSQL = StrSQL + " SELECT QryTransactionsTotal.TotalAfterTax AS XX,"
        StrSQL = StrSQL + " QryTransactionsTotal.PaymentType FROM QryTransactionsTotal "
        StrSQL = StrSQL + " Where ( Transactions.Transaction_Type=" & IntTransType
        StrSQL = StrSQL + " AND Month(Transactions.Transaction_Date)=" & Month(Date) & ""
        StrSQL = StrSQL + " AND Day(Transactions.Transaction_Date)=" & Day(Date) & ")"
        StrSQL = StrSQL + " AND Year(Transactions.Transaction_Date)=" & year(Date) & ""
        StrSQL = StrSQL + ")XTable "
        StrSQL = StrSQL + "Group By PaymentType  Order By PaymentType "
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch + adAsyncExecute

    Do While rs.State = adStateExecuting

        If SystemOptions.BolStopUpdateTask = True Then
            rs.Cancel
            Set rs = Nothing
            SystemOptions.BolUpdateTaskInProgress = False
            Exit Sub
        End If

        DoEvents
    Loop

    DoEvents

    If Not (rs.BOF Or rs.EOF) Then
        DblTemp = 0
        DblCount = 0

        For i = 1 To rs.RecordCount

            If rs("PaymentType").value = 0 Then
                DblCashValue = IIf(IsNull(rs("Res").value), 0, rs("Res").value)
            Else
                DblDueValue = IIf(IsNull(rs("Res").value), 0, rs("Res").value)
            End If

            DblTemp = DblTemp + IIf(IsNull(rs("Res").value), 0, rs("Res").value)
            DblCount = DblCount + IIf(IsNull(rs("CountRes").value), 0, rs("CountRes").value)
            rs.MoveNext
        Next i

        xGroup.Items(2).Caption = "ÅÌãÇáì ÞíãÉ " & StrGroupTitle & " : " & DblTemp
        xGroup.Items(1).Caption = "ÚÏÏ ÝæÇÊíÑ Çáíæã : " & DblCount

        xGroup.Items(4).Caption = "ÞíãÉ " & StrGroupTitle & " ÇáäÞÏíÉ : " & DblCashValue
        xGroup.Items(5).Caption = "ÞíãÉ " & StrGroupTitle & " ÇáÃÌáÉ : " & DblDueValue
    End If

    '----------------------------------------------------------

    DoEvents

    '----------------------------------------------------------
    'ÚÑÖ ÇÎÑ 5 ÝæÇÊíÑ
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT  TOP 5 dbo.Transactions.Transaction_ID," & "dbo.Transactions.Transaction_Serial, " & "dbo.QryOneTransactionTotal(dbo.Transactions.Transaction_ID) AS XX ,"
        StrSQL = StrSQL + " dbo.TblCustemers.CusName FROM dbo.Transactions INNER JOIN "
        StrSQL = StrSQL + " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
        StrSQL = StrSQL + " Where ( Transactions.Transaction_Type=" & IntTransType & ""
        StrSQL = StrSQL + " AND Month(Transactions.Transaction_Date)=" & Month(Date) & ""
        StrSQL = StrSQL + " AND Day(Transactions.Transaction_Date)=" & Day(Date) & ")"
        StrSQL = StrSQL + " AND Year(Transactions.Transaction_Date)=" & year(Date) & ""
        StrSQL = StrSQL + " ORDER BY dbo.Transactions.Transaction_ID DESC"
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then

        DoEvents
        StrSQL = "SELECT TOP 5  QryTransactionsTotal.Transaction_ID, QryTransactionsTotal" & ".Transaction_Serial, QryTransactionsTotal.TotalAfterTax AS XX, TblCustemers.CusName" & " FROM TblCustemers INNER JOIN QryTransactionsTotal ON TblCustemers.CusID =" & "QryTransactionsTotal.CusID  "
        StrSQL = StrSQL + " Where (QryTransactionsTotal.Transaction_Type=" & IntTransType
        StrSQL = StrSQL + " AND Month( QryTransactionsTotal.Transaction_Date)=" & Month(Date) & ""
        StrSQL = StrSQL + " AND Day( QryTransactionsTotal.Transaction_Date)=" & Day(Date) & ")"
        StrSQL = StrSQL + " AND Year( QryTransactionsTotal.Transaction_Date)=" & year(Date) & ""
        StrSQL = StrSQL + " ORDER BY  QryTransactionsTotal.Transaction_ID DESC"
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch + adAsyncExecute

    Do While rs.State = adStateExecuting

        If SystemOptions.BolStopUpdateTask = True Then
            rs.Cancel
            rs.Close
            SystemOptions.BolUpdateTaskInProgress = False
            Exit Sub
        End If

        DoEvents
    Loop

    With FG

        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst

            For i = .FixedRows To rs.RecordCount
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
                .TextMatrix(i, 2) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
                .TextMatrix(i, 3) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(i, 4) = IIf(IsNull(rs("XX").value), "", rs("XX").value)
                rs.MoveNext
            Next i

        End If

        .AutoSize 0, .Cols - 1, False
    End With

    '-------------------------------
    'ÇáÑÓã ÇáÈíÇäì ááÍÑßÉ Ýì ÎáÇá ÇáÃÓÈæÚ
    If Not xChart Is Nothing Then
        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "Select  convert(nvarchar(50),Transaction_Date,111)as Transaction_Date,"
            StrSQL = StrSQL + " Sum(TotalAfterTax) As xx "
            StrSQL = StrSQL + " From ( "
            StrSQL = StrSQL + " SELECT  Transaction_Date, TotalAfterTax"
            StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal"
            StrSQL = StrSQL + " Where (Transaction_Type = " & IntTransType & ")"
            StrSQL = StrSQL + " AND  (QryTransactionsTotal.Transaction_Date) >=" & SQLDate(GetWeekStartEND(Date, 0), True) & ""
            StrSQL = StrSQL + " AND  (QryTransactionsTotal.Transaction_Date) <=" & SQLDate(GetWeekStartEND(Date, 1), True) & ""
            StrSQL = StrSQL + " )XTable "
            StrSQL = StrSQL + " Group BY Transaction_Date"
            StrSQL = StrSQL + " Order BY Transaction_Date ASC"
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then

            DoEvents
            StrSQL = "Select  Transaction_Date,"
            StrSQL = StrSQL + " Sum(TotalAfterTax) As xx "
            StrSQL = StrSQL + " From ( "
            StrSQL = StrSQL + " SELECT  Transaction_Date, TotalAfterTax"
            StrSQL = StrSQL + " FROM QryTransactionsTotal"
            StrSQL = StrSQL + " Where (Transaction_Type = " & IntTransType & ")"
            StrSQL = StrSQL + " AND  (QryTransactionsTotal.Transaction_Date) >=" & SQLDate(GetWeekStartEND(Date, 0), True) & ""
            StrSQL = StrSQL + " AND  (QryTransactionsTotal.Transaction_Date) <=" & SQLDate(GetWeekStartEND(Date, 1), True) & ""
            StrSQL = StrSQL + " )XTable "
            StrSQL = StrSQL + " Group BY Transaction_Date"
            StrSQL = StrSQL + " Order BY Transaction_Date ASC"
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch + adAsyncExecute

        Do While rs.State = adStateExecuting

            If SystemOptions.BolStopUpdateTask = True Then
                rs.Cancel
                rs.Close
                SystemOptions.BolUpdateTaskInProgress = False
                Exit Sub
            End If

            DoEvents
        Loop

        '    With xChart
        '        .Gallery = Gallery_Bar
        '        .Chart3D = True
        '        .AxisX.LabelAngle = 0
        '        .WallWidth = 2
        '        .View3D = False
        '        .ShowTips = True
        '        .LegendBox = False
        '        .AllowEdit = False
        '        .MultipleColors = True
        '        .ContextMenus = False
        '        .DataSourceSettings.DataType.Item(0) = DataType_Value
        '        .DataSourceSettings.DataType.Item(1) = DataType_Label
        '        Set .DataSourceAdo = Rs
        '    End With
    End If

    '-------------------------------
    'ãÈíÚÇÊ ÇáÃÓÈæÚ ÇáÍÇáì
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "Select Count(XX) as CountRes, Sum(XX)as Res  " & " From  (Select dbo.QryOneTransactionTotal(Transaction_ID)as XX " & " From Transactions   Where ( Transactions.Transaction_Type=" & IntTransType
        StrSQL = StrSQL + " AND  (Transactions.Transaction_Date) >=" & SQLDate(GetWeekStartEND(Date, 0), True) & ""
        StrSQL = StrSQL + " AND  (Transactions.Transaction_Date) <=" & SQLDate(GetWeekStartEND(Date, 1), True) & ""
        StrSQL = StrSQL + "))XTable"
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then

        DoEvents
        StrSQL = "Select Count(XX) as CountRes, Sum(XX)as Res  " & " From  (Select QryTransactionsTotal.TotalAfterTax as XX " & " From  QryTransactionsTotal   Where ( QryTransactionsTotal.Transaction_Type=" & IntTransType
        StrSQL = StrSQL + " AND  (QryTransactionsTotal.Transaction_Date) >=" & SQLDate(GetWeekStartEND(Date, 0), True) & ""
        StrSQL = StrSQL + " AND  (QryTransactionsTotal.Transaction_Date) <=" & SQLDate(GetWeekStartEND(Date, 1), True) & ""
        StrSQL = StrSQL + "))XTable"
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch + adAsyncExecute

    Do While rs.State = adStateExecuting

        If SystemOptions.BolStopUpdateTask = True Then
            rs.Cancel
            rs.Close
            SystemOptions.BolUpdateTaskInProgress = False
            Exit Sub
        End If

        DoEvents
    Loop

    If Not (rs.BOF Or rs.EOF) Then
        DblTemp = IIf(IsNull(rs("Res").value), 0, rs("Res").value)
        DblCount = IIf(IsNull(rs("CountRes").value), 0, rs("CountRes").value)

        If SystemOptions.UserInterface = ArabicInterface Then
            StrTemp = "ÅÌãÇáì " & StrGroupTitle & " Ýì ÇáÃÓÈæÚ ÇáÍÇáì: " & DblTemp & " (" & DblCount & " ÝÇÊæÑÉ ) "
        Else
            StrTemp = "Total OF " & StrGroupTitle & " In the Current Week : " & DblTemp & " (" & DblCount & " Invoices ) "
        End If

        xGroup.Items(13).Caption = StrTemp
    End If

    '-------------------------------
    'ãÈíÚÇÊ ÇáÔåÑ ÇáÌÇÑì
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "Select Count(XX) as CountRes, Sum(XX)as Res  " & " From  (Select dbo.QryOneTransactionTotal(Transaction_ID)as XX " & " From Transactions   Where ( Transactions.Transaction_Type=" & IntTransType
        StrSQL = StrSQL + " AND  Month(Transactions.Transaction_Date)=" & Month(Date)
        StrSQL = StrSQL + " And Year(Transactions.Transaction_Date)=" & year(Date) & "))XTable"
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then

        DoEvents
        StrSQL = "Select Count(XX) as CountRes, Sum(XX)as Res  " & " From  (Select  QryTransactionsTotal.TotalAfterTax as XX " & " From  QryTransactionsTotal   Where ( QryTransactionsTotal.Transaction_Type=" & IntTransType
        StrSQL = StrSQL + " AND  Month( QryTransactionsTotal.Transaction_Date)=" & Month(Date)
        StrSQL = StrSQL + " And Year( QryTransactionsTotal.Transaction_Date)=" & year(Date) & "))XTable"
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncFetch + adAsyncExecute

    Do While rs.State = adStateExecuting

        If SystemOptions.BolStopUpdateTask = True Then
            rs.Cancel
            rs.Close
            SystemOptions.BolUpdateTaskInProgress = False
            Exit Sub
        End If

        DoEvents
    Loop

    If Not (rs.BOF Or rs.EOF) Then
        DblTemp = IIf(IsNull(rs("Res").value), 0, rs("Res").value)
        DblCount = IIf(IsNull(rs("CountRes").value), 0, rs("CountRes").value)

        If SystemOptions.UserInterface = ArabicInterface Then
            StrTemp = "ÅÌãÇáì " & StrGroupTitle & " Ýì ÇáÔåÑ ÇáÌÇÑì : " & DblTemp & " (" & DblCount & " ÝÇÊæÑÉ ) "
        Else
            StrTemp = "Total OF " & StrGroupTitle & " In the Current Month : " & DblTemp & " (" & DblCount & " Invoices ) "
        End If

        xGroup.Items(14).Caption = StrTemp
    End If

    '-------------------------------
    Exit Sub
hErr:
    Stop
    Resume
    Msg = "ÍÏË ÎØÇ "
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    Msg = Msg & Chr(13) & "ModTaskPanel:GetTransactionGroup"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

