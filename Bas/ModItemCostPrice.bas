Attribute VB_Name = "ModItemCostPrice"
Option Explicit

Private Type CostTrans
    Transactionid As Long
    costPrice As Currency
End Type
Public Function getPriceBySerial(Item_ID As Double, ItemSerial As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = " SELECT     dbo.Transaction_Details.Price"
sql = sql & " FROM         dbo.Transactions INNER JOIN"
sql = sql & "                        dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
sql = sql & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
sql = sql & "  WHERE     (dbo.Transaction_Details.ItemSerial = '" & ItemSerial & "') AND (  dbo.TransactionTypes.StockEffect = 1) AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
'Sql = Sql & "  WHERE     (dbo.Transaction_Details.ItemSerial = '" & ItemSerial & "') AND (  Transactions.Transaction_Type=22) AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
'
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        getPriceBySerial = IIf(IsNull(rs("Price").value), 0, rs("Price").value)
    Else
        getPriceBySerial = 0
    End If

    rs.Close


End Function
Public Function MinAddDate(Item_ID As Long, _
                           ToDate As Date, _
                           Optional ByRef MinDate As Date)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = "SELECT     MIN(dbo.Transactions.Transaction_Date) AS MinAddDate"
    sql = sql & " FROM         dbo.Transactions INNER JOIN"
    sql = sql & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    sql = sql & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    sql = sql & " WHERE     (dbo.TransactionTypes.StockEffect = 1) AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
    sql = sql & " AND (dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & ")"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        MinDate = IIf(IsNull(rs("MinAddDate").value), "1/1/2012", rs("MinAddDate").value)
    Else
        MinDate = "01/01/2012"
    End If

    rs.Close
End Function

Public Function GetMinDateOfQty2(LngItemID As Long, _
                                 ToDate As Date, _
                                 Optional ByRef ActQty As Double = 0, _
                                 Optional ByRef MinDate As Date, _
                                 Optional UpToTransID As Long = 0)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    '«·Õ’Ê· ⁄·Ï «Œ—  «—ÌŒ ﬂ„Ì … ’›—
 
    ' sql = " SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_Date, ISNULL(ROUND(dbo.GetItemqtytodate(dbo.Transactions.Transaction_Date, dbo.Transaction_Details.Item_ID, "
    'sql = sql & "  0), 2), 0) AS BeforeQty"
    'sql = sql & " FROM         dbo.Transactions INNER JOIN"
    'sql = sql & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    '   sql = sql & " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    'sql = sql & " WHERE     (dbo.TransactionTypes.StockEffect <> 0)"
    'sql = sql & " AND (dbo.Transactions.Transaction_Date <=" & SQLDate(todate, True) & ")"
    'sql = sql & "  and (dbo.Transaction_Details.Item_ID = " & LngItemID & ") AND (ISNULL(ROUND(dbo.GetItemqtytodate(dbo.Transactions.Transaction_Date, dbo.Transaction_Details.Item_ID, 0), 2), 0)"
    'sql = sql & " = 0) "
    ' If UpToTransID <> 0 Then
    '        sql = sql & "      AND (dbo.Transaction_Details.Transaction_ID <>    " & UpToTransID & ")"
    ' End If
 
    'sql = sql & " ORDER BY dbo.Transactions.Transaction_Date DESC"

    sql = "SELECT  MAX(Transaction_Date) Transaction_Date ,BeforeQty"
    sql = sql & "  from"
    sql = sql & " ("
    sql = sql & " SELECT     dbo.Transactions.Transaction_Date , ISNULL(ROUND(dbo.GetItemqtytodate(dbo.Transactions.Transaction_Date, dbo.Transaction_Details.Item_ID,"
    sql = sql & " 0), 2), 0) AS BeforeQty"
    sql = sql & " FROM         dbo.Transactions INNER JOIN"
    sql = sql & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    sql = sql & " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    sql = sql & " Where (dbo.TransactionTypes.StockEffect = -1) "
    sql = sql & " AND (dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & ")"

    If UpToTransID <> 0 Then
        sql = sql & "      AND (dbo.Transaction_Details.Transaction_ID <>    " & UpToTransID & ")"
    End If
 
    sql = sql & " And (dbo.Transaction_Details.Item_ID = " & LngItemID & ")"
    sql = sql & " GROUP BY dbo.Transactions.Transaction_Date, ISNULL(ROUND(dbo.GetItemqtytodate(dbo.Transactions.Transaction_Date, dbo.Transaction_Details.Item_ID, 0), 2), 0)"
    sql = sql & " )X"
    sql = sql & " Where BeforeQty = -1"
    sql = sql & " GROUP BY Transaction_Date,BeforeQty"
    sql = sql & " ORDER BY  Transaction_Date DESC"

  '  rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

  '  If rs.RecordCount > 0 Then
        '⁄‰œ ÊÃÊœ ﬂ„Ì… ’›
  '      MinDate = IIf(IsNull(rs("Transaction_Date").value), Null, rs("Transaction_Date").value)
  '      MinDate = DateAdd("D", 1, MinDate)
 
        ' ActQty = 0
  '  Else
        '«ﬁ·  «—ÌŒ «÷«›…
  '      MinAddDate LngItemID, todate, MinDate
   
  '  End If
   
  '  rs.Close
End Function

Public Function GetMinDateOfQty(LngItemID As Long, _
                                ToDate As Date, _
                                Optional ByRef ActQty As Double = 0, _
                                Optional ByRef MinDate As Date, _
                                Optional UpToTransID As Long = 0)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    sql = " SELECT     MIN(MinDate) AS MinDate, round(ActQty,3) as sumQty"
    sql = sql & " FROM         (SELECT     TOP 100 PERCENT dbo.Transaction_Details.Item_ID, SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS ActQty,"
    sql = sql & "  dbo.TblItems.ItemCode, dbo.Transactions.Transaction_Date AS MinDate"
    sql = sql & " FROM         dbo.Transactions INNER JOIN"
    sql = sql & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    sql = sql & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    sql = sql & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    sql = sql & " Where (dbo.TransactionTypes.StockEffect <> 0)  "

    If UpToTransID <> 0 Then
        sql = sql & " and  (dbo.Transaction_Details.Transaction_ID  <>" & UpToTransID & ")"
    End If

    sql = sql & " GROUP BY dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.Transactions.Transaction_Date"
    sql = sql & " HAVING      (dbo.Transaction_Details.Item_ID = " & LngItemID & ")) xx"
    sql = sql & "  GROUP BY ActQty"
    sql = sql & "  Having (round(ActQty,3) > 0)"
    sql = sql & " and  MIN(MinDate) <='" & Format((ToDate), "MM/DD/YYYY") & "'"
    sql = sql & " ORDER BY MIN(MinDate) "

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        'MinDate = Null
        ActQty = 0
    Else

        If Round(rs("sumQty").value, 3) <= 0 Then
            ActQty = 0
                       
        Else
            '
            MinDate = IIf(IsNull(rs("MinDate").value), Null, rs("MinDate").value)
            ActQty = IIf(IsNull(rs("sumQty").value), Null, rs("sumQty").value)
             
        End If
 
    End If
   
    rs.Close

End Function

Public Function GetCostItemPrice(LngItemID As Long, _
                                 Optional ByVal HaveSerial As Integer = 0, _
                                 Optional StrItemSerial As String = "", _
                                 Optional ByRef StrTransID As String = "", _
                                 Optional IntCostType As StockCostType = LastPurPriceType, _
                                 Optional DblQty As Double, _
                                 Optional FromDate As Variant = Null, _
                                 Optional ToDate As Variant = Null, _
                                 Optional UpToTransID As Long = 0, _
                                 Optional UnitID As Long = 0, Optional StoreID As Double) As Variant
    Dim ActQty As Double
    Dim MinDate As Date
        On Error Resume Next
    'LngItemID:—ﬁ„ «·’‰› «·„—«œ Õ”«»… ”⁄— «· ﬂ·›… ·Â
    'HaveSerial:«·’‰› Ì ⁄«„· »‰Ÿ«„ «·”Ì—»«·
    'StrItemSerial:—ﬁ„ «·”Ì—Ì«· ··’‰› «‰ ÊÃœ
    'On Error GoTo hErr
    Dim rs As ADODB.Recordset
    Dim rsITem As ADODB.Recordset
    Dim StrSQL As String
    Dim BolHaveSerial As Boolean
    Dim DblCostItemPrice As Double
    Dim Cmd As ADODB.Command
    Dim Par As ADODB.Parameter
    Dim DblTemp As Variant
    Dim DblTempQty As Double
    Dim DblOneUnitPrice As Currency
    Dim TempCostTrans As CostTrans
    Dim Msg As String

    
    
    If LngItemID = 89 Then
        'Stop
        LngItemID = LngItemID
    End If

    Dim RsUnitData As ADODB.Recordset
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim QtyBySmalltUnit As Double
    Dim SecOrder As Integer
    Dim UnitFactor As Double
    
    Dim SecOrderCurrent As Integer
    Dim UnitFactorCurrent As Double
 '   Dim SecOrder As Integer
    LngCurItemID = LngItemID
    LngUnitID = UnitID
    QtyBySmalltUnit = 1
    If LngUnitID = 0 Then
        StrSQL = "Select * From TblItemsUnits Where  DefaultUnit = 1 and ItemID=" & LngCurItemID
        'StrSQL = StrSQL + " AND UnitID=" & LngUnitID
        Set RsUnitData = New ADODB.Recordset
        RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
            QtyBySmalltUnit = RsUnitData("QtyBySmalltUnit").value
            LngUnitID = RsUnitData("UnitID").value
            SecOrder = val(RsUnitData("SecOrder").value & "")
            UnitFactor = val(RsUnitData("UnitFactor").value & "")
            SecOrderCurrent = val(RsUnitData("SecOrder").value & "")
            UnitFactorCurrent = val(RsUnitData("UnitFactor").value & "")
            
    
        Else
            QtyBySmalltUnit = 1
            LngUnitID = 1
            SecOrder = 1
            SecOrderCurrent = 1
            UnitFactorCurrent = 1
            
        End If
    Else
         StrSQL = "Select * From TblItemsUnits Where  UnitId = " & LngUnitID & " and ItemID=" & LngCurItemID
        'StrSQL = StrSQL + " AND UnitID=" & LngUnitID
        Set RsUnitData = New ADODB.Recordset
        RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
            'if val(RsUnitData!
            QtyBySmalltUnit = RsUnitData("UnitFactor").value
            LngUnitID = RsUnitData("UnitID").value
            UnitFactor = val(RsUnitData("UnitFactor").value & "")
            SecOrder = val(RsUnitData("SecOrder").value & "")
            SecOrderCurrent = val(RsUnitData("SecOrder").value & "")
            UnitFactorCurrent = val(RsUnitData("UnitFactor").value & "")

        Else
            QtyBySmalltUnit = 1
            LngUnitID = 1
            SecOrder = 1
            UnitFactor = 1
            SecOrderCurrent = 1
            UnitFactorCurrent = 1

        End If
    
    End If
    'HaveSerial=0 >> NOT Sent to the function
    'HaveSerial=1 >>   Sent to the function the item have serial number
    'HaveSerial=2 >> Sent to the Function the item NOT have Serial number
    If IntCostType = LastPurPriceType Then
        If HaveSerial = 0 Then
            StrSQL = "Select * From tblItems Where ItemID=" & LngItemID & ""
            Set rsITem = New ADODB.Recordset
            rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rsITem.BOF Or rsITem.EOF) Then
                BolHaveSerial = IIf(rsITem("HaveSerial").value = True, True, False)
            End If

        Else
            BolHaveSerial = IIf(HaveSerial = 1, True, False)
        End If

        If BolHaveSerial = True And StrItemSerial = "" Then
            Exit Function
        End If

        '›Ï «·ŒÿÊ… «·√Ê·Ï ‰Õ«Ê· «‰ ‰« Ï »«Œ— ”⁄— ‘—«¡
   'salim     DblTemp = GetPrice(LngItemID, 22, BolHaveSerial, StrItemSerial, StrTransID, FromDate, ToDate)

DblTemp = getcostbuylastinvoice(CLng(LngItemID), CDate(ToDate), LngUnitID)

    '    If DblTemp = 0 Then '·«ÌÊÃœ «Œ— ”⁄— ‘—«¡
    '        DblTemp = GetPrice(LngItemID, 3, BolHaveSerial, StrItemSerial, StrTransID, ToDate)  '‰Õ«Ê· «·Õ’Ê· ⁄·Ï ”⁄— «·—’Ìœ «·√›  «ÕÏ
'
'            If DblTemp = 0 Then
'                DblTemp = GetPrice(LngItemID, 9, BolHaveSerial, StrItemSerial, StrTransID, ToDate) '«Œ— ‘Ï¡ ÂÊ «·Õ’Ê· ⁄·Ï ”⁄— «Œ— „— Ã⁄ „»Ì⁄« 
'            End If
'        End If

If DblTemp = 0 Then
If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
        
            If Not IsNull(ToDate) Then
      
                GetMinDateOfQty2 LngItemID, Format(CDate(ToDate), "DD/MM/YYYY"), ActQty, MinDate, UpToTransID
                                 If SystemOptions.CostStarting = True Then
                                    Dim FirstPeriodDateInthisYear  As Date
                                    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                                
                                     MinDate = DateAdd("d", -1, FirstPeriodDateInthisYear)
                                End If
                                 If SystemOptions.AllowCostPerStore = False Then
                                           StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," & UpToTransID & ")" & "QryItemsTransactionsTotals "
                                  Else
                                          StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotalsByStores(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & StoreID & "," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotalsByStores "
                                          
                                End If
        
            Else
        
                StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(28, 3,20, '01/01/1900', ' 01/01/2079 '," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotals "
            End If
        
            StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
            StrSQL = StrSQL + " AND  TotalQty <>0"
           Set rsITem = New ADODB.Recordset
           DblTemp = 0
            rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rsITem.BOF Or rsITem.EOF) Then
                If Not IsNull(rsITem("AvCost").value) Then
                    ' DblTemp = RsItem("AvCost").value
                    DblTemp = Round(rsITem("Total").value / rsITem("TotalQty").value, 7)
                
                End If

            Else
                '«·Õ’Ê· ⁄·Ï „ Ê”ÿ «· ﬂ·›… „‰ Œ·«· „— Ã⁄ «·„»Ì⁄« 
                StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
            
                If Not IsNull(ToDate) Then
                    StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(0, 0,19, '01/01/1900', ' " & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotals "
                Else
                    StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(0, 0,19, '01/01/1900', ' 01/01/2079 '," & LngItemID & "," & UpToTransID & "  )" & "QryItemsTransactionsTotals "
                End If
            
                StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
                Set rsITem = New ADODB.Recordset
                rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rsITem.BOF Or rsITem.EOF) Then
                    If Not IsNull(rsITem("AvCost").value) Then
                        'DblTemp = RsItem("AvCost").value
                        'DblTemp = Format(RsItem("AvCost").value, SystemOptions.SysDefCurrencyForamt)
                        DblTemp = Round(rsITem("Total").value / rsITem("TotalQty").value, 4)
                        
                        
                    End If
                End If

                DblTemp = 0
            End If
        End If

End If
     ElseIf IntCostType = WeightAverage Then

        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
        
            If Not IsNull(ToDate) Then
      
                        GetMinDateOfQty2 LngItemID, Format(CDate(ToDate), "DD/MM/YYYY"), ActQty, MinDate, UpToTransID
                         If SystemOptions.CostStarting = True Then
                            'Dim FirstPeriodDateInthisYear  As Date
                            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                        
                            ' MinDate = DateAdd("d", -1, FirstPeriodDateInthisYear)
                           ' „ «·«Ìﬁ«› ›Ì 03 12 2020
                              MinDate = DateAdd("d", 0, FirstPeriodDateInthisYear)
                        End If
                                
                                
                                 If SystemOptions.AllowCostPerStore = False Then
                                           StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," & UpToTransID & ")" & "QryItemsTransactionsTotals "
                                           
                                  Else
                                          StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotalsByStores(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & StoreID & "," & LngItemID & "," & UpToTransID & "  )" & "QryItemsTransactionsTotalsByStores "
                                End If
        
            Else
        
                StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(28, 3,20, '01/01/1900', ' 01/01/2079 '," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotals "
                
            End If
        
            StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
            'StrSQL = StrSQL + " and UnitID=" & LngUnitID & ""
            StrSQL = StrSQL + " AND  TotalQty <>0"
           Set rsITem = New ADODB.Recordset
           DblTemp = 0
            rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
            If Not (rsITem.BOF Or rsITem.EOF) Then
                If Not IsNull(rsITem("AvCost").value) Then
                    ' DblTemp = RsItem("AvCost").value
                    DblTemp = Round(rsITem("Total").value / rsITem("TotalQty").value, 7)
                
                End If

            Else
            
                StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
                
                If Not IsNull(ToDate) Then
                
                'GetMinDateOfQty2 LngItemID, Format(CDate(ToDate), "DD/MM/YYYY"), ActQty, MinDate, UpToTransID
                
                MinDate = "1-1-2000"
                If SystemOptions.CostStarting = True Then
                'Dim FirstPeriodDateInthisYear  As Date
                'getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                
                ' MinDate = DateAdd("d", -1, FirstPeriodDateInthisYear)
                ' „ «·«Ìﬁ«› ›Ì 03 12 2020
                'MinDate = DateAdd("d", 0, FirstPeriodDateInthisYear)
                MinDate = "1-1-2000"
                End If
                
                
                If SystemOptions.AllowCostPerStore = False Then
                         StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," & UpToTransID & ")" & "QryItemsTransactionsTotals "
                         
                Else
                        StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotalsByStores(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & StoreID & "," & LngItemID & "," & UpToTransID & "  )" & "QryItemsTransactionsTotalsByStores "
                End If
                
                Else
                
                StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(28, 3,20, '01/01/1900', ' 01/01/2079 '," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotals "
                
                End If
                
                StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
                'StrSQL = StrSQL + " and UnitID=" & LngUnitID & ""
                StrSQL = StrSQL + " AND  TotalQty <>0"
                Set rsITem = New ADODB.Recordset
                DblTemp = 0
                rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                
                
                If Not (rsITem.BOF Or rsITem.EOF) Then
                    If Not IsNull(rsITem("AvCost").value) Then
                    ' DblTemp = RsItem("AvCost").value
                        DblTemp = Round(rsITem("Total").value / rsITem("TotalQty").value, 7)
                
                    End If
                End If
                '«·Õ’Ê· ⁄·Ï „ Ê”ÿ «· ﬂ·›… „‰ Œ·«· „— Ã⁄ «·„»Ì⁄« 
          '      StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
            
          '      If Not IsNull(ToDate) Then
          '          StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(0, 0,19, '01/01/1900', ' " & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," &  UpToTransID & "," & LngUnitID & " )" & "QryItemsTransactionsTotals "
          '      Else
          '          StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(0, 0,19, '01/01/1900', ' 01/01/2079 '," & LngItemID & "," &  UpToTransID & "," & LngUnitID & " )" & "QryItemsTransactionsTotals "
          '      End If
          '
          '      StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
              GoTo xl:
                
                
                
                
                
                
               '**********************************************************************
                          ' GetMinDateOfQty2 LngItemID, Format(CDate(ToDate), "DD/MM/YYYY"), ActQty, MinDate, UpToTransID
                          MinDate = "01/01/1900"
                          StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
                          
                                             If SystemOptions.AllowCostPerStore = False Then
                                           StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotals "
                                  Else
                                          StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotalsByStores(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & StoreID & "," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotalsByStores "
                                End If
                                
               '**************************************************************************
                
                Set rsITem = New ADODB.Recordset
                rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rsITem.BOF Or rsITem.EOF) Then
                    If Not IsNull(rsITem("AvCost").value) Then
                        'DblTemp = RsItem("AvCost").value
                        'DblTemp = Format(RsItem("AvCost").value, SystemOptions.SysDefCurrencyForamt)
                        DblTemp = Round(rsITem("Total").value / rsITem("TotalQty").value, 4)
                        
                        
                    End If
                End If

'                DblTemp = 0
            End If
        End If
xl:
    ElseIf IntCostType = FirstInFirstOut Then

        If DblQty = 0 Then
            '›Ï Õ«·… «‰ ÌﬂÊ‰ «·ﬂ„Ì… «·„ »Ìﬁ… „‰ —’Ìœ «·’‰›  ”«ÊÏ ’›—
            '›Ï Â–Â «·Õ«·…  ﬂÊ‰ ﬁÌ„… «·—’Ìœ ’›—
            DblTemp = 0
        Else
            '«·‰ﬁÿ… «·„Â„… Â‰« ÂÊ «‰‰« ‰⁄„· ﬂ„Ì… «·„Œ“Ê‰ «· Ï ”Ã·  ··’‰› ›Ï „‰
            '„‰ Œ·«· ›Ê« Ì— «·„‘—Ì«  Ê«·—’Ìœ «·√›  «ÕÏ
            'ÕÌÀ „‰Â« ‰ﬁœ— «‰ ‰Õ”» ﬁÌ„…
            Set rsITem = New ADODB.Recordset
            StrSQL = "Select * From dbo.QryItemFifoTransactions(" & LngItemID & ",'1,3,16')"
            'ÌÃ» «‰ ÌﬂÊ‰ «· — Ì»  ‰«“·Ï
            StrSQL = StrSQL + " Order By Transaction_ID DESC "
            rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (rsITem.BOF Or rsITem.EOF) Then
                DblTempQty = rsITem("TotalQty").value

                If DblTempQty >= DblQty Then
                    '«·ﬂ„Ì… «·„ÊÃÊœ… ›Ï Â–Â «·›« Ê—… «ﬂ»— „‰ «·—’Ìœ «·‰Â«∆Ï
                    '«Ê «‰ «·ﬂ„Ì… ›Ï «·›« Ê—…  ”«ÊÏ ‰›” «·ﬂ„Ì… «·„ »ﬁÌ… ﬂ—’Ìœ ‰Â«∆Ï
                    'If RsItem("Total").Value = 0 Then Stop
                    'DblOneUnitPrice = DblTempQty / RsItem("Total").Value
                    DblOneUnitPrice = rsITem("Total").value / DblTempQty
                    DblTemp = DblOneUnitPrice * DblQty
                ElseIf DblTempQty < DblQty Then
                    DblTempQty = 0

                    Do While DblTempQty < DblQty
                        DblTempQty = DblTempQty + rsITem("TotalQty").value

                        If DblTempQty <= DblQty Then
                            DblTemp = DblTemp + (rsITem("Total").value)
                        Else
                            'Stop
                            DblOneUnitPrice = rsITem("Total").value / rsITem("TotalQty").value
                            DblTemp = DblTemp + (DblOneUnitPrice * (DblQty - (DblTempQty - rsITem("TotalQty").value)))
                        End If

                        rsITem.MoveNext
                    Loop

                End If

            Else
            End If
        End If

    ElseIf IntCostType = ModernWeightAverage Then
        
        
        
                         Dim CostPriceNew As Double
                         If LngItemID = 89 Then
            LngItemID = LngItemID
        End If
                        If SystemOptions.AllowCostnNewShape = True Then
                            If IsDate(ToDate) = False Then ToDate = Date
                        
'                                getItemCostData CDate(ToDate), CDbl(LngItemID), StoreID, CDbl(UpToTransID), , , , CostPriceNew, True, LngUnitID 'CDbl(DblTemp)
'
'                                GetCostItemPrice = val(Format(CostPriceNew, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
'                                If CostPriceNew < 0 Then
'                                    CostPriceNew = CostPriceNew
'                                End If
'                                GetCostItemPrice = GetCostItemPrice * QtyBySmalltUnit
'                                Exit Function
                                
                                 getItemCostData CDate(ToDate), CDbl(LngItemID), StoreID, CDbl(UpToTransID), , , , CostPriceNew, True, , UnitFactor, SecOrder     'CDbl(DblTemp)
                                GetCostItemPrice = val(Format(CostPriceNew, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
                                If SecOrderCurrent > SecOrder Then
                                    If UnitFactorCurrent < 1 Then
                                        GetCostItemPrice = GetCostItemPrice * UnitFactorCurrent
                                    Else
                                        GetCostItemPrice = GetCostItemPrice / UnitFactorCurrent
                                    End If
                                ElseIf SecOrderCurrent < SecOrder Then
                                    GetCostItemPrice = GetCostItemPrice * UnitFactor
                                Else
                                    GetCostItemPrice = GetCostItemPrice * QtyBySmalltUnit
                                End If
                                
                                Exit Function
                      Else
                        
                        
        
                            TempCostTrans = CalModernWeightAverage(LngItemID, UpToTransID, , (ToDate), StoreID)
                     
                            GetCostItemPrice = TempCostTrans.costPrice * QtyBySmalltUnit
                        
                            StrTransID = TempCostTrans.Transactionid
                            Exit Function
    End If

    'GetCostItemPrice = DblTemp * QtyBySmalltUnit

     
                       
            
End If
endme:
    GetCostItemPrice = val(Format(DblTemp * QtyBySmalltUnit, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
    
    
    Dim mCostNew2 As Double
     If val(Format(DblTemp * QtyBySmalltUnit, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#"))) <= 0 Then
            Dim rsItemCost As ADODB.Recordset
         '   mCostNew2 = getcostbuylastinvoice(CDbl(LngItemID), CDate(ToDate), LngUnitID, UnitFactor, SecOrder)
            
            If mCostNew2 <= 0 Then
                StrSQL = " Select UnitPurPrice from TblItemsUnits where ItemID = " & val(LngItemID) & "  and UnitId = " & val(LngUnitID)
                Set rsItemCost = New ADODB.Recordset
                rsItemCost.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                If Not rsItemCost.EOF Then
                    mCostNew2 = val(rsItemCost!UnitPurPrice & "")
                End If
            End If
            GetCostItemPrice = mCostNew2 '* QtyBySmalltUnit
     End If
If SystemOptions.AllowCostBySerial = True Then
            StrSQL = "Select * From tblItems Where ItemID=" & LngItemID & ""
            Set rsITem = New ADODB.Recordset
            rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rsITem.BOF Or rsITem.EOF) Then
                BolHaveSerial = IIf(rsITem("HaveSerial").value = True, True, False)
            End If
            
If BolHaveSerial = True Then
   GetCostItemPrice = getPriceBySerial(CDbl(LngItemID), StrItemSerial)
End If

End If

    Exit Function
hErr:

    If SystemOptions.SysRegisterState = DevelopVersion Then
        Stop
        'Resume
    End If

    Msg = "ÕœÀ Œÿ«...!!!" & " GetCostItemPrice:"
    Msg = Msg & CHR(13) & "Err.Description:" & Err.Description
    Msg = Msg & CHR(13) & "Err.Number:" & Err.Number
    Msg = Msg & CHR(13) & "Err.Source:" & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function

Public Function GetCostItemPriceByGard(LngItemID As Long, _
                                 Optional ByVal HaveSerial As Integer = 0, _
                                 Optional StrItemSerial As String = "", _
                                 Optional ByRef StrTransID As String = "", _
                                 Optional IntCostType As StockCostType = LastPurPriceType, _
                                 Optional DblQty As Double, _
                                 Optional FromDate As Variant = Null, _
                                 Optional ToDate As Variant = Null, _
                                 Optional UpToTransID As Long = 0, _
                                 Optional UnitID As Long = 0, Optional StoreID As Double) As Variant
    Dim ActQty As Double
    Dim MinDate As Date
        On Error Resume Next
    'LngItemID:—ﬁ„ «·’‰› «·„—«œ Õ”«»… ”⁄— «· ﬂ·›… ·Â
    'HaveSerial:«·’‰› Ì ⁄«„· »‰Ÿ«„ «·”Ì—»«·
    'StrItemSerial:—ﬁ„ «·”Ì—Ì«· ··’‰› «‰ ÊÃœ
    'On Error GoTo hErr
    Dim rs As ADODB.Recordset
    Dim rsITem As ADODB.Recordset
    Dim StrSQL As String
    Dim BolHaveSerial As Boolean
    Dim DblCostItemPrice As Double
    Dim Cmd As ADODB.Command
    Dim Par As ADODB.Parameter
    Dim DblTemp As Variant
    Dim DblTempQty As Double
    Dim DblOneUnitPrice As Currency
    Dim TempCostTrans As CostTrans
    Dim Msg As String

    
    
    If LngItemID = 7462 Then
        'Stop
        LngItemID = LngItemID
    End If

    Dim RsUnitData As ADODB.Recordset
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim QtyBySmalltUnit As Double
    Dim SecOrder As Integer
    Dim UnitFactor As Double
    
    Dim SecOrderCurrent As Integer
    Dim UnitFactorCurrent As Double
 '   Dim SecOrder As Integer
    LngCurItemID = LngItemID
    LngUnitID = UnitID
    QtyBySmalltUnit = 1
    If LngUnitID = 0 Then
        StrSQL = "Select * From TblItemsUnits Where  DefaultUnit = 1 and ItemID=" & LngCurItemID
        'StrSQL = StrSQL + " AND UnitID=" & LngUnitID
        Set RsUnitData = New ADODB.Recordset
        RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
            QtyBySmalltUnit = RsUnitData("QtyBySmalltUnit").value
            LngUnitID = RsUnitData("UnitID").value
            SecOrder = val(RsUnitData("SecOrder").value & "")
            UnitFactor = val(RsUnitData("UnitFactor").value & "")
            SecOrderCurrent = val(RsUnitData("SecOrder").value & "")
            UnitFactorCurrent = val(RsUnitData("UnitFactor").value & "")
            
    
        Else
            QtyBySmalltUnit = 1
            LngUnitID = 1
            SecOrder = 1
            SecOrderCurrent = 1
            UnitFactorCurrent = 1
            
        End If
    Else
         StrSQL = "Select * From TblItemsUnits Where  UnitId = " & LngUnitID & " and ItemID=" & LngCurItemID
        'StrSQL = StrSQL + " AND UnitID=" & LngUnitID
        Set RsUnitData = New ADODB.Recordset
        RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
            'if val(RsUnitData!
            QtyBySmalltUnit = RsUnitData("UnitFactor").value
            LngUnitID = RsUnitData("UnitID").value
            UnitFactor = val(RsUnitData("UnitFactor").value & "")
            SecOrder = val(RsUnitData("SecOrder").value & "")
            SecOrderCurrent = val(RsUnitData("SecOrder").value & "")
            UnitFactorCurrent = val(RsUnitData("UnitFactor").value & "")

        Else
            QtyBySmalltUnit = 1
            LngUnitID = 1
            SecOrder = 1
            UnitFactor = 1
            SecOrderCurrent = 1
            UnitFactorCurrent = 1

        End If
    
    End If
    'HaveSerial=0 >> NOT Sent to the function
    'HaveSerial=1 >>   Sent to the function the item have serial number
    'HaveSerial=2 >> Sent to the Function the item NOT have Serial number
    If IntCostType = LastPurPriceType Then
        If HaveSerial = 0 Then
            StrSQL = "Select * From tblItems Where ItemID=" & LngItemID & ""
            Set rsITem = New ADODB.Recordset
            rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rsITem.BOF Or rsITem.EOF) Then
                BolHaveSerial = IIf(rsITem("HaveSerial").value = True, True, False)
            End If

        Else
            BolHaveSerial = IIf(HaveSerial = 1, True, False)
        End If

        If BolHaveSerial = True And StrItemSerial = "" Then
            Exit Function
        End If

        '›Ï «·ŒÿÊ… «·√Ê·Ï ‰Õ«Ê· «‰ ‰« Ï »«Œ— ”⁄— ‘—«¡
   'salim     DblTemp = GetPrice(LngItemID, 22, BolHaveSerial, StrItemSerial, StrTransID, FromDate, ToDate)

DblTemp = getcostbuylastinvoice(CLng(LngItemID), CDate(ToDate), LngUnitID)

    '    If DblTemp = 0 Then '·«ÌÊÃœ «Œ— ”⁄— ‘—«¡
    '        DblTemp = GetPrice(LngItemID, 3, BolHaveSerial, StrItemSerial, StrTransID, ToDate)  '‰Õ«Ê· «·Õ’Ê· ⁄·Ï ”⁄— «·—’Ìœ «·√›  «ÕÏ
'
'            If DblTemp = 0 Then
'                DblTemp = GetPrice(LngItemID, 9, BolHaveSerial, StrItemSerial, StrTransID, ToDate) '«Œ— ‘Ï¡ ÂÊ «·Õ’Ê· ⁄·Ï ”⁄— «Œ— „— Ã⁄ „»Ì⁄« 
'            End If
'        End If

If DblTemp = 0 Then
If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
        
            If Not IsNull(ToDate) Then
      
                GetMinDateOfQty2 LngItemID, Format(CDate(ToDate), "DD/MM/YYYY"), ActQty, MinDate, UpToTransID
                                 If SystemOptions.CostStarting = True Then
                                    Dim FirstPeriodDateInthisYear  As Date
                                    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                                
                                     MinDate = DateAdd("d", -1, FirstPeriodDateInthisYear)
                                End If
                                 If SystemOptions.AllowCostPerStore = False Then
                                           StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals_CostSafe(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," & UpToTransID & ")" & "QryItemsTransactionsTotals_CostSafe "
                                  Else
                                          StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotalsByStores(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & StoreID & "," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotalsByStores "
                                          
                                End If
        
            Else
        
                StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals_CostSafe(28, 3,20, '01/01/1900', ' 01/01/2079 '," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotals_CostSafe "
            End If
        
            StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
            StrSQL = StrSQL + " AND  TotalQty <>0"
           Set rsITem = New ADODB.Recordset
           DblTemp = 0
            rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rsITem.BOF Or rsITem.EOF) Then
                If Not IsNull(rsITem("AvCost").value) Then
                    ' DblTemp = RsItem("AvCost").value
                    DblTemp = Round(rsITem("Total").value / rsITem("TotalQty").value, 7)
                
                End If

            Else
                '«·Õ’Ê· ⁄·Ï „ Ê”ÿ «· ﬂ·›… „‰ Œ·«· „— Ã⁄ «·„»Ì⁄« 
                StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
            
                If Not IsNull(ToDate) Then
                    StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals_CostSafe(0, 0,19, '01/01/1900', ' " & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotals_CostSafe "
                Else
                    StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals_CostSafe(0, 0,19, '01/01/1900', ' 01/01/2079 '," & LngItemID & "," & UpToTransID & "  )" & "QryItemsTransactionsTotals_CostSafe "
                End If
            
                StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
                Set rsITem = New ADODB.Recordset
                rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rsITem.BOF Or rsITem.EOF) Then
                    If Not IsNull(rsITem("AvCost").value) Then
                        'DblTemp = RsItem("AvCost").value
                        'DblTemp = Format(RsItem("AvCost").value, SystemOptions.SysDefCurrencyForamt)
                        DblTemp = Round(rsITem("Total").value / rsITem("TotalQty").value, 4)
                        
                        
                    End If
                End If

                DblTemp = 0
            End If
        End If

End If
     ElseIf IntCostType = WeightAverage Then

        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
        
            If Not IsNull(ToDate) Then
      
                        GetMinDateOfQty2 LngItemID, Format(CDate(ToDate), "DD/MM/YYYY"), ActQty, MinDate, UpToTransID
                         If SystemOptions.CostStarting = True Then
                            'Dim FirstPeriodDateInthisYear  As Date
                            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                        
                            ' MinDate = DateAdd("d", -1, FirstPeriodDateInthisYear)
                           ' „ «·«Ìﬁ«› ›Ì 03 12 2020
                              MinDate = DateAdd("d", 0, FirstPeriodDateInthisYear)
                        End If
                                  MinDate = "01/01/1900"
                                
                                 If SystemOptions.AllowCostPerStore = False Then
                                           StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals_CostSafe(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," & UpToTransID & ")" & "QryItemsTransactionsTotals_CostSafe "
                                           
                                  Else
                                          StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotalsByStores(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & StoreID & "," & LngItemID & "," & UpToTransID & "  )" & "QryItemsTransactionsTotalsByStores "
                                End If
        
            Else
        
                StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals_CostSafe(28, 3,20, '01/01/1900', ' 01/01/2079 '," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotals_CostSafe "
                
            End If
        
            StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
            'StrSQL = StrSQL + " and UnitID=" & LngUnitID & ""
            StrSQL = StrSQL + " AND  TotalQty <>0"
           Set rsITem = New ADODB.Recordset
           DblTemp = 0
            rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
            If Not (rsITem.BOF Or rsITem.EOF) Then
                If Not IsNull(rsITem("AvCost").value) Then
                    ' DblTemp = RsItem("AvCost").value
                    DblTemp = Round(rsITem("Total").value / rsITem("TotalQty").value, 7)
                
                End If

            Else
            
                StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
                
                If Not IsNull(ToDate) Then
                
                'GetMinDateOfQty2 LngItemID, Format(CDate(ToDate), "DD/MM/YYYY"), ActQty, MinDate, UpToTransID
                
                  MinDate = "01/01/1900"
                If SystemOptions.CostStarting = True Then
                'Dim FirstPeriodDateInthisYear  As Date
                'getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                
                ' MinDate = DateAdd("d", -1, FirstPeriodDateInthisYear)
                ' „ «·«Ìﬁ«› ›Ì 03 12 2020
                'MinDate = DateAdd("d", 0, FirstPeriodDateInthisYear)
                  MinDate = "01/01/1900"
                End If
                
                
                If SystemOptions.AllowCostPerStore = False Then
                         StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals_CostSafe(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," & UpToTransID & ")" & "QryItemsTransactionsTotals_CostSafe "
                         
                Else
                        StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotalsByStores(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & StoreID & "," & LngItemID & "," & UpToTransID & "  )" & "QryItemsTransactionsTotalsByStores "
                End If
                
                Else
                
                         StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals_CostSafe(28, 3,20, '01/01/1900', ' 01/01/2079 '," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotals_CostSafe "
                
                End If
                
                StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
                'StrSQL = StrSQL + " and UnitID=" & LngUnitID & ""
                StrSQL = StrSQL + " AND  TotalQty <>0"
                Set rsITem = New ADODB.Recordset
                DblTemp = 0
                rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                
                
                If Not (rsITem.BOF Or rsITem.EOF) Then
                    If Not IsNull(rsITem("AvCost").value) Then
                    ' DblTemp = RsItem("AvCost").value
                        DblTemp = Round(rsITem("Total").value / rsITem("TotalQty").value, 7)
                
                    End If
                End If
                '«·Õ’Ê· ⁄·Ï „ Ê”ÿ «· ﬂ·›… „‰ Œ·«· „— Ã⁄ «·„»Ì⁄« 
          '      StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
            
          '      If Not IsNull(ToDate) Then
          '          StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals_CostSafe(0, 0,19, '01/01/1900', ' " & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," &  UpToTransID & "," & LngUnitID & " )" & "QryItemsTransactionsTotals_CostSafe "
          '      Else
          '          StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals_CostSafe(0, 0,19, '01/01/1900', ' 01/01/2079 '," & LngItemID & "," &  UpToTransID & "," & LngUnitID & " )" & "QryItemsTransactionsTotals_CostSafe "
          '      End If
          '
          '      StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
              GoTo xl:
                
                
                
                
                
                
               '**********************************************************************
                          ' GetMinDateOfQty2 LngItemID, Format(CDate(ToDate), "DD/MM/YYYY"), ActQty, MinDate, UpToTransID
                          MinDate = "01/01/1900"
                          StrSQL = "SELECT  ItemID, ItemCode, ItemName, Total, TotalQty, " & " CONVERT(money, Total / TotalQty, 3) AS AvCost "
                          
                                             If SystemOptions.AllowCostPerStore = False Then
                                           StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotals_CostSafe(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotals_CostSafe "
                                  Else
                                          StrSQL = StrSQL + " FROM dbo.QryItemsTransactionsTotalsByStores(28, 3,20, '" & Format(CDate(MinDate), "MM/DD/YYYY") & "', '" & Format(CDate(ToDate), "MM/DD/YYYY") & "'," & StoreID & "," & LngItemID & "," & UpToTransID & " )" & "QryItemsTransactionsTotalsByStores "
                                End If
                                
               '**************************************************************************
                
                Set rsITem = New ADODB.Recordset
                rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rsITem.BOF Or rsITem.EOF) Then
                    If Not IsNull(rsITem("AvCost").value) Then
                        'DblTemp = RsItem("AvCost").value
                        'DblTemp = Format(RsItem("AvCost").value, SystemOptions.SysDefCurrencyForamt)
                        DblTemp = Round(rsITem("Total").value / rsITem("TotalQty").value, 4)
                        
                        
                    End If
                End If

'                DblTemp = 0
            End If
        End If
xl:
    ElseIf IntCostType = FirstInFirstOut Then

        If DblQty = 0 Then
            '›Ï Õ«·… «‰ ÌﬂÊ‰ «·ﬂ„Ì… «·„ »Ìﬁ… „‰ —’Ìœ «·’‰›  ”«ÊÏ ’›—
            '›Ï Â–Â «·Õ«·…  ﬂÊ‰ ﬁÌ„… «·—’Ìœ ’›—
            DblTemp = 0
        Else
            '«·‰ﬁÿ… «·„Â„… Â‰« ÂÊ «‰‰« ‰⁄„· ﬂ„Ì… «·„Œ“Ê‰ «· Ï ”Ã·  ··’‰› ›Ï „‰
            '„‰ Œ·«· ›Ê« Ì— «·„‘—Ì«  Ê«·—’Ìœ «·√›  «ÕÏ
            'ÕÌÀ „‰Â« ‰ﬁœ— «‰ ‰Õ”» ﬁÌ„…
            Set rsITem = New ADODB.Recordset
            StrSQL = "Select * From dbo.QryItemFifoTransactions(" & LngItemID & ",'1,3,16')"
            'ÌÃ» «‰ ÌﬂÊ‰ «· — Ì»  ‰«“·Ï
            StrSQL = StrSQL + " Order By Transaction_ID DESC "
            rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (rsITem.BOF Or rsITem.EOF) Then
                DblTempQty = rsITem("TotalQty").value

                If DblTempQty >= DblQty Then
                    '«·ﬂ„Ì… «·„ÊÃÊœ… ›Ï Â–Â «·›« Ê—… «ﬂ»— „‰ «·—’Ìœ «·‰Â«∆Ï
                    '«Ê «‰ «·ﬂ„Ì… ›Ï «·›« Ê—…  ”«ÊÏ ‰›” «·ﬂ„Ì… «·„ »ﬁÌ… ﬂ—’Ìœ ‰Â«∆Ï
                    'If RsItem("Total").Value = 0 Then Stop
                    'DblOneUnitPrice = DblTempQty / RsItem("Total").Value
                    DblOneUnitPrice = rsITem("Total").value / DblTempQty
                    DblTemp = DblOneUnitPrice * DblQty
                ElseIf DblTempQty < DblQty Then
                    DblTempQty = 0

                    Do While DblTempQty < DblQty
                        DblTempQty = DblTempQty + rsITem("TotalQty").value

                        If DblTempQty <= DblQty Then
                            DblTemp = DblTemp + (rsITem("Total").value)
                        Else
                            'Stop
                            DblOneUnitPrice = rsITem("Total").value / rsITem("TotalQty").value
                            DblTemp = DblTemp + (DblOneUnitPrice * (DblQty - (DblTempQty - rsITem("TotalQty").value)))
                        End If

                        rsITem.MoveNext
                    Loop

                End If

            Else
            End If
        End If

    ElseIf IntCostType = ModernWeightAverage Then
        
        
        
                         Dim CostPriceNew As Double
                         If LngItemID = 89 Then
            LngItemID = LngItemID
        End If
                        If SystemOptions.AllowCostnNewShape = True Then
                            If IsDate(ToDate) = False Then ToDate = Date
                        
'                                getItemCostData CDate(ToDate), CDbl(LngItemID), StoreID, CDbl(UpToTransID), , , , CostPriceNew, True, LngUnitID 'CDbl(DblTemp)
'
'                                GetCostItemPriceByGard = val(Format(CostPriceNew, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
'                                If CostPriceNew < 0 Then
'                                    CostPriceNew = CostPriceNew
'                                End If
'                                GetCostItemPriceByGard = GetCostItemPrice * QtyBySmalltUnit
'                                Exit Function
                                
                                 getItemCostData CDate(ToDate), CDbl(LngItemID), StoreID, CDbl(UpToTransID), , , , CostPriceNew, True, , UnitFactor, SecOrder     'CDbl(DblTemp)
                                GetCostItemPriceByGard = val(Format(CostPriceNew, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
                                If SecOrderCurrent > SecOrder Then
                                    If UnitFactorCurrent < 1 Then
                                        GetCostItemPriceByGard = GetCostItemPriceByGard * UnitFactorCurrent
                                    Else
                                        GetCostItemPriceByGard = GetCostItemPriceByGard / UnitFactorCurrent
                                    End If
                                ElseIf SecOrderCurrent < SecOrder Then
                                    GetCostItemPriceByGard = GetCostItemPriceByGard * UnitFactor
                                Else
                                    GetCostItemPriceByGard = GetCostItemPriceByGard * QtyBySmalltUnit
                                End If
                                
                                Exit Function
                      Else
                        
                        
        
                            TempCostTrans = CalModernWeightAverage(LngItemID, UpToTransID, , (ToDate), StoreID)
                     
                            GetCostItemPriceByGard = TempCostTrans.costPrice * QtyBySmalltUnit
                        
                            StrTransID = TempCostTrans.Transactionid
                            Exit Function
    End If

    'GetCostItemPriceByGard = DblTemp * QtyBySmalltUnit

     
                       
            
End If
endme:
    GetCostItemPriceByGard = val(Format(DblTemp * QtyBySmalltUnit, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
    
    
    Dim mCostNew2 As Double
     If val(Format(DblTemp * QtyBySmalltUnit, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#"))) <= 0 Then
            Dim rsItemCost As ADODB.Recordset
         '   mCostNew2 = getcostbuylastinvoice(CDbl(LngItemID), CDate(ToDate), LngUnitID, UnitFactor, SecOrder)
            
            If mCostNew2 < 0 Then
                StrSQL = " Select UnitPurPrice from TblItemsUnits where ItemID = " & val(LngItemID) & "  and UnitId = " & val(LngUnitID)
                Set rsItemCost = New ADODB.Recordset
                rsItemCost.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                If Not rsItemCost.EOF Then
                    mCostNew2 = val(rsItemCost!UnitPurPrice & "")
                End If
            End If
            GetCostItemPriceByGard = mCostNew2 '* QtyBySmalltUnit
     End If
If SystemOptions.AllowCostBySerial = True Then
            StrSQL = "Select * From tblItems Where ItemID=" & LngItemID & ""
            Set rsITem = New ADODB.Recordset
            rsITem.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rsITem.BOF Or rsITem.EOF) Then
                BolHaveSerial = IIf(rsITem("HaveSerial").value = True, True, False)
            End If
            
If BolHaveSerial = True Then
   GetCostItemPriceByGard = getPriceBySerial(CDbl(LngItemID), StrItemSerial)
End If

End If

    Exit Function
hErr:

    If SystemOptions.SysRegisterState = DevelopVersion Then
        Stop
        'Resume
    End If

    Msg = "ÕœÀ Œÿ«...!!!" & " GetCostItemPriceByGard:"
    Msg = Msg & CHR(13) & "Err.Description:" & Err.Description
    Msg = Msg & CHR(13) & "Err.Number:" & Err.Number
    Msg = Msg & CHR(13) & "Err.Source:" & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Function



Public Function GetPrice(LngItemID As Long, _
                         IntTransType As Integer, _
                         Optional BolHaveSerial As Boolean = False, _
                         Optional StrItemSerial As String = "", _
                         Optional ByRef StrTransID As String, _
                         Optional FromDate As Variant = Null, _
                         Optional ToDate As Variant = Null) As Double
    Dim Cmd As ADODB.Command
    Dim Par As ADODB.Parameter
    Dim ParItem_ID As ADODB.Parameter
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        Set Cmd = New ADODB.Command
        Set Par = New ADODB.Parameter
        Set ParItem_ID = New ADODB.Parameter

        If BolHaveSerial = False Then
            StrSQL = "Select * From QryLastItemsPurPrice Where Item_ID=" & LngItemID & ""
            StrSQL = StrSQL + " Order By Transaction_ID ASC"
        Else
            StrSQL = "Select * From QryLastItemsPurPrice Where Item_ID=" & LngItemID & ""

            If StrItemSerial <> "" Then
                StrSQL = StrSQL + " AND ItemSerial='" & StrItemSerial & "'"
            End If
        End If

        Cmd.CommandText = StrSQL
        Cmd.CommandType = adCmdUnknown
        Set Cmd.ActiveConnection = Cn
        Par.Direction = adParamInput
        Par.Name = "X"
        Par.value = IntTransType
        Par.type = adInteger
    
        ParItem_ID.Direction = adParamInput
        ParItem_ID.Name = "Y"
        ParItem_ID.value = LngItemID
        ParItem_ID.type = adInteger
    
        Cmd.Parameters.Append Par
        Cmd.Parameters.Append ParItem_ID
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic
        rs.LockType = adLockReadOnly
        'Set Rs = Cmd.Execute
        rs.Open Cmd
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then

        If BolHaveSerial = False Then
            If Not (IsNull(ToDate) Or IsEmpty(ToDate)) Then
                StrSQL = "Select * From dbo.QryLastItemsPurPrice (" & IntTransType & "," & LngItemID & "," & SQLDate(CDate(ToDate), True) & ") QryLastItemsPurPrice "
            Else
                ' DateTime
                StrSQL = "Select * From dbo.QryLastItemsPurPrice (" & IntTransType & "," & LngItemID & ",DEFAULT) QryLastItemsPurPrice "
            End If

            StrSQL = StrSQL + " Order By Transaction_ID ASC"
        Else
            'StrSQL = "Select * From dbo.QryLastItemsPurPrice (" & IntTransType & "," & LngItemID & ") QryLastItemsPurPrice "
            'If StrItemSerial <> "" Then
            '    StrSQL = StrSQL + " Where ItemSerial='" & StrItemSerial & "'"
            'End If
            'StrSQL = StrSQL + " Order By Transaction_ID ASC"
            StrSQL = "SELECT dbo.Transactions.Transaction_ID, dbo.Transaction_Details.Price, dbo.Transactions.Transaction_Type,"
            StrSQL = StrSQL + " dbo.Transaction_Details.ItemSerial "
            StrSQL = StrSQL + " FROM dbo.Transactions INNER JOIN"
            StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID "
            StrSQL = StrSQL + " = dbo.Transaction_Details.Transaction_ID"
            StrSQL = StrSQL + " Where dbo.Transactions.Transaction_Type=" & IntTransType & ""
            StrSQL = StrSQL + " AND dbo.Transaction_Details.Item_ID=" & LngItemID & ""
            StrSQL = StrSQL + " AND dbo.Transaction_Details.ItemSerial='" & StrItemSerial & "'"
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    End If

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
        GetPrice = IIf(IsNull(rs("Price").value), 0, rs("Price").value)
        StrTransID = IIf(IsNull(rs("Transaction_ID").value), "N/A", rs("Transaction_ID").value)
    Else
        GetPrice = 0
        StrTransID = ""
    End If

    rs.Close
    Set rs = Nothing
    Set Cmd = Nothing

End Function

Public Function GetItemStockToDate(LngItemID As Long, _
                                   LngTransID As Long, _
                                   LngStoreID As Long, _
                                   ToDate As Date) As Single
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim SngTemp As Single
    'ﬁ„  »⁄„· Â–Â «·œ«·…
    '··√” ⁄·«„ ⁄‰ —’Ìœ ’‰› „⁄Ì‰
    'ﬁ»·  «—ÌŒ   Ê—ﬁ„ Õ—ﬂ… „⁄Ì‰…  ›Ï «·»—‰«„Ã
    StrSQL = "SELECT     SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) as SumQty "
    StrSQL = StrSQL & " FROM         dbo.Transaction_Details INNER JOIN"
    StrSQL = StrSQL & " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    StrSQL = StrSQL & "  WHERE     (dbo.TransactionTypes.StockEffect <> 0) AND (dbo.Transaction_Details.Item_ID = " & LngItemID & " )  "
    StrSQL = StrSQL & " AND (dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & ")"
    StrSQL = StrSQL & "  AND (dbo.Transaction_Details.Transaction_ID <> " & LngTransID & ")"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        SngTemp = IIf(IsNull(rs("SumQty").value), 0, rs("SumQty").value)
    Else
        SngTemp = 0
    End If

    rs.Close
    Set rs = Nothing

    If SngTemp < 0 Then
        SngTemp = 0
    End If

    GetItemStockToDate = SngTemp
End Function

Public Function GetItemStockToTrans(LngItemID As Long, _
                                    LngTransID As Long) As Single
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim SngTemp As Single
    'ﬁ„  »⁄„· Â–Â «·œ«·…
    '··√” ⁄·«„ ⁄‰ —’Ìœ ’‰› „⁄Ì‰
    'ﬁ»· œŒÊ· Õ—ﬂ… „⁄Ì‰… ›Ï «·»—‰«„Ã
    StrSQL = "SELECT TOP 100 PERCENT ItemID, ItemCode, ItemName, GroupID, SUM(QTY) AS SumQty"
    StrSQL = StrSQL + " FROM dbo.QryGardToTrans(" & LngTransID & ")QryGardToTrans "
    StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
    StrSQL = StrSQL + " GROUP BY ItemID, ItemCode, ItemName, GroupID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        SngTemp = IIf(IsNull(rs("SumQty").value), 0, rs("SumQty").value)
    Else
        SngTemp = 0
    End If

    rs.Close
    Set rs = Nothing
    GetItemStockToTrans = SngTemp
End Function

Private Function CalModernWeightAverage(LngItemID As Long, _
                                        Optional LngToInvTransID As Long = 0, _
                                        Optional StoreID As Long, _
                                        Optional ToDate As Date, Optional StoreId1 As Double) As CostTrans 'xx

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim Msg As String
    Dim SngQty As Double
    Dim SngTemp1 As Double, SngTemp2 As Double, SngTemp3 As Double, SngTemp4 As Double
    Dim NetCost As CostTrans

    Dim LngTransID As Long
    Dim SngTransQty As Double
    Dim SngXPrice As Double
    Dim SngBeforeQty As Double
    Dim SngBeforeCostPrice As Double
    Dim SngNewCostPrice As Double

    If LngItemID = 0 Then
        Exit Function
    End If
  If LngToInvTransID <> 0 Then
     StrSQL = "SELECT       isnull(round (dbo.GetItemqtytodate( Transactions.Transaction_Date ,Item_ID ," & LngToInvTransID & "),2),0) as BeforeQty ,  dbo.Transactions.Transaction_Date,   Transactions.Transaction_ID, dbo.Transaction_Details.Quantity AS xqty, dbo.Transaction_Details.Price AS xprice"
Else
  StrSQL = "SELECT       isnull(round (dbo.GetItemqtytodate( Transactions.Transaction_Date ,Item_ID ,0),2),0) as BeforeQty ,  dbo.Transactions.Transaction_Date,   Transactions.Transaction_ID, dbo.Transaction_Details.Quantity AS xqty, dbo.Transaction_Details.Price AS xprice"
  End If
  'GetItemqtytodate2015
   
   
   StrSQL = "SELECT       isnull(round (dbo.GetItemqtytodate2015( Transactions.Transaction_Date ,Item_ID ,dbo.Transactions.Transaction_ID," & LngToInvTransID & "),2), 0) as BeforeQty ,  dbo.Transactions.Transaction_Date,   Transactions.Transaction_ID, dbo.Transaction_Details.Quantity AS xqty, dbo.Transaction_Details.Price AS xprice"
   
    StrSQL = StrSQL + " FROM         dbo.Transaction_Details INNER JOIN"
    StrSQL = StrSQL + " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
    StrSQL = StrSQL + " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    StrSQL = StrSQL + " WHERE     (dbo.TransactionTypes.StockEffect = 1) AND (dbo.Transaction_Details.Item_ID = " & LngItemID & ") AND (dbo.Transactions.Transaction_Date <= " & SQLDate(ToDate, True) & " ) "
 
    StrSQL = StrSQL + " AND (dbo.Transaction_Details.Transaction_ID <> " & LngToInvTransID & ")"
    If SystemOptions.AllowCostPerStore Then
        StrSQL = StrSQL + " AND (dbo.Transactions.StoreId = " & StoreID & ")"
        StrSQL = StrSQL + " ORDER BY dbo.Transactions.StoreID, dbo.Transaction_Details.Transaction_ID"
    Else
        StrSQL = StrSQL + " ORDER BY dbo.Transaction_Details.Transaction_ID"
    End If
    
 'StoreId
    

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    'BeforeQty
Dim mTotalCost As Double
    If Not (rs.BOF Or rs.EOF) Then

        For i = 1 To rs.RecordCount
            LngTransID = IIf(IsNull(rs("Transaction_ID").value), 0, rs("Transaction_ID").value)
            SngTransQty = IIf(IsNull(rs("XQty").value), 0, rs("XQty").value)
            SngXPrice = IIf(IsNull(rs("XPrice").value), 0, rs("XPrice").value)
            SngBeforeQty = IIf(IsNull(rs("BeforeQty").value), 0, rs("BeforeQty").value)
            
           
            If SngBeforeQty < 0 Then
                'SngBeforeQty = 0
            ' SngBeforeQty = GetItemStockToDate(LngItemID, val(rs("Transaction_ID").value), StoreId, rs("Transaction_Date").value)
             End If
             
       
            '«·„⁄«œ·…
            '(«·ﬂ„Ì… «·»«ﬁÌ… »”⁄— «· ﬂ·›… «·√ŒÌ—  „÷—Ê»« ›Ï  ”⁄— «· ﬂ·›…«·√ŒÌ—)
            '+
            '(«·ﬂ„Ì… «·Ê«—œ… »”⁄— «· ﬂ·›… «·ÃœÌœ „÷—Ê»« ›Ï ”⁄— «· ﬂ·›… «·ÃœÌœ)
            '„ﬁ”„« ⁄·Ï
            '≈Ã„«·Ï «·ﬂ„Ì… «·ÃœÌœ… Ê«·ﬁœÌ„…
            If rs!Transaction_Date = "01/04/2020" Then
                SngBeforeQty = 0
                SngTemp4 = SngXPrice
                GoTo NextRow
            End If
                     If LngItemID = 89 Then
                 LngItemID = LngItemID
                End If
            If rs("XQty").value > 0.001 Then
            If SngNewCostPrice * SngBeforeQty > 0 Then
                SngTemp1 = SngNewCostPrice * SngBeforeQty
            Else
                SngTemp1 = 0
            End If
            If SngTemp1 < 0 Then
                SngTemp1 = SngTemp1
            End If
            
            '(SngBeforeQty + SngTransQty) * SngBeforeCostPrice
            SngTemp2 = SngTransQty * SngXPrice
            
            If SngBeforeQty < 0 Then
                SngBeforeQty = SngBeforeQty
            End If
            If SngBeforeQty = 0 Then
                SngTemp3 = 0
            Else
                SngTemp3 = SngBeforeQty + SngTransQty
            End If
            If mTotalCost < 0 Then
                mTotalCost = mTotalCost
            End If
            mTotalCost = SngTemp2 + SngTemp1
            If (SngTemp3) <> 0 Then
                SngTemp4 = (SngTemp1 + SngTemp2) / (SngTemp3)
            Else
                SngTemp4 = 0
            End If
            If SngTemp4 <= 0 Then
                SngTemp4 = getcostbuylastinvoice(CDbl(LngItemID), CDate(rs!Transaction_Date & ""))
            End If
            If SngTemp4 > 200 Then
                SngTemp4 = SngTemp4
            End If
NextRow:
            SngNewCostPrice = Round(SngTemp4, SystemOptions.SysDefCurrencyForamt)
            End If
            
            rs.MoveNext
            
            SngBeforeCostPrice = SngNewCostPrice
        Next i

        NetCost.Transactionid = LngTransID
        NetCost.costPrice = SngNewCostPrice
    Else
    End If
    
        
    
    CalModernWeightAverage = NetCost
End Function

Private Function CalModernWeightAveragexxx(LngItemID As Long, _
                                           Optional LngToInvTransID As Long = 0) As CostTrans

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim Msg As String
    Dim SngQty As Single
    Dim SngTemp1 As Single, SngTemp2 As Single, SngTemp3 As Single, SngTemp4 As Single
    Dim NetCost As CostTrans

    Dim LngTransID As Long
    Dim SngTransQty As Single
    Dim SngXPrice As Single
    Dim SngBeforeQty As Single
    Dim SngBeforeCostPrice As Single
    Dim SngNewCostPrice As Single

    If LngItemID = 0 Then
        Exit Function
    End If

    StrSQL = "Select * From RptItemTransCus"
    StrSQL = StrSQL + " Where Item_ID=" & LngItemID
    StrSQL = StrSQL + " AND (Transaction_Type=1 OR  Transaction_Type=3)"

    If LngToInvTransID <> 0 Then
        StrSQL = StrSQL + " AND Transaction_ID <" & LngToInvTransID
    End If

    StrSQL = StrSQL + " Order BY Transaction_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then

        For i = 1 To rs.RecordCount
            LngTransID = IIf(IsNull(rs("Transaction_ID").value), 0, rs("Transaction_ID").value)
            SngTransQty = IIf(IsNull(rs("XQty").value), 0, rs("XQty").value)
            SngXPrice = IIf(IsNull(rs("XPrice").value), 0, rs("XPrice").value)
            SngBeforeQty = GetItemStockToTrans(LngItemID, val(rs("Transaction_ID").value))
            '«·„⁄«œ·…
            '(«·ﬂ„Ì… «·»«ﬁÌ… »”⁄— «· ﬂ·›… «·√ŒÌ—  „÷—Ê»« ›Ï  ”⁄— «· ﬂ·›…«·√ŒÌ—)
            '+
            '(«·ﬂ„Ì… «·Ê«—œ… »”⁄— «· ﬂ·›… «·ÃœÌœ „÷—Ê»« ›Ï ”⁄— «· ﬂ·›… «·ÃœÌœ)
            '„ﬁ”„« ⁄·Ï
            '≈Ã„«·Ï «·ﬂ„Ì… «·ÃœÌœ… Ê«·ﬁœÌ„…

            SngTemp1 = SngBeforeQty * SngBeforeCostPrice
            SngTemp2 = SngTransQty * SngXPrice
            SngTemp3 = SngBeforeQty + SngTransQty

            If SngTemp3 <> 0 Then
                SngTemp4 = (SngTemp1 + SngTemp2) / SngTemp3
            Else
                SngTemp4 = 0
            End If

            SngNewCostPrice = Format(SngTemp4, SystemOptions.SysDefCurrencyForamt)
            rs.MoveNext
            SngBeforeCostPrice = SngNewCostPrice
        Next i

        NetCost.Transactionid = LngTransID
        NetCost.costPrice = SngNewCostPrice
    Else
    End If

    CalModernWeightAveragexxx = NetCost
End Function

Public Sub UpdateTransCost(LngTransID As Long)

    Dim rs              As ADODB.Recordset
    Dim StrSQL          As String
    Dim i               As Long, j As Long
    Dim Msg             As String
    Dim RsTrans         As ADODB.Recordset
    Dim DblItemValue    As Double
    Dim DblItemCost     As Double

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT TOP 100 PERCENT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial," & "dbo.Transactions.Transaction_Date,dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode," & "dbo.TblItems.ItemName, dbo.Transaction_Details.ItemSerial,dbo.Transaction_Details.Quantity," & "dbo.Transaction_Details.Price, dbo.Transaction_Details.CostPrice, dbo.Transaction_Details.CostTransID," & "dbo.Transaction_Details.ItemDiscountType,dbo.Transaction_Details.ItemDiscount," & "dbo.Transaction_Details.ItemProfit , dbo.Transaction_Details.Id, dbo.TblCustemers.CusName "
        StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN "
        StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = " & "dbo.Transaction_Details.Transaction_ID INNER JOIN dbo.TblItems ON dbo.Transaction_Details.Item_ID =" & "dbo.TblItems.ItemID INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID" & " Where(dbo.Transactions.Transaction_Type = 2) "
        StrSQL = StrSQL + " AND dbo.Transaction_Details.CostTransID=" & LngTransID & ""
        StrSQL = StrSQL + " ORDER BY dbo.Transactions.Transaction_ID, dbo.Transaction_Details.ID "
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT TOP 100 PERCENT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Date,Transaction_Details.Item_ID, TblItems.ItemCode,TblItems.ItemName," & "Transaction_Details.ItemSerial,Transaction_Details.Quantity,Transaction_Details.Price," & "Transaction_Details.CostPrice, Transaction_Details.CostTransID,Transaction_Details.ItemDiscountType," & "Transaction_Details.ItemDiscount,Transaction_Details.ItemProfit,Transaction_Details.Id , TblCustemers.CusName"
        StrSQL = StrSQL + " FROM TblCustemers INNER JOIN (TblItems INNER JOIN (Transactions INNER JOIN " & "Transaction_Details ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID) ON TblItems.ItemID =" & "Transaction_Details.Item_ID) ON TblCustemers.CusID = Transactions.CusID "
        StrSQL = StrSQL + " Where(Transactions.Transaction_Type = 2)  AND Transaction_Details.CostTransID=" & LngTransID & ""
        StrSQL = StrSQL + " ORDER BY Transactions.Transaction_ID,Transaction_Details.ID"
       
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "Â–Â «·Õ—ﬂ… ·Ì”  ·Â« ›Ê« Ì— »Ì⁄  ﬁ«∆„…  ﬂ·› Â« ⁄·ÌÂ«"
        '  MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    Load FrmItemsCostUpdate

    With FrmItemsCostUpdate.FG
        .rows = .FixedRows
        .AutoSize 0, .Cols - 1, False
    
        .rows = .FixedRows + rs.RecordCount

        For i = 1 To rs.RecordCount
            .TextMatrix(i, .ColIndex("Serial")) = i
            .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
            .TextMatrix(i, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
            .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(rs("Transaction_Date").value), "", rs("Transaction_Date").value)
            .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
            .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(rs("Item_ID").value), "", rs("Item_ID").value)
            .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
            .TextMatrix(i, .ColIndex("ItemSerial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
            .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(rs("Quantity").value), "", rs("Quantity").value)
            .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(rs("Price").value), "", rs("Price").value)
            '
            .TextMatrix(i, .ColIndex("ItemDiscountType")) = IIf(IsNull(rs("ItemDiscountType").value), 0, rs("ItemDiscountType").value)
            .TextMatrix(i, .ColIndex("ItemDiscount")) = IIf(IsNull(rs("ItemDiscount").value), 0, rs("ItemDiscount").value)
            '---------------------------------------------------------------------
            'Õ”«» ﬁÌ„… «·Œ’„ ⁄·Ï ﬂ· ’‰›
            '⁄‰ ÿ—Ìﬁ ÷—» «·ﬂ„Ì… ›Ï «·”⁄—
            DblItemValue = val(.TextMatrix(i, .ColIndex("Quantity"))) * val(.TextMatrix(i, .ColIndex("Price")))
        
            If val(.TextMatrix(i, .ColIndex("ItemDiscountType"))) = 0 Or val(.TextMatrix(i, .ColIndex("ItemDiscountType"))) = 1 Then
                '·«ÌÊÃœ Œ’„
                .TextMatrix(i, .ColIndex("DiscountValue")) = 0
            ElseIf val(.TextMatrix(i, .ColIndex("ItemDiscountType"))) = 2 Then
                'Œ’„ ﬁÌ„…
                .TextMatrix(i, .ColIndex("DiscountValue")) = DblItemValue - val(.TextMatrix(i, .ColIndex("ItemDiscount")))
            ElseIf val(.TextMatrix(i, .ColIndex("ItemDiscountType"))) = 3 Then
                'Œ’„ ‰”»…
                .TextMatrix(i, .ColIndex("DiscountValue")) = DblItemValue * (1 - (val(.TextMatrix(i, .ColIndex("ItemDiscount"))) / 100))
            ElseIf val(.TextMatrix(i, .ColIndex("ItemDiscountType"))) = 4 Then
                'Œ’„ ﬂ«„·(„Ã«‰Ï)·
                .TextMatrix(i, .ColIndex("DiscountValue")) = DblItemValue
            End If
        
            '--------------------------------------------------------------------
            .TextMatrix(i, .ColIndex("CostPrice")) = IIf(IsNull(rs("CostPrice").value), "", rs("CostPrice").value)
            .TextMatrix(i, .ColIndex("CostTransID")) = IIf(IsNull(rs("CostTransID").value), "", rs("CostTransID").value)
            .TextMatrix(i, .ColIndex("ItemProfit")) = IIf(IsNull(rs("ItemProfit").value), "", rs("ItemProfit").value)
            .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
            rs.MoveNext
        Next i

        .AutoSize 0, .Cols - 1, False
        '--------------------------------------------------------------------------
        Set RsTrans = New ADODB.Recordset
        StrSQL = "Select * From Transaction_Details Where Transaction_ID=" & LngTransID
        StrSQL = StrSQL + " Order By Item_ID,Transaction_Details.ID"
        RsTrans.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        RsTrans.MoveFirst

        For i = 0 To RsTrans.RecordCount - 1

            If IsNull(RsTrans("ItemSerial").value) Then

                'NO Serial
                For j = .FixedRows To .rows - 1

                    If .TextMatrix(j, .ColIndex("Item_ID")) = RsTrans("Item_ID").value Then
                        .TextMatrix(j, .ColIndex("NewCostPrice")) = RsTrans("Price").value
                    
                    End If

                Next j

            Else

                'Have Serial
                For j = .FixedRows To .rows - 1

                    If .TextMatrix(j, .ColIndex("Item_ID")) = RsTrans("Item_ID").value Then
                        If .TextMatrix(j, .ColIndex("ItemSerial")) = RsTrans("ItemSerial").value Then
                            .TextMatrix(j, .ColIndex("NewCostPrice")) = RsTrans("Price").value
                        End If
                    End If

                Next j

            End If

            RsTrans.MoveNext
        Next i

        '--------------------------------------------------------------------------
        For i = .FixedRows To .rows - 1
            '----------------------------------------------------------------------
            'Õ”«» ﬁÌ„… «·—»Õ «·ÃœÌœ… »⁄œ «· ⁄œÌ·
            'Ì”«ÊÏ ≈Ã„«·Ï ﬁÌ„… «·’‰› „ÿ—ÊÕ „‰Â«
            '{ﬁÌ„… «·Œ’„ + ﬁÌ„… «· ﬂ·›…}
        
            'ﬁÌ„… «·’‰›
            DblItemValue = val(.TextMatrix(i, .ColIndex("Quantity"))) * val(.TextMatrix(i, .ColIndex("Price")))
        
            'ﬁÌ„… «· ﬂ·›…
            DblItemCost = val(.TextMatrix(i, .ColIndex("Quantity"))) * val(.TextMatrix(i, .ColIndex("NewCostPrice")))
                
            .TextMatrix(i, .ColIndex("NewItemProfit")) = DblItemValue - (DblItemCost + .TextMatrix(i, .ColIndex("DiscountValue")))
        Next i

        '--------------------------------------------------------------------------
        For i = .FixedRows To .rows - 1
            StrSQL = "Update Transaction_Details"
            StrSQL = StrSQL + " Set Transaction_Details.CostPrice=" & val(.TextMatrix(i, .ColIndex("NewCostPrice")))
            StrSQL = StrSQL + ",ItemProfit=" & val(.TextMatrix(i, .ColIndex("NewItemProfit")))
            StrSQL = StrSQL + " Where Transaction_Details.ID=" & val(.TextMatrix(i, .ColIndex("ID")))
            Cn.Execute StrSQL, , adExecuteNoRecords
        Next i

        For i = .FixedRows To .rows - 1

            If val(.TextMatrix(i, .ColIndex("NewCostPrice"))) <> val(.TextMatrix(i, .ColIndex("CostPrice"))) Then
                .cell(flexcpBackColor, i, 1, i, .Cols - 1) = &H8080FF
            Else
                .cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbGreen
            End If

        Next i

    End With

    Screen.MousePointer = vbDefault
    FrmItemsCostUpdate.show vbModal
End Sub

