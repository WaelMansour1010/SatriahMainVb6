VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsItemsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowZeroItemsStock(StrSQL As String, _
                              m_PrintTarget As PrintTarget, _
                              IntReportStyle As Integer)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim RsData As ADODB.Recordset
    Dim cReportViewer As ClsReportViewer
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim StrReportTitle As String

    Dim StrFileName As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        If IntReportStyle = 0 Then
            StrFileName = App.path & "\Reports\ItemsZeroStockStyle2.rpt"
        ElseIf IntReportStyle = 1 Then
            StrFileName = App.path & "\Reports\ItemsZeroStock.rpt"
        End If

    Else

        If IntReportStyle = 0 Then
            StrFileName = App.path & "\Reports\ItemsZeroStockStyle2_Eng.rpt"
        ElseIf IntReportStyle = 1 Then
            StrFileName = App.path & "\Reports\ItemsZeroStock_Eng.rpt"
        End If
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    If SystemOptions.UserInterface = ArabicInterface Then
        StrReportTitle = " ﬁ—Ì— »«·√’‰«› –«  «·√—’œ… «·’›—Ì…"
    Else
        StrReportTitle = " ﬁ—Ì— »«·√’‰«› –«  «·√—’œ… «·’›—Ì…"
    End If

    xReport.Database.SetDataSource RsData
    xReport.reporttitle = StrReportTitle
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_PrintTarget, xReport.reporttitle, , , , StrFileName

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
    'Resume
End Sub

Public Function TypicalProduction(ID As Integer, _
                                  PeriodName As String, _
                                  branchname As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim CViewer As ClsReportViewer
    On Error GoTo ErrTrap

    If ID = 0 Then
        Exit Function
    End If

    StrSQL = "SELECT    Allocations1, Allocations, RecordDate, intervalid, branch_no, Remarks, Expenses, SalaryVouchersTotals, MaterialIssueVoucherTotals, AccDep, total, UnitValue, SaleValue, TotalProductionQty, "
    StrSQL = StrSQL & "  TotalsalesQty , NoteID, NoteSerial, NoteSerial1, ID, SaLePayValue, profit,SalesValue1,SalesValue2,ProfitService"
    StrSQL = StrSQL & " from dbo.TblTypicalProduction"
    StrSQL = StrSQL & "  WHERE     id =" & ID
    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If Dir(App.path & "\Reports\Inventory\TypicalProductions.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\ItemsReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\Inventory\TypicalProductions.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(5).AddCurrentValue PeriodName
        xReport.ParameterFields(6).AddCurrentValue branchname
        
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , App.path & "\Reports\Inventory\TypicalProductions.rpt"
    Screen.MousePointer = vbDefault
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Function ItemData(ItemID As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim CViewer As ClsReportViewer
    On Error GoTo ErrTrap

    If ItemID = 0 Then
        Exit Function
    End If

    StrSQL = "select * From QryItemsReport where ItemID=" & ItemID
    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If Dir(App.path & "\Reports\ItemsReport.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\ItemsReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ItemsReport.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , App.path & "\Reports\ItemsReport.rpt"
    Screen.MousePointer = vbDefault
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Function GetUnitFactorAndName(ItemID As Long, SecOrder As Integer, Optional ByRef UnitFactor As Double, Optional ByRef UnitName As String)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = "SELECT      dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItemsUnits.DefaultUnit, dbo.TblItemsUnits.ItemID, dbo.TblItemsUnits.SecOrder"
    sql = sql & "  ,    dbo.TblItemsUnits.UnitFactor FROM         dbo.TblItemsUnits LEFT OUTER JOIN"
    sql = sql & "   dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
    sql = sql & " WHERE     (dbo.TblItemsUnits.ItemID = " & ItemID & ") AND (dbo.TblItemsUnits.SecOrder = " & SecOrder & ")"
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        UnitFactor = IIf(IsNull(rs("UnitFactor").value), 0, rs("UnitFactor").value)
        If SystemOptions.UserInterface = ArabicInterface Then
        UnitName = IIf(IsNull(rs("unitname").value), "", rs("unitname").value)
       Else
       UnitName = IIf(IsNull(rs("unitnamee").value), "", rs("unitnamee").value)
       End If
    Else
        UnitFactor = 0
        UnitName = ""
    End If
   
    rs.Close
End Function

Public Function GetDefaultItemUnit(ItemID As Long, _
                                   Optional ByRef UnitID As Integer, _
                                   Optional ByRef UnitName As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = " SELECT     dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItemsUnits.DefaultUnit, dbo.TblItemsUnits.ItemID"
    sql = sql & " FROM         dbo.TblItemsUnits LEFT OUTER JOIN"
    sql = sql & " dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
    sql = sql & "  WHERE     (dbo.TblItemsUnits.DefaultUnit = 1) AND (dbo.TblItemsUnits.ItemID = " & ItemID & ")"
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        UnitID = IIf(IsNull(rs("UnitID").value), 0, rs("UnitID").value)
If SystemOptions.UserInterface = ArabicInterface Then
        UnitName = IIf(IsNull(rs("UnitName").value), 0, rs("UnitName").value)
Else
UnitName = IIf(IsNull(rs("UnitNamee").value), 0, rs("UnitNamee").value)
End If

    Else
        UnitID = 0
        UnitName = ""
    End If
   
    rs.Close
End Function
Public Function ItemCart(LngItemID As Long, _
                         LngStoreID As Long, _
                         FromDate As Variant, _
                         ToDate As Variant, _
                         Optional BolShowTable As Boolean = False, _
                         Optional m_PrintTarget As PrintTarget = WindowTarget, _
                         Optional BolShowunit As Boolean, _
                         Optional SerialNo As String = "", _
                         Optional lotNo As String = "", _
                         Optional ColorID As Integer = 0, _
                         Optional sizeid As Integer = 0, _
                         Optional ClassId As Integer = 0, _
                         Optional colorname As String, _
                         Optional sizename As String, _
                         Optional classname As String, _
                         Optional StoreName As String, _
                         Optional BolShowOrder As Boolean, _
                         Optional order_no As String = "", Optional ByVal InOut As Integer = 0) As Boolean
    
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsBegine As ADODB.Recordset
    Dim RsEndQty As ADODB.Recordset
    Dim DblBegineQty As Double
    Dim DblBegineQtybyunit  As Double

    Dim DblEndQty As Double
    Dim RsData As ADODB.Recordset
    Dim FrmReport As FrmReportViewer
    Dim cReportViewer As ClsReportViewer
    Dim StrRS As ADODB.Recordset
    Dim i As Integer
    Dim Msg As String
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim cProgress As ClsProgress

    'On Error GoTo ErrTrap
    If BolShowTable = False Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrFileName = App.path & "\Reports\ItemCart.rpt"
                    Else
                     StrFileName = App.path & "\Reports\ItemCartE.rpt"
                    End If
    
    Else
                        If SystemOptions.UserInterface = ArabicInterface Then

        StrFileName = App.path & "\Reports\ItemCartTable.rpt"
        Else
        StrFileName = App.path & "\Reports\ItemCartTablee.rpt"
        End If
    End If

    If BolShowunit = True Then
     If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ItemCart1.rpt"
      Else
      StrFileName = App.path & "\Reports\ItemCart1E.rpt"
      End If
    End If

    If BolShowTable = True And BolShowunit = True Then
         If SystemOptions.UserInterface = ArabicInterface Then

        StrFileName = App.path & "\Reports\ItemCartTableBigUnit.rpt"
        Else
         StrFileName = App.path & "\Reports\ItemCartTableBigUnitE.rpt"
        End If
        
    End If

    If SystemOptions.UserInterface = EnglishInterface Then
   '     StrFileName = App.path & "\Reports\ItemCartE.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    If Not IsNull(FromDate) Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT Sum(IIf(Transactions.Transaction_Type=1 Or Transactions.Transaction_Type=9" & " Or Transactions.Transaction_Type=11 Or Transactions.Transaction_Type=12 Or " & "Transactions.Transaction_Type=3,Transaction_Details.Quantity*1,IIf(Transactions.Transaction_Type=2 " & "Or Transactions.Transaction_Type=8 Or Transactions.Transaction_Type=10 Or " & "Transactions.Transaction_Type=13 Or Transactions.Transaction_Type=5," & "Transaction_Details.Quantity*-1,0))) AS QTY, TblItems.ItemID, TblItems.ItemCode," & "TblItems.ItemName, Transaction_Details.ItemSerial, TblItems.HaveSerial," & "Transactions.StoreID, TblStore.StoreName "
            StrSQL = StrSQL + " FROM (TblStore INNER JOIN Transactions ON TblStore.StoreID = " & "Transactions.StoreID) INNER JOIN (TblItems INNER JOIN Transaction_Details ON " & "TblItems.ItemID = Transaction_Details.Item_ID) ON Transactions.Transaction_ID " & "= Transaction_Details.Transaction_ID "
            StrSQL = StrSQL + " Where (TblItems.ItemID=" & LngItemID & ") AND (Transactions.StoreID=" & LngStoreID & ")"
            StrSQL = StrSQL + " AND Transactions.Transaction_Date <" & SQLDate(CDate(FromDate), True) & ""
            StrSQL = StrSQL + " GROUP BY TblItems.ItemID, TblItems.ItemCode, TblItems.ItemName," & "Transaction_Details.ItemSerial, TblItems.HaveSerial, Transactions.StoreID, TblStore.StoreName"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        
            StrSQL = "SELECT     TOP 100 PERCENT QryQuantity.* "
            StrSQL = StrSQL + " FROM dbo.QryQuantity(" & SQLDate(CDate(FromDate) - 1, True) & ", 0) QryQuantity "
            '      StrSQL = StrSQL + " Where (ItemID=" & LngItemID & ") AND " & _
                   "(StoreID=" & LngStoreID & ")"
    
            StrSQL = StrSQL + " Where (ItemID=" & LngItemID & ")  "

            If LngStoreID <> 0 Then
                StrSQL = StrSQL & " AND (StoreID=" & LngStoreID & ")"
            End If
        
        End If

        '   Set RsBegine = New ADODB.Recordset
        '   RsBegine.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        '   If Not (RsBegine.BOF Or RsBegine.EOF) Then
        '       DblBegineQty = IIf(IsNull(RsBegine("Qty").value), "", RsBegine("Qty").value)
        '   End If
    Else
        DblBegineQty = 0
    End If

    DblBegineQty = ItemOpeningBalances(LngItemID, LngStoreID, ColorID, sizeid, ClassId, FromDate, 0, order_no)

    If BolShowunit = True Then
        Dim bigunitFactor As Double
        bigunitFactor = ItemBigUnit(LngItemID)

        If bigunitFactor <> 0 Then
        
            DblBegineQtybyunit = DblBegineQty / bigunitFactor
        End If
    End If

    'Set RsEndQty = New ADODB.Recordset
    'If SystemOptions.SysDataBaseType = AccessDataBase Then
    '    StrSQL = "Select * From QryGardComplete Where ItemID=" & LngItemID & " AND StoreID=" & LngStoreID & ""
    'ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
    '    StrSQL = "Select * From dbo.QryGardComplete(0)QryGardComplete Where ItemID=" & LngItemID & " AND StoreID=" & LngStoreID & ""
    'End If
    'RsEndQty.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If Not (RsEndQty.BOF Or RsEndQty.EOF) Then
    '    DblEndQty = 0
    '    RsEndQty.MoveFirst
    '   For I = 0 To RsEndQty.RecordCount - 1
    '       DblEndQty = DblEndQty + IIf(IsNull(RsEndQty("Qty").value), 0, RsEndQty("Qty").value)
    '       RsEndQty.MoveNext
    '   Next I
    'End If

    'StrSQL = "Select * From QryItemCart"
    'StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
    'StrSQL = StrSQL + " And StoreID=" & LngStoreID & ""

    StrSQL = "SELECT  dbo.Transactions.TransactionComment,  dbo.TblItems.barcodeno,   dbo.Transaction_Details.lotNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.TransactionTypes.TransactionTypeName, "
    StrSQL = StrSQL + "  dbo.Transaction_Details.ItemSerial ,dbo.TransactionTypes.TransactionEnglishName, dbo.TransactionTypes.StockEffect, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.Quantity AS transqty,"
    StrSQL = StrSQL + "  dbo.Transactions.NoteSerial1, dbo.Transactions.CusID, dbo.Transactions.StoreID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    StrSQL = StrSQL + "  dbo.TblStore.StoreName, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId, dbo.TblItemsColors.ColorName,"
    StrSQL = StrSQL + "  dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName, isnull(dbo.TransactionTypes.StockEffect * dbo.Transaction_Details.Quantity,0) AS actualqty,"
    StrSQL = StrSQL + "  dbo.Transactions.Transaction_Serial, dbo.Transactions.Nots2, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.Groups.GroupName,"
    StrSQL = StrSQL + "  dbo.Transactions.Transaction_ID , dbo.Transaction_Details.Price AS unitprice, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice,       dbo.TblItemsclasses.SizeName AS ClassName,"
   StrSQL = StrSQL + "   dbo.Transactions.CashCustomerName, dbo.Transaction_Details.FromStoreAr, dbo.Transaction_Details.FromstoreEn , dbo.Transaction_Details.TostoreAr, dbo.Transaction_Details.TostoreEn "
  StrSQL = StrSQL + "   ,     dbo.TblItems.ItemNamee, dbo.Groups.GroupNamee, dbo.TblStore.StoreNamee, dbo.TblUnites.UnitNamee,  dbo.Groups.Fullcode  FROM         dbo.Transactions INNER JOIN"
  
    StrSQL = StrSQL + "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL + "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL + "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL + "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL + "  dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL + "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL + "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL + "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL + "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL + "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL + "  WHERE   (dbo.TransactionTypes.StockEffect <> 0)   "
    If InOut = 2 Then
        StrSQL = StrSQL + "  and (Transactions.Transaction_Type=19 or Transactions.Transaction_Type=2 or Transactions.Transaction_Type=5 or Transactions.Transaction_Type=8 or Transactions.Transaction_Type=10 or Transactions.Transaction_Type=13 OR  Transactions.Transaction_Type=16  OR  Transactions.Transaction_Type=27  )"
    ElseIf InOut = 1 Then
        StrSQL = StrSQL + "  and (Transactions.Transaction_Type=20 or  Transactions.Transaction_Type=1 or Transactions.Transaction_Type=3 or Transactions.Transaction_Type=9 or Transactions.Transaction_Type=11 or Transactions.Transaction_Type=12 OR  Transactions.Transaction_Type=15  OR  Transactions.Transaction_Type=28 )"
    End If
    'StrSQL = StrSQL + "  (dbo.Transaction_Details.Item_ID = " & LngItemID & ") AND (dbo.Transactions.StoreID = " & LngStoreID & ") "
    If LngItemID <> 0 Then
        StrSQL = StrSQL + " and  (dbo.Transaction_Details.Item_ID = " & LngItemID & ")"
    End If

    If ColorID <> 0 Then
        StrSQL = StrSQL & " AND (dbo.Transaction_Details.ColorID=" & ColorID & ")"
    End If
    
    If sizeid <> 0 Then
        StrSQL = StrSQL & " AND (dbo.Transaction_Details.ItemSize=" & sizeid & ")"
    End If
    
    If ClassId <> 0 Then
        StrSQL = StrSQL & " AND (dbo.Transaction_Details.ClassId=" & ClassId & ")"
    End If
    
    If lotNo <> "" Then
        StrSQL = StrSQL + " and  (dbo.Transaction_Details.lotNo ='" & lotNo & "')"
    End If
           
    If order_no <> "" Then
        StrSQL = StrSQL + " and  (dbo.Transaction_Details.order_no ='" & order_no & "')"
    End If
            
    If LngStoreID <> 0 Then
        StrSQL = StrSQL & " AND (dbo.Transactions.StoreID=" & LngStoreID & ")"
    End If

    If SerialNo <> "" Then
        StrSQL = StrSQL & "AND (dbo.Transaction_Details.ItemSerial = N'" & SerialNo & "') "
    End If
          
    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
    End If

    If Not IsNull(ToDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
    End If

    StrSQL = StrSQL + " Order By dbo.Transactions.Transaction_Date,dbo.Transactions.Transaction_ID"

'    Set cProgress = New ClsProgress
'    cProgress.ProgressType = Waiting
'    cProgress.StartProgress
    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute

    Do While RsData.State = adStateExecuting
        DoEvents
    Loop

    If RsData.BOF Or RsData.EOF Then
'        cProgress.FinishProgress
'        cProgress.StopProgess
'        Set cProgress = Nothing
        '    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        '    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '    Screen.MousePointer = vbDefault
        '    ItemCart = False
        '    Exit Function

        If BolShowTable = False Then
            StrFileName = App.path & "\Reports\ItemCartnoitem.rpt"

        Else
            StrFileName = App.path & "\Reports\ItemCartTablenoitem.rpt"
        End If

    Else
        ItemCart = True
    End If

If SystemOptions.UserInterface = ArabicInterface Then
    StrReportTitle = "   ﬁ—Ì— ﬂ«—  «·’‰›  "

    If StoreName <> "" Then
        StrReportTitle = StrReportTitle & CHR(13) & " «·„Œ“‰:  " & StoreName
    End If

    If colorname <> "" Then
        StrReportTitle = StrReportTitle & CHR(13) & " «··Ê‰:  " & colorname
    End If

    If sizename <> "" Then
        StrReportTitle = StrReportTitle & CHR(13) & " «·„ﬁ«”:  " & sizename
    End If

    If classname <> "" Then
        StrReportTitle = StrReportTitle & CHR(13) & " «·›—“:  " & classname
    End If

    If lotNo <> "" Then
        StrReportTitle = StrReportTitle & CHR(13) & " ···Êÿ —ﬁ„ " & lotNo
    End If
 
    If order_no <> "" Then

        StrReportTitle = StrReportTitle & CHR(13) & " ··‘Õ‰… —ﬁ„ " & order_no
    End If
 
 Else
 'english
 StrReportTitle = " Items Card  "

    If StoreName <> "" Then
        StrReportTitle = StrReportTitle & CHR(13) & " Store:  " & StoreName
    End If

    If colorname <> "" Then
        StrReportTitle = StrReportTitle & CHR(13) & " Color:  " & colorname
    End If

    If sizename <> "" Then
        StrReportTitle = StrReportTitle & CHR(13) & " Size:  " & sizename
    End If

    If classname <> "" Then
        StrReportTitle = StrReportTitle & CHR(13) & " Class:  " & classname
    End If

    If lotNo <> "" Then
        StrReportTitle = StrReportTitle & CHR(13) & " Lot No " & lotNo
    End If
 
    If order_no <> "" Then

        StrReportTitle = StrReportTitle & CHR(13) & " Shipment No" & order_no
    End If
 
 
 End If
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    xReport.reporttitle = StrReportTitle
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name

    xReport.ParameterFields(4).AddCurrentValue DblBegineQty

    xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), CDate(FromDate))
    xReport.ParameterFields(7).AddCurrentValue DblEndQty

    'xReport.ParameterFields(6).AddCurrentValue IIf(IsNull(todate), CDate("01/01/1000"), CDate(todate))
    xReport.ParameterFields(6).AddCurrentValue IIf(IsNull(ToDate), CDate("01/01/1000"), (ToDate))

    If BolShowunit = True Then
        If BolShowTable = True Then
            If RsData.BOF Or RsData.EOF Then
            Else
                xReport.ParameterFields(10).AddCurrentValue val(DblBegineQtybyunit)
            End If

        Else
        
            If RsData.BOF Or RsData.EOF Then
                                                            
            Else
                xReport.ParameterFields(17).AddCurrentValue val(DblBegineQtybyunit)
            End If
        End If

    End If

    If RsData.BOF Or RsData.EOF Then
        Dim itemcode As String
        Dim ItemName As String
 
        GetItemData CInt(LngItemID), itemcode, ItemName
        xReport.ParameterFields(17).AddCurrentValue itemcode
        xReport.ParameterFields(18).AddCurrentValue ItemName
    End If

    Dim UnitID As Integer
    Dim UnitName As String
    Dim unitfator As Double
    GetDefaultItemUnit LngItemID, UnitID, UnitName

    xReport.ParameterFields(8).AddCurrentValue UnitName

    'If BolShowunit = True Or BolShowTable = True Then GoTo xx
    If BolShowTable = True Then GoTo xx
    'xReport.ParameterFields(8).AddCurrentValue UnitName

    GetUnitFactorAndName LngItemID, 2, unitfator, UnitName

    If unitfator <> 0 Then
        xReport.ParameterFields(9).AddCurrentValue CStr(unitfator)
        xReport.ParameterFields(10).AddCurrentValue UnitName
    Else
        xReport.ParameterFields(9).AddCurrentValue "0"
        xReport.ParameterFields(10).AddCurrentValue "0"
    End If

    GetUnitFactorAndName LngItemID, 3, unitfator, UnitName

    If unitfator <> 0 Then
        xReport.ParameterFields(11).AddCurrentValue CStr(unitfator)
        xReport.ParameterFields(12).AddCurrentValue UnitName
    Else
        xReport.ParameterFields(11).AddCurrentValue "0"
        xReport.ParameterFields(12).AddCurrentValue "0"
    End If

    GetUnitFactorAndName LngItemID, 4, unitfator, UnitName

    If unitfator <> 0 Then
        xReport.ParameterFields(13).AddCurrentValue CStr(unitfator)
        xReport.ParameterFields(14).AddCurrentValue UnitName
    Else
        xReport.ParameterFields(13).AddCurrentValue "0"
        xReport.ParameterFields(14).AddCurrentValue "0"
    End If

    GetUnitFactorAndName LngItemID, 5, unitfator, UnitName

    If unitfator <> 0 Then
        xReport.ParameterFields(15).AddCurrentValue CStr(unitfator)
        xReport.ParameterFields(16).AddCurrentValue UnitName
    Else
        xReport.ParameterFields(15).AddCurrentValue "0"
        xReport.ParameterFields(16).AddCurrentValue "0"
    End If

xx:

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_PrintTarget, xReport.reporttitle, , , , StrFileName, , StrSQL

    If Not cProgress Is Nothing Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If

    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:

    If Not cProgress Is Nothing Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If

    Msg = "·«Ì„ﬂ‰ ⁄—÷ «· ﬁ—Ì— ...!!!"
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Source
    Screen.MousePointer = vbDefault
End Function
 

Public Sub ItemsOutInTransactions(IntStoreID As Integer, _
                                  Optional M_Date As Date, _
                                  Optional m_PrintTarget As PrintTarget = WindowTarget)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim FrmReport As FrmReportViewer
    Dim Msg As String
    Dim StrFileName As String
    Dim cReportViewer As ClsReportViewer

    On Error GoTo ErrTrap

    StrFileName = App.path & "\Reports\ItemsOutInTransactions.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ItemsOutInTransactions.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    Set rs = GetItemsRsTransactions(IntStoreID, M_Date)

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
    
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue M_Date
        xReport.reporttitle = " ﬁ—Ì— »Õ—ﬂ… «·„Œ“Ê‰ ›Ï ÌÊ„ " & CStr(Format(M_Date, "yyyy/M/d"))
        Screen.MousePointer = vbDefault
    End If

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_PrintTarget, "", , , , StrFileName
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Sub

Public Function ShowGuarantee(IntType As Integer, _
                              IntInterval As Integer, _
                              IntReportDisplayTyppe As ReportDisplayDataType, _
                              m_PrintTarget As PrintTarget)
    'IntType=0
    'IntType=1
    'IntType=2
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim Msg As String
    Dim StrRS As New ADODB.Recordset
    Dim StrFileName As String
    Dim cReportViewer As ClsReportViewer

    On Error GoTo ErrTrap

    StrFileName = App.path & "\Reports\GuaranteeAlram.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    StrSQL = "Select * From QryGuaranteeAlram Where"

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        If IntType = 0 Then
            StrSQL = StrSQL + " ExpireDate <=#" & SQLDate(Date) & "#"
        ElseIf IntType = 1 Then
            StrSQL = StrSQL + " ExpireDate <=#" & SQLDate(DateAdd("d", IntInterval, Date)) & "# And ExpireDate >=#" & SQLDate(Date) & "#"
            'Me.Lbl(4).Caption = " «—ÌŒ ‰Â«Ì… «·› —…:" & Format(DateAdd("d", IntInterval, Date), "yyyy/M/d")
        ElseIf IntType = 2 Then
            StrSQL = StrSQL + " ExpireDate <=#" & SQLDate(DateAdd("d", IntInterval, Date)) & "#"
            'Me.Lbl(4).Caption = " «—ÌŒ ‰Â«Ì… «·› —…:" & Format(DateAdd("d", IntInterval, Date), "yyyy/M/d")
        End If

    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then

        If IntType = 0 Then
            StrSQL = StrSQL + " ExpireDate <='" & SQLDate(Date) & "'"
        ElseIf IntType = 1 Then
            StrSQL = StrSQL + " ExpireDate <='" & SQLDate(DateAdd("d", IntInterval, Date)) & "' And ExpireDate >='" & SQLDate(Date) & "'"
            'Me.Lbl(4).Caption = " «—ÌŒ ‰Â«Ì… «·› —…:" & Format(DateAdd("d", IntInterval, Date), "yyyy/M/d")
        ElseIf IntType = 2 Then
            StrSQL = StrSQL + " ExpireDate <='" & SQLDate(DateAdd("d", IntInterval, Date)) & "'"
            'Me.Lbl(4).Caption = " «—ÌŒ ‰Â«Ì… «·› —…:" & Format(DateAdd("d", IntInterval, Date), "yyyy/M/d")
        End If
    End If

    Set rs = New ADODB.Recordset
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
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If IntType = 0 Then
            xReport.reporttitle = " ﬁ—Ì— »«·√’‰«› «· Ï ≈‰ Â  „œ… ÷„«‰Â«"
        ElseIf IntType = 1 Then
            xReport.reporttitle = " ﬁ—Ì— »«·√’‰«› «· Ï ”Ê›  ‰ ÂÏ „œ… ÷„«‰Â« ›Ï Œ·«· " & IntInterval & " ÌÊ„ "
        ElseIf IntType = 2 Then
            xReport.reporttitle = " ﬁ—Ì— »«·√’‰«› «· Ï ≈‰ Â  «Ê ”Ê›  ‰ ÂÏ „œ… ÷„«‰Â« ›Ï Œ·«· " & IntInterval & " ÌÊ„ "
        End If

        Screen.MousePointer = vbDefault
    End If

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_PrintTarget, "", , , , StrFileName
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Function ShowItemStocks(LngGroupID As Integer, _
                               FromDate As Variant, _
                               ToDate As Variant, _
                               Optional m_Traget As PrintTarget = WindowTarget)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim cReportViewer As ClsReportViewer
    Dim StrFileName As String
 
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\NewItemsStock.rpt"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrFileName = App.path & "\Reports\NewItemsStock_Eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    StrSQL = "Select * From NewTempTable"

    If LngGroupID <> 0 Then
        StrSQL = StrSQL + " WHERE GroupId=" & LngGroupID
    End If

    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.reporttitle = " ﬁ—Ì— »Ã—œ «·„Œ“Ê‰"
    xReport.Database.SetDataSource RsData
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), FromDate)
    xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(ToDate), CDate("01/01/1000"), ToDate)
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_Traget, xReport.reporttitle, , , , StrFileName
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Screen.MousePointer = vbDefault
End Function

 

Public Sub ShowItemsProfits(FromDate As Variant, _
                            ToDate As Variant, _
                            Optional m_Traget As PrintTarget = WindowTarget, Optional BranchID As Integer = 0, _
                                    Optional branchname As String, Optional BrnchIDes As String, Optional BranshesActiv As String)
    
 

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim cReportViewer As ClsReportViewer
    Dim StrFileName As String
  Dim branchtext As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ItemsProfits.rpt"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsE.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
   ' StrSQL = "SELECT    dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, SUM(dbo.Transaction_Details.Quantity) AS SumQty, "
   ' StrSQL = StrSQL + " SUM(dbo.Transaction_Details.Quantity * dbo.Transaction_Details.Price*ISNULL (Currency_rate,1)) AS Total, SUM(dbo.Transaction_Details.ItemProfit) AS SumProfit"
   ' StrSQL = StrSQL + " FROM dbo.Transaction_Details INNER JOIN"
   ' StrSQL = StrSQL + "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
   ' StrSQL = StrSQL + " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
   ' StrSQL = StrSQL + " Where (dbo.Transactions.Transaction_Type = 9  OR  dbo.Transactions.Transaction_Type = 21)"




  StrSQL = " SELECT     dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, SUM(dbo.Transaction_Details.Quantity) AS SumQty,"
  StrSQL = StrSQL + "  SUM(dbo.Transaction_Details.showQty * dbo.Transaction_Details.showPrice * ISNULL(dbo.Transactions.Currency_rate, 1)) AS Total,"
  StrSQL = StrSQL + "  ISNULL("
  StrSQL = StrSQL + " SUM(dbo.Transaction_Details.ItemProfit),"
  StrSQL = StrSQL + " (   SUM(dbo.Transaction_Details.showQty * dbo.Transaction_Details.showprice * ISNULL(dbo.Transactions.Currency_rate, 1)) - sum(dbo.Transaction_Details.showqty *dbo.Transaction_Details.costprice"
  StrSQL = StrSQL + " )"
  StrSQL = StrSQL + " )*-1"
  StrSQL = StrSQL + " ) AS SumProfit"
  StrSQL = StrSQL + " ,  dbo.TblItems.ItemNamee  FROM         dbo.Transaction_Details INNER JOIN"
  StrSQL = StrSQL + " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
  StrSQL = StrSQL + " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
  StrSQL = StrSQL + " Where (dbo.Transactions.Transaction_Type = 9  OR  dbo.Transactions.Transaction_Type = 21)"
 If BrnchIDes <> "-1" And BrnchIDes <> "" Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId in (" & BrnchIDes & ")"
  End If
  If BranshesActiv <> "-1" And BranshesActiv <> "" Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId in (" & BranshesActiv & ")"
  End If
  

  If BranchID <> 0 Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId= " & BranchID
  If SystemOptions.UserInterface = ArabicInterface Then
  branchtext = "«·›—⁄  : " & branchname
  Else
   branchtext = "Branch " & branchname
  End If
  End If
 
 
 
    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
    End If

    If Not IsNull(ToDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
    End If

    StrSQL = StrSQL + " GROUP BY dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,  dbo.TblItems.ItemNamee"

    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
If SystemOptions.UserInterface = ArabicInterface Then

    xReport.reporttitle = " ﬁ—Ì— »√—»«Õ «·√’‰«›" & CHR(13) & branchtext
Else
    xReport.reporttitle = "Total Items Profits" & CHR(13) & branchtext

End If

    xReport.Database.SetDataSource RsData
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    'xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(3).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), CStr(FromDate))
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(ToDate), CDate("01/01/1000"), CStr(ToDate))
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_Traget, xReport.reporttitle, , , , StrFileName
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub



  Public Sub ShowItemsProfitsByinvoice(FromDate As Variant, _
                                     ToDate As Variant, _
                                     Optional m_Traget As PrintTarget = WindowTarget, Optional BranchID As Integer = 0, _
                                    Optional branchname As String, Optional BrnchIDes As String, Optional BranshesActiv As String)
 
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim cReportViewer As ClsReportViewer
    Dim StrFileName As String
  Dim branchtext As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsInvoice.rpt"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsInvoiceE.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    StrSQL = " SELECT     dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_ID, SUM(dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice) AS totalsale, "
    StrSQL = StrSQL & " SUM(dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.CostPrice) AS totalcost, dbo.Transactions.Transaction_Date"
    StrSQL = StrSQL & "  ,  dbo.Transactions.Transaction_Type   FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_Type = 21 or dbo.Transactions.Transaction_Type = 9)"
 
  If BrnchIDes <> "" And BrnchIDes <> "-1" Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId in(" & BrnchIDes & ")"
  End If
 If BranshesActiv <> "" And BranshesActiv <> "-1" Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId in(" & BranshesActiv & ")"
  End If
   If BranchID <> 0 Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId= " & BranchID
  If SystemOptions.UserInterface = ArabicInterface Then
  branchtext = "«·›—⁄  : " & branchname
  Else
   branchtext = "Branch  : " & branchname
  End If
  End If


    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
    End If

    If Not IsNull(ToDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
    End If
 
    StrSQL = StrSQL & "  GROUP BY dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type"

    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
  If SystemOptions.UserInterface = ArabicInterface Then
    xReport.reporttitle = " ﬁ—Ì— »√—»«Õ «·√’‰«› ÿ»ﬁ« ··›Ê« Ì—" & CHR(13) & branchtext
    Else
    xReport.reporttitle = "Bill Profits" & CHR(13) & branchtext
    End If
    xReport.Database.SetDataSource RsData
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    'xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(3).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), CStr(FromDate))
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(ToDate), CDate("01/01/1000"), CStr(ToDate))
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_Traget, xReport.reporttitle, , , , StrFileName
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub




 
 
  
   Public Sub ShowItemsProfitsByitems(FromDate As Variant, _
                                   ToDate As Variant, _
                                   Optional m_Traget As PrintTarget = WindowTarget, Optional BranchID As Integer = 0, _
                                    Optional branchname As String, Optional sql As String, Optional reportid As Integer = 0, Optional BrnchIDes As String, Optional BranshesActiv As String, Optional CurrenrEmployeeNames As String, Optional Item_ID As Double, Optional GroupID As Double, Optional GroupName As String, Optional StoreName As String)
 '*********************************************************************************
 '*********************************************************************************
 '*********************************************************************************'*********************************************************************************'*********************************************************************************
 '*********************************************************************************
 '*********************************************************************************
 'Œœ »«·ﬂ „‘ Â‰«
 'Œœ »«·ﬂ „‘ Â‰«
 'Œœ »«·ﬂ „‘ Â‰«
 'Œœ »«·ﬂ „‘ Â‰«
  '*********************************************************************************
   '*********************************************************************************
    '*********************************************************************************
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim cReportViewer As ClsReportViewer
    Dim StrFileName As String
 On Error Resume Next
    On Error GoTo ErrTrap
If reportid = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsbyitems.rpt"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsbyitems-E.rpt"
    End If
ElseIf reportid = 1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsbyitems1.rpt"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsbyitems1-E.rpt"
    End If


ElseIf reportid = 61 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsbyitems1Customers.rpt"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsbyitems1Customers-E.rpt"
    End If
    
    
ElseIf reportid = 2 Then

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsbyitems2.rpt"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsbyitems2-E.rpt"
    End If
ElseIf reportid = 3 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsbyitems3.rpt"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        StrFileName = App.path & "\Reports\ItemsProfitsbyitems3-E.rpt"
    End If

End If


    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
Dim branchtext As String
    Set RsData = New ADODB.Recordset
    'StrSQL = " SELECT  dbo.Transaction_Details.TotalDiscountPerLine,dbo.Transaction_Details.ItemDiscountValue ,    dbo.Transactions.Transaction_Type,  dbo.Transactions.NoteSerial1, dbo.Transaction_Details.CostPrice, dbo.Transaction_Details.showPrice, dbo.Transactions.Transaction_ID, "
    'StrSQL = StrSQL & "  dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice AS totalsale,"
    'StrSQL = StrSQL & " dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.CostPrice AS totalcost, dbo.Transactions.Transaction_Date"
    'StrSQL = StrSQL & "  , dbo.TblItems.ItemCode,  dbo.TblItems.ItemName, dbo.TblItems.ItemNamee FROM         dbo.Transactions INNER JOIN"
    'StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    'StrSQL = StrSQL & "   INNER JOIN dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    
    '************************************12022019*********************
    StrSQL = "SELECT     tblitems.DefaultSupplier, dbo.Transaction_Details.CostPrice,TblItems.ItemWithOutVAT, dbo.TblCustemers.fullcode, tblemployee.emp_name, tblemployee.emp_nameE, CONVERT(varchar, "
 StrSQL = StrSQL & "                        month(dbo.Transactions.Transaction_Date)) + '-' + CONVERT(varchar, year(dbo.Transactions.Transaction_Date)) AS Monthdate1, dbo.Transaction_Details.SHOWQTY,"
 StrSQL = StrSQL & "                        dbo.Transaction_Details.SHOWpRICE, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date,"
StrSQL = StrSQL & "                         dbo.TransactionTypes.TransactionTypeName, 1 AS UnitFactor, dbo.TransactionTypes.TransactionEnglishName, dbo.TransactionTypes.StockEffect,"
 StrSQL = StrSQL & "                        dbo.Transactions.Transaction_Type, dbo.Transactions.PaymentType, dbo.Transactions.CusID, dbo.Transactions.StoreID, dbo.Transactions.UserID,"
  StrSQL = StrSQL & "                       dbo.Transactions.Emp_ID, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transaction_Details.UnitId,"
 StrSQL = StrSQL & "                        dbo.Transaction_Details.OpeningBurcahseValue, dbo.Transaction_Details.OpeningBurcahseQty, dbo.Transaction_Details.OpeningSalesQty,"
StrSQL = StrSQL & "                         dbo.Transaction_Details.OpeningSalesValue, dbo.TblUnites.UnitName, dbo.TblStore.StoreName, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "                         dbo.Transactions.PayedValue, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.Tel,"
StrSQL = StrSQL & "                         dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee,"
StrSQL = StrSQL & "                         dbo.Transaction_Details.OpeningReSalesQty, dbo.Transaction_Details.OpeningReSalesValue, (ISNULL(dbo.Transaction_Details.TotalDiscountPerLine, 0)"
StrSQL = StrSQL & "                         * ISNULL(dbo.Transactions.Currency_rate, 1)) AS TotalDiscountPerLine, ISNULL(dbo.Transactions.Currency_rate, 1) AS Currency_rate,"
 StrSQL = StrSQL & "                        'ItemDiscountValue' = CASE WHEN ItemDiscountType = 1 THEN 0 WHEN ItemDiscountType = 2 THEN ((ItemDiscount * ISNULL(dbo.Transactions.Currency_rate, 1)))"
StrSQL = StrSQL & "                         WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100))"
StrSQL = StrSQL & "                         * ISNULL(dbo.Transactions.Currency_rate, 1)) WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice)"
StrSQL = StrSQL & "                         * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END, (showPrice * ISNULL(dbo.Transactions.Currency_rate, 1)) AS localprice,"
StrSQL = StrSQL & "                         (showPrice * ISNULL(dbo.Transactions.Currency_rate, 1) * dbo.Transaction_Details.showqty) AS LineNet, NOTESERIAL1, dbo.TblItems.ItemNameE,"
StrSQL = StrSQL & "                         dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress,"
StrSQL = StrSQL & "                         dbo.Transactions.CashCustomerComment, dbo.Transactions.VAT, dbo.Transactions.VATNO, dbo.Transactions.Transaction_NetValue,"
StrSQL = StrSQL & "                         dbo.Transaction_Details.Vat AS VatDet, dbo.Transaction_Details.Vatyo, dbo.Transactions.BasedOn, 1 AS MiniSellingPrice,"
StrSQL = StrSQL & "                         TblCustemers_1.CusName AS defualtCusName, TblCustemers_1.CusNamee AS defualtCusNamee"
StrSQL = StrSQL & "    FROM         dbo.Transactions INNER JOIN"
 StrSQL = StrSQL & "                        dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
StrSQL = StrSQL & "                         dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
 StrSQL = StrSQL & "                        dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
  StrSQL = StrSQL & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
  StrSQL = StrSQL & "                       dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
  StrSQL = StrSQL & "                       dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID INNER JOIN"
 StrSQL = StrSQL & "                        dbo.TblCustemers TblCustemers_1 ON dbo.TblItems.DefaultSupplier = TblCustemers_1.CusID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                        dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
   StrSQL = StrSQL & "                     dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
   StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID"
                      
    
    
    
    
    StrSQL = StrSQL & "  WHERE     (dbo.Transactions.Transaction_Type = 21 OR   dbo.Transactions.Transaction_Type = 9)"
    
        If GroupID <> 0 Then
        sql = sql & " AND (dbo.TblItems.GroupID = " & GroupID & ")"
    End If

 If BrnchIDes <> "-1" And BrnchIDes <> "" Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId in (" & BrnchIDes & ")"
  End If
  If BranshesActiv <> "-1" And BranshesActiv <> "" Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId in (" & BranshesActiv & ")"
  End If
  
   If BranchID <> 0 Then
  StrSQL = StrSQL & "   and dbo.Transactions.BranchId= " & BranchID
            If SystemOptions.UserInterface = ArabicInterface Then
                       branchtext = "«·›—⁄  : " & branchname
             Else
                         branchtext = "Branch  : " & branchname
             End If
  End If
  If GroupName <> "" Then
              If SystemOptions.UserInterface = ArabicInterface Then
                       GroupName = "«·„Ã„Ê⁄Â  : " & GroupName
             Else
                         GroupName = "Group  : " & GroupName
             End If
             
   End If
   
    If StoreName <> "" Then
              If SystemOptions.UserInterface = ArabicInterface Then
                       GroupName = "«·„” Êœ⁄  : " & StoreName
             Else
                         GroupName = "Store  : " & StoreName
             End If
             
   End If
  
  
If CurrenrEmployeeNames <> "" Then
 branchtext = branchtext & CHR(13) & CurrenrEmployeeNames
 Else
 If SystemOptions.UserInterface = ArabicInterface Then
 branchtext = branchtext & CHR(13) & "«·ﬂ·"
 Else
 branchtext = branchtext & CHR(13) & "Over All"
 End If
 End If
  If GroupName <> "" Then
branchtext = branchtext & CHR(13) & GroupName
End If
    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
    End If

    If Not IsNull(ToDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
    End If
 
 

 
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        Else
        Msg = "No Data"
        End If
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
If SystemOptions.UserInterface = ArabicInterface Then
    xReport.reporttitle = "    " & CHR(13) & branchtext
 Else
 xReport.reporttitle = "Detailed Items Profit" & CHR(13) & branchtext
 End If
    xReport.Database.SetDataSource RsData
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    'xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(3).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), CStr(FromDate))
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(ToDate), CDate("01/01/1000"), CStr(ToDate))
    xReport.ParameterFields(5).AddCurrentValue val(GetTotalSales(21, FromDate, ToDate, BranshesActiv, BrnchIDes, BranchID) - GetTotalSales(9, FromDate, ToDate, BranshesActiv, BrnchIDes, BranchID))
    xReport.ParameterFields(6).AddCurrentValue branchtext
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_Traget, xReport.reporttitle, , , , StrFileName, , sql
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub





'Public Sub ShowItemsAging(FromDate As Variant, _
'                          ToDate As Variant, _
'                          Optional m_Traget As PrintTarget = WindowTarget, _
'                          Optional reportno As Integer, _
'                          Optional GroupName As String = "", _
'                          Optional ItemName As String = "")
'
'    Dim xApp As New CRAXDRT.Application
'    Dim xReport As New CRAXDRT.Report
'    Dim StrSQL As String
'    Dim RsData As ADODB.Recordset
'    Dim StrRS As ADODB.Recordset
'    Dim Msg As String
'    Dim cReportViewer As ClsReportViewer
'    Dim StrFileName As String
'
'    On Error GoTo ErrTrap
'
'    If reportno = 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            StrFileName = App.path & "\Reports\Inventory\ItemAgenig.rpt"
'        ElseIf SystemOptions.UserInterface = EnglishInterface Then
'            StrFileName = App.path & "\Reports\Inventory\ItemAgenige.rpt"
'        End If
'
'    Else
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            StrFileName = App.path & "\Reports\Inventory\ItemAgenig2.rpt"
'        ElseIf SystemOptions.UserInterface = EnglishInterface Then
'            StrFileName = App.path & "\Reports\Inventory\ItemAgenig2E.rpt"
'        End If
'
'    End If
'
'    If Dir(StrFileName) = "" Then
'        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
'        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
'        Msg = Msg & StrFileName
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        Exit Sub
'    End If
'Set RsData = New ADODB.Recordset
'
'StrSQL = ""
'StrSQL = StrSQL & " SELECT dbo.TblTempItemAging.ItemId, "
'StrSQL = StrSQL & "        dbo.TblTempItemAging.LateID, "
'StrSQL = StrSQL & "        dbo.TblTempItemAging.Id, "
'StrSQL = StrSQL & "        dbo.TblTempItemAging.QtyBalance, "
'StrSQL = StrSQL & "        dbo.TblTempItemAging.BalanceValue, "
'StrSQL = StrSQL & "        dbo.TblTempItemAging.AvgCost, "
'StrSQL = StrSQL & "        dbo.TblItems.ItemCode, "
'StrSQL = StrSQL & "        dbo.TblItems.ItemName, "
'StrSQL = StrSQL & "        dbo.TblItems.ItemNamee, "
'StrSQL = StrSQL & "        dbo.AgengItem_type.Name, "
'StrSQL = StrSQL & "        dbo.AgengItem_type.namee "
'StrSQL = StrSQL & " FROM dbo.TblTempItemAging "
'StrSQL = StrSQL & " INNER JOIN dbo.TblItems "
'StrSQL = StrSQL & "     ON dbo.TblTempItemAging.ItemId = dbo.TblItems.ItemID "
'StrSQL = StrSQL & " INNER JOIN dbo.AgengItem_type "
'StrSQL = StrSQL & "     ON dbo.TblTempItemAging.LateID = dbo.AgengItem_type.Id "
'
'RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'     If RsData.BOF Or RsData.EOF Then
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
'        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'
'    Screen.MousePointer = vbArrowHourglass
'    Set xReport = xApp.OpenReport(StrFileName)
'
'  If SystemOptions.UserInterface = ArabicInterface Then
'    xReport.reporttitle = " ﬁ—Ì—   «·√’‰«› «·—«ﬂœ…"
'Else
'xReport.reporttitle = "Non Moving Items"
'End If
'    If GroupName <> "" Then
'      If SystemOptions.UserInterface = ArabicInterface Then
'
'        xReport.reporttitle = "„Ã„Ê⁄Â " & GroupName
'      Else
'      xReport.reporttitle = "Group " & GroupName
'      End If
'    End If
'
'    If ItemName <> "" Then
'          If SystemOptions.UserInterface = ArabicInterface Then
'
'        xReport.reporttitle = xReport.reporttitle & CHR(13) & "«·’„› " & ItemName
'        Else
'        xReport.reporttitle = xReport.reporttitle & CHR(13) & "Item " & ItemName
'        End If
'    End If
'
'    xReport.Database.SetDataSource RsData
'    Set StrRS = New ADODB.Recordset
'    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
'    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
'    xReport.ParameterFields(3).AddCurrentValue user_name
'    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), CStr(FromDate))
'    xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(ToDate), CDate("01/01/1000"), CStr(ToDate))
'    Set cReportViewer = New ClsReportViewer
'    cReportViewer.FireReport xReport, m_Traget, xReport.reporttitle, , , , StrFileName, , StrSQL
'    Screen.MousePointer = vbDefault
'    Exit Sub
'ErrTrap:
'    Screen.MousePointer = vbDefault
'End Sub
'

Public Sub ShowItemsAging(FromDate As Variant, _
                          ToDate As Variant, _
                          Optional m_Traget As PrintTarget = WindowTarget, _
                          Optional reportno As Integer, _
                          Optional GroupID As Long = 0, _
                          Optional Item_ID As Long = 0, _
                          Optional GroupName As String = "", _
                          Optional ItemName As String = "")

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim cReportViewer As ClsReportViewer
    Dim StrFileName As String

    On Error GoTo ErrTrap

    If reportno = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\Inventory\ItemAgenig.rpt"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            StrFileName = App.path & "\Reports\Inventory\ItemAgenige.rpt"
        End If
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\Inventory\ItemAgenig2.rpt"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            StrFileName = App.path & "\Reports\Inventory\ItemAgenig2E.rpt"
        End If
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset

StrSQL = ""
StrSQL = StrSQL & "EXEC dbo.Rpt_ItemAging "
StrSQL = StrSQL & SQLDate(CDate(FromDate), True) & ", "
StrSQL = StrSQL & SQLDate(CDate(ToDate), True) & ", "
StrSQL = StrSQL & GroupID & ", "
StrSQL = StrSQL & Item_ID

    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.reporttitle = " ﬁ—Ì— «·√’‰«› «·—«ﬂœ…"
    Else
        xReport.reporttitle = "Non Moving Items"
    End If

    If GroupName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = xReport.reporttitle & CHR(13) & "«·„Ã„Ê⁄… " & GroupName
        Else
            xReport.reporttitle = xReport.reporttitle & CHR(13) & "Group " & GroupName
        End If
    End If

    If ItemName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = xReport.reporttitle & CHR(13) & "«·’‰› " & ItemName
        Else
            xReport.reporttitle = xReport.reporttitle & CHR(13) & "Item " & ItemName
        End If
    End If

    xReport.Database.SetDataSource RsData

    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), CStr(FromDate))
    xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(ToDate), CDate("01/01/1000"), CStr(ToDate))

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_Traget, xReport.reporttitle, , , , StrFileName, , StrSQL

    Screen.MousePointer = vbDefault
    Exit Sub

ErrTrap:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation, App.Title
End Sub
Public Sub ShowStagnantItems(Optional StrStartDate As String = "", _
                             Optional IntReportStyle As Integer = 0)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim cReportViewer As ClsReportViewer
    Dim Msg As String
    Dim StrReportTitle As String
    Dim cOptions As ClsCompanyInfo

    Dim StrFileName As String
    On Error GoTo ErrTrap

    If IntReportStyle = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\StagnantItems.rpt"
        Else
            StrFileName = App.path & "\Reports\StagnantItems_Eng.rpt"
        End If

    ElseIf IntReportStyle = 1 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\StagnantItemsTree.rpt"
        Else
            StrFileName = App.path & "\Reports\StagnantItemsTree_Eng.rpt"
        End If
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If IntReportStyle = 0 Then
        StrSQL = "Select * From TempPrintStagnantItems Order By ID"
    Else
        StrSQL = "Select * From TempPrintStagnantItems " & "Order By ParentGroupID ASC,GroupID ASC,ItemID ASC"
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    If StrStartDate = "" Then
        Msg = IIf(SystemOptions.UserInterface = ArabicInterface, "»œ«Ì… «·»—‰«„Ã", "Programe Beginning")
    Else
        Msg = StrStartDate
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        StrReportTitle = " ﬁ—Ì— »«·√’‰«›  „‰– " & Msg & ""
    Else
        StrReportTitle = "Report By StagnantItems Since " & Msg & ""
    End If

    xReport.Database.SetDataSource RsData
    'xReport.OpenSubreport("SubReport1").Database.SetDataSource RsData
    xReport.reporttitle = StrReportTitle
    Set cOptions = New ClsCompanyInfo
    xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, WindowTarget, xReport.reporttitle, , , , StrFileName

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
    'Resume
End Sub

Public Sub ShowStockMovement(StrCaption As String, _
                             IntReportType As Integer, _
                             Optional IntReportStyle As Integer = 0)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim cReportViewer As ClsReportViewer
    Dim cOptions  As ClsCompanyInfo
    Dim Msg As String
    Dim StrReportTitle As String
    Dim StrFileName As String

    'IntReportType=0 Stock Movement Report
    'IntReportType=1 Stock Count Report

    'IntReportStyle=0 Table Report
    'IntReportStyle=1 Tree Report

    On Error GoTo ErrTrap

    If IntReportType = 0 Then
        If IntReportStyle = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\StockMovement.rpt"
            Else
                StrFileName = App.path & "\Reports\StockMovement_Eng.rpt"
            End If

        ElseIf IntReportStyle = 1 Then

            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\StockMovementTree.rpt"
            Else
                StrFileName = App.path & "\Reports\StockMovementTree_Eng.rpt"
            End If
        End If

    ElseIf IntReportType = 1 Then

        If IntReportStyle = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\StockMovementCount.rpt"
            Else
                StrFileName = App.path & "\Reports\StockMovementCount_Eng.rpt"
            End If

        ElseIf IntReportStyle = 1 Then

            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\StockMovementCountTree.rpt"
            Else
                StrFileName = App.path & "\Reports\StockMovementCountTree_Eng.rpt"
            End If
        End If
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If IntReportStyle = 0 Then
        StrSQL = "Select * From TempPrintStockMovement Order By ID"
    ElseIf IntReportStyle = 1 Then
        StrSQL = "Select * From RptTempPrintStockMovementTree"
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    If SystemOptions.UserInterface = ArabicInterface Then
        StrReportTitle = StrCaption
    Else
        StrReportTitle = StrCaption
    End If

    xReport.Database.SetDataSource RsData
    xReport.reporttitle = StrReportTitle
    Set cOptions = New ClsCompanyInfo
    xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, WindowTarget, xReport.reporttitle, , , , StrFileName

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
    'Resume
End Sub

Public Sub ManStockReports(LngStoreID As Long, _
                           IntType As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim cReportViewer As ClsReportViewer
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim StrReportTitle As String

    Dim StrFileName As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ManStockReport.rpt"
    Else
        StrFileName = App.path & "\Reports\ManStockReport.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    StrSQL = "SELECT     QryManStockComplete.QTY, QryManStockComplete.ItemID, QryManStockCompl" & "ete.ItemCode, QryManStockComplete.ItemName,                        QryManStockCo" & "mplete.HaveSerial, QryManStockComplete.ItemSerial, QryManStockComplete.TicketNO," & " QryManStockComplete.StoreID,                        QryManStockComplete.StoreNa" & "me, dbo.TblMaintenece.MaintananceID, dbo.TblMaintenece.DateGoIN, dbo.TblMaintene" & "ce.ReciptNumber,                        dbo.TblMaintenece.ManOperationTypeID FRO" & "M         dbo.QryManStockComplete(0) QryManStockComplete INNER JOIN             " & "          dbo.QryManLastTicketTrans ON QryManStockComplete.TicketNO = dbo.QryMan" & "LastTicketTrans.TicketNO INNER JOIN                       dbo.TblMaintenece ON d" & "bo.QryManLastTicketTrans.LastTrans = dbo.TblMaintenece.MaintananceID"
    StrSQL = StrSQL + " Where QryManStockComplete.StoreID=" & LngStoreID

    If IntType = 1 Then
        StrSQL = StrSQL + " AND dbo.TblMaintenece.ManOperationTypeID=1"
    ElseIf IntType = 3 Then
        StrSQL = StrSQL + " AND dbo.TblMaintenece.ManOperationTypeID=3"
    Else
    End If

    StrSQL = StrSQL + " Order By QryManStockComplete.TicketNO"
    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    StrReportTitle = " ﬁ—Ì— Ã—œ «·’Ì«‰…"
    xReport.Database.SetDataSource RsData
    xReport.reporttitle = StrReportTitle
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, WindowTarget, xReport.reporttitle, , , , StrFileName

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
    'Resume
End Sub

Public Sub SupManStock(LngCusID As Long)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim cReportViewer As ClsReportViewer
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim StrReportTitle As String

    Dim StrFileName As String
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\SupManStockReport.rpt"
    Else
        StrFileName = App.path & "\Reports\SupManStockReport.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    StrSQL = StrSQL + " SELECT * From RptSupManStock "

    If LngCusID <> 0 Then
        StrSQL = StrSQL + " Where RptSupManStock.CusID=" & LngCusID
    End If

    StrSQL = StrSQL + " Order By RptSupManStock.ItemName"
    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    StrReportTitle = " ﬁ—Ì— Ã—œ «·’Ì«‰…"
    xReport.Database.SetDataSource RsData
    xReport.reporttitle = StrReportTitle
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, WindowTarget, xReport.reporttitle, , , , StrFileName

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
    'Resume
End Sub

Public Function ShowItemTransCustomer(LngItemID As Long, _
                                      LngCusID As Long, _
                                      IntTransType As Integer, _
                                      LngStoreID As Long, _
                                      FromDate As Variant, _
                                      ToDate As Variant, _
                                      Optional m_Traget As PrintTarget = WindowTarget) As Boolean

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsData As ADODB.Recordset
    Dim cReportViewer As ClsReportViewer
    Dim i As Integer
    Dim Msg As String
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim cProgress As ClsProgress
    Dim cOptions As ClsCompanyInfo

    On Error GoTo ErrTrap

    '------------------------------
    If (LngCusID <> 0 And LngItemID <> 0) Then
        '’‰› „Õœœ „⁄ ⁄„Ì· «Ê „Ê—œ „Õœœ
        StrFileName = App.path & "\Reports\ItemsTransCus.rpt"
    ElseIf (LngCusID = 0 And LngItemID <> 0) Then
        '’‰› „Õœœ „⁄ Ã„Ì⁄ «·⁄„·«¡ «Ê «·„Ê—œÌ‰
        StrFileName = App.path & "\Reports\ItemsTransCusAll.rpt"
    ElseIf (LngCusID <> 0 And LngItemID = 0) Then
        '⁄„Ì· «Ê „Ê—œ „Õœœ „⁄ ﬂ· «·⁄„·«¡ «Ê «·„Ê—œÌ‰
        StrFileName = App.path & "\Reports\ItemsTransCusAllItems.rpt"
    ElseIf (LngCusID = 0 And LngItemID = 0) Then
        Msg = "»—Ã«¡ ≈»·«€ «·œ⁄„ «·›‰Ï »ÊÃÊœ  ÕœÌÀ €Ì— ﬂ«„· ›Ï «· ﬁ«—Ì—..!!!"
        Msg = Msg & CHR(13) & "LngCusID = 0 And LngItemID = 0"
        Msg = Msg & CHR(13) & "ClsItemsReports:ShowItemTransCustomer"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        ShowItemTransCustomer = False
        Exit Function
    End If

    '------------------------------
    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        ShowItemTransCustomer = False
        Exit Function
    End If

    StrSQL = "Select * From RptItemTransCus "
    StrSQL = StrSQL + " Where Transaction_Type=" & IntTransType & ""

    If LngItemID <> 0 Then
        StrSQL = StrSQL + " AND Item_ID=" & LngItemID & ""
    End If

    If LngCusID <> 0 Then
        StrSQL = StrSQL + " AND CusID=" & LngCusID & ""
    End If

    If LngStoreID <> 0 Then
        StrSQL = StrSQL + " AND StoreID=" & LngStoreID & ""
    End If

    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " AND Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
    End If

    If Not IsNull(ToDate) Then
        StrSQL = StrSQL + " AND Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
    End If

    StrSQL = StrSQL + " Order BY CusID,Item_ID,Transaction_ID"
    Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
    cProgress.StartProgress
    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute

    Do While RsData.State = adStateExecuting
        DoEvents
    Loop

    If RsData.BOF Or RsData.EOF Then
        cProgress.FinishProgress
        cProgress.StopProgess
        Set cProgress = Nothing
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        ShowItemTransCustomer = False
        Exit Function
    End If

    If IntTransType = 1 Then
        StrReportTitle = " ﬁ—Ì— »„‘ —Ì«   " & ""
    ElseIf IntTransType = 2 Then
        StrReportTitle = " ﬁ—Ì— »„»Ì⁄«   " & ""
    ElseIf IntTransType = 5 Then
        StrReportTitle = " ﬁ—Ì— „— Ã⁄ „‘ —Ì   " & ""
    ElseIf IntTransType = 9 Then
        StrReportTitle = " ﬁ—Ì— »„— Ã⁄ „»Ì⁄«   " & ""
    End If

    If LngItemID = 0 Then
        StrReportTitle = StrReportTitle & "«·√’‰«›"
    ElseIf LngItemID <> 0 Then
        StrReportTitle = StrReportTitle & "’‰›"
    End If

    If LngCusID <> 0 Then
        StrReportTitle = StrReportTitle & " „⁄ ⁄„Ì· «Ê „Ê—œ"
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    xReport.reporttitle = StrReportTitle
    Set cOptions = New ClsCompanyInfo
    xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), FromDate)
    xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(ToDate), CDate("01/01/1000"), ToDate)

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_Traget, xReport.reporttitle, , , , StrFileName

    If Not cProgress Is Nothing Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If

    Screen.MousePointer = vbDefault
    ShowItemTransCustomer = True
    Exit Function
ErrTrap:
    Screen.MousePointer = vbDefault
    ShowItemTransCustomer = False

End Function

Public Sub ShowGroupsItemsTrans(StrSQL As String, _
                                m_PrintTarget As PrintTarget, _
                                IntTransType As Integer, _
                                Optional FromDate As Variant = Null, _
                                Optional ToDate As Variant)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim RsData As ADODB.Recordset
    Dim cReportViewer As ClsReportViewer
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim StrReportTitle As String
    Dim cOptions As ClsCompanyInfo
    Dim StrFileName As String

    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\RptGroupsItemsTransactions.rpt"
    Else
        StrFileName = App.path & "\Reports\RptGroupsItemsTransactions_Eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)

    If SystemOptions.UserInterface = ArabicInterface Then
        If IntTransType = 1 Then
            StrReportTitle = " ﬁ—Ì— »«·„‘ —Ì«  „‰ „Ã„Ê⁄«  «·√’‰«›"
        ElseIf IntTransType = 2 Then
            StrReportTitle = " ﬁ—Ì— »«·„»Ì⁄«  „‰ „Ã„Ê⁄«  «·√’‰«›"
        ElseIf IntTransType = 9 Then
            StrReportTitle = " ﬁ—Ì— »„— Ã⁄ «·„»Ì⁄«  „‰ „Ã„Ê⁄«  «·√’‰«›"
        ElseIf IntTransType = 5 Then
            StrReportTitle = " ﬁ—Ì— „— Ã⁄ «·„‘ —Ì«   „‰ „Ã„Ê⁄«  «·√’‰«›"
        End If

    Else

        If IntTransType = 1 Then
            StrReportTitle = " ﬁ—Ì— »«·„‘ —Ì«  „‰ „Ã„Ê⁄«  «·√’‰«›"
        ElseIf IntTransType = 2 Then
            StrReportTitle = " ﬁ—Ì— »«·„»Ì⁄«  „‰ „Ã„Ê⁄«  «·√’‰«›"
        ElseIf IntTransType = 9 Then
            StrReportTitle = " ﬁ—Ì— »„— Ã⁄ «·„»Ì⁄«  „‰ „Ã„Ê⁄«  «·√’‰«›"
        ElseIf IntTransType = 5 Then
            StrReportTitle = " ﬁ—Ì— „— Ã⁄ «·„‘ —Ì«   „‰ „Ã„Ê⁄«  «·√’‰«›"
        End If
    End If

    xReport.Database.SetDataSource RsData
    xReport.reporttitle = StrReportTitle
    Set cOptions = New ClsCompanyInfo
 
    xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(FromDate), CDate("01/01/1000"), FromDate)
    xReport.ParameterFields(5).ClearCurrentValueAndRange
    xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(ToDate), CDate("01/01/1000"), ToDate)
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_PrintTarget, xReport.reporttitle, , , , StrFileName

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
    'Resume
End Sub

Public Sub ShowItemsInOutDes(LngGroupID As Long, _
                             LngStoreID As Long, _
                             FromDate As Variant, _
                             ToDate As Variant, _
                             Optional m_PrintTarget As PrintTarget = ToWindow)
    
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim RsBegine As ADODB.Recordset
    Dim RsEndQty As ADODB.Recordset
    Dim DblBegineQty As Double
    Dim DblEndQty As Double
    Dim RsData As ADODB.Recordset
    Dim FrmReport As FrmReportViewer
    Dim cReportViewer As ClsReportViewer
    Dim StrRS As ADODB.Recordset
    Dim i As Integer
    Dim Msg As String

    Dim StrFileName As String
    On Error GoTo ErrTrap

    StrFileName = App.path & "\Reports\RptItemsInOutDes.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        Exit Sub
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT  * FROM dbo.QryItemsInOutTransactionsDes('01- 01-2001'," & "' 01- 01-2010') QryItemsInOutTransactionsDes "
        StrSQL = StrSQL + " Where QryItemsInOutTransactionsDes.StoreID=" & LngStoreID

        If LngGroupID <> 0 Then
            StrSQL = StrSQL + " AND QryItemsInOutTransactionsDes.GroupID=" & LngGroupID
        End If
    End If
    
    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷ ...!!"
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_PrintTarget, xReport.reporttitle, , , , StrFileName

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
    'Resume
End Sub

