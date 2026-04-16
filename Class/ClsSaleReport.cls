VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSaleReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Public Function ShowComposedSallingDataDetailed(SallingID As Long, _
                                                Optional IntReportStyle As Integer = 0, _
                                                Optional LngCusID As Long, _
                                                Optional payed As Double, _
                                                Optional Result As Double, _
                                                Optional CusCode As String, _
                                                Optional BillComment As String, _
                                                Optional InstComm As Double, _
                                                Optional reportid As Integer, _
                                                Optional Nationality As String, _
                                                Optional GID As Double, _
                                                Optional Transaction_Date As String, _
                                                Optional AdvPayment As String, _
                                                Optional LblInstallCount As String, _
                                                Optional LblPrecenValue As String, _
                                                Optional LblFirstInstallDate As String, _
                                                Optional Madyna As String, _
                                                Optional hay As String, _
                                                Optional Address As String, _
                                                Optional total As String, _
                                                Optional BranchID As Integer, _
                                                Optional Expensee As String)
    On Error Resume Next
    Dim xApp            As New CRAXDRT.Application
    Dim xReport         As New CRAXDRT.Report
    Dim StrSQL          As String
    Dim StrRS           As New ADODB.Recordset
    Dim rs              As New ADODB.Recordset
    Dim RsTemp          As New ADODB.Recordset
    Dim ShowTax         As Boolean
    Dim Msg             As String
    Dim FrmReport       As FrmReportViewer
    Dim StrFileName     As String
    Dim SngCusBalBefore As Single
    Dim SngCusBalAfter  As Single
    Dim cOptions        As ClsCompanyInfo
    Dim CViewer         As ClsReportViewer
    Dim MySQL           As New ADODB.Recordset

    ' On Error GoTo ErrTrap

    If SallingID = 0 Then
        Exit Function
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        If IntReportStyle = 0 Then
            StrFileName = App.path & "\Reports\salesInvoiceComposed.rpt"
                
        ElseIf IntReportStyle = 1 Then
            '                StrFileName = App.path & "\Reports\salesInvoiceComposed.rpt"
        ElseIf IntReportStyle = 2 Then
            StrFileName = App.path & "\Reports\salesInvoiceComposed2.rpt"
 
        End If

    Else

        If IntReportStyle = 0 Then
            StrFileName = App.path & "\Reports\salesInvoiceEng.rpt"
                
        ElseIf IntReportStyle = 1 Then
            '              StrFileName = App.path & "\Reports\salesInvoiceEng.rpt"
        ElseIf IntReportStyle = 2 Then
            StrFileName = App.path & "\Reports\salesInvoiceComposed2.rpt"
        End If

    End If
 
    If reportid = 2 Then
        StrFileName = App.path & "\Reports\salesInvoiceComposed2.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    StrSQL = " SELECT     TOP 100 PERCENT "
    '********************************************
    StrSQL = StrSQL & "   BanksData.BankName BankName,"
    StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
    StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
    StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
    StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
    '*************************************************
    StrSQL = StrSQL & "                   dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID, "
    StrSQL = StrSQL & "                  TblCustemers_2.CusName, TblCustemers_2.CusNamee, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name,"
    StrSQL = StrSQL & "                  dbo.TblEmployee.Emp_Namee, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.Price, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemPhoto, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName,"
    StrSQL = StrSQL & "                  dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.CashCustomerName,"
    StrSQL = StrSQL & "                  dbo.Transactions.NoteSerial1, dbo.Transaction_Details.TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount,"
    StrSQL = StrSQL & "                  dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType, dbo.Transactions.ManualNo1, dbo.Transactions.PaymentType,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.order_no, dbo.Transaction_Details.Transaction_Date AS Transaction_DateSub, dbo.Transactions.Transaction_Serial,"
    StrSQL = StrSQL & "                  dbo.Transactions.Transaction_HijriDate, dbo.Transactions.VATYou, dbo.Transactions.Gratuity, dbo.Transactions.DriverName, dbo.Transactions.EmbarNo,"
    StrSQL = StrSQL & "                  dbo.Transactions.Commission, dbo.Transactions.BoardNo, dbo.Transactions.FarmID, TblCustemers_2.CusName AS VendCusName,"
    StrSQL = StrSQL & "                  TblCustemers_2.CusNamee AS VendCusNameE, TblCustemers_2.Fullcode, dbo.Transactions.DTPickerAccTo, dbo.Transactions.DTPickerAccFrom,"
    StrSQL = StrSQL & "                  dbo.Transactions.NetValueComm, dbo.Transactions.ValueComm, dbo.Transactions.TypComm, dbo.Transactions.Phone2, dbo.Transactions.GVAT,"
    StrSQL = StrSQL & "                  dbo.Transactions.SAlGValue, dbo.Transactions.GTotal, dbo.Transactions.ChLoadVAT, dbo.Transactions.VATNO, dbo.Transactions.VAT,"
    StrSQL = StrSQL & "                  dbo.Transactions.Transaction_NetValue, dbo.Transaction_Details.CusID2, TblCustemers_3.CusName AS DetCusName, TblCustemers_3.CusNamee AS DetCusNameE,"
    StrSQL = StrSQL & "                  TblCustemers_3.Fullcode AS DetFullcode, dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.Vat AS DetVAT, dbo.Transaction_Details.Vatyo AS DetVatyo,"
    StrSQL = StrSQL & "                  dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee,"
    StrSQL = StrSQL & "                  TblCustemers_1.CusName AS FarmName, TblCustemers_1.CusNamee AS FarmNameE, TblCustemers_1.Fullcode AS FarmFullcode ,dbo.Transactions.FactoryExpenses,"
    
    StrSQL = StrSQL & "                     Transaction_Details.EmpID4 ,Transaction_Details.CusID2 ,Transaction_Details.SupplierID,"
    StrSQL = StrSQL & "                      EmpName4 = (Select Top 1 te.Emp_Name FROM TblEmployee te WHERE te.Emp_ID= Transaction_Details.EmpID4), "
    StrSQL = StrSQL & "                      CustName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.CusID2), "
    StrSQL = StrSQL & "                      SuppName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.SupplierID) "
    StrSQL = StrSQL & "                      ,transactions.PurchOrderNo Poa, SuppName = (Select top 1 cc.cusName from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29),"
    StrSQL = StrSQL & "                      SuppAddress= (Select top 1 cc.Address from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29)"

    StrSQL = StrSQL & "      FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "                  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers TblCustemers_2 ON dbo.Transaction_Details.CusID2 = TblCustemers_2.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers TblCustemers_3 ON dbo.Transactions.CusID = TblCustemers_3.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers TblCustemers_1 ON dbo.Transactions.FarmID = TblCustemers_1.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
    '***********************************
    StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
    StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
    StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
    StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
    '**************************************
    StrSQL = StrSQL & "  Where (dbo.Transactions.Transaction_ID = " & SallingID & ")"
    StrSQL = StrSQL & "  ORDER BY dbo.Transaction_Details.ID"

    If reportid = 2 Then
 
        StrSQL = " SELECT     TOP 100 PERCENT MAX(dbo.Transactions.Transaction_Date) AS Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID,"
        StrSQL = StrSQL & "  dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID,"
        StrSQL = StrSQL & "  dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Transaction_Details.Item_ID, SUM(dbo.Transaction_Details.ShowQty) AS ShowQty,"
        StrSQL = StrSQL & "  AVG(dbo.Transaction_Details.showPrice) AS showPrice, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId,"
        StrSQL = StrSQL & "  SUM(dbo.Transaction_Details.discountvalue) AS discountvalue, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName,"
        StrSQL = StrSQL & "  dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS classname, dbo.Transactions.CashCustomerName,"
        StrSQL = StrSQL & "  dbo.Transactions.NoteSerial1, SUM(dbo.Transaction_Details.TotalDiscountPerLine) AS TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType,"
        StrSQL = StrSQL & "  SUM(dbo.Transaction_Details.ItemDiscount) AS ItemDiscount, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
        StrSQL = StrSQL & "  dbo.Transactions.PaymentType,"
        
        StrSQL = StrSQL & "                     Transaction_Details.EmpID4 ,Transaction_Details.CusID2 ,Transaction_Details.SupplierID,"
        StrSQL = StrSQL & "                      EmpName4 = (Select Top 1 te.Emp_Name FROM TblEmployee te WHERE te.Emp_ID= Transaction_Details.EmpID4), "
        StrSQL = StrSQL & "                      CustName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.CusID2), "
        StrSQL = StrSQL & "                      SuppName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.SupplierID) "
        StrSQL = StrSQL & "                      ,transactions.PurchOrderNo Poa, SuppName = (Select top 1 cc.cusName from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29),"
        StrSQL = StrSQL & "                      SuppAddress= (Select top 1 cc.Address from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29)"

        StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
        StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
        StrSQL = StrSQL & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
        StrSQL = StrSQL & "  dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID INNER JOIN"
        StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID INNER JOIN"
        StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
        StrSQL = StrSQL & "  GROUP BY dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.UserID,"
        StrSQL = StrSQL & "  dbo.TblUsers.UserName, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Transaction_Details.Item_ID,"
        StrSQL = StrSQL & "  dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
        StrSQL = StrSQL & "  dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName,"
        StrSQL = StrSQL & "  dbo.Transactions.CashCustomerName, dbo.Transactions.NoteSerial1, dbo.Transaction_Details.ItemDiscountType, dbo.Transactions.Trans_Discount,"
        StrSQL = StrSQL & "  dbo.Transactions.Trans_DiscountType , dbo.Transactions.PaymentType"
        StrSQL = StrSQL & "   Having (dbo.Transactions.Transaction_ID = " & SallingID & ")"

    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Screen.MousePointer = vbArrowHourglass
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    Set cOptions = New ClsCompanyInfo
    xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(5).AddCurrentValue BillComment

    If val(InstComm) >= 0 Then
        xReport.ParameterFields(10).AddCurrentValue val(InstComm)
    End If

    '

    xReport.ParameterFields(8).AddCurrentValue WriteNo(CStr(Result), 0, True)
    xReport.ParameterFields(9).AddCurrentValue CusCode

    If IntReportStyle = 0 Then
        'xReport.ParameterFields(6).AddCurrentValue payed
        'xReport.ParameterFields(7).AddCurrentValue result
    End If

    If reportid = 1 Then
        xReport.ParameterFields(11).AddCurrentValue Nationality
        xReport.ParameterFields(12).AddCurrentValue CStr(GID)
        xReport.ParameterFields(13).AddCurrentValue ToHijriDate(Transaction_Date)

        xReport.ParameterFields(14).AddCurrentValue CStr(AdvPayment)
        xReport.ParameterFields(15).AddCurrentValue CStr(LblInstallCount)
        xReport.ParameterFields(16).AddCurrentValue CStr(LblPrecenValue)
        xReport.ParameterFields(17).AddCurrentValue CStr(LblFirstInstallDate)
  
    End If
    If SystemOptions.VATNoAccordActivity = False Then
        xReport.ParameterFields(18).AddCurrentValue cOptions.VATRegNo
    Else
        xReport.ParameterFields(18).AddCurrentValue GetRegVATNo(BranchID)
    End If
    xReport.ParameterFields(19).AddCurrentValue CStr(Expensee)
    If reportid = 2 Then
        'xReport.ParameterFields(11).AddCurrentValue nationality
        'xReport.ParameterFields(12).AddCurrentValue CStr(gid)
        'xReport.ParameterFields(13).AddCurrentValue ToHijriDate(Transaction_Date)

        'xReport.ParameterFields(14).AddCurrentValue CStr(AdvPayment)
        'xReport.ParameterFields(15).AddCurrentValue CStr(LblInstallCount)
        'xReport.ParameterFields(16).AddCurrentValue CStr(LblPrecenValue)
        'xReport.ParameterFields(17).AddCurrentValue CStr(LblFirstInstallDate)
 
        ' xReport.ParameterFields(18).AddCurrentValue CStr(Madyna)
        ' xReport.ParameterFields(19).AddCurrentValue CStr(hay)
        ' xReport.ParameterFields(20).AddCurrentValue CStr(Address)
        ' xReport.ParameterFields(21).AddCurrentValue CStr(total)
        ' xReport.ParameterFields(22).AddCurrentValue CStr(WriteNo(Format(total, "0.00"), 0, True, ".", , 0))
  
    End If

    'xReport.ParameterFields(6).AddCurrentValue "2"
    'xReport.ParameterFields(7).AddCurrentValue "1"

    xReport.reporttitle = "ÿ»«⁄… ›« Ê—… »Ì⁄ „Ã„⁄Â"
 
    Screen.MousePointer = vbDefault
    Set CViewer = New ClsReportViewer
    Dim P_Target As PrintTarget

    If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
        P_Target = PrinterTarget
    Else
        P_Target = WindowTarget
    End If

    CViewer.FireReport xReport, P_Target, , , , , StrFileName
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function


 

  

 
 
     

 

   
    




   Public Function DestructionReport(TransID As Double, Optional reportid As Integer = 0)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim Msg As String
    Dim FrmReport As FrmReportViewer
Dim FileName As String

    On Error GoTo ErrTrap

    If TransID = Null Then
        Exit Function
    End If

StrSQL = " SELECT       ProductionDate,ExpiryDate  , dbo.Transaction_Details.PrintName, dbo.Transaction_Details.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName, "
StrSQL = StrSQL & "                       dbo.TblStore.StoreName, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemCode, dbo.Transaction_Details.ItemCase,"
StrSQL = StrSQL & "                       dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transaction_Details.ItemDiscountType,"
StrSQL = StrSQL & "                       dbo.Transaction_Details.ItemDiscount, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType, dbo.Transactions.TaxFound,"
StrSQL = StrSQL & "                       dbo.Transactions.TaxValue, dbo.TblItems.HaveSerial, dbo.Transactions.Transaction_Serial, dbo.Transaction_Details.guaranteeTime, dbo.TblEmployee.Emp_Code,"
StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Name, dbo.Transactions.PaymentType, dbo.Transactions.CashCustomerName, dbo.Transaction_Details.Remarks,"
StrSQL = StrSQL & "                       dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transaction_Details.sallReturnPrice, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "                       dbo.Transactions.NoteSerial1, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.Tel,"
StrSQL = StrSQL & "                       dbo.Transaction_Details.project_ID1, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.Transaction_Details.Pand_ID, dbo.projects_des.des,"
StrSQL = StrSQL & "                       dbo.Transaction_Details.Oper_ID, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE, dbo.Transaction_Details.ID, dbo.TblItems.ItemNamee,"
StrSQL = StrSQL & "                       dbo.TblUnites.UnitNamee , dbo.TblStore.storenamee, dbo.transactions.ManualNO, dbo.transactions.TransactionComment"
StrSQL = StrSQL & "  FROM         dbo.projects RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblItems INNER JOIN"
StrSQL = StrSQL & "                       dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblStore INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblCustemers RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID ON dbo.TblStore.StoreID = dbo.Transactions.StoreID ON"
StrSQL = StrSQL & "                       dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblProcessDEF ON dbo.Transaction_Details.Oper_ID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.projects_des ON dbo.Transaction_Details.Pand_ID = dbo.projects_des.oprid AND dbo.projects_des.oprid <> 0 ON"
StrSQL = StrSQL & "                       dbo.projects.id = dbo.Transaction_Details.project_ID1 LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "  where Transactions.Transaction_ID=" & TransID
 StrSQL = StrSQL & " order by Transaction_Details.id "
    
    
    
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
If reportid = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
            FileName = App.path & "\Reports\Destruction.rpt"
            Else
            FileName = App.path & "\Reports\DestructionE.rpt"
            End If
ElseIf reportid = 55 Then
            If SystemOptions.UserInterface = ArabicInterface Then
            FileName = App.path & "\Reports\RetDestruction.rpt"
            Else
            FileName = App.path & "\Reports\RetDestructionE.rpt"
            End If
Else


            If SystemOptions.UserInterface = ArabicInterface Then
            FileName = App.path & "\Reports\StockSettl.rpt"
            Else
            FileName = App.path & "\Reports\StockSettle.rpt"
            End If


End If

    
    
    If Dir(FileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\Destruction.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        
        Screen.MousePointer = vbArrowHourglass
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(FileName)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.txtPath = FileName
        FrmReport.TXTSTRSQL = StrSQL
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.show

    
    Screen.MousePointer = vbDefault
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function '.rpt


' ﬁ—Ì— «·«Â·«ﬂ« 
  

' ﬁ—Ì— «· ÕÊÌ· „‰ „Œ“‰ «·Ï „Œ“‰
Public Function MovingItems(TransID As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim Msg As String
    Dim FrmReport As FrmReportViewer

    On Error GoTo ErrTrap

    If TransID = Null Then
        Exit Function
    End If

    StrSQL = "select * From QryMovingItemsReport where Transaction_ID=" & TransID

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If Dir(App.path & "\Reports\MovingItems.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\MovingItems.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\MovingItems.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function 'MovingItems.rpt

Public Function ShowSallingDataShort(SallingID As Long, _
                                     Optional m_Target As PrintTarget = WindowTarget, _
                                     Optional PrinterType As Integer = 0)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim Frm As FrmReportViewer
    Dim CViewer As ClsReportViewer
    Dim Msg As String

    Dim StrFileName As String
    On Error GoTo ErrTrap

    If SallingID = Null Then
        Exit Function
    End If

    StrSQL = "select * From QryBuyReportShort where Transaction_ID=" & SallingID
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If PrinterType = 0 Then
        StrFileName = App.path & "\Reports\SaleReportShort1.rpt"
    Else
        StrFileName = App.path & "\Reports\SaleReportShort_Cashier.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\SaleReportShort.rpt"
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
        xReport.ParameterFields(3).AddCurrentValue user_name
        StrSQL = "select * From Notes where Transaction_ID=" & SallingID
        StrSQL = StrSQL + " and NoteType=0"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
        End If

        Screen.MousePointer = vbDefault
    End If

    Set Frm = New FrmReportViewer
    Frm.CRViewer.ReportSource = xReport

    If m_Target = WindowTarget Then
        Frm.CRViewer.viewReport
        Frm.WindowState = vbMaximized

        DoEvents
        Frm.show
        Sendkeys "{RIGHT}"
    ElseIf m_Target = PrinterTarget Then
        xReport.PaperSize = crPaper10x14
        Set CViewer = New ClsReportViewer
        CViewer.FireReport xReport, PrinterTarget, "", , 1, , StrFileName, , StrSQL
        Unload Frm
    End If

    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Function PrintInvoiceReceipt(SallingID As Long, _
                                    Optional m_Target As PrintTarget = WindowTarget)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim Frm As FrmReportViewer
    Dim CViewer As ClsReportViewer
    Dim Msg As String
    Dim StrTemp As String

    Dim StrFileName As String
    On Error GoTo ErrTrap

    If SallingID = Null Then
        Exit Function
    End If

    'If SystemOptions.SysDataBaseType = AccessDataBase Then
    StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, TblNotesTypes.NotesTypeName," & "Notes.NoteSerial, Notes.Note_Value, Notes.BankID, BanksData.BankName, Notes.ChqueNum," & "Notes.DueDate, Transactions.Transaction_Serial, TransactionTypes.TransactionTypeName," & "Transactions.Transaction_ID, TblCustemers.CusName, TblEmployee.Emp_Name "
    StrSQL = StrSQL + " FROM TransactionTypes INNER JOIN ((TblEmployee INNER JOIN " & "(TblCustemers INNER JOIN Transactions ON TblCustemers.CusID = Transactions.CusID) " & "ON TblEmployee.Emp_ID = Transactions.Emp_ID) INNER JOIN (TblNotesTypes INNER JOIN " & "(BanksData RIGHT JOIN Notes ON BanksData.BankID = Notes.BankID) ON TblNotesTypes.NotesType" & " = Notes.NoteType) ON Transactions.Transaction_ID = Notes.Transaction_ID) ON " & "TransactionTypes.Transaction_Type = Transactions.Transaction_Type "
    StrSQL = StrSQL + " WHERE (((Notes.NoteType)=0 Or (Notes.NoteType)=2 Or (Notes.NoteType)=13 Or (Notes.NoteType)=14))"
    StrSQL = StrSQL + " AND Transactions.Transaction_ID=" & SallingID

    'End If
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    StrFileName = App.path & "\Reports\InvoiceReceipt.rpt"

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
        StrTemp = StrTemp & " " & " œÌ‰«— ·»ÌÌ ›ﬁÿ ·«€Ì— "
        xReport.ParameterFields(3).AddCurrentValue StrTemp
        '        xReport.ParameterFields(3).AddCurrentValue User_Name
        '        StrSQL = "select * From Notes where Transaction_ID=" & SallingID
        '        StrSQL = StrSQL + " and NoteType=0"
        '        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        '        If Not (RsTemp.EOF Or RsTemp.BOF) Then
        '            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").Value), "", Trim(RsTemp("Note_Value").Value))
        '        Else
        '             xReport.ParameterFields(4).AddCurrentValue ""
        '        End If
        Screen.MousePointer = vbDefault
    End If

    Set Frm = New FrmReportViewer
    Frm.CRViewer.ReportSource = xReport
Frm.TXTSTRSQL = StrSQL
    If m_Target = WindowTarget Then
        Frm.CRViewer.viewReport
        Frm.WindowState = vbMaximized

        DoEvents
        Frm.show
        Sendkeys "{RIGHT}"
    ElseIf m_Target = PrinterTarget Then
        xReport.PaperSize = crPaper10x14
        Set CViewer = New ClsReportViewer
        CViewer.FireReport xReport, PrinterTarget, "", , 1, , StrFileName, , StrSQL
        Unload Frm
    End If

    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function
  
Public Function ShowPrice(PriceID As Double, _
                          Optional reportid As Integer, _
                          Optional EmpName As String, _
                          Optional CUSTID As Double, _
                          Optional stringItems As String, _
                          Optional casingvalues As String, _
                          Optional totals As Double, _
                          Optional ByRef TDate As Date, _
                          Optional ByRef tdate1 As Date, _
                          Optional StrDaeen As String, _
                          Optional xx As Integer = 0)
    Dim xApp      As New CRAXDRT.Application
    Dim xReport   As New CRAXDRT.Report
    Dim StrSQL    As String
    Dim StrRS     As New ADODB.Recordset
    Dim rs        As New ADODB.Recordset
    Dim RsTemp    As New ADODB.Recordset
    Dim Msg       As String
    Dim FrmReport As FrmReportViewer
    On Error Resume Next
    'On Error GoTo ErrTrap

    If PriceID = Null Then
        Exit Function
    End If
 
    'UPDATE 29 10 2017
    StrSQL = " SELECT "
    '********************************************
    StrSQL = StrSQL & "   BanksData.BankName BankName,"
    StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
    StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
    StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
    StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
    '*************************************************
    StrSQL = StrSQL & "   TblCustemers.VATNO  ,Transactions.TotalTaxExempt  , transactions.QrCodeImage,transactions.ExtraValue,Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width,Transactions.PaymentConditions,Transactions.DeliveryMethod,Transactions.Packing, dbo.Transaction_Details.Height, dbo.Transaction_Details.Area, dbo.Transaction_Details.Remarks, "
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ProductionDate, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.LotNO, dbo.Transactions.Trans_Discount,TblItems.barCodeNO,"
    StrSQL = StrSQL & "                    dbo.Transactions.Trans_DiscountType, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_ID,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.order_no, dbo.Transactions.Currency_id,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ItemSize,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
    StrSQL = StrSQL & "                    dbo.TblItems.ItemNamee, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName,"
    StrSQL = StrSQL & "                    dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusNamee, dbo.Transactions.Transaction_Type, dbo.TblCustemers.Fullcode, dbo.TblCustemers.E_mail,"
    StrSQL = StrSQL & "                    dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.FaxNumber, dbo.Transaction_Details.ParrtNoCode, TblUnites_2.UnitNamee,"
    StrSQL = StrSQL & "                    dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.DepartementID, TblEmpDepartments_2.DepartmentName,"
    StrSQL = StrSQL & "                    TblEmpDepartments_2.DepartmentNamee, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transaction_Details.RequestLimit,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.LastPurchasePrice, dbo.Transaction_Details.LastPurchaseqty, dbo.Transaction_Details.AverageIssue,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.AverageIssueyraly, dbo.Transaction_Details.LastPurchaseDate, dbo.Transactions.InternalFlag, dbo.Transactions.FixesAssetsID,"
    StrSQL = StrSQL & "                    dbo.FixedAssets.code AS Fixedcode, dbo.FixedAssets.Name, dbo.FixedAssets.namee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name1,"
    StrSQL = StrSQL & "                    dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
    StrSQL = StrSQL & "                    dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblItems.Fullcode AS ItemFullcode,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.ItemCase, dbo.markaas_taklefa.account_name, dbo.markaas_taklefa.Code AS TaklfaCode,"
    StrSQL = StrSQL & "                    dbo.Transactions.CBoBasedON, dbo.Transactions.OrderType, dbo.Transactions.DeptID, TblEmpDepartments_1.DepartmentName AS DepartmentName2,"
    StrSQL = StrSQL & "                    TblEmpDepartments_1.DepartmentNamee AS DepartmentName2E, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_HijriDate,"
    StrSQL = StrSQL & "                    dbo.Transactions.Shipment_no, dbo.Transactions.shipped, dbo.Transactions.shipmentMethod, dbo.Transactions.ShipmentPrice, dbo.Transactions.ShipmentArae,"
    StrSQL = StrSQL & "                    dbo.Transactions.ShipmentStatus, dbo.Transactions.SippingDate, dbo.Transactions.shipmentType, dbo.Transactions.ShippingTypeID,"
    StrSQL = StrSQL & "                    dbo.Transactions.ShippingStatus, dbo.Transactions.Shipping_Pos, dbo.Transactions.PaymentID, dbo.TblPaymentType.PaymentName,"
    StrSQL = StrSQL & "                    dbo.TblPaymentType.PaymentNamee, dbo.Transactions.ShipingID, dbo.TblShipingData.Name AS ShipName, dbo.TblShipingData.NameE AS ShipNameE,"
    StrSQL = StrSQL & "                    dbo.Transactions.DeliverDate, dbo.TblCustemers.CusID, dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone,"
    StrSQL = StrSQL & "                    dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress, dbo.Transactions.CashCustomerComment,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.Remarks AS RemarksDet, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area AS AreaDet,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.Height AS HeightDet, dbo.Transaction_Details.Width AS WidthDet, Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width, dbo.Transaction_Details.ItemsDetailsNewidea,"
    StrSQL = StrSQL & "                    dbo.Transactions.GrantyPeriod, dbo.Transactions.NoDelivery, dbo.Transactions.Granty, dbo.Transactions.Cityid, dbo.Transactions.Days1,"
    StrSQL = StrSQL & "                    dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblCountriesGovernmentsCities.CityName, dbo.TblCountriesGovernmentsCities.NameE AS CityNameE,"
    StrSQL = StrSQL & "                    dbo.Transactions.GovernmentID, dbo.TblCountriesGovernments.GovernmentName, dbo.Transactions.PaymentType, dbo.Transactions.UserID, dbo.TblUsers.UserSign,"
    StrSQL = StrSQL & "                    dbo.TblUsers.UserName, dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.Groups.GroupCode, dbo.Transactions.Transaction_NetValue,"
    StrSQL = StrSQL & "                    dbo.Transactions.TransactionComment, dbo.TblCustemers.CusName, dbo.TblEmployee.Emp_Name, dbo.Transactions.Vat, dbo.Transaction_Details.Vatyo,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.MaxQty, dbo.Transaction_Details.MaxUnitID, TblUnites_1.UnitName AS MaxUnitName, TblUnites_1.UnitNamee AS MaxUnitNameE,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ManufactCompany, dbo.Transaction_Details.CountryOrg, TblUnites_2.UnitName, dbo.Transaction_Details.PrintName,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ItemBalance ,dbo.Transactions.CusBalance,"
    StrSQL = StrSQL & "                    Cur.Name as CurrencyName , Cur.Namee as CurrencyNamee,TblItems.itemphoto ,TblItems.ItemComment"
  
    StrSQL = StrSQL & "                    , Transaction_Details.EmpID4 ,Transaction_Details.CusID2 ,Transaction_Details.SupplierID,Transactions.QrCodeImage,"
    StrSQL = StrSQL & "                      EmpName4 = (Select Top 1 te.Emp_Name FROM TblEmployee te WHERE te.Emp_ID= Transaction_Details.EmpID4), "
    StrSQL = StrSQL & "                      CustName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.CusID2), "
    StrSQL = StrSQL & "                      SuppName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.SupplierID) "
    StrSQL = StrSQL & "                      ,transactions.PurchOrderNo Poa, SuppName = (Select top 1 cc.cusName from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29),"
    StrSQL = StrSQL & "                      SuppAddress= (Select top 1 cc.Address from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29),TblItems.ItemWithOutVAT,"
    StrSQL = StrSQL & "                      transactions.PONo,TblCustemers.ResponsibleContact,TblCustemers.Cus_mobile,TblCustemers.Cus_Phone,TblCustemers.VATNO,TblEmployee.Emp_mobile,Transaction_Details.id as MLineNo,IsNull(Transaction_Details.LineID,Transaction_Details.id) as LineID"
    
    StrSQL = StrSQL & "                      ,transactions.PONo , transactions.PaymentConditions, transactions.SumValueLine, transactions.SumVATLine"
  'ResponsibleContact,Cus_mobile,Cus_Phone,E_mail,Address,VATNO,PONo
  
    StrSQL = StrSQL & "     FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "                    dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblUnites TblUnites_1 ON dbo.Transaction_Details.MaxUnitID = TblUnites_1.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblCountriesGovernments ON dbo.Transactions.GovernmentID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblCountriesGovernmentsCities ON dbo.Transactions.Cityid = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblShipingData ON dbo.Transactions.ShipingID = dbo.TblShipingData.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblPaymentType ON dbo.Transactions.PaymentID = dbo.TblPaymentType.PaymentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblEmpDepartments TblEmpDepartments_1 ON dbo.Transactions.DeptID = TblEmpDepartments_1.DeparmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.markaas_taklefa ON dbo.Transactions.general_cost_center = dbo.markaas_taklefa.Code LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.FixedAssets ON dbo.Transactions.FixesAssetsID = dbo.FixedAssets.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblEmpDepartments TblEmpDepartments_2 ON dbo.Transactions.DepartementID = TblEmpDepartments_2.DeparmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblUnites TblUnites_2 ON dbo.Transaction_Details.UnitId = TblUnites_2.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL & "                        LEFT OUTER JOIN dbo.currency Cur ON dbo.Transactions.Currency_id = Cur.ID"
    '***********************************
    StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
    StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
    StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
    StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
    '**************************************
  
    StrSQL = StrSQL & "  Where (dbo.Transactions.Transaction_ID = " & PriceID & ")"
    'UPDATE 29 10 2017

    If reportid = 100 Or reportid = 12 Then
        StrSQL = "SELECT     "
        '********************************************
        StrSQL = StrSQL & "   BanksData.BankName BankName,Transactions.TotalTaxExempt  ,"
        StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
        StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
        StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
        StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
        '*************************************************
        StrSQL = StrSQL & "   dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.TransactionsCreteria.Enteredvalue, dbo.TransactionsCreteria.remarks, "
        StrSQL = StrSQL & " dbo.Transactions.NoteSerial1, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile,"
        StrSQL = StrSQL & " dbo.TblCustemers.Address , dbo.TransactionsCreteria.CreteriaID, dbo.tblScreenCriteria.name, dbo.tblScreenCriteria.namee,  dbo.Transactions.Days1,dbo.Transactions.Days2 "
        StrSQL = StrSQL & ", dbo.TransactionsCreteria.remarksE  FROM         dbo.Transactions INNER JOIN"
        StrSQL = StrSQL & " dbo.TransactionsCreteria ON dbo.Transactions.Transaction_ID = dbo.TransactionsCreteria.Transaction_ID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
        StrSQL = StrSQL & "   dbo.tblScreenCriteria ON dbo.TransactionsCreteria.CreteriaID = dbo.tblScreenCriteria.CriteriaID"
        '***********************************
        StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
        StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
        StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
        StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
        '**************************************
        StrSQL = StrSQL & "  WHERE     (dbo.Transactions.Transaction_ID = " & PriceID & ")"
 
    End If

    If reportid = 10 Or reportid = 11 Then
        StrSQL = "  SELECT    Transactions.TotalTaxExempt  ,"
        '********************************************
        StrSQL = StrSQL & "   BanksData.BankName BankName,"
        StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
        StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
        StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
        StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
        '*************************************************
        StrSQL = StrSQL & "    dbo.TBLRegularMaint.itemid, dbo.TBLRegularMaint.DateOfRegularMaint, dbo.TBLRegularMaint.GranteeType, "
        StrSQL = StrSQL & " dbo.TBLRegularMaint.GranteeStartDate, dbo.TBLRegularMaint.Done, dbo.TBLRegularMaint.DoneDate, dbo.TBLRegularMaint.DoneEmpID,"
        StrSQL = StrSQL & " dbo.TBLRegularMaint.ItemSerial, dbo.TBLRegularMaint.MaintenanceIDS, dbo.TBLRegularMaint.Transaction_ID, dbo.Transactions.Transaction_Date,"
        StrSQL = StrSQL & " dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
        StrSQL = StrSQL & " dbo.TblMaintenanceType.name AS maintenancenameA, dbo.TblMaintenanceType.namee AS maintenancenameE, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
        StrSQL = StrSQL & " dbo.TblItems.ItemNamee, dbo.TblMaintenanceType.REMARKS, dbo.TblMaintenanceType.[interval], dbo.TblMaintenanceType.intervaltype,"
        StrSQL = StrSQL & " dbo.TBLRegularMaint.[count] , dbo.Transactions.NoteSerial1, dbo.TblMaintenanceType.id, dbo.TBLRegularMaint.GranteeEndDate"
        StrSQL = StrSQL & "  ,   dbo.Transactions.TransactionComment  FROM         dbo.TblMaintenanceType RIGHT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TBLRegularMaint ON dbo.TblMaintenanceType.id = dbo.TBLRegularMaint.MaintenanceIDS LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblItems ON dbo.TBLRegularMaint.itemid = dbo.TblItems.ItemID LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblCustemers RIGHT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID ON dbo.TBLRegularMaint.Transaction_ID = dbo.Transactions.Transaction_ID"
    
        '***********************************
        StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
        StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
        StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
        StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
        '**************************************

        StrSQL = StrSQL & " WHERE     (dbo.Transactions.Transaction_ID = " & PriceID & ")"
        StrSQL = StrSQL & " ORDER BY dbo.TblMaintenanceType.id "
    End If

    If reportid = 61 Then

        Dim StrAccountCode As String
        StrAccountCode = "SELECT     dbo.TblCustemers.Account_Code"
        StrAccountCode = StrAccountCode & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
        StrAccountCode = StrAccountCode & "  dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID"
        StrAccountCode = StrAccountCode & " WHERE     (dbo.Transactions.Transaction_Type = 61) AND (dbo.Transactions.Transaction_Date =" & SQLDate(CDate(TDate), True) & ")"
        StrAccountCode = StrAccountCode & " GROUP BY dbo.TblCustemers.Account_Code"
        Dim FirstPeriodDateInthisYear As Date
        getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
  
        ' updateopeningbalanceNewFromsql XPDtbFrom.value, XPDtpTo.value, True, , val(Me.Dcbranch.BoundText), StrAccountCode, 30
            
        updateopeningbalanceNewFromsqlTrialBalance2 FirstPeriodDateInthisYear, TDate, True, , , StrAccountCode, 3
    
        StrSQL = "   SELECT     Transactions.TotalTaxExempt  , "
        '********************************************
        StrSQL = StrSQL & "   BanksData.BankName BankName,"
        StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
        StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
        StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
        StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
        '*************************************************
        StrSQL = StrSQL & "  dbo.Transactions.FixesAssetsID, dbo.TblEquipments.Code, dbo.TblEquipments.name, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, "
        StrSQL = StrSQL & "    dbo.TblProductionType.name AS productiontypename, dbo.TblProductionType.namee AS productiontypenamee, dbo.Transactions.ContactTime,"
        StrSQL = StrSQL & "   dbo.Transaction_Details.ShowQty, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
        StrSQL = StrSQL & "    dbo.Transactions.TransactionComment, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
        StrSQL = StrSQL & "   dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Remark,"
        StrSQL = StrSQL & "      dbo.Transactions.Transaction_Date, dbo.TblCountriesGovernmentsCities.CityName AS HeyName,TblItems.ItemComment "
        StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
        StrSQL = StrSQL & "      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
        StrSQL = StrSQL & "    dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
        StrSQL = StrSQL & "    dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
        StrSQL = StrSQL & "        dbo.TblItems TblItems_1 ON dbo.Transaction_Details.Item_ID = TblItems_1.ItemID LEFT OUTER JOIN"
        StrSQL = StrSQL & "      dbo.TblCountriesGovernmentsCities ON dbo.Transactions.Neighborhoodid = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
        StrSQL = StrSQL & "       dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "      dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
        StrSQL = StrSQL & "     dbo.TblProductionType ON dbo.Transactions.ProductionTypeid = dbo.TblProductionType.Id LEFT OUTER JOIN"
        StrSQL = StrSQL & "     dbo.TblEquipments ON dbo.Transactions.FixesAssetsID = dbo.TblEquipments.fixedAssetid"
        '***********************************
        StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
        StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
        StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
        StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
        '**************************************
    
        StrSQL = StrSQL & " WHERE     ( Transaction_Type=61 and dbo.Transactions.Transaction_Date =" & SQLDate(CDate(TDate), True) & ")"

        StrSQL = "SELECT   Transactions.TotalTaxExempt  ,"
        '********************************************
        StrSQL = StrSQL & "   BanksData.BankName BankName,"
        StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
        StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
        StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
        StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
        '*************************************************
        StrSQL = StrSQL & " dbo.Transactions.without,dbo.Transactions.wait,   dbo.Transactions.FixesAssetsID, dbo.TblEquipments.Code, dbo.TblEquipments.name, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, "
        StrSQL = StrSQL & " dbo.TblProductionType.name AS productiontypename, dbo.TblProductionType.namee AS productiontypenamee, dbo.Transactions.ContactTime,"
        StrSQL = StrSQL & " dbo.Transaction_Details.ShowQty, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.TransactionComment, dbo.TblEmployee.Emp_Code,"
        StrSQL = StrSQL & " dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Cus_Phone,"
        StrSQL = StrSQL & " dbo.TblCustemers.Remark, dbo.Transactions.Transaction_Date, dbo.TblCountriesGovernmentsCities.CityName AS HeyName, dbo.TblCustemers.Account_Code,"
        StrSQL = StrSQL & " isnull(dbo.ACCOUNTS.DepitBalance,0) AS DepitBalance, isnull(dbo.ACCOUNTS.CreditBalance,0) AS CreditBalance, isnull(dbo.ACCOUNTS.opening_balance,0) AS opening_balance, dbo.TblEquipments.Notes, dbo.TblItems.Wight, dbo.TblItems.Content,"
        StrSQL = StrSQL & " dbo.TblItems.Dippre , dbo.TblItems.Source, dbo.TblItems.Typenew, dbo.TblItems.itemname,TblItems.ItemComment"
        StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
        StrSQL = StrSQL & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblCountriesGovernmentsCities ON dbo.Transactions.Neighborhoodid = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblProductionType ON dbo.Transactions.ProductionTypeid = dbo.TblProductionType.Id LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblEquipments ON dbo.Transactions.FixesAssetsID = dbo.TblEquipments.fixedAssetid "
        '***********************************
        StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
        StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
        StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
        StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
        '**************************************
        StrSQL = StrSQL & " WHERE     ( Transaction_Type=61 and dbo.Transactions.Transaction_Date =" & SQLDate(CDate(TDate), True) & ")"
        StrSQL = StrSQL & " ORDER BY dbo.Transactions.Wait, dbo.Transactions.Without"
    
    End If

    If reportid = 62 Then

        StrSQL = " SELECT    Transactions.TotalTaxExempt  ,"
        '********************************************
        StrSQL = StrSQL & "   BanksData.BankName BankName,"
        StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
        StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
        StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
        StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
        '*************************************************
        StrSQL = StrSQL & " dbo.GetNoOfShipments(dbo.Transactions.Transaction_ID) AS noofShipments,   dbo.Transactions.Transaction_ID, dbo.gettotalShippedQty1(dbo.Transactions.Transaction_ID) AS totalShippedQty,"
        StrSQL = StrSQL & "                       dbo.GetminTimeForShipments(dbo.Transactions.Transaction_ID) AS MinTime, dbo.GetmaxTimeForShipments(dbo.Transactions.Transaction_ID) AS MaxTime,"
        StrSQL = StrSQL & "                       dbo.Transactions.Without, dbo.Transactions.Wait, dbo.Transactions.FixesAssetsID, dbo.TblEquipments.Code, dbo.TblEquipments.name, dbo.TblCustemers.CusName,"
        StrSQL = StrSQL & "                       dbo.TblCustemers.CusNamee, dbo.TblProductionType.name AS productiontypename, dbo.TblProductionType.namee AS productiontypenamee,"
        StrSQL = StrSQL & "                       dbo.Transactions.ContactTime, dbo.Transaction_Details.ShowQty, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.TransactionComment,"
        StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.Cus_mobile,"
        StrSQL = StrSQL & "                       dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Remark, dbo.Transactions.Transaction_Date, dbo.TblCountriesGovernmentsCities.CityName AS HeyName,"
        StrSQL = StrSQL & "                       dbo.TblCustemers.Account_Code, ISNULL(dbo.ACCOUNTS.DepitBalance, 0) AS DepitBalance, ISNULL(dbo.ACCOUNTS.CreditBalance, 0) AS CreditBalance,"
        StrSQL = StrSQL & "                       ISNULL(dbo.ACCOUNTS.opening_balance, 0) AS opening_balance, dbo.TblEquipments.Notes, dbo.TblItems.Wight, dbo.TblItems.Content, dbo.TblItems.Dippre,"
        StrSQL = StrSQL & "                       dbo.TblItems.Source, dbo.TblItems.Typenew, dbo.TblItems.ItemName, TblEmployee_1.Emp_Name AS Workername, TblEmployee_2.Emp_Name AS Helpername,"
        StrSQL = StrSQL & "                       TblEmployee_1.Emp_Namee AS Workernamee, TblEmployee_2.Emp_Namee AS Helpernamee,TblItems.ItemComment "
        StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
        StrSQL = StrSQL & "                       dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
        StrSQL = StrSQL & "                       dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
        StrSQL = StrSQL & "                       dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblEmployee TblEmployee_2 ON dbo.Transactions.empID2 = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblEmployee TblEmployee_1 ON dbo.Transactions.empID1 = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblCountriesGovernmentsCities ON dbo.Transactions.Neighborhoodid = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblProductionType ON dbo.Transactions.ProductionTypeid = dbo.TblProductionType.Id LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblEquipments ON dbo.Transactions.FixesAssetsID = dbo.TblEquipments.fixedAssetid"
        '***********************************
        StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
        StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
        StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
        StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
        '**************************************
        StrSQL = StrSQL & " WHERE       Transaction_Type=61 and dbo.Transactions.Transaction_Date >=" & SQLDate(CDate(TDate), True)
        StrSQL = StrSQL & " and      dbo.Transactions.Transaction_Date <=" & SQLDate(CDate(tdate1), True)
        StrSQL = StrSQL & " ORDER BY dbo.Transactions.Wait, dbo.Transactions.Without"

    End If

    If reportid = 16 Or reportid = 160 Then
        StrSQL = "SELECT     Transactions.TotalTaxExempt  ,"
        '********************************************
        StrSQL = StrSQL & "   BanksData.BankName BankName,"
        StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
        StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
        StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
        StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
        '*************************************************
        StrSQL = StrSQL & " dbo.ItemsDetails.[Count], dbo.ItemsDetails.ItemId, dbo.ItemsDetails.Transaction_ID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.UnitID, "
        StrSQL = StrSQL & " dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ClassId, dbo.ItemsDetails.ParrtNoCode, dbo.TblItems.ItemCode, dbo.TblItemsSizes.SizeComment,"
        StrSQL = StrSQL & " dbo.TblItemsSizes.SizeName, dbo.TblItemsColors.ColorName, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Date, dbo.Transactions.NoteSerial,"
        StrSQL = StrSQL & " dbo.Transactions.NoteSerial1, dbo.Transactions.CusID, dbo.Transactions.StoreID, dbo.Transactions.UserID, dbo.Transactions.Emp_ID,"
        StrSQL = StrSQL & " dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress,"
        StrSQL = StrSQL & " dbo.Transactions.CashCustomerComment, dbo.Transactions.TransactionComment, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
        StrSQL = StrSQL & "  dbo.TblStore.STORENAME , dbo.TblStore.StoreNamee, dbo.TblItems.itemname, dbo.TblItems.ItemNamee,TblItems.ItemComment"
        StrSQL = StrSQL & " FROM         dbo.ItemsDetails INNER JOIN"
        StrSQL = StrSQL & " dbo.TblItems ON dbo.ItemsDetails.ItemId = dbo.TblItems.ItemID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId INNER JOIN"
        StrSQL = StrSQL & " dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID INNER JOIN"
        StrSQL = StrSQL & " dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID "
        '***********************************
        StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
        StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
        StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
        StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
        '**************************************
        StrSQL = StrSQL & "  WHERE     (dbo.ItemsDetails.Transaction_ID = " & PriceID & ")"
        'StrSQL = StrSQL & ""
    End If
    StrSQL = StrSQL & " Order By IsNull(Transaction_Details.LineID,Transaction_Details.id) "
    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
    Dim FileName As String

    '  If SystemOptions.UserInterface = EnglishInterface Then
    
    If Dir(App.path & "\Reports\Inventory\PerformaInvoices.rpt") = "" Then
        Msg = "«· ﬁ—Ì— €Ì— „ÊÃÊœ!!" & CHR(13)
        'Msg = "E??C? C?E??I ?? ???I ??C C???? ?? ??C? C?E??C??" & Chr(13)
        Msg = "\Reports\ShowPrice.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Screen.MousePointer = vbArrowHourglass
 
        If reportid = 0 Then
     
            FileName = App.path & "\Reports\Inventory\PerformaInvoices.rpt"
     
        ElseIf reportid = 1 Then
            FileName = App.path & "\Reports\Inventory\PerformaInvoices1.rpt"
        ElseIf reportid = 7 Then
            '    FileName = App.path & "\Reports\Inventory\PerformaInvoices7.rpt"
            FileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PerformaInvoices7.rpt"
        
        ElseIf reportid = 8 Then
              
            ' FileName = App.path & "\Reports\Inventory\PerformaInvoices8.rpt"
            FileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PerformaInvoices8.rpt"
            
        ElseIf reportid = 9 Then
            FileName = App.path & "\Reports\Inventory\PerformaInvoices9.rpt"
            
        ElseIf reportid = 91 Then
            FileName = App.path & "\Reports\Inventory\PerformaInvoices91.rpt"
            
        ElseIf reportid = 92 Then
            FileName = App.path & "\Reports\Inventory\PerformaInvoices92.rpt"
            
        ElseIf reportid = 93 Then
            FileName = App.path & "\Reports\Inventory\PerformaInvoices93.rpt"
        ElseIf reportid = 94 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                FileName = App.path & "\Reports\Inventory\PerformaInvoices94.rpt"
            Else
                FileName = App.path & "\Reports\Inventory\PerformaInvoices94E.rpt"
            End If
        ElseIf reportid = 942 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                FileName = App.path & "\Reports\Inventory\PerformaInvoices942.rpt"
            Else
                FileName = App.path & "\Reports\Inventory\PerformaInvoices942E.rpt"
            End If
        ElseIf reportid = 95 Then
            FileName = App.path & "\Reports\Inventory\PerformaInvoices95.rpt"
        ElseIf reportid = 96 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                FileName = App.path & "\Reports\Inventory\PerformaInvoices96.rpt"
            Else
                FileName = App.path & "\Reports\Inventory\PerformaInvoices96e.rpt"
            End If
            
        ElseIf reportid = 97 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                FileName = App.path & "\Reports\Inventory\PerformaInvoices97.rpt"
            Else
                FileName = App.path & "\Reports\Inventory\PerformaInvoices97e.rpt"
            End If
        ElseIf reportid = 10 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                FileName = App.path & "\Reports\REPORTS NEW\GuranteeTransaction.rpt"
            Else
                FileName = App.path & "\Reports\REPORTS NEW\GuranteeTransaction.rpt"
            End If
            
        ElseIf reportid = 11 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                FileName = App.path & "\Reports\REPORTS NEW\GuranteeTransaction1.rpt"
            Else
                FileName = App.path & "\Reports\REPORTS NEW\GuranteeTransaction1.rpt"
            End If
            
        ElseIf reportid = 12 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                FileName = App.path & "\Reports\REPORTS NEW\GuranteeTransaction2.rpt"
            Else
                FileName = App.path & "\Reports\REPORTS NEW\GuranteeTransaction2.rpt"
            End If
            
        ElseIf reportid = 2 Then
     
            FileName = App.path & "\Reports\Inventory\PerformaInvoices2.rpt"
        ElseIf reportid = 3 Then
            FileName = App.path & "\Reports\Inventory\PerformaInvoices3.rpt"
        ElseIf reportid = 4 Then
            FileName = App.path & "\Reports\Inventory\PerformaInvoices4.rpt"
        ElseIf reportid = 5 Then
     
        ElseIf reportid = 6 Then
            FileName = App.path & "\Reports\Inventory\PerformaInvoices6.rpt"
          
        ElseIf reportid = 100 Then
          
            If SystemOptions.UserInterface = ArabicInterface Then
                FileName = App.path & "\Reports\REPORTS NEW\PO.rpt"
            Else
                FileName = App.path & "\Reports\REPORTS NEW\POe.rpt"
            End If
        
        ElseIf reportid = 61 Then
            FileName = App.path & "\Reports\REPORTS NEW\ResproductionDaily.rpt"
        
        ElseIf reportid = 62 Then
            FileName = App.path & "\Reports\REPORTS NEW\ResproductionDaily1.rpt"
        
            'FileName = App.path & "\Reports\Inventory\ResproductionDaily.rpt"
        ElseIf reportid = 16 Then
            ' FileName = App.path & "\Reports\salesInvoice16.rpt"
            FileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice16.rpt"
        ElseIf reportid = 160 Then
            ' FileName = App.path & "\Reports\salesInvoice16.rpt"
            FileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice16m.rpt"
       
        ElseIf reportid = 620 Then
            '    FileName = App.path & "\Reports\Inventory\PerformaInvoices7.rpt"
            FileName = App.path & "\Special\" & SystemOptions.Reportpath & "\IssueRequest.rpt"
        
        End If

        Set xReport = xApp.OpenReport(FileName)
        xReport.Database.SetDataSource rs
        If reportid = 16 Or reportid = 94 Or reportid = 942 Or reportid = 8 Or reportid = 9 Or reportid = 620 Or reportid = 91 Or reportid = 92 Or reportid = 93 Or reportid = 95 Or reportid = 96 Or reportid = 97 Then
            CreateLogo xReport
        End If
          
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(5).AddCurrentValue EmpName
         
        Screen.MousePointer = vbDefault
    End If

    If reportid = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "«„— »Ì⁄"
        Else
            xReport.reporttitle = "Sales Order "
        End If
    
    ElseIf reportid = 1 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "«„— ‘—«¡ "
        Else
            xReport.reporttitle = "Purchase Order "
        End If

    ElseIf reportid = 7 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = " «„— ‘—«¡ "
        Else
            xReport.reporttitle = "Purchase Order "
        End If
        
    ElseIf reportid = 8 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "  ÿ·» œ«Œ·Ì  "
        Else
            xReport.reporttitle = "Internal Order "
        End If
        
    ElseIf reportid = 9 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "  ”‰‰œ ÕÃ“    "
        Else
            xReport.reporttitle = "Book Inventories "
        End If
        
    ElseIf reportid = 91 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = " ÿ·» ⁄—÷ ”⁄—    "
        Else
            xReport.reporttitle = "Request a quote "
        End If
        
    ElseIf reportid = 92 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "  ⁄—Ê÷ «”⁄«— «·„‘ —Ì«    "
        Else
            xReport.reporttitle = "quote "
        End If
        
    ElseIf reportid = 93 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "  ÿ·» ⁄—÷  "
        Else
            xReport.reporttitle = "Request a quote "
        End If
        
    ElseIf reportid = 94 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "   ⁄—÷   "
        Else
            xReport.reporttitle = "    quote "
        End If
        
    ElseIf reportid = 95 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "    «„— »Ì⁄ „»œ∆Ì    "
        Else
            xReport.reporttitle = "Primary Sales Order "
        End If
        
    ElseIf reportid = 96 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "     «„— »Ì⁄        "
        Else
            xReport.reporttitle = "Primary Sales Order "
        End If
        
    ElseIf reportid = 2 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "›« Ê—… „»œ∆Ì…"
        Else
            xReport.reporttitle = " Performa Invoice "
        End If

    ElseIf reportid = 3 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = " ÿ·» œ«Œ·Ì"
        Else
            xReport.reporttitle = "Internal Order "
        End If
      
    ElseIf reportid = 10 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "»Ì«‰«  «·÷„«‰"
        Else
            xReport.reporttitle = "Gurantee Data"
        End If
        
    ElseIf reportid = 11 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "«” ·«„ „’⁄œ"
        Else
            xReport.reporttitle = "Delivery elevator"
        End If
        
    ElseIf reportid = 6 Then '??E O?C? U?? C??C??
 
        Dim CusID         As Double
        Dim CusName       As String
        Dim ExpireDateH   As String
        Dim company       As String
        Dim JobTitle      As String
        Dim salary        As Double
        Dim JobAddress    As String
        Dim JobTel        As String
        Dim JobTelConvert As String
        Dim HomeTel       As String
        Dim Mobile1       As String
        Dim Mobile2       As String
        Dim Dola          As String
        Dim Madena        As String
        Dim hay           As String
        Dim Natinality    As String
        Dim Sex           As String
        Dim CustGID       As Double
        Dim Address       As String
        GetCustomerAllData CUSTID, CusName, ExpireDateH, company, JobTitle, salary, JobAddress, JobTel, JobTelConvert, HomeTel, Mobile1, Mobile2, Dola, Madena, hay, Natinality, Sex, CustGID, Address

        xReport.ParameterFields(9).AddCurrentValue Natinality
        xReport.ParameterFields(10).AddCurrentValue CStr(CustGID)
        xReport.ParameterFields(11).AddCurrentValue ExpireDateH
        xReport.ParameterFields(12).AddCurrentValue company
        xReport.ParameterFields(13).AddCurrentValue JobTitle
        xReport.ParameterFields(14).AddCurrentValue CStr(salary)
        xReport.ParameterFields(15).AddCurrentValue JobAddress
        xReport.ParameterFields(16).AddCurrentValue Madena & " " & hay & " " & Address
        xReport.ParameterFields(17).AddCurrentValue JobTel
        xReport.ParameterFields(18).AddCurrentValue JobTelConvert
        xReport.ParameterFields(19).AddCurrentValue HomeTel

        xReport.ParameterFields(20).AddCurrentValue Mobile1
        xReport.ParameterFields(21).AddCurrentValue Mobile2
        xReport.ParameterFields(22).AddCurrentValue Sex
        xReport.ParameterFields(23).AddCurrentValue ToHijriDate(rs("Transaction_Date").value)

    End If
    If reportid = 96 Then
        On Error Resume Next
        xReport.ParameterFields(9).AddCurrentValue StrDaeen
    End If
    If reportid = 100 Or reportid = 12 Then
        xReport.ParameterFields(9).AddCurrentValue stringItems
        xReport.ParameterFields(10).AddCurrentValue casingvalues
        xReport.ParameterFields(11).AddCurrentValue totals
        xReport.ParameterFields(12).AddCurrentValue WriteNo(CStr(totals), 0, True)
    End If
      
    If reportid = 94 Or reportid = 942 Then
        On Error Resume Next
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.ParameterFields(9).AddCurrentValue WriteNo(CStr(totals), 0, True)
        Else
            xReport.ParameterFields(9).AddCurrentValue WriteNo(CStr(totals), 0, True)
        End If
    End If
      
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
  
    FrmReport.TXTSTRSQL = StrSQL
    FrmReport.show
    FrmReport.txtPath = FileName
    ' FrmReport.TXTSTRSQL = StrSQL
    Screen.MousePointer = vbDefault
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "Œÿ√" & CHR(13) & "·« ÌÊÃœ »Ì«‰« " & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function





 
 

Public Function PrintTemplate(TempID As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim Msg As String
    Dim FrmReport As FrmReportViewer

    On Error GoTo ErrTrap

    If TempID = Null Then
        Exit Function
    End If

    StrSQL = "select * From QryTemplate where TemplateID=" & TempID
    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If Dir(App.path & "\Reports\Template.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\Template.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\Template.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Sub EmpsSales(FromDate As Variant, _
                     ToDate As Variant, _
                     Optional EmpID As Long = 0)

    Dim StrSQL      As String
    Dim StrWhere    As String
    Dim BolBegine   As Boolean
    Dim RsData      As ADODB.Recordset
    Dim xApp        As New CRAXDRT.Application
    Dim xReport     As New CRAXDRT.Report
    Dim xViewer     As ClsReportViewer
    Dim StrRS       As ADODB.Recordset
    Dim Msg         As String
    Dim StrFileName As String
    On Error GoTo ErrTrap

    StrFileName = App.path & "\Reports\EmpsSales.rpt"

    If Dir(App.path & "\Reports\EmpsSales.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\EmpsSales.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT Sum(QryTransactionsTotal.TransSum) AS SumOfTransSum," & "Count(QryTransactionsTotal.Transaction_ID) AS TransCount, TblEmployee.Emp_ID," & "TblEmployee.Emp_Name "
        StrSQL = StrSQL + " FROM TblEmployee INNER JOIN QryTransactionsTotal ON " & "TblEmployee.Emp_ID = QryTransactionsTotal.Emp_ID "
    Else
        StrSQL = "SELECT Sum(QryTransactionsTotal.TransNet) AS SumOfTransSum," & "Count(QryTransactionsTotal.Transaction_ID) AS TransCount, TblEmployee.Emp_ID," & "TblEmployee.Emp_Name "
        StrSQL = StrSQL + " FROM TblEmployee INNER JOIN dbo.QryTransactionsTotal()QryTransactionsTotal ON " & "TblEmployee.Emp_ID = QryTransactionsTotal.Emp_ID "
    End If

    If EmpID <> 0 Then
        StrWhere = "Where TblEmployee.Emp_ID=" & EmpID
        BolBegine = True
    End If

    If Not IsNull(FromDate) Then
        If BolBegine = True Then
            StrWhere = StrWhere + " AND QryTransactionsTotal.Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
        
        Else
            BolBegine = True
            StrWhere = StrWhere + " Where QryTransactionsTotal.Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
        End If
    End If

    If Not IsNull(ToDate) Then
        If BolBegine = True Then
            StrWhere = StrWhere + " AND QryTransactionsTotal.Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
        
        Else
            BolBegine = True
            StrWhere = StrWhere + " Where QryTransactionsTotal.Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
        End If
    End If

    StrSQL = StrSQL + StrWhere
    StrSQL = StrSQL + " GROUP BY TblEmployee.Emp_ID, TblEmployee.Emp_Name"
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
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = " ﬁ—Ì— »ÕÃ„ „»Ì⁄«  «·„ÊŸ›Ì‰"
    Set xViewer = New ClsReportViewer
    xViewer.FireReport xReport, WindowTarget, xReport.reporttitle, , , , StrFileName
    Sendkeys "{RIGHT}"
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Public Sub EmpsSalesDetails(FromDate As Variant, _
                            ToDate As Variant, _
                            Optional EmpID As Long = 0)

    Dim StrSQL      As String
    Dim StrWhere    As String
    Dim BolBegine   As Boolean
    Dim RsData      As ADODB.Recordset
    Dim xApp        As New CRAXDRT.Application
    Dim xReport     As New CRAXDRT.Report
    Dim xViewer     As ClsReportViewer
    Dim StrRS       As ADODB.Recordset
    Dim Msg         As String
    Dim StrFileName As String
    On Error GoTo ErrTrap

    StrFileName = App.path & "\Reports\EmpSalesDetails.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        Msg = "⁄›Ê« Â–« «· ﬁ—Ì— €Ì— „ «Õ „⁄ «·‰”Œ… «·⁄«œÌ… „‰ «·»—‰«„Ã"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
        StrSQL = "Select * From dbo.RptQryEmpSalesDetails"
    End If

    If EmpID <> 0 Then
        StrWhere = " Where Emp_ID=" & EmpID
        BolBegine = True
    End If

    If Not IsNull(FromDate) Then
        If BolBegine = True Then
            StrWhere = StrWhere + " AND Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
        Else
            BolBegine = True
            StrWhere = StrWhere + " Where Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
        End If
    End If

    If Not IsNull(ToDate) Then
        If BolBegine = True Then
            StrWhere = StrWhere + " AND Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
        
        Else
            BolBegine = True
            StrWhere = StrWhere + " Where Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
        End If
    End If

    StrSQL = StrSQL + StrWhere
    StrSQL = StrSQL + " Order By Transaction_Serial "
    Cn.CommandTimeout = 180
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
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = " ﬁ—Ì—  ›’Ì·Ï »„»Ì⁄«  „ÊŸ›"
    Set xViewer = New ClsReportViewer
    xViewer.FireReport xReport, WindowTarget, xReport.reporttitle, , , , StrFileName
    Sendkeys "{RIGHT}"
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

'************************
Public Sub EmpsAllocationsDetails(FromDate As Variant, _
                                  ToDate As Variant, _
                                  Optional EmpID As Long = 0)

    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim RsData As ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim xViewer As ClsReportViewer
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim StrFileName As String
    Dim reporttitle As String
    On Error GoTo ErrTrap

    StrFileName = App.path & "\Reports\EmpAllocationsData.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        Msg = "⁄›Ê« Â–« «· ﬁ—Ì— €Ì— „ «Õ „⁄ «·‰”Œ… «·⁄«œÌ… „‰ «·»—‰«„Ã"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
        reporttitle = " ﬁ—Ì— „Œ’’«  «·«Ã«“… Ê‰Â«Ì… «·Œœ„… "
        StrSQL = "SELECT     TOP 100 PERCENT dbo.TblEmployee.Emp_Name, dbo.TblEmpAllocations.AllocationsType, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Namee, "
        StrSQL = StrSQL & "  SUM(dbo.TblEmpAllocationsDetails.[value]) As totals"
        StrSQL = StrSQL & " FROM         dbo.TblEmpAllocations INNER JOIN"
        StrSQL = StrSQL & " dbo.TblEmpAllocationsDetails ON dbo.TblEmpAllocations.EmpAllocationsid = dbo.TblEmpAllocationsDetails.EmpAllocationsid INNER JOIN"
        StrSQL = StrSQL & " dbo.TblEmployee ON dbo.TblEmpAllocationsDetails.Emp_id = dbo.TblEmployee.Emp_ID"
        StrSQL = StrSQL & " WHERE  1=1"

        If EmpID <> 0 Then
            StrSQL = StrSQL + " and  dbo.TblEmpAllocationsDetails.Emp_id =" & EmpID
      
        End If

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND dbo.TblEmpAllocations.RecordDate >=" & SQLDate(CDate(FromDate), True) & ""
            reporttitle = reporttitle & CHR(13) & " „‰ " & FromDate
        End If

        If Not IsNull(ToDate) Then
            StrSQL = StrSQL + " AND dbo.TblEmpAllocations.RecordDate <=" & SQLDate(CDate(ToDate), True) & ""
            reporttitle = reporttitle & CHR(13) & " «·Ï " & ToDate
        End If
 
        StrSQL = StrSQL & " GROUP BY dbo.TblEmployee.Emp_Name, dbo.TblEmpAllocations.AllocationsType, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Namee"
        StrSQL = StrSQL & " ORDER BY  CAST(dbo.TblEmployee.Emp_Code AS int) "
    End If

    Cn.CommandTimeout = 180
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
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = reporttitle
    Set xViewer = New ClsReportViewer
    xViewer.FireReport xReport, WindowTarget, xReport.reporttitle, , , , StrFileName
    Sendkeys "{RIGHT}"
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Public Sub EmpInvProfit(FromDate As Variant, _
                        ToDate As Variant, _
                        Optional EmpID As Long = 0)
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim RsData As ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim FrmReport As FrmReportViewer
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim StrFileName As String

    On Error GoTo ErrTrap

    StrFileName = App.path & "\Reports\EmpSalesProfit.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT TblEmployee.Emp_ID, TblEmployee.Emp_Name, Transactions.Transaction_ID," & "Transactions.Transaction_Serial, Transactions.Transaction_Date," & "QryTransactionsTotal.TransSum, QryTransactionsTotal.TransProfit , TblEmployee.Emp_Comm"
        StrSQL = StrSQL + " FROM TblEmployee INNER JOIN (Transactions INNER JOIN " & "QryTransactionsTotal ON Transactions.Transaction_ID = " & "QryTransactionsTotal.Transaction_ID) ON TblEmployee.Emp_ID = Transactions.Emp_ID "
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "select QryRepEmpSalesProfit.* From QryRepEmpSalesProfit"
    End If

    If EmpID <> 0 Then
        StrWhere = " Where Emp_ID=" & EmpID
        BolBegine = True
    End If

    If Not IsNull(FromDate) Then
        If BolBegine = True Then
            StrWhere = StrWhere + " AND Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
        
        Else
            BolBegine = True
            StrWhere = StrWhere + " Where Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
        End If
    End If

    If Not IsNull(ToDate) Then
        If BolBegine = True Then
            StrWhere = StrWhere + " AND Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
        
        Else
            BolBegine = True
            StrWhere = StrWhere + " Where Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
        End If
    End If

    StrSQL = StrSQL + StrWhere
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
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name

    xReport.reporttitle = " ﬁ—Ì— »ﬁÌ„… ⁄„Ê·«  «·„ÊŸ›Ì‰"
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.show

    Sendkeys "{RIGHT}"
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Public Sub ShowSalesProfit(FromDate As Variant, _
                           ToDate As Variant, _
                           Optional LngItemID As Long = 0)

    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim RsData As ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim m_Report As CRAXDRT.Report

    Dim Msg As String
    Dim StrTitle As String
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim cOptions As ClsCompanyInfo
    On Error GoTo ErrTrap

    StrFileName = App.path & "\Reports\SalesProfit.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    StrSQL = "Select * From QrySalesProfit "

    If Not IsNull(FromDate) Then
        If BolBegine = True Then
            StrWhere = StrWhere + " AND Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
        Else
            BolBegine = True
            StrWhere = StrWhere + " Where Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
        End If
    End If

    If Not IsNull(ToDate) Then
        If BolBegine = True Then
            StrWhere = StrWhere + " AND Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
        
        Else
            BolBegine = True
            StrWhere = StrWhere + " Where Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
        End If
    End If

    StrSQL = StrSQL + StrWhere
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
    xReport.OpenSubreport("SubReport1").Database.SetDataSource RsData
    xReport.OpenSubreport("SubRptCustomerProfits.rpt").Database.SetDataSource RsData
    xReport.OpenSubreport("SubRptDayProfits.rpt").Database.SetDataSource RsData

    Set cOptions = New ClsCompanyInfo
    xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    StrTitle = " ﬁ—Ì— »ﬁÌ„… √—»«Õ «·›Ê« Ì—"

    If Not IsNull(FromDate) Then
        StrTitle = StrTitle & " „‰ " & Format(FromDate, "yyyy/M/d")
    End If

    If Not IsNull(ToDate) Then
        StrTitle = StrTitle & " ≈·Ï " & Format(ToDate, "yyyy/M/d")
    End If

    xReport.reporttitle = StrTitle
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , StrFileName
    Sendkeys "{RIGHT}"
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Public Sub ShowItemSalesProfit(FromDate As Variant, _
                               ToDate As Variant, _
                               Optional LngItemID As Long = 0)

    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim RsData As ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim FrmReport As FrmReportViewer
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim StrTitle As String
    Dim StrFileName As String

    On Error GoTo ErrTrap
    StrFileName = App.path & "\Reports\SalesProfitItems.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    StrSQL = "Select * From QryItemsProfit "
    BolBegine = True
    StrWhere = " Where ItemID=" & LngItemID & ""

    If Not IsNull(FromDate) Then
        If BolBegine = True Then
            StrWhere = StrWhere + " AND Transaction_Date >=#" & Format(FromDate, "mm/dd/yyyy") & "#"
        Else
            BolBegine = True
            StrWhere = StrWhere + " Where Transaction_Date >=#" & Format(FromDate, "mm/dd/yyyy") & "#"
        End If
    End If

    If Not IsNull(ToDate) Then
        If BolBegine = True Then
            StrWhere = StrWhere + " AND Transaction_Date <=#" & Format(ToDate, "mm/dd/yyyy") & "#"
        Else
            BolBegine = True
            StrWhere = StrWhere + " Where Transaction_Date <=#" & Format(ToDate, "mm/dd/yyyy") & "#"
        End If
    End If

    StrSQL = StrSQL + StrWhere

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
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    StrTitle = " ﬁ—Ì— »ﬁÌ„… √—»«Õ ’‰›"

    If Not IsNull(FromDate) Then
        StrTitle = StrTitle & " „‰ " & Format(FromDate, "yyyy/M/d")
    End If

    If Not IsNull(ToDate) Then
        StrTitle = StrTitle & " ≈·Ï " & Format(ToDate, "yyyy/M/d")
    End If

    xReport.reporttitle = StrTitle

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.show

    Sendkeys "{RIGHT}"
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Public Sub EmpSalesComm(FromDate As Variant, _
                        ToDate As Variant, _
                        Optional EmpID As Long = 0)
    Dim StrSQL As String
    Dim StrWhere As String
    Dim StrGroup As String
    Dim BolBegine As Boolean
    Dim RsData As ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim FrmReport As FrmReportViewer
    Dim StrRS As ADODB.Recordset
    Dim Msg As String
    Dim StrFileName As String
    Dim cReportViewer As ClsReportViewer
    On Error GoTo ErrTrap

    StrFileName = App.path & "\Reports\EmpSalesComm.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        Exit Sub
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT QryEmpSalesCom.* "
        StrSQL = StrSQL + " FROM dbo.QryEmpSalesCom(" & SQLDate(CDate(FromDate), True) & "," & SQLDate(CDate(ToDate), True) & "," & EmpID & ") QryEmpSalesCom"
    End If

    StrSQL = StrSQL + " Order BY PaymentType,Transaction_Serial"
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
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name

    xReport.reporttitle = " ﬁ—Ì— »ﬁÌ„… ⁄„Ê·«  «·„ÊŸ›Ì‰"
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, WindowTarget, , , , , StrFileName

    Sendkeys "{RIGHT}"
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
   
Public Function ShowSallingData(SallingID As Long, _
                                Optional IntReportStyle As Integer = 0, _
                                Optional LngCusID As Long, _
                                Optional payed As Double, _
                                Optional Result As Double, _
                                Optional pos As Integer, _
                                Optional sTitle As String, _
                                Optional printername As String = "", _
                                Optional CopnInfor As String, _
                                Optional total As String, _
                                Optional BranchVAT As String, _
                                Optional BranchID As Integer, _
                                Optional PayDes As String, _
                                Optional NoteSerial1 As String, _
                                Optional ByRef getSallingID As Long, _
                                Optional X As Integer)
   
    On Error Resume Next
    Dim WritTotal       As String
    Dim xApp            As New CRAXDRT.Application
    Dim xReport         As New CRAXDRT.Report
    Dim StrSQL          As String
    Dim StrRS           As New ADODB.Recordset
    Dim rs              As New ADODB.Recordset
    Dim RsTemp          As New ADODB.Recordset
    Dim ShowTax         As Boolean
    Dim Msg             As String
    Dim FrmReport       As FrmReportViewer
    Dim StrFileName     As String
    Dim SngCusBalBefore As Single
    Dim SngCusBalAfter  As Single
    Dim cOptions        As ClsCompanyInfo
    Dim CViewer         As ClsReportViewer
    Dim MySQL           As New ADODB.Recordset
    WritTotal = total
    'On Error GoTo ErrTrap

    If SallingID = 0 And NoteSerial1 = "" Then
        Exit Function
    End If

    If SallingID = 0 Then
        SallingID = getTransactionIdnoteserial1(NoteSerial1, 21)
        getSallingID = SallingID
    End If
    Dim report0 As String
    Dim report1 As String

    If IntReportStyle = 0 Then 'pos
    
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\SaleReport.rpt"
        GetSaleReportPOS report0, report1
    
        If report1 <> "" Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\" & report1 & ".rpt"
        End If
    ElseIf IntReportStyle = 123 Then 'pos
    
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\SaleReportPo10.rpt"
        GetSaleReportPOS report0, report1
    
        If report1 <> "" Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\" & report1 & ".rpt"
        End If
    ElseIf IntReportStyle = 124 Then 'pos
    
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\SaleReportOrder.rpt"
        GetSaleReportPOS report0, report1
    
        If report1 <> "" Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\" & report1 & ".rpt"
        End If
        '  StrFileName = App.path & "\Reports\SaleReport.rpt"
    ElseIf IntReportStyle = 1000 Then 'pos
    
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\SaleReport22.rpt"
        GetSaleReportPOS report0, report1
    
        If report1 <> "" Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\" & report1 & ".rpt"
        End If
    
    ElseIf IntReportStyle = 1 Then
        StrFileName = App.path & "\Reports\SaleReportStyle2.rpt"
    ElseIf IntReportStyle = 2 Then
        StrFileName = App.path & "\Reports\SaleReportStyle3.rpt"
    ElseIf IntReportStyle = 3 Then
        StrFileName = App.path & "\Reports\SaleReportStyle4.rpt"
    ElseIf IntReportStyle = 4 Then
        StrFileName = App.path & "\Reports\SaleReportByGuest.rpt"

    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    MySQL.Open "select * from InstallMent where Type = 18", Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (MySQL.EOF Or MySQL.BOF) Then
        xReport.ParameterFields(8).AddCurrentValue IIf(IsNull(MySQL("Total").value), "", Trim(MySQL("Total").value))
        ''Else
        ''     xReport.ParameterFields(8).AddCurrentValue "0"
    End If

    '  StrSQL = "Select * From QryBuyReport where Transaction_ID=" & SallingID
    'StrSQL = StrSQL & " ORDER BY ID "
    StrSQL = " SELECT   transactions.TotalTaxExempt,"
    '********************************************
    StrSQL = StrSQL & "   BanksData.BankName BankName,"
    StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
    StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
    StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
    StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
    '*************************************************
    StrSQL = StrSQL & "  transactions.QrCodeImage,ExtraValue,FirstEntryDateDate, transactions.TimeIn, LIPD,RIPD,LADD,RADD,LSH,RSH,LPRISM,RPRISM,      ciban,        TblItems.barcodeno  , Chasee,KM  ,CBoBasedON,  dbo.Transactions.STableID, dbo.Transactions.PayDes, dbo.Transactions.CopyNO, dbo.Transaction_Details.PrintName, "
    StrSQL = StrSQL & "                    dbo.Transactions.Transaction_Date, TblBranchesData.branch_name,TblBranchesData.branch_namee,TblBranchesData.manger,dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID, dbo.TblCustemers.CusName,"
    StrSQL = StrSQL & "                    dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID,"
    StrSQL = StrSQL & "                    dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.showPrice, dbo.Transaction_Details.Price, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ColorID,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ClassId, dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode,ISNULL(NULLIF(dbo.Transaction_Details.PrintName, ''), dbo.TblItems.ItemName) AS ItemName, dbo.TblItems.ItemPhoto,"
    StrSQL = StrSQL & "                    dbo.TblItems.ItemNamee, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName,"
    StrSQL = StrSQL & "                    dbo.Transactions.CashCustomerName, dbo.Transactions.NoteSerial1, dbo.Transaction_Details.TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ItemDiscount, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType, dbo.Transactions.ManualNo1,"
    StrSQL = StrSQL & "                    dbo.Transactions.PaymentType, dbo.Transaction_Details.ParrtNoCode, dbo.Transaction_Details.ItemDetailedCode, dbo.Transactions.LocationID,"
    StrSQL = StrSQL & "                    dbo.TblCountriesGovernmentsCities.CityName, dbo.TblCountriesGovernmentsCities.NameE, dbo.Transactions.CustomerlocationID,"
    StrSQL = StrSQL & "                    dbo.TblCustomersLocations.Name AS CustomerLocation, dbo.TblCustomersLocations.NameE AS CustomerLocationE, dbo.Transactions.TransactionComment,"
    StrSQL = StrSQL & "                    dbo.Transactions.Doctype, dbo.TblDoCumentsTypes.name AS DoctypeName, dbo.TblDoCumentsTypes.namee AS DoctypeNameE, dbo.Transaction_Details.LotNO,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ReminRequ, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, "
    StrSQL = StrSQL & "                    Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width, dbo.Transactions.RemainValue, dbo.Transactions.NetTransactionNo, dbo.Transactions.PayedValue, dbo.Transactions.NetValue,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.LAXIS, dbo.Transaction_Details.LCYL, dbo.Transaction_Details.LSPH, dbo.Transaction_Details.RAXIS, dbo.Transaction_Details.RCYL,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.RSPH, dbo.Transaction_Details.ExpiryDate, dbo.Transactions.VATNO, dbo.Transactions.Typ, dbo.Transactions.ResonVAT,"
    StrSQL = StrSQL & "                    dbo.Transactions.VAT, dbo.Transactions.BatchNo, Transactions.PurchOrderNo,dbo.Transactions.BookingDate, dbo.Transactions.Granty, dbo.Transactions.NoDelivery,"
    StrSQL = StrSQL & "                    dbo.Transactions.Transaction_NetValue, dbo.Transaction_Details.Vatyo, dbo.Transaction_Details.Vat AS VatDet, dbo.Transaction_Details.H2,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.H1, dbo.Transaction_Details.W, dbo.Transaction_Details.L, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile,"
    StrSQL = StrSQL & "                    dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail, dbo.TblCustemers.Address, dbo.TblCustemers.Company,"
    StrSQL = StrSQL & "                    dbo.TblCustemers.ExpireDateH, dbo.TblCustemers.CustGID, dbo.TblCustemers.JobTitle, dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTel,"
    StrSQL = StrSQL & "                    dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.HomeTel, dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2, dbo.TblCustemers.BoxMil,"
    StrSQL = StrSQL & "                    dbo.TblCustemers.ZipCode, dbo.TblCustemers.RecordDate, dbo.TblCustemers.Entry, dbo.TblCustemers.JobName, dbo.TblCustemers.Map,"
    StrSQL = StrSQL & "                    dbo.TblCustemers.authorizationname, dbo.TblCustemers.authorizationNo, dbo.TblCustemers.BankAccount, dbo.TblCustemers.BankName,"
    StrSQL = StrSQL & "                    dbo.TblCustemers.RecordNo, dbo.TblCustemers.RsIDDateH, dbo.TblCustemers.IBAN, dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankAddress,"
    StrSQL = StrSQL & "                    dbo.TblCustemers.CustGIDPlace, dbo.TblCustemers.Mail2, dbo.TblCustemers.VATNO AS CusVATNO, dbo.Transactions.CashCustomerPhone,"
    StrSQL = StrSQL & "                    dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress, dbo.Transactions.CashCustomerComment, dbo.Transaction_Details.QtyFaqtors,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty, dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.Remarks,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.Remarks1, dbo.Transaction_Details.Remarks2, dbo.Transaction_Details.guaranteeTime, dbo.Transaction_Details.CostPrice,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ItemProfit, dbo.Transaction_Details.QtyBySmalltUnit, dbo.Transaction_Details.sallReturnPrice, dbo.Transaction_Details.ToTAlELSHahn,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.rate, dbo.Transaction_Details.LineShahn, dbo.Transaction_Details.LineExpenses, dbo.Transaction_Details.ProductionDate,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.StoreID2, TblStore_1.StoreName AS StoreName2, TblStore_1.StoreNamee AS StoreNamee2, dbo.Transactions.CompetitionNo,"
    StrSQL = StrSQL & "                    dbo.Transactions.CompetitionName, dbo.Transactions.DateBaptizing, dbo.Transaction_Details.CountryOrg, dbo.Transaction_Details.ManufactCompany,"
    StrSQL = StrSQL & "                    dbo.Transactions.order_no, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.MaxUnitID,"
    StrSQL = StrSQL & "                    TblUnites_1.UnitName AS MaxUnitName, TblUnites_1.UnitNamee AS MaxUnitNameE, dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode,"
    StrSQL = StrSQL & "                    dbo.Groups.GroupNamee , dbo.Transaction_Details.TotalPrice,"
    StrSQL = StrSQL & "                       TBLCarTypes.Name CarName,Cast(TblCardAuthorizationReform.YearFact AS NVARCHAR(10)) YearFact  ,       IsNull(TBLCarTypes.Namee,TBLCarTypes.Name)                    CarNamee,"
    '  StrSQL = StrSQL & "                       PlateNo = (SELECT Top 1 IsNull(BoardNo,nBoardNo)  from TblCusCar Where TblCusCar.CustomerID =TblCardAuthorizationReform.CusID)"1

    '1-
    StrSQL = StrSQL & "    LIPD,RIPD,LADD,RADD,LSH,RSH,LPRISM,RPRISM,        ciban,             PlateNo = (SELECT Top 1 IsNull(BoardNo,nBoardNo)  from TblCusCar Where TblCusCar.CustomerID =TblCardAuthorizationReform.CusID and TblCusCar.ID = TblCardAuthorizationReform.CarID),"
    StrSQL = StrSQL & "                      CarModelName = (Select Top 1 Model From TblCarModels Where TblCarModels.CarID = TblCardAuthorizationReform.CarID and TblCarModels.ID  =TblCardAuthorizationReform.CarModelID) ,"

    StrSQL = StrSQL & "                  TblCustemers.Code as cusCode,TblCustemers.FullCode as CusFullCode  , Transaction_Details.EmpID4 ,Transaction_Details.CusID2 ,Transaction_Details.SupplierID,Transactions.QrCodeImage,"
    StrSQL = StrSQL & "                      EmpName4 = (Select Top 1 te.Emp_Name FROM TblEmployee te WHERE te.Emp_ID= Transaction_Details.EmpID4), "
    StrSQL = StrSQL & "                      CustName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.CusID2), "
    StrSQL = StrSQL & "                      SuppName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.SupplierID) "

    StrSQL = StrSQL & "                      ,transactions.PurchOrderNo Poa, SuppName = (Select top 1 cc.cusName from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29),"
    StrSQL = StrSQL & "                      SuppAddress= (Select top 1 cc.Address from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29),TblItems.ItemWithOutVAT"
    
    StrSQL = StrSQL & "                      ,CarTypesName2 = (Select Top 1 tc.Name FROM TBLCarTypes tc WHERE tc.ID= Transactions.CarTypeID) "
    StrSQL = StrSQL & "                      ,CarModelsName2 = (Select Top 1 tc.Model FROM TblCarModels tc WHERE tc.ID= Transactions.FixesAssetsID) "
    StrSQL = StrSQL & "                      ,CarColorName2 = (Select Top 1 tc.Name FROM TblColor tc WHERE tc.ID= Transactions.ColorID2) "
    StrSQL = StrSQL & "                      ,Transactions.CarCurrentValue"
    StrSQL = StrSQL & "                      ,Transactions.CarPrevValue"
    StrSQL = StrSQL & "                      ,Transactions.CarEnginoil"
    StrSQL = StrSQL & "                      ,Transactions.CarGearOil"
    StrSQL = StrSQL & "                      ,Transactions.CarOilChangeDate,Transactions.TimeIn"
    StrSQL = StrSQL & "                      ,Transactions.km"
    StrSQL = StrSQL & "                      ,Transactions.Shaseh"
    StrSQL = StrSQL & "                      ,Transactions.CarMeter"
    StrSQL = StrSQL & "                      ,Transactions.PlateNo PlateNo2"
    StrSQL = StrSQL & "                      ,Transactions.PlateNo CarOilChangeDate"
    StrSQL = StrSQL & "                      ,Transactions.CashCustomerName CashCustomerName"
    
    StrSQL = StrSQL & "                      ,Transactions.YearFact YearFact2"
    
    StrSQL = StrSQL & "                      ,Transaction_Details.Cash"
    StrSQL = StrSQL & "                      ,Transaction_Details.Mada"
    StrSQL = StrSQL & "                      ,Transaction_Details.Visa"
    StrSQL = StrSQL & "                      ,Transaction_Details.Deferred,Transaction_Details.DeferredQty"

 
    StrSQL = StrSQL & "                      ,TblCustemers.StreetName , TblCustemers.BuildingNumber, TblCustemers.IdentificationCode, TblCustemers.PlotIdentification, TblCustemers.AdditionalStreetName, TblCustemers.CountrySubentity, TblCustemers.ZipCode, TblCustemers.CityName as CusCityName, TblCustemers.CitySubdivisionName, TblCustemers.PostalZone, TblCustemers.CountrySubentity, TblCustemers.Id700"
    StrSQL = StrSQL & "                      ,CompanyStreetName = (Select StreetName from TblOptions),"
    StrSQL = StrSQL & "                      CompanyBuildingNumber = (Select BuildingNumber from TblOptions),"
    StrSQL = StrSQL & "                      CompanyIdentificationCode = (Select IdentificationCode from TblOptions),"
    StrSQL = StrSQL & "                      CompanyPlotIdentification = (Select PlotIdentification from TblOptions),"
    StrSQL = StrSQL & "                      CompanyAdditionalStreetName = (Select AdditionalStreetName from TblOptions),"
    StrSQL = StrSQL & "                      CompanyCountrySubentity = (Select CountrySubentity from TblOptions),"
    StrSQL = StrSQL & "                      CompanyCityName = (Select CityName from TblOptions),"
    StrSQL = StrSQL & "                      CompanyCitySubdivisionName = (Select CitySubdivisionName from TblOptions),"
    StrSQL = StrSQL & "                      CompanyPostalZone = (Select PostalZone from TblOptions),"
    
    StrSQL = StrSQL & "                      transactions.PONo , transactions.PaymentConditions, transactions.SumValueLine, transactions.SumVATLine,"
  StrSQL = StrSQL & "                      TblBranchesData.VATNO as BrancheVATNO,"
   
    StrSQL = StrSQL & "                      TblBranchesData.VATRegNo BrancheVATRegNo,TblBranchesData.Company_Comment as BrancheRegNo ,TblBranchesData.StreetName BrancheStreetName,TblBranchesData.CityName BrancheCityName,TblBranchesData.BuildingNumber BrancheBuildingNumber,TblBranchesData.CitySubdivisionName BrancheCitySubdivisionName"
    StrSQL = StrSQL & "                      ,ActivitesTypeName = (Select tblActivitesType.name from  tblActivitesType where id =TblBranchesData.ActivityTypeId ),"
        StrSQL = StrSQL & "                      ActivitesTypeNamee = (Select tblActivitesType.namee from  tblActivitesType where id =TblBranchesData.ActivityTypeId )"
    
  
    StrSQL = StrSQL & "      FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "                    dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "                    dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    StrSQL = StrSQL & "                    dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id INNER JOIN"
    StrSQL = StrSQL & "                    dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID Left outer JOIN"
    StrSQL = StrSQL & "                    dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID INNER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "                    dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID INNER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblUnites TblUnites_1 ON dbo.Transaction_Details.MaxUnitID = TblUnites_1.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblStore TblStore_1 ON dbo.Transaction_Details.StoreID2 = TblStore_1.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblDoCumentsTypes ON dbo.Transactions.Doctype = dbo.TblDoCumentsTypes.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblCustomersLocations ON dbo.Transactions.CustomerlocationID = dbo.TblCustomersLocations.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblCountriesGovernmentsCities ON dbo.Transactions.LocationID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId"
    StrSQL = StrSQL & "                               LEFT OUTER JOIN TblCardAuthorizationReform"
    StrSQL = StrSQL & "                               ON TblCardAuthorizationReform.WorkOrder = Transactions.order_no"
    StrSQL = StrSQL & "                               AND Transactions.CBoBasedON = 7"
    StrSQL = StrSQL & "                               LEFT OUTER JOIN TBLCarTypes"
    StrSQL = StrSQL & "                               ON TBLCarTypes.id = TblCardAuthorizationReform.CarTypeID"
    '***********************************
    StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
    StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
    StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
    StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
    '**************************************
  
    StrSQL = StrSQL & "  Where (dbo.Transactions.Transaction_ID = " & SallingID & ")"
    StrSQL = StrSQL & " ORDER BY dbo.Transaction_Details.ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Screen.MousePointer = vbArrowHourglass
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    Set cOptions = New ClsCompanyInfo
    xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    If IntReportStyle = 0 And CopnInfor <> "" Then
        xReport.ParameterFields(2).AddCurrentValue CopnInfor
    End If

    If (IntReportStyle = 0 Or IntReportStyle = 123 Or IntReportStyle = 124) And CopnInfor = "" Then
        ' xReport.ParameterFields(9).AddCurrentValue WriteNo(WritTotal, 0, True)  'WriteNo(WritTotal)
        xReport.ParameterFields(9).AddCurrentValue CStr(WriteNo(Format(WritTotal, "0.00"), 0, True, ".", , 0))
    End If
    If IntReportStyle = 0 Or IntReportStyle = 4 Then
        xReport.ParameterFields(6).AddCurrentValue payed
        xReport.ParameterFields(7).AddCurrentValue Result
        xReport.ParameterFields(8).AddCurrentValue sTitle

    End If

    'xReport.ParameterFields(6).AddCurrentValue "2"
    'xReport.ParameterFields(7).AddCurrentValue "1"

    xReport.reporttitle = "ÿ»«⁄… ›« Ê—… »Ì⁄"
    '⁄—÷ ÷—Ì»… «·„»Ì⁄« 
    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)

    If ShowTax = True Then
        ' ŸÂ—
        xReport.ParameterFields(5).AddCurrentValue "1"
    Else
        '·‰  ŸÂ—
        xReport.ParameterFields(5).AddCurrentValue "2"
    End If
    Dim StrSQL1 As String
    StrSQL1 = "select * From Notes where Transaction_ID=" & SallingID
    StrSQL1 = StrSQL1 + " and NoteType=0"
    RsTemp.Open StrSQL1, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
    Else
        xReport.ParameterFields(4).AddCurrentValue "0"
    End If

    If IntReportStyle = 1 Then
        SngCusBalBefore = GetCustomerAccount(LngCusID, True, , SallingID)
        xReport.ParameterFields(6).AddCurrentValue CDbl(SngCusBalBefore)
   
        SngCusBalAfter = GetCustomerAccount(LngCusID, True)
        xReport.ParameterFields(7).AddCurrentValue CDbl(SngCusBalAfter)
    End If
    If xReport.ParameterFields.count > 10 Then
        xReport.ParameterFields(10).AddCurrentValue BranchVAT
    End If
    If xReport.ParameterFields.count > 23 Then
        If SystemOptions.VATNoAccordActivity = False Then
            xReport.ParameterFields(24).AddCurrentValue cOptions.VATRegNo
        Else
            xReport.ParameterFields(24).AddCurrentValue GetRegVATNo(BranchID)
        End If
        xReport.ParameterFields(25).AddCurrentValue PayDes
    End If
    Screen.MousePointer = vbDefault
    Set CViewer = New ClsReportViewer
    Dim P_Target As PrintTarget
    DoEvents
    DoEvents
    DoEvents
    If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
        '  P_Target = PrinterTarget
        If X = vbYes Then
            P_Target = PrinterTarget
        Else
            P_Target = WindowTarget
        End If
   
    Else
        P_Target = WindowTarget
    End If

    If pos = 1 Then
        P_Target = PrinterTarget
    End If
    DoEvents
    DoEvents
    DoEvents
    CViewer.FireReport xReport, P_Target, , , , , StrFileName, printername, StrSQL
        DoEvents
    DoEvents
    DoEvents
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function
Public Function ShowSallingDataDetailed(SallingID As Long, _
                                        Optional IntReportStyle As Integer = 0, _
                                        Optional LngCusID As Long, _
                                        Optional payed As Double, _
                                        Optional Result As String, _
                                        Optional CusCode As String, _
                                        Optional BillComment As String, _
                                        Optional InstComm As Double, _
                                        Optional reportid As Integer, _
                                        Optional Nationality As String, _
                                        Optional GID As Double, _
                                        Optional Transaction_Date As String, _
                                        Optional AdvPayment As String, _
                                        Optional LblInstallCount As String, _
                                        Optional LblPrecenValue As String, _
                                        Optional LblFirstInstallDate As String, _
                                        Optional Madyna As String, _
                                        Optional hay As String, _
                                        Optional Address As String, _
                                        Optional total As String, _
                                        Optional currencycode As Integer, _
                                        Optional LawFirmValue As Double = 0, _
                                        Optional BranchVAT As String, _
                                        Optional BranchID As Integer, _
                                        Optional PayDes As String, Optional ByVal mBalance As String = "", Optional X As Integer, Optional ByVal mPic As PictureBox = Nothing, Optional ByVal mGrid As Object = Nothing, Optional ByVal mPaymentName As Variant = Nothing, Optional ByVal mValues As Variant = Nothing)

    On Error Resume Next
    Dim xApp            As New CRAXDRT.Application
    Dim xReport         As New CRAXDRT.Report
    Dim StrSQL          As String
    Dim StrRS           As New ADODB.Recordset
    Dim rs              As New ADODB.Recordset
    Dim RsTemp          As New ADODB.Recordset
    Dim ShowTax         As Boolean
    Dim Msg             As String
    Dim FrmReport       As FrmReportViewer
    Dim StrFileName     As String
    Dim SngCusBalBefore As Single
    Dim SngCusBalAfter  As Single
    Dim cOptions        As ClsCompanyInfo
    Dim CViewer         As ClsReportViewer
    Dim MySQL           As New ADODB.Recordset
    If mBalance = "" Then mBalance = 0
    ' On Error GoTo ErrTrap

    If SallingID = 0 Then
        Exit Function
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        If IntReportStyle = 0 Then
            
            '     StrFileName = App.path & "\Reports\salesInvoice.rpt"
                
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice.rpt"
            If SystemOptions.PrintInvoiceByBranch = True Then
                StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice" & Current_branch & ".rpt"
            End If
                
            Dim report0 As String
            Dim report1 As String
         
            GetSaleReportPOS report0, report1
            If report0 <> "" Then
                StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\" & report0 & ".rpt"
            End If
     
        ElseIf IntReportStyle = 50 Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoiceWithOutVat.rpt"
        ElseIf IntReportStyle = 1 Then
            'StrFileName = App.path & "\Reports\salesInvoice.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice.rpt"
        ElseIf IntReportStyle = 2 Then
            'StrFileName = App.path & "\Reports\salesInvoice.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice.rpt"
        ElseIf IntReportStyle = 3 Then
            'StrFileName = App.path & "\Reports\salesInvoice.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice.rpt"
        ElseIf IntReportStyle = 4 Then
            ' StrFileName = App.path & "\Reports\salesInvoice4.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice4.rpt"
        
        ElseIf IntReportStyle = 5 Then
            ' StrFileName = App.path & "\Reports\salesInvoice4.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice5c.rpt"
        ElseIf IntReportStyle = 18 Then
            
            '     StrFileName = App.path & "\Reports\salesInvoice.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoiceItemDef.rpt"
               
        End If

    Else

        If IntReportStyle = 0 Then
            '   StrFileName = App.path & "\Reports\salesInvoiceEng.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoiceE.rpt"
         
        ElseIf IntReportStyle = 1 Then
            'StrFileName = App.path & "\Reports\salesInvoiceEng.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoiceE.rpt"
        ElseIf IntReportStyle = 2 Then
            'StrFileName = App.path & "\Reports\salesInvoiceEng.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoiceE.rpt"
        ElseIf IntReportStyle = 3 Then
            'StrFileName = App.path & "\Reports\salesInvoiceEng.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoiceE.rpt"
        ElseIf IntReportStyle = 4 Then
            ' StrFileName = App.path & "\Reports\salesInvoice4.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice4E.rpt"
        ElseIf IntReportStyle = 5 Then
            ' StrFileName = App.path & "\Reports\salesInvoice4.rpt"
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice5E.rpt"
        ElseIf IntReportStyle = 50 Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoiceWithOutVatE.rpt"
        End If

    End If

    If reportid = 1 Then
        'StrFileName = App.path & "\Reports\salesInvoice5.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice5.rpt"
    End If

    If reportid = 2 Then
        '  StrFileName = App.path & "\Reports\salesInvoice6.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice6.rpt"
    End If

    If reportid = 17 Then
        '    StrFileName = App.path & "\Reports\salesInvoice7.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice17.rpt"
    End If

    If reportid = 7 Then
        '    StrFileName = App.path & "\Reports\salesInvoice7.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice7.rpt"
    End If
    
    If reportid = 13 Then
        'StrFileName = App.path & "\Reports\salesInvoice13.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice13.rpt"
    End If
    
    If reportid = 14 Then
        'StrFileName = App.path & "\Reports\salesInvoice14.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice14.rpt"
    End If
    
    If reportid = 16 Then
        '   StrFileName = App.path & "\Reports\salesInvoice16.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\salesInvoice16.rpt"
    End If
    
    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    StrSQL = "SELECT   Transactions.TotalTaxExempt  ,"
    '********************************************
    StrSQL = StrSQL & "   BanksData.BankName BankName,"
    StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
    StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
    StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
    StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
    '*************************************************
    StrSQL = StrSQL & " Transactions.TotalTaxExempt,transactions.TimeIn,transactions.QrCodeImage,ExtraValue,transactions.QrCodeDataPath, Transaction_Details.Remarks as RemarksDet,  transactions.QrCodeData,  LIPD,RIPD,LADD,RADD,LSH,RSH,LPRISM,RPRISM, transactions.chkTaxExempt as chkTaxExempt2,   Transaction_Details.ID as Ser,Transactions.DateRec,dbo.RetriveRecivedQty( Transaction_Details.Item_ID,  dbo.Transactions.NoteSerial1  )  as resqty2,dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID, "
    StrSQL = StrSQL & "  dbo.TblCustemers.CusName,Transactions.CardId0,Transactions.CardId1,Transactions.PayedValue2 , Transactions.StillValue, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.UserID, dbo.TblUsers.UserName,"
    StrSQL = StrSQL & "  transactions.RecConditions,transactions.DownloadPort,Transactions.PurchOrderNo,transactions.PortLoading,transactions.PaymentConditions,transactions.ShipName,"
    StrSQL = StrSQL & " dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty,"
    StrSQL = StrSQL & " dbo.Transaction_Details.showPrice,Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width, dbo.Transaction_Details.Price, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ColorID,"
    StrSQL = StrSQL & " dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
    StrSQL = StrSQL & " dbo.TblItems.ItemPhoto, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
    StrSQL = StrSQL & " dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.CashCustomerName, dbo.Transactions.NoteSerial1, dbo.Transaction_Details.TotalDiscountPerLine,"
    StrSQL = StrSQL & "      dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
    StrSQL = StrSQL & "    dbo.Transactions.ManualNo1, dbo.Transactions.PaymentType, dbo.TblUnites.UnitNamee, dbo.Transactions.order_no, dbo.TblBranchesData.branch_id,"
    StrSQL = StrSQL & "         dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE, TblBranchesData.manger,dbo.TblCustemers.fullcode , dbo.Transactions.TransactionComment,"
    StrSQL = StrSQL & "                       TBLCarTypes.Name CarName,TBLCarTypes.Namee CarNamee,CBoBasedON,Cast(TblCardAuthorizationReform.YearFact AS NVARCHAR(10)) YearFact ,"
    'StrSQL = StrSQL & "                       PlateNo = (SELECT Top 1 IsNull(BoardNo,nBoardNo)  from TblCusCar Where TblCusCar.CustomerID =TblCardAuthorizationReform.CusID)"2
    '2-
    StrSQL = StrSQL & "                       PlateNo = (SELECT Top 1 IsNull(BoardNo,nBoardNo)  from TblCusCar Where TblCusCar.CustomerID =TblCardAuthorizationReform.CusID and TblCusCar.ID = TblCardAuthorizationReform.CarID),"
    StrSQL = StrSQL & "                      CarModelName = (Select Top 1 Model From TblCarModels Where TblCarModels.CarID = TblCardAuthorizationReform.CarTypeID and TblCarModels.ID  =TblCardAuthorizationReform.CarModelID) ,"

    StrSQL = StrSQL & "                    , Transaction_Details.EmpID4 ,Transaction_Details.CusID2 ,Transaction_Details.SupplierID,Transactions.QrCodeImage,"
    StrSQL = StrSQL & "                      EmpName4 = (Select Top 1 te.Emp_Name FROM TblEmployee te WHERE te.Emp_ID= Transaction_Details.EmpID4), "
    StrSQL = StrSQL & "                      CustName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.CusID2), "
    StrSQL = StrSQL & "                      SuppName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.SupplierID) "
    StrSQL = StrSQL & "                      ,transactions.PurchOrderNo Poa, SuppName = (Select top 1 cc.cusName from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29),"
    StrSQL = StrSQL & "                      SuppAddress= (Select top 1 cc.Address from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29)"

    StrSQL = StrSQL & " , dbo.Transaction_Details.Remarks,BalnceCust = " & mBalance & ",Transaction_Details.AreaL,Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width,TblItems.ItemWithOutVAT, "
    StrSQL = StrSQL & "                      transactions.PONo,TblCustemers.ResponsibleContact,TblCustemers.Cus_mobile,TblCustemers.Cus_Phone,TblCustemers.VATNO,TblEmployee.Emp_mobile,Transaction_Details.id as MLineNo,IsNull(Transaction_Details.LineID,Transaction_Details.id) as LineID,"
    StrSQL = StrSQL & "                      transactions.SumValueLine , transactions.SumVATLine"
    
    
    StrSQL = StrSQL & "                      ,TblCustemers.StreetName , TblCustemers.BuildingNumber, TblCustemers.IdentificationCode, TblCustemers.PlotIdentification, TblCustemers.AdditionalStreetName, TblCustemers.CountrySubentity, TblCustemers.ZipCode, TblCustemers.CityName, TblCustemers.CitySubdivisionName, TblCustemers.PostalZone, TblCustemers.CountrySubentity, TblCustemers.Id700"
    StrSQL = StrSQL & "                      ,CompanyStreetName = (Select StreetName from TblOptions),"
    StrSQL = StrSQL & "                      CompanyBuildingNumber = (Select BuildingNumber from TblOptions),"
    StrSQL = StrSQL & "                      CompanyIdentificationCode = (Select IdentificationCode from TblOptions),"
    StrSQL = StrSQL & "                      CompanyPlotIdentification = (Select PlotIdentification from TblOptions),"
    StrSQL = StrSQL & "                      CompanyAdditionalStreetName = (Select AdditionalStreetName from TblOptions),"
    StrSQL = StrSQL & "                      CompanyCountrySubentity = (Select CountrySubentity from TblOptions),"
    StrSQL = StrSQL & "                      CompanyCityName = (Select CityName from TblOptions),"
    StrSQL = StrSQL & "                      CompanyCitySubdivisionName = (Select CitySubdivisionName from TblOptions),"
    StrSQL = StrSQL & "                      CompanyPostalZone = (Select PostalZone from TblOptions),"
    
    StrSQL = StrSQL & "                      TblBranchesData.VATNO as BrancheVATNO,"
   
    StrSQL = StrSQL & "                      TblBranchesData.VATRegNo BrancheVATRegNo,TblBranchesData.Company_Comment as BrancheRegNo ,TblBranchesData.StreetName BrancheStreetName,TblBranchesData.CityName BrancheCityName,TblBranchesData.BuildingNumber BrancheBuildingNumber,TblBranchesData.CitySubdivisionName BrancheCitySubdivisionName"
    StrSQL = StrSQL & "                      ,ActivitesTypeName = (Select tblActivitesType.name from  tblActivitesType where id =TblBranchesData.ActivityTypeId ),"
        StrSQL = StrSQL & "                      ActivitesTypeNamee = (Select tblActivitesType.namee from  tblActivitesType where id =TblBranchesData.ActivityTypeId )"
    StrSQL = StrSQL & "                      FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    StrSQL = StrSQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL & " dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID INNER JOIN"
    StrSQL = StrSQL & " dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID INNER JOIN"
    StrSQL = StrSQL & " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "       dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    StrSQL = StrSQL & "   dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "                               LEFT OUTER JOIN TblCardAuthorizationReform"
    StrSQL = StrSQL & "                               ON TblCardAuthorizationReform.WorkOrder = Transactions.order_no"
    StrSQL = StrSQL & "                               AND Transactions.CBoBasedON = 7"
    StrSQL = StrSQL & "                               LEFT OUTER JOIN TBLCarTypes"
    StrSQL = StrSQL & "                               ON TBLCarTypes.id = TblCardAuthorizationReform.CarTypeID "
    '***********************************
    StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
    StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
    StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
    StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
    '**************************************
    StrSQL = StrSQL & "  where Transactions.Transaction_ID=" & SallingID
    StrSQL = StrSQL & "order by Transaction_Details.id"

    '01062015
    ''''//////////////
    If reportid = 2 Or reportid = 17 Or reportid = 1 Or reportid = 18 Then

        StrSQL = " SELECT     TOP 100 PERCENT  "
        '********************************************
        StrSQL = StrSQL & "  Transactions.TotalTaxExempt  , BanksData.BankName BankName,"
        StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
        StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
        StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
        StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
        '*************************************************
        StrSQL = StrSQL & " Transactions.CardId0,transactions.TimeIn,Transactions.CardId1,transactions.QrCodeImage,ExtraValue,dbo.Transactions.Transaction_Date,Transactions.DateRec, Transaction_Details.Remarks as RemarksDet, Transaction_Details.ID as Ser, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID,"
        StrSQL = StrSQL & "                    transactions.QrCodeDataPath,transactions.QrCodeData,   dbo.TblCustemers.CusName,Transactions.PayedValue2 , Transactions.StillValue, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.UserID, dbo.TblUsers.UserName,"
        StrSQL = StrSQL & "                      dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.Transactions.CashCustomerName, dbo.Transactions.NoteSerial1,"
        StrSQL = StrSQL & "                      dbo.Transactions.Trans_Discount,Transactions.PurchOrderNo, dbo.Transactions.Trans_DiscountType, dbo.Transactions.ManualNo1, dbo.Transactions.PaymentType, dbo.Transactions.order_no,"
        StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name,TblBranchesData.manger, dbo.TblBranchesData.branch_namee, dbo.TblCustemers.Fullcode, dbo.Transactions.ManualNO,"
        StrSQL = StrSQL & "                      dbo.Transactions.LocationID, dbo.EmpGroupDep.GroupName, dbo.TblCustemers.CountryID2, Nationality_1.name, Nationality_1.namee, dbo.TblCustemers.CustGID,"
        StrSQL = StrSQL & "                      dbo.TblCustemers.CustGIDPlace, dbo.TblCustemers.ExpireDateH, dbo.TblCustemers.BoxMil, dbo.TblCustemers.ZipCode, dbo.TblCustemers.HomeTel,"
        StrSQL = StrSQL & "                      dbo.TblCustemers.Mobile1, dbo.TblCustemers.Mobile2, dbo.TblCustemers.Company, dbo.TblCustemers.JobName, dbo.TblCustemers.Entry,"
        StrSQL = StrSQL & "                      dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Cus_Phone,"
        StrSQL = StrSQL & "                      dbo.TblCustemers.GovernmentID, dbo.TblCountriesGovernments.GovernmentName, dbo.TblCustemers.CityID, dbo.TblCountriesGovernmentsCities.CityName,"
        StrSQL = StrSQL & "                      dbo.TblCustemers.Address, dbo.TblGridguranteeSales.Gname, dbo.TblGridguranteeSales.Gid, dbo.TblGridguranteeSales.PlaceGID,"
        StrSQL = StrSQL & "                      dbo.TblGridguranteeSales.DateIDH, dbo.TblGridguranteeSales.City, dbo.TblGridguranteeSales.WorkAddress, dbo.TblGridguranteeSales.POB,"
        StrSQL = StrSQL & "                      dbo.TblGridguranteeSales.ZibCode, dbo.TblGridguranteeSales.District, dbo.TblGridguranteeSales.Street, dbo.TblGridguranteeSales.job,"
        StrSQL = StrSQL & "                      dbo.TblGridguranteeSales.WorkTel, dbo.TblGridguranteeSales.Ext, dbo.TblGridguranteeSales.Salary, dbo.TblGridguranteeSales.Mobile,"
        StrSQL = StrSQL & "                      dbo.TblGridguranteeSales.HomeTel AS G_HomeTel, dbo.TblGridguranteeSales.HomeAddress, dbo.TblGridguranteeSales.Relation,"
        StrSQL = StrSQL & "                      dbo.TblGridguranteeSales.Remarks AS G_Remarks, dbo.TblGridguranteeSales.ID, dbo.TblGridguranteeSales.NatioalID, Nationality_1.name AS Gnamtionality,"
        StrSQL = StrSQL & "                      Nationality_1.namee AS GnamtionalityE, dbo.TblGridguranteeSales.GuarType, dbo.Gbasic.name AS GType, dbo.Gbasic.namee AS GTypeE, dbo.Transactions.YMD,"
        StrSQL = StrSQL & "                      dbo.Transactions.QestEndtDate, dbo.Transactions.QestStartDate, dbo.Transactions.QestEndtDateH, dbo.Transactions.QestStartDateH, dbo.Transactions.QstNo,"
        StrSQL = StrSQL & "                      dbo.Transactions.QstValue, dbo.Transactions.TotalQest, dbo.Transactions.Sandts, Nationality_2.name AS CusNamtiolity,"
        StrSQL = StrSQL & "                      Nationality_2.namee AS CusNamtiolityE , dbo.Transactions.LawFirmValue "
        StrSQL = StrSQL & " ,dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, "
        StrSQL = StrSQL & "                       TBLCarTypes.Name CarName,TBLCarTypes.Namee CarNamee,CBoBasedON,Cast(TblCardAuthorizationReform.YearFact AS NVARCHAR(10)) YearFact ,"
        'StrSQL = StrSQL & "                       PlateNo = (SELECT Top 1 IsNull(BoardNo,nBoardNo)  from TblCusCar Where TblCusCar.CustomerID =TblCardAuthorizationReform.CusID)"3
        '3-
        StrSQL = StrSQL & "                       PlateNo = (SELECT Top 1 IsNull(BoardNo,nBoardNo)  from TblCusCar Where TblCusCar.CustomerID =TblCardAuthorizationReform.CusID and TblCusCar.ID = TblCardAuthorizationReform.CarID),"
        StrSQL = StrSQL & "                      CarModelName = (Select Top 1 Model From TblCarModels Where TblCarModels.CarID = TblCardAuthorizationReform.CarTypeID and TblCarModels.ID  =TblCardAuthorizationReform.CarModelID) "
        StrSQL = StrSQL & "                      ,BalnceCust = " & mBalance & ",Transaction_Details.AreaL,Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width"
        StrSQL = StrSQL & "                      , transactions.chkTaxExempt as chkTaxExempt2,Transactions.TotalTaxExempt"

        StrSQL = StrSQL & "                    , Transactions.CardId0,Transactions.CardId1,Transaction_Details.EmpID4 ,Transaction_Details.CusID2 ,Transaction_Details.SupplierID,Transactions.QrCodeImage,"
        StrSQL = StrSQL & "                      EmpName4 = (Select Top 1 te.Emp_Name FROM TblEmployee te WHERE te.Emp_ID= Transaction_Details.EmpID4), "
        StrSQL = StrSQL & "                      CustName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.CusID2), "
        StrSQL = StrSQL & "                      SuppName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.SupplierID) "
        StrSQL = StrSQL & "                      ,transactions.PurchOrderNo Poa, SuppName = (Select top 1 cc.cusName from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29),"
        StrSQL = StrSQL & "                      SuppAddress= (Select top 1 cc.Address from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29),TblItems.ItemWithOutVAT ,"
        
        StrSQL = StrSQL & "                      transactions.PONo,TblCustemers.ResponsibleContact,TblCustemers.Cus_mobile,TblCustemers.Cus_Phone,TblCustemers.VATNO,TblEmployee.Emp_mobile,Transaction_Details.id as MLineNo,IsNull(Transaction_Details.LineID,Transaction_Details.id) as LineID"
        
              
         StrSQL = StrSQL & "                      ,TblCustemers.StreetName , TblCustemers.BuildingNumber, TblCustemers.IdentificationCode, TblCustemers.PlotIdentification, TblCustemers.AdditionalStreetName, TblCustemers.CountrySubentity, TblCustemers.ZipCode, TblCustemers.CityName, TblCustemers.CitySubdivisionName, TblCustemers.PostalZone, TblCustemers.CountrySubentity, TblCustemers.Id700"
         StrSQL = StrSQL & "                      ,CompanyStreetName = (Select StreetName from TblOptions),"
         StrSQL = StrSQL & "                      CompanyBuildingNumber = (Select BuildingNumber from TblOptions),"
         StrSQL = StrSQL & "                      CompanyIdentificationCode = (Select IdentificationCode from TblOptions),"
         StrSQL = StrSQL & "                      CompanyPlotIdentification = (Select PlotIdentification from TblOptions),"
         StrSQL = StrSQL & "                      CompanyAdditionalStreetName = (Select AdditionalStreetName from TblOptions),"
         StrSQL = StrSQL & "                      CompanyCountrySubentity = (Select CountrySubentity from TblOptions),"
         StrSQL = StrSQL & "                      CompanyCityName = (Select CityName from TblOptions),"
         StrSQL = StrSQL & "                      CompanyCitySubdivisionName = (Select CitySubdivisionName from TblOptions),"
         StrSQL = StrSQL & "                      CompanyPostalZone = (Select PostalZone from TblOptions),"
         StrSQL = StrSQL & "                      CompanyCountrySubentity = (Select CountrySubentity from TblOptions),"
         StrSQL = StrSQL & "                      TblBranchesData.VATNO as BrancheVATNO,"
   
         StrSQL = StrSQL & "                      TblBranchesData.VATRegNo BrancheVATRegNo,TblBranchesData.Company_Comment as BrancheRegNo ,TblBranchesData.StreetName BrancheStreetName,TblBranchesData.CityName BrancheCityName,TblBranchesData.BuildingNumber BrancheBuildingNumber,TblBranchesData.CitySubdivisionName BrancheCitySubdivisionName"
        
        StrSQL = StrSQL & "                      ,ActivitesTypeName = (Select tblActivitesType.name from  tblActivitesType where id =TblBranchesData.ActivityTypeId ),"
        StrSQL = StrSQL & "                      ActivitesTypeNamee = (Select tblActivitesType.namee from  tblActivitesType where id =TblBranchesData.ActivityTypeId )"
        StrSQL = StrSQL & " FROM         dbo.Gbasic RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.TblGridguranteeSales ON dbo.Gbasic.Id = dbo.TblGridguranteeSales.GuarType LEFT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.Nationality Nationality_1 ON dbo.TblGridguranteeSales.NatioalID = Nationality_1.id RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.Transactions INNER JOIN"
        StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
        StrSQL = StrSQL & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
        StrSQL = StrSQL & "                      dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID INNER JOIN"
        StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID ON"
        StrSQL = StrSQL & "                      dbo.TblGridguranteeSales.Transec_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.TblCountriesGovernmentsCities ON dbo.TblCustemers.CityID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.TblCountriesGovernments ON dbo.TblCustemers.GovernmentID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.Nationality Nationality_2 ON dbo.TblCustemers.CountryID2 = Nationality_2.id LEFT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.EmpGroupDep ON dbo.Transactions.LocationID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
        StrSQL = StrSQL & "                               LEFT OUTER JOIN TblCardAuthorizationReform"
        StrSQL = StrSQL & "                               ON TblCardAuthorizationReform.WorkOrder = Transactions.order_no"
        StrSQL = StrSQL & "                               AND Transactions.CBoBasedON = 7"
        StrSQL = StrSQL & "                               LEFT OUTER JOIN TBLCarTypes"
        StrSQL = StrSQL & "                               ON TBLCarTypes.id = TblCardAuthorizationReform.CarTypeID"
        '***********************************
        StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
        StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
        StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
        StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
        '**************************************
        StrSQL = StrSQL & "  where Transactions.Transaction_ID=" & SallingID
        ''''/////
    Else

        StrSQL = " SELECT   Transactions.TotalTaxExempt  ,"
        '********************************************
        StrSQL = StrSQL & "   BanksData.BankName BankName,"
        StrSQL = StrSQL & "   BanksData.account_no BankAccountNo,"
        StrSQL = StrSQL & "   BanksData.IBan BankIBan,"
        StrSQL = StrSQL & "   Bankcurrency.name BankCurrencyName,"
        StrSQL = StrSQL & "   Transactions.PurchOrderNo ,"
        '*************************************************
        StrSQL = StrSQL & "  Transactions.TotalTaxExempt ,transactions.TimeIn ,Transactions.CardId0,Transactions.CardId1,FirstEntryDateDate, transactions.QrCodeImage,ExtraValue,transactions.QrCodeDataPath,transactions.QrCodeData, Transaction_Details.Remarks as RemarksDet, LIPD,RIPD,LADD,RADD,LSH,RSH,LPRISM,RPRISM,  CIBAN,TblItems.barcodeno  ,     Chasee,KM  ,   Transactions.DateRec,    doorno,dbo.Transaction_Details.ProductionDate, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.LotNO, dbo.RetriveRecivedQty(dbo.Transaction_Details.Item_ID, "
  
        StrSQL = StrSQL & "                   dbo.Transactions.NoteSerial1) AS resqty2, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID, dbo.TblCustemers.CusName,"
        StrSQL = StrSQL & "  transactions.RecConditions,Transaction_Details.ID as Ser,transactions.DownloadPort,transactions.PortLoading,transactions.PaymentConditions,transactions.ShipName,transactions.QrCodeImage,transactions.QrCodeDataPath,transactions.QrCodeData, "
        StrSQL = StrSQL & "                       dbo.TblCustemers.CusNamee ,Transactions.PayedValue2 , Transactions.StillValue, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name,"
        StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Namee, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.Price, dbo.Transaction_Details.Quantity,"
        StrSQL = StrSQL & "                       dbo.Transaction_Details.ColorID, Transactions.PurchOrderNo,dbo.Transaction_Details.ClassId, dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemPhoto, dbo.TblItems.ItemNamee,"
        StrSQL = StrSQL & "                       dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.CashCustomerName, dbo.Transactions.NoteSerial1,"
        StrSQL = StrSQL & "                       dbo.Transaction_Details.TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
        StrSQL = StrSQL & "                       dbo.Transactions.ManualNo1, dbo.Transactions.PaymentType, dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name,TblBranchesData.manger, dbo.TblBranchesData.branch_namee, dbo.TblCustemers.Fullcode,"
        StrSQL = StrSQL & "                       dbo.Transaction_Details.Remarks, dbo.Transactions.ManualNO, dbo.Transactions.LocationID, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.CustGIDPlace,"
        StrSQL = StrSQL & "                       dbo.TblCustemers.CustGID, dbo.TblCustemers.HomeTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.JobTel, dbo.TblCustemers.JobAddress, dbo.TblCustemers.Salary, dbo.TblCustemers.JobTitle,"
        StrSQL = StrSQL & "                       dbo.TblCustemers.Company, dbo.TblCustemers.ExpireDateH, dbo.TblCountriesGovernmentsCities.CityName, dbo.Transactions.CustomerlocationID, dbo.TblCustomersLocations.Name AS CustomersLocation,"
        StrSQL = StrSQL & "                       dbo.TblCustomersLocations.NameE AS CustomersLocationE, dbo.Transaction_Details.LAXIS, dbo.Transaction_Details.LCYL, dbo.Transaction_Details.LSPH, dbo.Transaction_Details.RAXIS,"
        StrSQL = StrSQL & "                       dbo.Transaction_Details.RCYL, dbo.Transaction_Details.RSPH, dbo.TblCountriesGovernmentsCities.NameE, dbo.Transactions.TransactionComment, dbo.Transactions.Doctype,"
        StrSQL = StrSQL & "                       dbo.TblDoCumentsTypes.name AS DoctypeName, dbo.TblDoCumentsTypes.namee AS DoctypeNameE, Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width, dbo.Transaction_Details.Area,"
        StrSQL = StrSQL & "                       dbo.Transaction_Details.NoCount, dbo.Transaction_Details.ReminRequ, dbo.Transactions.VATNO, dbo.Transactions.Typ, dbo.Transactions.ResonVAT, dbo.Transactions.VAT, dbo.Transactions.BatchNo,"
        StrSQL = StrSQL & "                       dbo.Transactions.BookingDate, dbo.Transactions.Granty, dbo.Transactions.NoDelivery, dbo.Transactions.GrantyPeriod, dbo.Transactions.Transaction_NetValue, dbo.Transactions.BasedOn,"
        StrSQL = StrSQL & "                       dbo.Transactions.DueDate, dbo.Transactions.CopyNO, dbo.Transaction_Details.Vatyo, dbo.Transaction_Details.H2, dbo.Transaction_Details.H1, dbo.Transaction_Details.Vat AS VatDet,"
        StrSQL = StrSQL & "                       dbo.Transaction_Details.W, dbo.Transaction_Details.L, dbo.Transactions.AddValue, dbo.Transactions.AddToTotal, dbo.Transactions.RecTime, dbo.Transaction_Details.PrintName,"
        StrSQL = StrSQL & "                       dbo.TblCustemers.VATNO AS CusVATNO, dbo.TblCustemers.BankCode, dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankAddress, dbo.TblCustemers.Mail2, dbo.TblCustemers.IBAN,"
        StrSQL = StrSQL & "                       dbo.TblCustemers.RecordNo, dbo.TblCustemers.RecordDateH, dbo.TblCustemers.BankAccount, dbo.TblCustemers.JobName, dbo.TblCustemers.authorizationname, dbo.TblCustemers.authorizationNo,"
        StrSQL = StrSQL & "                       dbo.TblCustemers.Entry, dbo.TblCustemers.RecordDate, dbo.TblCustemers.ZipCode, dbo.TblCustemers.BoxMil, dbo.TblCustemers.Mobile2, dbo.TblCustemers.Mobile1, dbo.TblCustemers.E_mail,"
        StrSQL = StrSQL & "                       dbo.TblCustemers.FaxNumber, dbo.TblCustemers.Address, dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress,"
        StrSQL = StrSQL & "                       dbo.Transactions.CashCustomerComment, dbo.Transaction_Details.QtyFaqtors, dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty,"
        StrSQL = StrSQL & "                       dbo.Transaction_Details.StoreID2, TblStore_1.StoreName AS StoreName2, TblStore_1.StoreNamee AS StoreNamee2, dbo.Transactions.RemainValue, dbo.Transactions.NetValue, dbo.Transactions.PayedValue,"
        StrSQL = StrSQL & "                       dbo.Transaction_Details.Remarks1, dbo.Transaction_Details.Remarks2, dbo.Transactions.PODays, dbo.Transactions.PayDes, dbo.Transactions.STableID, dbo.Transaction_Details.UnitId,"
        StrSQL = StrSQL & "                       dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.MaxUnitID, TblUnites_1.UnitName AS MaxUnitName, TblUnites_1.UnitNamee AS MaxUnitNameE,"
        StrSQL = StrSQL & "                       dbo.Transaction_Details.CountryOrg, dbo.Transaction_Details.ManufactCompany, dbo.Transactions.CompetitionNo, dbo.Transactions.CompetitionName, dbo.Transactions.DateBaptizing,"
        StrSQL = StrSQL & "                       dbo.Transactions.order_no1, dbo.Transactions.order_no, dbo.Transaction_Details.EmpID4, TblEmployee_1.Emp_Name AS Emp_NameLine, TblEmployee_1.Fullcode AS EmpLineFullcode,"
        StrSQL = StrSQL & "                       TblEmployee_1.Emp_Namee AS Emp_NameLineE ,dbo.Transaction_Details.TotalPrice,"
 
        StrSQL = StrSQL & "                       TBLCarTypes.Name CarName,IsNull(TBLCarTypes.Namee,TBLCarTypes.Name) CarNamee,CBoBasedON,"
        StrSQL = StrSQL & "                       Cast(TblCardAuthorizationReform.YearFact AS NVARCHAR(10)) as  YearFact ,"
  
        '  StrSQL = StrSQL & "                       0 as  YearFact2 ,"
        '4-
  
        StrSQL = StrSQL & "                        PlateNo = ("
        StrSQL = StrSQL & "                                  SELECT  TOP 1 ISNULL(BoardNo, nBoardNo)"
        StrSQL = StrSQL & "                                  From TblCusCar"
        StrSQL = StrSQL & "                                  Where TblCusCar.customerid = TblCardAuthorizationReform.CusID"
        StrSQL = StrSQL & "                                  and TblCusCar.ID = TblCardAuthorizationReform.CarID       ),"
        StrSQL = StrSQL & "                      CarModelName = (Select Top 1 Model From TblCarModels Where TblCarModels.CarID = TblCardAuthorizationReform.CarTypeID and TblCarModels.ID  =TblCardAuthorizationReform.CarModelID) ,"
        StrSQL = StrSQL & "                       TblCustemers.E_mail , TblCustemers.ResponsibleContact "
        StrSQL = StrSQL & "                      ,BalnceCust = " & mBalance & ",Transaction_Details.AreaL,Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width,"
        StrSQL = StrSQL & "                      Groups.GroupName,Groups.chkTaxExempt ,Transactions.TotalTaxExempt  "
        StrSQL = StrSQL & "                      , transactions.chkTaxExempt as chkTaxExempt2"
 
        StrSQL = StrSQL & "                    , Transactions.CardId0,Transactions.CardId1,Transaction_Details.EmpID4 ,Transaction_Details.CusID2 ,Transaction_Details.SupplierID,Transactions.QrCodeImage,"
        StrSQL = StrSQL & "                      EmpName4 = (Select Top 1 te.Emp_Name FROM TblEmployee te WHERE te.Emp_ID= Transaction_Details.EmpID4), "
        StrSQL = StrSQL & "                      CustName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.CusID2), "
        StrSQL = StrSQL & "                      SuppName4 = (Select Top 1 tc.CusName FROM TblCustemers tc WHERE tc.CusID= Transaction_Details.SupplierID) "
        StrSQL = StrSQL & "                      ,transactions.PurchOrderNo Poa, SuppName = (Select top 1 cc.cusName from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29),"
        StrSQL = StrSQL & "                      SuppAddress= (Select top 1 cc.Address from Transactions TT inner join TblCustemers cc On tt.CusId = cc.cusId  where TT.NoteSerial1 =transactions.PurchOrderNo  and tt.Transaction_Type = 29),TblItems.ItemWithOutVAT,"
        StrSQL = StrSQL & "                      transactions.PONo,TblCustemers.ResponsibleContact,TblCustemers.Cus_mobile,TblCustemers.Cus_Phone,TblCustemers.VATNO,TblEmployee.Emp_mobile,Transaction_Details.id as MLineNo,IsNull(Transaction_Details.LineID,Transaction_Details.id) as LineID,"
        StrSQL = StrSQL & "                      transactions.SumValueLine , transactions.SumVATLine"
        
    
       StrSQL = StrSQL & "                      ,TblCustemers.StreetName , TblCustemers.BuildingNumber, TblCustemers.IdentificationCode, TblCustemers.PlotIdentification, TblCustemers.AdditionalStreetName, TblCustemers.CountrySubentity, TblCustemers.ZipCode, TblCustemers.CityName, TblCustemers.CitySubdivisionName, TblCustemers.PostalZone, TblCustemers.CountrySubentity, TblCustemers.Id700"
       StrSQL = StrSQL & "                      ,CompanyStreetName = (Select StreetName from TblOptions),"
       StrSQL = StrSQL & "                      CompanyBuildingNumber = (Select BuildingNumber from TblOptions),"
       StrSQL = StrSQL & "                      CompanyIdentificationCode = (Select IdentificationCode from TblOptions),"
       StrSQL = StrSQL & "                      CompanyPlotIdentification = (Select PlotIdentification from TblOptions),"
       StrSQL = StrSQL & "                      CompanyAdditionalStreetName = (Select AdditionalStreetName from TblOptions),"
       StrSQL = StrSQL & "                      CompanyCountrySubentity = (Select CountrySubentity from TblOptions),"
       StrSQL = StrSQL & "                      CompanyCityName = (Select CityName from TblOptions),"
       StrSQL = StrSQL & "                      CompanyCitySubdivisionName = (Select CitySubdivisionName from TblOptions),"
       StrSQL = StrSQL & "                      CompanyPostalZone = (Select PostalZone from TblOptions),"
       StrSQL = StrSQL & "                      CompanyCountrySubentity = (Select CountrySubentity from TblOptions),"
       StrSQL = StrSQL & "                      TblBranchesData.VATNO as BrancheVATNO,"
   
       StrSQL = StrSQL & "                      TblBranchesData.VATRegNo BrancheVATRegNo,TblBranchesData.Company_Comment as BrancheRegNo ,TblBranchesData.StreetName BrancheStreetName,TblBranchesData.CityName BrancheCityName,TblBranchesData.BuildingNumber BrancheBuildingNumber,TblBranchesData.CitySubdivisionName BrancheCitySubdivisionName"
         StrSQL = StrSQL & "                      ,ActivitesTypeName = (Select tblActivitesType.name from  tblActivitesType where id =TblBranchesData.ActivityTypeId ),"
        StrSQL = StrSQL & "                      ActivitesTypeNamee = (Select tblActivitesType.namee from  tblActivitesType where id =TblBranchesData.ActivityTypeId )"
        StrSQL = StrSQL & "      FROM            dbo.Transactions INNER JOIN"
        StrSQL = StrSQL & "                       dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
        StrSQL = StrSQL & "                       dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
        StrSQL = StrSQL & "                       dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
        StrSQL = StrSQL & "                       dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID INNER JOIN"
        StrSQL = StrSQL & "                       dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID "
        StrSQL = StrSQL & "                       Inner Join"
        StrSQL = StrSQL & "                         Groups ON TblItems.GroupID = Groups.GroupID LEFT OUTER JOIN"
  
        StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblEmployee AS TblEmployee_1 ON dbo.Transaction_Details.EmpID4 = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblUnites AS TblUnites_1 ON dbo.Transaction_Details.MaxUnitID = TblUnites_1.UnitID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblStore AS TblStore_1 ON dbo.Transaction_Details.StoreID2 = TblStore_1.StoreID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblDoCumentsTypes ON dbo.Transactions.Doctype = dbo.TblDoCumentsTypes.id LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblCustomersLocations ON dbo.Transactions.CustomerlocationID = dbo.TblCustomersLocations.ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblCountriesGovernmentsCities ON dbo.Transactions.LocationID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
        StrSQL = StrSQL & "                       dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
        StrSQL = StrSQL & "                               LEFT OUTER JOIN TblCardAuthorizationReform"
        StrSQL = StrSQL & "                               ON TblCardAuthorizationReform.WorkOrder = Transactions.order_no"
        StrSQL = StrSQL & "                               AND Transactions.CBoBasedON = 7"
        StrSQL = StrSQL & "                               LEFT OUTER JOIN TBLCarTypes"
        StrSQL = StrSQL & "                               ON TBLCarTypes.id = TblCardAuthorizationReform.CarTypeID "
        
        '***********************************
        StrSQL = StrSQL & "     LEFT OUTER JOIN dbo.BanksData "
        StrSQL = StrSQL & "     ON dbo.Transactions.BankID = dbo.BanksData.BankID "
        StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.currency Bankcurrency "
        StrSQL = StrSQL & "    ON dbo.BanksData.Currency_ID = Bankcurrency.id "
        '**************************************
        StrSQL = StrSQL & "  where Transactions.Transaction_ID=" & SallingID
        '  StrSQL = StrSQL + "order by dbo.Transaction_Details.id"
        '01 06 2015
        '''/// 02 03 2016
    End If
    'StrSQL = StrSQL + "order by id"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Screen.MousePointer = vbArrowHourglass
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    Set cOptions = New ClsCompanyInfo
    xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(5).AddCurrentValue BillComment

    Dim i As Long
    For i = 1 To xReport.FormulaFields.count
        Select Case xReport.FormulaFields.Item(i).Name
            Case "{@PaymentName1}"
                xReport.FormulaFields.Item(i).text = "'" & mPaymentName(1) & "'"
            Case "{@PaymentName2}"
                xReport.FormulaFields.Item(i).text = "'" & mPaymentName(2) & "'"
            Case "{@PaymentName3}"
                xReport.FormulaFields.Item(i).text = "'" & mPaymentName(3) & "'"
        
            Case "{@PaymentName4}"
                xReport.FormulaFields.Item(i).text = "'" & mPaymentName(4) & "'"
            Case "{@PaymentName5}"
                xReport.FormulaFields.Item(i).text = "'" & mPaymentName(5) & "'"
        
            Case "{@Value1}"
                xReport.FormulaFields.Item(i).text = val(mValues(1))
            Case "{@Value2}"
                xReport.FormulaFields.Item(i).text = val(mValues(2))
            Case "{@Value3}"
                xReport.FormulaFields.Item(i).text = val(mValues(3))
            Case "{@Value4}"
                xReport.FormulaFields.Item(i).text = val(mValues(4))
            Case "{@Value5}"
                xReport.FormulaFields.Item(i).text = val(mValues(5))
            
            Case "{@TitleBy}"
                '      xReport.FormulaFields.Item(i).Text = "'" & mTitleBy & "'"
            
        End Select
    Next i

    If val(InstComm) >= 0 Then
        If xReport.ParameterFields.count >= 10 Then
            xReport.ParameterFields(10).AddCurrentValue val(InstComm)
        End If
    End If

    xReport.ParameterFields(6).AddCurrentValue CStr(SystemOptions.Count_ACCOUNT_digit)
 
    'xxxx
    If reportid = 17 Then
        xReport.ParameterFields(8).AddCurrentValue WriteNo(Round(LawFirmValue, SystemOptions.Count_ACCOUNT_digit), 0, True, ".", , , , currencycode)
    Else
        ' xReport.ParameterFields(8).AddCurrentValue WriteNo(Round(result, SystemOptions.Count_ACCOUNT_digit), 0, True, ".", , , , currencycode)                 ' WriteNo(CStr(result), 0, True, , , , , currencycode)
        ' xReport.ParameterFields(8).AddCurrentValue WriteNo(Round(result, SystemOptions.Count_ACCOUNT_digit), 0, True, ".", , , , currencycode)
        xReport.ParameterFields(8).AddCurrentValue CStr(WriteNo(Format(Result, "0.00"), 0, True, ".", , 0, , currencycode))
        If xReport.ParameterFields.count >= 30 Then
             xReport.ParameterFields(30).AddCurrentValue CStr(WriteNo(Format(Result, "0.00"), 0, True, ".", , 1, , currencycode))
        
             xReport.ParameterFields(31).AddCurrentValue GetCurrencyCode(CDbl(currencycode), "Code")
        End If
    End If
    xReport.ParameterFields(9).AddCurrentValue CusCode
    
    If SystemOptions.VATNoAccordActivity = False Then
        xReport.ParameterFields(24).AddCurrentValue cOptions.VATRegNo
    Else
        xReport.ParameterFields(24).AddCurrentValue GetRegVATNo(BranchID)
    End If
    xReport.ParameterFields(25).AddCurrentValue PayDes

    If IntReportStyle = 0 Then
        xReport.ParameterFields(11).AddCurrentValue ToHijriDate(Transaction_Date)
        'xReport.ParameterFields(6).AddCurrentValue payed
        'xReport.ParameterFields(7).AddCurrentValue result
    End If
    xReport.ParameterFields(12).AddCurrentValue BranchVAT
    If reportid = 7 Then
        xReport.ParameterFields(18).AddCurrentValue ToHijriDate(Date)
    
    End If

    If reportid = 1 Or reportid = 7 Or reportid = 13 Or reportid = 17 Then
        xReport.ParameterFields(11).AddCurrentValue Nationality
        xReport.ParameterFields(12).AddCurrentValue CStr(GID)
        xReport.ParameterFields(13).AddCurrentValue ToHijriDate(Transaction_Date)

        xReport.ParameterFields(14).AddCurrentValue CStr(AdvPayment)
        xReport.ParameterFields(15).AddCurrentValue CStr(LblInstallCount)
        xReport.ParameterFields(16).AddCurrentValue CStr(LblPrecenValue)
        xReport.ParameterFields(17).AddCurrentValue CStr(LblFirstInstallDate)
  
    End If

    If reportid = 2 Then
        xReport.ParameterFields(11).AddCurrentValue Nationality
        xReport.ParameterFields(12).AddCurrentValue CStr(GID)
        xReport.ParameterFields(13).AddCurrentValue ToHijriDate(Transaction_Date)

        xReport.ParameterFields(14).AddCurrentValue CStr(AdvPayment)
        xReport.ParameterFields(15).AddCurrentValue CStr(LblInstallCount)
        xReport.ParameterFields(16).AddCurrentValue CStr(LblPrecenValue)
        xReport.ParameterFields(17).AddCurrentValue CStr(LblFirstInstallDate)
 
        xReport.ParameterFields(18).AddCurrentValue CStr(Madyna)
        xReport.ParameterFields(19).AddCurrentValue CStr(hay)
        xReport.ParameterFields(20).AddCurrentValue CStr(Address)
        xReport.ParameterFields(21).AddCurrentValue CStr(total)
        xReport.ParameterFields(22).AddCurrentValue CStr(WriteNo(Format(total, "0.00"), 0, True, ".", , 0))
  
    End If

    'xReport.ParameterFields(6).AddCurrentValue "2"
    'xReport.ParameterFields(7).AddCurrentValue "1"
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.reporttitle = "ÿ»«⁄… ›« Ê—… »Ì⁄"
    Else
        xReport.reporttitle = "Print Sales Invoice"
    End If
    Screen.MousePointer = vbDefault
    Set CViewer = New ClsReportViewer
    Dim P_Target As PrintTarget

    'If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
    '    P_Target = PrinterTarget
    'Else
    '    P_Target = WindowTarget
    'End If

    If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
        '  P_Target = PrinterTarget
        If X = vbYes Or X = 1 Then
            P_Target = PrinterTarget
        Else
            P_Target = WindowTarget
        End If
   
    Else
        P_Target = WindowTarget
    End If
     
    Dim StrFileName33 As String
   
    If SystemOptions.IsQrCodePrint Then
        'LoadPictureFromDB mPic, rs, "QrCodeImage", StrFileName33
        CViewer.CreateLogo xReport, , Trim(StrFileName33), val(rs!Transaction_ID & ""), SystemOptions.IsHeaderPrint
    Else
        CViewer.CreateLogo xReport
    End If
    
    CViewer.FireReport xReport, P_Target, , , , , StrFileName, , StrSQL, , , SystemOptions.IsHeaderPrint
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    Else
        Msg = "Sorry " & CHR(13) & "Can Not Print" & CHR(13)
    End If
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Function ShowSallingDataDetailed2(SallingID As Long, _
                                         Optional IntReportStyle As Integer = 0, _
                                         Optional LngCusID As Long, _
                                         Optional payed As Double, _
                                         Optional Result As String, _
                                         Optional CusCode As String, _
                                         Optional BillComment As String, _
                                         Optional InstComm As Double, _
                                         Optional currcode As Integer = 0, _
                                         Optional BranchVAT As String, _
                                         Optional brnchid As Integer, _
                                         Optional PayDes As String)
    On Error Resume Next
    Dim xApp            As New CRAXDRT.Application
    Dim xReport         As New CRAXDRT.Report
    Dim StrSQL          As String
    Dim StrRS           As New ADODB.Recordset
    Dim rs              As New ADODB.Recordset
    Dim RsTemp          As New ADODB.Recordset
    Dim ShowTax         As Boolean
    Dim Msg             As String
    Dim FrmReport       As FrmReportViewer
    Dim StrFileName     As String
    Dim SngCusBalBefore As Single
    Dim SngCusBalAfter  As Single
    Dim cOptions        As ClsCompanyInfo
    Dim CViewer         As ClsReportViewer
    Dim MySQL           As New ADODB.Recordset

    ' On Error GoTo ErrTrap

    If SallingID = 0 Then
        Exit Function
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        If IntReportStyle = 0 Then
            StrFileName = App.path & "\Reports\salesInvoice2.rpt"
                
        ElseIf IntReportStyle = 1 Then
            StrFileName = App.path & "\Reports\salesInvoice.rpt"
        ElseIf IntReportStyle = 2 Then
            StrFileName = App.path & "\Reports\salesInvoice.rpt"
        ElseIf IntReportStyle = 3 Then
            StrFileName = App.path & "\Reports\salesInvoice.rpt"
        ElseIf IntReportStyle = 40 Then
            StrFileName = App.path & "\Reports\salesInvoice2.rpt"
        End If

    Else

        If IntReportStyle = 0 Then
            StrFileName = App.path & "\Reports\salesInvoice2eng.rpt"
                
        ElseIf IntReportStyle = 1 Then
            StrFileName = App.path & "\Reports\salesInvoiceEng.rpt"
        ElseIf IntReportStyle = 2 Then
            StrFileName = App.path & "\Reports\salesInvoiceEng.rpt"
        ElseIf IntReportStyle = 3 Then
            StrFileName = App.path & "\Reports\salesInvoiceEng.rpt"
        End If

    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    End If

    '555555555555555555555555
    'updateserial
 
    Dim Rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset

    Dim CurrentSerial As String
 
    Dim Sql2          As String

    sql = "SELECT     TOP 100 PERCENT dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.Transaction_ID"
    sql = sql & "  FROM         dbo.Transaction_Details INNER JOIN"
    sql = sql & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    sql = sql & "  Where (dbo.TblItems.HaveSerial = 1)"
    sql = sql & "  GROUP BY dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.Transaction_ID"
    sql = sql & "  Having (dbo.Transaction_Details.Transaction_ID = " & SallingID & ")"

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        Rs1.MoveFirst
        Dim i, j As Integer

        For i = 1 To Rs1.RecordCount
            CurrentSerial = ""
            Set rs2 = New ADODB.Recordset
            Sql2 = "SELECT     dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ItemSerial"
            Sql2 = Sql2 & " FROM         dbo.Transaction_Details INNER JOIN"
            Sql2 = Sql2 & " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
            Sql2 = Sql2 & " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
            Sql2 = Sql2 & " WHERE     (dbo.TblItems.HaveSerial = 1) AND (dbo.Transactions.Transaction_ID = " & SallingID & ")"
            Sql2 = Sql2 & " AND (dbo.Transaction_Details.Item_ID = " & IIf(IsNull(Rs1(0).value), 0, Rs1(0).value) & ")"
            rs2.Open Sql2, Cn, adOpenStatic, adLockReadOnly, adCmdText

            For j = 1 To rs2.RecordCount
                CurrentSerial = CurrentSerial & IIf(IsNull(rs2(1).value), "", rs2(1).value) & ","
                rs2.MoveNext
            Next j

            rs2.Close
            Cn.Execute "update TblItems set itemSerials ='" & CurrentSerial & "' where ItemID=" & IIf(IsNull(Rs1(0).value), 0, Rs1(0).value)
            Rs1.MoveNext
        
        Next i

        'GetItemOrderNo = IIf(IsNull(Rs1("ORDERnOMin").value), "", Rs1("ORDERnOMin").value)

    End If

    '777777777777777777777777

    '   StrSQL = " SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.ShowQty) AS qty, dbo.Transaction_Details.showPrice, dbo.TblUnites.UnitName, dbo.TblItems.ItemCode, "
    '   StrSQL = StrSQL & "                   dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitNamee, dbo.TblItems.itemSerials, dbo.Transactions.Transaction_Date,"
    '   StrSQL = StrSQL & "                  dbo.Transactions.NoteSerial1, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee,"
    '   StrSQL = StrSQL & "                  dbo.Transactions.project_id, dbo.Transactions.PaymentType, dbo.Transactions.CustomerlocationID, dbo.TblCustomersLocations.Name AS CustomerLocation,"
    '   StrSQL = StrSQL & "                  dbo.TblCustomersLocations.NameE AS CustomerLocationE, dbo.Transactions.LocationID, dbo.TblCountriesGovernmentsCities.CityName,"
    '   StrSQL = StrSQL & "                  dbo.TblCountriesGovernmentsCities.NameE, dbo.Transactions.TransactionComment, dbo.Transactions.Doctype, dbo.TblDoCumentsTypes.name AS DoctypeName,"
    '   StrSQL = StrSQL & "                  dbo.TblDoCumentsTypes.namee AS DoctypeNameE, dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ReminRequ, dbo.Transaction_Details.ColorID,"
    '   StrSQL = StrSQL & "                  dbo.TblItemsColors.ColorName, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, dbo.Transaction_Details.Height, dbo.Transaction_Details.Width,"
    '   StrSQL = StrSQL & "                  dbo.Transactions.VATNO, dbo.Transactions.Typ, dbo.Transactions.ResonVAT, dbo.Transactions.VAT, dbo.Transactions.BatchNo,"
    '   StrSQL = StrSQL & "                  dbo.Transactions.Transaction_NetValue, dbo.Transactions.DueDate, dbo.Transactions.AddValue, dbo.Transactions.AddToTotal, dbo.Transaction_Details.Vatyo,"
    '   StrSQL = StrSQL & "                  dbo.Transaction_Details.Vat AS VatDes, dbo.Transaction_Details.H2, dbo.Transaction_Details.H1, dbo.Transaction_Details.W, dbo.Transaction_Details.L,"
    '   StrSQL = StrSQL & "                  dbo.Transaction_Details.PrintName, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail, dbo.TblCustemers.Address, dbo.TblCustemers.Cus_Phone,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.VATNO AS CusVATNO, dbo.TblCustemers.Mail2,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.CustGIDPlace, dbo.TblCustemers.BankAddress, dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankCode, dbo.TblCustemers.IBAN,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.RsIDDateH, dbo.TblCustemers.RecordDateH, dbo.TblCustemers.RecordNo, dbo.TblCustemers.BankName, dbo.TblCustemers.BankAccount,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.RsIDDate, dbo.TblCustemers.RsDegree, dbo.TblCustemers.authorizationname, dbo.TblCustemers.authorizationNo, dbo.TblCustemers.Map,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.JobName, dbo.TblCustemers.Entry, dbo.TblCustemers.RecordDate, dbo.TblCustemers.ZipCode, dbo.TblCustemers.BoxMil, dbo.TblCustemers.Sex,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.Mobile2, dbo.TblCustemers.Mobile1, dbo.TblCustemers.HomeTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.JobTel,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.JobAddress, dbo.TblCustemers.Salary, dbo.TblCustemers.JobTitle, dbo.TblCustemers.Company, dbo.TblCustemers.ExpireDateH,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.CustGID, dbo.TblCustemers.CustomerTypeID, dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone,"
    '   StrSQL = StrSQL & "                  dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress, dbo.Transactions.CashCustomerComment, dbo.Transaction_Details.QtyFaqtors,"
    '   StrSQL = StrSQL & "                  dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty, dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.MaxUnitID,"
    '   StrSQL = StrSQL & "                  TblUnites_1.UnitName AS MaxUnitName, TblUnites_1.UnitNamee AS MaxUnitNameE"
    '   StrSQL = StrSQL & "      FROM         dbo.Transaction_Details INNER JOIN"
    '   StrSQL = StrSQL & "                  dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
    '   StrSQL = StrSQL & "                  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    '   StrSQL = StrSQL & "                  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    '   StrSQL = StrSQL & "                  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    '   StrSQL = StrSQL & "                  dbo.TblUnites TblUnites_1 ON dbo.Transaction_Details.MaxUnitID = TblUnites_1.UnitID LEFT OUTER JOIN"
    '   StrSQL = StrSQL & "                  dbo.TblDoCumentsTypes ON dbo.Transactions.Doctype = dbo.TblDoCumentsTypes.id LEFT OUTER JOIN"
    '   StrSQL = StrSQL & "                  dbo.TblCountriesGovernmentsCities ON dbo.Transactions.LocationID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
    '   StrSQL = StrSQL & "                  dbo.TblCustomersLocations ON dbo.Transactions.CustomerlocationID = dbo.TblCustomersLocations.ID LEFT OUTER JOIN"
    '   StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID CROSS JOIN"
    '   StrSQL = StrSQL & "                  dbo.TblItemsColors TblItemsColors_1"
    '   StrSQL = StrSQL & "            GROUP BY dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.showPrice, dbo.Transactions.Transaction_ID,"
    '   StrSQL = StrSQL & "                  dbo.Transactions.Transaction_Type, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
    '   StrSQL = StrSQL & "                   dbo.TblItems.itemSerials, dbo.Transactions.Transaction_Date, dbo.Transactions.NoteSerial1, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    '   StrSQL = StrSQL & "                  dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.project_id, dbo.Transactions.PaymentType, dbo.Transactions.CustomerlocationID,"
    '   StrSQL = StrSQL & "                  dbo.TblCustomersLocations.Name, dbo.TblCustomersLocations.NameE, dbo.Transactions.LocationID, dbo.TblCountriesGovernmentsCities.CityName,"
    '   StrSQL = StrSQL & "                  dbo.TblCountriesGovernmentsCities.NameE, dbo.Transactions.TransactionComment, dbo.Transactions.Doctype, dbo.TblDoCumentsTypes.name,"
    '   StrSQL = StrSQL & "                  dbo.TblDoCumentsTypes.namee, dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ReminRequ, dbo.Transaction_Details.ColorID,"
    '   StrSQL = StrSQL & "                  dbo.TblItemsColors.ColorName, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, dbo.Transaction_Details.Height, dbo.Transaction_Details.Width,"
    '   StrSQL = StrSQL & "                  dbo.Transactions.VATNO, dbo.Transactions.Typ, dbo.Transactions.ResonVAT, dbo.Transactions.VAT, dbo.Transactions.BatchNo,"
    '   StrSQL = StrSQL & "                  dbo.Transactions.Transaction_NetValue, dbo.Transactions.DueDate, dbo.Transactions.AddValue, dbo.Transactions.AddToTotal, dbo.Transaction_Details.Vatyo,"
    '   StrSQL = StrSQL & "                  dbo.Transaction_Details.Vat, dbo.Transaction_Details.H2, dbo.Transaction_Details.H1, dbo.Transaction_Details.W, dbo.Transaction_Details.L,"
    '   StrSQL = StrSQL & "                  dbo.Transaction_Details.PrintName, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail, dbo.TblCustemers.Address, dbo.TblCustemers.Cus_Phone,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.VATNO, dbo.TblCustemers.Mail2, dbo.TblCustemers.CustGIDPlace,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.BankAddress, dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankCode, dbo.TblCustemers.IBAN, dbo.TblCustemers.RsIDDateH,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.RecordDateH, dbo.TblCustemers.RecordNo, dbo.TblCustemers.BankName, dbo.TblCustemers.BankAccount, dbo.TblCustemers.RsIDDate,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.RsDegree, dbo.TblCustemers.authorizationname, dbo.TblCustemers.authorizationNo, dbo.TblCustemers.Map, dbo.TblCustemers.JobName,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.Entry, dbo.TblCustemers.RecordDate, dbo.TblCustemers.ZipCode, dbo.TblCustemers.BoxMil, dbo.TblCustemers.Sex, dbo.TblCustemers.Mobile2,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.Mobile1, dbo.TblCustemers.HomeTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.JobTel, dbo.TblCustemers.JobAddress,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.Salary, dbo.TblCustemers.JobTitle, dbo.TblCustemers.Company, dbo.TblCustemers.ExpireDateH, dbo.TblCustemers.CustGID,"
    '   StrSQL = StrSQL & "                  dbo.TblCustemers.CustomerTypeID, dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile,"
    '   StrSQL = StrSQL & "                  dbo.Transactions.CashCustomerAddress, dbo.Transactions.CashCustomerComment, dbo.Transaction_Details.QtyFaqtors, dbo.Transaction_Details.MixNo,"
    '   StrSQL = StrSQL & "                 dbo.Transaction_Details.MaxQty , dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.MaxUnitID, TblUnites_1.UnitName, TblUnites_1.UnitNamee"

    'Wael „—«Ã⁄…
    'transactions.QrCodeImage
    StrSQL = " SELECT     Sum(Transactions.TotalTaxExempt ) as TotalTaxExempt ,SUM(dbo.Transaction_Details.ShowQty) AS qty,ExtraValue, dbo.Transaction_Details.showPrice, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, "
    StrSQL = StrSQL & "                      dbo.TblItems.itemSerials, dbo.Transactions.Transaction_Date, dbo.Transactions.NoteSerial1, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    StrSQL = StrSQL & "                      dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.project_id, dbo.Transactions.PaymentType, dbo.Transactions.CustomerlocationID,"
    StrSQL = StrSQL & "                      dbo.TblCustomersLocations.Name AS CustomerLocation, dbo.TblCustomersLocations.NameE AS CustomerLocationE, dbo.Transactions.LocationID,"
    StrSQL = StrSQL & "                      dbo.TblCountriesGovernmentsCities.CityName, dbo.TblCountriesGovernmentsCities.NameE, dbo.Transactions.TransactionComment, dbo.Transactions.Doctype,"
    StrSQL = StrSQL & "                      dbo.TblDoCumentsTypes.name AS DoctypeName, dbo.TblDoCumentsTypes.namee AS DoctypeNameE, dbo.Transaction_Details.LotNO,"
    StrSQL = StrSQL & "                      dbo.Transaction_Details.ReminRequ, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area,"
    StrSQL = StrSQL & "                      Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width, dbo.Transactions.VATNO, dbo.Transactions.Typ, dbo.Transactions.ResonVAT, dbo.Transactions.VAT,"
    StrSQL = StrSQL & "                      dbo.Transactions.BatchNo, dbo.Transactions.Transaction_NetValue, dbo.Transactions.DueDate, dbo.Transactions.AddValue, dbo.Transactions.AddToTotal,"
    StrSQL = StrSQL & "                      dbo.Transaction_Details.Vatyo, dbo.Transaction_Details.Vat AS VatDes, dbo.Transaction_Details.H2, dbo.Transaction_Details.H1, dbo.Transaction_Details.W,"
    StrSQL = StrSQL & "                      dbo.Transaction_Details.L, dbo.Transaction_Details.PrintName, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail, dbo.TblCustemers.Address,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.VATNO AS CusVATNO,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.Mail2, dbo.TblCustemers.CustGIDPlace, dbo.TblCustemers.BankAddress, dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankCode,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.IBAN, dbo.TblCustemers.RsIDDateH, dbo.TblCustemers.RecordDateH, dbo.TblCustemers.RecordNo, dbo.TblCustemers.BankName,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.BankAccount, dbo.TblCustemers.RsIDDate, dbo.TblCustemers.RsDegree, dbo.TblCustemers.authorizationname, dbo.TblCustemers.authorizationNo,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.Map, dbo.TblCustemers.JobName, dbo.TblCustemers.Entry, dbo.TblCustemers.RecordDate, dbo.TblCustemers.ZipCode, dbo.TblCustemers.BoxMil,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.Sex, dbo.TblCustemers.Mobile2, dbo.TblCustemers.Mobile1, dbo.TblCustemers.HomeTel, dbo.TblCustemers.JobTelConvert,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.JobTel, dbo.TblCustemers.JobAddress, dbo.TblCustemers.Salary, dbo.TblCustemers.JobTitle, dbo.TblCustemers.Company,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.ExpireDateH, dbo.TblCustemers.CustGID, dbo.TblCustemers.CustomerTypeID, dbo.Transactions.CashCustomerName,"
    StrSQL = StrSQL & "                      dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress, dbo.Transactions.CashCustomerComment,"
    StrSQL = StrSQL & "                      dbo.Transaction_Details.QtyFaqtors, dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty, dbo.Transaction_Details.TypeVAT,"
    StrSQL = StrSQL & "                      dbo.Transaction_Details.StoreID2, TblStore_1.StoreName AS StoreNameDet, TblStore_1.StoreNamee AS StoreNameDetE, dbo.Transaction_Details.Remarks,"
    StrSQL = StrSQL & "                      dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.MaxUnitID, TblUnites_1.UnitName AS MaxUnitName,"
    StrSQL = StrSQL & "                      TblUnites_1.UnitNamee AS MaxUnitNameE ,dbo.Transaction_Details.TotalPrice,TblItems.ItemWithOutVAT"
    
    
    StrSQL = StrSQL & " FROM         dbo.Transaction_Details INNER JOIN"
    StrSQL = StrSQL & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "                      dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    StrSQL = StrSQL & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL & "                      dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblUnites TblUnites_1 ON dbo.Transaction_Details.MaxUnitID = TblUnites_1.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblStore TblStore_1 ON dbo.Transaction_Details.StoreID2 = TblStore_1.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblDoCumentsTypes ON dbo.Transactions.Doctype = dbo.TblDoCumentsTypes.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblCountriesGovernmentsCities ON dbo.Transactions.LocationID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblCustomersLocations ON dbo.Transactions.CustomerlocationID = dbo.TblCustomersLocations.ID"
    StrSQL = StrSQL & " GROUP BY dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.showPrice, dbo.Transactions.Transaction_ID,"
    StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Type, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.itemSerials,"
    StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Date, dbo.Transactions.NoteSerial1, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName,"
    StrSQL = StrSQL & "                      dbo.TblStore.StoreNamee, dbo.Transactions.project_id, dbo.Transactions.PaymentType, dbo.Transactions.CustomerlocationID, dbo.TblCustomersLocations.Name,"
    StrSQL = StrSQL & "                      dbo.TblCustomersLocations.NameE, dbo.Transactions.LocationID, dbo.TblCountriesGovernmentsCities.CityName, dbo.TblCountriesGovernmentsCities.NameE,"
    StrSQL = StrSQL & "                      dbo.Transactions.TransactionComment, dbo.Transactions.Doctype, dbo.TblDoCumentsTypes.name, dbo.TblDoCumentsTypes.namee, dbo.Transaction_Details.LotNO,"
    StrSQL = StrSQL & "                      dbo.Transaction_Details.ReminRequ, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area,"
    StrSQL = StrSQL & "                      Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width, dbo.Transactions.VATNO, dbo.Transactions.Typ, dbo.Transactions.ResonVAT, dbo.Transactions.VAT,"
    StrSQL = StrSQL & "                      dbo.Transactions.BatchNo, dbo.Transactions.Transaction_NetValue, dbo.Transactions.DueDate, dbo.Transactions.AddValue, dbo.Transactions.AddToTotal,"
    StrSQL = StrSQL & "                      dbo.Transaction_Details.Vatyo, dbo.Transaction_Details.Vat, dbo.Transaction_Details.H2, dbo.Transaction_Details.H1, dbo.Transaction_Details.W,"
    StrSQL = StrSQL & "                      dbo.Transaction_Details.L, dbo.Transaction_Details.PrintName, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail, dbo.TblCustemers.Address,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.VATNO, dbo.TblCustemers.Mail2,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.CustGIDPlace, dbo.TblCustemers.BankAddress, dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankCode, dbo.TblCustemers.IBAN,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.RsIDDateH, dbo.TblCustemers.RecordDateH, dbo.TblCustemers.RecordNo, dbo.TblCustemers.BankName, dbo.TblCustemers.BankAccount,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.RsIDDate, dbo.TblCustemers.RsDegree, dbo.TblCustemers.authorizationname, dbo.TblCustemers.authorizationNo, dbo.TblCustemers.Map,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.JobName, dbo.TblCustemers.Entry, dbo.TblCustemers.RecordDate, dbo.TblCustemers.ZipCode, dbo.TblCustemers.BoxMil, dbo.TblCustemers.Sex,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.Mobile2, dbo.TblCustemers.Mobile1, dbo.TblCustemers.HomeTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.JobTel,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.JobAddress, dbo.TblCustemers.Salary, dbo.TblCustemers.JobTitle, dbo.TblCustemers.Company, dbo.TblCustemers.ExpireDateH,"
    StrSQL = StrSQL & "                      dbo.TblCustemers.CustGID, dbo.TblCustemers.CustomerTypeID, dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone,"
    StrSQL = StrSQL & "                      dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress, dbo.Transactions.CashCustomerComment, dbo.Transaction_Details.QtyFaqtors,"
    StrSQL = StrSQL & "                      dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty, dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.StoreID2, TblStore_1.StoreName,"
    StrSQL = StrSQL & "                      TblStore_1.StoreNamee, dbo.Transaction_Details.Remarks, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.MaxUnitID,"
    StrSQL = StrSQL & "                      TblUnites_1.UnitName , TblUnites_1.UnitNamee ,dbo.Transaction_Details.TotalPrice"
    
    StrSQL = StrSQL & "  Having    (dbo.Transactions.Transaction_Type = 21)  and (dbo.Transactions.Transaction_ID = " & SallingID & ")"
    
    StrSQL = ""
    StrSQL = StrSQL & "      SELECT      Sum(Transactions.TotalTaxExempt ) as TotalTaxExempt ,  dbo.Transaction_Details.ShowQty AS qty, dbo.Transaction_Details.showPrice, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.itemSerials, dbo.Transactions.Transaction_Date,"
    StrSQL = StrSQL & "                           dbo.Transactions.NoteSerial1, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.project_id, dbo.Transactions.PaymentType,"
    StrSQL = StrSQL & "                           dbo.Transactions.CustomerlocationID, dbo.TblCustomersLocations.Name AS CustomerLocation, dbo.TblCustomersLocations.NameE AS CustomerLocationE, dbo.Transactions.LocationID,"
    StrSQL = StrSQL & "                           dbo.TblCountriesGovernmentsCities.CityName, dbo.TblCountriesGovernmentsCities.NameE, dbo.Transactions.TransactionComment, dbo.Transactions.Doctype, dbo.TblDoCumentsTypes.name AS DoctypeName,"
    StrSQL = StrSQL & "                           dbo.TblDoCumentsTypes.namee AS DoctypeNameE, dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ReminRequ, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area,"
    StrSQL = StrSQL & "                           Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width, dbo.Transactions.VATNO, dbo.Transactions.Typ, dbo.Transactions.ResonVAT, dbo.Transactions.VAT, dbo.Transactions.BatchNo,"
    StrSQL = StrSQL & "                           dbo.Transactions.Transaction_NetValue, dbo.Transactions.DueDate, dbo.Transactions.AddValue, dbo.Transactions.AddToTotal, dbo.Transaction_Details.Vatyo, dbo.Transaction_Details.Vat AS VatDes,"
    StrSQL = StrSQL & "                           dbo.Transaction_Details.H2, dbo.Transaction_Details.H1, dbo.Transaction_Details.W, dbo.Transaction_Details.L, dbo.Transaction_Details.PrintName, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail,"
    StrSQL = StrSQL & "                           dbo.TblCustemers.Address, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.VATNO AS CusVATNO, dbo.TblCustemers.Mail2,"
    StrSQL = StrSQL & "                           dbo.TblCustemers.CustGIDPlace, dbo.TblCustemers.BankAddress, dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankCode, dbo.TblCustemers.IBAN, dbo.TblCustemers.RsIDDateH, dbo.TblCustemers.RecordDateH,"
    StrSQL = StrSQL & "                           dbo.TblCustemers.RecordNo, dbo.TblCustemers.BankName, dbo.TblCustemers.BankAccount, dbo.TblCustemers.RsIDDate, dbo.TblCustemers.RsDegree, dbo.TblCustemers.authorizationname,"
    StrSQL = StrSQL & "                           dbo.TblCustemers.authorizationNo, dbo.TblCustemers.Map, dbo.TblCustemers.JobName, dbo.TblCustemers.Entry, dbo.TblCustemers.RecordDate, dbo.TblCustemers.ZipCode, dbo.TblCustemers.BoxMil, dbo.TblCustemers.Sex,"
    StrSQL = StrSQL & "                           dbo.TblCustemers.Mobile2, dbo.TblCustemers.Mobile1, dbo.TblCustemers.HomeTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.JobTel, dbo.TblCustemers.JobAddress, dbo.TblCustemers.Salary,"
    StrSQL = StrSQL & "                           dbo.TblCustemers.JobTitle, dbo.TblCustemers.Company, dbo.TblCustemers.ExpireDateH, dbo.TblCustemers.CustGID, dbo.TblCustemers.CustomerTypeID, dbo.Transactions.CashCustomerName,"
    StrSQL = StrSQL & "                           dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress, dbo.Transactions.CashCustomerComment, dbo.Transaction_Details.QtyFaqtors,"
    StrSQL = StrSQL & "                           dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty, dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.StoreID2, TblStore_1.StoreName AS StoreNameDet, TblStore_1.StoreNamee AS StoreNameDetE,"
    StrSQL = StrSQL & "                           dbo.Transaction_Details.Remarks, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.MaxUnitID, TblUnites_1.UnitName AS MaxUnitName,"
    StrSQL = StrSQL & "                           TblUnites_1.UnitNamee AS MaxUnitNameE, dbo.Transaction_Details.TotalPrice, dbo.Transactions.QrCodeData, dbo.Transactions.QrCodeDataPath, CAST(dbo.Transactions.QrCodeImage AS Varbinary) AS QrCodeImage,TblItems.ItemWithOutVAT"
    StrSQL = StrSQL & "  FROM            dbo.Transaction_Details INNER JOIN"
    StrSQL = StrSQL & "                           dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "                           dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "                           dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    StrSQL = StrSQL & "                           dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL & "                           dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                           dbo.TblUnites AS TblUnites_1 ON dbo.Transaction_Details.MaxUnitID = TblUnites_1.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                           dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                           dbo.TblStore AS TblStore_1 ON dbo.Transaction_Details.StoreID2 = TblStore_1.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                           dbo.TblDoCumentsTypes ON dbo.Transactions.Doctype = dbo.TblDoCumentsTypes.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                           dbo.TblCountriesGovernmentsCities ON dbo.Transactions.LocationID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                           dbo.TblCustomersLocations ON dbo.Transactions.CustomerlocationID = dbo.TblCustomersLocations.ID"
    StrSQL = StrSQL & "  Where    (dbo.Transactions.Transaction_Type = 21)  and (dbo.Transactions.Transaction_ID = " & SallingID & ")"
    StrSQL = StrSQL & "  GROUP BY dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.itemSerials, dbo.Transactions.Transaction_Date,"
    StrSQL = StrSQL & "                           dbo.Transactions.NoteSerial1, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.Transactions.project_id, dbo.Transactions.PaymentType,"
    StrSQL = StrSQL & "                           dbo.Transactions.CustomerlocationID, dbo.TblCustomersLocations.Name, dbo.TblCustomersLocations.NameE, dbo.Transactions.LocationID, dbo.TblCountriesGovernmentsCities.CityName,"
    StrSQL = StrSQL & "                           dbo.TblCountriesGovernmentsCities.NameE, dbo.Transactions.TransactionComment, dbo.Transactions.Doctype, dbo.TblDoCumentsTypes.name, dbo.TblDoCumentsTypes.namee, dbo.Transaction_Details.LotNO,"
    StrSQL = StrSQL & "                           dbo.Transaction_Details.ReminRequ, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width, dbo.Transactions.VATNO,"
    StrSQL = StrSQL & "                            dbo.Transactions.Typ, dbo.Transactions.ResonVAT, dbo.Transactions.VAT, dbo.Transactions.BatchNo, dbo.Transactions.Transaction_NetValue, dbo.Transactions.DueDate, dbo.Transactions.AddValue,"
    StrSQL = StrSQL & "                           dbo.Transactions.AddToTotal, dbo.Transaction_Details.Vatyo, dbo.Transaction_Details.Vat, dbo.Transaction_Details.H2, dbo.Transaction_Details.H1, dbo.Transaction_Details.W, dbo.Transaction_Details.L,"
    StrSQL = StrSQL & "                           dbo.Transaction_Details.PrintName, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail, dbo.TblCustemers.Address, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.ResponsibleContact,"
    StrSQL = StrSQL & "                           dbo.TblCustemers.VATNO, dbo.TblCustemers.Mail2, dbo.TblCustemers.CustGIDPlace, dbo.TblCustemers.BankAddress, dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankCode, dbo.TblCustemers.IBAN,"
    StrSQL = StrSQL & "                           dbo.TblCustemers.RsIDDateH, dbo.TblCustemers.RecordDateH, dbo.TblCustemers.RecordNo, dbo.TblCustemers.BankName, dbo.TblCustemers.BankAccount, dbo.TblCustemers.RsIDDate, dbo.TblCustemers.RsDegree,"
    StrSQL = StrSQL & "                           dbo.TblCustemers.authorizationname, dbo.TblCustemers.authorizationNo, dbo.TblCustemers.Map, dbo.TblCustemers.JobName, dbo.TblCustemers.Entry, dbo.TblCustemers.RecordDate, dbo.TblCustemers.ZipCode,"
    StrSQL = StrSQL & "                           dbo.TblCustemers.BoxMil, dbo.TblCustemers.Sex, dbo.TblCustemers.Mobile2, dbo.TblCustemers.Mobile1, dbo.TblCustemers.HomeTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.JobTel,"
    StrSQL = StrSQL & "                           dbo.TblCustemers.JobAddress, dbo.TblCustemers.Salary, dbo.TblCustemers.JobTitle, dbo.TblCustemers.Company, dbo.TblCustemers.ExpireDateH, dbo.TblCustemers.CustGID, dbo.TblCustemers.CustomerTypeID,"
    StrSQL = StrSQL & "                           dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress, dbo.Transactions.CashCustomerComment,"
    StrSQL = StrSQL & "                           dbo.Transaction_Details.QtyFaqtors, dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty, dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.StoreID2, TblStore_1.StoreName, TblStore_1.StoreNamee,"
    StrSQL = StrSQL & "                           dbo.Transaction_Details.Remarks, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.MaxUnitID, TblUnites_1.UnitName, TblUnites_1.UnitNamee,"
    StrSQL = StrSQL & "                           dbo.Transaction_Details.TotalPrice, dbo.Transactions.QrCodeData, dbo.Transactions.QrCodeDataPath, CAST(dbo.Transactions.QrCodeImage AS Varbinary)"

    '
    'StrSQL = StrSQL + "ORDER BY dbo.Transaction_Details.ID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Screen.MousePointer = vbArrowHourglass
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    Set cOptions = New ClsCompanyInfo
    
    xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(5).AddCurrentValue BillComment

    If val(InstComm) >= 0 Then
        xReport.ParameterFields(10).AddCurrentValue val(InstComm)
    End If

    '
    xReport.ParameterFields(8).AddCurrentValue CStr(WriteNo(Format(Result, "0.00"), 0, True, ".", , 0, , currcode))
    ' xReport.ParameterFields(8).AddCurrentValue WriteNo(CStr(result), 0, True, , , , , currcode)
    xReport.ParameterFields(9).AddCurrentValue CusCode
    xReport.ParameterFields(11).AddCurrentValue BranchVAT
    
    If SystemOptions.VATNoAccordActivity = False Then
        xReport.ParameterFields(24).AddCurrentValue cOptions.VATRegNo
    Else
        xReport.ParameterFields(24).AddCurrentValue GetRegVATNo(brnchid)
    End If
    xReport.ParameterFields(25).AddCurrentValue PayDes

    If IntReportStyle = 0 Then
        'xReport.ParameterFields(6).AddCurrentValue payed
        'xReport.ParameterFields(7).AddCurrentValue result
    End If

    'xReport.ParameterFields(6).AddCurrentValue "2"
    'xReport.ParameterFields(7).AddCurrentValue "1"

    xReport.reporttitle = "ÿ»«⁄… ›« Ê—… »Ì⁄"
 
    Screen.MousePointer = vbDefault
    Set CViewer = New ClsReportViewer
    Dim P_Target As PrintTarget

    If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
        P_Target = PrinterTarget
    Else
        P_Target = WindowTarget
    End If

    CViewer.FireReport xReport, P_Target, , , , , StrFileName, , StrSQL
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

