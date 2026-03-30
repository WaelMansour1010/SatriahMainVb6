VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBuyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

 
'Public Function ShowRecieveVoucherData(BuyingID As Long, _
                                       Optional ShortDetails As Boolean = False _
                                       , Optional title As String, Optional tostorename As String = "")

Public Function ShowMovingVOucherData(BuyingID As Long, _
                                      Optional ShortDetails As Boolean = False)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Reportname  As String
    Dim xstr As String
    On Error GoTo ErrTrap

    If BuyingID = 0 Then
        Exit Function
    End If

    'StrSQL = "select * From QryBuyReport where Transaction_ID=" & BuyingID
    'StrSQL = StrSQL + "order by id"
    If SystemOptions.UserInterface = ArabicInterface Then
        If ShortDetails = False Then
           ' ReportName = App.path & "\Reports\MovingVoucher.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\MovingVoucher.rpt"
        Else
            'ReportName = App.path & "\Reports\MovingVoucherShort.rpt"
             Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\MovingVoucherShort.rpt"
        End If

    Else
 
        If ShortDetails = False Then
            'ReportName = App.path & "\Reports\MovingVoucherEng.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\MovingVoucherEng.rpt"
        Else
            'ReportName = App.path & "\Reports\MovingVoucherShortEng.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\MovingVoucherShortEng.rpt"
        End If
        
    End If

   ' StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName,"
   ' StrSQL = StrSQL & "  dbo.TblItems.ItemCode, dbo.Transaction_Details.ItemCase, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.TblItems.HaveSerial,"
   ' StrSQL = StrSQL & " Transactions_1.StoreID, TblStore_1.StoreName AS secondStore, dbo.TblStore.StoreNamee, TblStore_1.StoreNamee AS SecondStorename, dbo.TblItems.ItemNamee,"
   ' StrSQL = StrSQL & " dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.CostPrice, dbo.Transaction_Details.ColorID,"
   ' StrSQL = StrSQL & " dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblBranchesData.branch_name,"
   ' StrSQL = StrSQL & " dbo.TblBranchesData.branch_namee, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
   ' StrSQL = StrSQL & " FROM         dbo.TblStore INNER JOIN"
   ' StrSQL = StrSQL & " dbo.Transactions INNER JOIN"
   ' StrSQL = StrSQL & " dbo.TblStore TblStore_1 INNER JOIN"
   ' StrSQL = StrSQL & " dbo.Transactions Transactions_1 ON TblStore_1.StoreID = Transactions_1.StoreID ON dbo.Transactions.Transaction_ID = Transactions_1.ReturnID ON"
   ' StrSQL = StrSQL & " dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
   ' StrSQL = StrSQL & " dbo.TblItems INNER JOIN"
   ' StrSQL = StrSQL & " dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
   ' StrSQL = StrSQL & " dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
   ' StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
   ' StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.Transaction_Details.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
   ' StrSQL = StrSQL & " dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
   ' StrSQL = StrSQL & " dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
   ' StrSQL = StrSQL & " dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"




    StrSQL = " SELECT   dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemCode,TblItems.barCodeNO ,"
    StrSQL = StrSQL & "                     dbo.Transaction_Details.ItemCase, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.TblItems.HaveSerial, Transactions_1.StoreID, TblStore_1.StoreName AS secondStore,"
    StrSQL = StrSQL & "                     dbo.TblStore.StoreNamee, TblStore_1.StoreNamee AS SecondStorename, dbo.TblItems.ItemNamee, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.CostPrice,"
    StrSQL = StrSQL & "                     dbo.Transaction_Details.ColorID, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, TblBranchesData_2.branch_name,"
    StrSQL = StrSQL & "                     TblBranchesData_2.branch_namee, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.Transactions.CashCustomerName,"
    StrSQL = StrSQL & "                     dbo.Transactions.TransactionComment, dbo.Transactions.CusID, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblCustemers.CusName,"
    StrSQL = StrSQL & "                     dbo.TblCustemers.CusNamee, dbo.Transactions.Phone, dbo.Transactions.InspectionReport, TblStore_1.Code AS StoreFullcode1, dbo.TblStore.Code AS StoreFullcode, dbo.TblItems.PartNo,"
    StrSQL = StrSQL & "                     dbo.Transaction_Details.Remarks, dbo.TblCustemers.Fullcode, dbo.TblStore.BranchId, TblBranchesData_1.branch_name AS Frombranch_name, TblBranchesData_1.branch_namee AS Frombranch_nameE,"
    StrSQL = StrSQL & "                     TblStore_1.BranchId AS ToBranchId, TblBranchesData_2.branch_name AS Tobranch_name, TblBranchesData_2.branch_namee AS Tobranch_nameE, dbo.GetUnitSalesPrice(dbo.Transaction_Details.Item_ID,"
    StrSQL = StrSQL & "                     dbo.Transaction_Details.UnitId) AS SalesPrice, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.Price , dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.LotNO"
    StrSQL = StrSQL & "     FROM            dbo.TblStore INNER JOIN"
    StrSQL = StrSQL & "                     dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "                     dbo.TblStore AS TblStore_1 INNER JOIN"
    StrSQL = StrSQL & "                     dbo.Transactions AS Transactions_1 ON TblStore_1.StoreID = Transactions_1.StoreID ON dbo.Transactions.Transaction_ID = Transactions_1.ReturnID ON"
    StrSQL = StrSQL & "                     dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
    StrSQL = StrSQL & "                     dbo.TblItems INNER JOIN"
    StrSQL = StrSQL & "                     dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblBranchesData AS TblBranchesData_2 ON TblStore_1.BranchId = TblBranchesData_2.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblBranchesData AS TblBranchesData_1 ON dbo.TblStore.BranchId = TblBranchesData_1.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblBranchesData AS TblBranchesData_3 ON dbo.Transaction_Details.BranchId = TblBranchesData_3.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
'Where (dbo.Transactions.Transaction_ID = 17)
'ORDER BY dbo.Transaction_Details.ID

    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_ID = " & BuyingID & ")"
    StrSQL = StrSQL + "order by dbo.Transaction_Details.id"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
xstr = StrSQL

    If Dir(Reportname) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Reportname
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(Reportname)
        xReport.Database.SetDataSource rs
         
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        StrSQL = "select * From Notes where Transaction_ID=" & BuyingID
        StrSQL = StrSQL + " and NoteType=0"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
    
        End If

        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , Reportname, , xstr
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
   ' Screen.MousePointer = vbDefault
End Function



Public Function ShowMovingVOucherDatax08062017(BuyingID As Long, _
                                      Optional ShortDetails As Boolean = False)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Reportname  As String
    On Error GoTo ErrTrap

    If BuyingID = 0 Then
        Exit Function
    End If

    'StrSQL = "select * From QryBuyReport where Transaction_ID=" & BuyingID
    'StrSQL = StrSQL + "order by id"
    If SystemOptions.UserInterface = ArabicInterface Then
        If ShortDetails = False Then
           ' ReportName = App.path & "\Reports\MovingVoucher.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\MovingVoucher.rpt"
        Else
            'ReportName = App.path & "\Reports\MovingVoucherShort.rpt"
             Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\MovingVoucherShort.rpt"
        End If

    Else
 
        If ShortDetails = False Then
            'ReportName = App.path & "\Reports\MovingVoucherEng.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\MovingVoucherEng.rpt"
        Else
            'ReportName = App.path & "\Reports\MovingVoucherShortEng.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\MovingVoucherShortEng.rpt"
        End If
        
    End If

   ' StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName,"
   ' StrSQL = StrSQL & "  dbo.TblItems.ItemCode, dbo.Transaction_Details.ItemCase, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.TblItems.HaveSerial,"
   ' StrSQL = StrSQL & " Transactions_1.StoreID, TblStore_1.StoreName AS secondStore, dbo.TblStore.StoreNamee, TblStore_1.StoreNamee AS SecondStorename, dbo.TblItems.ItemNamee,"
   ' StrSQL = StrSQL & " dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.CostPrice, dbo.Transaction_Details.ColorID,"
   ' StrSQL = StrSQL & " dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblBranchesData.branch_name,"
   ' StrSQL = StrSQL & " dbo.TblBranchesData.branch_namee, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
   ' StrSQL = StrSQL & " FROM         dbo.TblStore INNER JOIN"
   ' StrSQL = StrSQL & " dbo.Transactions INNER JOIN"
   ' StrSQL = StrSQL & " dbo.TblStore TblStore_1 INNER JOIN"
   ' StrSQL = StrSQL & " dbo.Transactions Transactions_1 ON TblStore_1.StoreID = Transactions_1.StoreID ON dbo.Transactions.Transaction_ID = Transactions_1.ReturnID ON"
   ' StrSQL = StrSQL & " dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
   ' StrSQL = StrSQL & " dbo.TblItems INNER JOIN"
   ' StrSQL = StrSQL & " dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
   ' StrSQL = StrSQL & " dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
   ' StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
   ' StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.Transaction_Details.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
   ' StrSQL = StrSQL & " dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
   ' StrSQL = StrSQL & " dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
   ' StrSQL = StrSQL & " dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"




StrSQL = " SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.Transaction_Details.Item_ID, "
StrSQL = StrSQL & "                      dbo.TblItems.ItemName, dbo.TblItems.ItemCode, dbo.Transaction_Details.ItemCase, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity,"
StrSQL = StrSQL & "                      dbo.TblItems.HaveSerial, Transactions_1.StoreID, TblStore_1.StoreName AS secondStore, dbo.TblStore.StoreNamee, TblStore_1.StoreNamee AS SecondStorename,"
StrSQL = StrSQL & "                      dbo.TblItems.ItemNamee, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.CostPrice, dbo.Transaction_Details.ColorID,"
StrSQL = StrSQL & "                      dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, TblBranchesData_2.branch_name,"
StrSQL = StrSQL & "                      TblBranchesData_2.branch_namee, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1,"
StrSQL = StrSQL & "                      dbo.Transactions.CashCustomerName, dbo.Transactions.TransactionComment, dbo.Transactions.CusID, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.Phone, dbo.Transactions.InspectionReport,"
StrSQL = StrSQL & "                      TblStore_1.Code AS StoreFullcode1, dbo.TblStore.Code AS StoreFullcode, dbo.TblItems.PartNo, dbo.Transaction_Details.Remarks, dbo.TblCustemers.Fullcode,"
StrSQL = StrSQL & "                      dbo.TblStore.BranchId, TblBranchesData_1.branch_name AS Frombranch_name, TblBranchesData_1.branch_namee AS Frombranch_nameE,"
StrSQL = StrSQL & "                      TblStore_1.BranchId AS ToBranchId, TblBranchesData_2.branch_name AS Tobranch_name, TblBranchesData_2.branch_namee AS Tobranch_nameE"
StrSQL = StrSQL & " FROM         dbo.TblStore INNER JOIN"
StrSQL = StrSQL & "                      dbo.Transactions INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblStore TblStore_1 INNER JOIN"
StrSQL = StrSQL & "                      dbo.Transactions Transactions_1 ON TblStore_1.StoreID = Transactions_1.StoreID ON dbo.Transactions.Transaction_ID = Transactions_1.ReturnID ON"
StrSQL = StrSQL & "                      dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems INNER JOIN"
StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
StrSQL = StrSQL & "                      dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_2 ON TblStore_1.BranchId = TblBranchesData_2.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_1 ON dbo.TblStore.BranchId = TblBranchesData_1.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_3 ON dbo.Transaction_Details.BranchId = TblBranchesData_3.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
'Where (dbo.Transactions.Transaction_ID = 17)
'ORDER BY dbo.Transaction_Details.ID

    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_ID = " & BuyingID & ")"
    StrSQL = StrSQL + "order by dbo.Transaction_Details.id"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(Reportname) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Reportname
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(Reportname)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        StrSQL = "select * From Notes where Transaction_ID=" & BuyingID
        StrSQL = StrSQL + " and NoteType=0"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
    
        End If

        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , Reportname
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function


'End Function
Public Function ShowRecieveVoucherData(BuyingID As Long, _
                                       Optional ShortDetails As Boolean = False _
                                       , Optional Title As String, Optional tostorename As String = "")
                                       
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Reportname  As String
    On Error GoTo ErrTrap

    If BuyingID = 0 Then
        Exit Function
    End If

    'StrSQL = "select * From QryBuyReport where Transaction_ID=" & BuyingID
    'StrSQL = StrSQL + "order by id"
    If SystemOptions.UserInterface = ArabicInterface Then
        If ShortDetails = False Then
           
                             
                             
         '   ReportName = App.path & "\Reports\RecieveVoucher2.rpt"
         Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\RecieveVoucher.rpt"
        Else
      '      ReportName = App.path & "\Reports\RecieveVoucher2.rpt"
      Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\RecieveVoucher.rpt"
        End If

    Else
 
        If ShortDetails = False Then
            'ReportName = App.path & "\Reports\RecieveVoucher2.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\RecieveVoucher.rpt"
        Else
            'ReportName = App.path & "\Reports\RecieveVoucher2.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\RecieveVoucher.rpt"
        End If
        
    End If

    StrSQL = "SELECT      dbo.Transaction_Details.ProductionDate,dbo.Transaction_Details.ExpiryDate,dbo.Transaction_Details.LotNO,  dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, dbo.TblStore.StoreName, "
StrSQL = StrSQL & "   dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ItemCase,"
StrSQL = StrSQL & "   dbo.Transaction_Details.Price, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.Transaction_Details.ColorID,"
StrSQL = StrSQL & "   dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & "   dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblUnites.UnitName, dbo.TblItems.SallingPrice, dbo.TblItems.PartNo, dbo.Transaction_Details.ShowQty,"
StrSQL = StrSQL & "   dbo.Transaction_Details.showPrice, dbo.Transactions.NoteSerial1, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
StrSQL = StrSQL & "   dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.CusID, dbo.TblCustemers.CusName,"
StrSQL = StrSQL & "   dbo.TblCustemers.CusNamee, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStore.StoreNamee,"
StrSQL = StrSQL & "   dbo.TblItems.ItemNamee, dbo.TblUnites.UnitNamee, dbo.Transactions.ManualNO, dbo.Transactions.PolicyNo, dbo.Transactions.ReciveOrderO,"
StrSQL = StrSQL & "   dbo.Transactions.DriverId, dbo.TblEmployee.Emp_Name AS drivername, dbo.TblEmployee.Emp_Namee AS drivernameen, dbo.Transactions.TransactionComment,"
StrSQL = StrSQL & "   dbo.TblStore.Code AS storecode, dbo.TblCustemers.Fullcode AS SupplierCode, dbo.Transactions.order_no, dbo.Transactions.Nots2 AS BillNo"
StrSQL = StrSQL & "   FROM         dbo.TblStore INNER JOIN"
StrSQL = StrSQL & "   dbo.Transactions ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblItems INNER JOIN"
StrSQL = StrSQL & "   dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
StrSQL = StrSQL & "   dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblEmployee ON dbo.Transactions.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_ID = " & BuyingID & ")"
    StrSQL = StrSQL + "order by dbo.Transaction_Details.id"



StrSQL = "SELECT     dbo.Transaction_Details.ProductionDate,dbo.Transaction_Details.ExpiryDate,dbo.Transaction_Details.LotNO,  dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, dbo.TblStore.StoreName, "
     StrSQL = StrSQL & "               dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ItemCase,"
     StrSQL = StrSQL & "               dbo.Transaction_Details.Price, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.Transaction_Details.ColorID,"
     StrSQL = StrSQL & "             dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
     StrSQL = StrSQL & "               dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblUnites.UnitName, dbo.TblItems.SallingPrice, dbo.TblItems.PartNo, dbo.Transaction_Details.ShowQty,"
     StrSQL = StrSQL & "               dbo.Transaction_Details.showPrice, dbo.Transactions.NoteSerial1, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
     StrSQL = StrSQL & "                dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.CusID, dbo.TblCustemers.CusName,"
     StrSQL = StrSQL & "                dbo.TblCustemers.CusNamee, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStore.StoreNamee,"
     StrSQL = StrSQL & "                 dbo.TblItems.ItemNamee, dbo.TblUnites.UnitNamee, dbo.Transactions.ManualNO, dbo.Transactions.PolicyNo, dbo.Transactions.ReciveOrderO,"
     StrSQL = StrSQL & "                 dbo.Transactions.DriverId, TblEmployee_1.Emp_Name AS drivername, TblEmployee_1.Emp_Namee AS drivernameen, dbo.Transactions.TransactionComment,"
     StrSQL = StrSQL & "                 dbo.TblStore.Code AS storecode, dbo.TblCustemers.Fullcode AS SupplierCode, dbo.Transactions.order_no, dbo.Transactions.Nots2 AS BillNo,"
     StrSQL = StrSQL & "                 dbo.Transactions.Emp_id , TblEmployee_1.emp_name, TblEmployee_1.fullcode, TblEmployee_1.Emp_Namee"
     StrSQL = StrSQL & "    FROM         dbo.TblItems INNER JOIN"
     StrSQL = StrSQL & "                 dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
     StrSQL = StrSQL & "                 dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblEmployee TblEmployee_1 ON dbo.Transactions.Emp_ID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblEmployee TblEmployee_2 ON dbo.Transactions.DriverId = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"

    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_ID = " & BuyingID & ")"
    StrSQL = StrSQL + "order by dbo.Transaction_Details.id"
 
 
 
 
 
 
 
 
 

StrSQL = " SELECT     dbo.Transaction_Details.ProductionDate, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.LotNO, dbo.Transactions.Transaction_ID, "
StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ItemSerial,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ItemCase, dbo.Transaction_Details.Price, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
StrSQL = StrSQL & "                      dbo.TblItems.GroupID, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId,"
StrSQL = StrSQL & "                      dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblUnites.UnitName, dbo.TblItems.SallingPrice,"
StrSQL = StrSQL & "                      dbo.TblItems.PartNo, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transactions.NoteSerial1, dbo.Transactions.Trans_Discount,"
StrSQL = StrSQL & "                      dbo.Transactions.Trans_DiscountType, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.CusID,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
StrSQL = StrSQL & "                      dbo.TblStore.StoreNamee, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitNamee, dbo.Transactions.ManualNO, dbo.Transactions.PolicyNo,"
StrSQL = StrSQL & "                      dbo.Transactions.ReciveOrderO, dbo.Transactions.DriverId, TblEmployee_1.Emp_Name AS drivername, TblEmployee_1.Emp_Namee AS drivernameen,"
StrSQL = StrSQL & "                      dbo.Transactions.TransactionComment, dbo.TblStore.Code AS StoreCode, dbo.TblCustemers.Fullcode AS SupplierCode, dbo.Transactions.order_no,"
StrSQL = StrSQL & "                      dbo.Transactions.Nots2 AS BillNo, dbo.Transactions.Emp_ID, TblEmployee_1.Emp_Name, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.CostPriceTk , dbo.Transactions.BillBasedOn, dbo.Transactions.CBoBasedON, dbo.Transactions.CashCustomerName"
StrSQL = StrSQL & " FROM         dbo.TblItems INNER JOIN"
StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
StrSQL = StrSQL & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.Transactions.Emp_ID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.Transactions.DriverId = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
StrSQL = StrSQL & "   Where (dbo.Transactions.Transaction_ID = " & BuyingID & ")"
StrSQL = StrSQL & "   ORDER BY dbo.Transaction_Details.ID"
'**************************************************************************
 
 
 '**************************************************************************
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(Reportname) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Reportname
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(Reportname)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        StrSQL = "select * From Notes where Transaction_ID=" & BuyingID
        StrSQL = StrSQL + " and NoteType=0"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
    
        End If

 xReport.ParameterFields(5).AddCurrentValue Title
  xReport.ParameterFields(6).AddCurrentValue tostorename

        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , Reportname
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function


Public Function ShowRecieveVoucherDatax15032017(BuyingID As Long, _
                                       Optional ShortDetails As Boolean = False _
                                       , Optional Title As String, Optional tostorename As String = "")
                                       
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Reportname  As String
    On Error GoTo ErrTrap

    If BuyingID = 0 Then
        Exit Function
    End If

    'StrSQL = "select * From QryBuyReport where Transaction_ID=" & BuyingID
    'StrSQL = StrSQL + "order by id"
    If SystemOptions.UserInterface = ArabicInterface Then
        If ShortDetails = False Then
           
                             
                             
         '   ReportName = App.path & "\Reports\RecieveVoucher2.rpt"
         Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\RecieveVoucher.rpt"
        Else
      '      ReportName = App.path & "\Reports\RecieveVoucher2.rpt"
      Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\RecieveVoucher.rpt"
        End If

    Else
 
        If ShortDetails = False Then
            'ReportName = App.path & "\Reports\RecieveVoucher2.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\RecieveVoucher.rpt"
        Else
            'ReportName = App.path & "\Reports\RecieveVoucher2.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\RecieveVoucher.rpt"
        End If
        
    End If

    StrSQL = "SELECT      dbo.Transaction_Details.ProductionDate,dbo.Transaction_Details.ExpiryDate,dbo.Transaction_Details.LotNO,  dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, dbo.TblStore.StoreName, "
StrSQL = StrSQL & "   dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ItemCase,"
StrSQL = StrSQL & "   dbo.Transaction_Details.Price, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.Transaction_Details.ColorID,"
StrSQL = StrSQL & "   dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & "   dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblUnites.UnitName, dbo.TblItems.SallingPrice, dbo.TblItems.PartNo, dbo.Transaction_Details.ShowQty,"
StrSQL = StrSQL & "   dbo.Transaction_Details.showPrice, dbo.Transactions.NoteSerial1, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
StrSQL = StrSQL & "   dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.CusID, dbo.TblCustemers.CusName,"
StrSQL = StrSQL & "   dbo.TblCustemers.CusNamee, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStore.StoreNamee,"
StrSQL = StrSQL & "   dbo.TblItems.ItemNamee, dbo.TblUnites.UnitNamee, dbo.Transactions.ManualNO, dbo.Transactions.PolicyNo, dbo.Transactions.ReciveOrderO,"
StrSQL = StrSQL & "   dbo.Transactions.DriverId, dbo.TblEmployee.Emp_Name AS drivername, dbo.TblEmployee.Emp_Namee AS drivernameen, dbo.Transactions.TransactionComment,"
StrSQL = StrSQL & "   dbo.TblStore.Code AS storecode, dbo.TblCustemers.Fullcode AS SupplierCode, dbo.Transactions.order_no, dbo.Transactions.Nots2 AS BillNo"
StrSQL = StrSQL & "   FROM         dbo.TblStore INNER JOIN"
StrSQL = StrSQL & "   dbo.Transactions ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblItems INNER JOIN"
StrSQL = StrSQL & "   dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
StrSQL = StrSQL & "   dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblEmployee ON dbo.Transactions.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_ID = " & BuyingID & ")"
    StrSQL = StrSQL + "order by dbo.Transaction_Details.id"



StrSQL = "SELECT     dbo.Transaction_Details.ProductionDate,dbo.Transaction_Details.ExpiryDate,dbo.Transaction_Details.LotNO,  dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, dbo.TblStore.StoreName, "
     StrSQL = StrSQL & "               dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ItemCase,"
     StrSQL = StrSQL & "               dbo.Transaction_Details.Price, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.Transaction_Details.ColorID,"
     StrSQL = StrSQL & "             dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
     StrSQL = StrSQL & "               dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblUnites.UnitName, dbo.TblItems.SallingPrice, dbo.TblItems.PartNo, dbo.Transaction_Details.ShowQty,"
     StrSQL = StrSQL & "               dbo.Transaction_Details.showPrice, dbo.Transactions.NoteSerial1, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
     StrSQL = StrSQL & "                dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.CusID, dbo.TblCustemers.CusName,"
     StrSQL = StrSQL & "                dbo.TblCustemers.CusNamee, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStore.StoreNamee,"
     StrSQL = StrSQL & "                 dbo.TblItems.ItemNamee, dbo.TblUnites.UnitNamee, dbo.Transactions.ManualNO, dbo.Transactions.PolicyNo, dbo.Transactions.ReciveOrderO,"
     StrSQL = StrSQL & "                 dbo.Transactions.DriverId, TblEmployee_1.Emp_Name AS drivername, TblEmployee_1.Emp_Namee AS drivernameen, dbo.Transactions.TransactionComment,"
     StrSQL = StrSQL & "                 dbo.TblStore.Code AS storecode, dbo.TblCustemers.Fullcode AS SupplierCode, dbo.Transactions.order_no, dbo.Transactions.Nots2 AS BillNo,"
     StrSQL = StrSQL & "                 dbo.Transactions.Emp_id , TblEmployee_1.emp_name, TblEmployee_1.fullcode, TblEmployee_1.Emp_Namee"
     StrSQL = StrSQL & "    FROM         dbo.TblItems INNER JOIN"
     StrSQL = StrSQL & "                 dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
     StrSQL = StrSQL & "                 dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblEmployee TblEmployee_1 ON dbo.Transactions.Emp_ID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblEmployee TblEmployee_2 ON dbo.Transactions.DriverId = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"

    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_ID = " & BuyingID & ")"
    StrSQL = StrSQL + "order by dbo.Transaction_Details.id"
 
 
 
 
 
 
 
 
 

StrSQL = " SELECT   dbo.Transaction_Details.ProductionDate,dbo.Transaction_Details.ExpiryDate,dbo.Transaction_Details.LotNO,     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, dbo.TblStore.StoreName,"
StrSQL = StrSQL & "                         dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ItemCase,"
StrSQL = StrSQL & "                         dbo.Transaction_Details.Price, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.Transaction_Details.ColorID,"
StrSQL = StrSQL & "                         dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & "                         dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblUnites.UnitName, dbo.TblItems.SallingPrice, dbo.TblItems.PartNo, dbo.Transaction_Details.ShowQty,"
StrSQL = StrSQL & "                         dbo.Transaction_Details.showPrice, dbo.Transactions.NoteSerial1, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
StrSQL = StrSQL & "                         dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.CusID, dbo.TblCustemers.CusName,"
StrSQL = StrSQL & "                         dbo.TblCustemers.CusNamee, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStore.StoreNamee,"
StrSQL = StrSQL & "                         dbo.TblItems.ItemNamee, dbo.TblUnites.UnitNamee, dbo.Transactions.ManualNO, dbo.Transactions.PolicyNo, dbo.Transactions.ReciveOrderO,"
StrSQL = StrSQL & "                         dbo.Transactions.DriverId, TblEmployee_1.Emp_Name AS drivername, TblEmployee_1.Emp_Namee AS drivernameen, dbo.Transactions.TransactionComment,"
StrSQL = StrSQL & "                         dbo.TblStore.Code AS storecode, dbo.TblCustemers.Fullcode AS SupplierCode, dbo.Transactions.order_no, dbo.Transactions.Nots2 AS BillNo,"
StrSQL = StrSQL & "                         dbo.Transactions.Emp_id , TblEmployee_1.emp_Name, TblEmployee_1.fullcode, TblEmployee_1.Emp_Namee, dbo.Transaction_Details.CostPriceTk"
StrSQL = StrSQL & "   FROM         dbo.TblItems INNER JOIN"
StrSQL = StrSQL & "                         dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
StrSQL = StrSQL & "                         dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblEmployee TblEmployee_1 ON dbo.Transactions.Emp_ID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblEmployee TblEmployee_2 ON dbo.Transactions.DriverId = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                         dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
StrSQL = StrSQL & "   Where (dbo.Transactions.Transaction_ID = " & BuyingID & ")"
StrSQL = StrSQL & "   ORDER BY dbo.Transaction_Details.ID"
'**************************************************************************
 
 
 '**************************************************************************
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(Reportname) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Reportname
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(Reportname)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        StrSQL = "select * From Notes where Transaction_ID=" & BuyingID
        StrSQL = StrSQL + " and NoteType=0"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
    
        End If

 xReport.ParameterFields(5).AddCurrentValue Title
  xReport.ParameterFields(6).AddCurrentValue tostorename

        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , Reportname
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function
   
Public Function ShowIssueVoucherData(BuyingID As Long, _
                                     Optional ShortDetails As Boolean = False, _
                                     Optional Title As String, _
                                     Optional tostorename As String = "", _
                                     Optional CarName As String = "", _
                                     Optional drivername As String = "")
    Dim xApp       As New CRAXDRT.Application
    Dim xReport    As New CRAXDRT.Report
    Dim StrSQL     As String
    Dim StrRS      As New ADODB.Recordset
    Dim RsTemp     As New ADODB.Recordset
    Dim rs         As New ADODB.Recordset
    Dim CViewer    As ClsReportViewer
    Dim Reportname As String
    On Error GoTo ErrTrap

    If BuyingID = 0 Then
        Exit Function
    End If

    'StrSQL = "select * From QryBuyReport where Transaction_ID=" & BuyingID
    'StrSQL = StrSQL + "order by id"
    If SystemOptions.UserInterface = ArabicInterface Then
        If ShortDetails = False Then
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\IssueVoucher.rpt"
            '       ReportName = App.path & "\Reports\IssueVoucher.rpt"
        Else
            '        ReportName = App.path & "\Reports\IssueVoucher.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\IssueVoucher.rpt"
        End If

    Else
 
        If ShortDetails = False Then
            'ReportName = App.path & "\Reports\IssueVoucher.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\IssueVoucher.rpt"
        Else
            'ReportName = App.path & "\Reports\IssueVoucher.rpt"
            Reportname = App.path & "\Special\" & SystemOptions.Reportpath & "\IssueVoucher.rpt"
        End If
        
    End If

    StrSQL = " SELECT     TOP 100 PERCENT dbo.Transaction_Details.ProductionDate, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.LotNO, "
    StrSQL = StrSQL & "                  dbo.Transactions.TransactionComment, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, dbo.TblStore.StoreName,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.Item_ID,CAST (  Transactions.InvoiceOrderNo AS NVARCHAR(10)) AS InvoiceOrderNo, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ItemCase,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.Price,Emp4.Emp_Name as TechName, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.Transaction_Details.ColorID,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
    StrSQL = StrSQL & "                  dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblUnites.UnitName, dbo.TblItems.SallingPrice, dbo.TblItems.PartNo, dbo.Transaction_Details.ShowQty,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.showPrice, dbo.Transactions.NoteSerial1, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.CusID, dbo.TblCustemers.CusName,"
    StrSQL = StrSQL & "                  dbo.TblCustemers.CusNamee, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStore.StoreNamee,"
    StrSQL = StrSQL & "                  dbo.TblItems.ItemNamee, dbo.TblUnites.UnitNamee, dbo.TblCustemers.Fullcode AS CusSupCode, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name,"
    StrSQL = StrSQL & "                  dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Fullcode AS EmpFullCode, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
    StrSQL = StrSQL & "                  dbo.Transactions.FixesAssetsID, dbo.FixedAssets.Name AS FaName, dbo.FixedAssets.Fullcode AS FaCode, dbo.FixedAssets.namee AS FaNameE,"
    StrSQL = StrSQL & "                  dbo.TblStore.Code AS StoreCode, dbo.Transactions.CashCustomerName, dbo.Transactions.Nots2, dbo.Transactions.order_no, dbo.Transactions.BillBasedOn,"
    StrSQL = StrSQL & "                  dbo.Transactions.ManualNo1, dbo.Transactions.ManualNo2, dbo.Transactions.ManualNO, dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.ISSUEDQTY, dbo.Transaction_Details.TotalPriceNoHours, dbo.Transaction_Details.TotalInvoiceQty,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.PriceNoHours, dbo.Transaction_Details.NoHours, dbo.Transaction_Details.Vat, dbo.Transaction_Details.Vatyo,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.IsExpirDate, dbo.Transaction_Details.L, dbo.Transaction_Details.W, dbo.Transaction_Details.H1, dbo.Transaction_Details.H2,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.OrderNo, dbo.Transaction_Details.StoreID2, TblStore_1.StoreName AS Expr1, TblStore_1.StoreNamee AS Expr2,"
    '''''''''''''''''''''''''''Ship
    StrSQL = StrSQL & "                  dbo.Transaction_Details.QtyFaqtors "
    StrSQL = StrSQL & "  ,  Transactions.ShipOrderNo, "
    StrSQL = StrSQL & "       Transactions.ShipEnquieryNo, "
    StrSQL = StrSQL & "       Transactions.ShipAccountNo, "
    StrSQL = StrSQL & "       Transactions.ShipCustomerName, "
    StrSQL = StrSQL & "       Transactions.ShipDistance, "
    StrSQL = StrSQL & "       Transactions.ShipArea, "
    StrSQL = StrSQL & "       Transactions.ShipSiteNo, "
    StrSQL = StrSQL & "       Transactions.ShipProjectName, "
    StrSQL = StrSQL & "       Transactions.ShipStructuralElement, "
    StrSQL = StrSQL & "       Transactions.ShipMixDescription, "
    StrSQL = StrSQL & "       Transactions.ShipDriverName, "
    StrSQL = StrSQL & "       Transactions.ShipPipeLine, "
    StrSQL = StrSQL & "       Transactions.ShipPump, "
    StrSQL = StrSQL & "       Transactions.ShipTruckNo, "
    StrSQL = StrSQL & "       Transactions.ShipIceTemp, "
    StrSQL = StrSQL & "       Transactions.ShipTotalDeleveryd, "
    StrSQL = StrSQL & "       Transactions.ShipThisLoad, "
    StrSQL = StrSQL & "       Transactions.ShipDayOrder, "
    StrSQL = StrSQL & "       Transactions.ShipTripNo, "
    StrSQL = StrSQL & "       Transactions.ShipPlantNo, "
    StrSQL = StrSQL & "       Transactions.ShipBatched, "
    StrSQL = StrSQL & "       Transactions.ShipRestunedPlant, "
    StrSQL = StrSQL & "       Transactions.ShipEndDischarge, "
    StrSQL = StrSQL & "       Transactions.ShipStartDisCharge, "
    StrSQL = StrSQL & "       Transactions.ShipOnSite,Transaction_Details.Remarks"
    '*************************************
    StrSQL = StrSQL & "   FROM         dbo.TblStore INNER JOIN"
    StrSQL = StrSQL & "                  dbo.Transactions ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItems INNER JOIN"
    StrSQL = StrSQL & "                  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
    StrSQL = StrSQL & "                  dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblStore TblStore_1 ON dbo.Transaction_Details.StoreID2 = TblStore_1.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.FixedAssets ON dbo.Transactions.FixesAssetsID = dbo.FixedAssets.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmpDepartments ON dbo.Transactions.DepartementID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    
    StrSQL = StrSQL & "                  Left Outer join TblEmployee Emp4 On dbo.Transaction_Details.EmpId4 = Emp4.Emp_ID "
    'Where (dbo.Transactions.Transaction_ID = 12)
    'ORDER BY dbo.Transaction_Details.ID
    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_ID = " & BuyingID & ")"
    StrSQL = StrSQL + "order by dbo.Transaction_Details.id"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(Reportname) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Reportname
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(Reportname)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        StrSQL = "select * From Notes where Transaction_ID=" & BuyingID
        StrSQL = StrSQL + " and NoteType=0"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
    
        End If
        
        If Title <> "" Then
            xReport.ParameterFields(5).AddCurrentValue Title
 
        End If

        xReport.ParameterFields(6).AddCurrentValue tostorename
        xReport.ParameterFields(7).AddCurrentValue CarName
        xReport.ParameterFields(8).AddCurrentValue drivername

        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , Reportname
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Function ShowBuyData(BuyingID As Long, _
                            Optional RecieveVoucherOrBill As Integer, _
                            Optional ShortDetails As Boolean = False, _
                            Optional Result As String, _
                            Optional ManualNO As String, _
                            Optional CURRENCYtEXT As String, _
                            Optional BranchID As Integer, _
                            Optional CurrnceRate As Double, _
                            Optional CurrancyID As Integer)
    On Error Resume Next
    Dim xApp       As New CRAXDRT.Application
    Dim xReport    As New CRAXDRT.Report
    Dim StrSQL     As String
    Dim StrRS      As New ADODB.Recordset
    Dim RsTemp     As New ADODB.Recordset
    Dim rs         As New ADODB.Recordset
    Dim CViewer    As ClsReportViewer
    Dim Reportname As String

    'On Error GoTo ErrTrap
    If BuyingID = 0 Then
        Exit Function
    End If

    'StrSQL = "select * From QryBuyReport where Transaction_ID=" & BuyingID
    'StrSQL = StrSQL + "order by id"
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = App.path & "\Reports\sBuyReport.rpt"
    Else
        Reportname = App.path & "\Reports\sBuyReport.rpt"
    End If

    StrSQL = " SELECT     Transactions.order_no, dbo.Transaction_Details.Remarks, dbo.Transaction_Details.NProductionOrderNO, dbo.Transactions.remark, dbo.Transactions.WorkOrderNO, "
    StrSQL = StrSQL & "                    dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID,"
    StrSQL = StrSQL & "                    dbo.TblCustemers.CusName,Transactions.Doctype,TblDoCumentsTypes.Name DocTypeName, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.Price,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemPhoto, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName,"
    StrSQL = StrSQL & "                    dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.CashCustomerName,"
    StrSQL = StrSQL & "                    dbo.Transactions.NoteSerial1,dbo.Transactions.project_id ,projects.Project_name,dbo.Transaction_Details.TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount,"
    StrSQL = StrSQL & "                    dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType, dbo.Transactions.ManualNo1, dbo.Transactions.PaymentType, dbo.TblStore.StoreNamee,"
    StrSQL = StrSQL & "                    dbo.TblUnites.UnitNamee, dbo.Transactions.InspectionReport, dbo.Transactions.TransactionComment, dbo.Transactions.BranchId,"
    StrSQL = StrSQL & "                    dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStore.Code AS StoreCode, dbo.Transaction_Details.LotNO,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.ProductionDate, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, dbo.Transactions.VATNO,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width, dbo.Transaction_Details.L,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.W, dbo.Transaction_Details.H1, dbo.Transaction_Details.H2, TblEmployee_1.Emp_Name, dbo.Transactions.empID1,"
    StrSQL = StrSQL & "                    TblEmployee_1.Emp_Name AS SuperVisorName, TblEmployee_1.Fullcode AS SuperVisorFullCode, TblEmployee_1.Emp_Namee AS SuperVisorNameE,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.Vat AS VATDet, dbo.Transaction_Details.Vatyo, dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty, dbo.Transactions.VAT,"
    StrSQL = StrSQL & "                    dbo.Transactions.VATCustoms, dbo.Transactions.VATCustoms1, TblEmployee_2.Emp_Name AS Emp_Name2, TblEmployee_2.Emp_Namee AS Emp_Namee2,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.Remarks1, dbo.Transaction_Details.Remarks2, dbo.Transaction_Details.LSPH, dbo.Transaction_Details.LCYL,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.RAXIS, dbo.Transaction_Details.RCYL, dbo.Transaction_Details.RSPH, dbo.Transaction_Details.TypeVAT,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.IssueSerial, dbo.Transaction_Details.ReceiveSerial, dbo.Transaction_Details.PriceNoHours, dbo.Transaction_Details.TotalPriceNoHours,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.NoHours , dbo.Transactions.BLDate, dbo.Transactions.ResonVAT, dbo.Transactions.Transaction_NetValue , dbo.Transactions.DueDate"
    StrSQL = StrSQL & "      ,ISNULL(TblItems.ItemWithOutVAT,0) ItemWithOutVAT ,IsNull(Transaction_Details.LineID,Transaction_Details.Id) as LineID ,Transactions.TypeTrans FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "                    dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_1 ON dbo.Transactions.empID1 = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_2 ON dbo.Transactions.Emp_ID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "    Left OUTER JOIn TblDoCumentsTypes On  TblDoCumentsTypes.Id = Transactions.Doctype and TblDoCumentsTypes.DocType = 22"
    StrSQL = StrSQL & "    Left OUTER JOIn projects On  projects.Id = dbo.Transactions.project_id"
    'new 21 10 2 019
 
    StrSQL = "   SELECT  dbo.Transactions.Currency_id, dbo.Transactions.Currency_rate,    dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE, dbo.Transactions.order_no, dbo.Transaction_Details.Remarks, dbo.Transaction_Details.NProductionOrderNO, dbo.Transactions.remark,"
    StrSQL = StrSQL & "                        dbo.Transactions.WorkOrderNO, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID,"
    StrSQL = StrSQL & "                        dbo.Transactions.StoreID, dbo.TblCustemers.CusName, dbo.Transactions.Doctype, dbo.TblDoCumentsTypes.name AS DoctypeName, dbo.TblCustemers.CusNamee,"
    StrSQL = StrSQL & "                        dbo.TblStore.StoreName, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID, dbo.Transaction_Details.Item_ID,"
    StrSQL = StrSQL & "                        dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.Price, dbo.Transaction_Details.Quantity,"
    StrSQL = StrSQL & "                        dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode,"
    StrSQL = StrSQL & "                        dbo.TblItems.ItemName, dbo.TblItems.ItemPhoto, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
    StrSQL = StrSQL & "                        dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.CashCustomerName, dbo.Transactions.NoteSerial1, dbo.Transactions.project_id,"
    StrSQL = StrSQL & "                        dbo.projects.Project_name, dbo.Transaction_Details.TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount,"
    StrSQL = StrSQL & "                        dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType, dbo.Transactions.ManualNo1, dbo.Transactions.PaymentType, dbo.TblStore.StoreNamee,"
    StrSQL = StrSQL & "                        dbo.TblUnites.UnitNamee, dbo.Transactions.InspectionReport, dbo.Transactions.TransactionComment, dbo.Transactions.BranchId,"
    StrSQL = StrSQL & "                        dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStore.Code AS StoreCode, dbo.Transaction_Details.LotNO,"
    StrSQL = StrSQL & "                        dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.ProductionDate, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, dbo.Transactions.VATNO,"
    StrSQL = StrSQL & "                        dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width, dbo.Transaction_Details.L,"
    StrSQL = StrSQL & "                        dbo.Transaction_Details.W, dbo.Transaction_Details.H1, dbo.Transaction_Details.H2, TblEmployee_1.Emp_Name, dbo.Transactions.empID1,"
    StrSQL = StrSQL & "                        TblEmployee_1.Emp_Name AS SuperVisorName, TblEmployee_1.Fullcode AS SuperVisorFullCode, TblEmployee_1.Emp_Namee AS SuperVisorNameE,"
    StrSQL = StrSQL & "                        dbo.Transaction_Details.Vat AS VatDet, dbo.Transaction_Details.Vatyo, dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty, dbo.Transactions.VAT,"
    StrSQL = StrSQL & "                        dbo.Transactions.VATCustoms, dbo.Transactions.VATCustoms1, TblEmployee_2.Emp_Name AS Emp_Name2, TblEmployee_2.Emp_Namee AS Emp_Namee2,"
    StrSQL = StrSQL & "                        dbo.Transaction_Details.Remarks1, dbo.Transaction_Details.Remarks2, dbo.Transaction_Details.LSPH, dbo.Transaction_Details.LCYL,"
    StrSQL = StrSQL & "                        dbo.Transaction_Details.RAXIS, dbo.Transaction_Details.RCYL, dbo.Transaction_Details.RSPH, dbo.Transaction_Details.TypeVAT,"
    StrSQL = StrSQL & "                        dbo.Transaction_Details.IssueSerial, dbo.Transaction_Details.ReceiveSerial, dbo.Transaction_Details.PriceNoHours, dbo.Transaction_Details.TotalPriceNoHours,"
    StrSQL = StrSQL & "                        dbo.Transaction_Details.NoHours , dbo.transactions.BLDate, dbo.transactions.ResonVAT, dbo.transactions.Transaction_NetValue, dbo.transactions.DueDate"
    StrSQL = StrSQL & "   ,ISNULL(TblItems.ItemWithOutVAT,0) ItemWithOutVAT ,"
    
    StrSQL = StrSQL & "   dbo.TblCustemers.Fullcode CustFullCode, ACCOUNTS_1.Account_Code, ACCOUNTS_1.Account_Name, ACCOUNTS_1.Account_NameEng, dbo.ACCOUNTS.Account_Code AS Account_Code2, dbo.ACCOUNTS.Account_Name AS Account_Name2,"
    StrSQL = StrSQL & "                            dbo.ACCOUNTS.Account_NameEng AS Account_NameEng2 ,pp2.Project_name Project_name2 ,pp2.Project_nameE Project_nameE2,pp2.Fullcode ProjectCode2,IsNull(Transaction_Details.LineID,Transaction_Details.Id) as LineID,Transactions.TypeTrans"
    StrSQL = StrSQL & "   FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "                        dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "                        dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblBoxesData ON dbo.Transactions.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblEmployee TblEmployee_1 ON dbo.Transactions.empID1 = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblEmployee TblEmployee_2 ON dbo.Transactions.Emp_ID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.TblDoCumentsTypes ON dbo.TblDoCumentsTypes.id = dbo.Transactions.Doctype AND dbo.TblDoCumentsTypes.Doctype = 22 LEFT OUTER JOIN"
    StrSQL = StrSQL & "                        dbo.projects ON dbo.projects.id = dbo.Transactions.project_id"
    StrSQL = StrSQL & "                        LEFT OUTER JOIN"
    StrSQL = StrSQL & "                                                 dbo.ACCOUNTS ON dbo.Transaction_Details.AccountCode = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
    StrSQL = StrSQL & "                                                 dbo.ACCOUNTS AS ACCOUNTS_1 ON dbo.Transaction_Details.Account_Code= ACCOUNTS_1.Account_Code "
    StrSQL = StrSQL & "                                                     LEFT OUTER JOIN dbo.projects pp2"
    StrSQL = StrSQL & "                                                 ON pp2.id = dbo.Transaction_Details.ProjectID"
    StrSQL = StrSQL & "                        "
    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_ID = " & BuyingID & ")"
   
    StrSQL = StrSQL + "order by dbo.Transaction_Details.id"

    If RecieveVoucherOrBill = 5 Then

        'StrSQL = "  SELECT    dbo.gettotalShippedQty1(Transactions.ProkerId)as totaRQty ,  Transactions.Transaction_ID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, Transactions_1.ProkerId, "
        'StrSQL = StrSQL & "   Transactions_1.Transaction_Type, dbo.Transaction_Details.ShipedQty, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.RecivedShippedQty,"
        'StrSQL = StrSQL & "   dbo.Transaction_Details.TotallRecivedShippedQty, Transactions_1.ProductionTypeid, dbo.TblProductionType.name AS productionname,"
        'StrSQL = StrSQL & "   dbo.TblProductionType.namee AS productionnamee, Transactions_1.Transaction_Date, Transactions_1.FixesAssetsID,"
        'StrSQL = StrSQL & "   Transactions_1.TransactionComment AS additions, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode AS ItemCode, dbo.TblItems.Dippre,"
        'StrSQL = StrSQL & "   dbo.TblItems.Content, dbo.TblItems.Wight, Transactions_1.Neighborhoodid, dbo.TblCountriesGovernmentsCities.CityName AS location, dbo.FixedAssets.Name,"
        'StrSQL = StrSQL & "   dbo.FixedAssets.Fullcode AS pumbcode, dbo.FixedAssets.namee AS punbname, Transactions_1.CarId, dbo.TblCarsData.BoardNO AS carCode,"
        'StrSQL = StrSQL & "   Transactions_1.empID1, TblEmployee_2.Emp_Name, TblEmployee_2.Fullcode AS empfullcode, TblEmployee_2.Emp_Namee, FixedAssets_1.Fullcode AS mixurecode,"
        'StrSQL = StrSQL & "   Transactions_1.ContactTime, Transactions_1.empID2, TblEmployee_2.Emp_Name AS supervisorname, TblEmployee_2.Fullcode AS supervisorCode,"
        'StrSQL = StrSQL & "   Transactions_1.NoteSerial1, Transactions_1.Transaction_Serial, TblEmployee_1.Emp_Code AS driverCode, TblEmployee_1.Emp_Name AS drivername,"
        'StrSQL = StrSQL & "   dbo.TblItems.Source , dbo.TblItems.Typenew"
        'StrSQL = StrSQL & "   FROM         dbo.TblEmployee TblEmployee_3 RIGHT OUTER JOIN"
        'StrSQL = StrSQL & "   dbo.TblCountriesGovernmentsCities RIGHT OUTER JOIN"
        'StrSQL = StrSQL & "   dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
        'StrSQL = StrSQL & "   dbo.Transactions Transactions_1 ON TblEmployee_2.Emp_ID = Transactions_1.empID2 ON"
        'StrSQL = StrSQL & "   dbo.TblCountriesGovernmentsCities.CityID = Transactions_1.Neighborhoodid ON TblEmployee_3.Emp_ID = Transactions_1.empID1 LEFT OUTER JOIN"
        'StrSQL = StrSQL & "   dbo.FixedAssets ON Transactions_1.FixesAssetsID = dbo.FixedAssets.id LEFT OUTER JOIN"
        'StrSQL = StrSQL & "   dbo.TblProductionType ON Transactions_1.ProductionTypeid = dbo.TblProductionType.Id RIGHT OUTER JOIN"
        'StrSQL = StrSQL & "   dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
        'StrSQL = StrSQL & "   dbo.FixedAssets FixedAssets_1 INNER JOIN"
        'StrSQL = StrSQL & "   dbo.TblCarsData ON FixedAssets_1.id = dbo.TblCarsData.fixedAssetid RIGHT OUTER JOIN"
        'StrSQL = StrSQL & "   dbo.Transactions Transactions INNER JOIN"
        'StrSQL = StrSQL & "   dbo.Transaction_Details ON Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
        'StrSQL = StrSQL & "   dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID ON dbo.TblCarsData.id = Transactions.CarId ON"
        'StrSQL = StrSQL & "   TblEmployee_1.Emp_ID = Transactions.DriverId ON Transactions_1.Transaction_ID = Transactions.ProkerId LEFT OUTER JOIN"
        'StrSQL = StrSQL & "   dbo.TblCustemers ON Transactions.CusID = dbo.TblCustemers.CusID"
 
        '   StrSQL = StrSQL & "   WHERE     (Transactions.Transaction_ID = " & BuyingID & ")"
        '    StrSQL = StrSQL + "order by dbo.Transaction_Details.id"

        'StrSQL = "SELECT     dbo.gettotalShippedQty1(Transactions.ProkerId) AS totaRQty, Transactions.Transaction_ID, dbo.TblCustemers.CusName, "
        'StrSQL = StrSQL & "   dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, Transactions_1.ProkerId, Transactions_1.Transaction_Type, dbo.Transaction_Details.ShipedQty,"
        'StrSQL = StrSQL & "   dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.RecivedShippedQty, dbo.Transaction_Details.TotallRecivedShippedQty, Transactions_1.ProductionTypeid,"
        'StrSQL = StrSQL & "  dbo.TblProductionType.name AS productionname, dbo.TblProductionType.namee AS productionnamee, Transactions_1.Transaction_Date,"
        'StrSQL = StrSQL & "  Transactions_1.FixesAssetsID, Transactions_1.TransactionComment AS additions, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
        'StrSQL = StrSQL & "  dbo.TblItems.Fullcode AS ItemCode, dbo.TblItems.Dippre, dbo.TblItems.Content, dbo.TblItems.Wight, Transactions_1.Neighborhoodid,"
        'StrSQL = StrSQL & "  dbo.TblCountriesGovernmentsCities.CityName AS location, dbo.FixedAssets.Name, dbo.FixedAssets.Fullcode AS pumbcode, dbo.FixedAssets.namee AS punbname,"
        'StrSQL = StrSQL & "  Transactions_1.CarId, dbo.TblCarsData.BoardNO AS carCode, Transactions_1.empID1, TblEmployee_2.Emp_Name, TblEmployee_2.Fullcode AS empfullcode,"
        'StrSQL = StrSQL & "  TblEmployee_2.Emp_Namee, FixedAssets_1.Fullcode AS mixurecode, Transactions_1.ContactTime, Transactions_1.empID2,"
        'StrSQL = StrSQL & "  TblEmployee_2.Emp_Name AS supervisorname, TblEmployee_2.Fullcode AS supervisorCode, Transactions_1.NoteSerial1, Transactions_1.Transaction_Serial,"
        'StrSQL = StrSQL & "  TblEmployee_1.Emp_Code AS driverCode, TblEmployee_1.Emp_Name AS drivername, dbo.TblItems.Source, dbo.TblItems.Typenew,"
        'StrSQL = StrSQL & "  TblEmployee_3.Emp_Name AS Workername, TblEmployee_3.Emp_Namee AS WorkernameEng"
        'StrSQL = StrSQL & "  FROM         dbo.TblEmployee TblEmployee_3 RIGHT OUTER JOIN"
        'StrSQL = StrSQL & "  dbo.TblCountriesGovernmentsCities RIGHT OUTER JOIN"
        'StrSQL = StrSQL & "  dbo.TblEmployee TblEmployee_2 INNER JOIN"
        'StrSQL = StrSQL & "  dbo.Transactions Transactions_1 ON TblEmployee_2.Emp_ID = Transactions_1.empID2 ON"
        'StrSQL = StrSQL & "  dbo.TblCountriesGovernmentsCities.CityID = Transactions_1.Neighborhoodid ON TblEmployee_3.Emp_ID = Transactions_1.empID1 LEFT OUTER JOIN"
        'StrSQL = StrSQL & "  dbo.FixedAssets ON Transactions_1.FixesAssetsID = dbo.FixedAssets.id LEFT OUTER JOIN"
        'StrSQL = StrSQL & "  dbo.TblProductionType ON Transactions_1.ProductionTypeid = dbo.TblProductionType.Id RIGHT OUTER JOIN"
        'StrSQL = StrSQL & "  dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
        'StrSQL = StrSQL & "  dbo.FixedAssets FixedAssets_1 INNER JOIN"
        'StrSQL = StrSQL & "  dbo.TblCarsData ON FixedAssets_1.id = dbo.TblCarsData.fixedAssetid RIGHT OUTER JOIN"
        'StrSQL = StrSQL & "     dbo.Transactions Transactions INNER JOIN"
        'StrSQL = StrSQL & "  dbo.Transaction_Details ON Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
        'StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID ON dbo.TblCarsData.id = Transactions.CarId ON"
        'StrSQL = StrSQL & "  TblEmployee_1.Emp_ID = Transactions.DriverId ON Transactions_1.Transaction_ID = Transactions.ProkerId LEFT OUTER JOIN"
        'StrSQL = StrSQL & "  dbo.TblCustemers ON Transactions.CusID = dbo.TblCustemers.CusID"
        'Where (Transactions.Transaction_ID = 1148)
        'ORDER BY dbo.Transaction_Details.ID"

        StrSQL = "SELECT     TOP 102 PERCENT dbo.gettotalShippedQty1(Transactions.ProkerId) AS totaRQty, Transactions.Transaction_ID, dbo.TblCustemers.CusName, "
        StrSQL = StrSQL & "   dbo.TblCustemers.CusNamee,TblDoCumentsTypes.Name DocTypeName, dbo.TblCustemers.Fullcode, Transactions_1.ProkerId, Transactions_1.Transaction_Type, dbo.Transaction_Details.ShipedQty,"
        StrSQL = StrSQL & "   dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.RecivedShippedQty, dbo.Transaction_Details.TotallRecivedShippedQty, Transactions_1.ProductionTypeid,"
        StrSQL = StrSQL & "   dbo.TblProductionType.name AS productionname,Transactions.Doctype, dbo.TblProductionType.namee AS productionnamee, Transactions_1.Transaction_Date,"
        StrSQL = StrSQL & "   Transactions_1.FixesAssetsID, Transactions_1.TransactionComment AS additions, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
        StrSQL = StrSQL & "   dbo.TblItems.Fullcode AS ItemCode,dbo.Transactions.project_id ,projects.Project_name, dbo.TblItems.Dippre, dbo.TblItems.Content, dbo.TblItems.Wight, Transactions_1.Neighborhoodid,"
        StrSQL = StrSQL & "   dbo.TblCountriesGovernmentsCities.CityName AS location, Transactions_1.CarId, dbo.TblCarsData.BoardNO, Transactions_1.empID1, TblEmployee_2.Emp_Name,"
        StrSQL = StrSQL & "   TblEmployee_2.Fullcode AS empfullcode, TblEmployee_2.Emp_Namee, Transactions_1.empID2, TblEmployee_2.Emp_Name AS supervisorname,"
        StrSQL = StrSQL & "                         TblEmployee_2.Fullcode AS supervisorCode, TblEmployee_1.Emp_Code AS driverCode, TblEmployee_1.Emp_Name AS drivername, dbo.TblItems.Source,"
        StrSQL = StrSQL & "   dbo.TblItems.Typenew, TblEmployee_3.Emp_Name AS Workername, TblEmployee_3.Emp_Namee AS WorkernameEng, dbo.TblCarsData.Notes AS carecode2,"
        StrSQL = StrSQL & "   Transactions.NoteSerial1, Transactions.ContactTime, TblEmployee_4.Emp_Name AS sabasupervisor, dbo.FixedAssets.id,"
        StrSQL = StrSQL & "   dbo.TblEquipments.Notes AS pubpcodenew"
        StrSQL = StrSQL & "   ,ISNULL(TblItems.ItemWithOutVAT,0) ItemWithOutVAT FROM         dbo.TblProductionType RIGHT OUTER JOIN"
        StrSQL = StrSQL & "   dbo.TblEquipments RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                         dbo.FixedAssets ON dbo.TblEquipments.fixedAssetid = dbo.FixedAssets.id RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                         dbo.TblEmployee TblEmployee_3 RIGHT OUTER JOIN"
        StrSQL = StrSQL & "   dbo.TblCountriesGovernmentsCities RIGHT OUTER JOIN"
        StrSQL = StrSQL & "   dbo.TblEmployee TblEmployee_2 INNER JOIN"
        StrSQL = StrSQL & "   dbo.Transactions Transactions_1 ON TblEmployee_2.Emp_ID = Transactions_1.empID2 ON"
        StrSQL = StrSQL & "   dbo.TblCountriesGovernmentsCities.CityID = Transactions_1.Neighborhoodid ON TblEmployee_3.Emp_ID = Transactions_1.empID1 ON"
        StrSQL = StrSQL & "   dbo.FixedAssets.id = Transactions_1.FixesAssetsID ON dbo.TblProductionType.Id = Transactions_1.ProductionTypeid RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                         dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
        StrSQL = StrSQL & "   dbo.TblCarsData RIGHT OUTER JOIN"
        StrSQL = StrSQL & "   dbo.TblEmployee TblEmployee_4 INNER JOIN"
        StrSQL = StrSQL & "   dbo.Transactions Transactions INNER JOIN"
        StrSQL = StrSQL & "   dbo.Transaction_Details ON Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
        StrSQL = StrSQL & "   dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID ON TblEmployee_4.Emp_ID = Transactions.empID1 ON"
        StrSQL = StrSQL & "   dbo.TblCarsData.id = Transactions.CarId ON TblEmployee_1.Emp_ID = Transactions.DriverId ON"
        StrSQL = StrSQL & "   Transactions_1.Transaction_ID = Transactions.ProkerId LEFT OUTER JOIN"
        StrSQL = StrSQL & "   dbo.TblCustemers ON Transactions.CusID = dbo.TblCustemers.CusID"
   
        StrSQL = StrSQL & "    Left OUTER JOIn TblDoCumentsTypes On  TblDoCumentsTypes.Id = Transactions.Doctype and TblDoCumentsTypes.DocType = 22"
        StrSQL = StrSQL & "    Left OUTER JOIn projects On  projects.Id = dbo.Transactions.project_id"

        StrSQL = "SELECT     TOP 100 PERCENT dbo.gettotalShippedQty1(Transactions.ProkerId) AS totaRQty, Transactions.Transaction_ID, dbo.TblCustemers.CusName, "
        StrSQL = StrSQL & "                        dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, Transactions_1.ProkerId, Transactions_1.Transaction_Type, dbo.Transaction_Details.ShipedQty,"
        StrSQL = StrSQL & "  dbo.Transaction_Details.ShowQty,TblDoCumentsTypes.Name DocTypeName,  dbo.Transaction_Details.RecivedShippedQty, dbo.Transaction_Details.TotallRecivedShippedQty, Transactions_1.ProductionTypeid,"
        StrSQL = StrSQL & "  dbo.TblProductionType.name AS productionname,Transactions.Doctype, dbo.TblProductionType.namee AS productionnamee, Transactions_1.Transaction_Date,"
        StrSQL = StrSQL & "  Transactions_1.FixesAssetsID, Transactions_1.TransactionComment AS additions, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
        StrSQL = StrSQL & "  dbo.TblItems.Fullcode AS ItemCode ,dbo.Transactions.project_id ,projects.Project_name, dbo.TblItems.Dippre, dbo.TblItems.Content, dbo.TblItems.Wight, Transactions_1.Neighborhoodid,"
        StrSQL = StrSQL & "  dbo.TblCountriesGovernmentsCities.CityName AS location, Transactions_1.CarId, dbo.TblCarsData.BoardNO, Transactions_1.empID1, TblEmployee_2.Emp_Name,"
        StrSQL = StrSQL & "  TblEmployee_2.Fullcode AS empfullcode, TblEmployee_2.Emp_Namee, Transactions_1.empID2, TblEmployee_2.Emp_Name AS supervisorname,"
        StrSQL = StrSQL & "  TblEmployee_2.Fullcode AS supervisorCode, TblEmployee_1.Emp_Code AS driverCode, TblEmployee_1.Emp_Name AS drivername, dbo.TblItems.Source,"
        StrSQL = StrSQL & "  dbo.TblItems.Typenew, TblEmployee_3.Emp_Name AS Workername, TblEmployee_3.Emp_Namee AS WorkernameEng, dbo.TblCarsData.Notes AS carecode2,"
        StrSQL = StrSQL & "  Transactions.NoteSerial1, Transactions.ContactTime, TblEmployee_4.Emp_Name AS sabasupervisor, dbo.FixedAssets.id,"
        StrSQL = StrSQL & "  dbo.TblEquipments.Notes AS pubpcodenew, dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate , dbo.Transaction_Details.ProductionDate"
        StrSQL = StrSQL & "  ,ISNULL(TblItems.ItemWithOutVAT,0) ItemWithOutVAT FROM         dbo.TblProductionType RIGHT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblEquipments RIGHT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.FixedAssets ON dbo.TblEquipments.fixedAssetid = dbo.FixedAssets.id RIGHT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblEmployee TblEmployee_3 RIGHT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblCountriesGovernmentsCities RIGHT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
        StrSQL = StrSQL & " dbo.Transactions Transactions_1 ON TblEmployee_2.Emp_ID = Transactions_1.empID2 ON"
        StrSQL = StrSQL & " dbo.TblCountriesGovernmentsCities.CityID = Transactions_1.Neighborhoodid ON TblEmployee_3.Emp_ID = Transactions_1.empID1 ON"
        StrSQL = StrSQL & " dbo.FixedAssets.id = Transactions_1.FixesAssetsID ON dbo.TblProductionType.Id = Transactions_1.ProductionTypeid RIGHT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblCarsData RIGHT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblEmployee TblEmployee_4 RIGHT OUTER JOIN"
        StrSQL = StrSQL & " dbo.Transactions Transactions INNER JOIN"
        StrSQL = StrSQL & " dbo.Transaction_Details ON Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID ON TblEmployee_4.Emp_ID = Transactions.empID1 ON"
        StrSQL = StrSQL & " dbo.TblCarsData.id = Transactions.CarId ON TblEmployee_1.Emp_ID = Transactions.DriverId ON"
        StrSQL = StrSQL & " Transactions_1.Transaction_ID = Transactions.ProkerId LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblCustemers ON Transactions.CusID = dbo.TblCustemers.CusID"
  
        StrSQL = StrSQL & "    Left OUTER JOIn TblDoCumentsTypes On  TblDoCumentsTypes.Id = Transactions.Doctype and TblDoCumentsTypes.DocType = 22"
        StrSQL = StrSQL & "    Left OUTER JOIn projects On  projects.Id = dbo.Transactions.project_id"
  
        'Where (Transactions.Transaction_ID = 1272)
        'ORDER BY dbo.Transaction_Details.ID"
 
        StrSQL = StrSQL & "   WHERE     (Transactions.Transaction_ID = " & BuyingID & ")"
        StrSQL = StrSQL + " order by dbo.Transaction_Details.id"

    End If

    If RecieveVoucherOrBill = 6 Then
        StrSQL = "SELECT     TOP 100 PERCENT dbo.Transaction_Details.NProductionOrderNO,TblDoCumentsTypes.Name DocTypeName,  dbo.Transactions.remark, dbo.Transactions.WorkOrderNO, dbo.Transactions.Transaction_Serial, "
        StrSQL = StrSQL & "  dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID, dbo.TblCustemers.CusName,"
        StrSQL = StrSQL & "  dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName,Transactions.Doctype, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID,"
        StrSQL = StrSQL & "  dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.Price,"
        StrSQL = StrSQL & "   dbo.Transaction_Details.Quantity,dbo.Transactions.project_id ,projects.Project_name, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId,"
        StrSQL = StrSQL & "  dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemPhoto, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName,"
        StrSQL = StrSQL & "  dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.CashCustomerName,"
        StrSQL = StrSQL & "  dbo.Transactions.NoteSerial1, dbo.Transaction_Details.TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount,"
        StrSQL = StrSQL & "  dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType, dbo.Transactions.ManualNo1, dbo.Transactions.PaymentType, dbo.TblStore.StoreNamee,"
        StrSQL = StrSQL & "  dbo.TblUnites.UnitNamee, dbo.Transactions.InspectionReport, dbo.Transactions.TransactionComment, dbo.Transactions.BranchId,"
        StrSQL = StrSQL & "  dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStore.Code AS StoreCode, dbo.TblItems.Typenew, dbo.TblItems.Source,"
        StrSQL = StrSQL & "  dbo.TblItems.Dippre, dbo.TblItems.Content, dbo.TblItems.Wight, dbo.TblEmployee.Emp_Code AS drivercode, dbo.TblEmployee.Emp_Name AS drivername,"
        StrSQL = StrSQL & "  dbo.Transactions.Neighborhoodid , dbo.TblCountriesGovernmentsCities.CityName , dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate , dbo.Transaction_Details.ProductionDate"
        StrSQL = StrSQL & "  ,ISNULL(TblItems.ItemWithOutVAT,0) ItemWithOutVAT  ,Transactions.TypeTrans FROM         dbo.Transactions INNER JOIN"
        StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
        StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblCountriesGovernmentsCities ON dbo.Transactions.Neighborhoodid = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.Transactions.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
        StrSQL = StrSQL & "    Left OUTER JOIn TblDoCumentsTypes On  TblDoCumentsTypes.Id = Transactions.Doctype and TblDoCumentsTypes.DocType = 22"
        StrSQL = StrSQL & "    Left OUTER JOIn projects On  projects.Id = dbo.Transactions.project_id"

        'Where (Transactions.Transaction_ID = 1272)
        'ORDER BY dbo.Transaction_Details.ID"
        'Where (dbo.Transactions.Transaction_ID = 1147)
        'ORDER BY dbo.Transaction_Details.ID"
        StrSQL = StrSQL & "   WHERE     (Transactions.Transaction_ID = " & BuyingID & ")"
        StrSQL = StrSQL + " order by dbo.Transaction_Details.id"

    End If
    
    Dim FileName As String

    If RecieveVoucherOrBill = 0 Then 'voucher
        '   xReport.ReportTitle = "«–‰ «” ·«„ »÷«⁄Â "
        If SystemOptions.UserInterface = ArabicInterface Then
            FileName = App.path & "\Reports\BuyReport.rpt"
        Else
            FileName = App.path & "\Reports\BuyReportE.rpt"
        End If
        
    ElseIf RecieveVoucherOrBill = 1 Then
        '   xReport.ReportTitle = "›« Ê—… „‘ —Ì«   "
        If SystemOptions.UserInterface = ArabicInterface Then
   
            FileName = App.path & "\Reports\BuyReport.rpt"
        Else
            FileName = App.path & "\Reports\BuyReportE.rpt"
        End If
        
    ElseIf RecieveVoucherOrBill = 2 Then
        '   xReport.ReportTitle = "«„— «‰ «Ã"
        If SystemOptions.UserInterface = ArabicInterface Then

            FileName = App.path & "\Reports\ProductionOrder.rpt"
        Else
            FileName = App.path & "\Reports\ProductionOrdere.rpt"
        End If

    ElseIf RecieveVoucherOrBill = 3 Then

        '   xReport.ReportTitle = "’—› „Ê«œ ··«‰ «Ã  "
        If SystemOptions.UserInterface = ArabicInterface Then
            FileName = App.path & "\Reports\ProductionIssueVoucherTotal.rpt"
        Else
            FileName = App.path & "\Reports\ProductionIssueVoucherTotal.rpt"
        End If
    ElseIf RecieveVoucherOrBill = 33 Then

        '   xReport.ReportTitle = "’—› „Ê«œ ··«‰ «Ã  "
        If SystemOptions.UserInterface = ArabicInterface Then
            FileName = App.path & "\Reports\ProductionIssueVoucher.rpt"
        Else
            FileName = App.path & "\Reports\ProductionIssueVouchere.rpt"
        End If
    ElseIf RecieveVoucherOrBill = 4 Then
        '   xReport.ReportTitle = "    «” ·«„ «‰ «Ã  «„  "
        FileName = App.path & "\Reports\ProductionRecieveVoucher.rpt"
                          
    ElseIf RecieveVoucherOrBill = 5 Then
        '   xReport.ReportTitle = "    «” ·«„  ’—› ··⁄„Ì· «‰ «Ã   «„  "
        FileName = App.path & "\Reports\REPORTS NEW\productionDeliverynotes.rpt"
    ElseIf RecieveVoucherOrBill = 6 Then
        FileName = App.path & "\Reports\ProductionOrder12.rpt"
    End If
                        
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(FileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\BuyReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(FileName)
        xReport.Database.SetDataSource rs
        Set StrRS = New ADODB.Recordset
        StrRS.Open "TblOptions", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        If SystemOptions.VATNoAccordActivity = False Then
            If xReport.ParameterFields.count > 11 Then
                xReport.ParameterFields(12).AddCurrentValue IIf(IsNull(StrRS("VATRegNo").value), "", StrRS("VATRegNo").value)
            End If
        Else
            If xReport.ParameterFields.count > 11 Then
                xReport.ParameterFields(12).AddCurrentValue GetRegVATNo(BranchID)
            End If
        End If
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        If xReport.ParameterFields.count > 11 Then
            xReport.ParameterFields(13).AddCurrentValue CurrnceRate
        End If
        '    xReport.ParameterFields(8).AddCurrentValue WriteNo(CStr(result), 0, True)
        
        xReport.ParameterFields(8).AddCurrentValue WriteNo(Format(CStr(Result), "0.00"), 0, True, ".", , 0, , CurrancyID)
        xReport.ParameterFields(11).AddCurrentValue WriteNo(Format(CStr(Result * CurrnceRate), "0.00"), 0, True, ".", , 0)
        If RecieveVoucherOrBill = 1 Then
            xReport.ParameterFields(10).AddCurrentValue CURRENCYtEXT
        End If

        StrSQL1 = "select * From Notes where Transaction_ID=" & BuyingID
        StrSQL1 = StrSQL1 + " and NoteType=0"
        RsTemp.Open StrSQL1, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
             
            If RecieveVoucherOrBill = 0 Then 'voucher
                If SystemOptions.UserInterface = ArabicInterface Then
                        
                    xReport.reporttitle = "«–‰ «” ·«„ »÷«⁄Â "
                Else
                    xReport.reporttitle = "Recieve Voucher "
                End If
            ElseIf RecieveVoucherOrBill = 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then

                    xReport.reporttitle = "›« Ê—… „‘ —Ì«   "
                Else
                    xReport.reporttitle = "Purchase Invoice"
                End If
                xReport.ParameterFields(9).AddCurrentValue ManualNO
                
                If SystemOptions.PursgaseWithoutDecimal = True Then
                    xReport.ParameterFields(4).AddCurrentValue "1"
                End If
                'ManualNO
            ElseIf RecieveVoucherOrBill = 2 Then
                If SystemOptions.UserInterface = ArabicInterface Then

                    xReport.reporttitle = "«„— «‰ «Ã"
                Else
                    xReport.reporttitle = "Production order"
                End If
            ElseIf RecieveVoucherOrBill = 3 Then

                If SystemOptions.UserInterface = ArabicInterface Then
                    xReport.reporttitle = "  ”‰œ ’—› „Ê«œ Œ«„ ··«‰ «Ã"
                Else
                    xReport.reporttitle = "Production Issue Voucher"
                End If
                           
            ElseIf RecieveVoucherOrBill = 4 Then
                If SystemOptions.UserInterface = ArabicInterface Then

                    xReport.reporttitle = "  ”‰œ «” ·«„ «‰ «Ã  «„"
                Else
                    xReport.reporttitle = "  Production Recieve Voucher"
                End If
            End If
    
        End If

        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , FileName, , StrSQL
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function



 
  Public Function ShowBuyDataShort(BuyingID As Long, Optional BranchID As Integer)
 On Error Resume Next
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim cOptions As ClsCompanyInfo
   ' On Error GoTo ErrTrap

    If BuyingID = 0 Then
        Exit Function
    End If

    StrSQL = "select * From QryBuyReportShort where Transaction_ID=" & BuyingID
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(App.path & "\Reports\BuyReportShort.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\BuyReportShort.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\BuyReportShort.rpt")
        xReport.Database.SetDataSource rs
        xReport.reporttitle = "ÿ»«⁄… ›« Ê—… ‘—«¡"
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        StrSQL1 = "select * From Notes where Transaction_ID=" & BuyingID
        StrSQL1 = StrSQL1 + " and NoteType=0"
        RsTemp.Open StrSQL1, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
        End If

        Dim SQLExp As String
        SQLExp = "select * from QryInst where transaction_id=" & BuyingID
        Dim RsExp As New ADODB.Recordset
        RsExp.Open SQLExp, Cn, adOpenStatic, adLockPessimistic, adCmdText

        If RsExp.RecordCount > 0 Then
            xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(Format(RsExp("BasicAmmount").value, "00.00")), 0, Format(RsExp("BasicAmmount").value, "00.00"))
        Else
            xReport.ParameterFields(5).AddCurrentValue CStr(0)
        End If

        RsExp.Close

        SQLExp = "select * from SumCheck where transaction_id=" & BuyingID
        RsExp.Open SQLExp, Cn, adOpenStatic, adLockPessimistic, adCmdText

        If RsExp.RecordCount > 0 Then
            xReport.ParameterFields(6).AddCurrentValue IIf(IsNull(Format(RsExp("sumcheck").value, "00.00")), 0, Format(RsExp("sumcheck").value, "00.00"))
        Else
            xReport.ParameterFields(6).AddCurrentValue CStr(0)
        End If

        RsExp.Close

        SQLExp = "select * from credit where transaction_id=" & BuyingID
        RsExp.Open SQLExp, Cn, adOpenStatic, adLockPessimistic, adCmdText

        If RsExp.RecordCount > 0 Then
            xReport.ParameterFields(7).AddCurrentValue IIf(IsNull(Format(RsExp("Note_Value").value, "00.00")), 0, Format(RsExp("Note_Value").value, "00.00"))
        Else
            xReport.ParameterFields(7).AddCurrentValue CStr(0)
        End If

        RsExp.Close

        Screen.MousePointer = vbDefault
    End If
     If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(12).AddCurrentValue cOptions.VATRegNo
    Else
    xReport.ParameterFields(12).AddCurrentValue GetRegVATNo(BranchID)
    End If
    '    Screen.MousePointer = vbArrowHourglass
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , App.path & "\Reports\BuyReportShort.rpt", , StrSQL
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function


