VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsOpeningBalanceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function ShowOpeningBalanceData(OpeningID As Double, _
                                       Optional ShortView As Boolean, Optional reprtno As Integer = 0)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrTrap

    If OpeningID = Null Then
        Exit Function
    End If

    'StrSQL = "select * From QryBalanceOpening where Transaction_ID=" & OpeningID

    StrSQL = "SELECT  dbo.Transaction_Details.ProductionDate,dbo.Transaction_Details.ExpiryDate,dbo.Transaction_Details.LotNO,    dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.Transaction_Details.Item_ID, "
    StrSQL = StrSQL & "  dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ItemCase, dbo.Transaction_Details.Price, dbo.TblItems.ItemCode,"
    StrSQL = StrSQL & "   dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId,"
    StrSQL = StrSQL & "   dbo.Transaction_Details.ClassId, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName,"
    StrSQL = StrSQL & "  dbo.TblUnites.unitname, dbo.TblItems.SallingPrice, dbo.TblItems.PartNo, dbo.Transaction_Details.ShowQty,dbo.Transaction_Details.ShowPrice,  dbo.TblStore.StoreNamee ,  dbo.TblItems.ItemNamee , dbo.TblUnites.UnitNamee"
    StrSQL = StrSQL & "  FROM         dbo.TblStore INNER JOIN"
    StrSQL = StrSQL & "   dbo.Transactions ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblItems INNER JOIN"
    StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
    StrSQL = StrSQL & "  dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "   dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "   dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_ID = " & OpeningID & ")"
    StrSQL = StrSQL + "order by id"

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
 
    Dim Reportname As String
 
    If SystemOptions.UserInterface = ArabicInterface Then
        If ShortView = False Then
            Reportname = (App.path & "\Reports\OpeningPalanceReport.rpt")
            
        Else
            Reportname = (App.path & "\Reports\OpeningPalanceReportShort.rpt")
        End If
 
    Else
 
        If ShortView = False Then
            Reportname = (App.path & "\Reports\OpeningPalanceReportENG.rpt")
            
        Else
            Reportname = (App.path & "\Reports\OpeningPalanceReportShortENG.rpt")
        End If
 
    End If
 
 
If reprtno = 1 Then
  Reportname = (App.path & "\Reports\REPORTS NEW\gardEntryReports.rpt")
End If

    If Reportname = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\OpeningPalanceReport.rpt"
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
        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    CreateLogo xReport
    FrmReport.show
    FrmReport.txtPath = Reportname
      FrmReport.TXTSTRSQL = StrSQL
    
    Screen.MousePointer = vbDefault
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Function ShowOpeningBalanceDataAuditing(OpeningID As Integer, _
                                               Optional ShortView As Boolean, Optional StoreID As Integer = 0)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrTrap

    If OpeningID = Null Then
        Exit Function
    End If

    'StrSQL = "select * From QryBalanceOpening where Transaction_ID=" & OpeningID
    OpeningID = 3
    StrSQL = "SELECT    dbo.TblItems.barcodeno, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.Transaction_Details.Item_ID, "
    StrSQL = StrSQL & "  dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ItemCase, dbo.Transaction_Details.Price, dbo.TblItems.ItemCode,"
    StrSQL = StrSQL & "   dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId,"
    StrSQL = StrSQL & "   dbo.Transaction_Details.ClassId, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName,"
    StrSQL = StrSQL & "  dbo.TblUnites.unitname, dbo.TblItems.SallingPrice, dbo.TblItems.PartNo, dbo.Transaction_Details.ShowQty,dbo.Transaction_Details.ShowPrice,  dbo.TblStore.StoreNamee ,  dbo.TblItems.ItemNamee , dbo.TblUnites.UnitNamee"
    StrSQL = StrSQL & "  FROM         dbo.TblStore INNER JOIN"
    StrSQL = StrSQL & "   dbo.Transactions ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblItems INNER JOIN"
    StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
    StrSQL = StrSQL & "  dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "   dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "   dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.Transaction_Type = " & OpeningID & ")"
    

If StoreID <> 0 Then
StrSQL = StrSQL & " and dbo.Transactions.StoreId =" & StoreID
End If
   
   StrSQL = StrSQL + "order by id"
   
    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
 
    Dim Reportname As String
 
    If SystemOptions.UserInterface = ArabicInterface Then
        If ShortView = False Then
            Reportname = (App.path & "\Reports\OpeningPalanceReportAudit.rpt")
            
        Else
            Reportname = (App.path & "\Reports\OpeningPalanceReportAudit.rpt")
        End If
 
    Else
 
        If ShortView = False Then
            Reportname = (App.path & "\Reports\OpeningPalanceReportENG.rpt")
            
        Else
            Reportname = (App.path & "\Reports\OpeningPalanceReportAuditE.rpt")
        End If
 
    End If
 
    If Reportname = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\OpeningPalanceReport.rpt"
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
        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    
    FrmReport.show
    FrmReport.txtPath = Reportname
  FrmReport.TXTSTRSQL = StrSQL
    Screen.MousePointer = vbDefault
    
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function
Public Function ShowFixedAssets(OpeningID As Integer, _
                                Optional BranchView As Boolean = True, _
                                Optional branch_no As Integer = 0, _
                                Optional group_id As Integer = 0, _
                                Optional Emp_id As Integer = 0, _
                                Optional FixedID As Integer = 0)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrTrap

    If OpeningID = Null Then
        Exit Function
    End If

    'StrSQL = "select * From QryBalanceOpening where Transaction_ID=" & OpeningID
    OpeningID = 3

    If BranchView = False Then
        StrSQL = " SELECT     dbo.FixedAssetsGroup.GroupName AS GroupName, dbo.FixedAssetsGroup.GroupNamee AS GroupNamee, dbo.FixedAssets.code, dbo.FixedAssets.Name, "
        StrSQL = StrSQL & "  dbo.FixedAssets.group_id, dbo.FixedAssets.Branch_NO, dbo.FixedAssets.Emp_id, dbo.FixedAssets.ReceiveDate, dbo.FixedAssets.CurrentValue,"
        StrSQL = StrSQL & "  dbo.FixedAssets.AccDepreciation, dbo.FixedAssets.Status_id, dbo.FixedAssets.Depreciation_Type_id, dbo.FixedAssets.DefaultAge,"
        StrSQL = StrSQL & "  dbo.FixedAssets.StartDepreciationDate, dbo.FixedAssets.LastDepreciationDate, dbo.FixedAssets.NoOfInstallments, dbo.FixedAssets.EXEInstallments,"
        StrSQL = StrSQL & "  dbo.FixedAssets.RemainInstallments, dbo.FixedAssets.PurchasePrice, dbo.FixedAssets.PurchaseDate, dbo.FixedAssets.PurchaseBillId,"
        StrSQL = StrSQL & "  dbo.FixedAssets.KhordaPrice, dbo.FixedAssets.InstallmentValue, dbo.FixedAssets.New_or_opening, dbo.FixedAssets.Notes, dbo.FixedAssets.Fullcode,"
        StrSQL = StrSQL & "  dbo.FixedAssets.prifix, dbo.FixedAssets.NoteID, dbo.FixedAssets.NoteSerial, dbo.FixedAssets.variance, dbo.FixedAssets.NoteSerial1,"
        StrSQL = StrSQL & "  dbo.FixedAssets.NoteID1 , dbo.FixedAssets.HaveDepreciation, dbo.FixedAssets.opening_balance_voucher_id"
        StrSQL = StrSQL & " , dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee"
        StrSQL = StrSQL & "  FROM         dbo.FixedAssets INNER JOIN"
        StrSQL = StrSQL & "  dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID"
        StrSQL = StrSQL & "  LEFT OUTER JOIN   dbo.TblEmployee ON dbo.FixedAssets.Emp_id = dbo.TblEmployee.Emp_ID"
        StrSQL = StrSQL + " order by dbo.FixedAssets.id"

    Else
        StrSQL = "SELECT     ISNULL(dbo.FixedAssets.SalePrice, 0) AS saleprice, dbo.FixedAssetsGroup.GroupName AS GroupName, dbo.FixedAssetsGroup.GroupNamee AS GroupNamee, "
        StrSQL = StrSQL & "               dbo.FixedAssets.code, dbo.FixedAssets.Name,dbo.FixedAssets.Namee, dbo.FixedAssets.group_id, dbo.FixedAssets.Branch_NO, dbo.FixedAssets.Emp_id, dbo.FixedAssets.ReceiveDate,"
        StrSQL = StrSQL & "               dbo.FixedAssets.CurrentValue, dbo.FixedAssets.AccDepreciation, dbo.FixedAssets.Status_id, dbo.FixedAssets.Depreciation_Type_id, dbo.FixedAssets.DefaultAge,"
        StrSQL = StrSQL & "               dbo.FixedAssets.StartDepreciationDate, dbo.FixedAssets.LastDepreciationDate, dbo.FixedAssets.NoOfInstallments, dbo.FixedAssets.EXEInstallments,"
        StrSQL = StrSQL & "               dbo.FixedAssets.RemainInstallments, dbo.FixedAssets.PurchasePrice, dbo.FixedAssets.PurchaseDate, dbo.FixedAssets.PurchaseBillId,"
        StrSQL = StrSQL & "               dbo.FixedAssets.KhordaPrice, dbo.FixedAssets.InstallmentValue, dbo.FixedAssets.New_or_opening, dbo.FixedAssets.Notes, dbo.FixedAssets.Fullcode,"
        StrSQL = StrSQL & "               dbo.FixedAssets.prifix, dbo.FixedAssets.NoteID, dbo.FixedAssets.NoteSerial, dbo.FixedAssets.variance, dbo.FixedAssets.NoteSerial1, dbo.FixedAssets.NoteID1,"
        StrSQL = StrSQL & "               dbo.FixedAssets.HaveDepreciation, dbo.FixedAssets.opening_balance_voucher_id, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
        StrSQL = StrSQL & "               dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.GetInstalAddValue(dbo.FixedAssets.id) AS AddValue, dbo.FixedAssets.id,"
        StrSQL = StrSQL & "               dbo.GetInstalAdNo(dbo.FixedAssets.id) AS QstAddedNo , dbo.GetQstAddValue(dbo.FixedAssets.id) AS QstIncValue , dbo.GetAssetDiscuValue(dbo.FixedAssets.id)AS DisValue "
        StrSQL = StrSQL & "    FROM         dbo.FixedAssets INNER JOIN"
        StrSQL = StrSQL & "               dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID INNER JOIN"
        StrSQL = StrSQL & "               dbo.TblBranchesData ON dbo.FixedAssets.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
        StrSQL = StrSQL & "               dbo.TblEmployee ON dbo.FixedAssets.Emp_id = dbo.TblEmployee.Emp_ID"
        StrSQL = StrSQL & "   where 1=1"

        If branch_no <> 0 Then
            StrSQL = StrSQL & " and  dbo.FixedAssets.Branch_NO=" & branch_no
        End If
         If FixedID <> 0 Then
            StrSQL = StrSQL & " and  dbo.FixedAssets.ID=" & FixedID
        End If

        If group_id <> 0 Then
            StrSQL = StrSQL & " and  dbo.FixedAssets.group_id=" & group_id
        End If

        If Emp_id <> 0 Then
            StrSQL = StrSQL & " and  dbo.FixedAssets.Emp_id=" & Emp_id
        End If

        StrSQL = StrSQL & "  ORDER BY dbo.FixedAssets.id"

    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
 
    Dim Reportname As String
 
    If SystemOptions.UserInterface = ArabicInterface Then
        If BranchView = False Then
            Reportname = (App.path & "\Reports\FixedASSETReportAudit.rpt")
            
        Else
            Reportname = (App.path & "\Reports\FixedASSETReportAudit1.rpt")
        End If
 
    Else
 
        If BranchView = False Then
            Reportname = (App.path & "\Reports\FixedASSETReportAudit.rpt")
            
        Else
            Reportname = (App.path & "\Reports\FixedASSETReportAudit1.rpt")
        End If
 
    End If
 
    If Reportname = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\OpeningPalanceReport.rpt"
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
        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
  '  Set FrmReport = New FrmReportViewer
  '  FrmReport.CRViewer.ReportSource = xReport
  '  FrmReport.CRViewer.viewReport
        
        xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , Reportname, , StrSQL
    
  '  FrmReport.show
  '  FrmReport.txtPath = Reportname
   ' FrmReport
    Screen.MousePointer = vbDefault
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

Public Function ShowFixedAssetsxx(OpeningID As Integer, _
                                Optional BranchView As Boolean = True, _
                                Optional branch_no As Integer = 0, _
                                Optional group_id As Integer = 0, _
                                Optional Emp_id As Integer = 0, _
                                Optional FixedID As Integer = 0)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrTrap

    If OpeningID = Null Then
        Exit Function
    End If

    'StrSQL = "select * From QryBalanceOpening where Transaction_ID=" & OpeningID
    OpeningID = 3

    If BranchView = False Then
        StrSQL = " SELECT     dbo.FixedAssetsGroup.GroupName AS GroupName, dbo.FixedAssetsGroup.GroupNamee AS GroupNamee, dbo.FixedAssets.code, dbo.FixedAssets.Name, "
        StrSQL = StrSQL & "  dbo.FixedAssets.group_id, dbo.FixedAssets.Branch_NO, dbo.FixedAssets.Emp_id, dbo.FixedAssets.ReceiveDate, dbo.FixedAssets.CurrentValue,"
        StrSQL = StrSQL & "  dbo.FixedAssets.AccDepreciation, dbo.FixedAssets.Status_id, dbo.FixedAssets.Depreciation_Type_id, dbo.FixedAssets.DefaultAge,"
        StrSQL = StrSQL & "  dbo.FixedAssets.StartDepreciationDate, dbo.FixedAssets.LastDepreciationDate, dbo.FixedAssets.NoOfInstallments, dbo.FixedAssets.EXEInstallments,"
        StrSQL = StrSQL & "  dbo.FixedAssets.RemainInstallments, dbo.FixedAssets.PurchasePrice, dbo.FixedAssets.PurchaseDate, dbo.FixedAssets.PurchaseBillId,"
        StrSQL = StrSQL & "  dbo.FixedAssets.KhordaPrice, dbo.FixedAssets.InstallmentValue, dbo.FixedAssets.New_or_opening, dbo.FixedAssets.Notes, dbo.FixedAssets.Fullcode,"
        StrSQL = StrSQL & "  dbo.FixedAssets.prifix, dbo.FixedAssets.NoteID, dbo.FixedAssets.NoteSerial, dbo.FixedAssets.variance, dbo.FixedAssets.NoteSerial1,"
        StrSQL = StrSQL & "  dbo.FixedAssets.NoteID1 , dbo.FixedAssets.HaveDepreciation, dbo.FixedAssets.opening_balance_voucher_id"
        StrSQL = StrSQL & " , dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee"
        StrSQL = StrSQL & "  FROM         dbo.FixedAssets INNER JOIN"
        StrSQL = StrSQL & "  dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID"
        StrSQL = StrSQL & "  LEFT OUTER JOIN   dbo.TblEmployee ON dbo.FixedAssets.Emp_id = dbo.TblEmployee.Emp_ID"
        StrSQL = StrSQL + " order by dbo.FixedAssets.id"

    Else
        StrSQL = "SELECT     ISNULL(dbo.FixedAssets.SalePrice, 0) AS saleprice, dbo.FixedAssetsGroup.GroupName AS GroupName, dbo.FixedAssetsGroup.GroupNamee AS GroupNamee, "
        StrSQL = StrSQL & "               dbo.FixedAssets.code, dbo.FixedAssets.Name, dbo.FixedAssets.group_id, dbo.FixedAssets.Branch_NO, dbo.FixedAssets.Emp_id, dbo.FixedAssets.ReceiveDate,"
        StrSQL = StrSQL & "               dbo.FixedAssets.CurrentValue, dbo.FixedAssets.AccDepreciation, dbo.FixedAssets.Status_id, dbo.FixedAssets.Depreciation_Type_id, dbo.FixedAssets.DefaultAge,"
        StrSQL = StrSQL & "               dbo.FixedAssets.StartDepreciationDate, dbo.FixedAssets.LastDepreciationDate, dbo.FixedAssets.NoOfInstallments, dbo.FixedAssets.EXEInstallments,"
        StrSQL = StrSQL & "               dbo.FixedAssets.RemainInstallments, dbo.FixedAssets.PurchasePrice, dbo.FixedAssets.PurchaseDate, dbo.FixedAssets.PurchaseBillId,"
        StrSQL = StrSQL & "               dbo.FixedAssets.KhordaPrice, dbo.FixedAssets.InstallmentValue, dbo.FixedAssets.New_or_opening, dbo.FixedAssets.Notes, dbo.FixedAssets.Fullcode,"
        StrSQL = StrSQL & "               dbo.FixedAssets.prifix, dbo.FixedAssets.NoteID, dbo.FixedAssets.NoteSerial, dbo.FixedAssets.variance, dbo.FixedAssets.NoteSerial1, dbo.FixedAssets.NoteID1,"
        StrSQL = StrSQL & "               dbo.FixedAssets.HaveDepreciation, dbo.FixedAssets.opening_balance_voucher_id, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
        StrSQL = StrSQL & "               dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.GetInstalAddValue(dbo.FixedAssets.id) AS AddValue, dbo.FixedAssets.id,"
        StrSQL = StrSQL & "               dbo.GetInstalAdNo(dbo.FixedAssets.id) AS QstAddedNo , dbo.GetQstAddValue(dbo.FixedAssets.id) AS QstIncValue"
        StrSQL = StrSQL & "    FROM         dbo.FixedAssets INNER JOIN"
        StrSQL = StrSQL & "               dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID INNER JOIN"
        StrSQL = StrSQL & "               dbo.TblBranchesData ON dbo.FixedAssets.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
        StrSQL = StrSQL & "               dbo.TblEmployee ON dbo.FixedAssets.Emp_id = dbo.TblEmployee.Emp_ID"
        StrSQL = StrSQL & "   where 1=1"

        If branch_no <> 0 Then
            StrSQL = StrSQL & " and  dbo.FixedAssets.Branch_NO=" & branch_no
        End If
         If FixedID <> 0 Then
            StrSQL = StrSQL & " and  dbo.FixedAssets.ID=" & FixedID
        End If

        If group_id <> 0 Then
            StrSQL = StrSQL & " and  dbo.FixedAssets.group_id=" & group_id
        End If

        If Emp_id <> 0 Then
            StrSQL = StrSQL & " and  dbo.FixedAssets.Emp_id=" & Emp_id
        End If

        StrSQL = StrSQL & "  ORDER BY dbo.FixedAssets.id"

    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
 
    Dim Reportname As String
 
    If SystemOptions.UserInterface = ArabicInterface Then
        If BranchView = False Then
            Reportname = (App.path & "\Reports\FixedASSETReportAudit.rpt")
            
        Else
            Reportname = (App.path & "\Reports\FixedASSETReportAudit1.rpt")
        End If
 
    Else
 
        If BranchView = False Then
            Reportname = (App.path & "\Reports\FixedASSETReportAudit.rpt")
            
        Else
            Reportname = (App.path & "\Reports\FixedASSETReportAudit1.rpt")
        End If
 
    End If
 
    If Reportname = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\OpeningPalanceReport.rpt"
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
        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    
    FrmReport.show
    FrmReport.txtPath = Reportname
    Screen.MousePointer = vbDefault
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function


Public Function ShowFixedAssetsxxx(OpeningID As Integer, _
                                Optional BranchView As Boolean = True, _
                                Optional branch_no As Integer = 0, _
                                Optional group_id As Integer = 0, _
                                Optional Emp_id As Integer = 0)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrTrap

    If OpeningID = Null Then
        Exit Function
    End If

    'StrSQL = "select * From QryBalanceOpening where Transaction_ID=" & OpeningID
    OpeningID = 3

    If BranchView = False Then
        StrSQL = " SELECT     dbo.FixedAssetsGroup.GroupName AS GroupName, dbo.FixedAssetsGroup.GroupNamee AS GroupNamee, dbo.FixedAssets.code, dbo.FixedAssets.Name, "
        StrSQL = StrSQL & "  dbo.FixedAssets.group_id, dbo.FixedAssets.Branch_NO, dbo.FixedAssets.Emp_id, dbo.FixedAssets.ReceiveDate, dbo.FixedAssets.CurrentValue,"
        StrSQL = StrSQL & "  dbo.FixedAssets.AccDepreciation, dbo.FixedAssets.Status_id, dbo.FixedAssets.Depreciation_Type_id, dbo.FixedAssets.DefaultAge,"
        StrSQL = StrSQL & "  dbo.FixedAssets.StartDepreciationDate, dbo.FixedAssets.LastDepreciationDate, dbo.FixedAssets.NoOfInstallments, dbo.FixedAssets.EXEInstallments,"
        StrSQL = StrSQL & "  dbo.FixedAssets.RemainInstallments, dbo.FixedAssets.PurchasePrice, dbo.FixedAssets.PurchaseDate, dbo.FixedAssets.PurchaseBillId,"
        StrSQL = StrSQL & "  dbo.FixedAssets.KhordaPrice, dbo.FixedAssets.InstallmentValue, dbo.FixedAssets.New_or_opening, dbo.FixedAssets.Notes, dbo.FixedAssets.Fullcode,"
        StrSQL = StrSQL & "  dbo.FixedAssets.prifix, dbo.FixedAssets.NoteID, dbo.FixedAssets.NoteSerial, dbo.FixedAssets.variance, dbo.FixedAssets.NoteSerial1,"
        StrSQL = StrSQL & "  dbo.FixedAssets.NoteID1 , dbo.FixedAssets.HaveDepreciation, dbo.FixedAssets.opening_balance_voucher_id"
        StrSQL = StrSQL & " , dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee"
        StrSQL = StrSQL & "  FROM         dbo.FixedAssets INNER JOIN"
        StrSQL = StrSQL & "  dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID"
        StrSQL = StrSQL & "  LEFT OUTER JOIN   dbo.TblEmployee ON dbo.FixedAssets.Emp_id = dbo.TblEmployee.Emp_ID"
        StrSQL = StrSQL + " order by dbo.FixedAssets.id"

    Else
        StrSQL = "SELECT    ISNULL(dbo.FixedAssets.SalePrice, 0) AS saleprice,dbo.FixedAssetsGroup.GroupName AS GroupName, dbo.FixedAssetsGroup.GroupNamee AS GroupNamee, dbo.FixedAssets.code, "
        StrSQL = StrSQL & "  dbo.FixedAssets.Name, dbo.FixedAssets.group_id, dbo.FixedAssets.Branch_NO, dbo.FixedAssets.Emp_id, dbo.FixedAssets.ReceiveDate,"
        StrSQL = StrSQL & "  dbo.FixedAssets.CurrentValue, dbo.FixedAssets.AccDepreciation, dbo.FixedAssets.Status_id, dbo.FixedAssets.Depreciation_Type_id, dbo.FixedAssets.DefaultAge,"
        StrSQL = StrSQL & "  dbo.FixedAssets.StartDepreciationDate, dbo.FixedAssets.LastDepreciationDate, dbo.FixedAssets.NoOfInstallments, dbo.FixedAssets.EXEInstallments,"
        StrSQL = StrSQL & "  dbo.FixedAssets.RemainInstallments, dbo.FixedAssets.PurchasePrice, dbo.FixedAssets.PurchaseDate, dbo.FixedAssets.PurchaseBillId,"
        StrSQL = StrSQL & "  dbo.FixedAssets.KhordaPrice, dbo.FixedAssets.InstallmentValue, dbo.FixedAssets.New_or_opening, dbo.FixedAssets.Notes, dbo.FixedAssets.Fullcode,"
        StrSQL = StrSQL & "  dbo.FixedAssets.prifix, dbo.FixedAssets.NoteID, dbo.FixedAssets.NoteSerial, dbo.FixedAssets.variance, dbo.FixedAssets.NoteSerial1,"
        StrSQL = StrSQL & "  dbo.FixedAssets.NoteID1, dbo.FixedAssets.HaveDepreciation, dbo.FixedAssets.opening_balance_voucher_id, dbo.TblEmployee.Emp_Name,"
        StrSQL = StrSQL & "  dbo.TblEmployee.Emp_Namee , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
        StrSQL = StrSQL & "  FROM         dbo.FixedAssets INNER JOIN"
        StrSQL = StrSQL & "  dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID INNER JOIN"
        StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.FixedAssets.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
        StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.FixedAssets.Emp_id = dbo.TblEmployee.Emp_ID where 1=1"

        If branch_no <> 0 Then
            StrSQL = StrSQL & " and  dbo.FixedAssets.Branch_NO=" & branch_no
        End If

        If group_id <> 0 Then
            StrSQL = StrSQL & " and  dbo.FixedAssets.group_id=" & group_id
        End If

        If Emp_id <> 0 Then
            StrSQL = StrSQL & " and  dbo.FixedAssets.Emp_id=" & Emp_id
        End If

        StrSQL = StrSQL & "  ORDER BY dbo.FixedAssets.id"

    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText
 
    Dim Reportname As String
 
    If SystemOptions.UserInterface = ArabicInterface Then
        If BranchView = False Then
            Reportname = (App.path & "\Reports\FixedASSETReportAudit.rpt")
            
        Else
            Reportname = (App.path & "\Reports\FixedASSETReportAudit1.rpt")
        End If
 
    Else
 
        If BranchView = False Then
            Reportname = (App.path & "\Reports\FixedASSETReportAudit.rpt")
            
        Else
            Reportname = (App.path & "\Reports\FixedASSETReportAudit1.rpt")
        End If
 
    End If
 
    If Reportname = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\OpeningPalanceReport.rpt"
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
        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    
    FrmReport.show
    FrmReport.txtPath = Reportname
    Screen.MousePointer = vbDefault
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

