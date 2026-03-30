VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsReturnBackReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 

 Public Function ShowReturnBack(ReturnID As Long, _
                               Optional shortR As Boolean = False, _
                               Optional Result As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSql As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If ReturnID = Null Then
        Exit Function
    End If

 
 'UPDATE 01 06 2015


    StrSql = " SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID, "
    StrSql = StrSql & "                  dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID,"
    StrSql = StrSql & "                  dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.Price,"
    StrSql = StrSql & "                  dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId,"
    StrSql = StrSql & "                  dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemPhoto, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName,"
    StrSql = StrSql & "                  dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.CashCustomerName,"
    StrSql = StrSql & "                  dbo.Transactions.NoteSerial1, dbo.Transaction_Details.TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount,"
    StrSql = StrSql & "                  dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType, dbo.Transactions.ManualNo1, dbo.Transactions.PaymentType,"
    StrSql = StrSql & "                  dbo.Transactions.ReasonReturns, dbo.Transactions.BillSupplier, dbo.Transactions.OrderSupply, dbo.Transactions.ReturnSerial, dbo.Transaction_Details.IsExpirDate,"
    StrSql = StrSql & "                  dbo.Transaction_Details.Width, dbo.Transaction_Details.Height, dbo.Transaction_Details.Area, dbo.Transaction_Details.NoCount,"
    StrSql = StrSql & "                  dbo.Transaction_Details.ExpiryDate , dbo.Transaction_Details.lotNo , dbo.Transactions.VATNO , dbo.Transactions.VAT, dbo.Transactions.ResonVAT, "
    StrSql = StrSql & "                  dbo.Transactions.Typ, dbo.Transaction_Details.Vat AS VatDet, dbo.Transaction_Details.Vatyo"
    StrSql = StrSql & "     FROM         dbo.Transactions INNER JOIN"
    StrSql = StrSql & "                  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSql = StrSql & "                  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    StrSql = StrSql & "                  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSql = StrSql & "                  dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID INNER JOIN"
    StrSql = StrSql & "                  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSql = StrSql & "                  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSql = StrSql & "                  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSql = StrSql & "                  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    StrSql = StrSql & "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"

    'StrSQL = StrSQL & "  where Transactions.Transaction_ID=" & SallingID

    StrSql = StrSql & "   WHERE     (dbo.Transactions.Transaction_ID = " & ReturnID & ")"

    StrSql = StrSql + "order by dbo.Transaction_Details.id"
'UPDATE 01 06 2015

'    StrSQL = StrSQL + "order by dbo.Transaction_Details.id"
    Dim filename As String
  '  FileName = App.path & "\Reports\ReturnBackReport.rpt"
filename = App.path & "\Special\" & SystemOptions.Reportpath & "\ReturnBackReport.rpt"

    rs.Open StrSql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(filename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = "\Reports\ReturnBackReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(filename)
        
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(8).AddCurrentValue WriteNo(Format(CStr(Result), "0.00"), 0, True, ".", , 0)
        xReport.reporttitle = "„—œÊœ«  «·„‘ —Ì« "
        StrSQL1 = "select * From Notes where Transaction_ID=" & ReturnID
        StrSQL1 = StrSQL1 + " and NoteType=0"
        RsTemp.Open StrSQL1, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.txtpath = filename
    FrmReport.TXTSTRSQL = StrSql

    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & Chr(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function



Public Function ShowReturnBackx03072017(ReturnID As Long, _
                               Optional shortR As Boolean = False, _
                               Optional Result As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSql As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If ReturnID = Null Then
        Exit Function
    End If

 
 'UPDATE 01 06 2015


    StrSql = " SELECT     dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID, dbo.TblCustemers.CusName, "
    StrSql = StrSql & "  dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID,"
    StrSql = StrSql & "  dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.Price,"
    StrSql = StrSql & "  dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId,"
    StrSql = StrSql & "  dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemPhoto, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName,"
    StrSql = StrSql & "  dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.CashCustomerName,"
    StrSql = StrSql & "  dbo.Transactions.NoteSerial1, dbo.Transaction_Details.TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount,"
    StrSql = StrSql & "  dbo.Transactions.Trans_Discount , dbo.Transactions.Trans_DiscountType, dbo.Transactions.ManualNo1, dbo.Transactions.paymenttype ,  dbo.Transactions.ReasonReturns, dbo.Transactions.BillSupplier, dbo.Transactions.OrderSupply , dbo.Transactions.ReturnSerial"
    StrSql = StrSql & "  FROM         dbo.Transactions INNER JOIN"
    StrSql = StrSql & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSql = StrSql & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    StrSql = StrSql & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSql = StrSql & "  dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID INNER JOIN"
    StrSql = StrSql & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSql = StrSql & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSql = StrSql & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSql = StrSql & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    StrSql = StrSql & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"

    'StrSQL = StrSQL & "  where Transactions.Transaction_ID=" & SallingID

    StrSql = StrSql & "   WHERE     (dbo.Transactions.Transaction_ID = " & ReturnID & ")"

    StrSql = StrSql + "order by dbo.Transaction_Details.id"
'UPDATE 01 06 2015

'    StrSQL = StrSQL + "order by dbo.Transaction_Details.id"
    Dim filename As String
  '  FileName = App.path & "\Reports\ReturnBackReport.rpt"
filename = App.path & "\Special\" & SystemOptions.Reportpath & "\ReturnBackReport.rpt"

    rs.Open StrSql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(filename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = "\Reports\ReturnBackReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(filename)
        
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(8).AddCurrentValue WriteNo(Format(CStr(Result), "0.00"), 0, True, ".", , 0)
        xReport.reporttitle = "„—œÊœ«  «·„‘ —Ì« "
        StrSql = "select * From Notes where Transaction_ID=" & ReturnID
        StrSql = StrSql + " and NoteType=0"
        RsTemp.Open StrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.txtpath = filename
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & Chr(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Function ShowReturnBackShort(ReturnID As Long)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSql As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If ReturnID = Null Then
        Exit Function
    End If

    StrSql = "select * From QryBuyReport where Transaction_ID=" & ReturnID
    rs.Open StrSql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(App.path & "\Reports\ReturnBackReportShort.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & Chr(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & Chr(13)
        Msg = "\Reports\ReturnBackReportShort.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
filename = App.path & "\Reports\ReturnBackReportShort.rpt"
        Set xReport = xApp.OpenReport(filename)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        StrSQL1 = "select * From Notes where Transaction_ID=" & ReturnID
        StrSQL1 = StrSQL1 + " and NoteType=0"
        
        RsTemp.Open StrSQL1, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
     
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.TXTSTRSQL = StrSql
    FrmReport.txtpath = filename
    FrmReport.show
    Screen.MousePointer = vbDefault
    '    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & Chr(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & Chr(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

