VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGardReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function ShowGardData(GardID As Integer, _
                             Optional IntStockType As Integer = 0, _
                             Optional Hideqty As Boolean, _
                             Optional FromDate As Date)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim StrFileName As String
    'IntStockType=0 ' ›’Ì·Ï
    'IntStockType=1 '≈Ã„«·Ï
    On Error GoTo ErrTrap

    If GardID = 0 Then
        Exit Function
    End If

    StrSQL = "Select * From QryGardReport where Transaction_ID=" & GardID
    StrSQL = StrSQL + " Order By GroupID,Item_ID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If IntStockType = 0 Then
        StrFileName = App.path & "\Reports\GardReportsDetails.rpt"
    ElseIf IntStockType = 1 Then
        StrFileName = App.path & "\Reports\GardReports.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
            
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = " ﬁ—Ì— «·Ã—œ"
        Else
            xReport.reporttitle = "Stock Adjustement"
        End If

        xReport.ApplicationName = xReport.reporttitle

        If Hideqty = True Then
            xReport.ParameterFields(4).AddCurrentValue "1"
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
        End If
        
        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
    FrmReport.show
    FrmReport.txtPath = StrFileName
    FrmReport.PreviewCaption = xReport.reporttitle
    
    Screen.MousePointer = vbDefault
     
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

  
  Public Function ShowGardData1(Optional StoreID As Integer, _
                              Optional FromDate As Date, _
                              Optional ToDate As Date, _
                              Optional Hideqty As Boolean, Optional freezstring As String, Optional StoreName As String = "", Optional LocName As String = "", Optional LocName2 As String = "")

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
  Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim StrFileName As String
    'IntStockType=0 ' ›’Ì·Ï
    'IntStockType=1 '≈Ã„«·Ï
'    On Error GoTo ErrTrap
  Dim mLocName As Integer
    StrSQL = "SELECT   dbo.TblItems.barcodeno ,     dbo.TblItems.Fullcode, dbo.TblItems.BinLocation, SUM(dbo.Transaction_Details.ShowQty * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.Transaction_Details.Item_ID AS ItemID, "
    StrSQL = StrSQL & "  dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.Transaction_Details.UnitId, dbo.TblItems.ItemCode,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemName, dbo.TblUnites.UnitName, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ClassId,"
    StrSQL = StrSQL & "  dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblItemsSizes.SizeName AS SizeName, dbo.TblItemsColors.ColorName , dbo.Transaction_Details.LotNO,   dbo.Transaction_Details.MixNo"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL + " where  dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & ""
    StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & ""
    StrSQL = StrSQL & "   AND (dbo.Transactions.StoreID =" & StoreID & ") and  dbo.TransactionTypes.StockEffect<>0"
    If Trim(LocName) <> "" Then
        
        mLocName = CInt(LocName)
        StrSQL = StrSQL & "   AND (dbo.Transactions.StoreID =" & StoreID & ") and  dbo.TblItems.ItemID In (Select ItemID From TblItemLoc Where TblItemLoc.StoreID = " & StoreID & " and Locid = " & mLocName & " )"
    End If
    StrSQL = StrSQL & "  GROUP BY  dbo.TblItems.barcodeno,dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID, dbo.TblStore.StoreName , dbo.Transaction_Details.UnitId,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemCode, dbo.TblItems.Fullcode,dbo.TblItems.BinLocation, dbo.TblItems.ItemName, dbo.TblUnites.UnitName, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize,"
    StrSQL = StrSQL & "   dbo.Transaction_Details.ClassId , dbo.TblItemsclasses.SizeName, dbo.TblItemsSizes.SizeName, dbo.TblItemsColors.ColorName , dbo.Transaction_Details.LotNO,   dbo.Transaction_Details.MixNo"
    
   ' StrSQL = StrSQL & "  Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) <> 0)"
    
    StrSQL = StrSQL & "  ORDER BY dbo.TblItems." & freezstring & ""
 
 Dim fromdateS As Date
 Dim todateS As Date
 Dim openingdate As Date
 Dim openingdateS As Date
 If Hideqty = False Then
   Cn.CommandTimeout = 10000
        fromdateS = Replace(Format$(FromDate, "DD/MM/yyyy"), "-", "/")
        todateS = Replace(Format$(ToDate, "DD/MM/yyyy"), "-", "/")
        openingdateS = DateAdd("D", -1, fromdateS)
      '  openingdateS = Replace(Format$(openingdate, "MM/DD/yyyy"), "-", "/")
 
Dim StrSQL1 As String, strsql2 As String, StrSQL3 As String, StrSQL4 As String
Dim StrSQL5 As String, StrSQL6 As String, StrSQL7 As String

StrSQL1 = "SELECT TI.ItemID, TI.ItemName, TD.ColorID, TD.ItemSize, TD.ClassId, TD.LotNO, TD.MixNo, " & _
"TI.barcodeno, TI.ItemCode, TU.UnitName, TC.SizeName AS ClassName, TS.SizeName AS SizeName, " & _
"TCol.ColorName, TD.NoCount, TD.Area, TD.Length, TD.Height, TD.Width, TD.IsExpirDate, TD.L, TD.W, "

strsql2 = "TD.H1, TD.H2, ISNULL(SUM(QryInv.outputValue), 0) AS outputValue, " & _
"ISNULL(SUM(QryInv.inputvalue), 0) AS inputvalue, ISNULL(SUM(QryInv.openingBalance), 0) AS openingValue, " & _
"0 AS Cost, ISNULL(OldQty.oldopening, 0) AS oldopening "
'StrSQL3 = "FROM dbo.TblItems TI " & _
'"LEFT JOIN dbo.Transaction_Details TD ON TD.Item_ID = TI.ItemID " & _
'"LEFT JOIN dbo.TblUnites TU ON TD.UnitId = TU.UnitID " & _
'"LEFT JOIN dbo.TblItemsclasses TC ON TD.ClassId = TC.SizeId " & _
'"LEFT JOIN dbo.TblItemsSizes TS ON TD.ItemSize = TS.SizeId " & _
'"LEFT JOIN dbo.TblItemsColors TCol ON TD.ColorID = TCol.ColorID " & _
'"LEFT JOIN ( SELECT * FROM dbo.QryItemsInventry4('06 Jun 2025', '23 Jun 2025', 2, NULL, NULL, NULL, NULL, NULL) ) QryInv " & _
'"ON QryInv.Item_ID = TI.ItemID " & _
'"AND ISNULL(QryInv.ColorID, -1) = ISNULL(TD.ColorID, -1) " & _
'"AND ISNULL(QryInv.ItemSize, 'NULL') = ISNULL(TD.ItemSize, 'NULL') " & _
'"AND ISNULL(QryInv.ClassId, -1) = ISNULL(TD.ClassId, -1) " & _
'"AND ISNULL(QryInv.LotNO, 'NULL') = ISNULL(TD.LotNO, 'NULL') " & _
'"AND ISNULL(QryInv.Length, -1) = ISNULL(TD.Length, -1) " & _
'"AND ISNULL(QryInv.Height, -1) = ISNULL(TD.Height, -1) " & _
'"AND ISNULL(QryInv.Width, -1) = ISNULL(TD.Width, -1) "

StrSQL3 = "FROM dbo.TblItems TI " & _
"LEFT JOIN dbo.Transaction_Details TD ON TD.Item_ID = TI.ItemID " & _
"LEFT JOIN dbo.TblUnites TU ON TD.UnitId = TU.UnitID " & _
"LEFT JOIN dbo.TblItemsclasses TC ON TD.ClassId = TC.SizeId " & _
"LEFT JOIN dbo.TblItemsSizes TS ON TD.ItemSize = TS.SizeId " & _
"LEFT JOIN dbo.TblItemsColors TCol ON TD.ColorID = TCol.ColorID " & _
"LEFT JOIN ( SELECT * FROM dbo.QryItemsInventry4(" & _
      SQLDate(FromDate, True) & ", " & _
      SQLDate(ToDate, True) & ", " & _
      StoreID & ", " & _
      "NULL, NULL, NULL, NULL, NULL" & _
" ) ) QryInv " & _
"ON QryInv.Item_ID = TI.ItemID " & _
"AND ISNULL(QryInv.ColorID, 0) = ISNULL(TD.ColorID, 0) " & _
"AND ISNULL(QryInv.ItemSize, 'NULL') = ISNULL(TD.ItemSize, 'NULL') " & _
"AND ISNULL(QryInv.ClassId, 0) = ISNULL(TD.ClassId, 0) " & _
"AND ISNULL(QryInv.LotNO, 'NULL') = ISNULL(TD.LotNO, 'NULL') "
'"AND ISNULL(QryInv.Length, 0) = ISNULL(TD.Length, 0) " & _
'"AND ISNULL(QryInv.Height, 0) = ISNULL(TD.Height, 0) " & _
'"AND ISNULL(QryInv.Width, 0) = ISNULL(TD.Width, 0) "

StrSQL4 = "LEFT JOIN ( SELECT TD2.Item_ID, TD2.ColorID, TD2.ItemSize, TD2.ClassId, TD2.LotNO, TD2.Length, TD2.Height, TD2.Width, " & _
"SUM(TD2.Quantity * TT.StockEffect) AS oldopening " & _
"FROM dbo.Transaction_Details TD2 " & _
"INNER JOIN dbo.Transactions T2 ON TD2.Transaction_ID = T2.Transaction_ID " & _
"INNER JOIN dbo.TransactionTypes TT ON T2.Transaction_Type = TT.Transaction_Type " & _
"WHERE TT.StockEffect <> 0 AND T2.Transaction_Date <= " & SQLDate(FromDate, True) & "  AND T2. Storeid = " & StoreID & "  " & _
"GROUP BY TD2.Item_ID, TD2.ColorID, TD2.ItemSize, TD2.ClassId, TD2.LotNO, TD2.Length, TD2.Height, TD2.Width ) OldQty " & _
"ON TI.ItemID = OldQty.Item_ID " & _
"AND TD.ColorID = OldQty.ColorID " & _
"AND TD.ItemSize = OldQty.ItemSize " & _
"AND TD.ClassId = OldQty.ClassId " & _
"AND TD.LotNO = OldQty.LotNO "
'"AND TD.Length = OldQty.Length " & _
'"AND TD.Height = OldQty.Height " & _
'"AND TD.Width = OldQty.Width "
StrSQL5 = "WHERE TI.ItemID IN (SELECT DISTINCT TD3.Item_ID FROM dbo.Transactions T3 " & _
"INNER JOIN dbo.TransactionTypes TT3 ON T3.Transaction_Type = TT3.Transaction_Type " & _
"INNER JOIN dbo.Transaction_Details TD3 ON T3.Transaction_ID = TD3.Transaction_ID " & _
"WHERE TT3.StockEffect <> 0 AND T3. Storeid = " & StoreID & "  ) " & _
"GROUP BY TI.ItemID, TI.ItemName, TD.ColorID, TD.ItemSize, TD.ClassId, TD.LotNO, TD.MixNo, TI.barcodeno, TI.ItemCode, " & _
"TU.UnitName, TC.SizeName, TS.SizeName, TCol.ColorName, TD.NoCount, TD.Area, " & _
"TD.IsExpirDate,  OldQty.oldopening " & _
" Having " & _
    "ISNULL(SUM(QryInv.outputValue), 0) <> 0 " & _
    "OR ISNULL(SUM(QryInv.inputvalue), 0) <> 0 " & _
    "OR ISNULL(SUM(QryInv.openingBalance), 0) <> 0 " & _
    "OR ISNULL(OldQty.oldopening, 0) <> 0 " & _
"ORDER BY TI.ItemName; "

Dim FinalSQL As String
FinalSQL = StrSQL1 & strsql2 & StrSQL3 & StrSQL4 & StrSQL5




StrSQL = FinalSQL 'StrSQL1 & strsql2 & StrSQL3 & StrSQL4 & StrSQL5 & StrSQL6 & StrSQL7




Dim MainSQL As String

StrSQL1 = "WITH Inv AS (" & _
" SELECT * FROM dbo.QryItemsInventry4(" & _
"     " & SQLDate(FromDate, True) & ", " & SQLDate(ToDate, True) & ", " & StoreID & ", NULL, NULL, NULL, NULL, NULL" & _
" )" & _
") "

strsql2 = "SELECT Inv.Item_ID, TI.ItemCode, TI.barcodeno, TI.ItemName,TI.ItemNamee, " & _
" Inv.ColorID, TCol.ColorName, Inv.ItemSize, TS.SizeName AS SizeName, " & _
" Inv.ClassId, TC.SizeName AS ClassName, Inv.LotNO, " & _
" Inv.openingValue, Inv.inputvalue, Inv.outputValue, Inv.openingBalance "

StrSQL3 = "FROM Inv " & _
" LEFT JOIN dbo.TblItems TI ON Inv.Item_ID = TI.ItemID " & _
" LEFT JOIN dbo.TblItemsColors TCol ON Inv.ColorID = TCol.ColorID " & _
" LEFT JOIN dbo.TblItemsSizes TS ON Inv.ItemSize = TS.SizeId " & _
" LEFT JOIN dbo.TblItemsclasses TC ON Inv.ClassId = TC.SizeId "

' «·›· —… «·‰Â«∆Ì… ⁄·Ï «·„Œ“‰ (StoreID) Ê«· Ê«—ÌŒ
StrSQL4 = "WHERE Inv.Item_ID IN (SELECT DISTINCT TD3.Item_ID " & _
" FROM dbo.Transactions T3 " & _
" INNER JOIN dbo.TransactionTypes TT3 ON T3.Transaction_Type = TT3.Transaction_Type " & _
" INNER JOIN dbo.Transaction_Details TD3 ON T3.Transaction_ID = TD3.Transaction_ID " & _
" WHERE TT3.StockEffect <> 0 " & _
" AND T3.Storeid = " & StoreID & " " & _
" AND T3.Transaction_Date >=" & SQLDate(FromDate, True) & _
" AND T3.Transaction_Date <= " & SQLDate(ToDate, True) & " ) " & _
"ORDER BY TI.ItemName, Inv.ColorID, Inv.ItemSize, Inv.ClassId, Inv.LotNO"

MainSQL = StrSQL1 & strsql2 & StrSQL3 & StrSQL4
StrSQL = MainSQL
' Ì„ﬂ‰ﬂ «” Œœ«„ MainSQL „»«‘—… „⁄ rs.Open √Ê √Ì  ‰›Ì– ·œÌﬂ

 StrSQL1 = "WITH Inv AS (" & _
" SELECT * FROM dbo.QryItemsInventry4(" & _
"     " & SQLDate(FromDate, True) & ", " & SQLDate(ToDate, True) & ", " & StoreID & ", NULL, NULL, NULL, NULL, NULL" & _
" )" & _
") "

strsql2 = "SELECT Inv.Item_ID, TI.ItemCode, TI.barcodeno, TI.ItemName,TI.ItemNamee, " & _
" Inv.ColorID, TCol.ColorName, Inv.ItemSize, TS.SizeName AS SizeName, " & _
" Inv.ClassId, TC.SizeName AS ClassName, Inv.LotNO, Inv.UnitId, TU.UnitName, " & _
"  Inv.openingBalance openingValue, Inv.inputvalue, Inv.outputValue, Inv.openingBalance "

StrSQL3 = "FROM Inv " & _
" LEFT JOIN dbo.TblItems TI ON Inv.Item_ID = TI.ItemID " & _
" LEFT JOIN dbo.TblUnites TU ON Inv.UnitId = TU.UnitID " & _
" LEFT JOIN dbo.TblItemsColors TCol ON Inv.ColorID = TCol.ColorID " & _
" LEFT JOIN dbo.TblItemsSizes TS ON Inv.ItemSize = TS.SizeId " & _
" LEFT JOIN dbo.TblItemsclasses TC ON Inv.ClassId = TC.SizeId "

StrSQL4 = "ORDER BY Ti.Itemcode , Inv.UnitId,Inv.ColorID, Inv.ItemSize, Inv.ClassId"
', Inv.Length, Inv.Height, Inv.Width "

MainSQL = StrSQL1 & strsql2 & StrSQL3 & StrSQL4
StrSQL = MainSQL
 End If
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If Hideqty = False Then
 StrFileName = App.path & "\Reports\Inventory\Gard20.rpt"
 Else
    StrFileName = App.path & "\Reports\Inventory\Gard2.rpt"
 End If
    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
            
             If Hideqty = False Then
                xReport.ParameterFields(8).AddCurrentValue StoreName
         
             
             End If
             
        For i = 1 To xReport.FormulaFields.count
            Select Case xReport.FormulaFields.Item(i).Name
            Case "{@LocName}"
                xReport.FormulaFields.Item(i).text = "'" & LocName2 & "'"
            End Select
        Next i
             
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "   ﬂ‘›Ê›«  «·Ã—œ  "
        Else
            xReport.reporttitle = "Stock Adjustement"
        End If

        xReport.ApplicationName = xReport.reporttitle

        If Hideqty = True Then
            xReport.ParameterFields(4).AddCurrentValue "1"
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
        End If

        xReport.ParameterFields(5).AddCurrentValue GetCurrentGardEmployee
        xReport.ParameterFields(6).AddCurrentValue Format(FromDate, "yyyy/m/d")
        xReport.ParameterFields(7).AddCurrentValue Format(ToDate, "yyyy/m/d")
         
        Screen.MousePointer = vbDefault
    End If
      Dim CViewer As ClsReportViewer
 Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, xReport.reporttitle, , , , StrFileName, , StrSQL
    
    
    '    Screen.MousePointer = vbArrowHourglass
'    Set FrmReport = New FrmReportViewer
'    FrmReport.CRViewer.ReportSource = xReport
'    FrmReport.CRViewer.viewReport
'    FrmReport.show
'    FrmReport.txtPath = StrFileName
'    FrmReport.PreviewCaption = xReport.reporttitle
'    FrmReport.TXTSTRSQL = StrSQL
    
    Screen.MousePointer = vbDefault
     
    'salah SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function



Public Function ShowGardData2(sql As String, _
                              Optional FromDate As Date, _
                              Optional ToDate As Date, _
                              Optional ItemName As String)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
 
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim StrFileName As String
    'IntStockType=0 ' ›’Ì·Ï
    'IntStockType=1 '≈Ã„«·Ï
    'On Error GoTo ErrTrap
  
    StrSQL = sql
 
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    StrFileName = App.path & "\Reports\Inventory\Gard5.rpt"
 
    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & CHR(13) & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & CHR(13) & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
            
        If SystemOptions.UserInterface = ArabicInterface Then
            xReport.reporttitle = "    ﬁ—Ì— ﬂ„Ì«  ’‰› "
        Else
            xReport.reporttitle = "items Data"
        End If

        xReport.ApplicationName = xReport.reporttitle

        If Hideqty = True Then
            xReport.ParameterFields(4).AddCurrentValue "1"
        Else
            xReport.ParameterFields(4).AddCurrentValue "0"
        End If

        xReport.ParameterFields(5).AddCurrentValue GetCurrentGardEmployee
        xReport.ParameterFields(6).AddCurrentValue Format(FromDate, "yyyy/m/d")
        xReport.ParameterFields(7).AddCurrentValue Format(ToDate, "yyyy/m/d")
        xReport.ParameterFields(8).AddCurrentValue ItemName
         
        Screen.MousePointer = vbDefault
    End If

    '    Screen.MousePointer = vbArrowHourglass
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.viewReport
   ' FrmReport.show modalless
    FrmReport.txtPath = StrFileName
    FrmReport.PreviewCaption = xReport.reporttitle
    
    Screen.MousePointer = vbDefault
     
    Sendkeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Function

