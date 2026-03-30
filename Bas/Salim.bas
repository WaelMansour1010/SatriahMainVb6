Attribute VB_Name = "Salim"
Public Workshopgroupid As Integer
Public publicCarId As Double

Public SelectedIssueVoucher As Boolean

Public openingbalanceDes As String

Public LogTextA As String

Public LogTexte As String

Public ScreenNameArabic As String

Public ScreenNameEnglish As String

Public FirstOpenOfForm As Boolean

Public GeneralPriceType As Integer

Public totalprofit As Double
Public ChangePW As Boolean
Public BackGroundImag As String

Public Declare Function GetProfileString _
               Lib "kernel32" _
               Alias "GetProfileStringA" (ByVal lpAppName As String, _
                                          ByVal lpKeyName As String, _
                                          ByVal lpDefault As String, _
                                          ByVal lpReturnedString As String, _
                                          ByVal nSize As Long) As Long
Public Function SerchItemspUBLIC(Optional str As String, Optional ByRef sql As String, Optional ByRef SQL1 As String)

If str <> "" Then
 
Dim StrWhere As String
  Dim astrSplit2tems2() As String
  Dim j As Integer
  Dim nElements As Integer
  Dim SearchString As String
StrWhere = ""
SearchString = ""



sql = " select  ItemID,barCodeNO   from  dbo.TblItems where TblItems.IsArchive=0"
If SystemOptions.UserInterface = ArabicInterface Then
SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
Else
SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where TblItems.IsArchive=0"
End If
'name
  If SystemOptions.ShowOnlyItemsOfSales = True Then
     StrWhere = StrWhere & "  and dbo.TblItems.GroupID in(  "
    StrWhere = StrWhere & " From dbo.Groups  WHERE     (ISNULL(POSGroup, 0) = 1))"
  End If
    
    If SystemOptions.WorkWithLINKEDiActivity = True Then
    
  StrWhere = StrWhere & "  and dbo.TblItems.GroupID in(  "
    StrWhere = StrWhere & " select GroupID from fullgroups ()  )"
    
     
End If

  If str = "" Then
GoTo View
End If


          astrSplit2tems2 = Split(str, " ")
          nElements = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
                        If nElements = 0 Then
                                    If SystemOptions.UserInterface = ArabicInterface Then
                                          StrWhere = StrWhere & " and (ItemName Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%'  or fullcode Like N'%" & Trim(str) & "%') "
                                  Else
                                          StrWhere = StrWhere & " and (ItemNamee Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%' or fullcode Like N'%" & Trim(str) & "%' ) "
                                  End If
                                  
                        End If
        If nElements > 0 Then
              SearchString = ""
                            For j = 0 To nElements
                            
                             SearchString = SearchString & "%" & Trim(astrSplit2tems2(j))
                                  
                             Next j
                                 SearchString = SearchString & "%"
                                    If SystemOptions.UserInterface = ArabicInterface Then
                        
                                          StrWhere = StrWhere + " and (ItemName Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
                                     Else
                                            StrWhere = StrWhere + " and (ItemNamee Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
                                     End If
         
      
         End If
        
    sql = sql & StrWhere
   SQL1 = SQL1 & StrWhere
     
   End If


View:
sql = sql + " Order BY barCodeNO "
       If SystemOptions.UserInterface = ArabicInterface Then
       
        SQL1 = SQL1 + " Order BY ItemName "
    Else
        SQL1 = SQL1 + " Order BY ItemNamee "
    End If
    

        
        
  
        
End Function

Function CheckREpettedAttributionContract(IDAC As Double) As Boolean



 Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

 
        sql = "select * from TblAttributionContract where 1=1 and IDAC =" & IDAC
         
            
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
      CheckREpettedAttributionContract = True
      Else
      CheckREpettedAttributionContract = False
   
    End If

    rs.Close
 
End Function
Function RepeatedCashingVchr(ID As Double, GeneralBoxId As Double, SubBoxId As Double, CashierID As Double, FromDate As Date, ToDate As Date, RecordDate As Date) As Boolean



 Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

 
        sql = "select * from tblGeneralCashing where 1=1 and id <>" & ID
        sql = sql & " and GeneralBoxId=" & GeneralBoxId
         sql = sql & " and CashierID=" & CashierID
          sql = sql & " and FromDate=" & SQLDate(FromDate, True)
           sql = sql & "and  ToDate=" & SQLDate(ToDate, True)
            
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
      RepeatedCashingVchr = True
      Else
      RepeatedCashingVchr = False
   
    End If

    rs.Close
 
End Function
Function ItemOpeningBalances(ItemID As Long, Optional StoreID As Long = 0, Optional ColorID As Integer = 0, Optional sizeid As Integer = 0, Optional ClassId As Integer = 0, Optional FromDate As Variant, Optional Cust As Integer = 0, Optional order_no As String = "")
    Dim openingdate As Date
    Dim sql As String
    Dim rs As New ADODB.Recordset

    If IsNull(FromDate) Then FromDate = "01/01/2000"

    openingdate = DateAdd("D", -1, FromDate)
 
    StrSQL = StrSQL & "select ItemCode, ItemName, ItemID , isnull(  dbo.GetItemqtytodatenew(" & SQLDate(openingdate, True) & ", dbo.TblItems.ItemID, "
    StrSQL = StrSQL & IIf(StoreID = 0, "null", (StoreID))
    StrSQL = StrSQL & ","
    StrSQL = StrSQL & IIf(ColorID = 0, "null", (ColorID))
    StrSQL = StrSQL & ","
    StrSQL = StrSQL & IIf(sizeid = 0, "null", (sizeid))
    StrSQL = StrSQL & ","
    StrSQL = StrSQL & IIf(ClassId = 0, "null", (ClassId))

    StrSQL = StrSQL & ","
    StrSQL = StrSQL & IIf(order_no = "", "null", (order_no))
 
    StrSQL = StrSQL & ","
    StrSQL = StrSQL & IIf(Cust = 0, "null", Cust)
    StrSQL = StrSQL & " ) ,0) AS oldopening"

    StrSQL = StrSQL & "  from dbo.TblItems"
    StrSQL = StrSQL & " Where (1 = 1) And (ItemID = " & ItemID & ")"
    StrSQL = StrSQL & "  ORDER BY ItemCode"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        ItemOpeningBalances = IIf(IsNull(rs("oldopening").value), 0, rs("oldopening").value)
    Else
        ItemOpeningBalances = 0
    End If

End Function

Public Function Translate(Translatetype As Integer, word As String) As String
On Error Resume Next
    Dim openingdate As Date
    Dim sql As String
    Dim rs As New ADODB.Recordset
        Dim VarSet As Variant
        If word = "" Then
            Exit Function
        End If
    word = Trim(word)
    Dim Translatestr As String
VarSet = Split(word, " ", , vbTextCompare)




    StrSQL = "SELECT     aname, Ename From dbo.edictionary"
 
If Translatetype = 0 Then 'arabic to english
 
'StrSQL = StrSQL & " where aname=N'" & VarSet(0) & "'"
'            If VarSet(1) <> Empty Then
'            StrSQL = StrSQL & " or  aname=N'" & VarSet(1) & "'"
'            End If
 StrSQL = StrSQL & " where  aname=N'" & word & "'"
Else
' StrSQL = StrSQL & " where Ename=N'" & VarSet(0) & "'"
'             If VarSet(1) <> Empty Then
'            StrSQL = StrSQL & " or  Ename=N'" & VarSet(1) & "'"
'            End If
StrSQL = StrSQL & " where  ename=N'" & word & "'"
End If
Translatestr = ""

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                        If Translatetype = 1 Then 'arabic to english
                       ' Translatestr = Translatestr & " " & IIf(IsNull(rs("aname").value), "", rs("aname").value)
                            Translatestr = IIf(IsNull(rs("aname").value), "", rs("aname").value)
                        Else
                       ' Translatestr = Translatestr & " " & IIf(IsNull(rs("Ename").value), "", rs("Ename").value)
                            Translatestr = IIf(IsNull(rs("Ename").value), "", rs("Ename").value)
                        End If
                        rs.MoveNext
             Next i
    Else
    Translatestr = ""
        
    End If
Translate = Translatestr
End Function
Public Function GetItemIDFromCode(Optional ItemIDCode As String, _
                                      Optional ByRef Item_ID As Integer, _
                                      Optional Item_id1 As Integer = 0, _
                                      Optional ByRef ItemCode1 As String)
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    If Emp_id1 <> 0 Then
        sql = "select * from TblItems where ItemID= " & Item_ID
    Else
 
        sql = "select * from TblItems where  fullcode ='" & ItemIDCode & "'"
    End If
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Item_ID = IIf(IsNull(rs("ItemID").value), 0, rs("ItemID").value)
        ItemCode1 = IIf(IsNull(rs("fullcode").value), 0, rs("fullcode").value)
 
    Else
        Item_ID = 0
    End If

    rs.Close

End Function
Public Function CheckManulaNoForTransaction(Optional Transaction_Type As Integer, _
                                      Optional CusID As Double, Optional ManualNO As String, Optional Transaction_ID As Double, Optional ByRef NoteSerial1 As String) As Boolean
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
     If ManualNO = "" Then
         CheckManulaNoForTransaction = False
         Exit Function
      End If
      
      
        sql = "SELECT     NoteSerial1,Transaction_ID, Transaction_Type, CusID, ManualNO"
sql = sql & " From dbo.transactions"
sql = sql & "  Where (Transaction_Type = " & Transaction_Type & ") And (CusID = " & CusID & ")"
      sql = sql & " And Transaction_ID <> " & Transaction_ID
      
      If ManualNO = "" Then
      sql = sql & " And ManualNO ='" & ManualNO & "'"
      End If
      
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
      CheckManulaNoForTransaction = True
      NoteSerial1 = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
      MsgBox "    «·—ﬁ„ «·ÌœÊÌ  „ﬂ—— ›Ì ›« Ê—Â —ﬁ„ " & NoteSerial1, vbInformation
      
    Else
        CheckManulaNoForTransaction = False
      NoteSerial1 = "" ' IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
        
    End If

    rs.Close

End Function

Public Function updateemployeeEmbarkation(ID As Integer)
    Dim workdate  As Date
    Dim workdateH As String
 
    Dim Emp_id    As Integer
    Dim StrSQL    As String
    Dim rs        As New ADODB.Recordset
    StrSQL = "SELECT     *   from dbo.TblEmbarkation WHERE     (id = " & ID & ")"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_id = IIf(IsNull(rs("Emp_ID").value), 0, rs("Emp_ID").value)
        workdate = IIf(IsNull(rs("workdate").value), Date, rs("workdate").value)
        workdateH = IIf(IsNull(rs("workdateh").value), "", rs("workdateh").value)
    Else
        Exit Function
    End If
    StrSQL = "update TblEmployee  set  workstate=1 ,jopstatusid=1 , LastDate=" & SQLDate(workdate, True) & ", LastDateH='" & workdateH & "'"
    StrSQL = StrSQL & " where Emp_ID =" & Emp_id
    Cn.Execute StrSQL

End Function
Function GetTermsTotals(oprid As Double, _
                        Optional ID As Double = 0, _
                        Optional bill_date As Date, _
                        Optional ByRef LineFinalWithoutVat As Double, _
                        Optional ByRef quntExc As Double, _
                        Optional Flag As String = "N", _
                        Optional customerOrSub As Integer, _
                        Optional subcontid As Double, _
                        Optional VatPercent As Double, _
                        Optional oldPerforValue As Double, _
                        Optional discountHasmyat As Double, _
                        Optional linenetaftermainDiscountWithvat As Double, Optional ProjectID As Integer = 0)
 
    If oprid = 0 Then
        GetTermsTotals = 0
        Exit Function
    End If
    Dim StrSQL As String
    Dim rs     As New ADODB.Recordset
    StrSQL = "SELECT    ( SUM(dbo.project_bill_details.linenetaftermainDiscountWithvat + dbo.project_bill_details.PerforVLineDiscount )) AS LineFinal, SUM(dbo.project_bill_details.quntExc) AS quntExc  ,sum(linenetaftermainDiscountBeforevat) as LineFinalWithoutVat   ,sum(PerforVLineDiscount) as oldPerforValue  "
    StrSQL = StrSQL & "  ,sum(linenetaftermainDiscountWithvat) as linenetaftermainDiscountWithvat"
    
  '  StrSQL = StrSQL & "  ,QtyOpen = (Select  sum(QtyOpen) from projects_des where projects_des.oprid = project_bill_details.oprid  and projects_des.project_id = project_billl.project_No)"
    StrSQL = StrSQL & "                FROM         dbo.project_bill_details INNER JOIN"
    StrSQL = StrSQL & "                                      dbo.project_billl ON dbo.project_bill_details.bill_id = dbo.project_billl.id"
                      
    StrSQL = StrSQL & "  Where ( 1=1)"
    StrSQL = StrSQL & "   AND (dbo.project_bill_details.oprid = " & oprid & ")"
    StrSQL = StrSQL & "   AND (dbo.project_billl.bill_date <= " & SQLDate(bill_date, True) & ")"

    If ID <> 0 Then
        'If Flag = "E" Then
        'StrSQL = StrSQL & "  AND (dbo.project_billl.id < " & ID & ")"
        'Else
        StrSQL = StrSQL & "  AND (dbo.project_billl.id <" & ID & ")"
        'End If
    End If

    If customerOrSub = 1 Then '„ﬁ«Ê·
        StrSQL = StrSQL & "  AND bill_to=1 and subContractorId= " & subcontid
    Else
        StrSQL = StrSQL & "  AND bill_to=0"
    End If

    StrSQL = StrSQL & "  GROUP BY dbo.project_bill_details.oprid"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    
    Dim rsDummy As New ADODB.Recordset
    Dim s As String
    s = "select  sum(QtyOpen) QtyOpen from projects_des where projects_des.oprid =" & oprid & "  and projects_des.project_id = " & ProjectID
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rsDummy.EOF Then
        quntExc = IIf(IsNull(rsDummy("QtyOpen").value), 0, rsDummy("QtyOpen").value)
    End If
    Debug.Print StrSQL
    If rs.RecordCount > 0 Then
        quntExc = IIf(IsNull(rs("quntExc").value), 0, rs("quntExc").value) + quntExc
        
        LineFinalWithoutVat = IIf(IsNull(rs("LineFinalWithoutVat").value), 0, rs("LineFinalWithoutVat").value) ' «·«Ã„«·Ì »œÊ‰ «·÷—Ì»Â «Ê »œÊ‰ Œ’„ Õ”‰ «·«œ«¡
        linenetaftermainDiscountWithvat = IIf(IsNull(rs("linenetaftermainDiscountWithvat").value), 0, rs("linenetaftermainDiscountWithvat").value) '' «·«Ã„«·Ì ‘«„·  «·÷—Ì»Â »œÊ‰  Œ’„ Õ”‰ «·«œ«¡
        oldPerforValue = IIf(IsNull(rs("oldPerforValue").value), 0, rs("oldPerforValue").value) ' «Ã„«·Ì Œ’„ Õ”‰ «·«œ«¡
         
        'discountHasmyat = IIf(IsNull(rs("discountHasmyat").value), 0, rs("discountHasmyat").value)
        'VatPercent = IIf(IsNull(rs("Vat").value), 0, rs("Vat").value)
  
    Else
        LineFinalWithoutVat = 0
        linenetaftermainDiscountWithvat = 0
        oldPerforValue = 0
     
        Exit Function
    End If
End Function

Function GetDefaultGoldPrice() As Double
 
    Dim StrSQL As String
    Dim rs     As New ADODB.Recordset
    StrSQL = "SELECT      AvPrice  From dbo.TblAveragGrm  Where (Defult = 1) "

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetDefaultGoldPrice = IIf(IsNull(rs("AvPrice").value), 0, rs("AvPrice").value)
  
    Else
        GetDefaultGoldPrice = 0
        Exit Function
    End If

End Function

Public Function updateemployeeEmbarkation1(ID As Integer)
    Dim FromDate As Date
 
    Dim Emp_id   As Integer
    Dim StrSQL   As String
    Dim rs       As New ADODB.Recordset
    StrSQL = "SELECT     *   from dbo.TblVocation WHERE     (id = " & ID & ")"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_id = IIf(IsNull(rs("EmpID").value), 0, rs("EmpID").value)
        FromDate = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
        fromdateH = IIf(IsNull(rs("FromDateH").value), "", rs("FromDateH").value)
    Else
        Exit Function
    End If
    StrSQL = "update TblEmployee  set  workstate=2 ,jopstatusid=2 , lastHolidaydate=" & SQLDate(FromDate, True) & ", lastHolidaydateH='" & fromdateH & "'"
    StrSQL = StrSQL & " where Emp_ID =" & Emp_id
    Cn.Execute StrSQL

End Function

Public Function updateemployeeEmbarkation2(ID As Integer)
    Dim workdate  As Date
    Dim workdateH As String
 
    Dim Emp_id    As Integer
    Dim StrSQL    As String
    
    Dim ToDepart  As Integer
    Dim ProjectTo As Integer
    Dim JobTo     As Integer

    Dim rs        As New ADODB.Recordset
    StrSQL = "SELECT     *   from dbo.TblMoveEmp1 WHERE     (id = " & ID & ")"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_id = IIf(IsNull(rs("EmpID").value), 0, rs("EmpID").value)
        ToDepart = IIf(IsNull(rs("ToDepart").value), 0, rs("ToDepart").value)
        ProjectTo = IIf(IsNull(rs("ProjectTo").value), 0, rs("ProjectTo").value)
        JobTo = IIf(IsNull(rs("JobTo").value), 0, rs("JobTo").value)
    Else
        Exit Function
    End If
    StrSQL = "update TblEmployee  set   DepartmentID=" & ToDepart & " , GroupID=" & ProjectTo & ",JobTypeID=" & JobTo
    StrSQL = StrSQL & " where Emp_ID =" & Emp_id
    Cn.Execute StrSQL

End Function

Function ItemBigUnit(ItemID As Long) As Double
    Dim openingdate As Date
    Dim sql As String
    Dim rs As New ADODB.Recordset
    StrSQL = "SELECT     MAX(UnitFactor) AS BigUnit  from dbo.TblItemsUnits WHERE     (ItemID = " & ItemID & ")"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        ItemBigUnit = IIf(IsNull(rs("BigUnit").value), 0, rs("BigUnit").value)
    Else
        ItemBigUnit = 0
    End If

End Function

Public Sub ShowResultM(val As String)

    Select Case val

        Case "1": MsgBox ("·ﬁœ  „  «·⁄„·Ì… »‰Ã«Õ") 'sent

        Case "2": MsgBox ("≈‰ —’Ìœﬂ ·œÏ „Ê»«Ì·Ì ﬁœ ≈‰ ÂÏ Ê·„ Ì⁄œ »Â √Ì —”«∆·. (·Õ· «·„‘ﬂ·… ﬁ„ »‘Õ‰ —’Ìœﬂ „‰ «·—”«∆· ·œÏ „Ê»«Ì·Ì. ·‘Õ‰ —’Ìœﬂ ≈ »⁄  ⁄·Ì„«  ‘Õ‰ «·—’Ìœ)") 'your balance = 0

        Case "3": MsgBox ("≈‰ —’Ìœﬂ «·Õ«·Ì ·« Ìﬂ›Ì ·≈ „«„ ⁄„·Ì… «·≈—”«·. (·Õ· «·„‘ﬂ·… ﬁ„ »‘Õ‰ —’Ìœﬂ „‰ «·—”«∆· ·œÏ „Ê»«Ì·Ì. ·‘Õ‰ —’Ìœﬂ ≈ »⁄  ⁄·Ì„«  ‘Õ‰ «·—’Ìœ)") 'your balance  not  enough"

        Case "4": MsgBox ("≈‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ··œŒÊ· ≈·Ï Õ”«» «·—”«∆· €Ì— ’ÕÌÕ ( √ﬂœ „‰ √‰ ≈”„ «·„” Œœ„ «·–Ì ≈” Œœ„ Â ÂÊ ‰›”Â «·–Ì  ” Œœ„Â ⁄‰œ œŒÊ·ﬂ ≈·Ï „Êﬁ⁄ „Ê»«Ì·Ì)") 'mobile not found

        Case "5": MsgBox ("Â‰«ﬂ Œÿ√ ›Ì ﬂ·„… «·„—Ê— ( √ﬂœ „‰ √‰ ﬂ·„… «·„—Ê— «· Ì  „ ≈” Œœ«„Â« ÂÌ ‰›”Â« «· Ì  ” Œœ„Â« ⁄‰œ œŒÊ·ﬂ „Êﬁ⁄ „Ê»«Ì·Ì,≈–« ‰”Ì  ﬂ·„… «·„—Ê— ≈÷€ÿ ⁄·Ï —«»ÿ ‰”Ì  ﬂ·„… «·„—Ê— · ’·ﬂ —”«·… ⁄·Ï ÃÊ«·ﬂ »—ﬁ„ «·„—Ê— «·Œ«’ »ﬂ)") 'password error

        Case "6": MsgBox ("≈‰ ’›Õ… «·≈—”«· ·« ÃÌ» ›Ì «·Êﬁ  «·Õ«·Ì (ﬁœ ÌﬂÊ‰ Â‰«ﬂ ÿ·» ﬂ»Ì— ⁄·Ï «·’›Õ… √Ê  Êﬁ› „ƒﬁ  ··’›Õ… ›ﬁÿ Õ«Ê· „—… √Œ—Ï √Ê  Ê«’· „⁄ «·œ⁄„ «·›‰Ì ≈–« ≈” „— «·Œÿ√)") 'page not response try send again

        Case "12": MsgBox ("≈‰ Õ”«»ﬂ »Õ«Ã… ≈·Ï  ÕœÌÀ Ì—ÃÏ „—«Ã⁄… «·œ⁄„ «·›‰Ì")

        Case "13": MsgBox ("≈‰ ≈”„ «·„—”· «·–Ì ≈” Œœ„ Â ›Ì Â–Â «·—”«·… ·„ Ì „ ﬁ»Ê·Â. (Ì—ÃÏ ≈—”«· «·—”«·… »≈”„ „—”· ¬Œ— √Ê  ⁄—Ì› ≈”„ «·„—”· ·œÏ „Ê»«Ì·Ì)") 'sender not accept

        Case "14": MsgBox "≈‰ ≈”„ «·„—”· «·–Ì ≈” Œœ„ Â €Ì— „⁄—› ·œÏ „Ê»«Ì·Ì. (Ì„ﬂ‰ﬂ  ⁄—Ì› ≈”„ «·„—”· „‰ Œ·«· ’›Õ… ≈÷«›… ≈”„ „—”·)" 'sender name not activated

        Case "15": MsgBox "ÌÊÃœ —ﬁ„ ÃÊ«· Œ«ÿ∆ ›Ì «·√—ﬁ«„ «· Ì ﬁ„  »«·≈—”«· ·Â«. ( √ﬂœ „‰ ’Õ… «·√—ﬁ«„ «· Ì  —Ìœ «·≈—”«· ·Â« Ê√‰Â« »«·’Ì€… «·œÊ·Ì…)"

        Case "16": MsgBox "«·—”«·… «· Ì ﬁ„  »≈—”«·Â« ·«  Õ ÊÌ ⁄·Ï ≈”„ „—”·. (√œŒ· ≈”„ „—”· ⁄‰œ ≈—”«·ﬂ «·—”«·…)"

        Case "17": MsgBox "·„ Ì „ «—”«· ‰’ «·—”«·…. «·—Ã«¡ «· √ﬂœ „‰ «—”«· ‰’ «·—”«·… Ê«· √ﬂœ „‰  ÕÊÌ· «·—”«·… «·Ï ÌÊ‰Ì ﬂÊœ («·—Ã«¡ «· √ﬂœ „‰ «” Œœ«„ «·œ«·… ConvertToUnicode)"

        Case "-1": MsgBox "·„ Ì „ «· Ê«’· „⁄ Œ«œ„ (Server) «·≈—”«· „Ê»«Ì·Ì »‰Ã«Õ. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"

        Case "-2": MsgBox "·„ Ì „ «·—»ÿ „⁄ ﬁ«⁄œ… «·»Ì«‰«  (Database) «· Ì  Õ ÊÌ ⁄·Ï Õ”«»ﬂ Ê»Ì«‰« ﬂ ·œÏ „Ê»«Ì·Ì. (ﬁœ ÌﬂÊ‰ Â‰«ﬂ „Õ«Ê·«  ≈—”«· ﬂÀÌ—…  „  „⁄« , √Ê ﬁœ ÌﬂÊ‰ Â‰«ﬂ ⁄ÿ· „ƒﬁ  ÿ—√ ⁄·Ï «·Œ«œ„ ≈–« ≈” „—  «·„‘ﬂ·… Ì—ÃÏ «· Ê«’· „⁄ «·œ⁄„ «·›‰Ì)"
    
        Case Else: MsgBox (val)
    End Select

End Sub

Public Function sendMessageM(UserName As String, _
                             Password As String, _
                             Msg As String, _
                             sender As String, _
                             Numbers As String) As String
  '  On Error Resume Next
    Dim s As String
'    Dim Inet1 As InternetExplorer
 If Numbers = "" Then
    Exit Function
 End If
    UserName = SystemOptions.SMSUserName
    Password = (SystemOptions.SMSPassWord)
    sender = SystemOptions.SenderName
  '  msg = "1"
    
    If SystemOptions.OPTWEB = 0 Then
    Msg = ConvertToUnicode(Msg)
    Password = URLEncode2(Password)
    's = "http://www.mobily.ws/api/msgSend.php?mobile=" & UserName & "&password=" & Password & "&numbers=" & Numbers & "&sender=" & sender & "&msg=" & Msg & "&applicationType=24"
   s = "http://alfa-cell.com/api/msgSend.php?mobile=" & UserName & "&password=" & Password & "&numbers=" & Numbers & "&sender=" & sender & "&msg=" & Msg & "&applicationType=72"
   sendMessageM = WebRequest(s)
    ElseIf SystemOptions.OPTWEB = 1 Then
    s = "http://elec.sa/sms/api/sendsms.php?username=" & UserName & "&password=" & Password & "&message=" & Msg & "&numbers=" & Numbers & "&sender=" & sender & "&unicode=UTF-8&return=string"
    sendMessageM = WebRequest(s)
     ElseIf SystemOptions.OPTWEB = 2 Then
 
s = "http://www.jawalbsms.ws/api.php/sendsms?user=" & UserName & "&pass=" & Password & "&to=" & Numbers & "&message= " & Msg & " &sender=" & sender & ""
sendMessageM = WebRequest(s)
ElseIf SystemOptions.OPTWEB = 3 Then

 s = "https://apps.gateway.sa/vendorsms/pushsms.aspx?user=" & UserName & "&password=" & Password & "&msisdn=" & Numbers & "&sid=" & sender & "&msg" & Msg & "&fl=0"
 sendMessageM = WebRequest(s)
'Me.WbHelp.Navigate s
'sendMessageM = WebRequest(s)
ElseIf SystemOptions.OPTWEB = 4 Then
    'Msg = ConvertToUnicode(Msg)
s = "http://www.hisms.ws/api.php?send_sms&username=" & UserName & "&password=" & Password & "&numbers=" & Numbers & "&sender=" & sender & "&message=" & (Msg)
's = URLEncode(s)
 sendMessageM = WebRequestPHP(s)

    End If
    
  'sendMessageM = mdifrmmain.Inet1.OpenURL(s)'SMS
    
    
End Function
 
Public Function CreateReportForProduction(FromDate As Date, _
                                          ToDate As Date)
    Dim sql As String
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    StrSQL = " SELECT   MaterialIssueVoucherTotals,AccDep,  AccountIntervalID, StartDate, EndDate, SalaryVouchersTotals, Expenses, Allocations, Allocations1, AdvancedPayments, total, SaleValue, SaLePayValue, "
    StrSQL = StrSQL & "  ServicesValue , profit"
    StrSQL = StrSQL & "  from dbo.TblAccountIntervals"
    StrSQL = StrSQL + "  WHERE  StartDate >=" & SQLDate(FromDate, True) & ""
    StrSQL = StrSQL + " and EndDate <=" & SQLDate(ToDate, True) & ""
 
    Set rs = New ADODB.Recordset
    '  rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
               
        For i = 1 To rs.RecordCount
            'rs("EndDate").Value
            ' rs("MaterialIssueVoucherTotals").value=MaterialIssueVoucherTotals
            rs("MaterialIssueVoucherTotals").value = Round(GetISSueVoucherForProductionValue(rs("StartDate").value, rs("EndDate").value, 0), 2)  '„Ê«œ Œ«„
            rs("AccDep").value = Round(gettotal(90, rs("StartDate").value, rs("EndDate").value), 2) '”‰œ«  «·«Â·«ﬂ
            rs("Expenses").value = Round(GetExpensestotal(rs("StartDate").value, rs("EndDate").value), 2)
            rs("SalaryVouchersTotals").value = Round(GetNetsalaryVouchers(66, rs("StartDate").value, rs("EndDate").value), 2)
            rs("Allocations").value = Round(gettotal(8023, rs("StartDate").value, rs("EndDate").value, 0), 2)
            rs("Allocations1").value = Round(gettotal(8023, rs("StartDate").value, rs("EndDate").value, 1), 2)
            rs("AdvancedPayments").value = Round(gettotal(8027, rs("StartDate").value, rs("EndDate").value, -1), 2)
                                
            rs("total").value = Round(val(rs("MaterialIssueVoucherTotals").value) + val(rs("AccDep").value) + val(rs("SalaryVouchersTotals").value) + val(rs("Expenses").value) + val(rs("Allocations").value) + val(rs("Allocations1").value) + val(rs("AdvancedPayments").value), 2)
                                
            '  Txttotal.text = Round(Txttotal.text, 2)
            rs("SaleValue").value = Round(GetSalesCost(rs("StartDate").value, rs("EndDate").value), 2) '   ﬂ·›… «·„»Ì⁄« 
            rs("SaLePayValue").value = Round(GetSalesValue(rs("StartDate").value, rs("EndDate").value, 0), 2) 'ﬁÌ„… „»Ì⁄«  «·› —…
            rs("ServicesValue").value = Round(GetSalesValue(rs("StartDate").value, rs("EndDate").value, 1), 2)      'ﬁÌ„… Œœ„«  «·› —…
                                
            rs("Profit").value = val(rs("ServicesValue").value) + val(rs("SaLePayValue").value) - val(rs("SaleValue").value)
            rs.update
            rs.MoveNext
        Next i
             
    End If

End Function

 Public Function getLastCostPriceForItems(Item_ID As Long, _
                                         unit_id As Long) As Double
    Dim sql As String
    Dim rs As New ADODB.Recordset
  
    Dim RsUnitData As ADODB.Recordset
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim QtyBySmalltUnit As Double
        
    LngCurItemID = LngItemID
    LngUnitID = UnitID

    StrSQL = "Select * From TblItemsUnits Where ItemID=" & Item_ID
    StrSQL = StrSQL + " AND UnitID=" & unit_id
    Set RsUnitData = New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
        QtyBySmalltUnit = RsUnitData("UnitFactor").value
    Else
        QtyBySmalltUnit = 1
    End If
 
    sql = " SELECT     dbo.Transaction_Details.Price AS LasCosttPrice"
    sql = sql & "  FROM         dbo.Transaction_Details INNER JOIN"
    sql = sql & " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    sql = sql & " Where (dbo.Transactions.Transaction_Type = 19) And (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
    sql = sql & "  GROUP BY dbo.Transaction_Details.Price"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        getLastCostPriceForItems = IIf(IsNull(rs("LasCosttPrice").value), 0, rs("LasCosttPrice").value) * QtyBySmalltUnit
    Else
        getLastCostPriceForItems = 0
    End If

End Function


Public Function ComposMessage(ScreenName As String, _
                              Optional value As Double, _
                              Optional des As String, _
                              Optional Data As String, _
                              Optional ByRef opt As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim CurrentMessage As String
 
    sql = "SELECT     dbo.TblMessageDefDES.PlainMessageID, dbo.TblPlainMessage.Name, dbo.TblPlainMessage.Namee, dbo.TblMessageDef.opt, "
    sql = sql & " dbo.TblMessageDef.Screenname"
    sql = sql & " FROM         dbo.TblMessageDef INNER JOIN"
    sql = sql & "  dbo.TblMessageDefDES ON dbo.TblMessageDef.id = dbo.TblMessageDefDES.lMessageDefID INNER JOIN"
    sql = sql & "  dbo.TblPlainMessage ON dbo.TblMessageDefDES.PlainMessageID = dbo.TblPlainMessage.id"
    sql = sql & " WHERE     (dbo.TblMessageDef.ScreenName = '" & ScreenName & "')"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ComposMessage = ""

    If rs.RecordCount > 0 Then
        opt = IIf(IsNull(rs("opt").value), 0, rs("opt").value)

        For i = 1 To rs.RecordCount

            If SystemOptions.UserInterface = ArabicInterface Then
                CurrentMessage = IIf(IsNull(rs("Name").value), "", rs("Name").value)
            Else
                CurrentMessage = IIf(IsNull(rs("Namee").value), "", rs("Namee").value)
            End If

            CurrentMessagee = IIf(IsNull(rs("Namee").value), "", rs("Namee").value)

            If CurrentMessagee = "Value" Then
                CurrentMessage = value
            ElseIf CurrentMessagee = "DES" Then
                CurrentMessage = des
            ElseIf CurrentMessagee = "Data" Then
                CurrentMessage = Data
            End If

            ComposMessage = ComposMessage & " " & CurrentMessage
            rs.MoveNext
        Next i

    Else
        ComposMessage = ""
        opt = -1
    End If
 
End Function

Public Function CreateVacationData(EmpID As Integer) As String
    Dim sql As String
'    Exit Function
    Dim rs As New ADODB.Recordset
 Dim FirstOpeningDate  As Date
 Dim Diff As Integer
 getFirstPeriodDateInthisYear2 FirstOpeningDate
    sql = " delete tblVacationData where  Status1 is null  and InstVacaID is null   and EmpID=" & EmpID & " "
 Cn.Execute sql
 
Dim i As Integer
Dim StartDate As Data
Dim IssueDate As Date
Dim due_period As Integer
Dim Due_period_no As Integer


 
  
Dim Holiday_period_no As Integer
Dim Contract_date1 As Date
Dim Holiday_period As Integer
'Dim Holiday_period As Integer
Dim value As Double
Dim currentdate As Date
        get_employee_information EmpID, IssueDate, , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , Contract_date1, , due_period, Due_period_no, Holiday_period_no, Holiday_period
currentdate = Contract_date1

For i = 1 To 20
 

                If due_period = 0 Then
                      currentdate = DateAdd("M", Due_period_no, currentdate)
                ElseIf due_period = 1 Then
                    currentdate = DateAdd("YYYY", Due_period_no, currentdate)
                
               ElseIf due_period = 2 Then
                    currentdate = DateAdd("d", Due_period_no, currentdate)
                        
                        
                End If



    If Holiday_period = 0 Then
                     value = Holiday_period_no
                ElseIf Holiday_period = 1 Then
                    value = Holiday_period_no * 30
                
                   ElseIf Holiday_period = 2 Then
             '      value = Holiday_period_no * 360
                   
                   
      End If

Diff = DateDiff("d", FirstOpeningDate, currentdate)
If Diff >= 0 Then
  sql = "insert into tblVacationData (EmpID,ExpectedacationDate,ExpectedacationDateH,value)  values (" & EmpID & "," & SQLDate(currentdate, True) & ",'" & ToHijriDate(currentdate) & "'," & value & ") "
  
  End If
  
Cn.Execute sql
Next i
 
     



End Function

Public Function GetCustomerEmail(CusID As Integer, Optional ByRef customername As String) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = " SELECT   * from TblCustemers  "
 
    sql = sql & " Where (CusID = " & CusID & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                        customername = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                Else
                         customername = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                End If
    
    
        GetCustomerEmail = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
    Else
    customername = ""
        GetCustomerEmail = ""
    End If

End Function
Public Function getownerId(Optional Aqarid As Variant = 0)
If Aqarid <> 0 Then
Dim Rs9  As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
Dim sql As String
Aqarid = val(Aqarid)
sql = "select * from tblaqar where Aqarid =" & Aqarid & ""
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
getownerId = IIf(IsNull(Rs9("ownerid").value), 0, Rs9("ownerid").value)
  End If
  End If
End Function


Public Function GetEmployeeNumber(Emp_id As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = " SELECT   * from TblEmployee  "
 
    sql = sql & " Where   (Emp_ID = " & Emp_id & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetEmployeeNumber = IIf(IsNull(rs("Emp_mobile").value), "", rs("Emp_mobile").value)
    Else
        GetEmployeeNumber = ""
    End If

End Function


Public Function GetEmployeeIDFROMUserID(UserID As Long) As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = " SELECT   * from TblUsers where UserID =" & UserID
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetEmployeeIDFROMUserID = IIf(IsNull(rs("Empid").value), 0, rs("Empid").value)
    Else
        GetEmployeeIDFROMUserID = 0
    End If

End Function

Public Function GetCustomerNumber(CusID As Variant) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
 CusID = val(CusID)
    sql = " SELECT   * from TblCustemers  "
 
    sql = sql & " Where sendMessage=1 and  (CusID = " & CusID & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetCustomerNumber = IIf(IsNull(rs("Cus_mobile").value), "", rs("Cus_mobile").value)
    Else
        GetCustomerNumber = ""
    End If

End Function

Public Function getStoreCoding(StoreID As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = " SELECT   isnull(Code,'0') Code "
    sql = sql & " from   dbo.TblStore"
    sql = sql & " WHERE     (StoreID = " & StoreID & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        getStoreCoding = rs!code & "" 'IIf(IsNull(rs("Code").value), 0, rs("Code").value)
    Else
        getStoreCoding = ""
    End If
rs.Close
Set rs = Nothing
End Function

Public Function getStoreInformatin(code As String) As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = " SELECT   * "
    sql = sql & " from   dbo.TblStore"
    sql = sql & " WHERE     (Code ='" & code & "')"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        getStoreInformatin = IIf(IsNull(rs("StoreID").value), 0, rs("StoreID").value)
    Else
        getStoreInformatin = 0
    End If
rs.Close
Set rs = Nothing
End Function


Public Function CheckStoreCoding(branch_id As Integer, Sanad_No As Integer) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = " SELECT   * "
    sql = sql & " from   dbo.sanad_numbering"
    sql = sql & " WHERE     (sanad_no = " & Sanad_No & ") AND (branch_no = " & branch_id & ") "
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        CheckStoreCoding = IIf(IsNull(rs("StoreCoding").value), False, rs("StoreCoding").value)
    Else
        CheckStoreCoding = False
    End If
rs.Close
Set rs = Nothing
End Function



Public Function getBranchCurrentAccount(branch_id As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = " SELECT   * "
    sql = sql & " from dbo.TblBranchesData"
    sql = sql & " Where (branch_id = " & branch_id & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        getBranchCurrentAccount = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
    Else
        getBranchCurrentAccount = ""
    End If

End Function


Public Function getTransactionIdnoteserial1(NoteSerial1 As String, TransType As Integer) As Double
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = " SELECT    Transaction_ID "
    sql = sql & " from dbo.Transactions"
    sql = sql & " Where (Transaction_Type = " & TransType & ") And (noteserial1 ='" & NoteSerial1 & "')"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        getTransactionIdnoteserial1 = IIf(IsNull(rs("Transaction_ID").value), -1, rs("Transaction_ID").value)
         
          
    Else
        getTransactionIdnoteserial1 = -1
    End If

End Function


Public Function getTransactionIdBytable(STableID As Double) As Double
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = " SELECT     max(Transaction_ID) as Transaction_ID "
    sql = sql & " from dbo.Transactions"
    sql = sql & " Where (STableID = " & STableID & ") And (Printed Is Null)"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        getTransactionIdBytable = IIf(IsNull(rs("Transaction_ID").value), -1, rs("Transaction_ID").value)
         
         If getTransactionIdBytable = -1 Then
         Cn.Execute "update Stables  set status=null where id =" & STableID
         End If
    Else
        getTransactionIdBytable = -1
    End If

End Function

Public Function getNoOfBranches() As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = "SELECT count(branch_id)as NoOfBranches From TblBranchesData"
   
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        getNoOfBranches = IIf(IsNull(rs("NoOfBranches").value), 0, rs("NoOfBranches").value)
    Else
        getNoOfBranches = 0
    End If

End Function

Public Function GetShiftId(LoginTime As String, _
                           Optional ByRef SeftCode As String, _
                           Optional ByRef fomshift1 As Date, _
                           Optional ByRef ToDate1 As Date)
    '17 11 2012
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim lTime As Date
    sql = "select * from TbLSheft  "
    'SeftCode
    'SheftName
    'ShiftFrom   ShiftTo
    '    SeftCode = 1
    '                    SheftName = 1
    '              Exit Function
    'Ltime = FormatDateTime(LoginTime, vbShortTime)
    Dim ToTime As String

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
           
            ToTime = FormatDateTime(rs("ShiftTo").value, vbShortTime)

            If mId(ToTime, 1, 2) = "00" Then
                ToTime = "23" & mId(ToTime, 3, 3)
            End If

            '  MsgBox DateDiff("N", FormatDateTime(LoginTime, vbShortTime), FormatDateTime(rs("ShiftTo").value, vbShortTime))
            If FormatDateTime(LoginTime, vbShortTime) >= FormatDateTime(rs("ShiftFrom").value, vbShortTime) And FormatDateTime(LoginTime, vbShortTime) < FormatDateTime(CDate(ToTime), vbShortTime) Then
                '   If FormatDateTime(LoginTime, vbShortTime) <= FormatDateTime(rs("ShiftFrom").value, vbShortTime) Then
                 
                SeftCode = IIf(IsNull(rs("SeftCode").value), 0, rs("SeftCode").value)
                SheftName = IIf(IsNull(rs("SheftName").value), 0, rs("SheftName").value)

                If Not IsNull(rs("ShiftFrom").value) Then
                    fomshift1 = FormatDateTime(rs("ShiftFrom").value, vbShortTime)
                         
                End If

                If Not IsNull(rs("ShiftTo").value) Then
                    ToDate1 = FormatDateTime(rs("ShiftTo").value, vbShortTime)
                End If

                Exit For
            End If
             
            rs.MoveNext
        Next i
  
    Else
        SeftCode = 0
        SheftName = ""
    End If

    If SeftCode = "" Then SeftCode = 0
    rs.Close
End Function

Public Function GetCashierdatax(ID As Integer, _
                           Optional ByRef BoxID As Integer, _
                           Optional ByRef PointID As Integer)
    '17 11 2012
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim lTime As Date
    sql = " SELECT     dbo.cachierData.id, dbo.cachierData.name, dbo.cachierData.username, dbo.cachierData.password, dbo.cachierData.point_name, dbo.cachierData.EmpID, "
sql = sql & "  dbo.cachierData.PointID , dbo.cachierData.Ctype, dbo.cachierData.namee, dbo.cachierData.BoxID, dbo.TblBoxesData.Account_Code"
sql = sql & " FROM         dbo.cachierData INNER JOIN"
sql = sql & "  dbo.TblBoxesData ON dbo.cachierData.BoxID = dbo.TblBoxesData.BoxID"
sql = sql & "  Where (dbo.cachierData.id = " & ID & ")"

 
    Dim ToTime As String

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
                BoxID = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
                PointID = IIf(IsNull(rs("PointID").value), 0, rs("PointID").value)
  
    Else
BoxID = 0
PointID = 0
    End If

    If SeftCode = "" Then SeftCode = 0
    rs.Close
End Function
Public Function updatePaymenttransactions(value As Integer, BoxID As Integer, FromDate As Date, ToDate As Date)
Dim StrSQL As String

If value = 1 Then
StrSQL = "  update TblTransactionPayments Set Collected = 1"
 StrSQL = StrSQL & "   WHERE (PaymentID <> 0) AND  (boxid = " & BoxID & ")"
 StrSQL = StrSQL & " AND  (dbo.TblTransactionPayments.Recorddate >='" & SQLDate(FromDate) & "'"
 StrSQL = StrSQL & " AND   dbo.TblTransactionPayments.Recorddate <='" & SQLDate(FromDate) & "')"
 Else
 StrSQL = "  update TblTransactionPayments Set Collected = NULL"
 StrSQL = StrSQL & "   WHERE (PaymentID <> 0) AND  (boxid = " & BoxID & ")"
 StrSQL = StrSQL & " AND  (dbo.TblTransactionPayments.Recorddate >='" & SQLDate(FromDate) & "'"
 StrSQL = StrSQL & " AND   dbo.TblTransactionPayments.Recorddate <='" & SQLDate(FromDate) & "')"

 
 End If
 Cn.Execute StrSQL
 
End Function
Public Function GetPointSAles1Old(BoxID As Integer, PaymentID As Integer, FromDate As Date, ToDate As Date, Optional Emp_id As Double, Optional ByVal mIsReturn As Integer = 0, Optional ByVal mIstaxFree As Integer = 0) As Double
  Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim lTime As Date
    sql = " SELECT     TOP 100 PERCENT SUM(dbo.TblTransactionPayments.[value]) AS totals"
sql = sql & " FROM         dbo.TblTransactionPayments INNER JOIN"
sql = sql & "  dbo.cachierData ON dbo.TblTransactionPayments.CurrentCashireID = dbo.cachierData.id INNER JOIN"
sql = sql & "   dbo.TblBoxesData ON dbo.cachierData.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
sql = sql & "   dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"

'***********new
sql = " SELECT        SUM([value]) AS totals"
sql = sql & "   From dbo.TblTransactionPayments"


'sql = sql & " Where (dbo.TblTransactionPayments.locked=1) And (dbo.TblTransactionPayments.BoxId = " & BoxID & ")"
'salimhere
sql = sql & " Where (dbo.TblTransactionPayments.BoxId = " & BoxID & ")"

sql = sql & " AND  (dbo.TblTransactionPayments.PaymentID=" & PaymentID & ")"
 
 sql = sql & " AND  (dbo.TblTransactionPayments.Recorddate >='" & SQLDate(FromDate) & "'"
 sql = sql & " AND   dbo.TblTransactionPayments.Recorddate <='" & SQLDate(ToDate) & "')"
 

 
 
 'salimhere
   'new salim
 
 If mIstaxFree = 1 Then
     If mIsReturn = 0 Then
             My_SQL = "SELECT     TOP 100 PERCENT SUM("
            My_SQL = My_SQL & " ISNULL(dbo.transactions.TotalTaxExempt,0) *"
            
            My_SQL = My_SQL & "  isnull( dbo.TblTransactionPayments.Effect,"
            My_SQL = My_SQL & "  case"
            
            My_SQL = My_SQL & "  when Transaction_Type=21 then 1"
            My_SQL = My_SQL & "  else  -1"
            My_SQL = My_SQL & "  End"
            
            My_SQL = My_SQL & "  )"
    Else
             My_SQL = "SELECT     TOP 100 PERCENT SUM("
            My_SQL = My_SQL & " ISNULL(transactions.TotalTaxExempt,0)) "
            
            
            My_SQL = My_SQL & "  "
    
    End If
    
    
'    If mIstaxFree = 1 Then
'        My_SQL = My_SQL & "  , dbo.transactions.TotalTaxExempt)) AS TotalValue,"
'    Else
'        My_SQL = My_SQL & "  , dbo.Transactions.Transaction_NetValue)) AS TotalValue,"
'    End If


 Else
     If mIsReturn = 0 Then
             My_SQL = "SELECT     TOP 100 PERCENT SUM("
            My_SQL = My_SQL & " ISNULL(dbo.TblTransactionPayments.[value],0)) *"
            
            My_SQL = My_SQL & "  isnull( dbo.TblTransactionPayments.Effect,"
            My_SQL = My_SQL & "  case"
            
            My_SQL = My_SQL & "  when Transaction_Type=21 then 1"
            My_SQL = My_SQL & "  else  -1"
            My_SQL = My_SQL & "  End"
            
            My_SQL = My_SQL & "  )),0-"
    Else
             My_SQL = "SELECT     TOP 100 PERCENT SUM("
            My_SQL = My_SQL & " ISNULL(dbo.TblTransactionPayments.[value],0)) "
            
            
            My_SQL = My_SQL & "  "
    
    End If
    
    
'    If mIstaxFree = 1 Then
'        My_SQL = My_SQL & "  , dbo.transactions.TotalTaxExempt)) AS TotalValue,"
'    Else
'        My_SQL = My_SQL & "  , dbo.Transactions.Transaction_NetValue)) AS TotalValue,"
'    End If

 End If


My_SQL = My_SQL & "  AS TotalValue, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId,"
My_SQL = My_SQL & "                       dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee,"
My_SQL = My_SQL & "                       dbo.TblPaymentType.branch_no, dbo.TblPaymentType.MaxValue, dbo.TblPaymentType.TypTran, dbo.Transactions.Emp_ID,"
My_SQL = My_SQL & "                       dbo.transactions.POSBillType,dbo.TblPaymentType.Accountsus"
My_SQL = My_SQL & " FROM         dbo.TblTransactionPayments RIGHT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.Transactions ON dbo.TblTransactionPayments.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                       dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
If mIstaxFree = 1 Then
'    My_SQL = My_SQL & "                   Inner join Transaction_Details On Transaction_Details.Transaction_ID = Transactions.Transaction_ID    "
End If
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Date = CONVERT(DATETIME, '2019-11-10 00:00:00', 102))"
 My_SQL = My_SQL & " WHERE 1=1 "
 My_SQL = My_SQL & "  AND  (Transactions.Transaction_Date >='" & SQLDate(FromDate) & "'"
 My_SQL = My_SQL & "  AND   Transactions.Transaction_Date <='" & SQLDate(ToDate) & "')"
 
My_SQL = My_SQL & " AND  ( isnull(dbo.TblTransactionPayments.PaymentID,0)=" & PaymentID & ")"

My_SQL = My_SQL & " and      (dbo.Transactions.POSBillType = 1 OR"
If mIsReturn = 0 Then
    My_SQL = My_SQL & "                       dbo.Transactions.POSBillType = 4) AND (dbo.Transactions.Transaction_Type = 21 OR"
    My_SQL = My_SQL & "                       dbo.Transactions.Transaction_Type = 9)"
ElseIf mIsReturn = 1 Then
    My_SQL = My_SQL & "                       dbo.Transactions.POSBillType = 4) AND (dbo.Transactions.Transaction_Type = 21 )"
ElseIf mIsReturn = 2 Then
    My_SQL = My_SQL & "                       dbo.Transactions.POSBillType = 4) AND (dbo.Transactions.Transaction_Type = 9 )"
End If
My_SQL = My_SQL & "AND (dbo.Transactions.Emp_ID = " & Emp_id & ")"

If mIstaxFree = 1 Then
   
    
'    My_SQL = My_SQL & "  and    (IsNull(chkTaxExempt,0) = 1"
'    My_SQL = My_SQL & "  Or IsNull(chkTaxExempt,0) = 0 and Transaction_Details.Item_Id In (Select tblItems.ItemID from tblItems where IsNull(tblItems.ItemWithOutVAT,0) <> 0 ))"

End If

My_SQL = My_SQL & " GROUP BY isnull(dbo.TblTransactionPayments.PaymentID,0), dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId,"
My_SQL = My_SQL & "                       dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee,"
My_SQL = My_SQL & "                       dbo.TblPaymentType.branch_no, dbo.TblPaymentType.MaxValue, dbo.TblPaymentType.TypTran, dbo.Transactions.Emp_ID,"
My_SQL = My_SQL & "                       dbo.transactions.POSBillType,dbo.TblPaymentType.Accountsus"

My_SQL = My_SQL & "  ORDER BY isnull(dbo.TblTransactionPayments.PaymentID,0)"


 
    Dim ToTime As String

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
           GetPointSAles1Old = IIf(IsNull(rs("TotalValue").value), 0, rs("TotalValue").value)
                 
  
    Else
    GetPointSAles1Old = 0
    End If

     
    rs.Close
    
End Function


Public Function GetPointSAles1(BoxID As Integer, PaymentID As Integer, FromDate As Date, ToDate As Date, Optional Emp_id As Double, Optional ByVal mIsReturn As Integer = 0, Optional ByVal mIstaxFree As Integer = 0) As Double
  Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim lTime As Date
    sql = " SELECT     TOP 100 PERCENT SUM(dbo.TblTransactionPayments.[value]) AS totals"
sql = sql & " FROM         dbo.TblTransactionPayments INNER JOIN"
sql = sql & "  dbo.cachierData ON dbo.TblTransactionPayments.CurrentCashireID = dbo.cachierData.id INNER JOIN"
sql = sql & "   dbo.TblBoxesData ON dbo.cachierData.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
sql = sql & "   dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"

'***********new
sql = " SELECT        SUM([value]) AS totals"
sql = sql & "   From dbo.TblTransactionPayments"


'sql = sql & " Where (dbo.TblTransactionPayments.locked=1) And (dbo.TblTransactionPayments.BoxId = " & BoxID & ")"
'salimhere
sql = sql & " Where (dbo.TblTransactionPayments.BoxId = " & BoxID & ")"

sql = sql & " AND  (dbo.TblTransactionPayments.PaymentID=" & PaymentID & ")"
 
 sql = sql & " AND  (dbo.TblTransactionPayments.Recorddate >='" & SQLDate(FromDate) & "'"
 sql = sql & " AND   dbo.TblTransactionPayments.Recorddate <='" & SQLDate(ToDate) & "')"
 

 
 
 'salimhere
   'new salim
 
 If mIstaxFree = 1 Then
     If mIsReturn = 0 Then
             My_SQL = "SELECT     TOP 100 PERCENT SUM("
            My_SQL = My_SQL & " ISNULL(dbo.transactions.TotalTaxExempt,0) *"
            
            My_SQL = My_SQL & "  isnull( dbo.TblTransactionPayments.Effect,"
            My_SQL = My_SQL & "  case"
            
            My_SQL = My_SQL & "  when Transaction_Type=21 then 1"
            My_SQL = My_SQL & "  else  -1"
            My_SQL = My_SQL & "  End"
            
            My_SQL = My_SQL & "  )"
    Else
             My_SQL = "SELECT     TOP 100 PERCENT SUM("
            My_SQL = My_SQL & " ISNULL(transactions.TotalTaxExempt,0)) "
            
            
            My_SQL = My_SQL & "  "
    
    End If
    
    
'    If mIstaxFree = 1 Then
'        My_SQL = My_SQL & "  , dbo.transactions.TotalTaxExempt)) AS TotalValue,"
'    Else
'        My_SQL = My_SQL & "  , dbo.Transactions.Transaction_NetValue)) AS TotalValue,"
'    End If


 Else
   
   
   If PaymentID = 0 Then
    My_SQL = ""
    My_SQL = My_SQL & "SELECT "
    My_SQL = My_SQL & "    ISNULL(CashInvoices.TotalNetCash, 0) + ISNULL(CashPayments.TotalPaidCash, 0) AS TotalValue, "
    My_SQL = My_SQL & "    ISNULL(pt.PaymentName, N'‰ﬁœÌ') AS PaymentName, "
    My_SQL = My_SQL & "    ISNULL(pt.BankId, 0) AS BankId, "
    My_SQL = My_SQL & "    ISNULL(pt.Accountsus, 0) AS Accountsus, "
    My_SQL = My_SQL & "    ISNULL(pt.Accountcom, 0) AS Accountcom, "
    My_SQL = My_SQL & "    ISNULL(pt.commision, 0) AS commision, "
    My_SQL = My_SQL & "    ISNULL(pt.PaymentNamee, N'‰ﬁœÌ') AS PaymentNamee, "
    My_SQL = My_SQL & "    ISNULL(pt.branch_no, 0) AS branch_no, "
    My_SQL = My_SQL & "    ISNULL(pt.MaxValue, 0) AS MaxValue, "
    My_SQL = My_SQL & "    ISNULL(pt.TypTran, 0) AS TypTran, "
    My_SQL = My_SQL & "    EmpData.Emp_ID, "
    My_SQL = My_SQL & "    EmpData.POSBillType, "
    My_SQL = My_SQL & "    ISNULL(pt.Accountsus, 0) AS Expr1 "
    My_SQL = My_SQL & "FROM "
    My_SQL = My_SQL & "   (SELECT Emp_ID, POSBillType FROM Transactions "
    My_SQL = My_SQL & "     WHERE Transaction_Date >= '" & SQLDate(FromDate) & "' "
    My_SQL = My_SQL & "       AND Transaction_Date <= '" & SQLDate(ToDate) & "' "
    My_SQL = My_SQL & "       AND Emp_ID = " & Emp_id & " "
    My_SQL = My_SQL & "       AND (POSBillType = 1 OR POSBillType = 4) "
    If mIsReturn = 0 Then
        My_SQL = My_SQL & "       AND (Transaction_Type = 21 OR Transaction_Type = 9) "
    ElseIf mIsReturn = 1 Then
        My_SQL = My_SQL & "       AND (Transaction_Type = 21) and isnull(PaymentType,0) <> 1 "
    ElseIf mIsReturn = 2 Then
        My_SQL = My_SQL & "       AND (Transaction_Type = 9) "
    End If
    My_SQL = My_SQL & "     GROUP BY Emp_ID, POSBillType "
    My_SQL = My_SQL & "   ) EmpData "
    My_SQL = My_SQL & "LEFT JOIN ( "
    My_SQL = My_SQL & "   SELECT SUM(Transaction_NetValue) AS TotalNetCash, Emp_ID, POSBillType "
    My_SQL = My_SQL & "   FROM Transactions "
    My_SQL = My_SQL & "   WHERE PaymentType = 0 "
    My_SQL = My_SQL & "     AND Transaction_Date >= '" & SQLDate(FromDate) & "' "
    My_SQL = My_SQL & "     AND Transaction_Date <= '" & SQLDate(ToDate) & "' "
    My_SQL = My_SQL & "     AND Emp_ID = " & Emp_id & " "
    My_SQL = My_SQL & "     AND (POSBillType = 1 OR POSBillType = 4) "
    If mIsReturn = 0 Then
        My_SQL = My_SQL & "     AND (Transaction_Type = 21 OR Transaction_Type = 9) "
    ElseIf mIsReturn = 1 Then
        My_SQL = My_SQL & "     AND (Transaction_Type = 21) and isnull(PaymentType,0) <> 1 "
    ElseIf mIsReturn = 2 Then
        My_SQL = My_SQL & "     AND (Transaction_Type = 9) "
    End If
    My_SQL = My_SQL & "   GROUP BY Emp_ID, POSBillType "
    My_SQL = My_SQL & ") CashInvoices "
    My_SQL = My_SQL & "   ON EmpData.Emp_ID = CashInvoices.Emp_ID AND EmpData.POSBillType = CashInvoices.POSBillType "
    My_SQL = My_SQL & "LEFT JOIN ( "
    My_SQL = My_SQL & "   SELECT SUM(tp.value) AS TotalPaidCash, t.Emp_ID, t.POSBillType "
    My_SQL = My_SQL & "   FROM TblTransactionPayments tp "
    My_SQL = My_SQL & "   INNER JOIN Transactions t ON tp.Transaction_ID = t.Transaction_ID "
    My_SQL = My_SQL & "   WHERE tp.PaymentID = 0 "
    My_SQL = My_SQL & "     AND t.Transaction_Date >= '" & SQLDate(FromDate) & "' "
    My_SQL = My_SQL & "     AND t.Transaction_Date <= '" & SQLDate(ToDate) & "' "
    My_SQL = My_SQL & "     AND t.Emp_ID = " & Emp_id & " "
    My_SQL = My_SQL & "     AND (t.POSBillType = 1 OR t.POSBillType = 4) "
    If mIsReturn = 0 Then
        My_SQL = My_SQL & "     AND (t.Transaction_Type = 21 OR t.Transaction_Type = 9) "
    ElseIf mIsReturn = 1 Then
        My_SQL = My_SQL & "     AND (t.Transaction_Type = 21) and isnull(t.PaymentType,0) <> 1 "
    ElseIf mIsReturn = 2 Then
        My_SQL = My_SQL & "     AND (t.Transaction_Type = 9) "
    End If
    My_SQL = My_SQL & "   GROUP BY t.Emp_ID, t.POSBillType "
    My_SQL = My_SQL & ") CashPayments "
    My_SQL = My_SQL & "   ON EmpData.Emp_ID = CashPayments.Emp_ID AND EmpData.POSBillType = CashPayments.POSBillType "
    My_SQL = My_SQL & "LEFT JOIN TblPaymentType pt ON pt.PaymentID = 0 "
    My_SQL = My_SQL & "ORDER BY EmpData.Emp_ID, EmpData.POSBillType "
Else
    My_SQL = ""
    My_SQL = My_SQL & "SELECT "
    My_SQL = My_SQL & "    ISNULL(CashPayments.TotalPaidCash, 0) AS TotalValue, "
    My_SQL = My_SQL & "    ISNULL(pt.PaymentName, N'œ›⁄Â') AS PaymentName, "
    My_SQL = My_SQL & "    ISNULL(pt.BankId, 0) AS BankId, "
    My_SQL = My_SQL & "    ISNULL(pt.Accountsus, 0) AS Accountsus, "
    My_SQL = My_SQL & "    ISNULL(pt.Accountcom, 0) AS Accountcom, "
    My_SQL = My_SQL & "    ISNULL(pt.commision, 0) AS commision, "
    My_SQL = My_SQL & "    ISNULL(pt.PaymentNamee, N'œ›⁄Â') AS PaymentNamee, "
    My_SQL = My_SQL & "    ISNULL(pt.branch_no, 0) AS branch_no, "
    My_SQL = My_SQL & "    ISNULL(pt.MaxValue, 0) AS MaxValue, "
    My_SQL = My_SQL & "    ISNULL(pt.TypTran, 0) AS TypTran, "
    My_SQL = My_SQL & "    EmpData.Emp_ID, "
    My_SQL = My_SQL & "    EmpData.POSBillType, "
    My_SQL = My_SQL & "    ISNULL(pt.Accountsus, 0) AS Expr1 "
    My_SQL = My_SQL & "FROM "
    My_SQL = My_SQL & "   (SELECT Emp_ID, POSBillType FROM Transactions "
    My_SQL = My_SQL & "     WHERE Transaction_Date >= '" & SQLDate(FromDate) & "' "
    My_SQL = My_SQL & "       AND Transaction_Date <= '" & SQLDate(ToDate) & "' "
    My_SQL = My_SQL & "       AND Emp_ID = " & Emp_id & " "
    My_SQL = My_SQL & "       AND (POSBillType = 1 OR POSBillType = 4) "
    If mIsReturn = 0 Then
        My_SQL = My_SQL & "       AND (Transaction_Type = 21 OR Transaction_Type = 9) "
    ElseIf mIsReturn = 1 Then
        My_SQL = My_SQL & "       AND (Transaction_Type = 21) and isnull(PaymentType,0) <> 1 "
    ElseIf mIsReturn = 2 Then
        My_SQL = My_SQL & "       AND (Transaction_Type = 9) "
    End If
    My_SQL = My_SQL & "     GROUP BY Emp_ID, POSBillType "
    My_SQL = My_SQL & "   ) EmpData "
    My_SQL = My_SQL & "LEFT JOIN ( "
    My_SQL = My_SQL & "   SELECT SUM(tp.value) AS TotalPaidCash, t.Emp_ID, t.POSBillType "
    My_SQL = My_SQL & "   FROM TblTransactionPayments tp "
    My_SQL = My_SQL & "   INNER JOIN Transactions t ON tp.Transaction_ID = t.Transaction_ID "
    My_SQL = My_SQL & "   WHERE tp.PaymentID = " & PaymentID & " "
    My_SQL = My_SQL & "     AND t.Transaction_Date >= '" & SQLDate(FromDate) & "' "
    My_SQL = My_SQL & "     AND t.Transaction_Date <= '" & SQLDate(ToDate) & "' "
    My_SQL = My_SQL & "     AND t.Emp_ID = " & Emp_id & " "
    My_SQL = My_SQL & "     AND (t.POSBillType = 1 OR t.POSBillType = 4) "
    If mIsReturn = 0 Then
        My_SQL = My_SQL & "     AND (t.Transaction_Type = 21 OR t.Transaction_Type = 9) "
    ElseIf mIsReturn = 1 Then
        My_SQL = My_SQL & "     AND (t.Transaction_Type = 21) and isnull(t.PaymentType,0) <> 1 "
    ElseIf mIsReturn = 2 Then
        My_SQL = My_SQL & "     AND (t.Transaction_Type = 9) "
    End If
    My_SQL = My_SQL & "   GROUP BY t.Emp_ID, t.POSBillType "
    My_SQL = My_SQL & ") CashPayments "
    My_SQL = My_SQL & "   ON EmpData.Emp_ID = CashPayments.Emp_ID AND EmpData.POSBillType = CashPayments.POSBillType "
    My_SQL = My_SQL & "LEFT JOIN TblPaymentType pt ON pt.PaymentID = " & PaymentID & " "
    My_SQL = My_SQL & "ORDER BY EmpData.Emp_ID, EmpData.POSBillType "
End If


End If


 
    Dim ToTime As String

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
           GetPointSAles1 = IIf(IsNull(rs("TotalValue").value), 0, rs("TotalValue").value)
                 
  
    Else
    GetPointSAles1 = 0
    End If

     
    rs.Close
    
End Function


Public Function GetPointSAles(Optional PointID As Integer, Optional PaymentID As Integer = -1, Optional effect As Integer = 1) As Double
  Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim lTime As Date
    GetPointSAles = 0
'    sql = " SELECT     TOP 100 PERCENT SUM(dbo.TblTransactionPayments.[value]) AS totals"
'sql = sql & " FROM         dbo.TblTransactionPayments INNER JOIN"
'sql = sql & "  dbo.cachierData ON dbo.TblTransactionPayments.CurrentCashireID = dbo.cachierData.id INNER JOIN"
'sql = sql & "   dbo.TblBoxesData ON dbo.cachierData.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
'sql = sql & "   dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
'sql = sql & " Where (dbo.TblTransactionPayments.locked Is Null) And (dbo.TblTransactionPayments.PointID = " & PointID & ")"

sql = "SELECT     TOP 100 PERCENT SUM([value] * ISNULL(Effect, 1)) AS totals"
sql = sql & " FROM         dbo.TblTransactionPayments  "
 sql = sql & " Where (dbo.TblTransactionPayments.locked Is Null or dbo.TblTransactionPayments.locked=0 ) And (dbo.TblTransactionPayments.PointID = " & PointID & ")"
If PaymentID <> -1 Then
sql = sql & " AND  (dbo.TblTransactionPayments.PaymentID=" & PaymentID & ")"
End If
 
 If effect = -1 Then
        sql = sql & " AND  (dbo.TblTransactionPayments.Effect=" & effect & ")"
 
 ElseIf effect = 1 Then '’«›Ì «·‰ﬁœÌ…
         sql = sql & " AND  (dbo.TblTransactionPayments.Effect= 1  or dbo.TblTransactionPayments.Effect is null )"
 
 ElseIf effect = 0 Then
 
  ElseIf effect = 2 Then
     sql = sql & " AND  (dbo.TblTransactionPayments.Effect= 1  or dbo.TblTransactionPayments.Effect is null )"
 Else
         sql = sql & " AND  (dbo.TblTransactionPayments.Effect= 1  or dbo.TblTransactionPayments.Effect is null )"
         
    End If
 
 
    Dim ToTime As String

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
                GetPointSAles = IIf(IsNull(rs("totals").value), 0, rs("totals").value)
                 
  
    Else
    GetPointSAles = 0
    End If

    
    rs.Close
    
End Function

Public Function GetShiftData(SeftCode As String, _
                           Optional ByRef ShiftFrom As Date, _
                           Optional ByRef ShiftTo As Date)
    '17 11 2012
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim lTime As Date
    sql = "select * from TbLSheft  where SeftCode=" & SeftCode
    'SeftCode
    'SheftName
    'ShiftFrom   ShiftTo
    '    SeftCode = 1
    '                    SheftName = 1
    '              Exit Function
    'Ltime = FormatDateTime(LoginTime, vbShortTime)
    Dim ToTime As String

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
                ShiftFrom = IIf(IsNull(rs("ShiftFrom").value), Time, rs("ShiftFrom").value)
                ShiftTo = IIf(IsNull(rs("ShiftTo").value), Time, rs("ShiftTo").value)
  
    Else
    ShiftFrom = Time
    ShiftTo = Time
    End If

    If SeftCode = "" Then SeftCode = 0
    rs.Close
End Function

Public Function RegisterItemData(ScreenName As String, transactiontype As String, itemcode As String, ItemName As String, Optional itemUnit As String = "", Optional Qty As String = "", Optional Price As String = "", Optional total As String = "", Optional Color As String = "", Optional size As String = "", Optional Class As String = "", Optional DiscountType As String = "", Optional discountvalue As String = "", Optional NoteSerial As Double = 0, Optional NoteSerial1 As String = "", Optional NotesType As Integer = -1)
    LogTextA = " ﬂÊœ «·’‰› " & itemcode & "       «”„ «·’‰›       " & ItemName & CHR(13)

    If itemUnit <> "" Then
        LogTextA = LogTextA & "  ⁄œÌ·  «·ÊÕœ… «·Ï " & itemUnit & CHR(13)
    End If
  
    If Qty <> "" Then
        LogTextA = LogTextA & "  ⁄œÌ· «·ﬂ„Ì…  «·Ï " & Qty & CHR(13)
        '         LogTextA = LogTextA & "  «·«Ã„«·Ì " & total & Chr(13)
         
    End If
  
    If IsNumeric(Price) And val(Price) > 0 Then
        LogTextA = LogTextA & " ⁄œÌ·  «·”⁄—«·Ï " & Round(Price, SystemOptions.SysDefCurrencyForamt) & CHR(13)
        '         LogTextA = LogTextA & "  «·«Ã„«·Ì " & total & Chr(13)
    End If
  
    If DiscountType <> "" Then
        LogTextA = LogTextA & "  ⁄œÌ· ‰Ê⁄ «·Œ’„  «·Ï" & DiscountType & CHR(13)
    End If
  
    If discountvalue <> "" Then
        LogTextA = LogTextA & "  ⁄œÌ· ﬁÌ„… «·Œ’„ «·Ï" & discountvalue & CHR(13)
    End If

    If Color <> "" Then
        LogTextA = LogTextA & "    ⁄œÌ· «··Ê‰  «·Ï " & Color & CHR(13)
    End If
  
    If size <> "" Then
        LogTextA = LogTextA & " ⁄œÌ·  «·„ﬁ«” «·Ï  " & size & CHR(13)
    End If
  
    If Class <> "" Then
        LogTextA = LogTextA & "  ⁄œÌ· «·›∆… «·Ï " & Class & CHR(13)
    End If
                                         
    LogTexte = "  ItemCode  " & itemcode & "    Name  " & ItemName

    If itemUnit <> "" Then
        LogTexte = LogTexte & "Modify  Unit  To :" & itemUnit & CHR(13)
    End If
  
    If Qty <> "" Then
        LogTexte = LogTexte & "Modify  Qty  To :" & Qty & CHR(13)
        '      LogTextE = LogTextE & "Modify  Total " & total & Chr(13)
         
    End If
  
      If IsNumeric(Price) And val(Price) > 0 Then
        LogTexte = LogTexte & "Modify  Price To : " & Round(Price, SystemOptions.SysDefCurrencyForamt) & CHR(13)
        '    LogTextE = LogTextE & "Modify  Total " & total & Chr(13)
    End If
  
    If DiscountType <> "" Then
        LogTexte = LogTexte & " Modify  Discount Type  To :" & DiscountType & CHR(13)
    End If
  
    If discountvalue <> "" Then
        LogTexte = LogTexte & " Modify Discount value To :" & discountvalue & CHR(13)
    End If

    If Color <> "" Then
        LogTexte = LogTexte & " Modify Color  To :" & Color & CHR(13)
    End If
  
    If size <> "" Then
        LogTexte = LogTexte & "Modify  Size To : " & size & CHR(13)
    End If
  
    If Class <> "" Then
        LogTexte = LogTexte & " Modify Class To : " & Class & CHR(13)
    End If
                                          
    AddToLogFile CInt(user_id), NotesType, Date, Time, LogTextA, LogTexte, ScreenName, transactiontype, "", "", NoteSerial, NoteSerial1
 
End Function

Public Function ShowForm(ByVal sFormName As String)
    Dim fTemp As Form
    
    If checkApility(sFormName) = False Then
        Exit Function
    End If

    'add the form to Forms collection
    Set fTemp = Forms.Add(sFormName)
    'show the Form
    Load fTemp
    fTemp.show
End Function

Function RegisterLogInOut(frmname As String, ScreennameA As String, ScreennameE As String, Optional InOut As String = "", Optional NotesType As Integer)

    If InOut = "1" Then
        LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & ScreennameA
        LogTexte = " Open Window " & ScreennameE
    Else
        LogTextA = "  «·Œ—ÊÃ  „‰ ‘«‘… " & ScreennameA
        LogTexte = " Exit Window " & ScreennameE

    End If

    'AddToLogFile CInt(user_id), notesType, Date, Time, LogTextA, LogTextE, frmname, "O", "", ""

End Function

Public Function padding(str As String, _
                        noofchar As Integer) As String
    Dim lenstr As Integer
    Dim Diff As Integer
    Dim newStr As String

    lenstr = Len(str)

    If noofchar > lenstr Then
        Diff = noofchar - lenstr
        newStr = String(Diff, " ")
        padding = str & newStr
                    
    End If

End Function

Public Function zeropadding(str As String, _
                            noofchar As Integer, _
                            Optional atend As Boolean) As String
    Dim lenstr As Integer
    Dim Diff   As Integer
    Dim newStr As String
    newStr = ""
    lenstr = Len(str)
   
    If atend = False Then
        ' zeropadding = newStr & str
        zeropadding = Format(str, String(noofchar, "0"))
    Else
        If noofchar > lenstr Then
            Diff = noofchar - lenstr
            newStr = String(Diff, "0")
        End If
        zeropadding = str & newStr
    End If
   
   zeropadding = ""
   newStr = ""
    lenstr = Len(str)

    If noofchar > lenstr Then
        Diff = noofchar - lenstr
        newStr = String(Diff, "0")
     
                    
    End If
    If atend = False Then
   zeropadding = newStr & str
   Else
   zeropadding = str & newStr
   End If
End Function


Sub getCashireData(ID As Integer, _
                   Optional ByRef PointID As Integer, _
                   Optional ByRef Pointname As String, _
                   Optional ByRef Balance As Double, Optional ByRef PettyId As Long, Optional ByRef pettyBalance As Double, Optional ByRef BoxID As Long, Optional EmpID As Double)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    
 Dim pettyAccount As String
 Dim BoxAccount As String
 

    sql = " SELECT  dbo.cachierData.PettyId,  dbo.cachierData.BoxID  ,dbo.cachierData.PointId, dbo.cachierData.name, dbo.Tblposdata.BoxName, dbo.Tblposdata.BoxNamee, dbo.Tblposdata.Account_Code, dbo.cachierData.password, "
    sql = sql & "   dbo.cachierData.UserName , dbo.cachierData.Ctype, dbo.cachierData.namee, dbo.cachierData.EmpID, dbo.cachierData.id,  dbo.Tblposdata.BranchId"
    sql = sql & "  FROM         dbo.cachierData INNER JOIN"
    sql = sql & "   dbo.Tblposdata ON dbo.cachierData.PointId = dbo.Tblposdata.BoxID"
   If EmpID = 0 Then
    sql = sql & "   WHERE      (dbo.cachierData.id = " & ID & ")"
    Else
    sql = sql & "   WHERE      (dbo.cachierData.EmpID = " & EmpID & ")"
    End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim FirstPeriod As Date

 
    If rs.RecordCount > 0 Then
        PointID = IIf(IsNull(rs("PointId").value), 0, rs("PointId").value)

        If SystemOptions.UserInterface = ArabicInterface Then
            Pointname = IIf(IsNull(rs("BoxName").value), 0, rs("BoxName").value)
        Else
            Pointname = IIf(IsNull(rs("BoxNamee").value), 0, rs("BoxNamee").value)
        End If
   PettyId = IIf(IsNull(rs("PettyId").value), 0, rs("PettyId").value) '??CE C???IE
   BoxID = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value) '??CE C???IE

                 '
     getFirstPeriodDateInthisYear FirstPeriod
      BoxAccount = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", BoxID)
        Balance = 0 ' GetActualAccountBalance(BoxAccount, , FirstPeriod, Date)
    pettyAccount = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", PettyId)
 pettyBalance = 0 ' GetActualAccountBalance(pettyAccount, , FirstPeriod, Date)
    Else
    PettyId = 0
    BoxID = 0
        PointID = 0
        Pointname = ""
        Balance = 0
    End If
   
    rs.Close
 
End Sub

Public Sub WriteCustomerBalPublic(Optional AccountCode As String = "", _
                                  Optional ByRef Balance As String, _
                                  Optional balanceString As String, _
                                  Optional ByRef balancetype As Integer, _
                                  Optional notes_all As Double, _
                                  Optional ByRef account_name As String, _
                                  Optional ByRef Account_NameEng As String, _
                                  Optional ByRef Account_Serial As String, _
                                  Optional ToDat As Date, _
                                  Optional SendDate As Integer = 0)
    Dim StrTemp             As String
    Dim SngCusBegainAccount As Double
    Dim FirstPeriod         As Date
 
    If AccountCode <> "" Then

        'SngCusBegainAccount = GetCustomerAccount(Val(Me.DBCboClientName.BoundText), True)
        getFirstPeriodDateInthisYear2 FirstPeriod
        If SendDate = 0 Then ToDat = Date
  
        SngCusBegainAccount = GetActualAccountBalance(AccountCode, 0, FirstPeriod, ToDat, , , , , account_name, Account_NameEng, Account_Serial)
 
        Balance = Abs(SngCusBegainAccount)

        If SngCusBegainAccount > 0 Then
            balancetype = 0

            If SystemOptions.UserInterface = ArabicInterface Then
                StrTemp = FormatNumber(SngCusBegainAccount, SystemOptions.SysDefCurrencyForamt, True, True, True) & " „œÌ‰ "
            Else
                StrTemp = FormatNumber(SngCusBegainAccount, SystemOptions.SysDefCurrencyForamt, True, True, True) & " Depit "
            End If

        ElseIf SngCusBegainAccount < 0 Then
            balancetype = 1

            If SystemOptions.UserInterface = ArabicInterface Then
                StrTemp = FormatNumber(SngCusBegainAccount, SystemOptions.SysDefCurrencyForamt, True, True, True) & " œ«∆‰ "
            Else
                StrTemp = FormatNumber(SngCusBegainAccount, SystemOptions.SysDefCurrencyForamt, True, True, True) & " Credit "
            End If

        Else

            If SystemOptions.UserInterface = ArabicInterface Then
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTemp = " Œ«·’ "
                Else
                    StrTemp = "  "
                End If
        
            Else
                StrTemp = "  "
            End If
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrTemp = "0" & " Œ«·’ "
        Else
            StrTemp = "0" & "  "
        End If
   
    End If

    balanceString = StrTemp
End Sub



Public Sub save_login_info1(DBPath As String, _
                            dbname As String, Optional ServersName As String)
    'On Error Resume Next
    If ServersName = "" Then
    SaveSetting SystemOptions.SysRegsAppPath, "ServerCon", "ServerName", "."
    Else
    SaveSetting SystemOptions.SysRegsAppPath, "ServerCon", "ServerName", Trim(ServersName)
    End If
    SaveSetting "Byte_DBS", "Setting", "DBPath", DBPath

    SaveSetting "Byte_DBS", "Setting", "DBname", dbname
    SaveSetting "Byte_DBS", "Setting", "ServersName", ServersName
  '  SaveSetting "Byte_DBS", "Setting", "ServerCon", ServersName
  '    SysSQLServerName
 SaveSetting SystemOptions.SysRegsAppPath, "ServerCon", "ServerName", ServersName
 
    SaveSetting "Byte_DBS", "Setting", "SysSQLServerUserId", SystemOptions.SysSQLServerUserId
    SaveSetting "Byte_DBS", "Setting", "SysSQLServerUserpassword", SystemOptions.SysSQLServerUserpassword
    
    

End Sub

Function updateAutoOpeningBalanceLineNo(notes_id As Integer)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    If notes_id <> 1 Then
        Exit Function
    End If
    sql = " SELECT     TOP 100 PERCENT opening_balance_voucher_id,value, Double_Entry_Vouchers_ID, Credit_Or_Debit, DEV_ID_Line_No, Notes_ID"
    sql = sql & " from dbo.DOUBLE_ENTREY_VOUCHERS1"
    sql = sql & "  Where (Notes_ID = " & notes_id & ")"
    sql = sql & "  ORDER BY Double_Entry_Vouchers_ID, Credit_Or_Debit, Notes_ID"

    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            rs("DEV_ID_Line_No").value = i
            rs.update
            rs.MoveNext
        Next i
 
    End If
   
    rs.Close

End Function

Public Function CreateEmployee(emp_Name As String, _
                               EmpNamee As String) As Integer
    Dim Emp_id As Integer
    Dim StrSQL As String
    Dim rs As New ADODB.Recordset
 
    Emp_id = CStr(new_id("TblEmployee", "Emp_ID", "", True))
   
    StrSQL = "select * from  TblEmployee  where Emp_ID=-1"

    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    rs.AddNew
    rs("Emp_ID").value = Emp_id
    rs("Emp_Name").value = emp_Name
    rs("Emp_Namee").value = Emp_Namee
    rs.update
   
    rs.Close
    CreateEmployee = Emp_id
End Function

Public Function UnitsHaveTransactionsProjects(UnitID As Long) As Boolean

End Function
Public Function UnitsHaveTransactions(UnitID As Long) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = " SELECT * from  TblItemsUnits where UnitID=" & UnitID
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        UnitsHaveTransactions = True
    Else
        UnitsHaveTransactions = False
    End If
   
    rs.Close
    
End Function
Public Function GetSaleReportPOS(Optional ByRef report0 As String, Optional ByRef report1 As String, Optional ByRef report2 As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    sql = "  select * from tblusers  where userID=" & user_id
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        report0 = IIf(IsNull(rs("ReportName").value), "", rs("ReportName").value)
        report1 = IIf(IsNull(rs("ReportName1").value), "", rs("ReportName1").value)
         report2 = IIf(IsNull(rs("ReportName2").value), "", rs("ReportName2").value)
    Else
    report0 = ""
    report1 = ""
    report2 = ""
        GetSaleReportPOS = ""
    End If
   
    rs.Close
    
End Function
Public Function CheckOrderNotInTransaction(Transaction_Type As Double, _
                                           order_no As String) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = " SELECT * from  Transactions  where Transaction_Type= " & Transaction_Type & "and  order_no='" & order_no & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        CheckOrderNotInTransaction = True
    Else
        CheckOrderNotInTransaction = False
    End If
   
    rs.Close
    
End Function

Public Function GetCustomerIdByAccountCodeLong(Account_code As String, Optional ByRef fullcode) As Long
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    sql = "  SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.Fullcode"
    sql = sql & "    FROM         dbo.TblCustemers INNER JOIN"
    sql = sql & "                  dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code"
    sql = sql & "  where dbo.TblCustemers.Account_Code='" & Account_code & "'"
    sql = sql & GetAccountByBarnchUser
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    fullcode = IIf(IsNull(rs("fullcode").value), 0, rs("fullcode").value)
        GetCustomerIdByAccountCodeLong = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
    Else
        GetCustomerIdByAccountCodeLong = 0
        fullcode = ""
        
    End If
   
    rs.Close
    
End Function



Public Function GetCustomerIdByAccountCode(Account_code As String, Optional ByRef fullcode) As Long
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    sql = "  SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.Fullcode"
    sql = sql & "    FROM         dbo.TblCustemers INNER JOIN"
    sql = sql & "                  dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code"
    sql = sql & "  where dbo.TblCustemers.Account_Code='" & Account_code & "'"
    sql = sql & GetAccountByBarnchUser
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    fullcode = IIf(IsNull(rs("fullcode").value), 0, rs("fullcode").value)
        GetCustomerIdByAccountCode = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
    Else
        GetCustomerIdByAccountCode = 0
        fullcode = ""
        
    End If
   
    rs.Close
    
End Function

 

Public Function CheckComponentsIntransactions(ComponentID As Integer) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = "  select * from TblChangedComponentRegister where ComponentID=" & ComponentID
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        CheckComponentsIntransactions = True
    Else
        CheckComponentsIntransactions = False
    End If
   
    rs.Close
    
End Function

Public Function CheckFixedAssetsDipre(ID As Integer) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = " Select  id,Name  From FixedAssets where id= " & ID & " and  HaveDepreciation=1 and (Status_id =2 or Status_id =3) and PurchasePrice>0 order by Name"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        CheckFixedAssetsDipre = True
    Else
        CheckFixedAssetsDipre = False
    End If
   
    rs.Close
     
End Function
'18092013

Public Function GetCustomerAllData(CusID As Double, _
                                   Optional ByRef CusName As String = "", _
                                   Optional ByRef ExpireDateH As String, _
                                   Optional ByRef company As String = "", _
                                   Optional ByRef JobTitle As String = "", _
                                   Optional ByRef salary As Double = 0, _
                                   Optional ByRef JobAddress As String = "", _
                                   Optional ByRef JobTel As String = "", _
                                   Optional ByRef JobTelConvert As String = "", _
                                   Optional ByRef HomeTel As String = "", _
                                   Optional ByRef Mobile1 As String = "", _
                                   Optional ByRef Mobile2 As String = "", _
                                   Optional ByRef Dola As String = "", _
                                   Optional ByRef Madena As String = "", _
                                   Optional ByRef hay As String = "", _
                                   Optional ByRef Natinality As String = "", _
                                   Optional ByRef Sex As String = "", _
                                   Optional ByRef CustGID As Double = 0, _
                                   Optional ByRef Address As String)

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = " SELECT     dbo.TblCountriesData.CountryName AS Dola, dbo.TblCountriesGovernments.GovernmentName AS Madena, dbo.TblCountriesGovernmentsCities.CityName AS hay, "
    sql = sql & "    dbo.TblCustemers.*, TblCountriesData_1.CountryName AS Natinality"
    sql = sql & "  FROM         dbo.TblCustemers INNER JOIN"
    sql = sql & "  dbo.TblCountriesData ON dbo.TblCustemers.CountryID = dbo.TblCountriesData.CountryID INNER JOIN"
    sql = sql & "  dbo.TblCountriesGovernments ON dbo.TblCustemers.GovernmentID = dbo.TblCountriesGovernments.GovernmentID INNER JOIN"
    sql = sql & "   dbo.TblCountriesGovernmentsCities ON dbo.TblCustemers.CityID = dbo.TblCountriesGovernmentsCities.CityID INNER JOIN"
    sql = sql & "   dbo.TblCountriesData TblCountriesData_1 ON dbo.TblCustemers.CountryID = TblCountriesData_1.CountryID"

    sql = sql & "  WHERE     (dbo.TblCustemers.CusID = " & CusID & ")"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            CusName = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
        Else
            CusName = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
        End If
 
        CustGID = IIf(IsNull(rs("CustGID").value), 0, rs("CustGID").value)
        ExpireDateH = IIf(IsNull(rs("ExpireDateH").value), ToHijriDate(Date), rs("ExpireDateH").value)
        company = IIf(IsNull(rs("Company").value), "", rs("Company").value)
        JobTitle = IIf(IsNull(rs("JobTitle").value), "", rs("JobTitle").value)
        salary = IIf(IsNull(rs("Salary").value), 0, rs("Salary").value)
        JobAddress = IIf(IsNull(rs("JobAddress").value), "", rs("JobAddress").value)
        JobTel = IIf(IsNull(rs("JobTel").value), "", rs("JobTel").value)
        JobTelConvert = IIf(IsNull(rs("JobTelConvert").value), "", rs("JobTelConvert").value)
        HomeTel = IIf(IsNull(rs("HomeTel").value), "", rs("HomeTel").value)
        Mobile1 = IIf(IsNull(rs("Mobile1").value), "", rs("Mobile1").value)
        Mobile2 = IIf(IsNull(rs("Mobile2").value), "", rs("Mobile2").value)

        Dola = IIf(IsNull(rs("Dola").value), "", rs("Dola").value)
        Madena = IIf(IsNull(rs("Madena").value), "", rs("Madena").value)
        hay = IIf(IsNull(rs("hay").value), "", rs("hay").value)
        Natinality = IIf(IsNull(rs("Natinality").value), "", rs("Natinality").value)
        Sex = IIf(IsNull(rs("Sex").value), "", rs("Sex").value)
        Address = IIf(IsNull(rs("Address").value), "", rs("Address").value)
        
    Else
 
    End If
   
    rs.Close
     
End Function

Public Function CheckItemsIntransactions(Item_ID As Integer, Optional ByVal mUnitId As Integer = 0) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = " SELECT     Item_ID"
    sql = sql & " from dbo.Transaction_Details"
    sql = sql & " WHERE     (Item_ID = " & Item_ID & ") "
    If mUnitId <> 0 Then
        sql = sql & " and (UnitId= " & mUnitId & ") "
    End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        CheckItemsIntransactions = True
    Else
        CheckItemsIntransactions = False
    End If
   
    rs.Close
    
End Function


Public Function GetCarName(CarID As Integer, _
                           Optional ByRef BoardNO As String, Optional ByRef CarsTypeId As Integer, Optional ByRef LastKMCounter As Double, Optional ByRef VehicleLong As Double, Optional ByRef EquQty As Double, Optional ByRef Emp_id As Integer)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = " SELECT  *     "
    sql = sql & " from dbo.TblCarsData"
    sql = sql & " WHERE     (id = " & CarID & ") "
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        BoardNO = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
        CarsTypeId = IIf(IsNull(rs("CarsTypeId").value), 0, rs("CarsTypeId").value)
        LastKMCounter = IIf(IsNull(rs("LastKMCounter").value), 0, rs("LastKMCounter").value)
        VehicleLong = IIf(IsNull(rs("VehicleLong").value), 0, rs("VehicleLong").value)
        EquQty = IIf(IsNull(rs("EquQty").value), 0, rs("EquQty").value)
 
Emp_id = IIf(IsNull(rs("Emp_id").value), 0, rs("Emp_id").value)

    Else
    Emp_id = 0
        BoardNO = ""
        CarsTypeId = 0
        LastKMCounter = 0
        VehicleLong = 0
        EquQty = 0
    End If
   
    rs.Close
    
End Function




Public Function GetEquipmentsData(ID As Integer, _
                      Optional ByRef EmpID1 As Integer, Optional ByRef empID2 As Integer)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 
    sql = "SELECT     dbo.TblEquipments.*, dbo.TblEmployee.Emp_Code AS operatorCode, dbo.TblEmployee.Emp_Name AS operatorName, "
    sql = sql & "     dbo.TblEmployee.Emp_Namee AS operatornamee, TblEmployee_1.Emp_Code AS SupervisionCode, TblEmployee_1.Emp_Name AS Supervisionname,"
    sql = sql & "   TblEmployee_1.Emp_Namee AS SupervisionnameE"
    sql = sql & " FROM         dbo.TblEquipments INNER JOIN"
    sql = sql & "    dbo.TblEmployee ON dbo.TblEquipments.empID1 = dbo.TblEmployee.Emp_ID INNER JOIN"
    sql = sql & "    dbo.TblEmployee TblEmployee_1 ON dbo.TblEquipments.empID2 = TblEmployee_1.Emp_ID"
    sql = sql & "  Where (dbo.TblEquipments.fixedAssetid = " & ID & ")"
    
       
    
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
      
        EmpID1 = IIf(IsNull(rs("empID1").value), 0, rs("empID1").value)
        empID2 = IIf(IsNull(rs("empID2").value), 0, rs("empID2").value)
 
    Else
    EmpID1 = 0
    empID2 = 0
    
     End If
   
    rs.Close
    
End Function
Public Function updateNotesValueAndNobytext(notes_id As Double, Optional totalvalue1 As Double = 0, Optional nolegal As Boolean = False, Optional ByVal mIsOpenBalance As Boolean = False)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim TotalValue As Double
    Dim TotalValue2 As Double
    Dim TotalText As String
    Dim TotalText2 As String
    Dim Account_Code_dynamic As String
    
   
    If mIsOpenBalance Then
        sql = "select sum([Value]) as total from DOUBLE_ENTREY_VOUCHERS1 where Credit_Or_Debit=0 and Notes_ID= " & notes_id
    Else
        sql = "select sum([Value]) as total from DOUBLE_ENTREY_VOUCHERS where Credit_Or_Debit=0 and Notes_ID= " & notes_id
    End If

If nolegal = True Then
             Account_Code_dynamic = get_account_code_branch(72, my_branch)
sql = sql & " and Account_Code not in ("
sql = sql & " SELECT     Account_Code "
sql = sql & "  From dbo.ACCOUNTS"
sql = sql & "  WHERE     (Parent_Account_Code =  '" & Account_Code_dynamic & "'))"


End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
     TotalValue2 = totalvalue1
     TotalValue = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
     TotalText = WriteNo(Format(TotalValue, "0.00"), 0, True, ".")
    ' TotalValue2 = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
     TotalText2 = WriteNo(Format(TotalValue2, "0.00"), 0, True, ".")
    Else
    If totalvalue1 <> 0 Then
        TotalValue = totalvalue1
        TotalValue2 = totalvalue1
        TotalText = WriteNo(Format(TotalValue, "0.00"), 0, True, ".")
        TotalText2 = WriteNo(Format(TotalValue2, "0.00"), 0, True, ".")
     Else
     TotalText = ""
     TotalText2 = ""
     TotalValue2 = 0
     TotalValue = 0
     End If
      '
    End If

    rs.Close
   ' If totalvalue1 <> 0 Then
   ' TotalValue = totalvalue1
   ' TotalText = WriteNo(Format(totalvalue1, "0.00"), 0, True, ".")
   ' End If
    
    '    sql = "Update Notes set Note_Value= " & totalvalue & ",note_value_by_characters='" & TotalText & "'"
    If mIsOpenBalance Then
        sql = "Update Notes1 set  Note_Value= " & TotalValue & ",note_value_by_characters='" & TotalText & "'  where NoteID=" & notes_id
    
        Cn.Execute sql
        'sql = "Update Notes1 set  Note_Value2= " & TotalValue2 & ",note_value_by_characters2='" & TotalText2 & "'  where NoteID=" & notes_id
        'Cn.Execute sql
    Else
        sql = "Update Notes set  Note_Value= " & TotalValue & ",note_value_by_characters='" & TotalText & "'  where NoteID=" & notes_id
    
        Cn.Execute sql
        sql = "Update Notes set  Note_Value2= " & TotalValue2 & ",note_value_by_characters2='" & TotalText2 & "'  where NoteID=" & notes_id
        Cn.Execute sql
    End If
End Function
 

Public Function GetComponentForAlllocations(Emp_id As Integer, _
                                            ID As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String

    sql = "select id from mofrad where  " & SearchFiled & " =1"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            str = str & IIf(IsNull(rs("id").value), "", rs("id").value) & ","
 
            rs.MoveNext
 
        Next i

    Else
        GetComponentForAlllocations = ""
 
    End If
 
    If Len(str) > 1 Then
        GetComponentForAlllocations = mId(str, 1, Len(str) - 1)
    End If

    rs.Close
End Function

Public Function GetEmpContarctingComponents(DefDataType As Integer, _
                                            Emp_id As Integer, _
                                            Optional ByRef Output As String, _
                                            Optional ByRef DivisionMonth As Integer, _
                                            Optional ByRef Issue_date As Date, Optional ByRef Divisionsalary As Double, Optional ByRef TicketValue As Double)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
    ' 0 «Ã«“…
    '1 “Ì«œ… ”‰ÊÌ…
    '2 ‰Â«Ì… Œœ„…

    str = ""
    sql = "SELECT   dbo.Contract.TicketValue ,  dbo.TblContractDetails.Mofradtype, dbo.Contract.Due_period_no, dbo.Contract.due_period,   dbo.Contract.Issue_date"
    sql = sql & " , dbo.Contract.salary_period_no,dbo.Contract.salary_period FROM         dbo.TblContractDetails INNER JOIN"
    sql = sql & " dbo.Contract ON dbo.TblContractDetails.Contract_ID = dbo.Contract.Contract_ID"
 
    sql = sql & " Where (dbo.TblContractDetails.Emp_id = " & Emp_id & ") And (DefDataType = " & DefDataType & ")"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
TicketValue = IIf(IsNull(rs("TicketValue").value), 0, rs("TicketValue").value)
            If rs("due_period").value = 0 Then 'month
                DivisionMonth = IIf(IsNull(rs("Due_period_no").value), 0, rs("Due_period_no").value)
            ElseIf rs("due_period").value = 1 Then 'year
                DivisionMonth = IIf(IsNull(rs("Due_period_no").value), 0, rs("Due_period_no").value) * 12
                 ElseIf rs("due_period").value = 2 Then 'day
                DivisionMonth = IIf(IsNull(rs("Due_period_no").value), 0, rs("Due_period_no").value) / 30
            Else
                DivisionMonth = -1
            End If

          If rs("salary_period").value = 0 Then 'day
                Divisionsalary = IIf(IsNull(rs("salary_period_no").value), 0, rs("salary_period_no").value)
            ElseIf rs("salary_period").value = 1 Then 'month
                Divisionsalary = IIf(IsNull(rs("salary_period_no").value), 0, rs("salary_period_no").value) * 30
            Else
            Divisionsalary = -1
            End If
            
            Issue_date = IIf(IsNull(rs("Issue_date").value), Date, rs("Issue_date").value)

            str = str & IIf(IsNull(rs("Mofradtype").value), "", rs("Mofradtype").value) & ","
 
            rs.MoveNext
 
        Next i

    Else
        Output = ""
 
    End If
 
    If Len(str) > 1 Then
        Output = mId(str, 1, Len(str) - 1)
    End If

    rs.Close
End Function

Public Function GetSpecificComponentIncalculations(SpecificComponentId As Integer, Optional ByRef Equation As Double) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String

    str = ""
 
    sql = "SELECT      dbo.tblSpecificComponentDetails.Equation , dbo.tblSpecificComponentDetails.ComponentID, dbo.tblSpecificComponent.ComponentID AS MAINcOMPONENTID"
    sql = sql & "  FROM         dbo.tblSpecificComponentDetails INNER JOIN"
    sql = sql & "   dbo.tblSpecificComponent ON dbo.tblSpecificComponentDetails.SpecificComponentId = dbo.tblSpecificComponent.SpecificComponentId"
    sql = sql & "   WHERE     (dbo.tblSpecificComponent.ComponentID = " & SpecificComponentId & ")"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
Equation = 1
    If rs.RecordCount > 0 Then
Equation = IIf(IsNull(rs("Equation").value) Or rs("Equation").value = 0, 1, rs("Equation").value)
        For i = 1 To rs.RecordCount
            str = str & IIf(IsNull(rs("ComponentID").value), "", rs("ComponentID").value) & ","
 
            rs.MoveNext
 
        Next i

    Else
        GetSpecificComponentIncalculations = "0"
 
    End If
 
    If Len(str) > 1 Then
        GetSpecificComponentIncalculations = mId(str, 1, Len(str) - 1)
    End If

    rs.Close
End Function

Public Function GetComponentIncalculations(ID As Integer) As String
    Dim sql         As String
    Dim rs          As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str         As String

    If ID = 8 Then
        SearchFiled = "OverTime"
    ElseIf ID = 9 Then
        SearchFiled = "Punch"
    ElseIf ID = 10 Then
        SearchFiled = "Discount"
    ElseIf ID = 11 Then
        SearchFiled = "Absence"
    ElseIf ID = 12 Then
        SearchFiled = "Late"
    ElseIf ID = 0 Then
        SearchFiled = "Aloc1"
    ElseIf ID = 1 Then
        SearchFiled = "Aloc2"
    ElseIf ID = 2 Then
        SearchFiled = "Aloc1"
        
    Else
        GetComponentIncalculations = ""
        Exit Function
    End If

    str = ""

    sql = "select id from mofrad where  " & SearchFiled & " =1"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            str = str & IIf(IsNull(rs("id").value), "", rs("id").value) & ","
 
            rs.MoveNext
 
        Next i

    Else
        GetComponentIncalculations = "0"
 
    End If
 
    If Len(str) > 1 Then
        GetComponentIncalculations = mId(str, 1, Len(str) - 1)
    End If

    rs.Close
End Function

Public Function GetAllProfitsAccounts() As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
    str = ""
 
    sql = "SELECT     Account_Code from dbo.ACCOUNTS WHERE     (Parent_Account_Code = N'" & get_account_code_branch(49, my_branch) & "')"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount

            If i = rs.RecordCount Then
                str = str & IIf(IsNull(rs("Account_code").value), "", "'" & rs("Account_code").value) & "'"

            Else
                str = str & IIf(IsNull(rs("Account_code").value), "", "'" & rs("Account_code").value) & "',"
            End If

            rs.MoveNext
 
        Next i

    Else
        GetAllProfitsAccounts = "0,0"
 
    End If

    GetAllProfitsAccounts = str
    rs.Close
End Function

Public Function GetAllCompositeAccounts(CombositAccountid As Long) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
    str = ""
 
    sql = "SELECT     * from TblCombositAccountDetails where CombositAccountid= " & CombositAccountid
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount

            If i = rs.RecordCount Then
                str = str & IIf(IsNull(rs("Account_code").value), "", "'" & rs("Account_code").value) & "'"

            Else
                str = str & IIf(IsNull(rs("Account_code").value), "", "'" & rs("Account_code").value) & "',"
            End If

            rs.MoveNext
 
        Next i

    Else
        GetAllCompositeAccounts = "0,0"
 
    End If

    GetAllCompositeAccounts = str
    rs.Close
End Function

Public Function GetAllprojectsAccounts(ID As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String

    str = ""
    'Account_code
    sql = "select *  from projects where id= " & ID
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            str = IIf(IsNull(rs("expanses_account").value), "", "'" & rs("expanses_account").value) & "',"
            str = str & IIf(IsNull(rs("REVENUE_account").value), "", "'" & rs("REVENUE_account").value) & "',"
             str = str & IIf(IsNull(rs("legal").value), "", "'" & rs("legal").value) & "',"
             'legal
             str = str & IIf(IsNull(rs("Salary_account").value), "", "'" & rs("Salary_account").value) & "',"
            str = str & IIf(IsNull(rs("Material_account").value), "", "'" & rs("Material_account").value) & "'"
 
            rs.MoveNext
 
        Next i

    Else
        GetAllprojectsAccounts "0,0"
 
    End If
 
    GetAllprojectsAccounts = str
 
    rs.Close
End Function



Public Function GetAllIDFromTblAttributionInstallmentDivided(REID As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
Dim i As Integer
    str = ""
    'Account_code
    sql = "  select * from TblAttributionInstallmentDivided  where REID=" & REID & "  order by idac "
    'Sql = "select *  from dbo.TblBalanceSheetDetails where BalanceSheetHeaderid= " & ID
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            str = str & IIf(IsNull(rs("id").value), "", "" & rs("id").value & ",")
  
            rs.MoveNext
 
        Next i

    Else
        GetAllIDFromTblAttributionInstallmentDivided = ""
 
    End If
 
    GetAllIDFromTblAttributionInstallmentDivided = mId(str, 1, Len(str) - 1)
 
    rs.Close
End Function

Public Function GetAllIDFromTblMinistrtyContract(VRID As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
Dim i As Integer
    str = ""
    'Account_code
    sql = "  select * from TblMinistryContract_Installment  where VRID=" & VRID & " order by idmc"
    'Sql = "select *  from dbo.TblBalanceSheetDetails where BalanceSheetHeaderid= " & ID
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            str = str & IIf(IsNull(rs("id").value), "", "" & rs("id").value & ",")
  
            rs.MoveNext
 
        Next i

    Else
        GetAllIDFromTblMinistrtyContract = ""
 
    End If
 
    GetAllIDFromTblMinistrtyContract = mId(str, 1, Len(str) - 1)
 
    rs.Close
End Function

Public Function GetAlLastAccounts(Parent_Account_Code As String) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
Dim i As Integer
    str = ""
    'Account_code
    sql = "select *  from dbo.ACCOUNTS where last_account=1 and Parent_Account_Code like '" & Parent_Account_Code & "%'"
    sql = sql & GetAccountByBarnchUser
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            str = str & IIf(IsNull(rs("Account_code").value), "", "'" & rs("Account_code").value & "',")
  
            rs.MoveNext
 
        Next i

    Else
        GetAlLastAccounts = ""
 
    End If
 If str <> "" Then
    GetAlLastAccounts = mId(str, 1, Len(str) - 1)
 End If
    rs.Close
End Function


 

Public Function GetAllTrialBalanceAccounts(ID As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
Dim i As Integer
    str = ""
    'Account_code
    sql = "select *  from dbo.TblBalanceSheetDetails where BalanceSheetHeaderid= " & ID
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            str = str & IIf(IsNull(rs("Account_code").value), "", "'" & rs("Account_code").value & "',")
  
            rs.MoveNext
 
        Next i

    Else
        GetAllTrialBalanceAccounts = ""
 
    End If
 
    GetAllTrialBalanceAccounts = mId(str, 1, Len(str) - 1)
 
    rs.Close
End Function

Public Function GetAllEmployeeAccounts(Emp_id As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String

    str = ""
    'Account_code
    sql = "select *  from TblEmployee where Emp_ID= " & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            str = IIf(IsNull(rs("Account_code").value), "", "'" & rs("Account_code").value & "',")
             str = str & IIf(IsNull(rs("Account_code1").value), "", "'" & rs("Account_code1").value & "',")
             str = str & IIf(IsNull(rs("Account_code5").value), "", "'" & rs("Account_code5").value & "',")
            str = str & IIf(IsNull(rs("Account_code2").value), "", "'" & rs("Account_code2").value & "'")
            If Not IsNull(rs("Account_code4").value) Then
            
                    str = str & IIf(IsNull(rs("Account_code3").value), "", ",'" & rs("Account_code3").value & "',")
                    str = str & IIf(IsNull(rs("Account_code4").value), "", "'" & rs("Account_code4").value & "'")
            Else
                      '    If Not IsNull(rs("Account_code3").value) Then
                                  str = str & IIf(IsNull(rs("Account_code3").value), "", "      ,'" & rs("Account_code3").value & "'")
                      '            Else
                      '            str = str & "'"
                      '  End If
          
            
            End If
            rs.MoveNext
 
        Next i

    Else
        GetAllEmployeeAccounts = "0,0"
 
    End If
 
    GetAllEmployeeAccounts = str
 
    rs.Close
End Function

Public Function GetEmployeeˆAvgChangedSalary(Emp_id As Integer, _
                                             StrWhere As String, _
                                             Actualyear As Integer) As Double
    Dim sql     As String
    Dim rs      As New ADODB.Recordset
    Dim Balance As Double
    If StrWhere = "" Then
        GetEmployeeˆAvgChangedSalary = 0
        Exit Function
    End If
 
    sql = "SELECT     AVG(dbo.TblChangedComponentRegisterDetails.[value]) AS TotalAvg"
    sql = sql & "  FROM         dbo.TblChangedComponentRegisterDetails INNER JOIN"
    sql = sql & "   dbo.TblChangedComponentRegister ON"
    sql = sql & "   dbo.TblChangedComponentRegisterDetails.ChangedComponentid = dbo.TblChangedComponentRegister.ChangedComponentid"
    sql = sql & "   WHERE     (dbo.TblChangedComponentRegister.ComponentID IN (" & StrWhere & ")) AND (dbo.TblChangedComponentRegisterDetails.Emp_id = " & Emp_id & ")"
    sql = sql & " and   (dbo.TblChangedComponentRegister.Actualyear = " & Actualyear & ") "
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetEmployeeˆAvgChangedSalary = IIf(IsNull(rs("TotalAvg").value), 0, rs("TotalAvg").value)
 
    Else
        GetEmployeeˆAvgChangedSalary = 0
 
    End If

    rs.Close
    
End Function

Public Function GetEmployeeChangedSalary(Emp_id As Integer, _
                                         ComponentID As Integer, _
                                         Actualyear As Integer, _
                                         Actualmonth As Integer) As Double
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    sql = "SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Total"
    sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
    sql = sql & "  dbo.TblChangedComponentRegisterDetails ON"
    sql = sql & "  dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid"
    sql = sql & " WHERE     (dbo.TblChangedComponentRegister.Actualmonth = " & Actualmonth & ") AND (dbo.TblChangedComponentRegister.Actualyear = " & Actualyear & ") AND"
    sql = sql & "  (dbo.TblChangedComponentRegister.ComponentID = " & ComponentID & ") AND (dbo.TblChangedComponentRegisterDetails.Emp_id = " & Emp_id & ")"

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetEmployeeChangedSalary = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
 
    Else
        GetEmployeeChangedSalary = 0
 
    End If

    rs.Close
    
End Function

Public Function GetEmployeeSalaryAccordingToComponentName(Emp_id As Integer, _
                                                      whrstr As String, Optional Ch As Integer) As String
    Dim sql As String
    Dim mofrad_name As String
    Dim valuee As Double
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim Mofradd As String
    Dim i As Integer
    Mofradd = ""
    'mofrad_name = ""
If Ch = 1 Then
    'sql = "SELECT     SUM([Value]) AS Total"
    'sql = sql & " from dbo.EmpSalaryComponent"
    'sql = sql & " WHERE     (emp_ID = " & Emp_id & ") AND (mofrad_type IN (" & whrstr & "))"
'    If whrstr = "" Then GetEmployeeSalaryAccordingToComponent = 0: Exit Function
    sql = "SELECT     dbo.EmpSalaryComponent.[Value],dbo.mofrdat.mofrad_name "
    sql = sql & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
    sql = sql & " dbo.mofrdat ON dbo.EmpSalaryComponent.AccountCode = dbo.mofrdat.mofrad_code"
    sql = sql & " WHERE   (dbo.EmpSalaryComponent.Flagx Is Null) And  (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
  '  If whrstr <> "" Then ' (dbo.EmpSalaryComponent.Flagx Is Null) And (dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.text) & ")"
  '   sql = sql & "   AND (dbo.mofrdat.mofrad_type IN (" & whrstr & "))"
      rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
      '  GetEmployeeSalaryAccordingToComponent = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
      For i = 1 To rs.RecordCount
       mofrad_name = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
 valuee = IIf(IsNull(rs("value").value), 0, rs("value").value)
 Mofradd = Mofradd & mofrad_name & "   " & valuee
 Mofradd = Mofradd + vbNewLine
 
 rs.MoveNext
      Next i
 
  '  End If
     End If
     
     Else
         'sql = "SELECT     SUM([Value]) AS Total"
    'sql = sql & " from dbo.EmpSalaryComponent"
    'sql = sql & " WHERE     (emp_ID = " & Emp_id & ") AND (mofrad_type IN (" & whrstr & "))"
'    If whrstr = "" Then GetEmployeeSalaryAccordingToComponent = 0: Exit Function
    sql = "SELECT   dbo.EmpSalaryComponent.[Value],dbo.EmpSalaryComponent.EntIncresDataH,dbo.mofrdat.mofrad_name "
    sql = sql & " FROM         dbo.EmpSalaryComponent LEFT OUTER JOIN"
    sql = sql & " dbo.mofrdat ON dbo.EmpSalaryComponent.AccountCode = dbo.mofrdat.mofrad_code"
    sql = sql & " WHERE  (dbo.EmpSalaryComponent.Flagx = 1) And  (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
 '   If whrstr <> "" Then ' (dbo.EmpSalaryComponent.Flagx Is Null) And (dbo.EmpSalaryComponent.Emp_id = " & val(Emp_id.text) & ")"
 '    sql = sql & "   AND (dbo.mofrdat.mofrad_type IN (" & whrstr & "))"
      rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
      '  GetEmployeeSalaryAccordingToComponent = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
      For i = 1 To rs.RecordCount
       mofrad_name = IIf(IsNull(rs("EntIncresDataH").value), "", rs("EntIncresDataH").value)
 valuee = IIf(IsNull(rs("value").value), 0, rs("value").value)
 Mofradd = Mofradd & valuee & "   " & mofrad_name
 Mofradd = Mofradd + vbNewLine

 rs.MoveNext
      Next i
 
 '   End If
     End If
     End If
    'WHERE     (dbo.EmpSalaryComponent.emp_ID = 126) AND (dbo.mofrdat.mofrad_type IN (1))"
   
    
 GetEmployeeSalaryAccordingToComponentName = Mofradd
    rs.Close
    
End Function


Public Function GetEmployeeSalaryAccordingToComponent(Emp_id As Integer, _
                                                      whrstr As String, Optional Ch As Integer = 0, Optional EntIncresDataM As Date = "01/01/1900", Optional MonthID As Integer = 0, Optional YearID As Integer = 0) As Double
    Dim sql As String
    Dim Flag As Boolean
    Dim mofrad_name As String
    Dim valuee As Double
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
 Flag = ChekEmpInProject(Emp_id, MonthID, YearID)
If Ch = 1 Then
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
'SQL = SQL & " Where (dbo.EmpSalaryComponent.Emp_id = 2)"
If Flag = True Then
sql = sql & " WHERE     (dbo.EmpSalaryComponent.Flagx =2) and (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
Else
sql = sql & " WHERE     (dbo.EmpSalaryComponent.Flagx Is Null) and (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
End If
    If whrstr <> "" Then
     sql = sql & "   AND (dbo.mofrdat.mofrad_type IN (" & whrstr & "))"
     
     End If
sql = sql & " )x"
     
     Else
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
'SQL = SQL & " Where (dbo.EmpSalaryComponent.Emp_id = 2)"
sql = sql & " WHERE     (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
If Flag = True Then
sql = sql & " and     (dbo.EmpSalaryComponent.Flagx =2)"
Else
sql = sql & " and    ( (dbo.EmpSalaryComponent.Flagx Is Null) or (dbo.EmpSalaryComponent.Flagx =1)) "
End If

    If whrstr <> "" Then
     sql = sql & "   AND (dbo.mofrdat.mofrad_type IN (" & whrstr & "))"
     
     End If
sql = sql & " )x"
   
 
     End If
If Not (EntIncresDataM) = "01/01/1900" Then
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
'SQL = SQL & " Where (dbo.EmpSalaryComponent.Emp_id = 2)"

If Flag = True Then
sql = sql & " Where (dbo.EmpSalaryComponent.Emp_id = " & Emp_id & ")"
sql = sql & " and     (dbo.EmpSalaryComponent.Flagx =2) "
Else
sql = sql & " WHERE  ( EntIncresDataM is null or    EntIncresDataM<= " & SQLDate(EntIncresDataM, True) & ") and  (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ")"
sql = sql & " and     ( (dbo.EmpSalaryComponent.Flagx Is Null) or (dbo.EmpSalaryComponent.Flagx =1)) "
End If
    If whrstr <> "" Then
     sql = sql & "   AND (dbo.mofrdat.mofrad_type IN (" & whrstr & "))"
     
     End If
sql = sql & " )x"
End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    If rs.RecordCount > 0 Then
        GetEmployeeSalaryAccordingToComponent = Abs(IIf(IsNull(rs("Total").value), 0, rs("Total").value))
    Else
        GetEmployeeSalaryAccordingToComponent = 0
 
    End If
 
    rs.Close
    
End Function

 

Public Function GetMofradUnit(ID As Integer, Optional ByRef avg As Double = 0) As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    sql = "select * from mofrad where  id =" & ID
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetMofradUnit = IIf(IsNull(rs("Unit").value), 0, rs("Unit").value)
 avg = IIf(IsNull(rs("avg").value), 0, rs("avg").value)
    Else
        GetMofradUnit = 0
 avg = 0
    End If

    rs.Close
End Function

Public Function GetNoOfHourPerMonth() As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    sql = " SELECT sum((Go_HourTime)-Bring_HourTime+(isnull(Go_HourTime1,0))-isnull(Bring_HourTime1,0)) *4 as whour from tblTimeSetting"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetNoOfHourPerMonth = IIf(IsNull(rs("whour").value), 0, rs("whour").value)
 
    Else
        GetNoOfHourPerMonth = 0
 
    End If

    rs.Close
End Function

Public Function get_transaction_NoteSerial1ByiD(Transaction_ID As Double, _
                                                Transaction_Type As String) As String

    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select * from Transactions where Transaction_ID=" & Transaction_ID
    sql = sql & " and " & Transaction_Type
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        get_transaction_NoteSerial1ByiD = ""
    Else
        get_transaction_NoteSerial1ByiD = IIf(IsNull(rs("NoteSerial1").value), 0, rs("NoteSerial1").value)
    End If

End Function

Public Function GetOpeningBalanceDateForType2(ReportDate As Date) As Date
    '·ﬁ«∆„… «·œŒ· Ê «·„Ì“«‰ ··„’—Ê›«  Ê «·«Ì—«œ« 

    Dim rs As ADODB.Recordset
    Dim sql As String
 
    sql = "  SELECT     TOP 100 PERCENT OpeneingbalancesDate, DATEDIFF(day, OpeneingbalancesDate,'" & SQLDate(ReportDate) & "') AS Diff"
    sql = sql & " from dbo.TblyearsData"
    sql = sql & " WHERE     (OpeneingbalancesDate <='" & SQLDate(ReportDate) & "')"
    sql = sql & " ORDER BY DATEDIFF(day, OpeneingbalancesDate,'" & SQLDate(ReportDate) & "')"

    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        GetOpeningBalanceDateForType2 = IIf(IsDate(rs("OpeneingbalancesDate").value), rs("OpeneingbalancesDate").value, ReportDate)
    Else
        GetOpeningBalanceDateForType2 = ReportDate
    End If

End Function

Public Function updateprofitAccount(Optional ActivityId As Integer = 0, _
                                    Optional BranchID As Integer = 0, _
                                    Optional reportEnddate As Date, Optional ManyBranch As String = "", Optional SumProfit As Boolean, Optional reportStartdate As Date)

    ' ”ÃÌ· «·«—»«Õ Ê «·Œ”«∆— ·ﬂ· ⁄«„
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim profit As Double
 Dim TotalProfitMaster As Double
    Dim StrSQL As String
    Dim Account_code As String
 StrSQL = " update ACCOUNTS set ProfitBalance=0"
 Cn.Execute StrSQL
 
    sql = "  SELECT      * from Tblyearsdata "
    'where CurrentYear=1 "
    totalprofit = 0
    TotalProfitMaster = 0
    'Õ”«» «·«» ··«—»«Õ Ê «·Œ”«∆—

    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
 
        For i = 1 To rs.RecordCount
            
            
            If (rs("Account_Code").value) <> "" And Not IsNull(rs("Account_Code").value) Then
                If DateDiff("D", reportEnddate, rs("DateEnd").value) >= 0 Or year(rs("datesatrt").value) = year(reportEnddate) Then
               ' If DateDiff("D", reportEnddate, rs("DateEnd").value) >= 0 Then
                    profit = 0 ' getprofitValue(rs("datesatrt").value, reportEnddate, ActivityId, BranchId)
                  '  profit = getprofitValue(rs("datesatrt").value, reportEnddate, ActivityId, BranchID)
                Else
                   
                    profit = getprofitValue(rs("datesatrt").value, rs("DateEnd").value, ActivityId, BranchID, ManyBranch)
                '    profit = 0
                End If
          
                  
       '        profit = 0
                totalprofit = totalprofit + profit
                TotalProfitMaster = TotalProfitMaster + totalprofit
                If profit = 0 Then
                    StrSQL = " update ACCOUNTS"
                    StrSQL = StrSQL & " SET opening_balance=0"
                      StrSQL = StrSQL & ", ProfitBalance=0"
                    StrSQL = StrSQL & "  where  Account_code ='" & rs("Account_Code").value & "'"
                Else
                    StrSQL = " update ACCOUNTS"
                    StrSQL = StrSQL & " SET opening_balance=opening_balance+ " & profit * -1
                      StrSQL = StrSQL & ", ProfitBalance=" & profit * -1
                    StrSQL = StrSQL & "  where  Account_code ='" & rs("Account_Code").value & "'"
                    
'                         StrSQL = " update ACCOUNTS"
'                    StrSQL = StrSQL & " SET opening_balance=opening_balance+ " & profit
'                      StrSQL = StrSQL & ", ProfitBalance=" & profit
'                    StrSQL = StrSQL & "  where  Account_code ='" & rs("Account_Code").value & "'"
                End If
                Cn.Execute StrSQL
'
'                StrSQL = " update ACCOUNTS"
'                StrSQL = StrSQL & " SET opening_balance= " & profit * -1
'                  StrSQL = StrSQL & ", ProfitBalance=" & profit * -1
'                StrSQL = StrSQL & "  where  Account_code ='" & rs("Account_Code").value & "'"
'                Cn.Execute StrSQL
            
            End If
    
            rs.MoveNext
        Next i
     
    End If
      Dim Account_Code_dynamic1 As String

    Account_Code_dynamic1 = get_account_code_branch(49, my_branch)
        
If SumProfit = True Then
  
StrSQL = " update ACCOUNTS set ProfitBalance="
  StrSQL = StrSQL & "  ("
  StrSQL = StrSQL & "  SELECT     SUM(ACCOUNTS.ProfitBalance)"
  StrSQL = StrSQL & "   From Accounts"
  StrSQL = StrSQL & "  where    (Parent_Account_Code = '" & Account_Code_dynamic1 & "')"
  StrSQL = StrSQL & "  )"
 StrSQL = StrSQL & "  where  last_account=0  and (  Account_code ='" & Account_Code_dynamic1 & "'"
 StrSQL = StrSQL & "  or  Account_code ='" & mId(Account_Code_dynamic1, 1, 2) & "'"
 StrSQL = StrSQL & "  or  Account_code ='" & mId(Account_Code_dynamic1, 1, 4) & "'"
 StrSQL = StrSQL & "  or  Account_code like'" & mId(Account_Code_dynamic1, 1, 6) & "%')"
 
  Cn.Execute StrSQL
            
End If
End Function

Public Function updateopeningbalanceNewFromsql2(Optional FromDate As Date, _
                                                Optional ToDate As Date, _
                                                Optional continous As Boolean = False, _
                                                Optional ActivityId As Integer = 0, _
                                                Optional BranchID As Integer = 0, _
                                                Optional Account_code As String = "", _
                                                Optional updatetype As Integer = 0, _
                                                Optional notes_all As Double, _
                                                Optional project_id As Double)

    '0 balance Sheet
    '1 trial balances
    Dim openingbalacedate As Date
    ' getOpeningBalancedate DTPickerAccFrom.value, DTPickerAccTo.value, returnedfromdate, returnedTOdate, year(DTPickerAccFrom.value), openingbalacedate
    getOpeningBalancedate , , , , year(ToDate), openingbalacedate, continous
 
    Dim StrSQL As String

    If openingbalacedate = FromDate Then
     
        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalance2('" & SQLDate(openingbalacedate) & "', Account_code,last_account),"
        StrSQL = StrSQL & " balance= dbo.GetBalance2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account)"

        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByActivity2('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),"
            StrSQL = StrSQL & " balance= dbo.GetBalanceByActivity2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account)"
  
        End If

        If BranchID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByBranch2('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),"
            StrSQL = StrSQL & " balance= dbo.GetBalanceByBranch2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account)"
  
        End If
        
        If project_id <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByProject2('" & SQLDate(openingbalacedate) & "'," & project_id & " )"
            ' StrSQL = StrSQL & " balance= dbo.GetBalanceByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(toDate) & "'," & BranchId & ", Account_code,last_account)"
  
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
     
        StrSQL = StrSQL & " set  balance= dbo.GetBalance2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalance2('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)  +   isnull(dbo.GetBalance2('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code,last_account),0) "
    
        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByActivity2('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByActivity2('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
   
        End If

        If BranchID <> 0 Then
  
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByBranch2('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByBranch2('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0) "
  
        End If
 
        If project_id <> 0 Then
            StrSQL = " update ACCOUNTS"
            '  StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByProject('" & SQLDate(Fromdate) & "'," & project_id & " )"
            ' StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(toDate) & "'," & BranchId & ", Account_code,last_account)"
  
            StrSQL = StrSQL & " set opening_balance= isnull(dbo.GetOpeningBalanceByProject2('" & SQLDate(openingbalacedate) & "'," & project_id & " ),0)  +   isnull(dbo.GetBalanceByproject2('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & project_id & " ),0) "
  
        End If
            
    End If

    Dim FirstStr  As String
    Dim SecondStr As String
    FirstStr = StrSQL

    If updatetype = 0 Then ' „Ì“«‰Ì…
        StrSQL = StrSQL & " WHERE    (last_account )<> 1  and  (AccountTypes = 1) " 'ok
    ElseIf updatetype = 1 Then  ' „Ì“«‰
        StrSQL = StrSQL & " WHERE     (last_account = 1)  and  (AccountTypes = 1)"
 
    ElseIf updatetype = 2 Then ' ﬁ«∆„… «·œŒ·
        GoTo Part2
        'StrSQL = StrSQL & " WHERE   (last_account )<> 1  and   (AccountTypes = 2) "

    ElseIf updatetype = 3 Then   ' ﬂ‘› Õ”«»
   
        If getAccountTypes(Account_code) <> 1 Then ' ·Ê ﬂ«‰ Õ”«» „’—Ê› «Ê «Ì—«œ
            GoTo Part2
        End If
 
        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_code ='" & Account_code & "'"
  
        End If
 
    ElseIf updatetype = 4 Then ' Õ”«» «” «–

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Parent_Account_Code ='" & Account_code & "'  or Account_code ='" & Account_code & "'"
            
            If getAccountTypes(Account_code) <> 1 Then ' ·Ê ﬂ«‰ Õ”«» „’—Ê› «Ê «Ì—«œ
                GoTo Part2
            End If
        End If

    ElseIf updatetype = 30 Then  '  ﬂ‘› Õ”«» „Ã„⁄

        If (Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_Code in (" & Account_code & ")"
                
        End If
                  
    Else

    End If

    Cn.CommandTimeout = 10000
    Cn.Execute StrSQL

    If updatetype = 0 Or updatetype = 30 Or getAccountTypes(Account_code) = 1 Then
        Exit Function
    End If

    '***********************************************Part2
Part2:
   
    openingbalacedate = GetOpeningBalanceDateForType2(FromDate)
    
    If openingbalacedate = FromDate Then
     
        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalance2('" & SQLDate(openingbalacedate) & "', Account_code,last_account),"
        StrSQL = StrSQL & " balance= dbo.GetBalance2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account)"

        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByActivity2('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),"
            StrSQL = StrSQL & " balance= dbo.GetBalanceByActivity2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account)"
  
        End If

        If BranchID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByBranch2('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),"
            StrSQL = StrSQL & " balance= dbo.GetBalanceByBranch2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account)"
  
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
     
        StrSQL = StrSQL & " set  balance= dbo.GetBalance2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalance2('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)  +   isnull(dbo.GetBalance2('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code,last_account),0) "
    
        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByActivity2('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByActivity2('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
   
        End If

        If BranchID <> 0 Then
  
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch2('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByBranch2('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByBranch2('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0) "
  
        End If
 
    End If

    If updatetype = 1 Then  ' „Ì“«‰
        'StrSQL = StrSQL & " WHERE     (last_account = 1) "
        StrSQL = StrSQL & " WHERE     (last_account = 1)   and  (AccountTypes = 2) "
    ElseIf updatetype = 2 Then ' ﬁ«∆„… «·œŒ·
        StrSQL = StrSQL & " WHERE   (last_account )<> 1  and   (AccountTypes = 2) "
    ElseIf updatetype = 3 Then ' ﬂ‘› Õ”«»
   
        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_code ='" & Account_code & "'"
  
        End If
 
    ElseIf updatetype = 4 Then ' Õ”«» «” «–

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Parent_Account_Code ='" & Account_code & "'  or Account_code ='" & Account_code & "'"
        
        End If
          
    ElseIf updatetype = 30 Then  '  ﬂ‘› Õ”«» „Ã„⁄

        If (Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_Code in (" & Account_code & ")"
                
        End If
                  
    Else

    End If

    Cn.CommandTimeout = 10000
    
    Cn.Execute StrSQL
    
    Debug.Print StrSQL
   
End Function



Public Function createIntervalAccount(Optional ActivityId As Integer = 0, _
                                    Optional BranchID As Integer = 0, _
                                    Optional fromdat1e As Date, Optional todata1 As Date)

On Error Resume Next
    Dim rs As ADODB.Recordset
    Dim sql As String
      Dim Rs1 As ADODB.Recordset
   
    Dim StartDateArray(12) As Date
     Dim EndDateArray(12) As Date
     Dim intervalcoun As Integer
     
    Dim Balance As Double
 
    Dim StrSQL As String
    Dim Account_code As String
 
    sql = "   SELECT     dbo.TblyearsData.CurrentYear, dbo.TblAccountIntervals.StartDate, dbo.TblAccountIntervals.EndDate FROM         dbo.TblyearsData INNER JOIN dbo.TblAccountIntervals ON dbo.TblyearsData.TblyearsDataid = dbo.TblAccountIntervals.TblyearsDataid WHERE     (dbo.TblyearsData.CurrentYear = 1)"
    Balance = 0
     

    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
 
        For i = 1 To rs.RecordCount
            If i > 11 Then GoTo xl
            If IsNull(rs("StartDate").value) Then
            intervalcoun = i
            i = 13
            Else
             StartDateArray(i) = IIf(IsNull(rs("StartDate").value), Null, rs("StartDate").value)
             EndDateArray(i) = IIf(IsNull(rs("EndDate").value), Null, rs("EndDate").value)
           rs.MoveNext
            End If
          
    
           
        Next i
     
    End If
xl:
Dim LocalStr As String
LocalStr = "update ACCOUNTS set   "
For X = 1 To 12
LocalStr = LocalStr & "interval" & X & "=0,"
Next X
Cn.Execute mId(LocalStr, 1, Len(LocalStr) - 1)

    sql = "  SELECT interval1,interval2,interval3,interval4,interval5,interval6,interval7,interval8,interval9,interval10,interval11,interval12,  ACCOUNTS.Account_Code, dbo.ACCOUNTS.account_serial , dbo.ACCOUNTS.account_name, dbo.ACCOUNTS.Account_NameEng FROM         dbo.ACCOUNTS left   JOIN dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code Where (last_account <> 1) And (AccountTypes = 2) ORDER BY Account_Serial"
    Balance = 0
     

    Set Rs1 = New ADODB.Recordset
    Rs1.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText



    If Rs1.RecordCount > 0 Then
 
        For i = 1 To Rs1.RecordCount 'Õ”«»« 
            
            For j = 1 To rs.RecordCount ' ”‰Ê« 
        
         '    StartDateArray(i) = IIf(IsNull(rs("StartDate").value), Null, rs("StartDate").value)
         '  rs("interval1" & i) = Balance(IIf(IsNull(rs("Account_Code").value), Null, rs("Account_Code").value))
        Dim FromDate As Date
         Dim ToDate As Date
           
          FromDate = StartDateArray(j)
          ToDate = EndDateArray(j)
          
          
      StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " SET opening_balance=0,"
        StrSQL = StrSQL & " " & "interval" & j & "= dbo.GetBalance('" & SQLDate(FromDate) & "', '" & SQLDate(ToDate) & "', Account_code,last_account , " & IIf(SystemOptions.IsHiddenUser, 1, 0) & ")"

        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= 0,"
            StrSQL = StrSQL & " " & "interval" & j & "= dbo.GetBalanceByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & ")"
  
        End If

        If BranchID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= 0,"
            StrSQL = StrSQL & " " & "interval" & j & "= dbo.GetBalanceByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & ")"
  
        End If

  StrSQL = StrSQL & " WHERE   (last_account )<> 1  and   (AccountTypes = 2) "

 Cn.Execute StrSQL



          
          
             Next j
            Rs1.MoveNext
        Next i
     
    End If
    
    





End Function
Public Function updateopeningbalanceNewFromsql(Optional FromDate As Date, _
                                               Optional ToDate As Date, _
                                               Optional continous As Boolean = False, _
                                               Optional ActivityId As Integer = 0, _
                                               Optional BranchID As Integer = 0, _
                                               Optional Account_code As String = "", _
                                               Optional updatetype As Integer = 0, _
                                               Optional notes_all As Double, _
                                               Optional project_id As Double, _
                                               Optional RegionID As Integer, _
                                               Optional withouopenening As Boolean)
    If withouopenening = True Then
      
        openingbalanceDes = " ﬂ‘› Õ—ﬂ… "
        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " SET opening_balance=  0"
        StrSQL = StrSQL & "  where  Account_code ='" & Account_code & "'"
        Cn.Execute StrSQL
        GoTo WithotFlag
    End If

    '0 balance Sheet
    '1 trial balances
    Dim openingbalacedate As Date
    ' getOpeningBalancedate DTPickerAccFrom.value, DTPickerAccTo.value, returnedfromdate, returnedTOdate, year(DTPickerAccFrom.value), openingbalacedate
    getOpeningBalancedate , , , , year(ToDate), openingbalacedate, continous
 
    'Dim StrSQL As String

    If openingbalacedate = FromDate Then
     
        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),"
        StrSQL = StrSQL & " balance= dbo.GetBalance('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & ")"

        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),"
            StrSQL = StrSQL & " balance= dbo.GetBalanceByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account)"
        End If
        If RegionID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByRegion('" & SQLDate(openingbalacedate) & "'," & RegionID & ", Account_code,last_account),"
            StrSQL = StrSQL & " balance= dbo.GetBalanceByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & RegionID & ", Account_code,last_account)"
        End If

        If BranchID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),"
            StrSQL = StrSQL & " balance= dbo.GetBalanceByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account)"
  
        End If
        
        If project_id <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByProject('" & SQLDate(openingbalacedate) & "'," & project_id & " )"
            ' StrSQL = StrSQL & " balance= dbo.GetBalanceByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(toDate) & "'," & BranchId & ", Account_code,last_account)"
  
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
     
        StrSQL = StrSQL & " set  balance= dbo.GetBalance('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account  , " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)  +   isnull(dbo.GetBalance('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code,last_account ," & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),0) "
    
        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
   
        End If
        If RegionID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & RegionID & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByRegion('" & SQLDate(openingbalacedate) & "'," & RegionID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByRegion('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & RegionID & ", Account_code,last_account),0) "
   
        End If

        If BranchID <> 0 Then
  
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByBranch('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0) "
  
        End If
 
        If project_id <> 0 Then
            StrSQL = " update ACCOUNTS"
            '  StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByProject('" & SQLDate(Fromdate) & "'," & project_id & " )"
            ' StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(Fromdate) & "','" & SQLDate(toDate) & "'," & BranchId & ", Account_code,last_account)"
  
            StrSQL = StrSQL & " set opening_balance= isnull(dbo.GetOpeningBalanceByProject('" & SQLDate(openingbalacedate) & "'," & project_id & " ),0)  +   isnull(dbo.GetBalanceByproject('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & project_id & ",'" & IIf(Account_code = "", "DEFAULT", Account_code) & "' ),0) "
  
        End If
            
    End If

    Dim FirstStr  As String
    Dim SecondStr As String
    FirstStr = StrSQL

    If updatetype = 0 Then ' „Ì“«‰Ì…
        StrSQL = StrSQL & " WHERE    (last_account )<> 1  and  (AccountTypes = 1) " 'ok
    ElseIf updatetype = 1 Then  ' „Ì“«‰
        StrSQL = StrSQL & " WHERE     (last_account = 1)  and  (AccountTypes = 1)"
 
    ElseIf updatetype = 2 Then ' ﬁ«∆„… «·œŒ·
        GoTo Part2
        'StrSQL = StrSQL & " WHERE   (last_account )<> 1  and   (AccountTypes = 2) "

    ElseIf updatetype = 3 Then   ' ﬂ‘› Õ”«»
   
        If getAccountTypes(Account_code) <> 1 Then ' ·Ê ﬂ«‰ Õ”«» „’—Ê› «Ê «Ì—«œ
            GoTo Part2
        End If
 
        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_code ='" & Account_code & "'"
  
        End If
 
    ElseIf updatetype = 4 Then ' Õ”«» «” «–

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Parent_Account_Code ='" & Account_code & "'  or Account_code ='" & Account_code & "'"
            
            If getAccountTypes(Account_code) <> 1 Then ' ·Ê ﬂ«‰ Õ”«» „’—Ê› «Ê «Ì—«œ
                GoTo Part2
            End If
        End If

    ElseIf updatetype = 30 Then  '  ﬂ‘› Õ”«» „Ã„⁄

        If (Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_Code in (" & Account_code & ")"
                
        End If
                  
    Else

    End If

    Cn.CommandTimeout = 10000
    Cn.Execute StrSQL

    If updatetype = 0 Or updatetype = 30 Or getAccountTypes(Account_code) = 1 Then
        Exit Function
    End If

    '***********************************************Part2
Part2:
   
    openingbalacedate = GetOpeningBalanceDateForType2(FromDate)
    
    If openingbalacedate = FromDate Then
     
        StrSQL = " update ACCOUNTS"
        StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),"
        StrSQL = StrSQL & " balance= dbo.GetBalance('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & ")"

        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),"
            StrSQL = StrSQL & " balance= dbo.GetBalanceByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account)"
  
        End If
        If RegionID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByRegion('" & SQLDate(openingbalacedate) & "'," & RegionID & ", Account_code,last_account),"
            StrSQL = StrSQL & " balance= dbo.GetBalanceByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & RegionID & ", Account_code,last_account)"
  
        End If

        If BranchID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),"
            StrSQL = StrSQL & " balance= dbo.GetBalanceByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account)"
  
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
     
        StrSQL = StrSQL & " set  balance= dbo.GetBalance('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "', Account_code,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),"
        StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) & "', Account_code,last_account),0)  +   isnull(dbo.GetBalance('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "', Account_code,last_account, " & IIf(SystemOptions.IsHiddenUser, 1, 0) & "),0) "
    
        If ActivityId <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByActivity('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & ActivityId & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByActivity('" & SQLDate(openingbalacedate) & "'," & ActivityId & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByActivity('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & ActivityId & ", Account_code,last_account),0) "
   
        End If
        If RegionID <> 0 Then
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByRegion('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & RegionID & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByRegion('" & SQLDate(openingbalacedate) & "'," & RegionID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByRegion('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & RegionID & ", Account_code,last_account),0) "
        End If

        If BranchID <> 0 Then
  
            StrSQL = " update ACCOUNTS"
            StrSQL = StrSQL & " set  balance= dbo.GetBalanceByBranch('" & SQLDate(FromDate) & "','" & SQLDate(ToDate) & "'," & BranchID & ", Account_code,last_account),"
            StrSQL = StrSQL & "  opening_balance= isnull(dbo.GetOpeningBalanceByBranch('" & SQLDate(openingbalacedate) & "'," & BranchID & ", Account_code,last_account),0)  +   isnull(dbo.GetBalanceByBranch('" & SQLDate(openingbalacedate) & "','" & SQLDate(FromDate1) & "'," & BranchID & ", Account_code,last_account),0) "
  
        End If
 
    End If

    If updatetype = 1 Then  ' „Ì“«‰
        'StrSQL = StrSQL & " WHERE     (last_account = 1) "
        StrSQL = StrSQL & " WHERE     (last_account = 1)   and  (AccountTypes = 2) "
    ElseIf updatetype = 2 Then ' ﬁ«∆„… «·œŒ·
        StrSQL = StrSQL & " WHERE   (last_account )<> 1  and   (AccountTypes = 2) "
    ElseIf updatetype = 3 Then ' ﬂ‘› Õ”«»
   
        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_code ='" & Account_code & "'"
  
        End If
 
    ElseIf updatetype = 4 Then ' Õ”«» «” «–

        If Trim(Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Parent_Account_Code ='" & Account_code & "'  or Account_code ='" & Account_code & "'"
        
        End If
          
    ElseIf updatetype = 30 Then  '  ﬂ‘› Õ”«» „Ã„⁄

        If (Account_code) <> "" Then
            StrSQL = StrSQL & "  where  Account_Code in (" & Account_code & ")"
                
        End If
                  
    Else

    End If

    Cn.CommandTimeout = 10000
    
    Cn.Execute StrSQL
    
    Debug.Print StrSQL
WithotFlag:
 
End Function


Public Function getAccountTypes(Account_code As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    On Error Resume Next
    sql = "select * from ACCOUNTS where  Account_Code ='" & Account_code & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    If rs.RecordCount > 0 Then
        getAccountTypes = IIf(IsNull(rs("AccountTypes").value), 0, rs("AccountTypes").value)
    Else
        getAccountTypes = 0
    End If

    rs.Close
 
End Function

Public Function GetusercodeByid(Optional UserID As Integer) As String

            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

 
 
   sql = " SELECT     TOP 100 PERCENT dbo.TblUsers.UserName, dbo.TblEmployee.Fullcode, dbo.TblUsers.UserID"
sql = sql & "  FROM         dbo.TblUsers LEFT OUTER JOIN"
sql = sql & "  dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
sql = sql & "  WHERE     (dbo.TblUsers.UserID = " & UserID & ")"
    
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetusercodeByid = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
  

    Else
        GetusercodeByid = ""
    End If

    rs.Close

End Function

Public Function GetEmployeeIDFromCode(Optional EmpCode As String, _
                                      Optional ByRef Emp_id As Integer, _
                                      Optional Emp_id1 As Integer = 0, _
                                      Optional ByRef EmpCode1 As String, Optional emp_Name As String, Optional flagAccount As Integer, Optional ByRef Account_code As String, Optional searchbyaccount As Boolean = False)
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    If Emp_id1 <> 0 Then
        sql = "select * from TblEmployee where Emp_Id= " & Emp_id1
    Else
 
        sql = "select * from TblEmployee where  FullCode ='" & EmpCode & "'"
    End If
 
    
If searchbyaccount = True Then
          If flagAccount = 0 Then
           sql = "select * from TblEmployee where  Account_Code ='" & Account_code & "'"
           ElseIf flagAccount = 1 Then
sql = "select * from TblEmployee where  Account_Code1 ='" & Account_code & "'"
             
           ElseIf flagAccount = 2 Then
             sql = "select * from TblEmployee where  Account_Code2 ='" & Account_code & "'"
            ElseIf flagAccount = 3 Then
             sql = "select * from TblEmployee where  Account_Code3 ='" & Account_code & "'"
            ElseIf flagAccount = 4 Then
sql = "select * from TblEmployee where  Account_Code4 ='" & Account_code & "'"
            ElseIf flagAccount = 5 Then
             sql = "select * from TblEmployee where  Account_Code5 ='" & Account_code & "'"
           End If
End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
Account_code = ""
    If rs.RecordCount > 0 Then
        Emp_id = IIf(IsNull(rs("Emp_Id").value), 0, rs("Emp_Id").value)
        EmpCode1 = IIf(IsNull(rs("FullCode").value), 0, rs("FullCode").value)
 
 
             If SystemOptions.UserInterface = ArabicInterface Then
             
                         emp_Name = IIf(IsNull(rs("Emp_Name").value), 0, rs("Emp_Name").value)
            
            Else
                            emp_Name = IIf(IsNull(rs("Emp_Namee").value), 0, rs("Emp_Namee").value)
            End If
            
            
            
           If flagAccount = 0 Then
             Account_code = IIf(IsNull(rs("Account_Code").value), 0, rs("Account_Code").value)
           ElseIf flagAccount = 1 Then
             Account_code = IIf(IsNull(rs("Account_Code1").value), 0, rs("Account_Code1").value)
             
           ElseIf flagAccount = 2 Then
             Account_code = IIf(IsNull(rs("Account_Code2").value), 0, rs("Account_Code2").value)
            ElseIf flagAccount = 3 Then
             Account_code = IIf(IsNull(rs("Account_Code3").value), 0, rs("Account_Code3").value)
            ElseIf flagAccount = 4 Then
             Account_code = IIf(IsNull(rs("Account_Code4").value), 0, rs("Account_Code4").value)
            ElseIf flagAccount = 5 Then
             Account_code = IIf(IsNull(rs("Account_Code5").value), 0, rs("Account_Code5").value)
           End If


    Else
        Emp_id = 0
    End If

    rs.Close

End Function

Public Function GetSalespersonDetail(Optional Emp_id As Integer) As String
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
 
     Dim StrSQL As String
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.*"
    StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TBLSalesRepData ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData.EmpID"
    StrSQL = StrSQL & " where dbo.TBLSalesRepData.EmpID=" & Emp_id
 
    StrSQL = StrSQL & " Order By Emp_Name ASC"
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
     
        GetSalespersonDetail = IIf(IsNull(rs("FullCode").value), 0, rs("FullCode").value)
 
    Else
     GetSalespersonDetail = ""
    End If

    rs.Close

End Function




Public Function GetCustomersDetail(Optional CusID As Variant, _
                                   Optional ByRef DefaultSalesPerson As Integer, _
                                   Optional fullcode As String = "", _
                                   Optional Custype As Integer = 0, _
                                   Optional DepitIntervalID As String = "D", _
                                   Optional DepitInterval As Integer = 0, Optional CusName As String, Optional creditlocked As Integer, Optional CurrncyID As Integer, Optional VATNO As String, Optional CPaymentType As Integer = 0, Optional chkTaxExempt As Boolean = False)
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    If fullcode = "" Then
        sql = "select * from TblCustemers where  CusID =" & CusID
        GoTo ll
    Else
        sql = "select * from TblCustemers where  FullCode ='" & fullcode & "'"
    End If

    If Custype <> 0 Then
        sql = sql & "  and  (type=" & Custype & " or CustomerandVendor=1)"
    End If

ll:
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        DefaultSalesPerson = IIf(IsNull(rs("EmpId").value), 0, rs("EmpId").value)
        CusID = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
        fullcode = IIf(IsNull(rs("FullCode").value), 0, rs("FullCode").value)
         creditlocked = IIf(IsNull(rs("creditlocked").value), 0, rs("creditlocked").value)
         CurrncyID = IIf(IsNull(rs("CurrncyID").value), MainCurrency(), rs("CurrncyID").value)
                    VATNO = IIf(IsNull(rs("VATNO").value), "", rs("VATNO").value)
                    CPaymentType = IIf(IsNull(rs("CPaymentType").value), 0, rs("CPaymentType").value)
        'creditlocked
        If SystemOptions.UserInterface = ArabicInterface Then
        CusName = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
        
        Else
        CusName = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
        End If
      '  chkTaxExempt = IIf(IsNull(rs("chkTaxExempt").value), False, rs("chkTaxExempt").value)
 DepitInterval = IIf(IsNull(rs("DepitInterval").value), 0, rs("DepitInterval").value)
 DepitIntervalID = IIf(IsNull(rs("DepitIntervalID").value), 0, rs("DepitIntervalID").value)
                 If DepitIntervalID = 0 Or DepitIntervalID = -1 Then
                        DepitIntervalID = "D"
                ElseIf DepitIntervalID = 1 Then
                        DepitIntervalID = "M"
                ElseIf DepitIntervalID = 2 Then
                        DepitIntervalID = "Y"
                End If
    Else
        DefaultSalesPerson = 0
 DepitInterval = 0
        CusID = 0
        DepitIntervalID = "D"
        
    End If

    rs.Close

End Function


Public Function GetCashCustomernamebycard(card As String, Optional ByRef Name = "", Optional ByRef phone As String = "", Optional ByRef discount As Double = 0) As String
         Dim RecordDate As Date
    Dim sql As String
    Dim rs As New ADODB.Recordset
  

 
        sql = "SELECT     *  From dbo.TblCusCsh  WHERE     (card = '" & card & "')"
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Name = IIf(IsNull(rs("name").value), "", rs("name").value)
    Else
    Name = IIf(IsNull(rs("namee").value), "", rs("namee").value)
    End If
    
        phone = IIf(IsNull(rs("tel").value), "", rs("tel").value)
        card = IIf(IsNull(rs("card").value), "", rs("card").value)
         discount = IIf(IsNull(rs("discount").value), "", rs("discount").value)
         RecordDate = IIf(IsNull(rs("RecordDate").value), "", rs("RecordDate").value)
         
       Dim NoofDays As Integer
       NoofDays = DateDiff("d", Date, RecordDate)
       If NoofDays < 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Â–« «·ﬂ«—  €Ì— „‰ ÂÌ ’·«ÕÌ …"
                Else
                MsgBox " Card Expire"
                End If
       discount = 0
       
       End If
       
   
    Else
        
    End If

    rs.Close

End Function


Public Function GetCashCustomernamebyphone(phone As String) As String
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

 
        sql = "SELECT     CashCustomerName, CashCustomerPhone From dbo.Transactions  WHERE     (CashCustomerPhone = '" & phone & "')"
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GetCashCustomernamebyphone = IIf(IsNull(rs("CashCustomerName").value), "", rs("CashCustomerName").value)
   
    Else
         GetCashCustomernamebyphone = ""
    End If

    rs.Close

End Function


Public Function GetProjectsDetail(Optional ID As Integer, _
                                   Optional ByRef DefaultSalesPerson As Integer, _
                                   Optional fullcode As String = "", _
                                   Optional Custype As Integer = 0)
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

 
        sql = "select * from projects  where  id =" & ID
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        DefaultSalesPerson = IIf(IsNull(rs("EmpId1").value), 0, rs("EmpId1").value)
   '     CusID = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
        fullcode = IIf(IsNull(rs("FullCode").value), 0, rs("FullCode").value)
 
    Else
        DefaultSalesPerson = 0
 
        fullcode = ""
    End If

    rs.Close

End Function



Public Function BankCollectData(BankID As Integer, _
                                Optional OperationType As Integer = -1, _
                                Optional branch_no As Integer = -1, _
                                Optional FromDate As Date, _
                                Optional ToDate As Date)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    StrSQL = " SELECT     SUM(dbo.TblBanksCollectDetails.[value]) AS totals"
    StrSQL = StrSQL & "  FROM         dbo.TblBanksCollect INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblBanksCollectDetails ON dbo.TblBanksCollect.id = dbo.TblBanksCollectDetails.TblBanksCollectId"
    StrSQL = StrSQL & " WHERE     (dbo.TblBanksCollect.bankid = " & BankID & ") "

    If OperationType <> -1 Then
        StrSQL = StrSQL & "AND (dbo.TblBanksCollect.OperationType = " & OperationType & ")"
    End If

    If branch_no <> -1 Then
        StrSQL = StrSQL & "AND (dbo.TblBanksCollect.branch_no = " & branch_no & ") "
    End If

    If Not (IsNull(FromDate)) Then
        StrSQL = StrSQL + " AND dbo.TblBanksCollect.RecordDate >=" & SQLDate(FromDate, True)
    End If

    If Not (IsNull(ToDate)) Then
        StrSQL = StrSQL + " AND dbo.TblBanksCollect.RecordDate<=" & SQLDate(ToDate, True)
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        BankCollectData = IIf(IsNull(rs("totals").value), 0, rs("totals").value)
    Else
        BankCollectData = 0
    
    End If

End Function

Public Function BankPendingCheques(BankID As Integer, _
                                   Optional OperationType As Integer = -1, _
                                   Optional branch_no As Integer = -1, _
                                   Optional FromDate As Date, _
                                   Optional ToDate As Date)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    'StrSQL = " SELECT     SUM(dbo.TblBanksCollectDetails.[value]) AS totals"
    'StrSQL = StrSQL & "  FROM         dbo.TblBanksCollect INNER JOIN"
    'StrSQL = StrSQL & "  dbo.TblBanksCollectDetails ON dbo.TblBanksCollect.id = dbo.TblBanksCollectDetails.TblBanksCollectId"
    ' StrSQL = StrSQL & " WHERE     (dbo.TblBanksCollect.bankid = " & BankID & ") "
 
    StrSQL = "SELECT     TOP 100 PERCENT SUM(ChequeValue) AS totals  "
    StrSQL = StrSQL & " From dbo.TblChecqueBoxContent1"
    StrSQL = StrSQL & " WHERE     (Payed = 0 OR"
    StrSQL = StrSQL & "  Payed IS NULL)"
    StrSQL = StrSQL & " and     (bankid = " & BankID & ") "

    If Not (IsNull(FromDate)) Then
        StrSQL = StrSQL + " AND DueDate>=" & SQLDate(FromDate, True)
    End If

    If Not (IsNull(ToDate)) Then
        StrSQL = StrSQL + " AND DueDate<=" & SQLDate(ToDate, True)
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        BankPendingCheques = IIf(IsNull(rs("totals").value), 0, rs("totals").value)
    Else
        BankPendingCheques = 0
    
    End If

End Function

Public Function BankDepositeData(BankID As Integer, _
                                 Optional box_or_bank As Integer = -1, _
                                 Optional branch_no As Integer = -1, _
                                 Optional FromDate As Date, _
                                 Optional ToDate As Date)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    'sql = "select * from ACCOUNTS where  Account_Code ='" & Account_code & "'"

    StrSQL = "SELECT     TOP 100 PERCENT SUM(dbo.TblBanksDepositeDetails.[value]) AS totals"
    StrSQL = StrSQL & " FROM         dbo.TblBanksDeposite INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblBanksDepositeDetails ON dbo.TblBanksDeposite.id = dbo.TblBanksDepositeDetails.TblBanksDepositeId LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblBoxesData ON dbo.TblBanksDepositeDetails.BoxID = dbo.TblBoxesData.BoxID"
    StrSQL = StrSQL & " WHERE     (dbo.TblBanksDeposite.bankid = " & BankID & ") "

    If box_or_bank <> -1 Then
        StrSQL = StrSQL & "AND (dbo.TblBanksDepositeDetails.box_or_bank = " & box_or_bank & ")"
    End If

    If branch_no <> -1 Then
        StrSQL = StrSQL & "AND (dbo.TblBanksDeposite.branch_no = " & branch_no & ") "
    End If
     
    If Not (IsNull(FromDate)) Then
        StrSQL = StrSQL + " AND dbo.TblBanksDeposite.RecordDate >=" & SQLDate(FromDate, True)
    End If

    If Not (IsNull(ToDate)) Then
        StrSQL = StrSQL + " AND dbo.TblBanksDeposite.RecordDate<=" & SQLDate(ToDate, True)
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        BankDepositeData = IIf(IsNull(rs("totals").value), 0, rs("totals").value)
    Else
        BankDepositeData = 0
    
    End If

End Function

Public Function GetActualAccountBalance(Account_code As String, _
                                        Optional BranchID As Integer = 0, _
                                        Optional FromDate As Date, _
                                        Optional ToDate As Date, _
                                        Optional ActivityId As Integer = 0, _
                                        Optional updatedata As Boolean = True, _
                                        Optional ByRef opening_balance As Double, _
                                        Optional notes_all As Double, _
                                        Optional ByRef account_name As String, _
                                        Optional ByRef Account_NameEng As String, _
                                        Optional ByRef Account_Serial As String) As Double

    If DateDiff("d", FromDate, ToDate) < 0 Then
        ToDate = FromDate
    End If
    
    If updatedata = True Then
        updateopeningbalanceNewFromsql FromDate, ToDate, True, ActivityId, BranchID, Account_code, 3
        
    End If
            
    Dim sql     As String
    Dim rs      As New ADODB.Recordset
    Dim Balance As Double
    sql = "select * from ACCOUNTS where  Account_Code ='" & Account_code & "'"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        account_name = IIf(IsNull(rs("account_name").value), "", rs("account_name").value)
        Account_NameEng = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
        Account_Serial = IIf(IsNull(rs("account_serial").value), "", rs("account_serial").value)
    
        Balance = IIf(IsNull(rs("opening_balance").value), 0, rs("opening_balance").value)
        opening_balance = Balance
        Balance = Balance + IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
        GetActualAccountBalance = Balance
        
    Else
        Balance = 0
 
    End If

    rs.Close
End Function

Public Function showLabel(NoteSerial As String, _
   oldnoteserial As String) As String

    If oldtxtNoteSerial1 = "" Then
        Exit Function
    End If
    If val(NoteSerial) = val(oldnoteserial) Then
        showLabel = ""
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            showLabel = " „  ⁄œÌ· —ﬁ„ Â–« «·„” ‰œ »‰«¡ ⁄·Ï  ⁄œÌ· «· «—ÌŒ Ê «·—ﬁ„ «·ﬁœÌ„ ··„” ‰œ ÂÊ :  " & oldnoteserial
            
        Else
            
            showLabel = " This Document Editing According To Date Editing Old No Is : " & oldnoteserial
        End If

    End If

End Function

Public Function getDocAccounts(ID As Integer, _
                               Optional ByRef Account_code1 As String, _
                               Optional ByRef Account_code2 As String, _
                               Optional ByRef Account_code3 As String, _
                               Optional ByRef Account_code4 As String, _
                               Optional ByRef Account_Code5 As String, _
                               Optional ByRef UseAccount_code1 As Integer, _
                               Optional ByRef UseAccount_code2 As Integer, _
                               Optional ByRef UseAccount_code3 As Integer, _
                               Optional ByRef UseAccount_code4 As Integer, _
                               Optional ByRef UseAccount_code5 As Integer, _
                               Optional ByRef UseCustomerAcc As Integer, Optional ByRef InvoiceTypeCodeID As Integer)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select * from TblDoCumentsTypes where  id =" & ID
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Account_code1 = IIf(IsNull(rs("Account_code1").value) Or CheckAccountToJE(IIf(IsNull(rs("Account_code1").value), "", rs("Account_code1").value)) = False, "", rs("Account_code1").value)
        Account_code2 = IIf(IsNull(rs("Account_code2").value) Or CheckAccountToJE(IIf(IsNull(rs("Account_code2").value), "", rs("Account_code2").value)) = False, "", rs("Account_code2").value)
        Account_code3 = IIf(IsNull(rs("Account_code3").value) Or CheckAccountToJE(IIf(IsNull(rs("Account_code3").value), "", rs("Account_code3").value)) = False, "", rs("Account_code3").value)
        Account_code4 = IIf(IsNull(rs("Account_code4").value) Or CheckAccountToJE(IIf(IsNull(rs("Account_code4").value), "", rs("Account_code4").value)) = False, "", rs("Account_code4").value)
        Account_Code5 = IIf(IsNull(rs("Account_code5").value) Or CheckAccountToJE(IIf(IsNull(rs("Account_code5").value), "", rs("Account_code5").value)) = False, "", rs("Account_code5").value)
        UseAccount_code1 = IIf(IsNull(rs("UseAccount_code1").value), 0, rs("UseAccount_code1").value)
        UseAccount_code2 = IIf(IsNull(rs("UseAccount_code2").value), 0, rs("UseAccount_code2").value)
        UseAccount_code3 = IIf(IsNull(rs("UseAccount_code3").value), 0, rs("UseAccount_code3").value)
        UseAccount_code4 = IIf(IsNull(rs("UseAccount_code4").value), 0, rs("UseAccount_code4").value)
        UseAccount_code5 = IIf(IsNull(rs("UseAccount_code5").value), 0, rs("UseAccount_code5").value)
        UseCustomerAcc = IIf(IsNull(rs("UseCustomerAcc").value), 0, rs("UseCustomerAcc").value)
        
                
  InvoiceTypeCodeID = IIf(IsNull(rs("InvoiceTypeCodeID").value), 0, rs("InvoiceTypeCodeID").value)
    End If

End Function

Public Function getemployeeCode(Emp_id As Integer, _
                                Optional ByRef Name As String, Optional ByRef datetype As Integer _
                                , Optional ByRef ddateh As String, Optional ByRef ddate As Date)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select * from TblEmployee where  Emp_ID= " & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        getemployeeCode = 0
        Name = ""
    Else
        getemployeeCode = IIf(IsNull(rs("fullcode").value), 0, rs("fullcode").value)

        If SystemOptions.UserInterface = ArabicInterface Then
            Name = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
        Else
            Name = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
        End If

If datetype = 0 Then ' ÕœÌÀ «ﬁ«„« 

ddate = IIf(IsNull(rs("DateEndekama").value), Date, rs("DateEndekama").value)
ddateh = IIf(IsNull(rs("DateEndekamaH").value), ToHijriDate(Date), rs("DateEndekamaH").value)

ElseIf datetype = 1 Then '—Œ’ ⁄„·
ddate = IIf(IsNull(rs("DateEndLinc").value), Date, rs("DateEndLinc").value)
ddateh = IIf(IsNull(rs("DateEndLincH").value), ToHijriDate(Date), rs("DateEndLincH").value)


ElseIf datetype = 2 Then 'ÃÊ«“« 

ddate = IIf(IsNull(rs("DateEndPasp").value), Date, rs("DateEndPasp").value)
ddateh = IIf(IsNull(rs("DateEndPasp").value), ToHijriDate(Date), ToHijriDate(rs("DateEndLincH").value))


ElseIf datetype = 3 Then 'Õ«›Ÿ… ‰›Ê”
ddateh = IIf(IsNull(rs("dateendpoketh").value), ToHijriDate(Date), rs("dateendpoketh").value)
ddate = IIf(IsNull(rs("dateendpoketh").value), Date, ToGregorianDate(rs("dateendpoketh").value))


 


End If




    End If

End Function

Public Function get_transactionData(FieldName As String, _
                                    Filedvalue As String, _
                                    RetuenFiledname As String, _
                                    Optional Transaction_Type As Integer = 0, Optional Transaction_Type2 As Integer = 0) As String

    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select * from Transactions where " & FieldName & "='" & Filedvalue & "'"

    If Transaction_Type <> 0 And Transaction_Type2 = 0 Then
        sql = sql & " and Transaction_Type=" & Transaction_Type
    End If

    If Transaction_Type <> 0 And Transaction_Type2 <> 0 Then
        sql = sql & " and  ( Transaction_Type=" & Transaction_Type & " or  Transaction_Type=" & Transaction_Type2 & ")"
    End If
     
'/If Transaction_Type = 30 Then
'sql = "SELECT     Transaction_ID From dbo.transactions Where (Transaction_Type = 30)"
'sql = SELECT     Transaction_ID, Transaction_Serial From dbo.transactions Where (Transaction_Type = 30) and Transaction_Serial=15

'End If

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        get_transactionData = 0
    Else
        get_transactionData = IIf(IsNull(rs(RetuenFiledname).value), "", rs(RetuenFiledname).value)
    End If

End Function
Public Function GetActualItemQty(Optional StoreID As Integer, _
                                 Optional FromDate As Date, _
                                 Optional ToDate As Date, _
                                 Optional ItemID As Long, _
                                 Optional UnitID As Long, _
                                 Optional itemsize As Long, _
                                 Optional ColorID As Long, _
                                 Optional ClassId As Long, Optional ExpiryDate As Variant, Optional myindex As Integer) As Double
    Dim StrSQL As String
    Dim RsDetails As New ADODB.Recordset
    Dim RowNum As Integer
    Dim LngNoteID As Long

    StrSQL = "SELECT     SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.Transaction_Details.Item_ID AS ItemID, "
    StrSQL = StrSQL & "  dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.TblItems.ItemCode,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemName,  dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ClassId,"
    StrSQL = StrSQL & "  dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblItemsSizes.SizeName AS SizeName, dbo.TblItemsColors.ColorName"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL + " where dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & ""
    StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & ""
    StrSQL = StrSQL & "   AND (dbo.Transactions.StoreID =" & StoreID & ")"
           
    StrSQL = StrSQL & "   AND (dbo.Transaction_Details.Item_ID =" & ItemID & ")"
   ' StrSQL = StrSQL & "   AND (dbo.Transaction_Details.UnitId =" & unitid & ")"
    StrSQL = StrSQL & "   AND (dbo.Transaction_Details.ItemSize =" & itemsize & ")"
    StrSQL = StrSQL & "   AND (dbo.Transaction_Details.ColorID =" & ColorID & ")"
    StrSQL = StrSQL & "   AND (dbo.Transaction_Details.ClassId =" & ClassId & ")"
        If (IsDate(ExpiryDate)) And myindex = 4 Then
        StrSQL = StrSQL & "   AND (dbo.Transaction_Details.ExpiryDate =" & SQLDate((ExpiryDate), True) & ")"
        
        ElseIf (myindex <> 4) Then
        StrSQL = StrSQL & "   AND (dbo.Transaction_Details.ExpiryDate  is null  )"
        End If

    StrSQL = StrSQL & "  GROUP BY dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID, dbo.TblStore.StoreName ,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemCode, dbo.TblItems.ItemName , dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize,"
    StrSQL = StrSQL & "   dbo.Transaction_Details.ClassId , dbo.TblItemsclasses.SizeName, dbo.TblItemsSizes.SizeName, dbo.TblItemsColors.ColorName"
    StrSQL = StrSQL & "  Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) <> 0)"
    'StrSQL = StrSQL & "  ORDER BY dbo.TblItems.ItemID,Transaction_Details.UnitId,Transaction_Details.ItemSize,Transaction_Details.ColorID,Transaction_Details.ClassId"
        StrSQL = StrSQL & "  ORDER BY  Item_ID ,Transaction_Details.ItemSize,Transaction_Details.ColorID,Transaction_Details.ClassId"

    Dim LngItemID As Long
    Dim LngUnitID As Long
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim UnitFactor As Double
GetUnitNoOfItems ItemID, UnitID, UnitFactor
If UnitFactor = 0 Then UnitFactor = 1
    If RsDetails.RecordCount > 0 Then
        GetActualItemQty = Round((IIf(IsNull(RsDetails("SUMQTY").value), 0, RsDetails("SUMQTY").value)) / UnitFactor, 5)
    Else
        GetActualItemQty = 0
    End If
 
End Function

Public Function GetQtyByBarcode(StoreID As Integer, Optional FromDate As Date, _
                                 Optional ToDate As Date, _
                                 Optional ParrtNoCode As String)
                   Dim StrSQL As String
    Dim RsDetails As New ADODB.Recordset
    Dim RowNum As Integer
    Dim LngNoteID As Long

    StrSQL = "SELECT     SUM(dbo.ItemsDetails.[Count] * dbo.ItemsDetails.EffectN) AS Qty, dbo.ItemsDetails.ParrtNoCode"
StrSQL = StrSQL + "   FROM         dbo.ItemsDetails INNER JOIN"
StrSQL = StrSQL + "                        dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID"

'StrSQL = StrSQL + "  WHERE     (dbo.Transactions.Transaction_Date >= CONVERT(DATETIME, '2015-01-01 00:00:00', 102) AND dbo.Transactions.Transaction_Date <= CONVERT(DATETIME,"

'StrSQL = StrSQL + "                        '2016-01-01 00:00:00', 102)) AND (dbo.Transactions.StoreID = " & StoreId & ")"
    StrSQL = StrSQL + "  where dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & ""
    StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & ""
    StrSQL = StrSQL & "   AND (dbo.Transactions.StoreID =" & StoreID & ")"
StrSQL = StrSQL & "   AND (dbo.ItemsDetails.ParrtNoCode ='" & ParrtNoCode & "')"


StrSQL = StrSQL + "  GROUP BY dbo.ItemsDetails.ParrtNoCode"
 
 
    Dim LngItemID As Long
    Dim LngUnitID As Long
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim UnitFactor As Double
'GetUnitNoOfItems ItemID, unitid, UnitFactor
    If RsDetails.RecordCount > 0 Then
        GetQtyByBarcode = Round((IIf(IsNull(RsDetails("Qty").value), 0, RsDetails("Qty").value)))
    Else
        GetQtyByBarcode = 0
    End If
        
        
                                 

End Function
Public Function GetUnitFactor(ByVal ItemID As Long, ByVal UnitID As Long) As Double
    On Error GoTo Fallback
    Dim rs As New ADODB.Recordset
    Dim sql As String

    If UnitID > 0 Then
        sql = "SELECT TOP 1 ISNULL(UnitFactor,1) AS F " & _
              "FROM dbo.TblItemsUnits WHERE ItemID=" & ItemID & " AND UnitID=" & UnitID
        rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF And Not IsNull(rs!F) Then
            GetUnitFactor = rs!F: rs.Close: Exit Function
        End If
        rs.Close
    End If

    ' ???? ???? ????? (????? ?? ?????? = ???? 1 ?? ??? ????)
    sql = "SELECT TOP 1 ISNULL(UnitFactor,1) AS F " & _
          "FROM dbo.TblItemsUnits WHERE ItemID=" & ItemID & " ORDER BY UnitFactor ASC"
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF And Not IsNull(rs!F) Then GetUnitFactor = rs!F Else GetUnitFactor = 1
    rs.Close
    Exit Function
Fallback:
    GetUnitFactor = 1
End Function


Public Function GetCurrentGardEmployee(Optional StoreID As Integer) As String

    Dim StrSQL  As String
    Dim i As Integer
    Dim ReturnStr As String
    StrSQL = "SELECT     Store_or_EmpName"
    StrSQL = StrSQL + " from dbo.TblStrartGardDetails"
    StrSQL = StrSQL + " WHERE     (TblStrartGardId ="
    StrSQL = StrSQL + " (SELECT     MAX(id) AS id"
    StrSQL = StrSQL + " FROM         dbo.TblStrartGard)) AND (Store_or_Emp = 1)"
    Dim rs As New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            ReturnStr = ReturnStr & CHR(13) & IIf(IsNull(rs("Store_or_EmpName").value), "", rs("Store_or_EmpName").value)
            rs.MoveNext
        Next i
              
    End If

    GetCurrentGardEmployee = ReturnStr
End Function
Public Function GetActualItemQtyNew(Optional StoreID As Integer, _
                                 Optional FromDate As Date, _
                                 Optional ToDate As Date, _
                                 Optional ItemID As Long, _
                                 Optional UnitID As Long, _
                                 Optional itemsize As Long, _
                                 Optional ColorID As Long, _
                                 Optional ClassId As Long, _
                                 Optional ExpiryDate As Variant, _
                                 Optional myindex As Integer) As Double

    Dim StrSQL As String
    Dim rs As New ADODB.Recordset
    Dim BaseQty As Double
    Dim TargetFactor As Double

    '⁄«„·  ÕÊÌ· «·ÊÕœ… «·„ÿ·Ê»… (·Ê UnitID=ÊÕœ… «·»Ì⁄ „À·«)
    TargetFactor = GetUnitFactor(ItemID, UnitID)
    If TargetFactor = 0 Then TargetFactor = 1

    StrSQL = ""
    StrSQL = StrSQL & "SELECT SUM( "
    StrSQL = StrSQL & "  (CAST(ISNULL(td.Quantity,0) AS decimal(18,6)) "
    StrSQL = StrSQL & "   / NULLIF(CAST(ISNULL(dbo.GetItemUnitFactor(td.Item_ID, td.UnitId),1) AS decimal(18,6)),0) "
    StrSQL = StrSQL & "  ) * CAST(ISNULL(tt.StockEffect,0) AS decimal(18,6)) "
    StrSQL = StrSQL & ") AS SumBaseQty "
    StrSQL = StrSQL & "FROM dbo.Transactions t "
    StrSQL = StrSQL & "INNER JOIN dbo.Transaction_Details td ON t.Transaction_ID = td.Transaction_ID "
    StrSQL = StrSQL & "INNER JOIN dbo.TransactionTypes tt ON t.Transaction_Type = tt.Transaction_Type "
    StrSQL = StrSQL & "WHERE 1=1 "
    StrSQL = StrSQL & "AND tt.StockEffect <> 0 "
    StrSQL = StrSQL & "AND t.StoreID = " & StoreID & " "
    StrSQL = StrSQL & "AND t.Transaction_Date < " & SQLDate(DateAdd("d", 1, ToDate), True) & " "

    StrSQL = StrSQL & "AND td.Item_ID = " & ItemID & " "
    StrSQL = StrSQL & "AND ISNULL(td.ItemSize,-1) = " & IIf(IsNull(itemsize), "-1", CStr(itemsize)) & " "
    StrSQL = StrSQL & "AND ISNULL(td.ColorID,-1) = " & IIf(IsNull(ColorID), "-1", CStr(ColorID)) & " "
    StrSQL = StrSQL & "AND ISNULL(td.ClassId,-1) = " & IIf(IsNull(ClassId), "-1", CStr(ClassId)) & " "

    If IsDate(ExpiryDate) And myindex = 4 Then
        StrSQL = StrSQL & "AND CONVERT(date, td.ExpiryDate) = " & SQLDate(CDate(ExpiryDate), True) & " "
    End If

    rs.Open StrSQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rs.EOF And Not IsNull(rs.Fields(0).value) Then
        BaseQty = rs.Fields(0).value
    Else
        BaseQty = 0
    End If
    rs.Close

    'BaseQty Â‰« »ÊÕœ… «·√”«” -> ÕÊ¯· ··ÊÕœ… «·„ÿ·Ê»…
    GetActualItemQtyNew = Round(BaseQty / TargetFactor, 5)

End Function

Public Function GetActualItemQtyOld(Optional StoreID As Integer, _
                                 Optional FromDate As Date, _
                                 Optional ToDate As Date, _
                                 Optional ItemID As Long, _
                                 Optional UnitID As Long, _
                                 Optional itemsize As Long, _
                                 Optional ColorID As Long, _
                                 Optional ClassId As Long, Optional ExpiryDate As Variant, Optional myindex As Integer) As Double
    Dim StrSQL As String
    Dim RsDetails As New ADODB.Recordset
    Dim RowNum As Integer
    Dim LngNoteID As Long

    StrSQL = "SELECT     SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.Transaction_Details.Item_ID AS ItemID, "
    StrSQL = StrSQL & "  dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.TblItems.ItemCode,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemName,  dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ClassId,"
    StrSQL = StrSQL & "  dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblItemsSizes.SizeName AS SizeName, dbo.TblItemsColors.ColorName"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL + " where dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & ""
    StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & ""
    StrSQL = StrSQL & "   AND (dbo.Transactions.StoreID =" & StoreID & ")"
           
    StrSQL = StrSQL & "   AND (dbo.Transaction_Details.Item_ID =" & ItemID & ")"
   ' StrSQL = StrSQL & "   AND (dbo.Transaction_Details.UnitId =" & unitid & ")"
    StrSQL = StrSQL & "   AND (dbo.Transaction_Details.ItemSize =" & itemsize & ")"
    StrSQL = StrSQL & "   AND (dbo.Transaction_Details.ColorID =" & ColorID & ")"
    StrSQL = StrSQL & "   AND (dbo.Transaction_Details.ClassId =" & ClassId & ")"
        If (IsDate(ExpiryDate)) And myindex = 4 Then
        StrSQL = StrSQL & "   AND (dbo.Transaction_Details.ExpiryDate =" & SQLDate((ExpiryDate), True) & ")"
        
        ElseIf (myindex <> 4) Then
        StrSQL = StrSQL & "   AND (dbo.Transaction_Details.ExpiryDate  is null  )"
        End If

    StrSQL = StrSQL & "  GROUP BY dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID, dbo.TblStore.StoreName ,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemCode, dbo.TblItems.ItemName , dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize,"
    StrSQL = StrSQL & "   dbo.Transaction_Details.ClassId , dbo.TblItemsclasses.SizeName, dbo.TblItemsSizes.SizeName, dbo.TblItemsColors.ColorName"
    StrSQL = StrSQL & "  Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) <> 0)"
    'StrSQL = StrSQL & "  ORDER BY dbo.TblItems.ItemID,Transaction_Details.UnitId,Transaction_Details.ItemSize,Transaction_Details.ColorID,Transaction_Details.ClassId"
        StrSQL = StrSQL & "  ORDER BY  Item_ID ,Transaction_Details.ItemSize,Transaction_Details.ColorID,Transaction_Details.ClassId"

    Dim LngItemID As Long
    Dim LngUnitID As Long
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim UnitFactor As Double
GetUnitNoOfItems ItemID, UnitID, UnitFactor
If UnitFactor = 0 Then UnitFactor = 1
    If RsDetails.RecordCount > 0 Then
        GetActualItemQtyOld = Round((IIf(IsNull(RsDetails("SUMQTY").value), 0, RsDetails("SUMQTY").value)) / UnitFactor, 5)
    Else
        GetActualItemQtyOld = 0
    End If
 
End Function
Public Function CheckAccountHaveDestributions(StrAccountCode As String) As Boolean
    CheckAccountHaveDestributions = False
    Dim StrSQL  As String
    'StrSQL = "SELECT     dbo.TblAccountsDestributions.AccountMaster, dbo.TblAccountsDestributionsDetails.ACode, dbo.TblAccountsDestributionsDetails.Percentage, "
    '   StrSQL = StrSQL + "  dbo.TblAccountsDestributions.DistType , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
    '   StrSQL = StrSQL + " FROM         dbo.TblAccountsDestributions INNER JOIN"
    '   StrSQL = StrSQL + " dbo.TblAccountsDestributionsDetails ON"
    '   StrSQL = StrSQL + " dbo.TblAccountsDestributions.TblAccountsDestributionsid = dbo.TblAccountsDestributionsDetails.TblAccountsDestributionsid INNER JOIN"
    '   StrSQL = StrSQL + "  dbo.TblBranchesData ON dbo.TblAccountsDestributionsDetails.ACode = dbo.TblBranchesData.branch_id"
    '   StrSQL = StrSQL + " WHERE     (dbo.TblAccountsDestributions.DistType IS NULL) AND (dbo.TblAccountsDestributions.AccountMaster = N'" & StrAccountCode & "')"
 
    StrSQL = "SELECT     dbo.TblAccountsDestributions.AccountMaster, dbo.TblAccountsDestributionsDetails.ACode, dbo.TblAccountsDestributionsDetails.Percentage, "
    StrSQL = StrSQL + " dbo.TblAccountsDestributions.DistType"
    StrSQL = StrSQL + " FROM         dbo.TblAccountsDestributions INNER JOIN"
    StrSQL = StrSQL + " dbo.TblAccountsDestributionsDetails ON"
    StrSQL = StrSQL + " dbo.TblAccountsDestributions.TblAccountsDestributionsid = dbo.TblAccountsDestributionsDetails.TblAccountsDestributionsid"
    StrSQL = StrSQL + "  WHERE     (dbo.TblAccountsDestributions.AccountMaster = N'" & StrAccountCode & "')"

    Dim rs As New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        CheckAccountHaveDestributions = True
    Else
        CheckAccountHaveDestributions = False
              
    End If
            
End Function

Public Function MonthLastDay(ByVal dCurrDate As Date) As Date
    
    Dim tmpDate As Date
 
    'Get First Day of Month
    tmpDate = DateAdd("d", (day(dCurrDate) - 1) * -1, dCurrDate)

    'Get First Day of Next Month
    tmpDate = DateAdd("m", 1, tmpDate)

    'Get Last Day of This Month
    tmpDate = DateAdd("d", -1, tmpDate)

    MonthLastDay = tmpDate
    Exit Function
    
    Dim dFirstDayNextMonth As Date
  
    On Error GoTo lbl_Error
 
    MonthLastDay = Empty
    dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
    MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
  
    Exit Function
lbl_Error:
    MsgBox Err.Description, vbOKOnly + vbExclamation
End Function

Public Function CountAs(str As String) As Integer
    Dim count As Integer

    For i = 1 To Len(str)

        If mId$(str, i, 1) = "a" Then count = count + 1
    Next

    CountAs = count
End Function

Function GetIntervalsFullData(AccountIntervalID As Integer, Optional ByRef StartDate As Date, Optional ByRef EndDate As Date)
    'x
    Dim StrSQL  As String

    StrSQL = " SELECT     TOP 100 PERCENT AccountIntervalID, StartDate, EndDate"
    StrSQL = StrSQL & "  from dbo.TblAccountIntervals"

    StrSQL = StrSQL + " where  AccountIntervalID=" & AccountIntervalID
    StartDate = Date
    EndDate = Date
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        StartDate = IIf(IsNull(RsUnitData("StartDate").value), Date, (RsUnitData("StartDate").value))
        EndDate = IIf(IsNull(RsUnitData("EndDate").value), Date, (RsUnitData("EndDate").value))
             
    End If

    RsUnitData.Close
              
End Function

Function GetProductionInventoryId(BranchID As Integer) As Integer
    'x
    Dim StrSQL  As String

    StrSQL = " SELECT     StoreID From dbo.Transactions"
    StrSQL = StrSQL + " where    (Transaction_Type = 28) and   BranchId=" & BranchID
            
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsUnitData.RecordCount) > 0 Then
                 
        GetProductionInventoryId = IIf(IsNull(RsUnitData("StoreID").value), 0, (RsUnitData("StoreID").value))
    Else
        GetProductionInventoryId = 0
    End If

    RsUnitData.Close
              
End Function

 
 Function GetUnitNoOfItems(LngCurItemID As Variant, LngUnitID As Long, Optional ByRef UnitFactor As Double)
    'x
    Dim StrSQL  As String

    StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
    StrSQL = StrSQL + " AND UnitID=" & LngUnitID
    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                 
        UnitFactor = IIf(IsNull(RsUnitData("UnitFactor").value), 1, (RsUnitData("UnitFactor").value))
               
    End If

    RsUnitData.Close
End Function
 Public Function GetDefaultItemUnit(ItemID As Long, _
                                   Optional ByRef UnitID As Long, _
                                   Optional ByRef UnitName As String, Optional ByRef UnitFactor As Double)
    Dim RsUnitData As New ADODB.Recordset
    
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName,TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
    Else
        StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitNamee UnitName," & "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
    End If
    StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID =" & "TblUnites.UnitID"
    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & ItemID
    StrSQL = StrSQL + " AND DefaultUnit=1"
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
        UnitID = IIf(IsNull(RsUnitData("UnitID").value), 0, RsUnitData("UnitID").value)
        UnitName = IIf(IsNull(RsUnitData("UnitName").value), "", RsUnitData("UnitName").value)
        UnitFactor = IIf(IsNull(RsUnitData("UnitFactor").value), 0, RsUnitData("UnitFactor").value)
    End If

    RsUnitData.Close
    Set RsUnitData = Nothing
         
End Function



Public Function DefaultPrinter() As String

    Dim strReturn As String
    Dim intReturn As Integer

    strReturn = Space(255)

    'This gets the default printer name
    intReturn = GetProfileString("Windows", ByVal "device", "", strReturn, Len(strReturn))

    If intReturn Then
        strReturn = (left(strReturn, InStr(strReturn, ",") - 1))
    End If

    DefaultPrinter = strReturn

End Function

Public Function setfoxyNo(Filedname As String) As String
    Dim lastNo As String
    Dim sql As String
    lastNo = CStr(new_id("foxy", Filedname, "", True))

    sql = "update    foxy"
    sql = sql & " Set " & Filedname & " = " & lastNo
    Cn.Execute sql
 
    setfoxyNo = lastNo
End Function

Public Function GetItemOrderNo(ItemID As Long) As String
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "Select MIN(oRDER_NO) AS ORDERnOMin From QryGardWithOrderNo where   ITEMID=" & ItemID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        GetItemOrderNo = IIf(IsNull(Rs1("ORDERnOMin").value), "", Rs1("ORDERnOMin").value)
    End If

End Function
Public Function GetallChilddata(group_id As Integer) As String
    On Error Resume Next
 
    Dim Rs1 As ADODB.Recordset
Dim i As Integer
    Dim sql As String
    Dim str As String
   Set Rs1 = New ADODB.Recordset
 
    sql = "SELECT * from Groups  where ParentID=" & group_id
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
                For i = 1 To Rs1.RecordCount
                GetallChilddata = GetallChilddata & GetallChilddata(IIf(IsNull(Rs1("GroupID").value), 0, Rs1("GroupID").value)) & "," & IIf(IsNull(Rs1("GroupID").value), 0, Rs1("GroupID").value)
                
                Rs1.MoveNext
                Next i
        Rs1.Close
    End If
 
End Function
Public Function GetPrefix(group_id As Double, _
                          Optional tablename As String) As String
    On Error Resume Next
    Dim fullCodeAll     As String
    
    Dim GroupCode       As String
    Dim ParentGroupCode As String
    Dim ParentID        As Double
    fullCodeAll = ""
    txtid.text = ""

    If SystemOptions.WorkWithBarCodeParent = False Then
        GetGroupData group_id, GroupCode, , ParentGroupCode, ParentID, tablename
        GetPrefix = GroupCode
        Exit Function
    End If
 
    If SystemOptions.WorkWithBarCodeParent = True Then
 
        GetGroupData group_id, GroupCode, , ParentGroupCode, ParentID, tablename
      
        If group_id = 0 Or group_id = 1 Then
            GetPrefix = GroupCode
            Exit Function
        End If

        '      fullCodeAll = fullCodeAll & SystemOptions.itemSeprator & GroupCode
        GetPrefix = GetPrefix(ParentID, tablename) & SystemOptions.itemSeprator & GroupCode
 
    End If
 
End Function
Public Function GetItemIDExpiry(ItemID As Long, _
                             Optional ByRef EXpirType As Integer, _
                             Optional ByRef EXpireValue As Integer)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
 
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from  TblItems  where ItemID=" & ItemID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            EXpirType = IIf(IsNull(Rs1("EXpirType").value), EXpirType, Rs1("EXpirType").value)
        EXpireValue = IIf(IsNull(Rs1("EXpireValue").value), EXpireValue, Rs1("EXpireValue").value)
                    If EXpireValue <> 0 Then
                                 EXpireValue = EXpireValue
                     End If
                     
        Else
        
                 
    End If

    Rs1.Close
   Exit Function
End Function
Public Function GetGroupData(GroupID As Double, _
                             Optional ByRef GroupCode As String, _
                             Optional ByRef GroupName As String, _
                             Optional ByRef ParentGroupCode As String, _
                             Optional ByRef ParentID As Double, _
                             Optional tablename As String, _
                             Optional ByRef EXpirType As Integer, _
                             Optional ByRef EXpireValue As Integer, _
                             Optional ByRef OverHead As Double, _
                             Optional ByRef chkTaxExempt As Integer)
    Dim rs  As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    If GroupID = 1 Or GroupID = 0 Then
        Exit Function
    End If
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from " & tablename & " where GroupID=" & GroupID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        GroupCode = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
        GroupName = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
        ParentID = IIf(IsNull(Rs1("ParentID").value), 0, Rs1("ParentID").value)
        EXpirType = IIf(IsNull(Rs1("EXpirType").value), -1, Rs1("EXpirType").value)
        EXpireValue = IIf(IsNull(Rs1("EXpireValue").value), -1, Rs1("EXpireValue").value)
        chkTaxExempt = IIf(IsNull(Rs1("chkTaxExempt").value), 0, Rs1("chkTaxExempt").value)
        
        OverHead = IIf(IsNull(Rs1("OverHead").value), 0, Rs1("OverHead").value)
                
    End If

    Rs1.Close
    Exit Function
    If ParentID = 1 Then
        Exit Function
    End If
    sql = "SELECT * from Groups where GroupID=" & ParentID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        ParentGroupCode = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
        ParentID = IIf(IsNull(Rs1("ParentID").value), 0, Rs1("ParentID").value)
    End If
    Rs1.Close
 
End Function
'
Public Function GetUserData(UserID As Long, _
                            Optional ByRef usertype As Integer, _
                            Optional ByRef BranchID As Integer, _
                            Optional ByRef StoreID As Integer, _
                            Optional ByRef BoxID As Integer, _
                            Optional ByRef BankID As Integer, _
                            Optional ByRef EmpID As Integer, _
                            Optional ByRef UserName As String, Optional CUSTID As Integer, Optional StoreId1 As Integer, Optional CUSTID1 As Integer, Optional ByRef boxid1 As Integer)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
On Error Resume Next
    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from TblUsers where UserID=" & UserID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then




AllowSelectEmp = IIf(IsNull(Rs1("AllowSelectEmp").value), 0, Rs1("AllowSelectEmp").value)
HidLowering = IIf(IsNull(Rs1("HidLowering").value), 0, Rs1("HidLowering").value)

        usertype = IIf(IsNull(Rs1("UserType").value), 0, Rs1("UserType").value)
        BranchID = IIf(IsNull(Rs1("BranchId").value), 0, Rs1("BranchId").value)
        StoreID = IIf(IsNull(Rs1("StoreID").value), 0, Rs1("StoreID").value)
       StoreId1 = IIf(IsNull(Rs1("StoreID1").value), 0, Rs1("StoreID1").value)

        BoxID = IIf(IsNull(Rs1("BoxID").value), 0, Rs1("BoxID").value)
           boxid1 = IIf(IsNull(Rs1("BoxID1").value), 0, Rs1("BoxID1").value)
                
        BankID = IIf(IsNull(Rs1("BankID").value), 0, Rs1("BankID").value)
        EmpID = IIf(IsNull(Rs1("EmpID").value), 0, Rs1("EmpID").value)
        UserName = IIf(IsNull(Rs1("UserName").value), "", Rs1("UserName").value)
        ChangePW = IIf(IsNull(Rs1("ChangePW").value), 0, Rs1("ChangePW").value)
        CUSTID = IIf(IsNull(Rs1("custid").value), 2, Rs1("custid").value)
        CUSTID1 = IIf(IsNull(Rs1("custid1").value), 2, Rs1("custid1").value)
       

        If UserID = 1 Then
            BranchID = 1
            StoreID = 1
       '     BoxID = 1
       '     BankID = 1
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
Public Function GetBranchData(BranchID As Integer, _
                              Optional ByRef StoreID As Integer, _
                              Optional ByRef BoxID As Integer, _
                              Optional ByRef BankID As Integer, _
                              Optional ByRef ActivityTypeId As Integer, _
                               Optional ByRef branch_name As String, _
                                Optional ByRef branch_namee As String)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT branch_nameE,branch_name,branch_id,ActivityTypeId  from TblBranchesData where Branch_Id=" & BranchID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        ActivityTypeId = IIf(IsNull(Rs1("ActivityTypeId").value), 0, Rs1("ActivityTypeId").value)
        
        branch_name = IIf(IsNull(Rs1("branch_name").value), 0, Rs1("branch_name").value)
        branch_namee = IIf(IsNull(Rs1("branch_nameE").value), 0, Rs1("branch_nameE").value)
        
    Else
        ActivityTypeId = 0
        
    End If

    Rs1.Close
 
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from TblStore where BranchId=" & BranchID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        StoreID = IIf(IsNull(Rs1("StoreID").value), "", Rs1("StoreID").value)
    Else
        StoreID = 0
    End If

    Rs1.Close
 
    sql = "SELECT * from TblBoxesData where BranchId=" & BranchID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        BoxID = IIf(IsNull(Rs1("BoxID").value), "", Rs1("BoxID").value)
    Else
        BoxID = 0
    End If

    Rs1.Close
 
    sql = "SELECT * from BanksData where BranchId=" & BranchID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        BankID = IIf(IsNull(Rs1("BankID").value), "", Rs1("BankID").value)
    Else
        BankID = 0
    End If

    Rs1.Close
 
End Function
Public Function GeTuserIDByEmpCode(fullcode As String) As Integer
    Dim rs  As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    If fullcode = "0000" Then
        GeTuserIDByEmpCode = 1
        Exit Function
    End If

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    
    sql = " SELECT     dbo.TblUsers.UserID, dbo.TblEmployee.Fullcode"
    sql = sql & "  FROM         dbo.TblUsers INNER JOIN"
    sql = sql & "  dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
    sql = sql & "   WHERE     (dbo.TblEmployee.Fullcode = N'" & fullcode & "')"
    
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        GeTuserIDByEmpCode = IIf(IsNull(Rs1("UserID").value), "", Rs1("UserID").value)
    Else
        GeTuserIDByEmpCode = 0
    End If
    
    Rs1.Close
End Function

Public Function GeTEmpIDByEmpCode(fullcode As String, Optional all As Boolean) As Integer
    Dim rs  As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    If fullcode = "0000" Then
        GeTEmpIDByEmpCode = 1
        Exit Function
    End If

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    
    sql = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Fullcode"

    If all = True Then
        sql = sql & "  FROM         dbo.TblEmployee  "

    Else
        sql = sql & "  FROM         dbo.TblUsers INNER JOIN"
        sql = sql & "  dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
    End If
    sql = sql & "   WHERE     (dbo.TblEmployee.Fullcode = N'" & fullcode & "')"
    
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        GeTEmpIDByEmpCode = IIf(IsNull(Rs1("Emp_ID").value), "", Rs1("Emp_ID").value)
    Else
        GeTEmpIDByEmpCode = 0
    End If
    
    Rs1.Close
End Function

Public Function GeBranchInfo(tablename As String, _
                             Filedname As String, _
                             ID As Integer) As Integer
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from " & tablename & " where  " & Filedname & "=" & ID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        BranchID = IIf(IsNull(Rs1("BranchId").value), "", Rs1("BranchId").value)
    Else
        BranchID = 0
    End If

    GeBranchInfo = BranchID
    Rs1.Close
End Function

Public Function GetID(tablename As String, _
                      Filedname As String, _
                      ReturnFiledname As String, _
                      value As String) As String
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from " & tablename & " where  " & Filedname & "='" & value & "'"
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        GetID = IIf(IsNull(Rs1(ReturnFiledname).value), "", Rs1(ReturnFiledname).value)
    Else
        GetID = 0
    End If
   
    Rs1.Close
End Function

Public Function GetInventoryBranch(StoreID As Integer)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from TblStore where StoreID=" & StoreID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        BranchID = IIf(IsNull(Rs1("BranchId").value), "", Rs1("BranchId").value)
    Else
        BranchID = 0
    End If

    GetInventoryBranch = BranchID
    Rs1.Close
End Function

Public Function GetEstimatedCost(Optional ByRef ItemID As Integer, _
                                 Optional ByRef GroupID As Variant, _
                                 Optional ByRef EstimatedCost As Double)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
'    sql = "SELECT     dbo.UnitsIndustrialCost.unitid, dbo.UnitsIndustrialCost.EstimatedTotalCost * dbo.currency.rate AS EstimatedCost"
''    sql = sql & "  FROM         dbo.UnitsIndustrialCost INNER JOIN"
 '   sql = sql & "  dbo.currency ON dbo.UnitsIndustrialCost.CurrencyID = dbo.currency.id"
 '   sql = sql & "  WHERE     (dbo.UnitsIndustrialCost.unitid = " & GroupID & ")"
 
 sql = " SELECT     SUM(dbo.UnitsIndustrialCostDetails.Cost * dbo.currency.rate) AS EstimatedCost"
  sql = sql & "  FROM         dbo.UnitsIndustrialCost INNER JOIN"
  sql = sql & "  dbo.currency ON dbo.UnitsIndustrialCost.CurrencyID = dbo.currency.id INNER JOIN"
  sql = sql & "  dbo.UnitsIndustrialCostDetails ON dbo.UnitsIndustrialCost.id = dbo.UnitsIndustrialCostDetails.UnitsIndustrialCostId"
  sql = sql & "   Where (dbo.UnitsIndustrialCost.unitid = " & GroupID & ")"

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        EstimatedCost = IIf(IsNull(Rs1("EstimatedCost").value), 0, Rs1("EstimatedCost").value)
    Else
        EstimatedCost = 0
    End If
     
    Rs1.Close
End Function

Public Function GetItemData(ItemID As Long, _
                            Optional ByRef itemcode As String, _
                            Optional ByRef ItemName As String, _
                            Optional ByRef GroupID As Variant)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from TblItems where ItemID=" & ItemID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        itemcode = IIf(IsNull(Rs1("ItemCode").value), "", Rs1("ItemCode").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                        ItemName = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                        
                 Else
                      ItemName = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
                        
                        
                 End If
 GroupID = IIf(IsNull(Rs1("GroupId").value), 0, Rs1("GroupId").value)
    End If
  
    Rs1.Close
End Function

Public Function GetActivityBranchs(ActivityTypeId As Integer, _
                                   filed As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql           As String
    Dim AccountsCodes As String
    AccountsCodes = ""
    sql = "Select  " & filed & " from branches where ActivityTypeId=" & ActivityTypeId
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        GetActivityBranchs = ""
        Exit Function
    End If
    For i = 1 To Rs3.RecordCount
        AccountsCodes = AccountsCodes & Rs3(filed).value & ","
        Rs3.MoveNext
    Next i

    AccountsCodes = mId(AccountsCodes, 1, Len(AccountsCodes) - 1)
    GetActivityBranchs = AccountsCodes
    Rs3.Close

End Function

Public Sub ShowItemsStatusReport(m_PrintTarget As PrintTarget, _
                                 Optional ToDate As Date)
    Dim MySQL        As String
    Dim RsData       As New ADODB.Recordset
    Dim xApp         As New CRAXDRT.Application
    Dim xReport      As CRAXDRT.Report
    Dim CViewer      As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo

    If Dir(App.path & "\Reports\Inventory\" & "ItemsStatus.rpt") = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    MySQL = "SELECT     dbo.QRyOrdersData.order_no, dbo.Items_Status_totals1.StoreNamex1, dbo.Items_Status_totals1.totalPurchaseQtyx1, "
    MySQL = MySQL & "   dbo.Items_Status_totals1.totalPurchaseValuex1, dbo.Items_Status_totals1.totalSalesQtyx1, dbo.Items_Status_totals1.SaleValueValuex1,"
    MySQL = MySQL & "  dbo.Items_Status_totals1.SaleValuex1, dbo.Items_Status_totals2.StoreNamex2, dbo.Items_Status_totals2.totalPurchaseQtyx2,"
    MySQL = MySQL & "   dbo.Items_Status_totals2.totalPurchaseValuex2, dbo.Items_Status_totals2.totalSalesQtyx2, dbo.Items_Status_totals2.SaleValueValuex2,"
    MySQL = MySQL & "   dbo.Items_Status_totals3.SaleValueValuex3, dbo.Items_Status_totals3.totalSalesQtyx3, dbo.Items_Status_totals3.totalPurchaseValuex3,"
    MySQL = MySQL & "    dbo.Items_Status_totals3.totalPurchaseQtyx3, dbo.Items_Status_totals3.StoreNamex3, dbo.Items_Status_totals4.StoreNamex4,"
    MySQL = MySQL & "  dbo.Items_Status_totals4.totalPurchaseQtyx4, dbo.Items_Status_totals4.totalPurchaseValuex4, dbo.Items_Status_totals4.totalSalesQtyx4,"
    MySQL = MySQL & "   dbo.Items_Status_totals4.SaleValueValuex4, dbo.Items_Status_totals7.SaleValueValuex7, dbo.Items_Status_totals7.totalSalesQtyx7,"
    MySQL = MySQL & "   dbo.Items_Status_totals7.totalPurchaseValuex7, dbo.Items_Status_totals7.StoreNamex7, dbo.Items_Status_totals5.StoreNamex5,"
    MySQL = MySQL & "  dbo.Items_Status_totals5.totalPurchaseQtyx5, dbo.Items_Status_totals5.totalPurchaseValuex5, dbo.Items_Status_totals5.totalSalesQtyx5,"
    MySQL = MySQL & "  dbo.Items_Status_totals5.SaleValueValuex5, dbo.Items_Status_totals6.StoreNamex6, dbo.Items_Status_totals6.totalPurchaseQtyx6,"
    MySQL = MySQL & "   dbo.Items_Status_totals6.totalPurchaseValuex6, dbo.Items_Status_totals6.totalSalesQtyx6, dbo.Items_Status_totals6.SaleValueValuex6,"
    MySQL = MySQL & "   dbo.Items_Status_totals6.SaleValuex6, dbo.Items_Status_totals7.totalPurchaseQtyx7, dbo.Items_Status_totals1.actualQtyx1, dbo.Items_Status_totals2.actualQtyx2,"
    MySQL = MySQL & "   dbo.Items_Status_totals3.actualQtyx3, dbo.Items_Status_totals4.actualQtyx4, dbo.Items_Status_totals6.actualQtyx6, dbo.Items_Status_totals5.actualQtyx5,"
    MySQL = MySQL & "   dbo.Items_Status_totals7.actualQtyx7 , dbo.QRyOrdersData.OrderArrivalDateMax"
    MySQL = MySQL & "  FROM         dbo.QRyOrdersData LEFT OUTER JOIN"
    MySQL = MySQL & "   dbo.Items_Status_totals6 ON dbo.QRyOrdersData.order_no = dbo.Items_Status_totals6.order_nox6 LEFT OUTER JOIN"
    MySQL = MySQL & "  dbo.Items_Status_totals5 ON dbo.QRyOrdersData.order_no = dbo.Items_Status_totals5.order_nox5 LEFT OUTER JOIN"
    MySQL = MySQL & "   dbo.Items_Status_totals7 ON dbo.QRyOrdersData.order_no = dbo.Items_Status_totals7.order_nox7 LEFT OUTER JOIN"
    MySQL = MySQL & "  dbo.Items_Status_totals4 ON dbo.QRyOrdersData.order_no = dbo.Items_Status_totals4.order_nox4 LEFT OUTER JOIN"
    MySQL = MySQL & "   dbo.Items_Status_totals3 ON dbo.QRyOrdersData.order_no = dbo.Items_Status_totals3.order_nox3 LEFT OUTER JOIN"
    MySQL = MySQL & "  dbo.Items_Status_totals2 ON dbo.QRyOrdersData.order_no = dbo.Items_Status_totals2.order_nox2 LEFT OUTER JOIN"
    MySQL = MySQL & "   dbo.Items_Status_totals1 ON dbo.QRyOrdersData.order_no = dbo.Items_Status_totals1.order_nox1"

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.UserInterface = EnglishInterface Then
        Set xReport = xApp.OpenReport(App.path & "\Reports\inventory\" & "ItemsStatus.rpt.rpt")
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngComment
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.EngCompanyName
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = "Items Status"
    Else

        Set xReport = xApp.OpenReport(App.path & "\Reports\Inventory\" & "ItemsStatus.rpt")
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabComment
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabCompanyName
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = " „Êﬁ› «·«’‰«› «·Õ«·Ì  Õ Ï" & ToDate

    End If

    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, App.path & "\Reports\Inventory\" & "ItemsStatus.rpt"
    Set xApp = Nothing
    Set xReport = Nothing
    Screen.MousePointer = vbDefault
End Sub

Public Function GetVoucherGLNO(Transaction_ID As Double, _
                               Optional ByRef NoteSerial1 As String) As String
    Dim rs  As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
    Dim sql As String
    If Transaction_ID = 0 Then
        Exit Function
    End If
    sql = "SELECT  *  from Transactions  where Transaction_ID=" & Transaction_ID
 
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        GetVoucherGLNO = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Else
        GetVoucherGLNO = ""
    End If
         
    sql = "SELECT  *  from Transactions  where Transaction_ID=" & Transaction_ID
 
    Set Rs1 = New ADODB.Recordset
    Rs1.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        NoteSerial1 = IIf(IsNull(Rs1("NoteSerial1").value), "", Rs1("NoteSerial1").value)
    Else
        NoteSerial1 = ""
    End If
     
End Function

Public Function DeleteTransactiomsVoucher(Transaction_ID As Double)
    If Transaction_ID = 0 Then
        Exit Function
    End If
    StrSqlDel = "delete From Transactions  where Transaction_ID=" & Transaction_ID
    Cn.Execute StrSqlDel, , adExecuteNoRecords
        
    StrSqlDel = "delete From Transaction_Details  where Transaction_ID=" & Transaction_ID
    Cn.Execute StrSqlDel, , adExecuteNoRecords
        
    StrSqlDel = "delete From Notes where Transaction_ID=" & Transaction_ID
    Cn.Execute StrSqlDel, , adExecuteNoRecords
        
    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & Transaction_ID
    Cn.Execute StrSQL, , adExecuteNoRecords
    StrSqlDel = "delete From ItemsDetails where Transaction_ID=" & Transaction_ID
    Cn.Execute StrSqlDel, , adExecuteNoRecords
     
End Function

Public Function getFirstPeriodDateInthisYear2(Optional ByRef FirstPeriodDateInthisYear As Date)
 
    Dim rs As ADODB.Recordset
    Dim sql As String
 
    sql = "SELECT     Min(OpeneingbalancesDate) AS OpeningBalanceDate FROM         dbo.TblyearsData"
 
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        FirstPeriodDateInthisYear = IIf(IsDate(rs("OpeningBalanceDate").value), rs("OpeningBalanceDate").value, Date)
    End If

End Function

Public Function getFirstPeriodDateInthisYear(Optional ByRef FirstPeriodDateInthisYear As Date)
 
    Dim rs As ADODB.Recordset
    Dim sql As String
 
    sql = "SELECT     MAX(OpeneingbalancesDate) AS OpeningBalanceDate FROM         dbo.TblyearsData  where CurrentYear=1 "
 
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        FirstPeriodDateInthisYear = IIf(IsDate(rs("OpeningBalanceDate").value), rs("OpeningBalanceDate").value, "1990-01-01")
    End If

End Function

Public Function GetOverHeadForItems(ItemID As Integer) As Double
 'xxxx
    Dim rs As ADODB.Recordset
    Dim sql As String
 
    sql = "  SELECT     TOP 100 PERCENT SUM(dbo.TblDistriExpensItemDet3.Vlue) AS TOTAL"
sql = sql & "  FROM         dbo.TblDistriExpensItemDet2 LEFT OUTER JOIN"
sql = sql & "   dbo.TblDistriExpensItemDet3 ON dbo.TblDistriExpensItemDet2.ID = dbo.TblDistriExpensItemDet3.IDDet LEFT OUTER JOIN"
sql = sql & "   dbo.ACCOUNTS ON REPLACE(REPLACE(dbo.TblDistriExpensItemDet3.Account_Code, CHAR(10), ''), CHAR(13), '') = dbo.ACCOUNTS.Account_Code"
sql = sql & "  GROUP BY dbo.TblDistriExpensItemDet2.ItemID"
sql = sql & "   Having (dbo.TblDistriExpensItemDet2.ItemID = " & ItemID & ")"
sql = sql & "  ORDER BY dbo.TblDistriExpensItemDet2.ItemID"

 'sql = sql & "  AND  (dbo.Transaction_Details.NProductionOrderNO = '" & WorkOrderNO & "')"
                      
 
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        GetOverHeadForItems = IIf(IsNull(rs("total").value), 0, rs("total").value)
    End If

End Function


Public Function GetProductionTotalIssue(WorkOrderNO As String) As Double
 
    Dim rs As ADODB.Recordset
    Dim sql As String
 
    sql = " SELECT     SUM(dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice) AS total"
sql = sql & "  FROM         dbo.Transactions INNER JOIN"
sql = sql & "    dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
sql = sql & "  WHERE     (dbo.Transactions.Transaction_Type = 27)" ' AND (dbo.Transactions.WorkOrderNO = '" & WorkOrderNO & "') OR"
 sql = sql & "  AND  (dbo.Transaction_Details.NProductionOrderNO = '" & WorkOrderNO & "')"
                      
 
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        GetProductionTotalIssue = IIf(IsNull(rs("total").value), 0, rs("total").value)
    End If

End Function

Public Function getOpeningBalancedate(Optional FromDate As Date, _
                                      Optional ToDate As Date, _
                                      Optional ByRef Retuenedfromdate As Date, _
                                      Optional ByRef Retuenedtodate As Date, _
                                      Optional ByRef YEARS As Integer, _
                                      Optional ByRef openingBalanceDate As Date, _
                                      Optional continous As Boolean = False)

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    ' If FrmAccountingReport.chkContinue.value = vbChecked Then
    ' continous = True
    ' Else
    ' continous = False
    'End If
 
    'sql = "SELECT     dbo.TblyearsData.[no], MIN(dbo.TblAccountIntervals.StartDate) AS OpeningBalanceDate"
    'sql = sql + " FROM         dbo.TblAccountIntervals INNER JOIN"
    'sql = sql + "  dbo.TblyearsData ON dbo.TblAccountIntervals.TblyearsDataid = dbo.TblyearsData.TblyearsDataid"
    'sql = sql + " GROUP BY dbo.TblyearsData.[no]"
    'sql = sql + "  HAVING      (dbo.TblyearsData.[no] = " & YEARS & ")"
    If continous = False Then
        sql = "SELECT     MAX(OpeneingbalancesDate) AS OpeningBalanceDate FROM         dbo.TblyearsData"
    Else
        sql = "SELECT     Min(OpeneingbalancesDate) AS OpeningBalanceDate FROM         dbo.TblyearsData"
    End If

    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        openingBalanceDate = IIf(IsDate(rs("OpeningBalanceDate").value), rs("OpeningBalanceDate").value, FromDate)
    End If

End Function


Public Function getprofitValue(Optional BegineDate As Date, _
                               Optional EndDate As Date, _
                               Optional ByRef ActivityId As Integer = 0, _
                               Optional ByRef branch_id As Integer = 0, Optional ManyBranch As String) As Double

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Dim totalExpenses As Double
    Dim totalRevenue As Double
    Dim debitSum As Double
    Dim creditSum As Double
  
    GoTo ll
    sql = "Select Sum(DEV_Value1)-Sum(DEV_Value2) as  result , Sum(DEV_Value1) as d,Sum(DEV_Value2) as  cre" & CHR(13)
    sql = sql & "   from" & CHR(13)
    sql = sql & "       (" & CHR(13)
    sql = sql & "   SELECT" & CHR(13)
    sql = sql & "   DEV_Value1=Case" & CHR(13)
    sql = sql & "   When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & "   Else 0" & CHR(13)
    sql = sql & "   END," & CHR(13)
    sql = sql & "   DEV_Value2=Case" & CHR(13)
    sql = sql & "   When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & "   Else 0" & CHR(13)
    sql = sql & "   End" & CHR(13)
    
    sql = sql & "       FROM         dbo.Notes RIGHT OUTER JOIN" & CHR(13)
    sql = sql & "   dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN" & CHR(13)
    sql = sql & "   dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id INNER JOIN" & CHR(13)
    sql = sql & "   dbo.tblActivitesType ON dbo.TblBranchesData.ActivityTypeId = dbo.tblActivitesType.id INNER JOIN" & CHR(13)
    sql = sql & "   dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code" & CHR(13)
    sql = sql & " WHERE     (dbo.ACCOUNTS.AccountTab = 2 or dbo.ACCOUNTS.AccountTab = 3) " & CHR(13)
    sql = sql & " and     dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate >=" & SQLDate(BegineDate, True) & CHR(13)
    sql = sql & " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <=" & SQLDate(EndDate, True) & CHR(13)
    sql = sql & " and (dbo.DOUBLE_ENTREY_VOUCHERS.Posted Is Null)"
    If branch_id <> 0 Then
        sql = sql & " and dbo.DOUBLE_ENTREY_VOUCHERS.branch_id=" & branch_id & CHR(13)
    End If

    If ActivityId <> 0 Then
        sql = sql & " and dbo.tblActivitesType.id =" & ActivityId & CHR(13)
    End If
   If ManyBranch <> "" Then
       sql = sql & " and dbo.DOUBLE_ENTREY_VOUCHERS.branch_id in (" & ManyBranch & ")" & CHR(13)
   End If
   
    sql = sql & "       )XTable" & CHR(13)
 
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        totalRevenue = IIf(IsNull(rs("result").value), 0, rs("result").value)
    End If

    rs.Close
ll:
    sql = "Select Sum(DEV_Value1)-Sum(DEV_Value2) as  result , Sum(DEV_Value1) as debitSum,Sum(DEV_Value2) as  creditSum" & CHR(13)
    sql = sql & "   from" & CHR(13)
    sql = sql & "       (" & CHR(13)
    sql = sql & "   SELECT" & CHR(13)
    sql = sql & "   DEV_Value1=Case" & CHR(13)
    sql = sql & "   When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & "   Else 0" & CHR(13)
    sql = sql & "   END," & CHR(13)
    sql = sql & "   DEV_Value2=Case" & CHR(13)
    sql = sql & "   When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & "   Else 0" & CHR(13)
    sql = sql & "   End" & CHR(13)
    
    sql = sql & "       FROM         dbo.Notes RIGHT OUTER JOIN" & CHR(13)
    sql = sql & "   dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN" & CHR(13)
    sql = sql & "   dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id INNER JOIN" & CHR(13)
    sql = sql & "   dbo.tblActivitesType ON dbo.TblBranchesData.ActivityTypeId = dbo.tblActivitesType.id INNER JOIN" & CHR(13)
    sql = sql & "   dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code" & CHR(13)
    sql = sql & " WHERE     (dbo.ACCOUNTS.AccountTab = 2 or dbo.ACCOUNTS.AccountTab = 3) " & CHR(13)
    sql = sql & " and     dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate >=" & SQLDate(BegineDate, True) & CHR(13)
    sql = sql & " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <=" & SQLDate(EndDate, True) & CHR(13)
     sql = sql & " and (dbo.DOUBLE_ENTREY_VOUCHERS.Posted Is Null)"
    If branch_id <> 0 Then
        sql = sql & " and dbo.DOUBLE_ENTREY_VOUCHERS.branch_id=" & branch_id & CHR(13)
    End If
   If ManyBranch <> "" Then
       sql = sql & " and dbo.DOUBLE_ENTREY_VOUCHERS.branch_id in (" & ManyBranch & ")" & CHR(13)
   End If
    If ActivityId <> 0 Then
        sql = sql & " and dbo.tblActivitesType.id =" & ActivityId & CHR(13)
    End If
   
    sql = sql & "       )XTable" & CHR(13)
 
 


 








sql = ""
sql = sql & ";WITH DEV AS (" & CHR(13)
sql = sql & "    SELECT" & CHR(13)
sql = sql & "        d.Value," & CHR(13)
sql = sql & "        d.Credit_Or_Debit," & CHR(13)
sql = sql & "        CASE WHEN d.Credit_Or_Debit = 0 THEN d.Value" & CHR(13)
sql = sql & "             WHEN d.Credit_Or_Debit = 1 THEN -d.Value" & CHR(13)
sql = sql & "        END AS DEV_Value" & CHR(13)
sql = sql & "    FROM DOUBLE_ENTREY_VOUCHERS d" & CHR(13)
sql = sql & "    JOIN ACCOUNTS a ON a.Account_Code = d.Account_Code" & CHR(13)
sql = sql & "    WHERE d.RecordDate >=" & SQLDate(BegineDate, True) & CHR(13)
sql = sql & "      AND d.RecordDate <=" & SQLDate(EndDate, True) & CHR(13)
sql = sql & "      AND a.AccountTypes = 2" & CHR(13)
sql = sql & "      AND a.last_account = 1" & CHR(13)
sql = sql & "      AND (d.Posted IS NULL)" & CHR(13)

' ›·« — «·›—⁄
If branch_id <> 0 Then
    sql = sql & "      AND d.branch_id = " & branch_id & CHR(13)
End If

If ManyBranch <> "" Then
    sql = sql & "      AND d.branch_id IN (" & ManyBranch & ")" & CHR(13)
End If

' ›· — «·‰‘«ÿ
If ActivityId <> 0 Then
    sql = sql & "      AND d.branch_id IN (SELECT branch_id FROM TblBranchesData WHERE ActivityTypeId = " & ActivityId & ")" & CHR(13)
End If

sql = sql & ")" & CHR(13)

sql = sql & "SELECT" & CHR(13)
sql = sql & "   SUM(CASE WHEN Credit_Or_Debit = 0 THEN Value ELSE 0 END) AS debitSum," & CHR(13)
sql = sql & "   SUM(CASE WHEN Credit_Or_Debit = 1 THEN Value ELSE 0 END) AS creditSum," & CHR(13)
sql = sql & "   SUM(DEV_Value) AS result" & CHR(13)
sql = sql & "FROM DEV" & CHR(13)

Set rs = New ADODB.Recordset
rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

If rs.RecordCount > 0 Then
    debitSum = IIf(IsNull(rs("debitSum").value), 0, rs("debitSum").value)
    creditSum = IIf(IsNull(rs("creditSum").value), 0, rs("creditSum").value)
End If

rs.Close

' ================================
'   «·—»Õ «·‰Â«∆Ì
' ================================
getprofitValue = creditSum - debitSum




    'getprofitValue = (totalExpenses)

   
    If creditSum - debitSum > 0 Then
        '—»Õ
    ElseIf creditSum - debitSum < 0 Then
        getprofitValue = getprofitValue * 1
    Else
        getprofitValue = 0
    End If
    
End Function


Public Function getprofitValuesalimsalim(Optional BegineDate As Date, _
                               Optional EndDate As Date, _
                               Optional ByRef ActivityId As Integer = 0, _
                               Optional ByRef branch_id As Integer = 0, Optional ManyBranch As String) As Double

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Dim totalExpenses As Double
    Dim totalRevenue As Double
    Dim debitSum As Double
    Dim creditSum As Double
  
    GoTo ll
    sql = "Select Sum(DEV_Value1)-Sum(DEV_Value2) as  result , Sum(DEV_Value1) as d,Sum(DEV_Value2) as  cre" & CHR(13)
    sql = sql & "   from" & CHR(13)
    sql = sql & "       (" & CHR(13)
    sql = sql & "   SELECT" & CHR(13)
    sql = sql & "   DEV_Value1=Case" & CHR(13)
    sql = sql & "   When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & "   Else 0" & CHR(13)
    sql = sql & "   END," & CHR(13)
    sql = sql & "   DEV_Value2=Case" & CHR(13)
    sql = sql & "   When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & "   Else 0" & CHR(13)
    sql = sql & "   End" & CHR(13)
    
    sql = sql & "       FROM         dbo.Notes INNER JOIN" & CHR(13)
    sql = sql & "   dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN" & CHR(13)
    sql = sql & "   dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id INNER JOIN" & CHR(13)
    sql = sql & "   dbo.tblActivitesType ON dbo.TblBranchesData.ActivityTypeId = dbo.tblActivitesType.id INNER JOIN" & CHR(13)
    sql = sql & "   dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code" & CHR(13)
    sql = sql & " WHERE     (dbo.ACCOUNTS.AccountTab = 2 or dbo.ACCOUNTS.AccountTab = 3) " & CHR(13)
    sql = sql & " and     dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate >=" & SQLDate(BegineDate, True) & CHR(13)
    sql = sql & " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <=" & SQLDate(EndDate, True) & CHR(13)
    sql = sql & " and (dbo.DOUBLE_ENTREY_VOUCHERS.Posted Is Null)"
    If branch_id <> 0 Then
        sql = sql & " and dbo.DOUBLE_ENTREY_VOUCHERS.branch_id=" & branch_id & CHR(13)
    End If

    If ActivityId <> 0 Then
        sql = sql & " and dbo.tblActivitesType.id =" & ActivityId & CHR(13)
    End If
   If ManyBranch <> "" Then
       sql = sql & " and dbo.DOUBLE_ENTREY_VOUCHERS.branch_id in (" & ManyBranch & ")" & CHR(13)
   End If
   
    sql = sql & "       )XTable" & CHR(13)
 
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        totalRevenue = IIf(IsNull(rs("result").value), 0, rs("result").value)
    End If

    rs.Close
ll:
    sql = "Select Sum(DEV_Value1)-Sum(DEV_Value2) as  result , Sum(DEV_Value1) as debitSum,Sum(DEV_Value2) as  creditSum" & CHR(13)
    sql = sql & "   from" & CHR(13)
    sql = sql & "       (" & CHR(13)
    sql = sql & "   SELECT" & CHR(13)
    sql = sql & "   DEV_Value1=Case" & CHR(13)
    sql = sql & "   When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & "   Else 0" & CHR(13)
    sql = sql & "   END," & CHR(13)
    sql = sql & "   DEV_Value2=Case" & CHR(13)
    sql = sql & "   When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & "   Else 0" & CHR(13)
    sql = sql & "   End" & CHR(13)
    
    sql = sql & "       FROM         dbo.Notes INNER JOIN" & CHR(13)
    sql = sql & "   dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN" & CHR(13)
    sql = sql & "   dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id INNER JOIN" & CHR(13)
    sql = sql & "   dbo.tblActivitesType ON dbo.TblBranchesData.ActivityTypeId = dbo.tblActivitesType.id INNER JOIN" & CHR(13)
    sql = sql & "   dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code" & CHR(13)
    sql = sql & " WHERE     (dbo.ACCOUNTS.AccountTab = 2 or dbo.ACCOUNTS.AccountTab = 3) " & CHR(13)
    sql = sql & " and     dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate >=" & SQLDate(BegineDate, True) & CHR(13)
    sql = sql & " and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <=" & SQLDate(EndDate, True) & CHR(13)
     sql = sql & " and (dbo.DOUBLE_ENTREY_VOUCHERS.Posted Is Null)"
    If branch_id <> 0 Then
        sql = sql & " and dbo.DOUBLE_ENTREY_VOUCHERS.branch_id=" & branch_id & CHR(13)
    End If

    If ActivityId <> 0 Then
        sql = sql & " and dbo.tblActivitesType.id =" & ActivityId & CHR(13)
    End If
   
    sql = sql & "       )XTable" & CHR(13)
 
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        'totalExpenses = IIf(IsNull(rs("result").value), 0, rs("result").value)
        debitSum = IIf(IsNull(rs("debitSum").value), 0, rs("debitSum").value)
        creditSum = IIf(IsNull(rs("creditSum").value), 0, rs("creditSum").value)

    End If

    rs.Close

    'getprofitValue = (totalExpenses)
    getprofitValuesalimsalim = creditSum - debitSum

    If creditSum - debitSum > 0 Then
        '—»Õ
    ElseIf creditSum - debitSum < 0 Then
        getprofitValuesalimsalim = getprofitValue * 1
    Else
        getprofitValuesalimsalim = 0
    End If
    
End Function

 Public Function updatallAccountBalances(Optional havedate As Boolean, _
                                        Optional FromDate As Date, _
                                        Optional ToDate As Date, _
                                        Optional Branch As Long = 0, _
                                        Optional ByRef openingbalacedate As Date)

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    sql = "Select * from ACCOUNTS  where last_account=0 "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs.RecordCount
        update_account_balance rs("Account_Code").value, True, FromDate, ToDate, Branch, openingbalacedate
        'update_account_opening_balance rs("Account_Code").value, havedate, fromdate, ToDate, branch
        rs.MoveNext
    Next i

End Function

Public Function updatallAccountOpeningBalances(Optional havedate As Boolean, _
                                               Optional FromDate As Date, _
                                               Optional ToDate As Date, _
                                               Optional Branch As Long = 0, _
                                               Optional ByRef openingbalacedate As Date)

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    sql = "Select * from ACCOUNTS"  'where last_account=1 "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs.RecordCount
        update_account_opening_balance rs("Account_Code").value, True, FromDate, ToDate - 1, Branch, openingbalacedate
        'update_account_opening_balance rs("Account_Code").value, havedate, fromdate, ToDate, branch
        rs.MoveNext
    Next i

End Function

Public Function CheckDelStatus(jopstatusid As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where jopstatusid=" & jopstatusid & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckDelStatus = False
    Else
        CheckDelStatus = True
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Sub FromExcel(ByRef mGrid As Object, _
                     ByRef mtmpGrd As Object, _
                     Frm As Form, _
                     Optional MainFormName As String = "", _
                     Optional ProgressBar As Object = Nothing, _
                     Optional ByVal XlsFileName As String = "", _
                     Optional ByVal MainTableName As String = "")

    ' If Not i Then Exit Sub
    '       Dim cProgress As ClsProgress
    Dim Hide As Integer
    '    Dim mtmpGrd As VSFlexGrid
    If XlsFileName = "" Then
        MsgBox "Õœœ «·„·› «Ê·«", vbCritical
        Exit Sub
        'XlsFileName = GetGridFileName(mGrid, MainFormName)
    End If
    'MsgBox "Before XlsFileName"
    If FileExists(XlsFileName) Then
        ' MsgBox "After  XlsFileName"
        mtmpGrd.FixedCols = 0
        mtmpGrd.FixedRows = 0
        
        '        MsgBox "Before loadgrid"
        mtmpGrd.loadgrid XlsFileName, flexFileExcel
        'MsgBox "Before loadgrid"
        mtmpGrd.backcolor = &HFFFFFF
        mtmpGrd.BackColorAlternate = &HE9E9E9
        mtmpGrd.BackColorBkg = &H8000000C
        mtmpGrd.BackColorFixed = &H8000000F
        mtmpGrd.BackColorFrozen = &HC0FFFF
        mtmpGrd.BackColorSel = &H8000000D
        mtmpGrd.ForeColor = &H80000008
        mtmpGrd.ForeColorFixed = &HFF0000
        mtmpGrd.ForeColorSel = &H8000000E
        mtmpGrd.GridColor = &H8000000F
        mtmpGrd.GridColorFixed = &H80000010
        mtmpGrd.FixedCols = 1
        mtmpGrd.FixedRows = 1
        '·«‰ Loaded ÌŒ ›Ì
        mtmpGrd.Cols = mGrid.Cols + 1
        mtmpGrd.ColKey(mtmpGrd.Cols - 1) = "Loaded"
        mtmpGrd.ColHidden(mtmpGrd.Cols - 1) = True
        mtmpGrd.AutoSize 0, mtmpGrd.Cols - 1
    End If
    mGrid.rows = 1
    mGrid.rows = mtmpGrd.rows

    '********************************
    '    If Not ProgressBar Is Nothing Then
    '        ProgressBar.Min = 1
    '        ProgressBar.Max = IIf(mGrid.Rows > 2, mGrid.Rows - 1, 2)    ' mGrid.Rows - 1
    '        ProgressBar.Visible = True
    '        '********************************
    '    End If
    '        Set cProgress = New ClsProgress
    '       cProgress.ProgressType = Waiting
       
    For i = 1 To mtmpGrd.rows - 1
        '        '********************************
        '        If Not ProgressBar Is Nothing Then
        '            ProgressBar.value = i
        '            DoEvents
        '            ProgressBar.Refresh
        '        End If
        '        cProgress.StartProgress
        '       DoEvents
        '        '********************************
        jj = 0
        For j = 1 To mGrid.Cols - 1
            If j = 18 Then
                j = 18
            End If
            If Not mGrid.ColHidden(j) Then
                jj = jj + 1
                If mGrid.ColKey(j) = "MainGroumName" Then
                    j = j
                End If
                If i = mGrid.rows Then
                    Exit Sub
                End If
                Debug.Print i & " " & mGrid.TextMatrix(i, j)
                If InStr(1, mGrid.ColComboList(j), "#") Then
                    Hide = 0
                    For H = j - 1 To 1 Step -1
                        Hide = Hide + IIf(mGrid.ColHidden(H), 1, 0)
                    Next
                    mGrid.TextMatrix(i, j) = mtmpGrd.TextMatrix(i, j - Hide)
                    'Replace(Trim(mtmpGrd.TextMatrix(i, jj)), "'", "")
                Else
                    mGrid.TextMatrix(i, j) = Replace(Trim(mtmpGrd.TextMatrix(i, jj)), "'", "")
                End If
                If Trim(mGrid.ColEditMask(j)) = "Date" Then
                    GetFieldID mGrid.ColEditMask(j), i, j, mGrid
                End If
                'pValue = Split(G.ColComboList(j), ";")
            Else
                j = j
                If j = 34 Then
                    j = j
                End If
                If Trim(mGrid.ColEditMask(j)) <> "" Then
                    GetFieldID mGrid.ColEditMask(j), i, j, mGrid, MainTableName
                End If
                If Trim(mGrid.ColComboList(j)) <> "" Then
                    GetIDCombo Trim(mGrid.ColComboList(j)), i, j, mGrid
                End If
            End If
            If Trim(Replace(Trim(mtmpGrd.TextMatrix(i, 1)), "'", "")) = "" Then
                mGrid.rows = i + 1:  Exit Sub
            End If
        Next
        ' DisplayOrderTotals
NextRow:
    Next
    '    '********************************
    '    If Not ProgressBar Is Nothing Then
    '        ProgressBar.Visible = False
    '    End If
    '           DoEvents
    '    cProgress.FinishProgress
    '    cProgress.StopProgess
    '    Set cProgress = Nothing
    '   MsgBox " „ «·«œ—«Ã"
    '********************************
End Sub

Private Sub GetFieldID(ByVal mTableColName As String, _
                       ByVal mRow As Long, _
                       ByVal mCol As Long, _
                       ByVal mGrid As Object, _
                       Optional ByVal MainTableName As String = "")
    Dim mTableName   As String
    Dim mFieldIDName As String
    Dim mFieldName   As String
    Dim xx           As Variant
    Dim mValue       As String
    Dim rsDummy      As New ADODB.Recordset
    Dim rsDummy2     As New ADODB.Recordset
    If mCol = 67 Then
        mCol = 67
    End If
    If mGrid.ColKey(mCol) = "NationlID" Then
        mCol = mCol
    End If
    Dim mValue2 As String
    If mGrid.ColKey(mCol) = "DeanID" Then
        mCol = mCol
    End If
    If mGrid.ColKey(mCol) = "DOBH" Then
        mCol = mCol
    End If
    If mTableColName = "Date" Then
        If CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 1 Then
            'If Trim(mGrid.TextMatrix(mRow, mCol - 1)) <> "" Then
            mGrid.TextMatrix(mRow, mCol) = Trim(mGrid.TextMatrix(mRow, mCol - 1))
            mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(mGrid.TextMatrix(mRow, mCol))
            'Else
            'End If
        ElseIf CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 2 Then
            If Trim(mGrid.TextMatrix(mRow, mCol - 1)) = "" Then
                mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(Trim(mGrid.TextMatrix(mRow, mCol)))
            Else
                mGrid.TextMatrix(mRow, mCol) = ToHijriDate(Trim(mGrid.TextMatrix(mRow, mCol - 1)))
            End If
        ElseIf CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 3 Then
            If mGrid.TextMatrix(mRow, mCol) <> "" Then
                mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(Trim(mGrid.TextMatrix(mRow, mCol)))
            End If
            'mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(mGrid.TextMatrix(mRow, mCol))
        Else
        
        End If
        Exit Sub
    End If
    xx = Split(mTableColName, ",")
    mTableName = xx(0)
    mFieldIDName = xx(1)
    mFieldName = xx(2)
    
    If mRow = 50 Then
        mRow = mRow
    End If
    If mRow = mGrid.rows Then
        Exit Sub
    End If
    mValue = Trim(mGrid.TextMatrix(mRow, mCol - 1))
    Dim strValue As String
    strValue = ""
    Dim mValue3 As String

    mValue3 = mValue
    If (right(mValue, 1)) = "Â" Then
        strValue = "…"
    ElseIf (right(mValue, 1)) = "…" Then
        strValue = "Â"
    
    End If
    If strValue <> "" Then
        mValue3 = Replace(mValue3, right(mValue3, 1), strValue)
    End If
    Dim mEngLett As String
    mEngLett = "e"
    Dim s As String
    mValue2 = mValue
    Select Case mTableName
        Case "jopstatus"
            If UCase(mValue) = "ACTIVE" Then
                mValue2 = "⁄·Ï ﬁÊ… «·⁄„·"
            
            End If
        Case "dean"
            If UCase(mValue) = "ISLAM" Then
                mValue2 = "„”·„"
            ElseIf UCase(mValue) = "CHRISTIAN" Then
                mValue2 = "„”ÌÕÏ"
            End If
        Case "Nationality"
            If UCase(mValue) = "JORDAN" Then
                mValue2 = "«—œ‰"
            ElseIf UCase(mValue) = "INDIA" Then
                mValue2 = "Â‰œ"
            ElseIf Trim(UCase(mValue)) = "" Then
                mValue2 = "”⁄ÊœÌ"
            ElseIf UCase(mValue) = "EGYPT" Then
                mValue2 = "„’—"
            ElseIf UCase(mValue) = "PAKISTAN" Then
                mValue2 = "»«ﬂ” «‰"
            ElseIf UCase(mValue) = "BANGLADESH" Then
                mValue2 = "»‰Ã·«œÌ‘"
            ElseIf UCase(mValue) = "SUDAN" Then
                mValue2 = "”Êœ«‰"
            ElseIf UCase(mValue) = "ETHIOPIA" Then
                mValue2 = "«ÀÌÊ»Ì«"
            
            ElseIf UCase(mValue) = "CAMEROON" Then
                mValue2 = "ﬂ«„Ì—Ê‰"
            ElseIf UCase(mValue) = "PALESTINE" Then
                mValue2 = "›·”ÿÌ‰"
            ElseIf UCase(mValue) = "SYRIA" Then
                mValue2 = "”Ê—Ì«"
            ElseIf UCase(mValue) = "JORDANIAN" Then
                mValue2 = "«—œ‰"
            ElseIf UCase(mValue) = "AMERICA" Then
                mValue2 = "«„—Ìﬂ«"
            ElseIf UCase(mValue) = "EGYPTIAN" Then
                mValue2 = "„’—"
            ElseIf UCase(mValue) = "KENYA" Then
                mValue2 = "ﬂÌ‰Ì«"
            ElseIf UCase(mValue) = "LEBANON" Then
                mValue2 = "·»‰«‰"
            ElseIf UCase(mValue) = "SIRLANKIAN" Then
                mValue2 = "”Ì—·«‰ﬂ"
            ElseIf UCase(mValue) = "YEMEN" Then
                mValue2 = "Ì„‰"
            ElseIf UCase(mValue) = "TUNIS" Then
                mValue2 = " Ê‰”"
            ElseIf UCase(mValue) = "MALAYSIA" Then
                mValue2 = "„«·Ì“Ì«"
            Else
                mValue2 = mValue
            
            End If
            If mValue = "" Then mValue2 = "”⁄ÊœÌ"
        Case Else
    End Select
    If mValue = "" Then
        Exit Sub
    End If
    mEngLett = "e"
    If UCase(mTableName) = "ACCOUNTS" Then
        mEngLett = "Eng"
    End If
    If UCase(mTableName) = "TBLCOUNTRIESGOVERNMENTS" Then
        mEngLett = ""
    End If
    
    s = "Select " & mFieldName & " ," & mFieldIDName & " ," & Trim(mFieldName) & mEngLett
    If UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Or UCase(mTableName) = "FIXEDASSETSGROUP" Then
        s = s & " ,ParentID,FullCode,GroupCode,Code,LastGroup "
    
    End If
    
    s = s & " from  " & mTableName
    s = s & " Where (" & mFieldName & " = '" & Trim(mValue2) & "' Or " & Trim(mFieldName) & mEngLett & "    = '" & Trim(mValue) & "')"
    s = s & " or (" & mFieldName & " = '" & Trim(mValue3) & "' Or " & Trim(mFieldName) & mEngLett & "   = '" & Trim(mValue3) & "')"
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    If rsDummy.EOF Then
        s = s & " Or ( " & mFieldName & " Like '%" & Trim(mValue2) & "%' Or " & Trim(mFieldName) & mEngLett & "    Like '%" & Trim(mValue) & "%')"
    End If
    If rsDummy.EOF And UCase(mTableName) = "ACCOUNTS" Then
        MsgBox "Â–« «·Õ”«» €Ì— „ÊÃÊœ ›Ï «·œ·Ì· " & mValue
        Exit Sub
    End If
    rsDummy.Close
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    If UCase(mTableName) = "GROUPS" And rsDummy.EOF Then
        rsDummy.Close
        s = "Select " & mFieldName & " ," & mFieldIDName & " ," & Trim(mFieldName) & "e   "
        If UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Or UCase(mTableName) = "FIXEDASSETSGROUP" Then
            s = s & " ,ParentID,FullCode,GroupCode,Code,LastGroup "
        End If
        Dim mValue4 As String
        mValue4 = Trim(mGrid.TextMatrix(mRow, mCol - 2))
        
        s = s & " from  " & mTableName
        s = s & " Where " & mFieldName & " Like '%" & Trim(mValue2) & "%' Or " & Trim(mFieldName) & "e Like '%" & Trim(mValue) & "%'"
        If mValue4 <> "" Then
            s = s & " Or Fullcode   Like '%" & Trim(mValue4) & "%' Or Code Like '%" & Trim(mValue4) & "%'"
        End If
        rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
        If rsDummy.EOF Then
            mValue4 = mValue4
        End If
    End If
    
    If Not rsDummy.EOF Then
        If UCase(mTableName) = "ACCOUNTS" Then
            mGrid.TextMatrix(mRow, mCol) = Trim(rsDummy.Fields.Item(Trim(mFieldIDName)) & "")
        Else
            mGrid.TextMatrix(mRow, mCol) = val(rsDummy.Fields.Item(Trim(mFieldIDName)) & "")
        End If
        If mGrid.ColKey(mCol) = "ParentID" Then
            mGrid.TextMatrix(mRow, mGrid.ColIndex("Code")) = Trim(mGrid.TextMatrix(mRow, mGrid.ColIndex("FullCode")))
            Dim mmm As String
            mmm = SearchInGrid(mGrid, mValue, "GroupName")
            If mmm <> "" Then
                'mGrid.TextMatrix(mRow, mGrid.ColIndex("GroupCode")) = GetNewGroupCode(Val(mGrid.TextMatrix(CLng(mmm), mGrid.ColIndex("NewId"))))
            End If
            mGrid.TextMatrix(mRow, mGrid.ColIndex("LastGroup")) = 0
        End If

    Else
        '         tRs!GroupCode = GetNewGroupCode(val(tGrd.TextMatrix(i, tGrd.ColIndex("ParentID"))), mTableName)
        '                If UCase(mTableName) <> "GROUPSCUSTOMERS" Then
        '                    tRs!GroupID = val(mMaxId)
       
        rsDummy.AddNew
        rsDummy(Trim(mFieldName)) = mValue
        rsDummy(Trim(mFieldName) & mEngLett) = mValue
        If mGrid.ColKey(mCol) = "ParentID" Then
            'rsDummy("ParentID") = mValue
            Dim mm As String
            mm = SearchInGrid(mGrid, mValue, "GroupName")
            If mm <> "" Then
                rsDummy("ParentID") = val(mGrid.TextMatrix(CLng(mm), mCol))
                rsDummy("FullCode") = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
                rsDummy("Code") = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
            Else
                xx = Split(Trim(mGrid.TextMatrix(mRow, mGrid.ColIndex("FullCode"))), "-")
                rsDummy("ParentID") = 1
                rsDummy("FullCode") = xx(0)
                rsDummy("Code") = xx(0)
            End If
            rsDummy("GroupCode") = GetNewGroupCode(val(rsDummy("ParentID") & ""), mTableName)
            
            rsDummy("LastGroup") = 0
            If mm <> "" Then
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("Code")) = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("GroupCode3")) = rsDummy("GroupCode") & ""
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("LastGroup")) = 0
            End If
        End If
        s = "Select Max(" & mFieldIDName & ")  as MaxID  from  " & mTableName
        
        rsDummy2.Open s, Cn, adOpenKeyset, adLockOptimistic
        Dim mMaxId As Long
        If Not rsDummy2.EOF Then
            mMaxId = val(rsDummy2!MaxID & "") + 1
        Else
            mMaxId = 1
        End If
        If UCase(mTableName) <> "GROUPSCUSTOMERS" Then
            rsDummy(Trim(mFieldIDName)) = mMaxId
        End If
        rsDummy(Trim(mFieldName)) = mValue
        If UCase(mTableName) = "GROUPS" Then
            rsDummy!GroupCode = GetNewGroupCode(1, mTableName)
            rsDummy!LastGroup = 0
            rsDummy!ParentID = 1
       
        End If
        rsDummy.update
        ' mGrid.TextMatrix(mRow, mGrid.ColIndex("NewId")) = mMaxId
        mGrid.TextMatrix(mRow, mCol) = rsDummy(Trim(mFieldIDName) & "")
    End If

End Sub

Private Function SearchInGrid(ByVal mGrd As Object, ByVal mTxt As String, ByVal mFldName As String) As String
Dim i As Long
For i = 1 To mGrd.rows - 1
    If Trim(mGrd.TextMatrix(i, mGrd.ColIndex(mFldName))) = mTxt Then
        SearchInGrid = i
        Exit Function
    End If
Next
SearchInGrid = ""
End Function
Function FileExists(FileName) As Boolean
    On Error GoTo CheckError        ' Turn on error trapping so error handler                            ' responds if any error is detected.
    FileExists = (Dir(FileName) <> "")
    Exit Function

CheckError:        ' Branch here if error occurs.    ' Define constants to represent Visual Basic error code.
    FileExists = False
    Resume Next
End Function

Private Sub GetIDCombo(ByVal mTableColID As String, _
                       ByVal mRow As Long, _
                       ByVal mCol As Long, _
                       ByVal mGrid As Object)
    Dim mTxt As String
    mTxt = Trim(mGrid.TextMatrix(mRow, mCol - 1))
    Select Case mTableColID
        Case "sexID"
            If mTxt = "Male" Or mTxt = "–ﬂ—" Then
                mTxt = 1
            Else
                mTxt = 2
            End If
        Case "MaritalStatusID"
            '    DcbMatrial.AddItem "√⁄“»"
            '      DcbMatrial.AddItem "„ “ÊÃ"
            If mTxt = "√⁄“»" Or mTxt = "Single" Then
                mTxt = 0
            ElseIf mTxt = "„ “ÊÃ" Or UCase(mTxt) = "MARRIED" Then
                mTxt = 1
            ElseIf mTxt = "„ÿ·ﬁ/„ÿ·›…" Or UCase(mTxt) = "DIVORCED" Then
                mTxt = 2
            ElseIf mTxt = "«—„·/√—„·…" Or UCase(mTxt) = "WIDOWED" Then
                mTxt = 3
        
            End If
        Case "Emp_Name1.Emp_Name2.Emp_Name3.Emp_Name4"
            mTxt = mGrid.TextMatrix(mRow, mCol - 4) + " " + mGrid.TextMatrix(mRow, mCol - 3) + " " + mGrid.TextMatrix(mRow, mCol - 2) + " " + mGrid.TextMatrix(mRow, mCol - 1)
        Case ""
    End Select
    mGrid.TextMatrix(mRow, mCol) = mTxt
End Sub

Public Function CheckDateIsHij(ByVal mDate As String) As Integer
    If Not IsDate(mDate) Then
        CheckDateIsHij = 3
        Exit Function
    End If
    If Trim(mDate) = "" Then
        CheckDateIsHij = 3
        Exit Function
    End If
    
    If year(mDate) < 1800 Then
        CheckDateIsHij = 1
    Else
        CheckDateIsHij = 2
    End If
End Function




Private Function GetNewGroupCode(LngParentGroupID As Long, mTableName As String) As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StrParentCode  As String
    Dim StrNewGroupCode As String
    Dim StrLastGroupCode As String
    Dim IntTemp As String

    On Error GoTo ErrTrap
    StrSQL = "Select GroupCode From Groups Where GroupID=" & LngParentGroupID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF Or rs.BOF) Then
        StrParentCode = IIf(IsNull(rs("GroupCode").value), "", rs("GroupCode").value)
    End If

    rs.Close
    Set rs = New ADODB.Recordset
    StrSQL = "Select * From Groups Where ParentID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        StrNewGroupCode = StrParentCode & "1"
    Else
        rs.MoveLast
        StrLastGroupCode = IIf(IsNull(rs("GroupCode").value), "", rs("GroupCode").value)
        IntTemp = val(mId(StrLastGroupCode, Len(StrParentCode) + 1))
        StrNewGroupCode = StrParentCode & CStr(IntTemp + 1)
    End If

    rs.Close
    Set rs = Nothing
    GetNewGroupCode = StrNewGroupCode
    Exit Function
ErrTrap:
End Function

Public Sub SENDSMS(phone, msgb)
    Dim xmlhttp As Object
    Dim URL As String
    Dim authHeader As String
    Dim requestBody As String

    ' Initialize the XMLHTTP object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    URL = "https://api.oursms.com/msgs/sms"
    authHeader = "Bearer UJ2dRMbVpQpm6xy0sVVY"
    requestBody = "{" & _
                  """src"": ""SRE-DEV""," & _
                Replace("""dests"": [""phone""],", "phone", phone) & _
                Replace("""body"": ""MSG"",", "MSG", msgb) & _
                  """priority"": 0," & _
                  """delay"": 0," & _
                  """validity"": 0," & _
                  """maxParts"": 0," & _
                  """dlr"": 0," & _
                  """prevDups"": 0," & _
                  """msgClass"": ""promotional""}"
 
    xmlhttp.Open "POST", URL, False
    
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "Authorization", authHeader

    xmlhttp.send requestBody

 
    If xmlhttp.Status = 200 Then
      '  MsgBox "SMS sent successfully: " & xmlhttp.responseText
    Else
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.responseText
    End If
    
 
    Set xmlhttp = Nothing
End Sub
Function CheckPhoneNumber(ByVal phoneNumber As String, Optional ByVal mLineNo As Long = 0) As String
    If left(phoneNumber, 1) = "0" Then
        CheckPhoneNumber = "966" & mId(phoneNumber, 2)
    ElseIf left(phoneNumber, 3) = "966" Then
        CheckPhoneNumber = phoneNumber
    ElseIf left(phoneNumber, 1) <> "0" Then
        CheckPhoneNumber = "966" & phoneNumber
    
    Else
        
'         If mLineNo <> 0 Then
'            MsgBox "—ﬁ„ «· ·Ì›Ê‰ Œ«ÿ∆ Ì—ÃÏ «·„—«Ã⁄…" & " ”ÿ— —ﬁ„ " & mLineNo & " ··—ﬁ„ «·Œ«ÿ∆ :" & phoneNumber
'        Else
'            MsgBox "—ﬁ„ «· ·Ì›Ê‰ Œ«ÿ∆ Ì—ÃÏ «·„—«Ã⁄…"
'        End If
        
        CheckPhoneNumber = "Invalid phone number"
    End If
    If Len(CheckPhoneNumber) <> 12 Then
        If mLineNo <> 0 Then
            MsgBox "—ﬁ„ «· ·Ì›Ê‰ Œ«ÿ∆ Ì—ÃÏ «·„—«Ã⁄…" & " ”ÿ— —ﬁ„ " & mLineNo & " ··—ﬁ„ «·Œ«ÿ∆ :" & phoneNumber
        Else
            MsgBox "—ﬁ„ «· ·Ì›Ê‰ Œ«ÿ∆ Ì—ÃÏ «·„—«Ã⁄…"
        End If
        
    End If
End Function
