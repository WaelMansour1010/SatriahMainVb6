Attribute VB_Name = "registry"
Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

Public State               As String
Public Alarm_start         As String
Public Alarm_end           As String
Public hide_logo           As Boolean
Public run_count           As Integer
Public key_for_me          As String
Public my_branch           As String
Public Current_branch      As Integer
Public Current_branchSql   As String
Public CurrentBranchName   As String
Public CurrentBranchNameE  As String
Public CurrentActivityName As String
Public Declare Function GetKeyboardLayoutName _
               Lib "user32" _
               Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function LoadKeyboardLayout _
               Lib "user32" _
               Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, _
                                            ByVal Flags As Long) As Long
Const KLF_ACTIVATE = &H1
' some languages code
Public Const LANG_ENGLISH As String = "00000409"
Public Const LANG_FRENCH  As String = "0000040C"
Public Const LANG_ARABIC  As String = "00000401"
Public Const LANG_GREEK   As String = "00000408"
Public Const LANG_ITALIAN As String = "00000400"
Public Const LANG_GERMAN  As String = "00000407"
Public user_name_id       As Integer
Public pass_word          As String
Public branch_id          As Integer
Public Activity_id        As Integer
Public language_id        As Integer
Public save_password      As Boolean
Public account_level      As Integer
Public WhereViewString    As String

 '
'
Function GetNumberAfterMonth(inputNumber As String, Optional ByVal YearDigit As Long = 4, Optional ByVal mBranchDigit As Integer = 1, Optional ByVal monthLength2 As Integer = 1) As Long
Dim monthStartPosition As Integer
    Dim monthLength As Integer
    Dim numberAfterMonth As Long
    
    '  ÕœÌœ „Êﬁ⁄ »œ«Ì… «·‘Â— ÊÿÊ·Â
    
   If YearDigit = 4 Then
        If monthLength2 = 1 Then
            monthStartPosition = 6 + mBranchDigit ' «·‘Â— Ì»œ√ „‰ «·Œ«‰… «·Œ«„”…
        Else
            monthStartPosition = 5 + mBranchDigit ' «·‘Â— Ì»œ√ „‰ «·Œ«‰… «·Œ«„”…
        End If
    Else
        If monthLength2 = 1 Then
            monthStartPosition = 4 + mBranchDigit ' «·‘Â— Ì»œ√ „‰ «·Œ«‰… «·Œ«„”…
        Else
            monthStartPosition = 3 + mBranchDigit ' «·‘Â— Ì»œ√ „‰ «·Œ«‰… «·Œ«„”…
        End If
   End If
   ' If monthLength = 1 Then monthStartPosition = monthStartPosition + 1:
    
    monthLength = 2 ' «·‘Â— ÂÊ —ﬁ„ –Ê Œ«‰ Ì‰
    
    
    ' «· Õﬁﬁ „‰ ÿÊ· «·”·”·… «·‰’Ì…
    If Len(inputNumber) >= monthStartPosition + monthLength Then
        ' «” Œ—«Ã «·—ﬁ„ «·–Ì Ì·Ì «·‘Â—
        numberAfterMonth = mId(inputNumber, monthStartPosition + monthLength)
        If numberAfterMonth = 0 Then
            numberAfterMonth = mId(inputNumber, monthStartPosition + monthLength - 1)
        End If
    Else
        ' ≈–« ﬂ«‰ ÿÊ· «·”·”·… «·‰’Ì… €Ì— ﬂ«›Ú° ≈—Ã«⁄ ”·”·… ›«—€…
        numberAfterMonth = 0
    End If
    
    GetNumberAfterMonth = numberAfterMonth
End Function


Public Function CalculateTimes(FromTime As String, _
   ToTime As String) As String
    Dim m_PreDate     As Date
    Dim IntMintsCount As Double
    Dim IntHoursCount As Double
    Dim IntMin        As Integer
    Dim StrSing       As String
    Dim StrTemp       As String
    ' Dim tempdate As String
    Dim date1         As Variant
    Dim date2         As Variant

    '  date1 = Format(Now, "yyyy/mm/dd hh:mm:ss")
    '    date2 = Format(Label1, "yyyy/mm/dd hh:mm:ss")
    
    If CDate(ToTime) < CDate(FromTime) Then
 
        date1 = Format("2012/5/15 " & Format(FromTime, "hh:mm:ss"), "yyyy/mm/dd hh:mm:ss")
        date2 = Format("2012/5/16 " & Format(ToTime, "hh:mm:ss"), "yyyy/mm/dd hh:mm:ss")
 
        FromTime = date1
        ToTime = date2
    End If
 
    m_DateCome = GetPresentTime
    m_DateOut = GetDepTime
    IntMintsCount = (DateDiff("n", FromTime, ToTime))
    '   If IntMintsCount < 0 Then
    '   IntMintsCount = Abs(IntMintsCount) + 720
    '   End If
    
    IntHoursCount = IntMintsCount \ 60
    IntMin = IntMintsCount Mod 60

    If IntHoursCount < 0 Or IntMin < 0 Then
        'Stop
        CalculateTimes = "???"
        Exit Function
      
    Else
        StrSing = ""
        StrTemp = StrSing & Format(IntHoursCount, "00") & ":" & Format(IntMin, "00")
       
    End If

    CalculateTimes = StrTemp

End Function

Public Function get_code(FIELD_no As Integer, _
                         my_branch As Integer, _
                         prifix As String, _
                         table_name As String) As String
    'On Error Resume Next

    Dim tempcode    As String
    Dim code        As String
    Dim coding_auto As Boolean
    Dim sql         As String
    Dim Rs3         As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset

    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset

    'Dim Sql As String
    Dim i As Integer

    sql = "select * from coding where FIELD_no=" & FIELD_no & " and  branch_no=" & my_branch & " and prifix='" & prifix & "'"
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    If Rs3.RecordCount > 0 Then

        coding_auto = Rs3("Auto").value

        If coding_auto = False Then
            get_code = "Error1"
            Exit Function
        End If

        If coding_auto = True Then
            no_of_digit = Rs3("no_of_digit").value
            Zeros = Rs3("zeros").value
                
            prifix = prifix_combo.text
                              
        End If
    End If

    sql = "select max(code)  as last_code from " & table_name & "  where branch_no=" & my_branch & " and prifix='" & prifix & "'"
    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If IsNull(Rs4("last_code").value) Then

        tempcode = prifix & "1"

        If Len(tempcode) < no_of_digit Then
            
            diffrent = no_of_digit - Len(tempcode)
            tempcode = prifix

            If Zeros = True Then

                For i = 1 To diffrent
                    tempcode = tempcode & "0"
                    code = code & "0"
                                   
                Next i

            End If
                
            fullcode = tempcode & "1"
            code = code & "1"
            
        Else

            If Len(tempcode) > no_of_digit Then
                MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
                fullcode = "error"
                Exit Function
              
            Else
                fullcode = tempcode
                code = 1
            End If
        End If

    Else
        tempcode = prifix & val(Adodccode2.Recordset.Fields!last_code) + 1

        If Len(tempcode) < no_of_digit Then
            
            diffrent = no_of_digit - Len(tempcode)
            tempcode = prifix

            If Zeros = True Then

                For i = 1 To diffrent
                    tempcode = tempcode & "0"
                    code = code & "0"
                Next i
                        
            End If

            fullcode = tempcode & val(Rs4("last_code").value) + 1
            code = code & val(Rs4("last_code").value + 1)
        Else

            If Len(tempcode) > no_of_digit Then
                MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ… ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â«"
                fullcode = "error"
                Exit Function
              
            Else
                fullcode = tempcode
                code = val(Rs4("last_code").value) + 1
            End If
        End If

    End If

    get_code = code
    '  If Adodc1.Recordset.RecordCount > 0 Then
    '   Adodc1.Recordset.Fields!fullcode = fullcode
    '   End If
End Function

Public Function get_currency_rate(ID As Integer) As Single
    'Transaction_Type=19 «–‰ ’—›
    'Transaction_Type=20 «–‰ «÷«›…
    On Error Resume Next
    Dim departement_name As String
    departement_name = 1
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select * from currency where id=" & ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_currency_rate = 1
        Exit Function
    End If
    get_currency_rate = IIf(IsNull(Rs3("rate").value), 1, Rs3("rate").value)
  
    Rs3.Close

End Function

Public Function check_bill_voucher(bill_id As Double, _
                                   Transaction_Type As Integer) As Double
    'Transaction_Type=19 «–‰ ’—›
    'Transaction_Type=20 «–‰ «÷«›…
    On Error Resume Next
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select * from Transactions where Transaction_Type=" & Transaction_Type & "and Nots='" & bill_id & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        check_bill_voucher = 0
        Exit Function
    End If
    check_bill_voucher = IIf(IsNull(Rs3("Transaction_ID").value), 0, Rs3("Transaction_ID").value)
    Rs3.Close

End Function

Public Function Update_opening_balance_screen_accounts()
 
    On Error Resume Next
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select * from accounts"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Exit Function
    Else
        For i = 1 To Rs3.RecordCount
            update_account_opening_balance Rs3("Account_Code").value
            Rs3.MoveNext
        Next i
    End If

End Function

Public Function get_transaction_id(Transaction_serial As String, _
                                   Transaction_Type As Integer) As Integer
    On Error Resume Next
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select * from Transactions where Transaction_Type=" & Transaction_Type & "and Transaction_Serial='" & Transaction_serial & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_transaction_id = 0
        Exit Function
    End If
    get_transaction_id = IIf(IsNull(Rs3("Transaction_ID").value), 0, Rs3("Transaction_ID").value)
    Rs3.Close

End Function

Public Function get_transaction_idByNoteSerial1(NoteSerial1 As String, _
                                                Transaction_Type As Integer) As Long
    On Error Resume Next
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select * from Transactions where Transaction_Type=" & Transaction_Type & "and NoteSerial1='" & NoteSerial1 & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_transaction_idByNoteSerial1 = 0
        Exit Function
    End If
    get_transaction_idByNoteSerial1 = IIf(IsNull(Rs3("Transaction_ID").value), 0, Rs3("Transaction_ID").value)
    Rs3.Close

End Function

'·«ÌÃ«œ  ﬂ·›… „»Ì⁄«  ’‰›  ›Ì ⁄„·Ì… „Õœœ…
Public Function Get_item_Cost_price(Transaction_ID As Integer, _
                                    Item_ID As Integer) As Integer
    'Transaction_Type=19 «–‰ ’—›
    'Transaction_Type=20 «–‰ «÷«›…
    On Error Resume Next
    Dim departement_name As String
    departement_name = 1
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select * from Transaction_Details where Transaction_ID=" & Transaction_ID & "and Item_ID=" & Item_ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Get_item_Cost_price = 0
        Exit Function
    End If
    Get_item_Cost_price = IIf(IsNull(Rs3("Price").value), 0, Rs3("Price").value)
    Rs3.Close

End Function
Public Function save_employee_prohectt_EndDate(EndDate As Date, _
                                               toid As Integer, _
                                               Emp_id As Integer)
                                              
    '  Dim sql As String
    '  sgl = "update  dbo.opr_employee_details  set enddate=" & EndDate
    '  sgl = sgl + " ,toid='" & toid
    '  sgl = sgl + "' ,[interval]='" & [interval]
    '  sgl = sgl + "' Where  (end_date IS NULL) and  Emp_ID = " & val(Emp_id) & " and toid<>" & toid
    '  Cn.Execute sgl, , adExecuteNoRecords
End Function
Public Function save_employee_current_status(ByVal project_id As Integer, _
                                             term_fullcode As String, _
                                             opr_fullcode As String, _
                                             Emp_id As Integer)
    Dim sql As String
    sgl = "update  TblEmployee  set project_id=" & IIf(Not IsNumeric(CStr(project_id)), 0, project_id)
    sgl = sgl + " ,term_fullcode='" & IIf(term_fullcode = "", "", term_fullcode)
    sgl = sgl + "' ,opr_fullcode='" & IIf(opr_fullcode = "", "", opr_fullcode)
    sgl = sgl + "' Where Emp_ID = " & val(Emp_id)
    Cn.Execute sgl, , adExecuteNoRecords
End Function

Public Function setfoxy() As Double
    Dim new_id1 As Double
    new_id1 = CStr(new_id("foxy", "id", "", True))
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs("id").value = new_id1
    rs.update
    setfoxy = new_id1
End Function
Public Function MyTime() As Double
    MyTime = Format(Now, "ddMMyyHHnnss")
    'MyTime = Format(Now, "ddMMyyyyHHnnss")
    '    MyTime = Format(Now, "ddMMyyyyHHnnss") & "." & right(Format(Timer, "#0.00"), 2)
    
End Function
Public Function get_opening_balance_voucher_id() As Double
    Dim newSeril As Double
    On Error Resume Next
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql      As String

        sql = "select max(opening_balance_voucher_id) As id from DOUBLE_ENTREY_VOUCHERS1 Where opening_balance_voucher_id > 0"
 
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
        If Rs3.RecordCount > 0 Then
            get_opening_balance_voucher_id = IIf(IsNull(Rs3("id").value), 0, Rs3("id").value) + 1
    
        Else
            get_opening_balance_voucher_id = 1
        End If
    Dim LngDevID As Long
    'LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)

    '  Cn.Execute "insert into  DOUBLE_ENTREY_VOUCHERS1 (opening_balance_voucher_id,DEV_ID_Line_No,Double_Entry_Vouchers_ID)  values (" & get_opening_balance_voucher_id & ",0," & LngDevID & ")"
  '  get_opening_balance_voucher_id = MyTime
End Function

Public Function get_balance(Account_code As String, _
                            Optional openingtype As Integer, _
                            Optional havedate As Boolean, _
                            Optional FromDate As Date, _
                            Optional ToDate As Date, _
                            Optional branch_id As Long) As Double
    Dim total_depit  As Single
    Dim total_credit As Single

    Dim total        As Single

    total_credit = 0: total_depit = 0: total = 0
 
    On Error Resume Next
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select sum(value) As total_depit from DOUBLE_ENTREY_VOUCHERS1 where Credit_Or_Debit=0 and  Account_Code LIKE '" & Account_code & "%'"

    If havedate = True Then
   
        sql = sql + " and     RecordDate >=" & SQLDate(FromDate, True) & ""
   
        sql = sql + " and RecordDate <=" & SQLDate(ToDate, True) & ""
  
    End If

    If branch_id <> 0 Then
        sql = sql + " and     branch_id  =" & branch_id
    End If

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        total_depit = 0
    Else
        total_depit = IIf(IsNull(Rs3("total_depit").value), 0, Rs3("total_depit").value)
    End If

    Rs3.Close
 
    sql = "select sum(value) As total_credit from DOUBLE_ENTREY_VOUCHERS1 where Credit_Or_Debit=1 and  Account_Code like'" & Account_code & "%'"
 
    If havedate = True Then
   
        sql = sql + " and     RecordDate >=" & SQLDate(FromDate, True) & ""
   
        sql = sql + " and RecordDate <=" & SQLDate(ToDate, True) & ""
  
    End If

    If branch_id <> 0 Then
        sql = sql + " and     branch_id  =" & branch_id
    End If

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        total_credit = 0
    Else
        total_credit = IIf(IsNull(Rs3("total_credit").value), 0, Rs3("total_credit").value)
    End If

    Rs3.Close
 
    get_balance = total_depit - total_credit

    If get_balance > 0 Then
        openingtype = 0
    Else
        openingtype = 1
    End If

    '1 œ∆«∆‰
    '2„œÌ‰
End Function

Public Function getActualpayedToContract(ContNo As Integer, _
                                         Optional ByRef RentValuePayed As Double, _
                                         Optional ByRef WaterPayed As Double, _
                                         Optional ByRef TelandNetPayed As Double, _
                                         Optional ByRef ElectricPayed As Double)
    Dim sql As String
    Dim i   As Integer
    Dim rs  As New ADODB.Recordset

    sql = "SELECT    sum( dbo.ContracttBillInstallmentsDone.RentValuePayed) as RentValuePayed, sum(dbo.ContracttBillInstallmentsDone.WaterPayed) as WaterPayed, sum(dbo.ContracttBillInstallmentsDone.TelandNetPayed)as TelandNetPayed,sum(dbo.ContracttBillInstallmentsDone.ElectricPayed) ElectricPayed "
    sql = sql & "  FROM         dbo.TblContract INNER JOIN"
    sql = sql & "   dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo INNER JOIN"
    sql = sql & "  dbo.ContracttBillInstallmentsDone ON dbo.TblContractInstallments.id = dbo.ContracttBillInstallmentsDone.istallid"
    sql = sql & "   WHERE     (dbo.TblContractInstallments.ContNo = " & ContNo & ")"
      
    rs.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        WaterPayed = IIf(IsNull(rs("WaterPayed").value), 0, rs("WaterPayed").value)
        RentValuePayed = IIf(IsNull(rs("RentValuePayed").value), 0, rs("RentValuePayed").value)
        TelandNetPayed = IIf(IsNull(rs("TelandNetPayed").value), 0, rs("TelandNetPayed").value)
        ElectricPayed = IIf(IsNull(rs("ElectricPayed").value), 0, rs("ElectricPayed").value)
    Else
        WaterPayed = 0
        RentValuePayed = 0
        TelandNetPayed = 0
        ElectricPayed = 0
    End If
End Function
   
Public Function saveContractInstallments(NoteID As Double, _
                                         RecordDate As Date, _
                                         RecorddateH As String, _
                                         Optional total As Double, _
                                         Optional RentValuePayed As Double, _
                                         Optional CommissionsPayed As Double, _
                                         Optional InsurancePayed As Double, _
                                         Optional WaterPayed As Double, _
                                         Optional ElectricPayed As Double, _
                                         Optional TelandNetPayed As Double, _
                                         Optional EmpID As Integer, _
                                         Optional ContNo As Integer)

    Dim sql As String
    Dim i   As Integer
    Dim rs  As New ADODB.Recordset

    sql = "delete from ContracttBillInstallmentsDone  where Noteid=" & NoteID
    Cn.Execute sql
 
    StrSQL = "SELECT     * from  ContracttBillInstallmentsDone  Where (NoteID = -1)"
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With FrmCashing1.Grid3

        For i = 1 To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                total = val(.TextMatrix(i, .ColIndex("RentValuePayed"))) + val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) + val(.TextMatrix(i, .ColIndex("InsurancePayed"))) + val(.TextMatrix(i, .ColIndex("WaterPayed"))) + val(.TextMatrix(i, .ColIndex("ElectricPayed"))) + val(.TextMatrix(i, .ColIndex("TelandNetPayed"))) + val(.TextMatrix(i, .ColIndex("VATPayed")))
                total = total + val(.TextMatrix(i, .ColIndex("VATArboon"))) + val(.TextMatrix(i, .ColIndex("RentArbon"))) + val(.TextMatrix(i, .ColIndex("ServiceArbon"))) + val(.TextMatrix(i, .ColIndex("ElectricArbon"))) + val(.TextMatrix(i, .ColIndex("WaterArbon"))) + val(.TextMatrix(i, .ColIndex("InsuranceArbon"))) + val(.TextMatrix(i, .ColIndex("CommissionsArbon")))
                total = total + val(.TextMatrix(i, .ColIndex("OldValuePayed")))
                If total > 0 Then
                    rs.AddNew
                    rs("istallid").value = val(.TextMatrix(i, .ColIndex("id"))) '??? C????
                    rs("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo"))) '??? C?I???
                    
                    rs("NoteID").value = NoteID ' ??? ??I C????

                    If val(.TextMatrix(i, .ColIndex("Result"))) < total Then
                        '                        rs("Value").value = val(.TextMatrix(I, .ColIndex("Result")))
                    Else
                        '                        rs("Value").value = total
                    End If
                    rs("Value").value = total
                    rs("total").value = val(.TextMatrix(i, .ColIndex("total")))
                    rs("RentValuePayed").value = val(.TextMatrix(i, .ColIndex("RentValuePayed"))) + val(.TextMatrix(i, .ColIndex("RentArbon")))
                    rs("VATPayed").value = val(.TextMatrix(i, .ColIndex("VATPayed"))) + val(.TextMatrix(i, .ColIndex("VATArboon")))
                    rs("CommissionsPayed").value = val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) + val(.TextMatrix(i, .ColIndex("CommissionsArbon")))
                    rs("InsurancePayed").value = val(.TextMatrix(i, .ColIndex("InsurancePayed"))) + val(.TextMatrix(i, .ColIndex("InsuranceArbon")))
                    rs("WaterPayed").value = val(.TextMatrix(i, .ColIndex("WaterPayed"))) + val(.TextMatrix(i, .ColIndex("WaterArbon")))
                    rs("ElectricPayed").value = val(.TextMatrix(i, .ColIndex("ElectricPayed"))) + val(.TextMatrix(i, .ColIndex("ElectricArbon")))
                    rs("TelandNetPayed").value = val(.TextMatrix(i, .ColIndex("TelandNetPayed"))) + val(.TextMatrix(i, .ColIndex("ServiceArbon")))
                    rs("OldValuePayed").value = val(.TextMatrix(i, .ColIndex("OldValuePayed")))
                    rs("empid").value = EmpID
                    rs("paymentType").value = val(.TextMatrix(i, .ColIndex("CommisionTypesid")))
                    Dim Rent         As Double
                    Dim InternalComm As Double
                    Dim ExternalComm As Double
                    Dim Revenue      As Double
 
                    GetCommisionPercentages val(.TextMatrix(i, .ColIndex("CommisionTypesid"))), EmpID, Rent, InternalComm, ExternalComm, Revenue
                    Dim OutContract As Integer
                    OutContract = checkOutContract(ContNo)
                    Dim Commisionvalue As Single
                    If OutContract = 0 Then
                        Commisionvalue = ((rs("RentValuePayed").value) * Rent) + (rs("CommissionsPayed").value * InternalComm)
                    Else
                        Commisionvalue = ((rs("RentValuePayed").value) * Rent) + (rs("CommissionsPayed").value * InternalComm)
                    End If

                    rs("CommisionValue").value = Round(Commisionvalue, 2)
                    rs("RecordDate").value = RecordDate
                    rs("RecordDateH").value = RecorddateH
                    rs.update
                    total = total - val(.TextMatrix(i, .ColIndex("Result")))
                End If
            End If

        Next i

    End With

End Function

Public Function saveContractInstallmentsxxyd(NoteID As Double, _
                                             RecordDate As Date, _
                                             RecorddateH As String, _
                                             Optional total As Double, _
                                             Optional RentValuePayed As Double, _
                                             Optional CommissionsPayed As Double, _
                                             Optional InsurancePayed As Double, _
                                             Optional WaterPayed As Double, _
                                             Optional ElectricPayed As Double, _
                                             Optional TelandNetPayed As Double, _
                                             Optional EmpID As Integer, _
                                             Optional ContNo As Integer)

    Dim sql As String
    Dim i   As Integer
    Dim rs  As New ADODB.Recordset

    sql = "delete from ContracttBillInstallmentsDone  where Noteid=" & NoteID
    Cn.Execute sql
 
    StrSQL = "SELECT     * from  ContracttBillInstallmentsDone  Where (NoteID = -1)"
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With FrmCashing1.Grid3

        For i = 1 To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                total = val(.TextMatrix(i, .ColIndex("RentValuePayed"))) + val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) + val(.TextMatrix(i, .ColIndex("InsurancePayed"))) + val(.TextMatrix(i, .ColIndex("WaterPayed"))) + val(.TextMatrix(i, .ColIndex("ElectricPayed"))) + val(.TextMatrix(i, .ColIndex("TelandNetPayed")))
                total = total + val(.TextMatrix(i, .ColIndex("RentArbon"))) + val(.TextMatrix(i, .ColIndex("ServiceArbon"))) + val(.TextMatrix(i, .ColIndex("ElectricArbon"))) + val(.TextMatrix(i, .ColIndex("WaterArbon"))) + val(.TextMatrix(i, .ColIndex("InsuranceArbon"))) + val(.TextMatrix(i, .ColIndex("CommissionsArbon")))
                If total > 0 Then
                    rs.AddNew
                    rs("istallid").value = val(.TextMatrix(i, .ColIndex("id"))) '—ﬁ„ «·ﬁ”ÿ
                    rs("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo"))) '—ﬁ„ «·œ›⁄Â
                    
                    rs("NoteID").value = NoteID ' —ﬁ„ ”‰œ «·’—›

                    If val(.TextMatrix(i, .ColIndex("Result"))) < total Then
                        '                        rs("Value").value = val(.TextMatrix(I, .ColIndex("Result")))
                    Else
                        '                        rs("Value").value = total
                    End If
                    rs("Value").value = total
                    
                    rs("total").value = val(.TextMatrix(i, .ColIndex("total")))
                    rs("RentValuePayed").value = val(.TextMatrix(i, .ColIndex("RentValuePayed"))) + val(.TextMatrix(i, .ColIndex("RentArbon")))
                    rs("CommissionsPayed").value = val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) + val(.TextMatrix(i, .ColIndex("CommissionsArbon")))
                    rs("InsurancePayed").value = val(.TextMatrix(i, .ColIndex("InsurancePayed"))) + val(.TextMatrix(i, .ColIndex("InsuranceArbon")))
                    rs("WaterPayed").value = val(.TextMatrix(i, .ColIndex("WaterPayed"))) + val(.TextMatrix(i, .ColIndex("WaterArbon")))
                    rs("ElectricPayed").value = val(.TextMatrix(i, .ColIndex("ElectricPayed"))) + val(.TextMatrix(i, .ColIndex("ElectricArbon")))
                    rs("TelandNetPayed").value = val(.TextMatrix(i, .ColIndex("TelandNetPayed"))) + val(.TextMatrix(i, .ColIndex("ServiceArbon")))

                    rs("empid").value = EmpID
                    rs("paymentType").value = val(.TextMatrix(i, .ColIndex("CommisionTypesid")))
                    Dim Rent         As Double
                    Dim InternalComm As Double
                    Dim ExternalComm As Double
                    Dim Revenue      As Double
 
                    GetCommisionPercentages val(.TextMatrix(i, .ColIndex("CommisionTypesid"))), EmpID, Rent, InternalComm, ExternalComm, Revenue
                    Dim OutContract As Integer
                    OutContract = checkOutContract(ContNo)
                    Dim Commisionvalue As Single
                    If OutContract = 0 Then
                        Commisionvalue = ((rs("RentValuePayed").value) * Rent) + (rs("CommissionsPayed").value * InternalComm)
                    Else
                        Commisionvalue = ((rs("RentValuePayed").value) * Rent) + (rs("CommissionsPayed").value * InternalComm)
                    End If

                    rs("CommisionValue").value = Round(Commisionvalue, 2)
                    rs("RecordDate").value = RecordDate
                    rs("RecordDateH").value = RecorddateH
                    rs.update
                    total = total - val(.TextMatrix(i, .ColIndex("Result")))
                End If
            End If

        Next i

    End With

End Function

Public Function saveContractInstallmentsxx(NoteID As Double, _
                                           RecordDate As Date, _
                                           RecorddateH As String, _
                                           Optional total As Double, _
                                           Optional RentValuePayed As Double, _
                                           Optional CommissionsPayed As Double, _
                                           Optional InsurancePayed As Double, _
                                           Optional WaterPayed As Double, _
                                           Optional ElectricPayed As Double, _
                                           Optional TelandNetPayed As Double, _
                                           Optional EmpID As Integer, _
                                           Optional ContNo As Integer)

    Dim sql As String
    Dim i   As Integer
    Dim rs  As New ADODB.Recordset

    sql = "delete from ContracttBillInstallmentsDone  where Noteid=" & NoteID
    Cn.Execute sql
 
    StrSQL = "SELECT     * from  ContracttBillInstallmentsDone  Where (NoteID = -1)"
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With FrmCashing1.Grid3

        For i = 1 To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                total = val(.TextMatrix(i, .ColIndex("RentValuePayed"))) + val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) + val(.TextMatrix(i, .ColIndex("InsurancePayed"))) + val(.TextMatrix(i, .ColIndex("WaterPayed"))) + val(.TextMatrix(i, .ColIndex("ElectricPayed"))) + val(.TextMatrix(i, .ColIndex("TelandNetPayed")))
                If total > 0 Then
                    rs.AddNew
                    rs("istallid").value = val(.TextMatrix(i, .ColIndex("id"))) '—ﬁ„ «·ﬁ”ÿ
                    rs("InstallNo").value = val(.TextMatrix(i, .ColIndex("InstallNo"))) '—ﬁ„ «·œ›⁄Â
                    
                    rs("NoteID").value = NoteID ' —ﬁ„ ”‰œ «·’—›

                    If val(.TextMatrix(i, .ColIndex("Result"))) < total Then
                        '                        rs("Value").value = val(.TextMatrix(I, .ColIndex("Result")))
                    Else
                        '                        rs("Value").value = total
                    End If
                    rs("Value").value = total
                    
                    rs("total").value = val(.TextMatrix(i, .ColIndex("total")))
                    rs("RentValuePayed").value = val(.TextMatrix(i, .ColIndex("RentValuePayed")))
                    rs("CommissionsPayed").value = val(.TextMatrix(i, .ColIndex("CommissionsPayed")))
                    rs("InsurancePayed").value = val(.TextMatrix(i, .ColIndex("InsurancePayed")))
                    rs("WaterPayed").value = val(.TextMatrix(i, .ColIndex("WaterPayed")))
                    rs("ElectricPayed").value = val(.TextMatrix(i, .ColIndex("ElectricPayed")))
                    rs("TelandNetPayed").value = val(.TextMatrix(i, .ColIndex("TelandNetPayed")))

                    rs("empid").value = EmpID
                    rs("paymentType").value = val(.TextMatrix(i, .ColIndex("CommisionTypesid")))
                    Dim Rent         As Double
                    Dim InternalComm As Double
                    Dim ExternalComm As Double
                    Dim Revenue      As Double
 
                    GetCommisionPercentages val(.TextMatrix(i, .ColIndex("CommisionTypesid"))), EmpID, Rent, InternalComm, ExternalComm, Revenue
                    Dim OutContract As Integer
                    OutContract = checkOutContract(ContNo)
                    Dim Commisionvalue As Single
                    If OutContract = 0 Then
                        Commisionvalue = ((rs("RentValuePayed").value) * Rent) + (rs("CommissionsPayed").value * InternalComm)
                    Else
                        Commisionvalue = ((rs("RentValuePayed").value) * Rent) + (rs("CommissionsPayed").value * InternalComm)
                    End If

                    rs("CommisionValue").value = Round(Commisionvalue, 2)
                    rs("RecordDate").value = RecordDate
                    rs("RecordDateH").value = RecorddateH
                    rs.update
                    total = total - val(.TextMatrix(i, .ColIndex("Result")))
                End If
            End If

        Next i

    End With

End Function
  
Public Function saveprojectBillPayment(Optional TxtNoteSerial As String = "", _
                                       Optional total As Double, _
                                       Optional NoteID As Double)
 
    Dim sql As String
    Dim i   As Integer
    Dim rs  As New ADODB.Recordset
  
    sql = "delete from ProjectBillBuy where TxtNoteSerial='" & TxtNoteSerial & "'"
    Cn.Execute sql

    '   rs.Open "ProjectBillBuy", Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    sql = "select * from ProjectBillBuy where 1=-1"
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    With FrmCashing.Grid

        For i = 1 To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                If val(.TextMatrix(i, .ColIndex("ActualTotal"))) > 0 Then
                    rs.AddNew
                    rs("noteid").value = NoteID
                    rs("Bill_id").value = val(.TextMatrix(i, .ColIndex("id"))) '—ﬁ„ «·„Œ’’
                    rs("TxtNoteSerial").value = TxtNoteSerial ' —ﬁ„ ”‰œ «·’—›
                    
                    rs("NoteSerial1").value = val(.TextMatrix(i, .ColIndex("NoteSerial1"))) ' —ﬁ„ ”‰œ «·„” Œ·’
                    rs("ManualNO").value = val(.TextMatrix(i, .ColIndex("ManualNO"))) ' —ﬁ„ ”‰œ «·„” Œ·’ ManualNO ' —ﬁ„ ”‰œ «·„” Œ·’

                    'If val(.TextMatrix(i, .ColIndex("Result"))) < Total Then
                    '    rs("Value").value = val(.TextMatrix(i, .ColIndex("Result")))
                    'Else
                    '    rs("Value").value = Total
                    'End If
                    rs("Value").value = val(.TextMatrix(i, .ColIndex("ActualTotal")))
                    rs("bill_to").value = val(.TextMatrix(i, .ColIndex("bill_to")))
                    rs("total").value = val(.TextMatrix(i, .ColIndex("total")))
       
                    rs("RecordDate").value = Date
      
                    rs.update
                    'Total = Total - val(.TextMatrix(i, .ColIndex("Result")))
                End If
            End If

        Next i

    End With

End Function
 
Public Function getinsttPayedTocontract(Optional ContNo As Double = 0, _
                                        Optional ByRef RentValuePayed As Double, _
                                        Optional ByRef CommissionsPayed As Double, _
                                        Optional ByRef InsurancePayed As Double, _
                                        Optional ByRef WaterPayed As Double, _
                                        Optional ByRef ElectricPayed As Double, _
                                        Optional ByRef TelandNetPayed As Double, _
                                        Optional ByRef TotalOldValue As Double, _
                                        Optional NoteID As Double, _
                                        Optional Typ As Integer = 0, _
                                        Optional ByRef VATPayed As Double) As Double
    On Error Resume Next

    Dim total As Single

    Dim Rs3   As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    sql = "select sum(value) As total   ,sum(RentValuePayed) as RentValuePayed,sum(CommissionsPayed) as CommissionsPayed"
    sql = sql & "  ,sum(InsurancePayed) as InsurancePayed,sum(WaterPayed) as WaterPayed"
    sql = sql & "  ,sum(ElectricPayed) as ElectricPayed,sum(TelandNetPayed) as TelandNetPayed ,sum(OldValuePayed) as TotalOldValue ,sum(VATPayed) as VATPayed"
    sql = sql & "  from ContracttBillInstallmentsDone  where istallid=" & ContNo
    If Typ <> 0 Then
        sql = sql & " and NoteID <>" & NoteID & ""
 
    End If
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        total = 0
        RentValuePayed = 0
        CommissionsPayed = 0
        InsurancePayed = 0
        WaterPayed = 0
        ElectricPayed = 0
        TelandNetPayed = 0
        TotalOldValue = 0
        VATPayed = 0
    Else

        total = IIf(IsNull(Rs3("total").value), 0, Rs3("total").value)
        VATPayed = IIf(IsNull(Rs3("VATPayed").value), 0, Rs3("VATPayed").value)
        RentValuePayed = IIf(IsNull(Rs3("RentValuePayed").value), 0, Rs3("RentValuePayed").value)
        CommissionsPayed = IIf(IsNull(Rs3("CommissionsPayed").value), 0, Rs3("CommissionsPayed").value)
        InsurancePayed = IIf(IsNull(Rs3("InsurancePayed").value), 0, Rs3("InsurancePayed").value)
        WaterPayed = IIf(IsNull(Rs3("WaterPayed").value), 0, Rs3("WaterPayed").value)
        ElectricPayed = IIf(IsNull(Rs3("ElectricPayed").value), 0, Rs3("ElectricPayed").value)
        TelandNetPayed = IIf(IsNull(Rs3("TelandNetPayed").value), 0, Rs3("TelandNetPayed").value)
        TotalOldValue = IIf(IsNull(Rs3("TotalOldValue").value), 0, Rs3("TotalOldValue").value)
    End If

    Rs3.Close
    getinsttPayedTocontract = total
End Function

Public Function getinsttPayedTocontractxxy(Optional ContNo As Double = 0, _
                                           Optional ByRef RentValuePayed As Double, _
                                           Optional ByRef CommissionsPayed As Double, _
                                           Optional ByRef InsurancePayed As Double, _
                                           Optional ByRef WaterPayed As Double, _
                                           Optional ByRef ElectricPayed As Double, _
                                           Optional ByRef TelandNetPayed As Double) As Double
    On Error Resume Next

    Dim total As Single

    Dim Rs3   As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    sql = "select sum(value) As total   ,sum(RentValuePayed) as RentValuePayed,sum(CommissionsPayed) as CommissionsPayed"
    sql = sql & "  ,sum(InsurancePayed) as InsurancePayed,sum(WaterPayed) as WaterPayed"
    sql = sql & "  ,sum(ElectricPayed) as ElectricPayed,sum(TelandNetPayed) as TelandNetPayed"
    sql = sql & "  from ContracttBillInstallmentsDone  where istallid=" & ContNo
 
    '   Optional  byref RentValuePayed  as Double,Optional  byref CommissionsPayed  as Double,  Optional  byref InsurancePayed  as Double,Optional  byref WaterPayed  as Double,  Optional  byref ElectricPayed  as Double,Optional  byref TelandNetPayed  as Double
  
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        total = 0
        RentValuePayed = 0
        CommissionsPayed = 0
        InsurancePayed = 0
        WaterPayed = 0
        ElectricPayed = 0
        TelandNetPayed = 0
    Else
        
        total = IIf(IsNull(Rs3("total").value), 0, Rs3("total").value)
        RentValuePayed = IIf(IsNull(Rs3("RentValuePayed").value), 0, Rs3("RentValuePayed").value)
        CommissionsPayed = IIf(IsNull(Rs3("CommissionsPayed").value), 0, Rs3("CommissionsPayed").value)
        InsurancePayed = IIf(IsNull(Rs3("InsurancePayed").value), 0, Rs3("InsurancePayed").value)
        WaterPayed = IIf(IsNull(Rs3("WaterPayed").value), 0, Rs3("WaterPayed").value)
        ElectricPayed = IIf(IsNull(Rs3("ElectricPayed").value), 0, Rs3("ElectricPayed").value)
        TelandNetPayed = IIf(IsNull(Rs3("TelandNetPayed").value), 0, Rs3("TelandNetPayed").value)
  
    End If

    Rs3.Close
    '  getinsttPayedTocontract = Total
End Function
Public Function getBillPayedToproject(Optional bill_id As Integer = 0, _
                                      Optional TxtNoteSerial As String = "") As Double
    On Error Resume Next

    Dim total As Single

    Dim Rs3   As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    sql = "select sum(value) As total  from ProjectBillBuy  where Bill_id=" & bill_id
    If TxtNoteSerial <> "" Then
        sql = sql & "  AND TxtNoteSerial='" & TxtNoteSerial & "'"

    End If

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        total = 0
    Else
        total = IIf(IsNull(Rs3("total").value), 0, Rs3("total").value)
    End If

    Rs3.Close
    getBillPayedToproject = total
End Function

Public Function GetItemCost(Optional project_id As Integer = 0, _
                            Optional opr_fullcode As String = "") As Double
    Dim total_depit  As Single
    Dim total_credit As Single

    Dim total        As Single

    total = 0
 
    On Error Resume Next
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer
 
    sql = "select TOTAL  from projects_des    where    fullcode='" & opr_fullcode & "' AND project_id= " & project_id
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        total = 0
    Else
        total = IIf(IsNull(Rs3("TOTAL").value), 0, Rs3("TOTAL").value)
    End If

    Rs3.Close
 
    GetItemCost = total
End Function

Public Function checkitems(Optional project_id As Integer = 0, _
                           Optional opr_fullcode As String = "", _
                           Optional currentvalue As Double) As Boolean
    Dim ActualTotal As Double
    Dim total       As Double
    Dim StrMSG      As String
    Dim IntResult   As Integer
    checkitems = True

    If opr_fullcode = "" Then
        Exit Function
    End If
    If currentvalue = 0 Then
        Exit Function
    End If
 
    ActualTotal = get_balanceFromGl("", project_id, opr_fullcode)
    total = GetItemCost(project_id, opr_fullcode)
   
    If currentvalue + ActualTotal > total Then
        StrMSG = "Â–… «·ﬁÌ„… «·„œŒ·…  Ã⁄· Â–« «·»‰œ Ì ŒÿÌ «·„ Êﬁ⁄ ·Â" & CHR(13)
        StrMSG = StrMSG & "«·„ Êﬁ⁄ ··»‰œ : " & total & CHR(13)
        StrMSG = StrMSG & "«·„‰’—› Õ Ï «·«‰  ··»‰œ : " & ActualTotal & CHR(13)
        StrMSG = StrMSG & "«·›—ﬁ : " & total - ActualTotal & CHR(13)
        StrMSG = StrMSG & "›Ì Õ«·… «·„Ê«›ﬁ… ⁄·Ï «·ﬁÌ„… «·„ﬂ Ê»… ”ÌﬂÊ‰ «·«‰Õ—«› : " & total - (currentvalue + ActualTotal) & CHR(13)
     
        StrMSG = StrMSG & "‰⁄„" & "-" & "Â·  —Ìœ «· ﬂ„·Â ⁄·Ï «Ì Õ«·" & CHR(13)
        StrMSG = StrMSG & "·«" & "-" & " ⁄œÌ· «·ﬁÌ„… " & CHR(13)
 
        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

        Select Case IntResult

            Case vbYes
                checkitems = True
         
            Case vbCancel
                checkitems = False

            Case vbNo
                checkitems = False
        End Select

    Else
        checkitems = True
    End If

End Function

Public Function get_balanceFromGl(Account_code As String, _
                                  Optional project_id As Integer = 0, _
                                  Optional opr_fullcode As String = "", Optional isOpenBalance As Boolean = False) As Double
    Dim total_depit  As Double
    Dim total_credit As Double

    Dim total_depitOpen  As Double
    Dim total_creditOpen As Double

    Dim total        As Single

    total_credit = 0: total_depit = 0: total = 0
 
    On Error Resume Next
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select sum(value) As total_depit from DOUBLE_ENTREY_VOUCHERS  where Credit_Or_Debit=0 "

    If Account_code <> "" Then
        sql = sql & " and  Account_Code='" & Account_code & "'"
    End If
 
    If opr_fullcode <> "" Then
        sql = sql & " and  opr_fullcode='" & opr_fullcode & "'"
    End If
 
    If project_id <> 0 Then
        sql = sql & " and project_id=" & project_id
    End If

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        total_depit = 0
    Else
        total_depit = IIf(IsNull(Rs3("total_depit").value), 0, Rs3("total_depit").value)
    End If

    Rs3.Close
 
    '  If opr_fullcode <> "" Then
    '  get_balanceFromGl = total_depit
    '  Exit Function
    '  End If
    sql = "select sum(value) As total_credit from DOUBLE_ENTREY_VOUCHERS  where Credit_Or_Debit=1 and  Account_Code='" & Account_code & "'"
  
    If opr_fullcode <> "" Then
        sql = sql & " and  opr_fullcode='" & opr_fullcode & "'"
    End If
 
    If project_id <> 0 Then
        sql = sql & "  and project_id=" & project_id
    End If
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        total_credit = 0
    Else
        total_credit = IIf(IsNull(Rs3("total_credit").value), 0, Rs3("total_credit").value)
    End If

    Rs3.Close
 
 
  If isOpenBalance Then
    
            sql = "select sum(value) As total_depit from DOUBLE_ENTREY_VOUCHERS1  where Credit_Or_Debit=0 "
        
            If Account_code <> "" Then
                sql = sql & " and  Account_Code LIKE '" & Account_code & "%'"
            End If
         
            If opr_fullcode <> "" Then
                sql = sql & " and  opr_fullcode='" & opr_fullcode & "'"
            End If
            
            If project_id <> 0 Then
                sql = sql & "  and project_id=" & project_id
            End If
         
          
            Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
            If Rs3.RecordCount = 0 Then
                total_depitOpen = 0
            Else
                total_depitOpen = IIf(IsNull(Rs3("total_depit").value), 0, Rs3("total_depit").value)
            End If
        Rs3.Close
    End If
 
 
 
     
    If isOpenBalance Then
    
              sql = "select sum(value) As total_credit from DOUBLE_ENTREY_VOUCHERS1  where Credit_Or_Debit=1 and  Account_Code LIKE '" & Account_code & "%'"
     
         
            If opr_fullcode <> "" Then
                sql = sql & " and  opr_fullcode='" & opr_fullcode & "'"
            End If


         
            If opr_fullcode <> "" Then
                sql = sql & " and  opr_fullcode='" & opr_fullcode & "'"
            End If
            
            If project_id <> 0 Then
                sql = sql & "  and project_id=" & project_id
            End If
         
          
         
            Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
            If Rs3.RecordCount = 0 Then
                total_creditOpen = 0
            Else
                total_creditOpen = IIf(IsNull(Rs3("total_credit").value), 0, Rs3("total_credit").value)
            End If
        
            Rs3.Close
    End If
    
 
 
    get_balanceFromGl = total_depit + (total_depitOpen - total_creditOpen) - total_credit

    '1 œ∆«∆‰
    '2„œÌ‰
End Function

Public Function getBalanceWithOpeningBalance(StrAccountCode As String, _
                                             Branch As Long, _
                                             ToDate As Date, _
                                             Balance As Double, _
                                             balance_type As Integer)
    Dim openingbalacedate As Date
    getOpeningBalancedate , , , , year(ToDate), openingbalacedate
    '   getOpeningBalancedate DTPickerAccFrom.value, DTPickerAccTo.value, returnedfromdate, returnedTOdate
    update_account_opening_balance StrAccountCode, True, openingbalacedate, ToDate, Branch, openingbalacedate, Balance, balance_type
     
End Function

Public Function get_balanceFromGlNew(Account_code As String, _
                                     Optional project_id As Integer = 0, _
                                     Optional opr_fullcode As String = "", _
                                     Optional havedate As Boolean, _
                                     Optional BegineDate As Date, _
                                     Optional EndDate As Date, _
                                     Optional ByRef total_depit As Double, _
                                     Optional ByRef total_credit As Double, _
                                     Optional ByRef total As Double, _
                                     Optional branch_id As Long = 0, Optional isOpenBalance As Boolean = False)
 
    On Error Resume Next
    Dim total_creditOpen As Double
    Dim total_depitOpen As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select sum(value) As total_depit from DOUBLE_ENTREY_VOUCHERS  where Credit_Or_Debit=0 "

    If Account_code <> "" Then
        sql = sql & " and  Account_Code LIKE '" & Account_code & "%'"
    End If
 
    If opr_fullcode <> "" Then
        sql = sql & " and  opr_fullcode='" & opr_fullcode & "'"
    End If
 
    If project_id <> 0 Then
        sql = sql & " and project_id=" & project_id
    End If
 
    If branch_id <> 0 Then
        sql = sql & " and branch_id=" & branch_id
    End If
 
    If havedate = True Then
   
        sql = sql + " and     RecordDate >=" & SQLDate(BegineDate, True) & ""
   
        sql = sql + " and RecordDate <=" & SQLDate(EndDate, True) & ""
  
    End If
  
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        total_depit = 0
    Else
        total_depit = IIf(IsNull(Rs3("total_depit").value), 0, Rs3("total_depit").value)
    End If

    Rs3.Close
    
    
    If isOpenBalance Then
    
                sql = "select sum(value) As total_depit from DOUBLE_ENTREY_VOUCHERS1  where Credit_Or_Debit=0 "
        
            If Account_code <> "" Then
                sql = sql & " and  Account_Code LIKE '" & Account_code & "%'"
            End If
         
            If opr_fullcode <> "" Then
                sql = sql & " and  opr_fullcode='" & opr_fullcode & "'"
            End If
         
            If project_id <> 0 Then
                sql = sql & " and project_id=" & project_id
            End If
         
            If branch_id <> 0 Then
                sql = sql & " and branch_id=" & branch_id
            End If
         
            If havedate = True Then
           
                sql = sql + " and     RecordDate >=" & SQLDate(BegineDate, True) & ""
           
                sql = sql + " and RecordDate <=" & SQLDate(EndDate, True) & ""
          
            End If
          
            Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
            If Rs3.RecordCount = 0 Then
                total_depitOpen = 0
            Else
                total_depitOpen = IIf(IsNull(Rs3("total_depit").value), 0, Rs3("total_depit").value)
            End If
        
    End If
 
    '  If opr_fullcode <> "" Then
    '  get_balanceFromGl = total_depit
    '  Exit Function
    '  End If
    sql = "select sum(value) As total_credit from DOUBLE_ENTREY_VOUCHERS  where Credit_Or_Debit=1 and  Account_Code LIKE '" & Account_code & "%'"
  
    If opr_fullcode <> "" Then
        sql = sql & " and  opr_fullcode='" & opr_fullcode & "'"
    End If
 
    If project_id <> 0 Then
        sql = sql & "  and project_id=" & project_id
    End If
 
    If branch_id <> 0 Then
        sql = sql & " and branch_id=" & branch_id
    End If
 
    If havedate = True Then
   
        sql = sql + " and     RecordDate >=" & SQLDate(BegineDate, True) & ""
   
        sql = sql + " and RecordDate <=" & SQLDate(EndDate, True) & ""
  
    End If
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        total_credit = 0
    Else
        total_credit = IIf(IsNull(Rs3("total_credit").value), 0, Rs3("total_credit").value)
    End If

    Rs3.Close
    
    If isOpenBalance Then
    
              sql = "select sum(value) As total_credit from DOUBLE_ENTREY_VOUCHERS1  where Credit_Or_Debit=1 and  Account_Code LIKE '" & Account_code & "%'"
          
            If opr_fullcode <> "" Then
                sql = sql & " and  opr_fullcode='" & opr_fullcode & "'"
            End If
         
            If project_id <> 0 Then
                sql = sql & "  and project_id=" & project_id
            End If
         
            If branch_id <> 0 Then
                sql = sql & " and branch_id=" & branch_id
            End If
         
            If havedate = True Then
           
                sql = sql + " and     RecordDate >=" & SQLDate(BegineDate, True) & ""
           
                sql = sql + " and RecordDate <=" & SQLDate(EndDate, True) & ""
          
            End If
         
            Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
            If Rs3.RecordCount = 0 Then
                total_creditOpen = 0
            Else
                total_creditOpen = IIf(IsNull(Rs3("total_credit").value), 0, Rs3("total_credit").value)
            End If
        
            Rs3.Close
    End If
    
    
    Dim openingbalacedate As Date
    '     getOpeningBalancedate openingbalacedate
    '    update_account_opening_balance Account_Code, True, BegineDate, Enddate, branch_id, openingbalacedate
     
    total = total_depit - total_credit
    Dim openingbalancevalue As Double
    openingbalancevalue = (total_depitOpen - total_creditOpen)
    total = total + openingbalancevalue
    'get_balanceFromGl = total_depit - total_credit

    '1 œ∆«∆‰
    '2„œÌ‰
End Function

Public Function check_opening_balance_notes()
    On Error Resume Next
    Dim departement_name As String
    departement_name = 1
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select * from notes1  where NoteID=1"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
 
        Dim RsNetes As ADODB.Recordset
        Set RsNetes = New ADODB.Recordset
        RsNetes.Open "notes1", Cn, adOpenStatic, adLockOptimistic, adCmdTable

        RsNetes.AddNew
        RsNetes("NoteID").value = 1
        RsNetes("NoteType").value = 101
 
        RsNetes("NoteSerial").value = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & 1
        RsNetes("NoteSerial1").value = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & 1
    
        RsNetes("numbering_type").value = sand_numbering_type(0) ' „”·”· «·ﬁÌœ
        RsNetes("numbering_type1").value = sand_numbering_type(3) ' „”·”· «·”‰œ
    
        RsNetes("sanad_year").value = year(Now)
        RsNetes("sanad_month").value = Month(Now)
        RsNetes("foxy_no").value = 1
        RsNetes("NoteDate").value = Now
        RsNetes("Note_Value").value = 0
        RsNetes("Double_Entry_Vouchers_ID").value = 0
        RsNetes("DAWRY").value = Check4.value
        RsNetes("KALEB").value = Check3.value
    
        RsNetes("Remark").value = "Opening Balance"
    
        RsNetes.update
 
    End If
  
End Function

Public Function update_account_balance(Account_code As String, _
                                       Optional havedate As Boolean, _
                                       Optional FromDate As Date, _
                                       Optional ToDate As Date, _
                                       Optional Branch As Long = 0, _
                                       Optional ByRef openingBalanceDate As Date, _
                                       Optional ByRef opening_balance As Double, _
                                       Optional ByRef opening_balanceType As Integer)
 
    Dim Balance As Double
    opening_balance = 0

    'opening_balance = get_balance(Account_Code, opening_balanceType, havedate, OpeningBalanceDate, OpeningBalanceDate, branch)
 
    get_balanceFromGlNew Account_code, , , True, FromDate, ToDate, , , Balance
    opening_balance = Balance ' + opening_balance

    If opening_balance > 0 Then
        opening_balanceType = 0
    Else
        opening_balanceType = 1
    End If

    sgl = "update  ACCOUNTS  set opening_balance= opening_balance+ " & Balance & ", opening_balance_type=" & opening_balanceType & " where  Account_Code='" & Account_code & "'"
    Cn.Execute sgl, , adExecuteNoRecords

End Function

Public Function update_account_opening_balance(Account_code As String, _
                                               Optional havedate As Boolean, _
                                               Optional FromDate As Date, _
                                               Optional ToDate As Date, _
                                               Optional Branch As Long = 0, _
                                               Optional ByRef openingBalanceDate As Date, _
                                               Optional ByRef opening_balance As Double, _
                                               Optional ByRef opening_balanceType As Integer)
 
    Dim Balance As Double
    opening_balance = 0
 
    opening_balance = get_balance(Account_code, opening_balanceType, havedate, openingBalanceDate, openingBalanceDate, Branch)

    If SystemOptions.UserInterface = ArabicInterface Then
        openingbalanceDes = "—’Ìœ «›  «ÕÌ «›  «ÕÌ ›Ì" & openingBalanceDate
    Else
        openingbalanceDes = "Opening Balance In " & openingBalanceDate
    End If

    If openingBalanceDate <> FromDate Then

        If SystemOptions.UserInterface = ArabicInterface Then
            openingbalanceDes = "—’Ìœ Õ Ï    " & FromDate - 1
        Else
            openingbalanceDes = " Balance Untill " & FromDate - 1
        End If

        get_balanceFromGlNew Account_code, , , True, openingBalanceDate, FromDate - 1, , , Balance
        opening_balance = Balance + opening_balance

        If opening_balance > 0 Then
            opening_balanceType = 0
        Else
            opening_balanceType = 1
        End If

    End If

    sgl = "update  ACCOUNTS  set opening_balance=  " & opening_balance & ", opening_balance_type=" & opening_balanceType & " where  Account_Code='" & Account_code & "'"
    Cn.Execute sgl, , adExecuteNoRecords

End Function

Public Function setfoxy_Line() As Double
    Dim last_line_id As Double

    last_line_id = CStr(new_id("foxy", "id1", "", True))
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id1").value = last_line_id
    setfoxy_Line = last_line_id
    rs.update
    
End Function

Public Function detect_employee_work_type() As Integer
    '0 work with out  create accounts
    '1 work with create accounts
   
    Dim rsOut As New ADODB.Recordset
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    detect_employee_work_type = 0

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!Create_employee_account = True Then
                
            detect_employee_work_type = 1
                
        Else
            detect_employee_work_type = 0
        End If
    
    End If

End Function

Public Function detect_inventory_work_type() As Integer
    '1 work with branch
    '2 work with inventory
    '3 work with groups
    Dim rsOut As New ADODB.Recordset
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!opt_group = True Then
     
            If rsOut!Opt_Inventory_create_account = 1 Then
                detect_inventory_work_type = 2
            ElseIf rsOut!opt_inv_and_branch_create_account = 1 Then
                detect_inventory_work_type = 3
            End If
     
        Else
            detect_inventory_work_type = 1
        End If
    End If

End Function

Public Function sand_numbering_type(Sanad_No As Integer) As Integer
    On Error Resume Next
    Dim departement_name As String
    departement_name = 1
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select * from sanad_numbering where branch_no=" & Current_branch & " and  sanad_no=" & Sanad_No
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        numbering_id = 0
        Exit Function
    End If
    sand_numbering_type = IIf(IsNull(Rs3("numbering_id").value), 0, Rs3("numbering_id").value)
    Rs3.Close

End Function

Public Function Get_currency_txt() As String
    On Error Resume Next
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select * from currency where basic=1"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Get_currency_txt = ""
        Exit Function
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Get_currency_txt = IIf(IsNull(Rs3("name").value), "", Rs3("name").value)
        Exit Function
    Else
        Get_currency_txt = IIf(IsNull(Rs3("namee").value), "", Rs3("namee").value)
        Exit Function
    End If

    Rs3.Close

End Function

Public Function ShowGL_ccOpening(Optional note_serial As String, _
                                 Optional date_from, _
                                 Optional NoteType As Integer, _
                                 Optional ByRef LngNoteID As Integer)
    
    Dim MySQL          As String
    Dim RsData         As New ADODB.Recordset
    Dim xApp           As New CRAXDRT.Application
    Dim xReport        As CRAXDRT.Report
    Dim CViewer        As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName    As String
    Dim Msg            As String
 
    'MySQL = "SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code, "
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.[Value], dbo.DOUBLE_ENTREY_VOUCHERS1.Credit_Or_Debit,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Description, dbo.DOUBLE_ENTREY_VOUCHERS1.RecordDate,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.ReceiptID, dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS1.OperaID,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.Transaction_ID, dbo.DOUBLE_ENTREY_VOUCHERS1.AdvanceID, dbo.DOUBLE_ENTREY_VOUCHERS1.UserID,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.Posted, dbo.DOUBLE_ENTREY_VOUCHERS1.PostedDate, dbo.DOUBLE_ENTREY_VOUCHERS1.PostedUserID,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Interval_ID, dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_Serial, dbo.DOUBLE_ENTREY_VOUCHERS1.credit_value,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.depet_value, dbo.DOUBLE_ENTREY_VOUCHERS1.des, dbo.DOUBLE_ENTREY_VOUCHERS1.currency,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.project_bill_no, dbo.DOUBLE_ENTREY_VOUCHERS1.valuee, dbo.DOUBLE_ENTREY_VOUCHERS1.rate,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Descriptione, dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No1,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.notes_all, dbo.DOUBLE_ENTREY_VOUCHERS1.project_id, dbo.DOUBLE_ENTREY_VOUCHERS1.opr_fullcode,"
    'MySQL = MySQL & " dbo.Notes1.Note_Value, dbo.Notes1.BankID, dbo.Notes1.ChqueNum, dbo.Notes1.DueDate, dbo.Notes1.NoteHijriDate, dbo.Notes1.MaintananceID,"
    'MySQL = MySQL & " dbo.Notes1.Member_ID, dbo.Notes1.ExpensesID, dbo.Notes1.CashingType, dbo.Notes1.CusID, dbo.Notes1.BoxID, dbo.Notes1.RevenuesID,"
    'MySQL = MySQL & " dbo.Notes1.RetrunNoteID, dbo.Notes1.NoteCashingType, dbo.Notes1.NotePosted, dbo.Notes1.PostedBy, dbo.Notes1.PostDate, dbo.Notes1.NumOrderInpot,"
    'MySQL = MySQL & " dbo.Notes1.ked_type, dbo.Notes1.Buy, dbo.Notes1.numbering_type, dbo.Notes1.sanad_year, dbo.Notes1.sanad_month, dbo.Notes1.type, dbo.Notes1.branch_no,"
    'MySQL = MySQL & " dbo.Notes1.user_name, dbo.Notes1.DEPARTEMENT, dbo.Notes1.sanad_type, dbo.Notes1.sanad_source, dbo.Notes1.DAWRY, dbo.Notes1.KALEB,"
    'MySQL = MySQL & " dbo.Notes1.projectAccountCode, dbo.Notes1.foxy_no, dbo.Notes1.person, dbo.Notes1.project_Expensen_account, dbo.Notes1.salary, dbo.Notes1.displayed,"
    'MySQL = MySQL & " dbo.Notes1.Adv_payment_value, dbo.Notes1.salary_or_advance, dbo.Notes1.EmpAccountCode, dbo.Notes1.project_depit_or_credit, dbo.Notes1.Cus_or_sub,"
    'MySQL = MySQL & " dbo.Notes1.numbering_type1, dbo.Notes1.NoteSerial1, dbo.Notes1.general_cost_center, dbo.Notes1.too, dbo.Notes1.NoteID,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.bill_id, dbo.ACCOUNTS.Account_NameEng, dbo.TblNotesTypes.NotesTypeName,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.opening_balance_voucher_id, dbo.Notes1.NoteSerial, dbo.Notes1.NoteDate,"
    'MySQL = MySQL & " dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No , dbo.Notes1.note_value_by_characters, dbo.Notes1.NoteType, dbo.Notes1.remark"
    'MySQL = MySQL & " FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 INNER JOIN"
    'MySQL = MySQL & " dbo.Notes1 ON dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID = dbo.Notes1.NoteID LEFT OUTER JOIN"
    'MySQL = MySQL & " dbo.TblNotesTypes ON dbo.Notes1.NoteType = dbo.TblNotesTypes.NotesType LEFT OUTER JOIN"
    'MySQL = MySQL & " dbo.TblBranchesData ON dbo.Notes1.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    'MySQL = MySQL & " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code = dbo.ACCOUNTS.Account_Code"
    'MySQL = MySQL & "   where  noteserial='" & note_serial & "'"
    updateAutoOpeningBalanceLineNo LngNoteID

    MySQL = "SELECT     TOP 100 PERCENT dbo.marakes_taklefa_temp.cost_center_id, dbo.marakes_taklefa_temp.cost_center, dbo.marakes_taklefa_temp.[value] AS CC_Valie, "
    MySQL = MySQL & "  dbo.marakes_taklefa_temp.depit_or_credit, dbo.ACCOUNTS.Account_Name, dbo.marakes_taklefa_temp.Project__code, dbo.marakes_taklefa_temp.Project_name,"
    MySQL = MySQL & "  dbo.ACCOUNTS.Account_Serial, dbo.marakes_taklefa_temp.Description, dbo.marakes_taklefa_temp.opr_type, dbo.marakes_taklefa_temp.[value] AS cc_valie1,"
    MySQL = MySQL & "  dbo.marakes_taklefa_temp.[value] AS DEV_Value1, dbo.marakes_taklefa_temp.[value] AS DEV_Value2, dbo.ACCOUNTS.Account_NameEng,"
    MySQL = MySQL & "  dbo.TblNotesTypes.NotesTypeName, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblNotesTypes.NotesTypeNamee,"
    MySQL = MySQL & "  dbo.TblBranchesData.branch_id, dbo.marakes_taklefa_temp.opr_id, dbo.marakes_taklefa_temp.account_no, dbo.marakes_taklefa_temp.account_type,"
    MySQL = MySQL & "  dbo.marakes_taklefa_temp.line_no, dbo.marakes_taklefa_temp.kedno, dbo.marakes_taklefa_temp.Foxy_no, dbo.marakes_taklefa_temp.user_id,"
    MySQL = MySQL & "  dbo.marakes_taklefa_temp.ok, dbo.marakes_taklefa_temp.record_date, dbo.marakes_taklefa_temp.general_des, dbo.marakes_taklefa_temp.auto_des,"
    MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS1.[Value], dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No,"
    MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS1.Credit_Or_Debit,"
    MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Description, dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID,"
    MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS1.ReceiptID, dbo.DOUBLE_ENTREY_VOUCHERS1.OperaID, dbo.DOUBLE_ENTREY_VOUCHERS1.Transaction_ID,"
    MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS1.AdvanceID, dbo.DOUBLE_ENTREY_VOUCHERS1.UserID, dbo.DOUBLE_ENTREY_VOUCHERS1.Posted,"
    MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS1.PostedDate, dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Interval_ID, dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_Serial,"
    MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS1.PostedUserID, dbo.DOUBLE_ENTREY_VOUCHERS1.credit_value, dbo.DOUBLE_ENTREY_VOUCHERS1.depet_value,"
    MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS1.des, dbo.DOUBLE_ENTREY_VOUCHERS1.currency, dbo.DOUBLE_ENTREY_VOUCHERS1.valuee,"
    MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS1.project_bill_no, dbo.DOUBLE_ENTREY_VOUCHERS1.rate, dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No1,"
    MySQL = MySQL & "    dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Descriptione, dbo.DOUBLE_ENTREY_VOUCHERS1.notes_all,"
    MySQL = MySQL & "   dbo.DOUBLE_ENTREY_VOUCHERS1.project_id, dbo.DOUBLE_ENTREY_VOUCHERS1.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS1.bill_id,"
    MySQL = MySQL & "   dbo.DOUBLE_ENTREY_VOUCHERS1.opening_balance_voucher_id, dbo.DOUBLE_ENTREY_VOUCHERS1.FixedAssetId,"
    MySQL = MySQL & "    dbo.DOUBLE_ENTREY_VOUCHERS1.FixedAssetgroupid, dbo.DOUBLE_ENTREY_VOUCHERS1.FixedAssetbranch_id, dbo.Notes1.NoteID, dbo.Notes1.NoteDate,"
    MySQL = MySQL & "    dbo.Notes1.NoteType, dbo.Notes1.Note_Value, dbo.Notes1.BankID, dbo.Notes1.ChqueNum, dbo.Notes1.DueDate, dbo.Notes1.NoteHijriDate,"
    MySQL = MySQL & "   dbo.Notes1.MaintananceID, dbo.Notes1.Member_ID, dbo.Notes1.Remark, dbo.Notes1.ExpensesID, dbo.Notes1.CashingType, dbo.Notes1.CusID, dbo.Notes1.BoxID,"
    MySQL = MySQL & "  dbo.Notes1.RevenuesID, dbo.Notes1.RetrunNoteID, dbo.Notes1.NoteCashingType, dbo.Notes1.NotePosted, dbo.Notes1.PostedBy, dbo.Notes1.PostDate,"
    MySQL = MySQL & "  dbo.Notes1.NumOrderInpot, dbo.Notes1.Buy, dbo.Notes1.ked_type, dbo.Notes1.numbering_type, dbo.Notes1.sanad_year, dbo.Notes1.sanad_month, dbo.Notes1.type,"
    MySQL = MySQL & "  dbo.Notes1.branch_no, dbo.Notes1.user_name, dbo.Notes1.DEPARTEMENT, dbo.Notes1.sanad_type, dbo.Notes1.sanad_source, dbo.Notes1.DAWRY,"
    MySQL = MySQL & "  dbo.Notes1.KALEB, dbo.Notes1.projectAccountCode, dbo.Notes1.person, dbo.Notes1.project_Expensen_account, dbo.Notes1.salary, dbo.Notes1.displayed,"
    MySQL = MySQL & "  dbo.Notes1.Adv_payment_value, dbo.Notes1.note_value_by_characters, dbo.Notes1.too, dbo.Notes1.general_cost_center, dbo.Notes1.numbering_type1,"
    MySQL = MySQL & "  dbo.Notes1.Cus_or_sub, dbo.Notes1.project_depit_or_credit, dbo.Notes1.EmpAccountCode, dbo.Notes1.salary_or_advance, dbo.Notes1.general_des_notes,"
    MySQL = MySQL & "  dbo.Notes1.BTCashAccountcode , dbo.Notes1.NoteSerial, dbo.Notes1.NoteSerial1"
    MySQL = MySQL & "  FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 INNER JOIN"
    MySQL = MySQL & "  dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code = dbo.ACCOUNTS.Account_Code INNER JOIN"
    MySQL = MySQL & "  dbo.Notes1 ON dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID = dbo.Notes1.NoteID INNER JOIN"
    MySQL = MySQL & "  dbo.TblNotesTypes ON dbo.Notes1.NoteType = dbo.TblNotesTypes.NotesType LEFT OUTER JOIN"
    MySQL = MySQL & "  dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & "  dbo.marakes_taklefa_temp ON dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No1 = dbo.marakes_taklefa_temp.line_no"
    'Where (dbo.Notes1.NoteSerial = 20121)
    'ORDER BY dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial

    MySQL = MySQL & "   where   dbo.Notes1.NoteID=" & LngNoteID

    If LngNoteID = 1 Then
        MySQL = MySQL + " Order By  DOUBLE_ENTREY_VOUCHERS1.Credit_Or_Debit , DOUBLE_ENTREY_VOUCHERS1.value"
        'strsql = strsql + " Order By DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No "
    Else

        MySQL = MySQL + " Order By DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No "

    End If

    'MySQL = MySQL & " WHERE     (dbo.Notes1.NoteSerial = '201201')"
    'MySQL = MySQL & " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS1.opening_balance_voucher_id"
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "GL_ccOPeningN.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "GL_ccOPeningNE.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Public Function ShowGL_cc(Optional note_serial As String, _
                          Optional date_from, _
                          Optional NoteType As Integer, _
                          Optional notes_id As Variant = -1, _
                          Optional notevale As String, _
                          Optional NoteSerial1 As Double, Optional ByVal mFileName As String = "", Optional ByVal mFileNameShow As String = "")
    
    If DoPremis(Do_Print, "FrmAccEditJournal", True) = False Then
        Exit Function
    End If
            
    Dim MySQL          As String
    Dim RsData         As New ADODB.Recordset
    Dim xApp           As New CRAXDRT.Application
    Dim xReport        As CRAXDRT.Report
    Dim CViewer        As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName    As String
    Dim Msg            As String
    
    Dim sql            As String
    Dim rs             As New ADODB.Recordset
    Dim TotalValue     As Double
    Dim TotalText      As String

    'MySQL = "Select * From GL_CC  where notetype=" & NoteType & "  and noteserial='" & note_serial & "' order by dev_id_line_no"
    MySQL = "Select * From GL_CC  where  noteserial='" & val(note_serial) & "'"

    If (NoteType = 80) Then
        MySQL = MySQL & " and notetype = 80 "
        If NoteSerial1 <> 0 Then
            MySQL = MySQL & " AND (NoteSerial1 = " & NoteSerial1 & ")"
        End If
  
    ElseIf (NoteType = 8063) Then
        MySQL = MySQL & " and notetype = 8063 "
        If NoteSerial1 <> 0 Then
            MySQL = MySQL & " AND (NoteSerial1 = " & NoteSerial1 & ")"
        End If
            
    ElseIf (NoteType = 3) Then
        MySQL = MySQL & " and notetype = 3"
        If NoteSerial1 <> 0 Then
            MySQL = MySQL & " AND (NoteSerial1 = " & NoteSerial1 & ")"
        End If
    ElseIf (NoteType = 350) Then
        MySQL = MySQL & " and notetype = 350"
        If NoteSerial1 <> 0 Then
            MySQL = MySQL & " AND (NoteSerial1 = " & NoteSerial1 & ")"
        End If
            
    ElseIf (NoteType = 57) Then
        MySQL = MySQL & " and notetype = 57"
        If NoteSerial1 <> 0 Then
            MySQL = MySQL & " AND (NoteSerial1 = " & NoteSerial1 & ")"
        End If
            
    End If

    MySQL = MySQL & " order by dev_id_line_no"

    If notes_id <> -1 Then
    
        MySQL = "Select * From GL_CC  where  notes_id=" & notes_id & " order by dev_id_line_no"
    Else
    
    End If
    If NoteType = 53 Then
        Account_Code_dynamic = get_account_code_branch(72, my_branch)
        MySQL = "SELECT     TOP 100 PERCENT *, dbo.BanksData.BankName AS BankNameA, dbo.BanksData.BankNamee AS BankNameeE"
        MySQL = MySQL & " FROM         dbo.GL_CC LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.BanksData ON dbo.GL_CC.BankID = dbo.BanksData.BankID"
        MySQL = MySQL & " WHERE     (dbo.GL_CC.NoteSerial = '" & note_serial & "')"

        MySQL = MySQL & " and GL_CC.Account_Code not in ("
        MySQL = MySQL & " SELECT     Account_Code "
        MySQL = MySQL & "  From dbo.ACCOUNTS"
        MySQL = MySQL & "  WHERE     (Parent_Account_Code =  '" & Account_Code_dynamic & "'))"

        MySQL = MySQL & " ORDER BY dbo.GL_CC.DEV_ID_Line_No"
 
    End If
 
    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If

    If SystemOptions.DateOpt = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "GL_cc.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "GL_CCE.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\" & "GL_ccH.rpt"
        Else
            StrFileName = App.path & "\Reports\" & "GL_CCH.rpt"
        End If

    End If

    If NoteType = 53 Then
        If SystemOptions.DateOpt = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\" & "ExpensesVchrMultinew.rpt"
            Else
                StrFileName = App.path & "\Reports\" & "ExpensesVchrMultinew.rpt"
            End If

        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\" & "ExpensesVchrMultiH.rpt"
            Else
                StrFileName = App.path & "\Reports\" & "ExpensesVchrMultiHE.rpt"
            End If

        End If
    End If

    If NoteType = 57 Then
        If SystemOptions.DateOpt = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\" & "GL_ccM.rpt"
            Else
                StrFileName = App.path & "\Reports\" & "GL_ccMe.rpt"
            End If

        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\" & "GL_ccM.rpt"
            Else
                StrFileName = App.path & "\Reports\" & "GL_ccM.rpt"
            End If

        End If
    End If

    If mFileName <> "" Then
        StrFileName = App.path & "\Reports\" & mFileName
    End If
    If Dir(StrFileName) = "" Then
        'Getf Msgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo



If Trim(mFileNameShow) <> "" Then
    For i = 1 To xReport.FormulaFields.count
        Select Case xReport.FormulaFields.Item(i).Name
        Case "{@mFileNameShow}"
            xReport.FormulaFields.Item(i).text = "'" & mFileNameShow & "'"
       
            
        End Select
    Next i
End If

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        ' xReport.ParameterFields(4).AddCurrentValue get_branch_name(Val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    If NoteType = 53 Then
        xReport.ParameterFields(6).AddCurrentValue CStr(notevale)
    End If
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False

    If PrintBranchINGE = True Then
        xReport.ParameterFields(4).AddCurrentValue "1"
    Else
        xReport.ParameterFields(4).AddCurrentValue "0"
    End If

    If PrintCCinGE = True Then
        xReport.ParameterFields(5).AddCurrentValue "1"
    Else
        xReport.ParameterFields(5).AddCurrentValue "0"
    End If

    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    '       CreateLogo xReport, val(Current_branch)
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL
 
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Public Function SwitchKeyboardLang(ByVal strLangID As String)
    'Returns TRUE when the KeyboardLayout was set properly, FALSE otherwise
    Dim strRet As String
    On Error Resume Next
    strRet = String(9, 0)
    GetKeyboardLayoutName strRet

    If strRet = (strLangID & CHR(0)) Then
        ' you are try to switch to the already selected language
        ' so return without doing anything
        SwitchKeyboardLang = True
        Exit Function
    Else
        strRet = String(9, 0)
        strRet = LoadKeyboardLayout((strLangID & CHR(0)), KLF_ACTIVATE)
    End If

    GetKeyboardLayoutName strRet ' Test if switch successed

    If strRet = (strLangID) Then
        SwitchKeyboardLang = True
    End If

End Function

Public Function ExactAge(date1 As Variant, _
   date2 As Variant) As String

    Dim yer  As Integer, mon As Integer, d As Integer
    Dim dt   As Date
    Dim sAns As String

    If Not IsDate(date1) Then
        ExactAge = ""
        Exit Function
    End If
    dt = CDate(date1)

    If dt > date2 Then
        ExactAge = "0-0-0"
        Exit Function
    End If
    yer = year(dt)
    mon = Month(dt)
    d = day(dt)
    
    yer = year(date2) - yer
    mon = Month(date2) - mon
    d = day(date2) - d

    If Sgn(d) = -1 Then
        d = 30 - Abs(d)
        mon = mon - 1
    End If

    If Sgn(mon) = -1 Then
        mon = 12 - Abs(mon)
        yer = yer - 1
    End If
    
    sAns = yer & "-" & mon & "-" & d

    ExactAge = sAns

End Function

Public Function Notes_codingByUser(my_branch As Integer, _
                                   date1 As Date, _
                                   Optional departement_name As Integer = 1) As String
    On Error Resume Next
    Dim start_at       As Double
    Dim end_at         As Single
    Dim auto_sanad_no  As String
    Dim NO             As Single
    Dim numbering_type As Integer
    auto_sanad_no = ""

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset

    Dim sql                 As String
    Dim i                   As Integer
    Dim JLCodeBasedOnBranch As Boolean
    JLCodeBasedOnBranch = False
    '**********************
    'JLCodeBasedOnBranch = True

    Dim mNoOfUser     As Integer
    Dim mUserIdSerial As String
    Dim mLenUser      As Integer

    Dim mFormatUser   As String

    mFormatUser = ""
    Dim mm As Integer
    mm = 1
    For mm = 1 To SystemOptions.NoOFDigitUserVouc
        If user_id < 9 And SystemOptions.NoOFDigitUserVouc > 1 And SystemOptions.NoOFDigitUserVouc > mm Then
            mFormatUser = mFormatUser & "0"
        ElseIf user_id > 9 And user_id < 100 And SystemOptions.NoOFDigitUserVouc > 1 And SystemOptions.NoOFDigitUserVouc > mm + 1 Then
            mFormatUser = mFormatUser & "0"
        ElseIf user_id > 99 And SystemOptions.NoOFDigitUserVouc > 1 And SystemOptions.NoOFDigitUserVouc > mm + 2 Then
            mFormatUser = mFormatUser & "0"
        End If

    Next

    'If user_id > 9 And user_id < 100 Then
    mUserIdSerial = mFormatUser & user_id
    'End If

    If SystemOptions.JLCodeBasedOnBranch = True Then
    
        Dim mWhere3 As String
        If my_branch > 9 Then
            mWhere3 = " SUBSTRING(CAST(NoteSerial1 AS VARCHAR(50))," & SystemOptions.NoOFDigitUserTrans + 3 & ", 2) = " & my_branch
        
        Else
            mWhere3 = " SUBSTRING(CAST(NoteSerial1 AS VARCHAR(50)), " & SystemOptions.NoOFDigitUserTrans + 3 & ", 1) = " & my_branch
        End If
        mWhere3 = mWhere3 & " and   branch_no= " & my_branch
 
        '  Notes_codingByUser = Note_codingNew(my_branch, date1, 0, 200)
        '    Exit Function

    End If
    Dim mWhere4 As String
    
    If SystemOptions.IsSerialByUserVouch Then
       
        mWhere4 = "SUBSTRING(CAST(cast(NoteSerial AS BIGINT) AS VARCHAR(100)),2 , " & SystemOptions.NoOFDigitUserVouc & ") = " & user_id
        mWhere4 = mWhere4 & " AND " & user_id & "  IN ("
        mWhere4 = mWhere4 & " SELECT UserID FROM DOUBLE_ENTREY_VOUCHERS AS dev WHERE dev.Notes_ID = Notes.NoteID)"

    Else
        mUserIdSerial = ""
        ' mWhere4 = mWhere4 & " and UserID = " & user_id
    End If

    '******************

    sql = "select ISNULL(numbering_id, 0), ISNULL(start_at, 0) start_at,ISNULL(end_at, 0) end_at from sanad_numbering where branch_no=" & my_branch & " and  sanad_no=0"
        
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly ' , adCmdText
  
    If rs.EOF = 0 Then
        numbering_type = 0
    Else
        numbering_type = rs!numbering_id 'IIf(IsNull(rs("numbering_id").value), 0, rs("numbering_id").value)
        start_at = rs!start_at 'IIf(IsNull(rs("start_at").value), 0, rs("start_at").value)
        end_at = rs!end_at 'IIf(IsNull(rs("end_at").value), 0, rs("end_at").value)

    End If

    If numbering_type = 1 Then
        If SystemOptions.JLCodeBasedOnBranch = False Then
            sql = "select max(cast(NoteSerial AS BIGINT)) as last_sand_no from  Notes WHERE NoteType<>1"  'where      numbering_type=" & numbering_type
        Else
            sql = "select max(cast(NoteSerial AS BIGINT)) as last_sand_no from  Notes where    NoteType<>1 AND  branch_no= " & my_branch '& "  and     numbering_type=" & numbering_type
            sql = sql & " and " & mWhere3
        End If
        sql = sql & "  and   NoteType <>1 "
        sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        Set Rs3 = New ADODB.Recordset
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly  ', adCmdText

        If Not IsNull(Rs3("last_sand_no").value) Then
            If end_at = 0 Then end_at = val(Rs3("last_sand_no").value) + 1: GoTo XL1
               
            If Rs3("last_sand_no").value >= end_at Then
                Notes_codingByUser = "error"
                Exit Function
            End If
        End If
XL1:
    ElseIf numbering_type = 2 Then '„ ’· ”‰‘Â—ÌÊÌ
 
        If SystemOptions.JLCodeBasedOnBranch = False Then
            sql = "select max(cast(NoteSerial AS BIGINT)) as last_sand_no from  Notes where    year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            If SystemOptions.JLCodeBasedOnBranch = False Then
                '201910 0001
                sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS bigint) AS varchar(100)), 5 + " & SystemOptions.NoOFDigitUserVouc + 1 & " , 2) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 1, 2))"
                sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS bigint) AS varchar(100)), " & SystemOptions.NoOFDigitUserVouc + 2 & ", 4) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 7, 4))"

            End If
        Else
            sql = "select max(cast(NoteSerial AS BIGINT)) as last_sand_no from  Notes where  branch_no= " & my_branch & " and sanad_year=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and sanad_month=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            ' sql = sql & " and " & mWhere3
        End If
        sql = sql & "  and   NoteType <>1 "
        sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        sql = sql & " and " & IIf(mWhere3 = "", " 1 = 1 ", mWhere3)
        Set Rs3 = New ADODB.Recordset
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly  ' , adCmdText

        If Not IsNull(Rs3("last_sand_no").value) Then
            NO = mId(Rs3("last_sand_no").value, 7 + SystemOptions.NoOFDigitUserVouc + 1, Len(Rs3("last_sand_no").value) - 6)

            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Notes_codingByUser = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ

        If JLCodeBasedOnBranch = False Then
            sql = "select max(cast(NoteSerial AS BIGINT)) as last_sand_no from  Notes where     year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            If SystemOptions.JLCodeBasedOnBranch = False Then
                sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS bigint) AS varchar(100)), " & SystemOptions.NoOFDigitUserVouc + 1 & ", 4) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 7, 4))"
            End If
        Else
            sql = "select max(cast(NoteSerial AS BIGINT)) as last_sand_no from  Notes where    branch_no= " & my_branch & "  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            sql = sql & " and " & mWhere3
        End If
  
        sql = sql & "  and   NoteType <>1 "
        sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        Set Rs3 = New ADODB.Recordset
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly ' , adCmdText
 
        If Not IsNull(Rs3("last_sand_no").value) Then
            NO = mId(Rs3("last_sand_no").value, 5 + SystemOptions.NoOFDigitUserVouc + 1, Len(Rs3("last_sand_no").value) - 4)

            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Notes_codingByUser = "error"
                Exit Function
            End If
        End If
 
    End If

    Dim Askcount As Integer
    Askcount = SystemOptions.Ked_digit ' GetSetting(StrAppRegPath, "Setting", "Count_Ked_digit", 0)
         
    Dim first_serial As Boolean
           
    Dim mNum         As Integer
    Dim mNum2        As Integer
    If SystemOptions.JLCodeBasedOnBranch = True Then
        mNum = 9
        mNum2 = 7
    Else
        mNum = 7
        mNum2 = 5
    End If

    If Rs3.RecordCount = 0 Or IsNull(Rs3("last_sand_no").value) Then
        first_serial = True

        If numbering_type = 0 Then
                 
        ElseIf numbering_type = 1 Then
            auto_sanad_no = start_at
        ElseIf numbering_type = 2 Then
            '  auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            auto_sanad_no = year(date1) & Format(Month(date1), String(2, "0")) & Format(start_at, String(Askcount, "0"))
            
            ' year(date1) & Format(Month(date1), String(2, "0"))
        ElseIf numbering_type = 3 Then
            '    auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(start_at, String(Askcount, "0"))
            auto_sanad_no = year(date1) & Format(start_at, String(Askcount, "0"))
        
        End If

    Else

        If numbering_type = 0 Then
                
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Rs3("last_sand_no").value + 1
        ElseIf numbering_type = 2 Then
              
            NO = mId(Rs3("last_sand_no").value, mNum + SystemOptions.NoOFDigitUserVouc + 1, Len(Rs3("last_sand_no").value) - 6)
            'auto_sanad_no = Mid(rs3("last_sand_no").value, 1, 6) & Format((NO + 1), String(Askcount, "0"))
            auto_sanad_no = year(date1) & Format(Month(date1), String(2, "0")) & Format((NO + 1), String(Askcount, "0"))
        ElseIf numbering_type = 3 Then
         
            NO = mId(Rs3("last_sand_no").value, mNum2 + SystemOptions.NoOFDigitUserVouc + 1, Len(Rs3("last_sand_no").value) - 4)
            'auto_sanad_no = Mid(rs3("last_sand_no").value, 1, 4) & Format((NO + 1), String(Askcount, "0"))
            auto_sanad_no = year(date1) & Format((NO + 1), String(Askcount, "0"))
        End If
 
    End If
    Rs3.Close
    brancHcode = zeropadding(CStr(my_branch), Int(SystemOptions.BranchDigit))
    If SystemOptions.JLCodeBasedOnBranch = True Then
        Notes_codingByUser = "1" & mUserIdSerial & brancHcode & auto_sanad_no
    Else
        Notes_codingByUser = "1" & mUserIdSerial & auto_sanad_no
    End If
    
    'If first_serial = False Then
    'auto_sanad_no = Mid(auto_sanad_no, 2, Len(auto_sanad_no))
    'End If
    'Notes_coding = my_branch & auto_sanad_no
    'Notes_codingByUser = auto_sanad_no
  
End Function
  
Public Function Notes_coding(my_branch As Integer, _
                             date1 As Date, _
                             Optional departement_name As Integer = 1) As String
    On Error Resume Next
    Dim start_at       As Double
    Dim end_at         As Single
    Dim auto_sanad_no  As String
    Dim NO             As Single
    Dim numbering_type As Integer
    auto_sanad_no = ""

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset

    Dim sql                 As String
    Dim i                   As Integer
    Dim JLCodeBasedOnBranch As Boolean
    JLCodeBasedOnBranch = SystemOptions.JLCodeBasedOnBranch
    '**********************
    'JLCodeBasedOnBranch = True
    If SystemOptions.IsSerialByUserVouch Then

        Notes_coding = Notes_codingByUser(my_branch, date1, departement_name)
        Exit Function
    End If
    If JLCodeBasedOnBranch = True Then
        Notes_coding = Note_codingNew(my_branch, date1, 0, 200)
        Exit Function

    End If
    '******************
    Dim mWhere3 As String
    If my_branch > 9 Then
        mWhere3 = " SUBSTRING(CAST(CAST(NoteSerial AS BigInt) AS VARCHAR(50)), 1, 2) = " & my_branch
    Else
        mWhere3 = " SUBSTRING(CAST(CAST(NoteSerial AS BigInt) AS VARCHAR(50)), 1, 1) = " & my_branch
    End If

    sql = "select Isnull(numbering_id,0) numbering_id "
    sql = sql & " , isnull(start_at,0) start_at , "
    sql = sql & "  isnull( end_at,0) end_at from sanad_numbering where branch_no=" & my_branch & " and  sanad_no=0"
        
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly
  
    If rs.EOF Then
        numbering_type = 0
    Else
        numbering_type = rs!numbering_id 'IIf(IsNull(rs("numbering_id").value), 0, rs("numbering_id").value)
        start_at = rs!start_at 'IIf(IsNull(rs("start_at").value), 0, rs("start_at").value)
        end_at = rs!end_at 'IIf(IsNull(rs("end_at").value), 0, rs("end_at").value)
    End If

    If numbering_type = 1 Then
        If JLCodeBasedOnBranch = False Then
            sql = "select max(NoteSerial) as last_sand_no from  Notes WHERE NoteType<>1"  'where      numbering_type=" & numbering_type
        Else
            sql = "select max(NoteSerial) as last_sand_no from  Notes where    NoteType<>1 AND  branch_no= " & my_branch '& "  and     numbering_type=" & numbering_type
            sql = sql & " and " & mWhere3
        End If
        sql = sql & "  and   NoteType <>1 "
        Set Rs3 = New ADODB.Recordset
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly ', adCmdText

        If Not IsNull(Rs3("last_sand_no").value) Then
        
            If end_at = 0 Then end_at = val(Rs3("last_sand_no").value) + 1: GoTo XL1
               
            If Rs3("last_sand_no").value >= end_at Then
                Notes_coding = "error"
                Exit Function
            End If
        End If
XL1:
    ElseIf numbering_type = 2 Then '„ ’· ”‰‘Â—ÌÊÌ
 
        If JLCodeBasedOnBranch = False Then
            sql = "select max(NoteSerial) as last_sand_no from  Notes where    year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            If SystemOptions.JLCodeBasedOnBranch = False Then
                '2019010002
                sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS bigint) AS varchar(100)), 5, 2) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 1, 2))"
                sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS bigint) AS varchar(100)), 1, 4) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 7, 4))"
            End If
        Else
            sql = "select max(NoteSerial) as last_sand_no from  Notes where  branch_no= " & my_branch & " and sanad_year=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and sanad_month=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        sql = sql & "  and   NoteType <>1 "
        Set Rs3 = New ADODB.Recordset
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly  ', adCmdText

        If Not IsNull(Rs3("last_sand_no").value) Then
            NO = mId(Rs3("last_sand_no").value, 7, Len(Rs3("last_sand_no").value) - 6)

            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Notes_coding = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ

        If JLCodeBasedOnBranch = False Then
            sql = "select max(NoteSerial) as last_sand_no from  Notes where     year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            If SystemOptions.JLCodeBasedOnBranch = False Then
                sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS bigint) AS varchar(100)), 1, 4) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 7, 4))"
            End If
        Else
            sql = "select max(NoteSerial) as last_sand_no from  Notes where    branch_no= " & my_branch & "  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            sql = sql & " and " & mWhere3
        End If
  
        sql = sql & "  and   NoteType <>1 "
      Set Rs3 = New ADODB.Recordset
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly ' , adCmdText
 
        If Not IsNull(Rs3("last_sand_no").value) Then
            NO = mId(Rs3("last_sand_no").value, 5, Len(Rs3("last_sand_no").value) - 4)

            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Notes_coding = "error"
                Exit Function
            End If
        End If
 
    End If

    Dim Askcount As Integer
    Askcount = SystemOptions.Ked_digit ' GetSetting(StrAppRegPath, "Setting", "Count_Ked_digit", 0)
         
    Dim first_serial As Boolean

    If Rs3.EOF Or IsNull(Rs3("last_sand_no").value) Then
        first_serial = True

        If numbering_type = 0 Then
                 
        ElseIf numbering_type = 1 Then
            auto_sanad_no = start_at
        ElseIf numbering_type = 2 Then
            '  auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            auto_sanad_no = year(date1) & Format(Month(date1), String(2, "0")) & Format(start_at, String(Askcount, "0"))
            
            ' year(date1) & Format(Month(date1), String(2, "0"))
        ElseIf numbering_type = 3 Then
            '    auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(start_at, String(Askcount, "0"))
            auto_sanad_no = year(date1) & Format(start_at, String(Askcount, "0"))
        
        End If

    Else

        If numbering_type = 0 Then
                
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Rs3("last_sand_no").value + 1
        ElseIf numbering_type = 2 Then
              
            NO = mId(Rs3("last_sand_no").value, 7, Len(Rs3("last_sand_no").value) - 6)
            'auto_sanad_no = Mid(rs3("last_sand_no").value, 1, 6) & Format((NO + 1), String(Askcount, "0"))
            auto_sanad_no = year(date1) & Format(Month(date1), String(2, "0")) & Format((NO + 1), String(Askcount, "0"))
        ElseIf numbering_type = 3 Then
         
            NO = mId(Rs3("last_sand_no").value, 5, Len(Rs3("last_sand_no").value) - 4)
            'auto_sanad_no = Mid(rs3("last_sand_no").value, 1, 4) & Format((NO + 1), String(Askcount, "0"))
            auto_sanad_no = year(date1) & Format((NO + 1), String(Askcount, "0"))
        End If
 
    End If

    Rs3.Close
 
    Notes_coding = auto_sanad_no
  
End Function

Public Function Note_codingNew(my_branch As Integer, _
                               date1 As Date, _
                               Sanad_No As Integer, _
                               NoteType As Integer, _
                               Optional departement_name As Integer = 1, _
                               Optional Transaction_Type As Integer = 0, _
                               Optional Prefix As String = "", _
                               Optional StoreID As Integer = 0) As String
    
    On Error Resume Next
    Dim start_at       As Integer
    Dim end_at         As Integer
    Dim auto_sanad_no  As String
    Dim NO             As Integer
    Dim numbering_type As Integer
    Dim noOfDigit      As Double
    Dim Zeros          As Double
    Dim StoreCoding    As Double
    Dim YearDigit      As Double
    Dim branchpadidng  As Integer
    Dim storepadding   As Integer

    auto_sanad_no = ""
 
    Dim first_serial As Boolean
    Dim rs           As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset

    Dim sql       As String
    Dim i         As Integer
    Dim storecode As String
    Dim mWhere3   As String
    If my_branch > 9 Then
        mWhere3 = " SUBSTRING(CAST(CAST(NoteSerial AS BigInt) AS VARCHAR(50)), 1, 2) = " & my_branch
    Else
        mWhere3 = " SUBSTRING(CAST(CAST(NoteSerial AS BigInt) AS VARCHAR(50)), 1, 1) = " & my_branch
    End If
    first_serial = False
    '  sql = "select * from sanad_numbering where branch_no=" & my_branch & " and  sanad_no=" & Sanad_No
    sql = "SELECT ISNULL(numbering_id, 0) numbering_id, "
    sql = sql & "       ISNULL(start_at, 0) start_at, "
    sql = sql & "       ISNULL(end_at, 0) end_at, "
    sql = sql & "       ISNULL(no_of_digit, 0) no_of_digit, "
    sql = sql & "       ISNULL(zeros, 0) zeros, "
    sql = sql & "       ISNULL(YearDigit, 0) YearDigit "
    sql = sql & "FROM sanad_numbering "
    sql = sql & "WHERE branch_no = " & my_branch & " "
    sql = sql & "      AND sanad_no = " & Sanad_No & ";"
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly ' , adCmdText
  
    If rs.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = IIf(IsNull(rs("numbering_id").value), 0, rs("numbering_id").value)
        start_at = IIf(IsNull(rs("start_at").value), 0, rs("start_at").value)
        end_at = IIf(IsNull(rs("end_at").value), 0, rs("end_at").value)
        noOfDigit = IIf(IsNull(rs("no_of_digit").value), 0, rs("no_of_digit").value)
        Zeros = IIf(IsNull(rs("zeros").value), 0, rs("zeros").value)
        '   StoreCoding = IIf(IsNull(rs("StoreCoding").value), 0, rs("StoreCoding").value)
        YearDigit = IIf(IsNull(rs("YearDigit").value), 4, rs("YearDigit").value)
        If noOfDigit = 0 Then noOfDigit = 3
        '  storepadding = SystemOptions.StoreDigit - 1
        If YearDigit = 0 Then YearDigit = 4
        
        branchpadidng = SystemOptions.BranchDigit - 1
        
    End If
    If val(my_branch) < 10 Then
        branchpadidng = 1
    ElseIf val(my_branch) >= 10 And Len(my_branch) <= 99 Then
        branchpadidng = 2
    Else
        branchpadidng = 3
    End If
    If numbering_type = 1 Then ' «·Ì
        sql = "select max(NoteSerial) as last_sand_no from  Notes    where  NoteType<>1 AND   branch_no= " & my_branch
        sql = sql & "  and   NoteType <>1 "
        sql = sql & " and " & mWhere3
        '   Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rs3("last_sand_no").value) Then
        
            startrreadding = branchpadidng + 1
            noofreadinchar = startrreadding - 1
 
            NO = mId(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
           
            If end_at = 0 Then end_at = val(Rs3("last_sand_no").value) + 1: GoTo xl
               
            If Rs3("last_sand_no").value >= end_at Then
                Note_codingNew = "error"
                Exit Function
            End If
        End If

xl:
             
    ElseIf numbering_type = 2 Then ' „ ’· ‘Â—Ì

        sql = "select max(NoteSerial) as last_sand_no from  Notes where     branch_no= " & my_branch & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
   
        sql = sql & "  and   NoteType <>1 "
        sql = sql & " and " & mWhere3
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
        If Not IsNull(Rs3("last_sand_no").value) Then
            startrreadding = branchpadidng + YearDigit + 3
            noofreadinchar = startrreadding - 1
  
            NO = mId(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
             
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Note_codingNew = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ
 
        sql = "select max(NoteSerial) as last_sand_no from  Notes where  branch_no= " & my_branch & "and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        sql = sql & "  and   NoteType <>1 "
        sql = sql & " and " & mWhere3
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If Not IsNull(Rs3("last_sand_no").value) Then
     
            startrreadding = branchpadidng + YearDigit + 1
            noofreadinchar = startrreadding - 1
 
            NO = mId(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Note_codingNew = "error"
                Exit Function
            End If
        End If
 
    End If

    Dim Askcount As Double
    'Askcount = 3
    Askcount = noOfDigit

    If Askcount = 0 Then Askcount = 3

    If Rs3.RecordCount = 0 Or IsNull(Rs3("last_sand_no").value) Then
        first_serial = True

        If numbering_type = 0 Then
            'ÌœÊÌ
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Format(start_at, String(Askcount, "0"))
   
        ElseIf numbering_type = 2 Then
        
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            End If
        
        ElseIf numbering_type = 3 Then
        
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format(start_at, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(start_at, String(Askcount, "0"))
            End If
       
        End If

    Else

        If numbering_type = 0 Then
                
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Format((NO + 1), String(Askcount, "0"))
        ElseIf numbering_type = 2 Then
            If StoreCoding = True And StoreID <> 0 Then
              
                If YearDigit = 2 Then
                    '            no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                    '  no = Mid(Rs3("last_sand_no").value, 10, Len(Rs3("last_sand_no").value) - 9)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                End If
             
            Else
             
                If YearDigit = 2 Then
                    ' no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                    '  no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                End If
             
            End If
        ElseIf numbering_type = 3 Then
            If StoreCoding = True And StoreID <> 0 Then
            
                If YearDigit = 2 Then
                    '  no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                               
                    '          no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format((NO + 1), String(Askcount, "0"))
                    
                End If
                      
            Else
              
                If YearDigit = 2 Then
                    'no = Mid(Rs3("last_sand_no").value, 4, Len(Rs3("last_sand_no").value) - 3)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                               
                    '            no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format((NO + 1), String(Askcount, "0"))
                    
                End If
              
            End If
                      
        End If

    End If

    Rs3.Close
    'Dim storeADDZero As String
    'storeADDZero = IIf(StoreID < 10, "0", "")
    Dim brancHcode As String
 
    brancHcode = zeropadding(CStr(my_branch), Int(SystemOptions.BranchDigit))
    'storecode = zeropadding(storecode, Int(SystemOptions.StoreDigit))

    '    If numbering_type = 1 Then Note_codingNew = auto_sanad_no: Exit Function
    
    If first_serial = True Then
        If auto_sanad_no <> "" Then
            
            If StoreCoding = True And StoreID <> 0 Then
                '       Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
                Note_codingNew = brancHcode & storecode & auto_sanad_no
            Else
                Note_codingNew = brancHcode & auto_sanad_no
            End If
        
        Else
            Note_codingNew = auto_sanad_no
        End If

    Else
        '     Voucher_coding = my_branch & auto_sanad_no
        If StoreCoding = True And StoreID <> 0 Then
            ' Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
            Note_codingNew = brancHcode & storecode & auto_sanad_no
        Else
            Note_codingNew = brancHcode & auto_sanad_no
        End If
    End If

End Function
 
'
'

Private Function GetSalesCounterStartValue(ByVal start_at As Double) As Long
    If start_at <= 0 Then
        GetSalesCounterStartValue = 0
    Else
        GetSalesCounterStartValue = CLng(start_at) - 1
    End If
End Function

Private Function GetNextSalesCounterValue(ByVal Transaction_Type As Integer, _
                                          ByVal BranchID As Integer, _
                                          ByVal start_at As Double, _
                                          ByRef NextCounterValue As Long) As Boolean
    Dim CnCounter  As ADODB.Connection
    Dim RsCounter  As ADODB.Recordset
    Dim sqlCounter As String
    Dim seedValue  As Long

    On Error GoTo ErrHandler

    seedValue = GetSalesCounterStartValue(start_at)

    Set CnCounter = New ADODB.Connection
    CnCounter.ConnectionString = Cn.ConnectionString
    CnCounter.Open
    CnCounter.BeginTrans

    sqlCounter = "SELECT CounterValue "
    sqlCounter = sqlCounter & "FROM SerialCounters WITH (UPDLOCK, HOLDLOCK) "
    sqlCounter = sqlCounter & "WHERE TransactionType = " & Transaction_Type
    sqlCounter = sqlCounter & " AND BranchID = " & BranchID

    Set RsCounter = New ADODB.Recordset
    RsCounter.Open sqlCounter, CnCounter, adOpenKeyset, adLockOptimistic

    If RsCounter.EOF Then
        RsCounter.Close
        sqlCounter = "INSERT INTO SerialCounters (TransactionType, BranchID, CounterValue, LastUpdated) "
        sqlCounter = sqlCounter & "VALUES (" & Transaction_Type & ", " & BranchID & ", " & seedValue & ", GETDATE())"
        CnCounter.Execute sqlCounter
    Else
        RsCounter.Close
    End If

    sqlCounter = "UPDATE SerialCounters "
    sqlCounter = sqlCounter & "SET CounterValue = CounterValue + 1, LastUpdated = GETDATE() "
    sqlCounter = sqlCounter & "WHERE TransactionType = " & Transaction_Type
    sqlCounter = sqlCounter & " AND BranchID = " & BranchID
    CnCounter.Execute sqlCounter

    sqlCounter = "SELECT CounterValue "
    sqlCounter = sqlCounter & "FROM SerialCounters WITH (HOLDLOCK) "
    sqlCounter = sqlCounter & "WHERE TransactionType = " & Transaction_Type
    sqlCounter = sqlCounter & " AND BranchID = " & BranchID

    Set RsCounter = New ADODB.Recordset
    RsCounter.Open sqlCounter, CnCounter, adOpenForwardOnly, adLockReadOnly

    If RsCounter.EOF Then GoTo ErrHandler

    NextCounterValue = CLng(RsCounter!CounterValue)
    RsCounter.Close
    CnCounter.CommitTrans

    GetNextSalesCounterValue = True

ExitHandler:
    On Error Resume Next
    If Not RsCounter Is Nothing Then
        If RsCounter.State = adStateOpen Then RsCounter.Close
    End If
    If Not CnCounter Is Nothing Then
        If CnCounter.State = adStateOpen Then CnCounter.Close
    End If
    Set RsCounter = Nothing
    Set CnCounter = Nothing
    Exit Function

ErrHandler:
    GetNextSalesCounterValue = False
    On Error Resume Next
    If Not CnCounter Is Nothing Then
        If CnCounter.State = adStateOpen Then CnCounter.RollbackTrans
    End If
    Resume ExitHandler
End Function

Public Function Voucher_coding(my_branch As Integer, _
                               date1 As Date, _
                               Sanad_No As Integer, _
                               NoteType As Integer, _
                               Optional departement_name As Integer = 1, _
                               Optional Transaction_Type As Integer = 0, _
                               Optional Prefix As String = "", _
                               Optional StoreID As Integer = 0, _
                               Optional BillType As Integer = 0, _
                               Optional MosemID As Double = 0, _
                               Optional ByVal mTableName As String = "", _
                               Optional ByVal mUserId As Long = 0, _
                               Optional ByRef mSerInv As Long = 0, _
                               Optional ByVal mSerPosString As String = "") As String
                               
                               
    
    On Error Resume Next
    If my_branch = 0 Then
        Exit Function
    End If
    If mUserId = 0 Then mUserId = user_id
    'SystemOptions.IsByNewCoding = False

    If SystemOptions.IsByNewCoding Then
       ' Voucher_coding = Voucher_codingByBreaks(my_branch, date1, Sanad_No, NoteType, departement_name, Transaction_Type, Prefix, StoreID, BillType, MosemID, mTableName, mUserID, mSerInv)
       ' Exit Function
    End If
    If SystemOptions.IsSerialByUserTrans Then
        Voucher_coding = Voucher_codingByUser(my_branch, date1, Sanad_No, NoteType, departement_name, Transaction_Type, Prefix, StoreID, BillType, MosemID, mTableName, mUserId)
        Exit Function
    End If

    Dim start_at       As Double
    Dim end_at         As Double
    Dim auto_sanad_no  As String
    Dim NO             As Double
    Dim numbering_type As Integer
    Dim noOfDigit      As Double
    Dim Zeros          As Double
    Dim StoreCoding    As Double
    Dim YearDigit      As Double
    Dim branchpadidng  As Integer
    Dim storepadding   As Integer

    auto_sanad_no = ""
 
    Dim first_serial As Boolean
    Dim rs           As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset

    Dim sql       As String
    Dim i         As Integer
    Dim storecode As String
    Dim Askcount          As Double
    Dim mSalesCounterValue As Long

    first_serial = False
    ' sql = "select * from sanad_numbering where branch_no=" & my_branch & " and  sanad_no=" & Sanad_No
   
    sql = "SELECT  isnull(Prefix,0)  Prefix,isnull(numbering_id,0)  numbering_id ,  ISNULL(start_at, 0) start_at, "
    sql = sql & "       ISNULL(end_at, 0) end_at, "
    sql = sql & "       ISNULL(no_of_digit, 0) no_of_digit, "
    sql = sql & "       ISNULL(zeros, 0) zeros, "
    sql = sql & "       ISNULL(StoreCoding, 0) StoreCoding, "
    sql = sql & "       ISNULL(YearDigit, 0) YearDigit "
    sql = sql & "FROM sanad_numbering "
    sql = sql & "WHERE branch_no =  " & my_branch
    sql = sql & "      AND sanad_no = " & Sanad_No & ""
   If Prefix <> "" Then
        sql = sql & "      AND Prefix  = '" & Prefix & "';"
   End If
    
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly
  
    If Not rs.EOF = 0 Then
        numbering_type = 0
    Else
        numbering_type = rs!numbering_id 'IIf(IsNull(rs("numbering_id").value), 0, rs("numbering_id").value)
        start_at = rs!start_at 'IIf(IsNull(rs("start_at").value), 0, rs("start_at").value)
        end_at = rs!end_at 'IIf(IsNull(rs("end_at").value), 0, rs("end_at").value)
        noOfDigit = rs!no_of_digit 'IIf(IsNull(rs("no_of_digit").value), 0, rs("no_of_digit").value)
        If noOfDigit = 0 Then noOfDigit = 3
        Zeros = rs!Zeros 'IIf(IsNull(rs("zeros").value), 0, rs("zeros").value)
        StoreCoding = rs!StoreCoding 'IIf(IsNull(rs("StoreCoding").value), 0, rs("StoreCoding").value)
        YearDigit = rs!YearDigit 'IIf(IsNull(rs("YearDigit").value), 4, rs("YearDigit").value)
        
        storepadding = SystemOptions.StoreDigit - 1
        If YearDigit = 0 Then YearDigit = 4
        
        branchpadidng = SystemOptions.BranchDigit - 1

        If StoreCoding = True Then
            If StoreID <> 0 Then
                storecode = getStoreCoding(StoreID)
            End If
        End If
        
    End If
    
    
     Dim brancHcode As String
 
    brancHcode = zeropadding(CStr(my_branch), Int(SystemOptions.BranchDigit))
    
    Dim mWhere3 As String
    If Sanad_No = 75 Then '”‰œ œ›⁄«  «·⁄ﬁ«— ›« Ê—Â «·÷—Ì»Â
        
        
        If my_branch > 9 < my_branch < 100 Then
                
            mWhere3 = " SUBSTRING(CAST(CAST(TblContractInstallments.NoteSerial1 AS BigInt) AS VARCHAR(50)), 1, 2) = " & my_branch
        ElseIf my_branch > 99 Then
            mWhere3 = " SUBSTRING(CAST(CAST(TblContractInstallments.NoteSerial1 AS BigInt) AS VARCHAR(50)), 1, 3) = " & my_branch
        Else
            mWhere3 = " SUBSTRING(CAST(CAST(TblContractInstallments.NoteSerial1 AS BigInt) AS VARCHAR(50)), 1, 1) = " & my_branch
        End If
    Else
        
        If start_at = 0 Then
            If my_branch > 9 < my_branch < 100 Then
        
                mWhere3 = " SUBSTRING(CAST(CAST(NoteSerial1 AS BigInt) AS VARCHAR(50)), 1, 2) = " & my_branch
            ElseIf my_branch > 99 Then
                mWhere3 = " SUBSTRING(CAST(CAST(NoteSerial1 AS BigInt) AS VARCHAR(50)), 1, 3) = " & my_branch
            
            Else
                mWhere3 = " SUBSTRING(CAST(CAST(NoteSerial1 AS BigInt) AS VARCHAR(50)), 1, 1) = " & my_branch
            End If
        Else
            mWhere3 = " 1 = 1 "
        End If
    End If

    ' mWhere3 = " 1 = 1 "
    If Transaction_Type = 21 And numbering_type <> 0 Then
        Askcount = noOfDigit
        If Askcount = 0 Then Askcount = 3

        If GetNextSalesCounterValue(Transaction_Type, my_branch, start_at, mSalesCounterValue) = False Then
            Voucher_coding = "error"
            Exit Function
        End If

        mSerInv = mSalesCounterValue

        If end_at <> 0 And mSalesCounterValue > end_at Then
            Voucher_coding = "error"
            Exit Function
        End If

        If numbering_type = 1 Then
            auto_sanad_no = CStr(mSalesCounterValue)
        ElseIf numbering_type = 2 Then
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(mSalesCounterValue, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(mSalesCounterValue, String(Askcount, "0"))
            End If
        ElseIf numbering_type = 3 Then
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format(mSalesCounterValue, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(mSalesCounterValue, String(Askcount, "0"))
            End If
        End If

        GoTo BuildSalesVoucherCode
    End If
    If numbering_type = 1 Then ' «·Ì
        If Transaction_Type = 0 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and   NoteType=" & NoteType ' & " and   numbering_type1=" & numbering_type
            Select Case Sanad_No

                Case 5
                    sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and not(BTCashAccountcode is null )"
                Case 1
                    If SystemOptions.ExpensesCoding Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type
               
                    ElseIf SystemOptions.ExpensesCoding2 = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 14) " ' and   numbering_type1=" & numbering_type
                    End If
                Case 4
                    If SystemOptions.ExpensesCoding = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
                    Else
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
                    End If
                    ' ﬂÊÌœ ”‰œ «· À”Ìÿ ‰›” ”‰œ «·ﬁ»÷
                Case 25
                    If SystemOptions.InstallmntsvchrCoding = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 18) " '  and   numbering_type1=" & numbering_type
                    End If
                Case 2
                    If SystemOptions.InstallmntsvchrCoding = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 18 )" ' And numbering_type1 = " & numbering_type"
                    Else
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType  ' & " and   numbering_type1=" & numbering_type
                    End If
                Case 26
                    If SystemOptions.InstallmntsvchrCoding = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
                    End If
                Case 2
                    If SystemOptions.InstallmntsvchrCoding = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
                    Else
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type
                    End If
                    '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ «·’—›
   
                Case Is = 1
                    If SystemOptions.ExpensesCoding2 = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 14) " ' and   numbering_type1=" & numbering_type
                    End If
                Case 16
                    If SystemOptions.ExpensesCoding2 = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 3 or NoteType= 14 ) " ' and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
                    Else
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType  '& " and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
                    End If
                Case 50
           
                    sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
                Case 58
           
                    sql = "select max(NoteSerial1) as last_sand_no from  TblExchange    where  BranchID= " & my_branch
   
                Case 60
           
                    sql = "select max(NoteSerial1) as last_sand_no from  TblContract    where  Branch_NO= " & my_branch
                Case 62
            
                    sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= " & NoteType & ") "
             
                Case 64
                    sql = "select max(CAST (NoteSerial1 AS BigInt))as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   "
                Case 65, 84
                    If SystemOptions.AllowProjectBill2Serial = True Then
                        If BillType = 1 Then
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1 "
                        Else
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0 "
                        End If
                
                    Else
            
                        sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "   "
                    End If
            
                Case 83
                    If SystemOptions.AllowProjectBill2Serial = True Then
                        If BillType = 1 Then
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "  and bill_to=1 "
                        Else
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract where  Branch_NO= " & my_branch & "  and bill_to=0 "
                        End If
                
                    Else
            
                        sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract where  Branch_NO= " & my_branch & "   "
                    End If
 
                Case 66
                    sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  "
                    sql = sql & "  and ( Transaction_Type=990 or Transaction_Type=18)"
                Case 67
                    sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  "
                    sql = sql & "  and  (Transaction_Type=66 or Transaction_Type=991) "
                Case 68
                    sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  "
                    sql = sql & "  and  (ImportExport=0 ) "
                Case 69
                    sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  "
                    sql = sql & "  and  (ImportExport=1 ) "
  
                Case 70
                    sql = "select max (NoteSerial1 )  as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  "
                    sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
   
                Case 71
                    sql = "select max (NoteSerial1 )  as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  "
                    sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
       
                Case 72
                    sql = "select max (NoteSerial1 )  as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  "
                    sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
                Case 74
                    sql = "select max (NoteSerial1 )  as last_sand_no from  notes_all where  branch_no= " & my_branch & "  and  notetype=370 "
                    sql = sql & " and " & mWhere3
                Case 75
                    sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS BigInt)) as last_sand_no "
                    sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
                    sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
                    sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")"
                Case 76
                    sql = "select max (NoteSerial1 )  as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "  "
                    sql = sql & " and " & mWhere3
 
            End Select
        Else
       
            sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
            sql = sql & " and " & mWhere3
        End If
        '    If Transaction_Type <> 0 Then
        '        sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
        '        sql = sql & " and " & mWhere3
        '    End If
        
        If Prefix = "" Then
            sql = sql & "  and   Prefix is null"
 
        Else
            sql = sql & "  and   Prefix='" & Prefix & "'"
        End If
        sql = sql & " and " & mWhere3
  
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly  'adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rs3("last_sand_no").value) Then
        
            If end_at = 0 Then end_at = val(Rs3("last_sand_no").value) + 1
               
            If Rs3("last_sand_no").value >= end_at Then
                Voucher_coding = "error"
                Exit Function
            End If
        End If
    ElseIf numbering_type = 2 Then ' „ ’· ‘Â—Ì

        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        sql = sql & " and " & mWhere3
        If Trim(mTableName & "") = "" Then
            If Transaction_Type = 0 Then
                Select Case Sanad_No

                    Case Is = 5
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and not(BTCashAccountcode is null )"
                        sql = sql & " and " & mWhere3
                    Case 1
                        If SystemOptions.ExpensesCoding = True Then
                            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            sql = sql & " and " & mWhere3
                        End If
                    Case 4
                        If SystemOptions.ExpensesCoding = True Then
                            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
                            sql = sql & " and " & mWhere3
                        Else
                            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
                            sql = sql & " and " & mWhere3
                        End If
        
                    Case 25
                        If SystemOptions.InstallmntsvchrCoding = True Then
                            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            sql = sql & " and " & mWhere3
                        End If
       
                    Case 2
                        If SystemOptions.InstallmntsvchrCoding = True Then
                            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
                            sql = sql & " and " & mWhere3
                        Else
                            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "       and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
                            sql = sql & " and " & mWhere3
                        End If
                        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ«  «·’—›
                    Case 1
                        If SystemOptions.ExpensesCoding2 = True Then
                            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            sql = sql & " and " & mWhere3
                        End If
        
                    Case 16
                        If SystemOptions.ExpensesCoding2 = True Then
                            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)     and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
                            sql = sql & " and " & mWhere3
                        Else
                            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
                            sql = sql & " and " & mWhere3
                        End If
                    Case 50
                        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  BranchID= " & my_branch & " and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
                    Case 58
           
                        sql = "select max(NoteSerial1) as last_sand_no from  TblExchange where  BranchID= " & my_branch & " and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
 
                    Case 60
         
                        sql = "select max(NoteSerial1) as last_sand_no from  TblContract where  Branch_NO= " & my_branch & " and year(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
                        sql = sql & " and " & mWhere3
            
                    Case 62
           
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= " & NoteType & ")   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
           
                    Case 64
                        sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Case 66
                        sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR  Transaction_Type=18)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Case 67
                        sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Case 68
                        sql = "select max  max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=0)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Case 69
                        sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=1)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Case 70
                        sql = "select max (NoteSerial1 ) as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
         
                    Case 71
                        sql = "select max (NoteSerial1 ) as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Case 72
                        sql = "select max (NoteSerial1 ) as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Case 74
                        sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "   and  notetype=370   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Case 76
                        sql = "select max (NoteSerial1 ) as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "      and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Case 85
                        sql = " SELECT     max(dbo.TblContractInstallments.NoteSerial1H )as last_sand_no  "
                        sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
                        sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
                        sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")and year(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        Case 86
                        sql = "select max (InvoiceID ) as last_sand_no from  tblEInvoice where  year(IssueDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(IssueDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        'where  BranchId= " & my_branch & "
                        sql = sql & " and " & mWhere3
                    Case 65, 84
                        If SystemOptions.AllowProjectBill2Serial = True Then
                            If BillType = 1 Then
                                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            Else
                                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            End If
                        Else
            
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        End If
                    Case 83
                        If SystemOptions.AllowProjectBill2Serial = True Then
                            If BillType = 1 Then
                                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            Else
                                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            End If
                        Else
            
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        End If
        
                    Case 81
                        sql = "select max (NoteSerial1 ) as last_sand_no from  TblHandWages where  BranchId= " & my_branch & "      and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        If mSerPosString <> "" Then
                             sql = sql & " and   IsNull(TblHandWages.SerPos,0) = " & val(mSerPosString)
                         Else
                             sql = sql & " and   IsNull(TblHandWages.SerPos,0) = 0"
                             
                         End If
                        sql = sql & " and " & mWhere3
         
                End Select
 
                sql = sql & " and " & mWhere3
            Else 'Transaction_Type <> 0
        
                If StoreCoding = True Then
                    If Transaction_Type = 10 Then
                        sql = "select max(  (NoteSerial1    )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & "   Or Transaction_Type= 992)  and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Else
                        sql = "select max(  (NoteSerial1    )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    End If
                Else
                    If SystemOptions.BranchDigit > 1 Then
                 
                        '   sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        'edit edit salim here
                        If Transaction_Type = 10 Then
                            sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & "  Or Transaction_Type= 992)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            sql = sql & " and " & mWhere3
                        Else
                            sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            sql = sql & " and " & mWhere3
                        End If
                        'edit edit salim here
            
                    Else
                        If Transaction_Type = 10 Then
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and  ( Transaction_Type=" & Transaction_Type & "   Or Transaction_Type= 992)  and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            sql = sql & " and " & mWhere3
                        Else
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            If mSerPosString <> "" Then
                                sql = sql & " and   IsNull(Transactions.SerPos,0) = " & val(mSerPosString)
                            Else
                                sql = sql & " and   IsNull(Transactions.SerPos,0) = 0"
                                sql = sql & " and " & mWhere3
                            End If
                            
                        End If
                    End If
       
                End If
                If Prefix = "" Then
                    sql = sql & "  and   Prefix is null"
 
                Else
                    sql = sql & "  and   Prefix='" & Prefix & "'"
                End If
  
                If StoreCoding = True And StoreID <> 0 Then
                    sql = sql & "  and   StoreID=" & StoreID
                End If
  
            End If
    
            If Prefix = "" Then
                If Sanad_No <> 58 And Sanad_No <> 50 And Sanad_No <> 60 Then
                    sql = sql & "  and   Prefix is null"
                End If
 
            Else
                sql = sql & "  and   Prefix='" & Prefix & "'"
            End If
        Else
            'If mTableName <> "" Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  " & mTableName & "  where  BranchId= " & my_branch & "      and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
             If mSerPosString <> "" Or mTableName = "TblHandWages" Then
                sql = sql & " and   IsNull(" & mTableName & ".SerPos,0) = " & val(mSerPosString)
            End If

        End If
        
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly ', adCmdText

        Dim startrreadding As Integer
        Dim noofreadinchar As Integer
        If val(Rs3("last_sand_no").value & "") = 0 Then
            first_serial = True
        End If
        If Not IsNull(Rs3("last_sand_no").value) Then
            If StoreCoding = True And StoreID <> 0 Then
                startrreadding = SystemOptions.BranchDigit + SystemOptions.StoreDigit + YearDigit + noOfDigit
                noofreadinchar = startrreadding - 1
            Else
                startrreadding = SystemOptions.BranchDigit + YearDigit + noOfDigit '+ Len(start_at)
                If Transaction_Type = 0 And (Sanad_No <> 66 And Sanad_No <> 67) Then
                    startrreadding = 1 + YearDigit + noOfDigit
                End If
                noofreadinchar = startrreadding - 1
            End If
            If (Len(Rs3("last_sand_no").value) > 9 And noOfDigit < 5 And YearDigit < 4) Or (Len(Rs3("last_sand_no").value) > 11 And noOfDigit < 5 And YearDigit = 4) Then
                If Len(brancHcode) = 1 Then
                    NO = mId(Rs3("last_sand_no").value, startrreadding - 1, Len(Rs3("last_sand_no").value))
                Else
                    NO = mId(Rs3("last_sand_no").value, startrreadding - 0, Len(Rs3("last_sand_no").value))
                End If
                
                If mSerPosString <> "" Then
                    If Len(brancHcode & mSerPosString) = 1 Then
                        NO = mId(Rs3("last_sand_no").value, startrreadding + 1, Len(Rs3("last_sand_no").value) - noofreadinchar)
                    Else
                        NO = mId(Rs3("last_sand_no").value, startrreadding + Len(brancHcode & mSerPosString) - 1, Len(Rs3("last_sand_no").value) - noofreadinchar)
                    End If

                End If
            Else
                If Len(brancHcode) = 1 Then
                    NO = mId(Rs3("last_sand_no").value, startrreadding + 1, Len(Rs3("last_sand_no").value) - noofreadinchar)
                Else
                    NO = mId(Rs3("last_sand_no").value, startrreadding + 0, Len(Rs3("last_sand_no").value) - noofreadinchar)
                End If
            End If
            If Len(CStr(NO)) > noOfDigit Then noOfDigit = noOfDigit + 1
            If Len(CStr(NO)) > noOfDigit Or NO = 0 Then
                
                NO = GetNumberAfterMonth(Rs3("last_sand_no") & "", YearDigit, Len(brancHcode), Len(Month(date1)))
            End If
            If Len(CStr(NO)) > noOfDigit Then noOfDigit = Len(CStr(NO))
            
            
           If start_at = 1 Then
                NO = right(Rs3("last_sand_no").value, noOfDigit)
            
           '„—«Ã⁄… ÷—Ê—Ï Ãœ« Ãœ« Ãœ« Ê«∆· Ê”«„Ì
            Else
              NO = mId(Rs3("last_sand_no").value, startrreadding + 1, Len(Rs3("last_sand_no").value) - noofreadinchar)
          '„—«Ã⁄… ÷—Ê—Ï Ãœ« Ãœ« Ãœ« Ê«∆· Ê”«„Ì
            End If
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Voucher_coding = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ
        If Transaction_Type = 0 Then
            Select Case Sanad_No
                Case 64
                    sql = "select max(CAST (NoteSerial1 AS BigInt))as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Case 65, 84
                    If SystemOptions.AllowProjectBill2Serial = True Then
                        If BillType = 1 Then
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                        Else
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                        End If
                    Else
                        sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                    End If
        
                Case 83
                    If SystemOptions.AllowProjectBill2Serial = True Then
                        If BillType = 1 Then
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                        Else
                            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                        End If
                    Else
            
                        sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                    End If
                Case 5
                    sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and not(BTCashAccountcode is null )"
                Case 1
                    If SystemOptions.ExpensesCoding = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)     and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
                    End If
                Case 4
                    If SystemOptions.ExpensesCoding = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
                    Else
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
                    End If
        
                Case 25
                    If SystemOptions.InstallmntsvchrCoding = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)      and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
                    End If
        
                Case 2
                    If SystemOptions.InstallmntsvchrCoding = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                    Else
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                    End If
         
                    '”‰œ«  «· ÕÊÌ· ‰›”  —ﬁ”„ ”‰œ«  «·’—›
                Case 1
                    If SystemOptions.ExpensesCoding2 = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
                    End If
        
                Case 16
                    If SystemOptions.ExpensesCoding2 = True Then
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)        and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
                    Else
                        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
                    End If
        
                Case 50
      
                    sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  BranchID= " & my_branch & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
      
                Case 58
      
                    '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
                    'sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
                    sql = "select max(NoteSerial1) as last_sand_no from  TblExchange where  BranchID= " & my_branch & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
                Case 60
                    sql = "select max(NoteSerial1) as last_sand_no from  TblContract where  Branch_NO= " & my_branch & " and year(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
     
                Case 62
           
                    sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= " & NoteType & ")    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Case 66
                    sql = "select  max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR Transaction_Type=18)  and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Case 67
                    sql = "select  max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and  (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
     
                Case 68
                    sql = "select  max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and  (ImportExport=0)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Case 69
                    sql = "select  max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and  (ImportExport=1)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Case 70
                    sql = "select  max( (NoteSerial1)  as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Case 71
                    sql = "select  max( (NoteSerial1)  as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Case 72
                    sql = "select  max( (NoteSerial1)  as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Case 74
                    sql = "select  max( (NoteSerial1)  as last_sand_no from  notes_all where  branch_no= " & my_branch & " and  notetype=370    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
Case 57
                    sql = "select  max (NoteSerial1)  as last_sand_no from  notes where  branch_no= " & my_branch & " and  notetype=57  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Case 76
                    sql = "select max (NoteSerial1 ) as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "     and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Case 75
                    sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS BigInt)) as last_sand_no  "
                    sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
                    sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
                    sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")and year(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " "
      
            End Select
            ' sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            'TblContract  Branch_NO
        Else
            '  If Transaction_Type <> 0 Then
            If StoreCoding = True Then
                sql = "select  max(  (NoteSerial1  ))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                        
            Else
                sql = "select  max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            End If
            
            If StoreCoding = True And StoreID <> 0 Then
                sql = sql & "  and   StoreID=" & StoreID
            End If
        End If
        
        If Prefix = "" Then
            If Sanad_No = 58 Or Sanad_No = 60 Then
            Else
                sql = sql & "  and   Prefix is null"
            End If
        Else
            sql = sql & "  and   Prefix='" & Prefix & "'"
        End If
  
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly ', adCmdText
 
        If Not IsNull(Rs3("last_sand_no").value) Then
            If StoreCoding = True And StoreID <> 0 Then
                                           
                startrreadding = SystemOptions.BranchDigit + SystemOptions.StoreDigit + YearDigit + 1
                noofreadinchar = startrreadding - 1
         
            Else
                If val(getNoOfBranches) > 9 Then
         
                    If mId(Rs3("last_sand_no").value, 1, 1) = "0" Then
         
                        startrreadding = SystemOptions.BranchDigit + YearDigit + 1
                    Else
             
                        If val(my_branch) > 9 Then
                            startrreadding = SystemOptions.BranchDigit + YearDigit + 1
                        Else
                            startrreadding = SystemOptions.BranchDigit + YearDigit
                        End If
             
                    End If
             
                Else
                    startrreadding = SystemOptions.BranchDigit + YearDigit
                    'noofreadinchar = startrreadding
                    If Transaction_Type <> 0 Then
                        If SystemOptions.BranchDigit = 1 Then
                            startrreadding = startrreadding + 1
                        End If
                    Else
                        startrreadding = startrreadding + 1
                    End If
                End If
      
                '                If Transaction_Type = 0 Then
                '                    'startrreadding = 1 + YearDigit + 1
                '                End If
                   
                '                               If YearDigit = 2 Then
                '           no = Mid(Rs3("last_sand_no").value, 4, Len(Rs3("last_sand_no").value) - 3)
                '        Else
                '        no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                '        End If
         
            End If
            noofreadinchar = startrreadding - 1
            NO = mId(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Voucher_coding = "error"
                Exit Function
            End If
        End If
 
    End If

    'Askcount = 3
    Askcount = noOfDigit

    If Askcount = 0 Then Askcount = 3

    If Rs3.EOF Or IsNull(Rs3("last_sand_no").value) Then
        first_serial = True

        If numbering_type = 0 Then
                 
        ElseIf numbering_type = 1 Then
            auto_sanad_no = start_at
        ElseIf numbering_type = 2 Then
        
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            End If
        
        ElseIf numbering_type = 3 Then
        
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format(start_at, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(start_at, String(Askcount, "0"))
            End If
       
        End If

    Else

        If numbering_type = 0 Then
                
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Rs3("last_sand_no").value + 1
        ElseIf numbering_type = 2 Then
            If StoreCoding = True And StoreID <> 0 Then
              
                If YearDigit = 2 Then
                    '            no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                    '  no = Mid(Rs3("last_sand_no").value, 10, Len(Rs3("last_sand_no").value) - 9)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                End If
             
            Else
             
                If YearDigit = 2 Then
                    ' no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                    '  no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                End If
             
            End If
        ElseIf numbering_type = 3 Then
            If StoreCoding = True And StoreID <> 0 Then
            
                If YearDigit = 2 Then
                    '  no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                               
                    '          no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format((NO + 1), String(Askcount, "0"))
                    
                End If
                      
            Else
              
                If YearDigit = 2 Then
                    'no = Mid(Rs3("last_sand_no").value, 4, Len(Rs3("last_sand_no").value) - 3)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                               
                    '            no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format((NO + 1), String(Askcount, "0"))
                    
                End If
              
            End If
                      
        End If

    End If

BuildSalesVoucherCode:
    If Not Rs3 Is Nothing Then
        If Rs3.State = adStateOpen Then Rs3.Close
    End If
    'Dim storeADDZero As String
    'storeADDZero = IIf(StoreID < 10, "0", "")
    
 
    brancHcode = zeropadding(CStr(my_branch), Int(SystemOptions.BranchDigit))
    storecode = zeropadding(storecode, Int(SystemOptions.StoreDigit))

    If numbering_type = 1 Then
        Voucher_coding = auto_sanad_no
        Exit Function
    End If
    
    If first_serial = True Then
        If auto_sanad_no <> "" Then
            
            If StoreCoding = True And StoreID <> 0 Then
                '       Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
                Voucher_coding = brancHcode & storecode & auto_sanad_no
            Else
                If mSerPosString <> "" Then
                    Voucher_coding = mSerPosString & brancHcode & auto_sanad_no
                Else
                    Voucher_coding = brancHcode & auto_sanad_no
                End If
            End If
        
        Else
            Voucher_coding = auto_sanad_no
        End If

    Else
        '     Voucher_coding = my_branch & auto_sanad_no
        If StoreCoding = True And StoreID <> 0 Then
            ' Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
            Voucher_coding = brancHcode & storecode & auto_sanad_no
        Else
            If mSerPosString <> "" Then
                Voucher_coding = mSerPosString & brancHcode & auto_sanad_no
            Else
                Voucher_coding = brancHcode & auto_sanad_no
            End If
        End If
    End If

End Function

Public Function Voucher_codingByUser(my_branch As Integer, _
                                     date1 As Date, _
                                     Sanad_No As Integer, _
                                     NoteType As Integer, _
                                     Optional departement_name As Integer = 1, _
                                     Optional Transaction_Type As Integer = 0, _
                                     Optional Prefix As String = "", _
                                     Optional StoreID As Integer = 0, _
                                     Optional BillType As Integer = 0, _
                                     Optional MosemID As Double = 0, _
                                     Optional ByVal mTableName As String = "", _
                                     Optional ByVal mUserId As Long = 0) As String
    
    On Error Resume Next
    If my_branch = 0 Then
        Exit Function
    End If
    Dim start_at       As Double
    Dim end_at         As Double
    Dim auto_sanad_no  As String
    Dim NO             As Double
    Dim numbering_type As Integer
    Dim noOfDigit      As Double
    Dim Zeros          As Double
    Dim StoreCoding    As Double
    Dim s              As String
    Dim YearDigit      As Double
    Dim branchpadidng  As Integer
    Dim storepadding   As Integer
    If mUserId = 0 Then mUserId = user_id
    Dim mNoOfUser     As Integer
    Dim mUserIdSerial As String
    Dim mLenUser      As Integer

    Dim mFormatUser   As String

    mFormatUser = ""
    Dim mm As Integer
    mm = 1
    For mm = 1 To SystemOptions.NoOFDigitUserTrans
        If mUserId < 10 And SystemOptions.NoOFDigitUserTrans > 1 And SystemOptions.NoOFDigitUserTrans > mm Then
            mFormatUser = mFormatUser & "0"
        ElseIf mUserId > 9 And mUserId < 100 And SystemOptions.NoOFDigitUserTrans > 1 And SystemOptions.NoOFDigitUserTrans > mm + 1 Then
            mFormatUser = mFormatUser & "0"
        ElseIf mUserId > 99 And SystemOptions.NoOFDigitUserTrans > 1 And SystemOptions.NoOFDigitUserTrans > mm + 2 Then
            mFormatUser = mFormatUser & "0"
        End If

    Next

    'If mUserId > 9 And mUserId < 100 Then
    mUserIdSerial = mFormatUser & mUserId
    'End If

    auto_sanad_no = ""
 
    Dim first_serial As Boolean
    Dim rs           As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset

    Dim sql       As String
    Dim i         As Integer
    Dim storecode As String
    Dim Askcount          As Double
    Dim mSalesCounterValue As Long

    first_serial = False
   sql = "SELECT ISNULL(numbering_id, 0) numbering_id, "
sql = sql & "       ISNULL(start_at, 0) start_at, "
sql = sql & "       ISNULL(end_at, 0) end_at, "
sql = sql & "       ISNULL(no_of_digit, 0) no_of_digit, "
sql = sql & "       ISNULL(zeros, 0) zeros, "
sql = sql & "       ISNULL(StoreCoding, 0) StoreCoding, "
sql = sql & "       ISNULL(YearDigit, 0) YearDigit, "
sql = sql & "       ISNULL(IsBreaks, 0) IsBreaks, "
sql = sql & "       ISNULL(IsCodeByBranch, 0) IsCodeByBranch, "
sql = sql & "       ISNULL(IsSerialByUser, 0) IsSerialByUser, "
sql = sql & "       ISNULL(Breaks, 0) Breaks "
  sql = sql & "   from sanad_numbering where branch_no=" & my_branch & " and  sanad_no=" & Sanad_No
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        numbering_type = 0
                
    Else
        numbering_type = rs!numbering_id 'IIf(IsNull(rs("numbering_id").value), 0, rs("numbering_id").value)
        start_at = rs!start_at 'IIf(IsNull(rs("start_at").value), 0, rs("start_at").value)
        end_at = rs!end_at 'IIf(IsNull(rs("end_at").value), 0, rs("end_at").value)
        noOfDigit = rs!no_of_digit 'IIf(IsNull(rs("no_of_digit").value), 0, rs("no_of_digit").value)
        If noOfDigit = 0 Then noOfDigit = 3
        Zeros = rs!Zeros 'IIf(IsNull(rs("zeros").value), 0, rs("zeros").value)
        StoreCoding = rs!StoreCoding 'IIf(IsNull(rs("StoreCoding").value), 0, rs("StoreCoding").value)
        YearDigit = rs!YearDigit 'IIf(IsNull(rs("YearDigit").value), 4, rs("YearDigit").value)
        
        storepadding = SystemOptions.StoreDigit - 1
        If YearDigit = 0 Then YearDigit = 4
        
        branchpadidng = SystemOptions.BranchDigit - 1

        If StoreCoding = True Then
            If StoreID <> 0 Then
                storecode = getStoreCoding(StoreID)
            End If
        End If
        
    End If
    
    Dim mWhere4     As String
    Dim mWhereUser  As String
    Dim mWhereUser2 As String
    
    If SystemOptions.IsSerialByUserTrans Then
       
        mWhere4 = "SUBSTRING(CAST(cast(NoteSerial1 AS BIGINT) AS VARCHAR(100)),2 , " & SystemOptions.NoOFDigitUserTrans & ") = " & mUserId
        mWhereUser = mWhereUser & " AND " & mUserId & "  IN ("
        mWhereUser = mWhereUser & " SELECT UserID FROM DOUBLE_ENTREY_VOUCHERS AS dev WHERE dev.Notes_ID = Notes.NoteID)"
        mWhereUser2 = " And UserID = " & mUserId

    Else
        mUserIdSerial = ""
        ' mWhere4 = mWhere4 & " and UserID = " & mUserId
    End If
    
    Dim mWhere3 As String
    If my_branch > 9 Then
        mWhere3 = " SUBSTRING(CAST(cast(NoteSerial1 AS BIGINT) AS VARCHAR(50))," & SystemOptions.NoOFDigitUserTrans + 2 & ", 2) = " & my_branch
    Else
        mWhere3 = " SUBSTRING(CAST(cast(NoteSerial1 AS BIGINT) AS VARCHAR(50)), " & SystemOptions.NoOFDigitUserTrans + 2 & ", 1) = " & my_branch
    End If
    ' mWhere3 = " 1 = 1 "
    If Transaction_Type = 21 And numbering_type <> 0 Then
        Askcount = noOfDigit
        If Askcount = 0 Then Askcount = 3

        If GetNextSalesCounterValue(Transaction_Type, my_branch, start_at, mSalesCounterValue) = False Then
            Voucher_codingByUser = "error"
            Exit Function
        End If

        If end_at <> 0 And mSalesCounterValue > end_at Then
            Voucher_codingByUser = "error"
            Exit Function
        End If

        If numbering_type = 1 Then
            auto_sanad_no = CStr(mSalesCounterValue)
        ElseIf numbering_type = 2 Then
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(mSalesCounterValue, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(mSalesCounterValue, String(Askcount, "0"))
            End If
        ElseIf numbering_type = 3 Then
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format(mSalesCounterValue, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(mSalesCounterValue, String(Askcount, "0"))
            End If
        End If

        GoTo BuildSalesVoucherCodeByUser
    End If
    If numbering_type = 1 Then ' «·Ì
        sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and   NoteType=" & NoteType ' & " and   numbering_type1=" & numbering_type
        sql = sql & mWhereUser
        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and not(BTCashAccountcode is null )"
            sql = sql & mWhereUser
            
        End If
   
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
        End If
   
        If Sanad_No = 4 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            End If
            sql = sql & mWhereUser
        End If
   
        ' ﬂÊÌœ ”‰œ «· À”Ìÿ ‰›” ”‰œ «·ﬁ»÷
        If Sanad_No = 25 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 18) " '  and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
            
        End If
   
        If Sanad_No = 2 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 18 )" ' And numbering_type1 = " & numbering_type"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType  ' & " and   numbering_type1=" & numbering_type
            End If
            sql = sql & mWhereUser
        End If
   
        If Sanad_No = 26 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
        End If
   
        If Sanad_No = 2 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type
            End If
            sql = sql & mWhereUser
        End If

        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ «·’—›
   
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 14) " ' and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
        End If
   
        If Sanad_No = 16 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 3 or NoteType= 14 ) " ' and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType  '& " and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            End If
            sql = sql & mWhereUser
        End If
      
        If Sanad_No = 50 Then
           
            sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            sql = sql & mWhereUser2
        End If
        
        If Sanad_No = 58 Then
           
            sql = "select max(NoteSerial1) as last_sand_no from  TblExchange    where  BranchID= " & my_branch
            sql = sql & mWhereUser2
        End If
        
        If Sanad_No = 60 Then
           
            sql = "select max(NoteSerial1) as last_sand_no from  TblContract    where  Branch_NO= " & my_branch
            sql = sql & mWhereUser2
        End If
        
        If Sanad_No = 62 Then
            
            sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= " & NoteType & ") "
            sql = sql & mWhereUser
        End If
        
        If Sanad_No = 64 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt))as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   "
            sql = sql & mWhereUser2
        End If
        
        If Sanad_No = 65 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1 "
                Else
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0 "
                End If
            Else
            
                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "   "
            End If
        End If
          
        If Sanad_No = 83 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "  and bill_to=1 "
                Else
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "  and bill_to=0 "
                End If
            Else
            
                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "   "
            End If
        End If
            
        If Sanad_No = 66 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  "
            sql = sql & "  and ( Transaction_Type=990 or Transaction_Type=18)"
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 67 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  "
            sql = sql & "  and  (Transaction_Type=66 or Transaction_Type=991) "
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 68 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  "
            sql = sql & "  and  (ImportExport=0 ) "
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 69 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  "
            sql = sql & "  and  (ImportExport=1 ) "
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 70 Then
            sql = "select max (NoteSerial1 )  as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  "
            sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 71 Then
            sql = "select max (NoteSerial1 )  as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  "
            sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
            sql = sql & mWhereUser2
        End If
        
        If Sanad_No = 72 Then
            sql = "select max (NoteSerial1 )  as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  "
            sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 74 Then
            sql = "select max (NoteSerial1 )  as last_sand_no from  notes_all where  branch_no= " & my_branch & "  and  notetype=370 "
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 75 Then
            '  Sql = "select max(CAST (dbo.TblContractInstallments.NoteSerial1 AS BigInt))   as last_sand_no from  TblContractInstallments where  branch_no= " & my_branch & "  "
            sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS BigInt)) as last_sand_no "
            sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
            sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
            sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")"
        End If
        If Sanad_No = 76 Then
            sql = "select max (NoteSerial1 )  as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "  "
            sql = sql & " and " & mWhere3
        End If
        If Transaction_Type <> 0 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
        If Transaction_Type <> 0 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
        
        If Prefix = "" Then
            sql = sql & "  and   Prefix is null"
 
        Else
            sql = sql & "  and   Prefix='" & Prefix & "'"
        End If
        sql = sql & " and " & mWhere3
        sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
  
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly  ', adCmdText

        If Not IsNull(Rs3("last_sand_no").value) Then
        
            If end_at = 0 Then end_at = val(Rs3("last_sand_no").value) + 1
               
            If Rs3("last_sand_no").value >= end_at Then
                Voucher_codingByUser = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 2 Then ' „ ’· ‘Â—Ì

        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        sql = sql & " and " & mWhere3
        sql = sql & mWhereUser
        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and not(BTCashAccountcode is null )"
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser
        End If
    
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        End If
    
        If Sanad_No = 4 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        
        End If
    
        If Sanad_No = 25 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        End If
    
        If Sanad_No = 2 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "       and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        End If
    
        If Sanad_No = 16 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)     and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        
        End If
        
        If Sanad_No = 50 Then
      
            '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  BranchID= " & my_branch & " and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
        End If
        If Sanad_No = 58 Then
      
            '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            sql = "select max(NoteSerial1) as last_sand_no from  TblExchange where  BranchID= " & my_branch & " and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
        
        End If
        
        If Sanad_No = 60 Then
         
            sql = "select max(NoteSerial1) as last_sand_no from  TblContract where  Branch_NO= " & my_branch & " and year(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
            sql = sql & " and " & mWhere3
        
        End If
        
        'TblContract  Branch_NO
        'Dim stockSettelmentsstr As String
        'stockSettelmentsstr = ""
        
        If Sanad_No = 62 Then
           
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= " & NoteType & ")   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser
        End If
        ''////
        If Sanad_No = 64 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 66 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR  Transaction_Type=18)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 67 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 68 Then
            sql = "select max  max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=0)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
          
        End If
        If Sanad_No = 69 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=1)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 70 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 71 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 72 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 74 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "   and  notetype=370   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
         
        If Sanad_No = 76 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "      and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 75 Then
            sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS BigInt))as last_sand_no  "
            sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
            sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
            sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")and year(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                
            'Sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "     and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
                
        If Sanad_No = 65 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                Else
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    'sql = sql & " and " & mWhere3
                End If
            Else
            
                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                'sql = sql & " and " & mWhere3
            End If
        End If
               
        If Sanad_No = 83 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                Else
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    'sql = sql & " and " & mWhere3
                End If
            Else
            
                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                'sql = sql & " and " & mWhere3
            End If
        End If
            
        sql = sql & " and " & mWhere3
        ''//////
        If Transaction_Type <> 0 Then
            '   If Transaction_Type = 15 Or Transaction_Type = 16 Then
            '    stockSettelmentsstr
            '   End If
        
            If StoreCoding = True Then
                'Or Transaction_Type= 992)
                If Transaction_Type = 10 Then
                    sql = "select max(  (NoteSerial1    )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & "   Or Transaction_Type= 992)  and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    sql = sql & " and " & mWhere3
                     
                Else
                    sql = "select max(  (NoteSerial1    )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    sql = sql & " and " & mWhere3
                End If
                sql = sql & mWhereUser2
            Else
                If SystemOptions.BranchDigit > 1 Then
                 
                    '   sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    'edit edit salim here
                    If Transaction_Type = 10 Then
                        sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & "  Or Transaction_Type= 992)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Else
                        sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    End If
                    sql = sql & mWhereUser2
                    'edit edit salim here
            
                Else
                    If Transaction_Type = 10 Then
                        sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and  ( Transaction_Type=" & Transaction_Type & "   Or Transaction_Type= 992)  and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    Else
                        sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    End If
                    sql = sql & mWhereUser2
                End If
       
            End If
            
            '
            If Prefix = "" Then
                sql = sql & "  and   Prefix is null"
 
            Else
                sql = sql & "  and   Prefix='" & Prefix & "'"
            End If
  
            If StoreCoding = True And StoreID <> 0 Then
                sql = sql & "  and   StoreID=" & StoreID
            End If
  
        End If
    
        If Prefix = "" Then
            If Sanad_No <> 58 And Sanad_No <> 50 And Sanad_No <> 60 Then
                sql = sql & "  and   Prefix is null"
            End If
 
        Else
            sql = sql & "  and   Prefix='" & Prefix & "'"
        End If
        
        If mTableName <> "" Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  " & mTableName & "  where  BranchId= " & my_branch & "      and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3

        End If
        
        s = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        
        ' sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        
        Rs3.Open s, Cn, adOpenForwardOnly, adLockReadOnly  ' , adCmdText
        If Not Rs3.EOF Then
            If val(Rs3!last_sand_no & "") = 0 Then
                s = sql
            End If
        Else
            s = sql
        End If
       Set Rs3 = New ADODB.Recordset
        Rs3.Open s, Cn, adOpenForwardOnly, adLockReadOnly ', adCmdText
        Dim startrreadding As Integer
        Dim noofreadinchar As Integer
        If Not IsNull(Rs3("last_sand_no").value) Then
            If StoreCoding = True And StoreID <> 0 Then
                startrreadding = SystemOptions.BranchDigit + SystemOptions.StoreDigit + YearDigit + noOfDigit
                noofreadinchar = startrreadding - 1
                '    If YearDigit = 2 Then
                     
                '      no = Mid(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
                '    Else
                     
                '    no = Mid(Rs3("last_sand_no").value, 10, Len(Rs3("last_sand_no").value) - 9)
                '   End If
               
            Else
           
                startrreadding = SystemOptions.BranchDigit + YearDigit + noOfDigit
                If Transaction_Type = 0 And (Sanad_No <> 66 And Sanad_No <> 67) Then
                    startrreadding = 1 + YearDigit + noOfDigit
                End If
           
                noofreadinchar = startrreadding - 1
                   
                '                      If YearDigit = 2 Then
                '             no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                '           Else
                '
                '           no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                '          End If
           
            End If
            NO = mId(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
            NO = right(Rs3("last_sand_no").value, noofreadinchar - 1)
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Voucher_codingByUser = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ
        If Sanad_No = 64 Then
            sql = "select max(CAST (NoteSerial1 AS BigInt))as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        
        If Sanad_No = 65 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Else
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                End If
            Else
            
                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & " and " & mWhere3
            End If
        End If
        
        If Sanad_No = 83 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract  where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Else
                    sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                End If
            Else
            
                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  SubcontractorContract where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & " and " & mWhere3
            End If
        End If
 
        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        sql = sql & mWhereUser
        sql = sql & " and " & mWhere3
        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and not(BTCashAccountcode is null )"
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
      
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)     and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & mWhereUser
                sql = sql & " and " & mWhere3
            End If
        
        End If
     
        If Sanad_No = 4 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            End If
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
          
        If Sanad_No = 25 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)      and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & mWhereUser
                sql = sql & " and " & mWhere3
            End If
        
        End If
     
        If Sanad_No = 2 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            End If
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›”  —ﬁ”„ ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & mWhereUser
                sql = sql & " and " & mWhere3
            End If
        
        End If
     
        If Sanad_No = 16 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)        and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            End If
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
    
        If Sanad_No = 50 Then
      
            '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            'sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
            sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  BranchID= " & my_branch & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 58 Then
      
            '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            'sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
            sql = "select max(NoteSerial1) as last_sand_no from  TblExchange where  BranchID= " & my_branch & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            sql = sql & " and " & mWhere3
        End If
        
        If Sanad_No = 60 Then
            sql = "select max(NoteSerial1) as last_sand_no from  TblContract where  Branch_NO= " & my_branch & " and year(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
   
        End If
        
        If Sanad_No = 62 Then
           
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= " & NoteType & ")    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 66 Then
            sql = "select  max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR Transaction_Type=18)  and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            sql = sql & mWhereUser2
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 67 Then
            sql = "select  max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and  (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            sql = sql & mWhereUser2
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 68 Then
            sql = "select  max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and  (ImportExport=0)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 69 Then
            sql = "select  max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and  (ImportExport=1)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 70 Then
            sql = "select  max( (NoteSerial1)  as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 71 Then
            sql = "select  max( (NoteSerial1)  as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 72 Then
            sql = "select  max( (NoteSerial1)  as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 74 Then
            sql = "select  max( (NoteSerial1)  as last_sand_no from  notes_all where  branch_no= " & my_branch & " and  notetype=370    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            sql = sql & mWhereUser2
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 76 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "     and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 75 Then
            sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS BigInt)) as last_sand_no  "
            sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
            sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
            sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")and year(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " "
                
            'Sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "     and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        
        'TblContract  Branch_NO
        If Transaction_Type <> 0 Then
            If StoreCoding = True Then
                sql = "select  max(  (NoteSerial1  ))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                        
            Else
                sql = "select  max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            End If
            
            If StoreCoding = True And StoreID <> 0 Then
                sql = sql & "  and   StoreID=" & StoreID
            End If
            sql = sql & mWhereUser2
            sql = sql & " and " & mWhere3
        End If
        
        If Prefix = "" Then
            If Sanad_No = 58 Or Sanad_No = 60 Then
            Else
                sql = sql & "  and   Prefix is null"
            End If
        Else
            sql = sql & "  and   Prefix='" & Prefix & "'"
        End If
  
        sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        Set Rs3 = New ADODB.Recordset
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly  ', adCmdText
 
        If Not IsNull(Rs3("last_sand_no").value) Then
            If StoreCoding = True And StoreID <> 0 Then
                                           
                startrreadding = SystemOptions.BranchDigit + SystemOptions.StoreDigit + YearDigit + 1
                noofreadinchar = startrreadding - 1
                'If YearDigit = 2 Then
                '            no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                '         Else
                '         no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                '         End If
         
            Else
                If val(getNoOfBranches) > 9 Then
         
                    If mId(Rs3("last_sand_no").value, 1, 1) = "0" Then
         
                        startrreadding = SystemOptions.BranchDigit + YearDigit + 1
                    Else
             
                        If val(my_branch) > 9 Then
                            startrreadding = SystemOptions.BranchDigit + YearDigit + 1
                        Else
                            startrreadding = SystemOptions.BranchDigit + YearDigit
                        End If
             
                    End If
             
                Else
                    startrreadding = SystemOptions.BranchDigit + YearDigit
                    'noofreadinchar = startrreadding
                    If Transaction_Type <> 0 Then
                        If SystemOptions.BranchDigit = 1 Then
                            startrreadding = startrreadding + 1
                        End If
                    Else
                        startrreadding = startrreadding + 1
                    End If
                End If
      
                If Transaction_Type = 0 Then
                    'startrreadding = 1 + YearDigit + 1
                End If
                   
                '                               If YearDigit = 2 Then
                '           no = Mid(Rs3("last_sand_no").value, 4, Len(Rs3("last_sand_no").value) - 3)
                '        Else
                '        no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                '        End If
         
            End If
            noofreadinchar = startrreadding - 1
            NO = mId(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Voucher_codingByUser = "error"
                Exit Function
            End If
        End If
 
    End If

    'Askcount = 3
    Askcount = noOfDigit

    If Askcount = 0 Then Askcount = 3

    If Rs3.RecordCount = 0 Or IsNull(Rs3("last_sand_no").value) Then
        first_serial = True

        If numbering_type = 0 Then
                 
        ElseIf numbering_type = 1 Then
            auto_sanad_no = start_at
        ElseIf numbering_type = 2 Then
        
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            End If
        
        ElseIf numbering_type = 3 Then
        
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format(start_at, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(start_at, String(Askcount, "0"))
            End If
       
        End If

    Else

        If numbering_type = 0 Then
                
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Rs3("last_sand_no").value + 1
        ElseIf numbering_type = 2 Then
            If StoreCoding = True And StoreID <> 0 Then
              
                If YearDigit = 2 Then
                    '            no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                    '  no = Mid(Rs3("last_sand_no").value, 10, Len(Rs3("last_sand_no").value) - 9)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                End If
             
            Else
             
                If YearDigit = 2 Then
                    ' no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                    '  no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                End If
             
            End If
        ElseIf numbering_type = 3 Then
            If StoreCoding = True And StoreID <> 0 Then
            
                If YearDigit = 2 Then
                    '  no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                               
                    '          no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format((NO + 1), String(Askcount, "0"))
                    
                End If
                      
            Else
              
                If YearDigit = 2 Then
                    'no = Mid(Rs3("last_sand_no").value, 4, Len(Rs3("last_sand_no").value) - 3)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                               
                    '            no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format((NO + 1), String(Askcount, "0"))
                    
                End If
              
            End If
                      
        End If

    End If

BuildSalesVoucherCodeByUser:
    If Not Rs3 Is Nothing Then
        If Rs3.State = adStateOpen Then Rs3.Close
    End If
    'Dim storeADDZero As String
    'storeADDZero = IIf(StoreID < 10, "0", "")
    Dim brancHcode As String
 
    brancHcode = zeropadding(CStr(my_branch), Int(SystemOptions.BranchDigit))
    storecode = zeropadding(storecode, Int(SystemOptions.StoreDigit))

    If numbering_type = 1 Then
    Voucher_codingByUser = auto_sanad_no
    Exit Function
    End If
    If first_serial = True Then
        If auto_sanad_no <> "" Then
            
            If StoreCoding = True And StoreID <> 0 Then
                '       Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
                Voucher_codingByUser = "1" & mUserIdSerial & brancHcode & storecode & auto_sanad_no
            Else
                Voucher_codingByUser = "1" & mUserIdSerial & brancHcode & auto_sanad_no
            End If
        
        Else
            Voucher_codingByUser = auto_sanad_no
        End If

    Else
        '     Voucher_coding = my_branch & auto_sanad_no
        If StoreCoding = True And StoreID <> 0 Then
            ' Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
            Voucher_codingByUser = "1" & mUserIdSerial & brancHcode & storecode & auto_sanad_no
        Else
            Voucher_codingByUser = "1" & mUserIdSerial & brancHcode & auto_sanad_no
        End If
    End If

End Function

Public Function Voucher_codingoriginaloriginal(my_branch As Integer, _
                                               date1 As Date, _
                                               Sanad_No As Integer, _
                                               NoteType As Integer, _
                                               Optional departement_name As Integer = 1, _
                                               Optional Transaction_Type As Integer = 0, _
                                               Optional Prefix As String = "", _
                                               Optional StoreID As Integer = 0, _
                                               Optional BillType As Integer = 0, _
                                               Optional MosemID As Double = 0) As String
    
    On Error Resume Next
    If my_branch = 0 Then
        Exit Function
    End If
    Dim start_at       As Double
    Dim end_at         As Double
    Dim auto_sanad_no  As String
    Dim NO             As Double
    Dim numbering_type As Integer
    Dim noOfDigit      As Double
    Dim Zeros          As Double
    Dim StoreCoding    As Double
    Dim YearDigit      As Double
    Dim branchpadidng  As Integer
    Dim storepadding   As Integer

    auto_sanad_no = ""
 
    Dim first_serial As Boolean
    Dim rs           As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset

    Dim sql       As String
    Dim i         As Integer
    Dim storecode As String

    first_serial = False
    sql = "select * from sanad_numbering where branch_no=" & my_branch & " and  sanad_no=" & Sanad_No
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        numbering_type = 0
                
    Else
        numbering_type = IIf(IsNull(rs("numbering_id").value), 0, rs("numbering_id").value)
        start_at = IIf(IsNull(rs("start_at").value), 0, rs("start_at").value)
        end_at = IIf(IsNull(rs("end_at").value), 0, rs("end_at").value)
        noOfDigit = IIf(IsNull(rs("no_of_digit").value), 0, rs("no_of_digit").value)
        If noOfDigit = 0 Then noOfDigit = 3
        Zeros = IIf(IsNull(rs("zeros").value), 0, rs("zeros").value)
        StoreCoding = IIf(IsNull(rs("StoreCoding").value), 0, rs("StoreCoding").value)
        YearDigit = IIf(IsNull(rs("YearDigit").value), 4, rs("YearDigit").value)
        
        storepadding = SystemOptions.StoreDigit - 1
        If YearDigit = 0 Then YearDigit = 4
        
        branchpadidng = SystemOptions.BranchDigit - 1

        If StoreCoding = True Then
            If StoreID <> 0 Then
                storecode = getStoreCoding(StoreID)
            End If
        End If
        
    End If

    If numbering_type = 1 Then ' «·Ì
        sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and   NoteType=" & NoteType ' & " and   numbering_type1=" & numbering_type

        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and not(BTCashAccountcode is null )"
        End If
   
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type
            End If
        End If
   
        If Sanad_No = 4 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            End If
        End If
   
        ' ﬂÊÌœ ”‰œ «· À”Ìÿ ‰›” ”‰œ «·ﬁ»÷
        If Sanad_No = 25 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 18) " '  and   numbering_type1=" & numbering_type
            End If
        End If
   
        If Sanad_No = 2 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 18 )" ' And numbering_type1 = " & numbering_type"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType  ' & " and   numbering_type1=" & numbering_type
            End If
        End If
   
        If Sanad_No = 26 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
            End If
        End If
   
        If Sanad_No = 2 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type
            End If
        End If

        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ «·’—›
   
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 14) " ' and   numbering_type1=" & numbering_type
            End If
        End If
   
        If Sanad_No = 16 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 3 or NoteType= 14 ) " ' and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType  '& " and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            End If
        End If
      
        If Sanad_No = 50 Then
           
            sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            
        End If
        
        If Sanad_No = 58 Then
           
            sql = "select max(NoteSerial1) as last_sand_no from  TblExchange    where  BranchID= " & my_branch
            
        End If
        
        If Sanad_No = 60 Then
           
            sql = "select max(NoteSerial1) as last_sand_no from  TblContract    where  Branch_NO= " & my_branch
            
        End If
        
        If Sanad_No = 62 Then
            
            sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= " & NoteType & ") "
             
        End If
        
        If Sanad_No = 64 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT))as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   "
        End If
        
        If Sanad_No = 65 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1 "
                Else
                    sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0 "
                End If
            Else
            
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "   "
            End If
        End If
        If Sanad_No = 66 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  "
            sql = sql & "  and ( Transaction_Type=990 or Transaction_Type=18)"
        End If
        If Sanad_No = 67 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  "
            sql = sql & "  and  (Transaction_Type=66 or Transaction_Type=991) "
        End If
        If Sanad_No = 68 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  "
            sql = sql & "  and  (ImportExport=0 ) "
        End If
        If Sanad_No = 69 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  "
            sql = sql & "  and  (ImportExport=1 ) "
        End If
        If Sanad_No = 70 Then
            sql = "select max (NoteSerial1 )  as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  "
            sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
        End If
        If Sanad_No = 71 Then
            sql = "select max (NoteSerial1 )  as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  "
            sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
        End If
        
        If Sanad_No = 72 Then
            sql = "select max (NoteSerial1 )  as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  "
            sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
        End If
        If Sanad_No = 74 Then
            sql = "select max (NoteSerial1 )  as last_sand_no from  notes_all where  branch_no= " & my_branch & "  and  notetype=370 "
        End If
        If Sanad_No = 75 Then
            '  Sql = "select max(CAST (dbo.TblContractInstallments.NoteSerial1 AS FLOAT))   as last_sand_no from  TblContractInstallments where  branch_no= " & my_branch & "  "
            sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS FLOAT)) as last_sand_no "
            sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
            sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
            sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")"
        End If
        If Sanad_No = 76 Then
            sql = "select max (NoteSerial1 )  as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "  "
        End If
        If Transaction_Type <> 0 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
  
        End If
        If Transaction_Type <> 0 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
        End If
        
        If Prefix = "" Then
            sql = sql & "  and   Prefix is null"
 
        Else
            sql = sql & "  and   Prefix='" & Prefix & "'"
        End If
  
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rs3("last_sand_no").value) Then
        
            If end_at = 0 Then end_at = val(Rs3("last_sand_no").value) + 1
               
            If Rs3("last_sand_no").value >= end_at Then
                Voucher_codingoriginal = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 2 Then ' „ ’· ‘Â—Ì

        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)

        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and not(BTCashAccountcode is null )"
        End If
    
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        
            End If
        End If
    
        If Sanad_No = 4 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
            End If
        
        End If
    
        If Sanad_No = 25 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        
            End If
        End If
    
        If Sanad_No = 2 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "       and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
            End If
        
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        
            End If
        End If
    
        If Sanad_No = 16 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)     and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
            End If
        
        End If
        
        If Sanad_No = 50 Then
      
            '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  BranchID= " & my_branch & " and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
        End If
        If Sanad_No = 58 Then
      
            '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            sql = "select max(NoteSerial1) as last_sand_no from  TblExchange where  BranchID= " & my_branch & " and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
        
        End If
        
        If Sanad_No = 60 Then
         
            sql = "select max(NoteSerial1) as last_sand_no from  TblContract where  Branch_NO= " & my_branch & " and year(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
        End If
        
        'TblContract  Branch_NO
        'Dim stockSettelmentsstr As String
        'stockSettelmentsstr = ""
        
        If Sanad_No = 62 Then
           
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= " & NoteType & ")   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
           
        End If
        ''////
        If Sanad_No = 64 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 66 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR  Transaction_Type=18)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 67 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 68 Then
            sql = "select max  max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=0)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 69 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=1)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 70 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 71 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 72 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 74 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "   and  notetype=370   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
         
        If Sanad_No = 76 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "      and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 75 Then
            sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS FLOAT))as last_sand_no  "
            sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
            sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
            sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")and year(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                
            'Sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "     and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
                
        If Sanad_No = 65 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                Else
                    sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                End If
            Else
            
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            End If
        End If
  
        ''//////
        If Transaction_Type <> 0 Then
            '   If Transaction_Type = 15 Or Transaction_Type = 16 Then
            '    stockSettelmentsstr
            '   End If
        
            If StoreCoding = True Then
                'Or Transaction_Type= 992)
                If Transaction_Type = 10 Then
                    sql = "select max(  (NoteSerial1    )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & "   Or Transaction_Type= 992)  and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                Else
                    sql = "select max(  (NoteSerial1    )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                End If
            Else
                If SystemOptions.BranchDigit > 1 Then
                 
                    '   sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    'edit edit salim here
                    If Transaction_Type = 10 Then
                        sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & "  Or Transaction_Type= 992)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    Else
                        sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    End If
                    'edit edit salim here
            
                Else
                    If Transaction_Type = 10 Then
                        sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and  ( Transaction_Type=" & Transaction_Type & "   Or Transaction_Type= 992)  and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    Else
                        sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    End If
                End If
       
            End If
            
            '
            If Prefix = "" Then
                sql = sql & "  and   Prefix is null"
 
            Else
                sql = sql & "  and   Prefix='" & Prefix & "'"
            End If
  
            If StoreCoding = True And StoreID <> 0 Then
                sql = sql & "  and   StoreID=" & StoreID
            End If
  
        End If
    
        If Prefix = "" Then
            If Sanad_No <> 58 And Sanad_No <> 50 And Sanad_No <> 60 Then
                sql = sql & "  and   Prefix is null"
            End If
 
        Else
            sql = sql & "  and   Prefix='" & Prefix & "'"
        End If
            
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        Dim startrreadding As Integer
        Dim noofreadinchar As Integer
        If Not IsNull(Rs3("last_sand_no").value) Then
            If StoreCoding = True And StoreID <> 0 Then
                startrreadding = SystemOptions.BranchDigit + SystemOptions.StoreDigit + YearDigit + noOfDigit
                noofreadinchar = startrreadding - 1
                '    If YearDigit = 2 Then
                     
                '      no = Mid(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
                '    Else
                     
                '    no = Mid(Rs3("last_sand_no").value, 10, Len(Rs3("last_sand_no").value) - 9)
                '   End If
               
            Else
           
                startrreadding = SystemOptions.BranchDigit + YearDigit + noOfDigit
                If Transaction_Type = 0 And (Sanad_No <> 66 And Sanad_No <> 67) Then
                    startrreadding = 1 + YearDigit + noOfDigit
                End If
           
                noofreadinchar = startrreadding - 1
                   
                '                      If YearDigit = 2 Then
                '             no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                '           Else
                '
                '           no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                '          End If
           
            End If
            NO = mId(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
            NO = right(Rs3("last_sand_no").value, noOfDigit)
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Voucher_codingoriginal = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ
        If Sanad_No = 64 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT))as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        
        If Sanad_No = 65 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Else
                    sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                End If
            Else
            
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            End If
        End If
 
        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)

        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and not(BTCashAccountcode is null )"
        End If
      
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)     and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            End If
        
        End If
     
        If Sanad_No = 4 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            End If
        
        End If
          
        If Sanad_No = 25 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)      and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            End If
        
        End If
     
        If Sanad_No = 2 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            End If
        
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›”  —ﬁ”„ ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            End If
        
        End If
     
        If Sanad_No = 16 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)        and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            End If
        
        End If
    
        If Sanad_No = 50 Then
      
            '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            'sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
            sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  BranchID= " & my_branch & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
   
        End If
        If Sanad_No = 58 Then
      
            '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            'sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
            sql = "select max(NoteSerial1) as last_sand_no from  TblExchange where  BranchID= " & my_branch & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
   
        End If
        
        If Sanad_No = 60 Then
            sql = "select max(NoteSerial1) as last_sand_no from  TblContract where  Branch_NO= " & my_branch & " and year(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
   
        End If
        
        If Sanad_No = 62 Then
           
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= " & NoteType & ")    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
        End If
        If Sanad_No = 66 Then
            sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR Transaction_Type=18)  and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 67 Then
            sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and  (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 68 Then
            sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and  (ImportExport=0)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 69 Then
            sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and  (ImportExport=1)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 70 Then
            sql = "select  max( (NoteSerial1)  as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 71 Then
            sql = "select  max( (NoteSerial1)  as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 72 Then
            sql = "select  max( (NoteSerial1)  as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 74 Then
            sql = "select  max( (NoteSerial1)  as last_sand_no from  notes_all where  branch_no= " & my_branch & " and  notetype=370    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 76 Then
            sql = "select max (NoteSerial1 ) as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "     and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 75 Then
            sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS FLOAT)) as last_sand_no  "
            sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
            sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
            sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")and year(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " "
                
            'Sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "     and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        
        'TblContract  Branch_NO
        If Transaction_Type <> 0 Then
            If StoreCoding = True Then
                sql = "select  max(  (NoteSerial1  ))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                        
            Else
                sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            End If
            
            If StoreCoding = True And StoreID <> 0 Then
                sql = sql & "  and   StoreID=" & StoreID
            End If
        End If
        
        If Prefix = "" Then
            If Sanad_No = 58 Or Sanad_No = 60 Then
            Else
                sql = sql & "  and   Prefix is null"
            End If
        Else
            sql = sql & "  and   Prefix='" & Prefix & "'"
        End If
  
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If Not IsNull(Rs3("last_sand_no").value) Then
            If StoreCoding = True And StoreID <> 0 Then
                                           
                startrreadding = SystemOptions.BranchDigit + SystemOptions.StoreDigit + YearDigit + 1
                noofreadinchar = startrreadding - 1
                'If YearDigit = 2 Then
                '            no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                '         Else
                '         no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                '         End If
         
            Else
                If val(getNoOfBranches) > 9 Then
         
                    If mId(Rs3("last_sand_no").value, 1, 1) = "0" Then
         
                        startrreadding = SystemOptions.BranchDigit + YearDigit + 1
                    Else
             
                        If val(my_branch) > 9 Then
                            startrreadding = SystemOptions.BranchDigit + YearDigit + 1
                        Else
                            startrreadding = SystemOptions.BranchDigit + YearDigit
                        End If
             
                    End If
             
                Else
                    startrreadding = SystemOptions.BranchDigit + YearDigit
                    'noofreadinchar = startrreadding
                    If Transaction_Type <> 0 Then
                        If SystemOptions.BranchDigit = 1 Then
                            startrreadding = startrreadding + 1
                        End If
                    Else
                        startrreadding = startrreadding + 1
                    End If
                End If
      
                If Transaction_Type = 0 Then
                    'startrreadding = 1 + YearDigit + 1
                End If
                   
                '                               If YearDigit = 2 Then
                '           no = Mid(Rs3("last_sand_no").value, 4, Len(Rs3("last_sand_no").value) - 3)
                '        Else
                '        no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                '        End If
         
            End If
            noofreadinchar = startrreadding - 1
            NO = mId(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Voucher_codingoriginal = "error"
                Exit Function
            End If
        End If
 
    End If

    Dim Askcount As Double
    'Askcount = 3
    Askcount = noOfDigit

    If Askcount = 0 Then Askcount = 3

    If Rs3.RecordCount = 0 Or IsNull(Rs3("last_sand_no").value) Then
        first_serial = True

        If numbering_type = 0 Then
                 
        ElseIf numbering_type = 1 Then
            auto_sanad_no = start_at
        ElseIf numbering_type = 2 Then
        
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            End If
        
        ElseIf numbering_type = 3 Then
        
            If YearDigit = 2 Then
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format(start_at, String(Askcount, "0"))
            Else
                auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(start_at, String(Askcount, "0"))
            End If
       
        End If

    Else

        If numbering_type = 0 Then
                
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Rs3("last_sand_no").value + 1
        ElseIf numbering_type = 2 Then
            If StoreCoding = True And StoreID <> 0 Then
              
                If YearDigit = 2 Then
                    '            no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                    '  no = Mid(Rs3("last_sand_no").value, 10, Len(Rs3("last_sand_no").value) - 9)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                End If
             
            Else
             
                If YearDigit = 2 Then
                    ' no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                    '  no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
                End If
             
            End If
        ElseIf numbering_type = 3 Then
            If StoreCoding = True And StoreID <> 0 Then
            
                If YearDigit = 2 Then
                    '  no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                               
                    '          no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format((NO + 1), String(Askcount, "0"))
                    
                End If
                      
            Else
              
                If YearDigit = 2 Then
                    'no = Mid(Rs3("last_sand_no").value, 4, Len(Rs3("last_sand_no").value) - 3)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format((NO + 1), String(Askcount, "0"))
                Else
                               
                    '            no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                    'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                    auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format((NO + 1), String(Askcount, "0"))
                    
                End If
              
            End If
                      
        End If

    End If

    Rs3.Close
    'Dim storeADDZero As String
    'storeADDZero = IIf(StoreID < 10, "0", "")
    Dim brancHcode As String
 
    brancHcode = zeropadding(CStr(my_branch), Int(SystemOptions.BranchDigit))
    storecode = zeropadding(storecode, Int(SystemOptions.StoreDigit))

    If numbering_type = 1 Then
        Voucher_codingoriginal = auto_sanad_no
        Exit Function
    End If
    If first_serial = True Then
        If auto_sanad_no <> "" Then
            
            If StoreCoding = True And StoreID <> 0 Then
                '       Voucher_codingoriginal = my_branch & storeADDZero & StoreID & auto_sanad_no
                Voucher_codingoriginal = brancHcode & storecode & auto_sanad_no
            Else
                Voucher_codingoriginal = brancHcode & auto_sanad_no
            End If
        
        Else
            Voucher_codingoriginal = auto_sanad_no
        End If

    Else
        '     Voucher_codingoriginal = my_branch & auto_sanad_no
        If StoreCoding = True And StoreID <> 0 Then
            ' Voucher_codingoriginal = my_branch & storeADDZero & StoreID & auto_sanad_no
            Voucher_codingoriginal = brancHcode & storecode & auto_sanad_no
        Else
            Voucher_codingoriginal = brancHcode & auto_sanad_no
        End If
    End If

End Function
 
Public Function OpeningVoucher_coding(my_branch As Integer, _
                                      date1 As Date, _
                                      Sanad_No As Integer, _
                                      NoteType As Integer, _
                                      Optional departement_name As Integer = 1, _
                                      Optional Transaction_Type As Integer = 0) As String
    On Error Resume Next
    Dim start_at       As Integer
    Dim end_at         As Integer
    Dim auto_sanad_no  As String
    Dim NO             As Integer
    Dim numbering_type As Integer
    auto_sanad_no = ""
 
    Dim first_serial As Boolean

    Dim rs           As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset

    Dim sql As String
    Dim i   As Integer
    first_serial = False
    sql = "select * from sanad_numbering where branch_no=" & my_branch & " and  sanad_no=" & Sanad_No
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = IIf(IsNull(rs("numbering_id").value), 0, rs("numbering_id").value)
        start_at = IIf(IsNull(rs("start_at").value), 0, rs("start_at").value)
        end_at = IIf(IsNull(rs("end_at").value), 0, rs("end_at").value)

    End If

    If numbering_type = 1 Then
        sql = "select max(NoteSerial1) as last_sand_no from  Notes1    where  branch_no= " & my_branch & "  and   NoteType=" & NoteType ' & " and   numbering_type1=" & numbering_type

        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes1 where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and not(BTCashAccountcode is null )"
        End If
   
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes1    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type
            End If
        End If
   
        If Sanad_No = 4 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes1 where  branch_no= " & my_branch & "  and    (NoteType= 3 or NoteType= 5 )  and  (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes1  where  branch_no= " & my_branch & "  and    NoteType=" & NoteType & " and  (BTCashAccountcode is null )"
            End If
        End If
   
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rs3("last_sand_no").value) Then
        
            If end_at = 0 Then end_at = val(Rs3("last_sand_no").value) + 1
               
            If Rs3("last_sand_no").value >= end_at Then
                OpeningVoucher_coding = "error"
                Exit Function
            End If
        End If
 
    ElseIf numbering_type = 2 Then

        sql = "select max(NoteSerial1) as last_sand_no from  Notes1 where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)

        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes1 where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and not(BTCashAccountcode is null )"
        End If
    
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes1 where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        
            End If
        
        End If
    
        If Sanad_No = 4 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes1 where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes1 where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
            End If
        
        End If
    
        If Transaction_Type <> 0 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
  
        End If
    
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rs3("last_sand_no").value) Then
            NO = mId(Rs3("last_sand_no").value, 7, Len(Rs3("last_sand_no").value) - 6)

            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                OpeningVoucher_coding = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then
 
        sql = "select max(NoteSerial1) as last_sand_no from  Notes1 where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
  
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If Not IsNull(Rs3("last_sand_no").value) Then
            NO = mId(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)

            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                OpeningVoucher_coding = "error"
                Exit Function
            End If
        End If
 
    End If

    Dim Askcount As Integer
    Askcount = 3

    If Rs3.RecordCount = 0 Or IsNull(Rs3("last_sand_no").value) Then
        first_serial = True

        If numbering_type = 0 Then
                 
        ElseIf numbering_type = 1 Then
            auto_sanad_no = start_at
        ElseIf numbering_type = 2 Then
            auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
        ElseIf numbering_type = 3 Then
            auto_sanad_no = mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(start_at, String(Askcount, "0"))
        End If

    Else

        If numbering_type = 0 Then
                
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Rs3("last_sand_no").value + 1
        ElseIf numbering_type = 2 Then
              
            NO = mId(Rs3("last_sand_no").value, 7, Len(Rs3("last_sand_no").value) - 6)
            auto_sanad_no = mId(Rs3("last_sand_no").value, 1, 6) & Format((NO + 1), String(Askcount, "0"))
        
        ElseIf numbering_type = 3 Then
         
            NO = mId(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
            auto_sanad_no = mId(Rs3("last_sand_no").value, 1, 5) & Format((NO + 1), String(Askcount, "0"))
        End If

    End If

    Rs3.Close

    If first_serial = True Then
        If auto_sanad_no <> "" Then
            OpeningVoucher_coding = my_branch & auto_sanad_no
        Else
            OpeningVoucher_coding = auto_sanad_no
        End If

    Else
        OpeningVoucher_coding = auto_sanad_no
    End If

End Function

Public Sub LoadSettings()
    'load database location
    On Error Resume Next

    State = GetSetting("Win_Sys_EX_B", "Setting", "State")
    run_count = GetSetting("Win_Sys_EX_B", "Setting", "run_count")
    key_for_me = GetSetting("Win_Sys_EX_B", "Setting", "key_for_me")
    Alarm_start = GetSetting("Win_Sys_EX_B", "Setting", "Alarm_start")
    Alarm_end = GetSetting("Win_Sys_EX_B", "Setting", "Alarm_end")

End Sub

Public Function get_balance_P(Account_Serial As String) As Double
    Dim total_credit As Double
    Dim total_depit  As Double
    Dim total        As Double
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    total_credit = 0: total_depit = 0:

    total = 0

    sql = "select sum(DEV_Value) As total_credit from RptLedger_Sub where Credit_Or_Debit=0 and  Account_Serial='" & Account_Serial & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Not IsNull(Rs3("total_credit").value) Then
        total_credit = Rs3("total_credit").value
    Else
        total_credit = 0
    End If

    Rs3.Close

    sql = "select sum(DEV_Value) As total_depit from RptLedger_Sub where Credit_Or_Debit=1 and  Account_Serial='" & Account_Serial & "'"
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not IsNull(Rs3("total_depit")) Then
        total_depit = Rs3("total_depit")
    Else
        total_depit = 0
    End If

    'Total = total_credit - total_depit
    get_balance_P = total_credit - total_depit

    '1 œ∆«∆‰
    '2„œÌ‰
End Function

Public Function get_late_COLOR(ID As Integer, _
                               Optional ByRef Name As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "Select * from Ageng_type where id=" & ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.EOF Then
        get_late_COLOR = ""
        Exit Function
    End If
    Name = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
    get_late_COLOR = IIf(IsNull(Rs3("COLOR").value), "", Rs3("COLOR").value)
 
    Rs3.Close

End Function

Public Function get_item_group_account_in_branch(ItemID As String, _
                                                 branch_id As Integer, _
                                                 account_type_code As Integer) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql     As String
    Dim i       As Integer
    Dim GroupID As Integer
    sql = "select ItemCode,GroupID  from TblItems where  ItemID  ='" & ItemID & "'"
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount = 0 Then
        get_item_group_account_in_branch = "Error"
        Exit Function
    End If
    GroupID = IIf(IsNull(Rs3("GroupID").value), 0, Rs3("GroupID").value)
    Rs3.Close
    
    sql = "select * from groups_account_in_inventory where  group_id  =" & GroupID & " and branch_id=" & branch_id & " and account_type_code='" & account_type_code & "'"
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Rs3.RecordCount = 0 Then
        get_item_group_account_in_branch = "Error"
        Exit Function
    End If
    get_item_group_account_in_branch = IIf(IsNull(Rs3("account_code").value), "Error", Rs3("account_code").value)
    Exit Function
    Rs3.Close
  
End Function

Public Function get_item_group_account_inventory(ItemID As String, _
                                                 inventory_id As Integer, _
                                                 account_type_code As Integer) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql     As String
    Dim i       As Integer
    Dim GroupID As Integer
    sql = "select ItemCode,GroupID  from TblItems where  ItemID  ='" & ItemID & "'"
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs3.RecordCount = 0 Then
        get_item_group_account_inventory = "Error"
        Exit Function
    End If
    GroupID = IIf(IsNull(Rs3("GroupID").value), 0, Rs3("GroupID").value)
    Rs3.Close
    
    sql = "select * from groups_account_in_inventory where  group_id  =" & GroupID & " and inventory_id=" & inventory_id & " and account_type_code='" & account_type_code & "'"
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Rs3.RecordCount = 0 Then
        get_item_group_account_inventory = "Error"
        Exit Function
    End If
    get_item_group_account_inventory = IIf(IsNull(Rs3("account_code").value), "Error", Rs3("account_code").value)
 
    Rs3.Close
  
End Function

Public Function get_item_id(itemcode As String) As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select ItemID  from TblItems where  ItemCode='" & itemcode & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_item_id = 0
        Exit Function
    End If
    get_item_id = IIf(IsNull(Rs3("ItemID").value), 0, Rs3("ItemID").value)
 
    Rs3.Close
  
End Function

Public Function get_project_id(Account_code As String, _
                               account_name As String) As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select id  from projects where " & account_name & "='" & Account_code & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_project_id = 0
        Exit Function
    End If
    get_project_id = IIf(IsNull(Rs3("id").value), 0, Rs3("id").value)
   
    Rs3.Close
  
End Function

Public Function get_project_customer_account(ID As Integer, _
                                             account_name As String, _
                                             Optional CustomerType As Integer = 0) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    If CustomerType = 0 Then
        sql = " SELECT     dbo.TblCustemers.Account_Code, dbo.TblCustemers.Account_Code1, dbo.TblCustemers.Account_Code2"
        sql = sql & " FROM         dbo.projects INNER JOIN"
        sql = sql & "  dbo.TblCustemers ON dbo.projects.End_user_id = dbo.TblCustemers.CusID"
        sql = sql & "  Where (dbo.Projects.id = " & ID & ")"
    Else
        sql = "  SELECT     dbo.projects.id, dbo.TblCustemers.Account_Code, dbo.TblCustemers.Account_Code1, dbo.TblCustemers.Account_Code2, dbo.projects.sub_contractor_id"
        sql = sql & "  FROM         dbo.projects INNER JOIN"
        sql = sql & "    dbo.TblCustemers ON dbo.projects.sub_contractor_id = dbo.TblCustemers.CusID"
        sql = sql & "  Where (dbo.Projects.id = " & ID & ")"

    End If

    'sql = "select " & account_name & "  from projects where id= " & id
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_project_customer_account = ""
        Exit Function
    End If
    get_project_customer_account = IIf(IsNull(Rs3(account_name).value), 0, Rs3(account_name).value)
    
    Rs3.Close
  
End Function

Public Function get_project_customer_id(ID As Integer, _
                                        account_name As String) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset

    Dim sql As String
    Dim i   As Integer

    ' sql = "select " & account_name & "  from projects where id= " & id
    sql = "SELECT     End_user_id, sub_contractor_id from dbo.Projects  WHERE     (id = " & ID & ") "
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount > 0 Then
        '        sql = "select *   from TblCustemers where Account_Code='" & Rs3(account_name).value & "'"
        ' Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Rs3.RecordCount > 0 Then
            If account_name = "End_user_Account" Then
                get_project_customer_id = val(IIf(IsNull(Rs3("End_user_id").value), 0, Rs3("End_user_id").value))
                Exit Function
            Else
                get_project_customer_id = val(IIf(IsNull(Rs3("sub_contractor_id").value), 0, Rs3("sub_contractor_id").value))
                Exit Function
            End If
         
            Rs3.Close
            Rs4.Close
    
        End If

    Else
        get_project_customer_id = 0
    End If
  
End Function

Public Function get_item_qty(Item_ID As Integer) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "select sum( QTY) as totalqty from QryItemsQTY(" & Item_ID & ")"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_item_qty = 0
        Exit Function
    End If
    get_item_qty = IIf(IsNull(Rs3("totalqty").value), 0, Rs3("totalqty").value)
   
    Rs3.Close
    Exit Function
End Function

Public Function get_item_Order_qty(Item_ID As Integer) As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql    As String
    Dim i      As Integer
    Dim netqty As Integer
    'Sql = "Select sum(Quantity)As total_order_qty from Transaction_Details where  not(order_id is null) and Item_ID=" & Item_id
    sql = "SELECT     dbo.Transactions.Transaction_Type, SUM(dbo.Transaction_Details.Quantity) AS Total_order_qty, dbo.Transaction_Details.Item_ID"
    sql = sql + " FROM         dbo.Transactions INNER JOIN"
    sql = sql + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    sql = sql + " GROUP BY dbo.Transactions.Transaction_Type, dbo.Transaction_Details.Item_ID"
    sql = sql + " HAVING      (dbo.Transactions.Transaction_Type = 6) AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_item_Order_qty = 0
        Exit Function
    End If
    netqty = IIf(IsNull(Rs3("total_order_qty").value), 0, Rs3("total_order_qty").value) - get_item_qty(Item_ID)

    If netqty < 0 Then netqty = 0
    get_item_Order_qty = netqty
    Rs3.Close
    Exit Function

End Function

Public Function get_item_Reserved_qty(Item_ID As Integer) As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer
    sql = "Select sum(Quantity)As total_rsv_qty from Transaction_Details where  not(Project_id is null) and Item_ID=" & Item_ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_item_Reserved_qty = 0
        Exit Function
    End If
    get_item_Reserved_qty = IIf(IsNull(Rs3("total_rsv_qty").value), 0, Rs3("total_rsv_qty").value)
  
    Rs3.Close

End Function

Public Function get_late_location2(days As Integer) As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "Select * from dbo.AgengItem_type"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_late_location2 = 0
        Exit Function
    End If
    For i = 1 To Rs3.RecordCount
  
        If Rs3("to").value <> 0 Then
            If days >= Rs3("from").value And days <= Rs3("to").value Then
                get_late_location2 = Rs3("id").value
                Exit Function
            End If

        Else

            If days >= Rs3("from").value Then
                get_late_location2 = Rs3("id").value
                Exit Function
            End If
  
        End If
  
        Rs3.MoveNext
    Next i
 
    Rs3.Close

End Function

Public Function get_late_location(days As Integer) As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i   As Integer

    sql = "Select * from Ageng_type"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_late_location = 0
        Exit Function
    End If
  
    For i = 1 To Rs3.RecordCount
  
        If Rs3("to").value <> 0 Then
            If days >= Rs3("from").value And days <= Rs3("to").value Then
                get_late_location = Rs3("id").value
                Exit Function
            End If

        Else

            If days >= Rs3("from").value Then
                get_late_location = Rs3("id").value
                Exit Function
            End If
  
        End If
  
        Rs3.MoveNext
    Next i
 
    Rs3.Close

End Function

Public Function get_Cheque_report_no(BankID As Integer) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    sql = "Select report_no from BanksData where BankID=" & BankID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_Cheque_report_no = ""
        Exit Function
    End If
    If IsNull(Rs3("report_no").value) Then
        get_Cheque_report_no = ""
        Exit Function
    End If
    If Not IsNull(Rs3("report_no").value) Then
        get_Cheque_report_no = Rs3("report_no").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_user_name(UserID As Integer) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    sql = "Select UserName from TblUsers where UserID=" & UserID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_user_name = ""
        Exit Function
    End If
    If IsNull(Rs3("UserName").value) Then
        get_user_name = ""
        Exit Function
    End If
    If Not IsNull(Rs3("UserName").value) Then
        get_user_name = Rs3("UserName").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_Notes_id(NoteSerial As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    sql = "Select * from Notes where NoteSerial='" & NoteSerial & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_Notes_id = 0
        Exit Function
    End If
    If IsNull(Rs3("NoteID").value) Then
        get_Notes_id = 0
        Exit Function
    End If
    If Not IsNull(Rs3("NoteID").value) Then
        get_Notes_id = Rs3("NoteID").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_branch_name(ID As Integer, _
                                Optional ByRef activityName As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset

    Dim sql        As String
    Dim ActivityId As Integer

    sql = "Select branch_nameE,branch_name,branch_id,ActivityTypeId from TblBranchesData where branch_id=" & ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_branch_name = ""
        Exit Function
    End If
    ActivityId = IIf(IsNull(Rs3("ActivityTypeId").value), 0, Rs3("ActivityTypeId").value)

    If SystemOptions.UserInterface = EnglishInterface Then
        If IsNull(Rs3("branch_nameE").value) Then
            get_branch_name = ""
            Exit Function
        End If
        If Not IsNull(Rs3("branch_nameE").value) Then
            get_branch_name = Rs3("branch_nameE").value
        End If
    Else

        If IsNull(Rs3("branch_name").value) Then
            get_branch_name = ""
        End If
        Exit Function
        If Not IsNull(Rs3("branch_name").value) Then
            get_branch_name = Rs3("branch_name").value
        End If
    End If
  
    Rs3.Close

    sql = "Select * from tblActivitesType where id=" & ActivityId
 
    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs4.RecordCount = 0 Then
        get_branch_name = ""
        Exit Function
    End If
    If SystemOptions.UserInterface = EnglishInterface Then
        If IsNull(Rs4("Namee").value) Then
            activityName = ""
            Exit Function
        End If
        If Not IsNull(Rs4("Namee").value) Then
            activityName = Rs4("Namee").value
            Exit Function
        End If
    Else

        If IsNull(Rs4("Name").value) Then
            activityName = ""
            Exit Function
        End If
        If Not IsNull(Rs4("Name").value) Then
            activityName = Rs4("Name").value
            Exit Function
        End If
    End If
  
    Rs4.Close

End Function

Public Function GET_ACCOUNT_CURRENCY(ID As Integer) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    sql = "Select * from branches where id=" & ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        GET_ACCOUNT_CURRENCY = ""
        Exit Function
    End If
    If SystemOptions.UserInterface = EnglishInterface Then
        If IsNull(Rs3("branch_nameE").value) Then
            GET_ACCOUNT_CURRENCY = ""
            Exit Function
        End If
        If Not IsNull(Rs3("branch_nameE").value) Then
            GET_ACCOUNT_CURRENCY = Rs3("branch_nameE").value
            Exit Function
        End If
    Else

        If IsNull(Rs3("branch_name").value) Then
            GET_ACCOUNT_CURRENCY = ""
            Exit Function
        End If
        If Not IsNull(Rs3("branch_name").value) Then
            GET_ACCOUNT_CURRENCY = Rs3("branch_name").value
            Exit Function
        End If
    End If
  
    Rs3.Close
End Function

Public Function GET_ACCOUNT_name_by_Code(Account_code As String, Optional ByVal mSerial As String = "") As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    sql = "Select * from ACCOUNTS where Account_Code='" & Account_code & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        GET_ACCOUNT_name_by_Code = ""
        Exit Function
    End If
    If mSerial = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            GET_ACCOUNT_name_by_Code = IIf(IsNull(Rs3("Account_NameEng").value), "", Rs3("Account_NameEng").value)
             
        Else
            GET_ACCOUNT_name_by_Code = IIf(IsNull(Rs3("Account_Name").value), "", Rs3("Account_Name").value)
             
        End If
    Else
        GET_ACCOUNT_name_by_Code = IIf(IsNull(Rs3("Account_Serial").value), "", Rs3("Account_Serial").value)
    End If
    Rs3.Close
End Function

Public Function get_note_type_name(note_type As Integer, _
                                   Optional ByRef NotesTypeNameE As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

    sql = "Select * from TblNotesTypes where NotesType=" & note_type
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_note_type_name = ""
        Exit Function
    End If
    If IsNull(Rs3("NotesTypeName").value) Then
        get_note_type_name = ""
        Exit Function
    End If
    If Not IsNull(Rs3("NotesTypeName").value) Then
        get_note_type_name = Rs3("NotesTypeName").value
        NotesTypeNameE = IIf(IsNull(Rs3("NotesTypeNamee").value), "", Rs3("NotesTypeNameE").value)
        Exit Function
    End If

    Rs3.Close

End Function
 
Public Function get_Financial_market_data(FinancialMarketId As Integer, _
                                          Optional ByRef FinancialMarketCode As String, _
                                          Optional ByRef FinancialMarketName As String, _
                                          Optional ByRef HyperLink As String, _
                                          Optional ByRef CountryID As Integer, _
                                          Optional ByRef TableID As Integer, _
                                          Optional ByRef RowID As Integer, _
                                          Optional ByRef CellID As Integer, _
                                          Optional ByRef HyperLinkGeneral As String)

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from FinancialMarkets where FinancialMarketId=" & FinancialMarketId
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Exit Function
    End If

    If Not IsNull(Rs3("FinancialMarketCode").value) Then FinancialMarketCode = Rs3("FinancialMarketCode").value
    If Not IsNull(Rs3("FinancialMarketName").value) Then FinancialMarketName = Rs3("FinancialMarketName").value
    If Not IsNull(Rs3("HyperLink").value) Then HyperLink = Rs3("HyperLink").value
    If Not IsNull(Rs3("HyperLinkGeneral").value) Then HyperLinkGeneral = Rs3("HyperLinkGeneral").value
     
    If Not IsNull(Rs3("CountryID").value) Then CountryID = val(Rs3("CountryID").value)
    If Not IsNull(Rs3("TableID").value) Then TableID = val(Rs3("TableID").value)
    If Not IsNull(Rs3("RowID").value) Then RowID = val(Rs3("RowID").value)
    If Not IsNull(Rs3("CellID").value) Then CellID = val(Rs3("CellID").value)
        
    Rs3.Close

End Function

Public Function get_FixedAsset_Account(group_id As Integer, _
                                       branch_id As Integer, _
                                       Optional Account_code As String = "") As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    ' sql = "Select * from FixedAssetsGroupsAccount where account_type_code=24 and group_id=" & group_id & " and branch_id=" & branch_id
    sql = "Select * from FixedAssetsGroup where   GroupID=" & group_id
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Account_code = "" Then
        If Rs3.RecordCount = 0 Then
            get_FixedAsset_Account = ""
            Exit Function
        End If
        If IsNull(Rs3("Account_Code").value) Then
            get_FixedAsset_Account = ""
            Exit Function
        End If
        If Not IsNull(Rs3("Account_Code").value) Then
            get_FixedAsset_Account = Rs3("account_code").value
            Exit Function
        End If
    Else

        If Rs3.RecordCount = 0 Then
            get_FixedAsset_Account = ""
            Exit Function
        End If
        If IsNull(Rs3(Account_code).value) Then
            get_FixedAsset_Account = ""
            Exit Function
        End If
        If Not IsNull(Rs3(Account_code).value) Then
            get_FixedAsset_Account = Rs3(Account_code).value
            Exit Function
        End If
  
    End If
  
    Rs3.Close

End Function

Public Function get_Revenue_id(Account_code As String, _
                               Optional ByRef AccountName) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from TblRevenuesTypes where Account_Code='" & Account_code & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_Revenue_id = 0
        Exit Function
    End If
    If IsNull(Rs3("RevenuesID").value) Then
        get_Revenue_id = 0
        Exit Function
    End If
    AccountName = IIf(IsNull(Rs3("RevenuesName").value), "", Rs3("RevenuesName").value)
  
    If Not IsNull(Rs3("RevenuesID").value) Then
        get_Revenue_id = Rs3("RevenuesID").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_Expenses_id(Account_code As String, _
                                Optional ByRef AccountName) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from ExpensesType where Account_Code='" & Account_code & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_Expenses_id = 0
        Exit Function
    End If
    If IsNull(Rs3("id").value) Then
        get_Expenses_id = 0
        Exit Function
    End If
    AccountName = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
  
    If Not IsNull(Rs3("id").value) Then
        get_Expenses_id = Rs3("id").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_Customer_id(Account_code As String) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select CusID from TblCustemers where Account_Code='" & Account_code & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_Customer_id = 0
        Exit Function
    End If
    If IsNull(Rs3("CusID").value) Then
        get_Customer_id = 0
        Exit Function
    End If
    If Not IsNull(Rs3("CusID").value) Then
        get_Customer_id = Rs3("CusID").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function CHECK_LAST_ACCOUNT(account As String) As Boolean
    Dim rs As ADODB.Recordset
    StrSQL = "Select * From Accounts Where Account_Code='" & account & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        
        If rs("last_account").value = True Then
            CHECK_LAST_ACCOUNT = True
            Exit Function
        Else
            CHECK_LAST_ACCOUNT = False
            Exit Function
        End If

    Else
        CHECK_LAST_ACCOUNT = True
        Exit Function
    End If
  
End Function

Public Function CountA(ByVal sText As String) As Long
    Dim bArr() As Byte
    Dim i      As Long
    Dim count  As Long
 
    For i = 1 To Len(sText)

        ' if this char is a space, increase the counter
        If mId(sText, i, 1) = "a" Then count = count + 1
    Next

    CountA = count
End Function

Public Function create_Branch_group(branch_id As Integer, _
                                    group_id As Integer, _
                                    group_name As String) As Boolean
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql         As String
    Dim X           As String
    Dim group_nameA As String
    Dim group_namee As String
    sql = "Select * from branches where  branch_id =" & branch_id
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        create_Branch_group = False
        Exit Function
    End If
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "[groups_account_in_inventory]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    If Not IsNull(Rs3("a1").value) And Rs3("a1").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a1").value) = False Then
               
            group_nameA = group_name

            If SystemOptions.UserInterface = ArabicInterface Then
                group_nameA = "    Õ”«»  ﬂ·›… «·„»Ì⁄« " + "  " + group_nameA
                group_namee = group_name + " " + "Cost Of Sale Acc."
            Else
                group_nameA = "    Õ”«»  ﬂ·›… «·„»Ì⁄« " + "  " + group_nameA
                group_namee = group_name + " " + "Cost Of Sale Acc."
            End If
        
            X = ModAccounts.AddNewAccount(Rs3("a1").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 1 '  ﬂ·›… «·„»Ì⁄« 
            rs("account_code").value = X
              
            rs.update
        End If
    End If
  
    If Not IsNull(Rs3("a2").value) And Rs3("a2").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a2").value) = False Then
            group_nameA = group_name

            If SystemOptions.UserInterface = ArabicInterface Then
                group_nameA = "    Õ”«»   «·„»Ì⁄« " + "  " + group_nameA
                group_namee = group_name + " " + " Sale Acc."
            Else
                group_nameA = "    Õ”«»   «·„»Ì⁄« " + "  " + group_nameA
                group_namee = group_name + " " + " Sale Acc."
            End If
        
            X = ModAccounts.AddNewAccount(Rs3("a2").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 2 '   «·„»Ì⁄« 
            rs("account_code").value = X
               
            rs.update
        End If
    End If
 
    If Not IsNull(Rs3("a3").value) And Rs3("a3").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a3").value) = False Then
            group_nameA = group_name

            If SystemOptions.UserInterface = ArabicInterface Then
                group_nameA = "    Õ”«»  „—œÊœ«  «·„»Ì⁄« " + "  " + group_nameA
                group_namee = group_name + " " + " Sale  Return Acc."
            Else
                group_nameA = "    Õ”«»  „—œÊœ«  «·„»Ì⁄« " + "  " + group_nameA
                group_namee = group_name + " " + " Sale  Return Acc."
            End If
        
            X = ModAccounts.AddNewAccount(Rs3("a3").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 3 '„—œÊœ«  «·„»Ì⁄« 
            rs("account_code").value = X
              
            rs.update
        End If
    End If
  
    If Not IsNull(Rs3("a4").value) And Rs3("a4").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a4").value) = False Then
            group_nameA = group_name

            If SystemOptions.UserInterface = ArabicInterface Then
                group_nameA = "    Õ”«» «·„‘ —Ì« " + "  " + group_nameA
                group_namee = group_name + " " + " Purchase Acc."
            Else
                group_nameA = "    Õ”«» «·„‘ —Ì« " + "  " + group_nameA
                group_namee = group_name + " " + " Purchase Acc."
            End If
        
            X = ModAccounts.AddNewAccount(Rs3("a4").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 4 '«·„‘ —Ì« 
            rs("account_code").value = X
              
            rs.update
        End If
    End If

    If Not IsNull(Rs3("a5").value) And Rs3("a5").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a5").value) = False Then
            group_nameA = group_name

            If SystemOptions.UserInterface = ArabicInterface Then
                group_nameA = "    Õ”«»  „—œÊœ«  «·„‘ —Ì« " + "  " + group_nameA
                group_namee = group_name + " " + " Purchase Return Acc."
            Else
                group_nameA = "    Õ”«»  „—œÊœ«  «·„‘ —Ì« " + "  " + group_nameA
                group_namee = group_name + " " + " Purchase Return Acc."
            End If
        
            X = ModAccounts.AddNewAccount(Rs3("a5").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 5 '„—œÊœ«  «·„‘ —Ì« 
            rs("account_code").value = X
               
            rs.update
        End If

    End If

    If Not IsNull(Rs3("a12").value) And Rs3("a12").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a12").value) = False Then
            group_nameA = group_name

            If SystemOptions.UserInterface = ArabicInterface Then
                group_nameA = "    Õ”«»  Œ’„ „”„ÊÕ »Â" + "  " + group_nameA
                group_namee = group_name + " " + " Discount allowed Acc."
            Else
                group_nameA = "    Õ”«»  Œ’„ „”„ÊÕ »Â" + "  " + group_nameA
                group_namee = group_name + " " + " Discount allowed Acc."
            End If
        
            X = ModAccounts.AddNewAccount(Rs3("a12").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 12 'Œ’„ „”„ÊÕ »…
            rs("account_code").value = X
               
            rs.update
        End If
    End If
  
    If Not IsNull(Rs3("a13").value) And Rs3("a13").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a13").value) = False Then
            group_nameA = group_name

            If SystemOptions.UserInterface = ArabicInterface Then
                group_nameA = "    Õ”«»  Œ’„ „ﬂ ”» " + "  " + group_nameA
                group_namee = group_name + " " + "Unearned discount Acc."
            Else
                group_nameA = "    Õ”«»  Œ’„ „ﬂ ”» " + "  " + group_nameA
                group_namee = group_name + " " + "Unearned discount Acc."
            End If
        
            X = ModAccounts.AddNewAccount(Rs3("a13").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 13 'Œ’„ „ﬂ ”»
            rs("account_code").value = X
               
            rs.update
        End If
    End If

    rs.Close
    Rs3.Close

    create_Branch_group = True
End Function

Public Function create_Branch_FixedAssets_group(branch_id As Integer, _
                                                group_id As Integer, _
                                                group_name As String) As Boolean
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql         As String
    Dim X           As String
    Dim group_nameA As String
    Dim group_namee As String
    sql = "Select * from branches where  branch_id =" & branch_id
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        create_Branch_FixedAssets_group = False
        Exit Function
    End If
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "[FixedAssetsGroupsAccount]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    If Not IsNull(Rs3("a24").value) And Rs3("a24").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a24").value) = False Then
            group_nameA = group_name
       
            group_nameA = "Õ”«» ﬁÌ„… «·«’Ê· «·À«» … " + " ›—⁄ " + Rs3("branch_name").value + "-" + group_nameA
            group_namee = "Fixed Assets value Account   " + Rs3("branch_namee").value + "-" + group_name
          
            X = ModAccounts.AddNewAccount(Rs3("a24").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 24 ' ﬁÌ„… «·«’Ê· «·À«» …
            rs("account_code").value = X
              
            rs.update
        End If
    End If
   
    If Not IsNull(Rs3("a25").value) And Rs3("a25").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a25").value) = False Then
            group_nameA = group_name
       
            group_nameA = "Õ”«» „’—Ê› «·«Â·«ﬂ " + " ›—⁄ " + Rs3("branch_name").value + "-" + group_nameA
            group_namee = "Damage Expenses" + Rs3("branch_namee").value + "-" + group_name
          
            X = ModAccounts.AddNewAccount(Rs3("a25").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 25 '„’—Ê› «·«Â·«ﬂ
            rs("account_code").value = X
              
            rs.update
        End If
    End If
    
    If Not IsNull(Rs3("a26").value) And Rs3("a26").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a26").value) = False Then
            group_nameA = group_name
       
            group_nameA = "Õ”«» „Ã„⁄ «·«Â·«ﬂ " + " ›—⁄ " + Rs3("branch_name").value + "-" + group_nameA
            group_namee = "Damage Expenses  Sum" + Rs3("branch_namee").value + "-" + group_name
          
            X = ModAccounts.AddNewAccount(Rs3("a26").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 26 '    „Ã„⁄ «·«Â·«ﬂ
            rs("account_code").value = X
              
            rs.update
        End If
    End If
   
    If Not IsNull(Rs3("a31").value) And Rs3("a31").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a31").value) = False Then
            group_nameA = group_name
       
            group_nameA = "Õ”«» «—»«Õ »Ì⁄ «.À«» … " + " ›—⁄ " + Rs3("branch_name").value + "-" + group_nameA
            group_namee = " FA Sales Profit" + Rs3("branch_namee").value + "-" + group_name
          
            X = ModAccounts.AddNewAccount(Rs3("a31").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 31 '       Õ”«» «—»«Õ »Ì⁄ «.À«» …
            rs("account_code").value = X
              
            rs.update
        End If
    End If
    
    If Not IsNull(Rs3("a40").value) And Rs3("a40").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("a40").value) = False Then
            group_nameA = group_name
       
            group_nameA = "Õ”«» Œ”«—… »Ì⁄ «.À«» … " + " ›—⁄ " + Rs3("branch_name").value + "-" + group_nameA
            group_namee = " FA Sales loss" + Rs3("branch_namee").value + "-" + group_name
          
            X = ModAccounts.AddNewAccount(Rs3("a40").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("branch_id").value = branch_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 40 '      Õ”«» Œ”«—… »Ì⁄ «.À«» …
            rs("account_code").value = X
              
            rs.update
        End If
    End If
    
    rs.Close
    Rs3.Close

    create_Branch_FixedAssets_group = True
End Function

Public Function create_inventory_group(inv_id As Integer, _
                                       group_id As Integer, _
                                       group_name As String) As Boolean
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql         As String
    Dim X           As String
    Dim group_namee As String
    Dim group_nameA As String
    sql = "Select * from TblStore where  StoreID =" & inv_id
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        create_inventory_group = False
        Exit Function
    End If
    If Rs3.RecordCount > 0 Then
        group_name = IIf(Not IsNull(Rs3("StoreName").value), "", Rs3("StoreName").value) + "  " + group_name
    End If
  
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "[groups_account_in_inventory]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    If Not IsNull(Rs3("Account_Code").value) And Rs3("Account_Code").value <> "" Then
        If CHECK_LAST_ACCOUNT(Rs3("Account_Code").value) = False Then
            group_nameA = group_name

            If SystemOptions.UserInterface = ArabicInterface Then
                group_nameA = " Õ”«»  «·„Œ“Ê‰ " + "  " + group_nameA
                group_namee = group_name + " " + "Stock Acc."
            Else
                group_nameA = " Õ”«»  «·„Œ“Ê‰ " + "  " + group_nameA
                group_namee = group_name + " " + "Stock Acc."
            End If
        
            X = ModAccounts.AddNewAccount(Rs3("Account_Code").value, group_nameA, True, False, group_namee)
            rs.AddNew
            rs("inventory_id").value = inv_id
            rs("group_id").value = group_id
            rs("account_type_code").value = 0 'Õ”«»  «·„Œ“Ê‰
            rs("account_code").value = X
              
            rs.update
        End If
    End If

    group_nameA = group_name

    If CHECK_LAST_ACCOUNT(Rs3("Account_Code1").value) = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            group_nameA = "  Õ”«» Œ”«∆— ›ﬁœ Ê ·› " + "  " + group_nameA
            group_namee = group_name + " " + "Loass & Damage Account Acc.  "
        Else
            group_nameA = "  Õ”«» Œ”«∆— ›ﬁœ Ê ·› " + "  " + group_nameA
            group_namee = group_name + " " + "Loass & Damage Account Acc.  "
        End If
        
        X = ModAccounts.AddNewAccount(Rs3("Account_Code1").value, group_nameA, True, False, group_namee)
        rs.AddNew
        rs("inventory_id").value = inv_id
        rs("group_id").value = group_id
        rs("account_type_code").value = 10 'Õ”«» Œ”«∆— ›ﬁœ Ê ·›
        rs("account_code").value = X
        
        rs.update
    End If
 
    group_nameA = group_name

    If CHECK_LAST_ACCOUNT(Rs3("Account_Code2").value) = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            group_nameA = " Õ”«» «· ”ÊÌ«  «·Ã—œÌ…" + "  " + group_nameA
            group_namee = group_name + " " + " Inventory adjust. Acc."
        Else
            group_nameA = " Õ”«» «· ”ÊÌ«  «·Ã—œÌ…" + "  " + group_nameA
            group_namee = group_name + " " + " Inventory adjust. Acc."
        End If
        
        X = ModAccounts.AddNewAccount(Rs3("Account_Code2").value, group_nameA, True, False, group_namee)
        rs.AddNew
        rs("inventory_id").value = inv_id
        rs("group_id").value = group_id
        rs("account_type_code").value = 11 'Õ”«» «· ”ÊÌ«  «·Ã—œÌ…
        rs("account_code").value = X
        
        rs.update
    End If

    group_nameA = group_name

    If CHECK_LAST_ACCOUNT(Rs3("Account_Code3").value) = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            group_nameA = " Õ”«» Âœ«Ì« Ê⁄Ì‰«  " + "  " + group_nameA
            group_namee = group_name + " " + " Sample & Gifts Acc. "
        Else
            group_nameA = " Õ”«» Âœ«Ì« Ê⁄Ì‰«  " + "  " + group_nameA
            group_namee = group_name + " " + " Sample & Gifts Acc. "
        End If
        
        X = ModAccounts.AddNewAccount(Rs3("Account_Code3").value, group_nameA, True, False, group_namee)
        rs.AddNew
        rs("inventory_id").value = inv_id
        rs("group_id").value = group_id
        rs("account_type_code").value = 17 'Õ”«» Âœ«Ì« Ê⁄Ì‰« 
        rs("account_code").value = X
        
        rs.update
    End If

    rs.Close
    Rs3.Close

    create_inventory_group = True
End Function

Public Function get_Employee_project_information(Emp_id As Integer)
    '    Dim Rs3 As ADODB.Recordset
    '    Set Rs3 = New ADODB.Recordset
    '    Dim sql As String
    '    sql = "Select * from emp_all_details_with_project_name where  Emp_Id =" & Emp_id
 
    '    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    '    If Rs3.RecordCount = 0 Then Rs3.Close: Exit Function
    '    Unload InfoBar
    '    InfoBar.show
    '    InfoBar.LblEmpID = IIf(IsNull(Rs3("Emp_Id").value), "", Rs3("Emp_Id").value)
    '    InfoBar.LblEmpCode = IIf(IsNull(Rs3("Emp_code").value), "", Rs3("Emp_code").value)
    '    InfoBar.LblEmpName = IIf(IsNull(Rs3("Emp_name").value), "", Rs3("Emp_name").value)
    '    InfoBar.lblprojectName = IIf(IsNull(Rs3("project_name").value), "", Rs3("project_name").value)
    '    InfoBar.LblTermName = IIf(IsNull(Rs3("term_fullcode").value), "", Rs3("term_fullcode").value)
    '    InfoBar.LblOprname = IIf(IsNull(Rs3("opr_fullcode").value), "", Rs3("opr_fullcode").value)
  
    '    Rs3.Close:  Exit Function

End Function

Public Function get_Customer_Account(CusID As Double) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select Account_Code from TblCustemers where  CusID =" & CusID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_Customer_Account = 0
        Exit Function
    End If
    If IsNull(Rs3("Account_Code").value) Then
        get_Customer_Account = 0
        Exit Function
    End If
    If Not IsNull(Rs3("Account_Code").value) Then
        get_Customer_Account = Rs3("Account_Code").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_expanses_id(Account_code As String) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select id from ExpensesType where Account_Code='" & Account_code & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_expanses_id = 0
        Exit Function
    End If
    If IsNull(Rs3("id").value) Then
        get_expanses_id = 0
        Exit Function
    End If
    If Not IsNull(Rs3("id").value) Then
        get_expanses_id = Rs3("id").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_opr_expenses_total(fullcode As String, _
                                       to_date As Date) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
 
    sql = "SELECT     opr_fullcode, SUM(value) AS total"
    sql = sql + " from DOUBLE_ENTREY_VOUCHERS"
    sql = sql + " WHERE      recorddate<='" & SQLDate(to_date)
    sql = sql + "' and  opr_fullcode='" & fullcode & "'"
    sql = sql + " GROUP BY opr_fullcode"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_opr_expenses_total = 0
        Exit Function
    End If
    If IsNull(Rs3("total").value) Then
        get_opr_expenses_total = 0
        Exit Function
    End If
    If Not IsNull(Rs3("total").value) Then
        get_opr_expenses_total = Rs3("total").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_opr_material_total(fullcode As String, _
                                       to_date As Date) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
 
    sql = "SELECT     opr_fullcode, SUM(total) AS total"
    sql = sql + " from dbo.opr_qty_total"
    sql = sql + " WHERE      Transaction_Date<='" & SQLDate(to_date)
    sql = sql + "' and  opr_fullcode='" & fullcode & "'"
    sql = sql + " GROUP BY opr_fullcode"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_opr_material_total = 0
        Exit Function
    End If
    If IsNull(Rs3("total").value) Then
        get_opr_material_total = 0
        Exit Function
    End If
    If Not IsNull(Rs3("total").value) Then
        get_opr_material_total = Rs3("total").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_notes_foxy_no(salary As String, _
   filed As String) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from Notes where NoteType =5 or NoteType =200 and salary='" & salary & "'"

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_notes_foxy_no = 0
        Exit Function
    End If
    If IsNull(Rs3(filed).value) Then
        get_notes_foxy_no = 0
        Exit Function
    End If
    If Not IsNull(Rs3(filed).value) Then
        get_notes_foxy_no = Rs3(filed).value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_COST_CENTER_NAME(code As String, _
   filed As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from markaas_taklefa where code='" & code & "'"

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_COST_CENTER_NAME = ""
        Exit Function
    End If
    If IsNull(Rs3(filed).value) Then
        get_COST_CENTER_NAME = ""
        Exit Function
    End If
    If Not IsNull(Rs3(filed).value) Then
        get_COST_CENTER_NAME = Rs3(filed).value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_EMPLOYEEIdFromAccountCode(Account_code As String, _
                                              filed As String) As Long
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from TblEmployee where Account_code='" & Account_code & "'"

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_EMPLOYEEIdFromAccountCode = 0
        Exit Function
    End If
    If IsNull(Rs3(filed).value) Then
        get_EMPLOYEEIdFromAccountCode = 0
        Exit Function
    End If
    If Not IsNull(Rs3(filed).value) Then
        get_EMPLOYEEIdFromAccountCode = Rs3(filed).value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_StoreBYSalesPerson(SalesPersonid As Double) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select StoreID from TblStore where SalesPersonid=" & SalesPersonid

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_StoreBYSalesPerson = 0
        Exit Function
    End If
    If IsNull(Rs3("StoreID").value) Then
        get_StoreBYSalesPerson = 0
        Exit Function
    End If
    If Not IsNull(Rs3("StoreID").value) Then
        get_StoreBYSalesPerson = Rs3("StoreID").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_EMPLOYEE_Data(Emp_id As Integer, _
                                  filed As String, _
                                  Optional ByRef jobstatus As Integer = 0) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from TblEmployee where Emp_ID=" & Emp_id

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_EMPLOYEE_Data = ""
        Exit Function
    End If
    If IsNull(Rs3(filed).value) Then
        get_EMPLOYEE_Data = ""
        Exit Function
    End If
    If Not IsNull(Rs3(filed).value) Then
        get_EMPLOYEE_Data = Rs3(filed).value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_EMPLOYEE_COST_CENTER_NAME(ID As String, _
                                              filed As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from markaas_taklefa where ACCOUNT_NO='" & ID & "'"

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_EMPLOYEE_COST_CENTER_NAME = ""
        Exit Function
    End If
    If IsNull(Rs3(filed).value) Then
        get_EMPLOYEE_COST_CENTER_NAME = ""
        Exit Function
    End If
    If Not IsNull(Rs3(filed).value) Then
        get_EMPLOYEE_COST_CENTER_NAME = Rs3(filed).value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_NO_OF_row(kedno As Double, _
                              account_no As String, _
                              LineNo1 As Double) As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
  
    sql = "select * from marakes_taklefa_temp   where kedno =" & kedno & " and account_no='" & account_no & "' and  line_no=" & LineNo1

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_NO_OF_row = 0
        Exit Function
    End If
    get_NO_OF_row = Rs3.RecordCount
    Rs3.Close
    Exit Function
    

End Function

Public Function get_EMPLOYEE_Account(ID As String, _
   filed As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from TblEmployee where Emp_ID=" & ID

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_EMPLOYEE_Account = ""
        Exit Function
    End If
    If IsNull(Rs3(filed).value) Then
        get_EMPLOYEE_Account = ""
        Exit Function
    End If
    If Not IsNull(Rs3(filed).value) Then
        get_EMPLOYEE_Account = Rs3(filed).value
        Exit Function
    End If
    Rs3.Close

End Function
 
Public Function Change_filed_value(ID As Double, _
                                   search_FILED As String, _
                                   filed As String, _
                                   table As String, _
                                   value) As Boolean
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  *   from " & table & " where " & search_FILED & "=" & ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Change_filed_value = False
        Exit Function
    End If
    Rs3(filed).value = value
    Rs3.update
    Rs3.Close
    Change_filed_value = True

End Function

Public Function get_project_Account(ID As Integer, _
   filed As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from projects where id=" & ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_project_Account = ""
        Exit Function
    End If
    If IsNull(Rs3(filed).value) Then
        get_project_Account = ""
        Exit Function
    End If
    
    If Not IsNull(Rs3(filed).value) Then
        get_project_Account = Rs3(filed).value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function replace_in_data_base(table As String, _
                                     filed As String, _
                                     str1 As String, _
                                     str2 As String) As Boolean
    sql = " Update   " & table & "  Set    " & filed & " = replace(" & filed & ", '" & str1 & "', '" & str2 & "');"
    Cn.Execute sql
    replace_in_data_base = True
End Function

Public Function get_bank_Account(ID As Long, _
   filed As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from BanksData where BankID=" & ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_bank_Account = ""
        Exit Function
    End If
    If IsNull(Rs3(filed).value) Then
        get_bank_Account = ""
        Exit Function
    End If
    If Not IsNull(Rs3(filed).value) Then
        get_bank_Account = Rs3(filed).value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function get_store_Account(ID As Integer, _
   filed As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from TblStore where StoreID=" & ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        get_store_Account = ""
        Exit Function
    End If
    If IsNull(Rs3(filed).value) Then
        get_store_Account = ""
        Exit Function
    End If
    If Not IsNull(Rs3(filed).value) Then
        get_store_Account = Rs3(filed).value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function CheckAccountToJE(Account_code As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql         As String
    Dim AccountName As String
    branch_id = 1
    AccountName = "a" & account_index
    sql = "Select * from ACCOUNTS where   Account_Code='" & Account_code & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        CheckAccountToJE = False
        Exit Function
    End If
  
    If IsNull(rs("Account_Code").value) Or rs("Account_Code").value = "" Then
        CheckAccountToJE = False
        Exit Function
    End If
    If Not IsNull(rs("Account_Code").value) Then
        CheckAccountToJE = True
        Exit Function
  
    End If
  
    rs.Close

End Function

Public Function get_account_code_branch(account_index As Integer, _
                                        branch_id As String, Optional ByVal mTxt As String = "a") As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql         As String
    Dim AccountName As String
    branch_id = 1
    AccountName = mTxt & account_index
    sql = "Select * from branches " 'where branch_id='" & branch_id & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        get_account_code_branch = "NO branch"
        Exit Function
    End If
    rs.MoveFirst
    
    If IsNull(rs(AccountName).value) Or rs(AccountName).value = "" Then
        get_account_code_branch = "NO account"
        Exit Function
    End If
    
    If mTxt = "a" Then
    If Not IsNull(rs(AccountName).value) Then
        If CheckAccountToJE(rs(AccountName).value) = True Then
            get_account_code_branch = rs(AccountName).value
            Exit Function
        Else
            get_account_code_branch = "NO account"
            Exit Function
        End If
  
    End If
    Else
        get_account_code_branch = rs(AccountName).value
        Exit Function
    
   End If
    rs.Close
End Function

Public Sub save_confoguration(State As String, _
                              run_count As Integer, _
                              key_for_me As String)
    On Error Resume Next
    SaveSetting "Win_Sys_EX_B", "Setting", "state", State
    SaveSetting "Win_Sys_EX_B", "Setting", "run_count", run_count
    SaveSetting "Win_Sys_EX_B", "Setting", "key_for_me", key_for_me

End Sub

Private Function ConvertDate(ByRef StringIn As String, _
                             ByRef OldCalender As Integer, _
                             ByVal NewCalender As Integer, _
                             ByRef NewFormat As String) As String
    If StringIn = "" Then
        Exit Function
    End If
    On Error Resume Next
    Dim SavedCal As Integer
    Dim d        As Date, s As String
    SavedCal = Calendar
    Calendar = OldCalender
    d = CDate(StringIn)
    Calendar = NewCalender
    s = CStr(d)
    ConvertDate = Format(s, NewFormat)
    Calendar = SavedCal
End Function

Public Function ToHijriDate(ByVal GregorianDate As String) As String
    Dim HijriDate As String, DateFormat As String
    ' DateFormat = "long date"
    
    DateFormat = "dd-mm-yyyy"
    HijriDate = ConvertDate(GregorianDate, vbCalGreg, vbCalHijri, DateFormat)
    ToHijriDate = HijriDate
    
End Function

Public Function ToGregorianDate(ByVal HijriDate As String) As Date
    Dim GregorianDate As String, DateFormat As String
    If HijriDate = "" Then
        Exit Function
    End If
    DateFormat = "dd/mm/yyyy"
    
    GregorianDate = ConvertDate(HijriDate, vbCalHijri, vbCalGreg, DateFormat)
    If DateDiff("D", "01/01/1900", GregorianDate) < 0 Then
        GregorianDate = Date
    End If
    ToGregorianDate = GregorianDate
End Function

Public Function GET_COST_PRICE_FOR_PRODUCT_ITEM(LngItemID As Long) As Long
    '131315
    
    GET_COST_PRICE_FOR_PRODUCT_ITEM = 0
    Exit Function
    Dim StrSQL  As String
    Dim RsParts As ADODB.Recordset
    Dim i       As Integer
    GET_COST_PRICE_FOR_PRODUCT_ITEM = 0
    StrSQL = "SELECT TableID, ItemID, PartItemID, PartItemQty, PartItemPrice"
    StrSQL = StrSQL + " FROM TblItemsParts Where ItemID=" & LngItemID
    StrSQL = StrSQL + " Order By TableID"
    Set RsParts = New ADODB.Recordset
    RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsParts.EOF Or RsParts.BOF) Then

        For i = 0 To RsParts.RecordCount - 1
            GET_COST_PRICE_FOR_PRODUCT_ITEM = GET_COST_PRICE_FOR_PRODUCT_ITEM + ModItemCostPrice.GetCostItemPrice(RsParts("PartItemID").value, 0, , , SystemOptions.SysMainStockCostMethod) * RsParts("PartItemQty").value
            RsParts.MoveNext
        Next i

    End If

End Function

Public Function getCustomerAgeingData(StrAccountCode As String, _
                                      Optional ByRef salesPersonName As String, _
                                      Optional allCustomer As Boolean = False, _
                                      Optional OnlyCheck As Boolean = False) As String
    Dim NameOfAgeType        As String

    Dim late_interval        As Integer
    Dim Dean_age             As Integer
                         
    Dim column_location      As Integer
    Dim column_COLOR         As String
    Dim customerid           As Long
    Dim i                    As Integer
    Dim sql                  As String
    Dim DefaultSalesPersonId As Integer
    Dim Rs3                  As New ADODB.Recordset

    If allCustomer = True Then
        GoTo ll:
    End If

    customerid = GetCustomerIdByAccountCode(StrAccountCode)

    GetCustomersDetail customerid, DefaultSalesPersonId
    getemployeeCode DefaultSalesPersonId, salesPersonName
 
    If customerid = 0 Then
        getCustomerAgeingData = ""

        Exit Function
    Else
        getCustomerAgeingData = customerid
    End If

    If OnlyCheck = True Then
        Exit Function
    End If
    getCustomerAgeingData = ""
ll:
    sql = "SELECT     dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.NoteSerial1, CompanyCreditValues.*"
    sql = sql & " FROM         dbo.CompanyCreditValues() CompanyCreditValues INNER JOIN"
    sql = sql & "  dbo.TblCustemers ON CompanyCreditValues.CusID = dbo.TblCustemers.CusID INNER JOIN"
    sql = sql & " dbo.Transactions ON CompanyCreditValues.TransactionsID = dbo.Transactions.Transaction_ID"

    If allCustomer = True Then
        sql = sql & "  WHERE     (CompanyCreditValues.RequiredValue > 0)  "
    Else
        sql = sql & "  WHERE     (CompanyCreditValues.RequiredValue > 0) and TblCustemers.CusID=" & customerid
    End If
 
    Dim str        As String
    Dim Note_Value As Double
    str = "delete TblTempCustomerAging"
 
    Cn.Execute str
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Exit Function
    End If
 
    If Rs3.RecordCount > 0 Then
      
        Rs3.MoveFirst
         
        For i = 1 To Rs3.RecordCount
              
            'CurrentString = IIf(IsNull(Rs3.Fields("NoteSerial1").value), _
             "", Rs3.Fields("NoteSerial1").value)
            'CurrentString = padding(Trim(CurrentString), 20)
 
            '                  getCustomerAgeingData = getCustomerAgeingData & CurrentString
 
            'CurrentString = IIf(IsNull(Rs3.Fields("transactiontypename").value), _
             "", Rs3.Fields("transactiontypename").value)
                       
            'CurrentString = padding(Trim(CurrentString), 20)
 
            '           getCustomerAgeingData = getCustomerAgeingData & vbTab & CurrentString
            'RequiredValue
            'Note_Value
                        
            CurrentString = IIf(IsNull(Rs3.Fields("RequiredValue").value), 0, Rs3.Fields("RequiredValue").value)
                       
            Note_Value = IIf(IsNull(Rs3.Fields("RequiredValue").value), 0, Rs3.Fields("RequiredValue").value)
                       
            'CurrentString = padding(Trim(CurrentString), 20)
 
            '                   getCustomerAgeingData = getCustomerAgeingData & CurrentString
                       
            '                      CurrentString = IIf(IsNull(Rs3.Fields("duedate").value), _
                                   "", Rs3.Fields("duedate").value)
 
            'CurrentString = padding(Trim(CurrentString), 20)
 
            '           getCustomerAgeingData = getCustomerAgeingData & CurrentString
            '         getCustomerAgeingData = getCustomerAgeingData & Chr(13)
          
            late_interval = DateDiff("d", Rs3.Fields("duedate").value, Date, vbSaturday)
                       
            CurrentString = late_interval
 
            'CurrentString = padding(Trim(CurrentString), 20)
 
            '          getCustomerAgeingData = getCustomerAgeingData & vbTab & CurrentString
                       
            column_location = get_late_location(late_interval)
            column_COLOR = get_late_COLOR(column_location, NameOfAgeType)
                      
            CurrentString = NameOfAgeType

            'CurrentString = padding(Trim(CurrentString), 20)
            If allCustomer = False Then
                add_record_to_table "TblTempCustomerAging", "CustD,LateID,DueValue  ", customerid & " ," & column_location & " ," & Note_Value & "", "CustD", 0
            Else
                add_record_to_table "TblTempCustomerAging", "CustD,LateID,DueValue  ", Rs3("CusID").value & " ," & column_location & " ," & Note_Value & "", "CustD", 0
            End If

            '          getCustomerAgeingData = getCustomerAgeingData & vbTab & CurrentString & Chr(13)
                        
            Rs3.MoveNext
        Next i
 
    End If

    Rs3.Close
 
    Dim StrSQL As String
 
    Dim Rs4    As New ADODB.Recordset
  
    StrSQL = "SELECT     TOP 100 PERCENT dbo.TblTempCustomerAging.CustD, SUM(dbo.TblTempCustomerAging.DueValue) AS DuevalueSum, dbo.Ageng_type.Name, CONVERT(varchar(5), "
    StrSQL = StrSQL & "   dbo.Ageng_type.[From]) + ' - ' + CONVERT(varchar(5), dbo.Ageng_type.[To]) AS DES"
    StrSQL = StrSQL & "  FROM         dbo.TblTempCustomerAging LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.Ageng_type ON dbo.TblTempCustomerAging.LateID = dbo.Ageng_type.id"
    StrSQL = StrSQL & "  GROUP BY dbo.TblTempCustomerAging.CustD, dbo.Ageng_type.Name, dbo.Ageng_type.[From], dbo.Ageng_type.[To], dbo.Ageng_type.id"
    StrSQL = StrSQL & "   ORDER BY dbo.Ageng_type.id"
    Debug.Print StrSQL
    CurrentString = ""
    getCustomerAgeingData = ""
    Rs4.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Rs4.RecordCount = 0 Then
        Exit Function
    End If
 
    If Rs4.RecordCount > 0 Then
      
        Rs4.MoveFirst
         
        For i = 1 To Rs4.RecordCount

            CurrentString = IIf(IsNull(Rs4.Fields("DuevalueSum").value), "", Rs4.Fields("DuevalueSum").value)

            CurrentString = padding(Trim(Format(val(CurrentString), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))), 20)
            getCustomerAgeingData = getCustomerAgeingData & "  " & CurrentString

            CurrentString = IIf(IsNull(Rs4.Fields("DES").value), "", Rs4.Fields("DES").value)

            CurrentString = padding(Trim(CurrentString), 20)
            getCustomerAgeingData = getCustomerAgeingData & "         " & CurrentString

            'CurrentString = IIf(IsNull(Rs4.Fields("Name").value), _
             "", Rs4.Fields("Name").value)

            'CurrentString = padding(Trim(CurrentString), 20)
            CurrentString = ""
            getCustomerAgeingData = getCustomerAgeingData & "        " & CurrentString & CHR(13)

            Rs4.MoveNext
        Next i

    End If

End Function

Public Sub ShowReport(Optional StrAccountCode As String, _
                      Optional StrAccountName As String, _
                      Optional FromDate As Date, _
                      Optional ToDate As Date, _
                      Optional manyString As Boolean = False, _
                      Optional BranchID As Integer = 0)
    Dim cAccountReport As ClsAccReports
    Set cAccountReport = New ClsAccReports
    Dim salesPersonName    As String
    Dim CustomerAgeingData As String
    cAccountReport.BegineDate = FromDate '
    cAccountReport.EndDate = ToDate
    
    If CheckUserNotPermAccounts(CDbl(user_id), StrAccountCode) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·« Ì„ﬂ‰ «·«ÿ·«⁄ ⁄·Ì Â–« «·Õ”«»", vbCritical
        Else
            MsgBox "can't view this acc statement", vbCritical
        End If
        Exit Sub
    End If
    Dim ShowAgingReport As String

    CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName, , True)
      
    If CustomerAgeingData <> "" Then
        If ViewAging = True Then
 
            If SystemOptions.UserInterface = ArabicInterface Then
                X = MsgBox("Â·  —Ìœ ⁄—÷ «⁄„«— «·œÌÊ‰ ›Ì Õ«·… «·⁄„·«¡ ", vbCritical + vbYesNoCancel)
            Else
                X = MsgBox("Show Aging Report ", vbCritical + vbYesNoCancel)
            End If

            If X = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
    
            If X = vbYes Then
                ShowAgingReport = "1"
                CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName)

            Else
                ShowAgingReport = 0
                CustomerAgeingData = getCustomerAgeingData(StrAccountCode, salesPersonName)

            End If

        End If

    End If
    
    updateopeningbalanceNewFromsql FromDate, ToDate, True, , BranchID, StrAccountCode, 3

    cAccountReport.ShowLedger StrAccountCode, StrAccountName, manyString, FromDate, ToDate, , BranchID, CustomerAgeingData, salesPersonName   'Current_branch
    Set cAccountReport = Nothing

End Sub

Public Function getLastDataBaseUpdateDate(Optional ByRef funId As Integer) As String
    On Error GoTo errorTrab
    Dim rs  As ADODB.Recordset
    Dim sql As String

    Set rs = New ADODB.Recordset
    sql = "Select * From Systemversion where id=1"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic

    If rs.RecordCount > 0 Then
        getLastDataBaseUpdateDate = IIf(IsNull(rs("version").value), 0, rs("version").value)
        funId = IIf(IsNull(rs("funId").value), 0, rs("funId").value)
    Else
        getLastDataBaseUpdateDate = ""
        funId = 0
    End If
    Exit Function
errorTrab:
    funId = 0
End Function

Public Function getoprTitle() As String

    If SystemOptions.UserInterface = ArabicInterface Then
        If SystemOptions.ProcessPeriodType = 0 Then
            getoprTitle = "ÌÊ„"
        ElseIf SystemOptions.ProcessPeriodType = 1 Then
            getoprTitle = "‘Â—"
        ElseIf SystemOptions.ProcessPeriodType = 2 Then
            getoprTitle = "”‰…"
        ElseIf SystemOptions.ProcessPeriodType = 3 Then
            getoprTitle = "«”»Ê⁄"
        End If

    Else

        If SystemOptions.ProcessPeriodType = 0 Then
            getoprTitle = "Days"
        ElseIf SystemOptions.ProcessPeriodType = 1 Then
            getoprTitle = "Months"
        ElseIf SystemOptions.ProcessPeriodType = 2 Then
            getoprTitle = "Years"
        ElseIf SystemOptions.ProcessPeriodType = 3 Then
            getoprTitle = "Weeks"
        End If
    
    End If

End Function

Public Function saveOperationDates(project_id As Integer, _
   project_startDate As Date)
    Dim rs             As ADODB.Recordset
    Dim sql            As String
    Dim i              As Integer
    Dim StartWeek      As Date
    Dim EndWeek        As Date
    Dim EarlyStartWeek As Date
    Dim EarlyEndWeek   As Date
    Dim X              As Double

    sql = "select * from terms_operations where ended=0 and project_id=" & project_id
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
 
            X = IIf(Not IsNumeric(rs("EarlyStartWeek").value), 0, rs("EarlyStartWeek").value)

            If SystemOptions.ProcessPeriodType = 0 Then

                EarlyStartWeek = DateAdd("d", X, project_startDate)    'day
            ElseIf SystemOptions.ProcessPeriodType = 1 Then
                EarlyStartWeek = DateAdd("m", X, project_startDate) 'Month
            ElseIf SystemOptions.ProcessPeriodType = 2 Then
                EarlyStartWeek = DateAdd("yyyy", X, project_startDate) 'Year
            ElseIf SystemOptions.ProcessPeriodType = 3 Then
                EarlyStartWeek = DateAdd("ww", X, project_startDate) 'week
            End If

            X = IIf(Not IsNumeric(rs("EarlyEndWeek").value), 0, rs("EarlyEndWeek").value)

            If SystemOptions.ProcessPeriodType = 0 Then

                EarlyEndWeek = DateAdd("d", X, project_startDate)    'day
            ElseIf SystemOptions.ProcessPeriodType = 1 Then
                EarlyEndWeek = DateAdd("m", X, project_startDate) 'Month
            ElseIf SystemOptions.ProcessPeriodType = 2 Then
                EarlyEndWeek = DateAdd("yyyy", X, project_startDate) 'Year
            ElseIf SystemOptions.ProcessPeriodType = 3 Then
                EarlyEndWeek = DateAdd("ww", X, project_startDate) 'week
            End If

            X = IIf(Not IsNumeric(rs("StartWeek").value), 0, rs("StartWeek").value)

            If SystemOptions.ProcessPeriodType = 0 Then
 
                StartWeek = DateAdd("d", X, project_startDate)    'day
            ElseIf SystemOptions.ProcessPeriodType = 1 Then
                StartWeek = DateAdd("m", X, project_startDate) 'Month
            ElseIf SystemOptions.ProcessPeriodType = 2 Then
                StartWeek = DateAdd("yyyy", X, project_startDate) 'Year
            ElseIf SystemOptions.ProcessPeriodType = 3 Then
                StartWeek = DateAdd("ww", X, project_startDate) 'week
            End If
 
            X = IIf(Not IsNumeric(rs("EndWeek").value), 0, rs("EndWeek").value)

            If SystemOptions.ProcessPeriodType = 0 Then

                EndWeek = DateAdd("d", X, project_startDate)    'day
            ElseIf SystemOptions.ProcessPeriodType = 1 Then
                EndWeek = DateAdd("m", X, project_startDate) 'Month
            ElseIf SystemOptions.ProcessPeriodType = 2 Then
                EndWeek = DateAdd("yyyy", X, project_startDate) 'Year
            ElseIf SystemOptions.ProcessPeriodType = 3 Then
                EndWeek = DateAdd("ww", X, project_startDate) 'week
            End If

            rs("start_date").value = StartWeek
            rs("end_date").value = EndWeek
            rs("EarlyStartDate").value = EarlyStartWeek
            rs("EarlyEndDate").value = EarlyEndWeek
            rs.update
            rs.MoveNext
        Next i

    End If

End Function

Public Function getEmployeeCash(EmpID As Integer)
    Dim sql          As String
    Dim rs           As ADODB.Recordset
    Dim i            As Integer
    Dim Balance      As Double
    Dim Account_code As String
    Balance = 0
 
    sql = "SELECT *    from TblBoxesData  WHERE     empid = " & EmpID

    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adcmtext

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            Account_code = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            Balance = Balance + get_balanceFromGl(Account_code)

            rs.MoveNext
        Next i

    End If

    getEmployeeCash = Balance
End Function

Public Function getEmployeeAdvance(EmpID As Integer)
    Dim sql As String
    Dim rs  As ADODB.Recordset
 
    sql = "SELECT     SUM(dbo.TblEmpAdvanceDetails.PartValue)  AS totalAdvance" & " FROM         dbo.TblEmpAdvance INNER JOIN " & "                       dbo.TblEmpAdvanceDetails ON dbo.TblEmpAdvance.AdvanceID = dbo.TblEmpAdvanceDetails.AdvanceID " & "  WHERE      dbo.TblEmpAdvanceDetails.Payed IS NULL and dbo.TblEmpAdvance.Emp_ID= " & EmpID

    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adcmtext

    If rs.RecordCount > 0 Then
        getEmployeeAdvance = IIf(IsNull(rs("totalAdvance").value), 0, rs("totalAdvance").value)
    End If

End Function

Public Function getProjectDuration(project_id As Integer, _
                                   Optional ByRef duration1 As Double, _
                                   Optional ByRef duration2 As Double) As Double
    Dim sql As String
    Dim rs  As ADODB.Recordset
 
    sql = " select max (EndWeek) as endweekmax from  terms_operations where project_id=" & project_id
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adcmtext

    If rs.RecordCount > 0 Then
        getProjectDuration = IIf(IsNull(rs("endweekmax").value), 0, rs("endweekmax").value)
    End If

End Function

'FixedAssetsFunctions
Public Function GetAndCalculateAll(FixedassetId As Integer, _
                                   Optional DepreciationPercentag As Double, _
                                   Optional ByRef noOfInstallments As Integer, _
                                   Optional ByRef Age As Integer, _
                                   Optional purchaseprice As Double, _
                                   Optional KhordaPrice As Double, _
                                   Optional AccDepreciation As Double, _
                                   Optional ByRef currentvalue As Double, _
                                   Optional ByRef installValue As Double, _
                                   Optional ByRef EXEInstallments As Double, _
                                   Optional ByRef RemainInstallments As Double, _
                                   Optional MinusValue As Double)
    On Error Resume Next

    If DepreciationPercentag = 0 Then
        noOfInstallments = 0
        AccDepreciation = 0
        KhordaPrice = 0
        currentvalue = purchaseprice
        noOfInstallments = 0
        installValue = 0
        EXEInstallments = 0
        MinusValue = 0
        Exit Function
    End If
 
    noOfInstallments = 100 / DepreciationPercentag * 12
    Age = noOfInstallments
    '    currentvalue = (purchaseprice - (AccDepreciation + KhordaPrice)) - MinusValue
    currentvalue = Round((purchaseprice - (AccDepreciation + 0)) - MinusValue, 2)
    installValue = Round((purchaseprice - KhordaPrice) / noOfInstallments, 2)
    EXEInstallments = Round(AccDepreciation / installValue, 0)
    RemainInstallments = noOfInstallments - EXEInstallments

End Function

Public Function getFixedAsstName(ID As Integer, _
   filed As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  " & filed & " from FixedAssets where id=" & ID
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        getFixedAsstName = ""
        code = ""
        Exit Function
    End If
    If IsNull(Rs3(filed).value) Then
        getFixedAsstName = ""
        code = ""
        Exit Function
    End If
    If Not IsNull(Rs3(filed).value) Then
        getFixedAsstName = Rs3(filed).value
        Exit Function
    End If
    Rs3.Close

End Function
 
Public Function GetAllDataAboutFixedAsset(FixedAsseId As Integer, _
                                          Optional ByRef Name As String, _
                                          Optional ByRef group_id As Integer, _
                                          Optional ByRef branch_no As Integer, _
                                          Optional ByRef Emp_id As Integer, _
                                          Optional ByRef ReceiveDate As Date, _
                                          Optional ByRef currentvalue As Double, _
                                          Optional ByRef AccDepreciation As Double, _
                                          Optional ByRef Status_id As Integer, _
                                          Optional ByRef Depreciation_Type_id As Integer, _
                                          Optional ByRef DefaultAge As Integer, _
                                          Optional ByRef StartDepreciationDate As Date, _
                                          Optional ByRef LastDepreciationDate As Date, _
                                          Optional ByRef noOfInstallments As Double, _
                                          Optional ByRef EXEInstallments As Double, _
                                          Optional ByRef RemainInstallments As Double, _
                                          Optional ByRef purchaseprice As Double, _
                                          Optional ByRef PurchaseDate As Date, _
                                          Optional ByRef PurchaseBillId As Double, _
                                          Optional ByRef KhordaPrice As Double, _
                                          Optional ByRef Installmentvalue As Double, _
                                          Optional ByRef New_or_opening As Integer, _
                                          Optional ByRef notes As String, _
                                          Optional ByRef fullcode As String, _
                                          Optional DepitAccount As String, Optional CreditAccount As String, Optional Account_Code5 As String, Optional ParetnAccount As String, Optional GroupName As String)
  
    Dim sql As String
    Dim rs  As ADODB.Recordset
    Set rs = New ADODB.Recordset
    '    sql = "Select * from  FixedAssets where id=" & FixedAsseId
    sql = "SELECT     dbo.FixedAssets.*,FixedAssets.fullcode as fafullcode, dbo.FixedAssetsGroup.*"
    sql = sql & " FROM         dbo.FixedAssets INNER JOIN"
    sql = sql & "  dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID"
    sql = sql & " WHERE     (dbo.FixedAssets.id = " & FixedAsseId & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Name = IIf(IsNull(rs("Name").value), "", rs("Name").value)
        group_id = IIf(IsNull(rs("group_id").value), 0, val(rs("group_id").value))
        branch_no = IIf(IsNull(rs("Branch_NO").value), 0, val(rs("Branch_NO").value))
        Emp_id = IIf(IsNull(rs("Emp_id").value), 0, (rs("Emp_id").value))
        ReceiveDate = IIf(IsNull(rs("ReceiveDate").value), Date, rs("ReceiveDate").value)
        currentvalue = IIf(IsNull(rs("CurrentValue").value), 0, val(rs("CurrentValue").value))
        DepitAccount = IIf(IsNull(rs("Account_Code1").value), "", (rs("Account_Code1").value))
        CreditAccount = IIf(IsNull(rs("Account_Code2").value), "", (rs("Account_Code2").value))
        Account_Code5 = IIf(IsNull(rs("Account_Code5").value), "", (rs("Account_Code5").value))
        ParetnAccount = IIf(IsNull(rs("ParetnAccount").value), "", (rs("ParetnAccount").value))
 
        GroupName = IIf(IsNull(rs("GroupName").value), "", (rs("GroupName").value))
        GroupNamee = IIf(IsNull(rs("GroupNamee").value), "", (rs("GroupNamee").value))
 
        AccDepreciation = IIf(IsNull(rs("AccDepreciation").value), 0, val(rs("AccDepreciation").value))

        Status_id = IIf(IsNull(rs("Status_id").value), 0, val(rs("Status_id").value))
        Depreciation_Type_id = IIf(IsNull(rs("Depreciation_Type_id").value), 0, val(rs("Depreciation_Type_id").value))
        DefaultAge = IIf(IsNull(rs("DefaultAge").value), 0, val(rs("DefaultAge").value))
        StartDepreciationDate = IIf(IsNull(rs("StartDepreciationDate").value), Date, rs("StartDepreciationDate").value)
        LastDepreciationDate = IIf(IsNull(rs("LastDepreciationDate").value), Date, rs("LastDepreciationDate").value)
        noOfInstallments = IIf(IsNull(rs("NoOfInstallments").value), 0, val(rs("NoOfInstallments").value))
        EXEInstallments = IIf(IsNull(rs("EXEInstallments").value), 0, val(rs("EXEInstallments").value))
        RemainInstallments = IIf(IsNull(rs("RemainInstallments").value), 0, val(rs("RemainInstallments").value))
        purchaseprice = IIf(IsNull(rs("PurchasePrice").value), 0, val(rs("PurchasePrice").value))
        PurchaseDate = IIf(IsNull(rs("PurchaseDate").value), Date, (rs("PurchaseDate").value))
        PurchaseBillId = IIf(IsNull(rs("PurchaseBillId").value), 0, val(rs("PurchaseBillId").value))
        KhordaPrice = IIf(IsNull(rs("KhordaPrice").value), 0, val(rs("KhordaPrice").value))
        Installmentvalue = IIf(IsNull(rs("InstallmentValue").value), 0, val(rs("InstallmentValue").value))
        New_or_opening = IIf(IsNull(rs("New_or_opening").value), 0, val(rs("New_or_opening").value))
        fullcode = IIf(IsNull(rs("fafullcode").value), "", rs("fafullcode").value)
        notes = IIf(IsNull(rs("Notes").value), "", rs("Notes").value)
    End If

End Function
Public Function AddInstallment(FixedAssetInstallmentsid As Integer, _
                               FixedAsseId As Integer, _
                               Optional currentvalue As Double, _
                               Optional InstallmentID As Integer, _
                               Optional InstallmentDate As Date, _
                               Optional ByRef AccDepreciation As Double, _
                               Optional ByRef Installmentvalue As Double, _
                               Optional ByRef RemainInstallments As Double, _
                               Optional ByRef InstallmentProduct As Integer)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "FixedAssetInstallmentsDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
    rs("FixedAssetInstallmentsid").value = FixedAssetInstallmentsid
    rs("FixedAssetID").value = FixedAsseId
    rs("CurrentValue").value = currentvalue
    rs("InstallmentID").value = InstallmentID
    rs("InstallmentValue").value = Installmentvalue
    rs("InstallmentDate").value = InstallmentDate
    rs("AccDepreciation").value = AccDepreciation
    rs("RemainInstallments").value = RemainInstallments
    rs("month").value = Month(InstallmentDate)
    rs("year").value = year(InstallmentDate)
    rs("InstallmentProduct").value = InstallmentProduct

    rs.update
 
End Function

Public Function updateFixedAsseTInstallmentInformations(FixedAsseId As Integer, _
                                                        Optional PurcahsePrice As Double, _
                                                        Optional currentvalue As Double, _
                                                        Optional ByRef InstallmentID As Integer, _
                                                        Optional ByRef InstallmentDate As Date, _
                                                        Optional ByRef AccDepreciation As Double, _
                                                        Optional ByRef Installmentvalue As Double, _
                                                        Optional ByRef RemainInstallments As Double, _
                                                        Optional NewAsset As Boolean, _
                                                        Optional First_Installment As Boolean)
    Dim KhordaPrice        As Double
    Dim noOfInstallments   As Double
    Dim delsql             As String
    Dim InstallmentProduct As Integer

    If NewAsset = True And First_Installment = True Then    'ÃœÌœ «Ê· ﬁ”ÿ
        delsql = "Delete FixedAssetInstallmentsDetails where FixedAssetID=" & FixedAsseId & "and InstallmentID=0"
        Cn.Execute delsql
        GetAllDataAboutFixedAsset FixedAsseId, , , , , , currentvalue, AccDepreciation, , , , , , noOfInstallments, , RemainInstallments, PurcahsePrice, , , KhordaPrice, Installmentvalue
        InstallmentID = 0
        InstallmentProduct = 0
        AccDepreciation = 0
        RemainInstallments = noOfInstallments
        Installmentvalue = Round((PurcahsePrice - KhordaPrice) / noOfInstallments, 2)
        Installmentvalue = 0
        AddInstallment 0, FixedAsseId, currentvalue, InstallmentID, InstallmentDate, AccDepreciation, Installmentvalue, RemainInstallments, InstallmentProduct
    ElseIf NewAsset = True And First_Installment = False Then 'ÃœÌœ Ê·Ì” «Ê· ﬁ”ÿ
   
    ElseIf NewAsset = False And First_Installment = True Then ' «›  «ÕÌ «Ê· ﬁ”ÿ
        delsql = "Delete FixedAssetInstallmentsDetails where FixedAssetID=" & FixedAsseId & "and InstallmentID=0"
        Cn.Execute delsql
        GetAllDataAboutFixedAsset FixedAsseId, , , , , , currentvalue, AccDepreciation, , , , , InstallmentDate, noOfInstallments, , RemainInstallments, PurcahsePrice, , , KhordaPrice, Installmentvalue
        InstallmentID = 0
        Installmentvalue = AccDepreciation
        InstallmentProduct = noOfInstallments - RemainInstallments
        AddInstallment 0, FixedAsseId, currentvalue, InstallmentID, InstallmentDate, AccDepreciation, Installmentvalue, RemainInstallments, InstallmentProduct
    ElseIf NewAsset = False And First_Installment = False Then ' «›  «ÕÌÊ·Ì”  «Ê· ﬁ”ÿ
    
    End If
   
End Function

Public Function CheCkInstallmentCount(FixedassetId As Integer) As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  count (FixedAssetID ) As InstallmentCount from FixedAssetInstallmentsDetails where FixedAssetID=" & FixedassetId
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        CheCkInstallmentCount = 0
        Exit Function
    End If
    If IsNull(Rs3("InstallmentCount").value) Then
        CheCkInstallmentCount = 0
        Exit Function
    End If
    If Not IsNull(Rs3("InstallmentCount").value) Then
        CheCkInstallmentCount = Rs3("InstallmentCount").value - 1
        Exit Function
    End If
    Rs3.Close

End Function

Public Function CheckLastInstallmentDate(Month As Integer, _
                                         year As Integer, _
                                         Optional BranchID As Integer) As Boolean
    CheckLastInstallmentDate = False
    Dim sql As String
    Dim rs  As ADODB.Recordset

    '  sql = "Select max(Month) As LastMonth  From FixedAssetInstallments where year =" & year
    sql = "Select max(Month) As LastMonth  From FixedAssetInstallments where year(RecordDate) =" & year
    
    '
    sql = sql & "  and BranchId=" & BranchID
    
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        CheckLastInstallmentDate = True
    ElseIf IsNull(rs("LastMonth").value) Then
        CheckLastInstallmentDate = True
    ElseIf Month - rs("LastMonth").value > 1 Then
        CheckLastInstallmentDate = False
    ElseIf Month - rs("LastMonth").value = 1 Then
        CheckLastInstallmentDate = True
    ElseIf rs("LastMonth").value - Month >= 0 Then
        CheckLastInstallmentDate = False

        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ «’œ«— «ﬁ”«ÿ ·Â–« «·‘Â—  ·«‰Â „’œ—„‰ ﬁ»·"
            MsgBox Msg, vbInformation
        Else
            Msg = "Cant Create Depreciation Installment For this Month , already Created"
            MsgBox Msg, vbInformation
        End If
        Exit Function
    End If

    If CheckLastInstallmentDate = False Then

        'CboYear.ListIndex = -1
        'CmbMonth.ListIndex = -1
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ «’œ«— «ﬁ”«ÿ ·Â–« «·‘Â— ÌÊÃœ «ﬁ”«ÿ ”«»ﬁ… €Ì— „’œ—…"
            MsgBox Msg, vbInformation
        Else
            Msg = "Cant Create Depreciation Installment For this Month , Check Last  Installment Date"
            MsgBox Msg, vbInformation
        End If

    End If

End Function

Public Function GetFixedAssetHistory(FixedassetId As Integer, _
                                     Optional ByRef AccDepreciation As Double, _
                                     Optional ByRef RemianInstallments As Double, _
                                     Optional ByRef CurrentInstalmentNo As Double, _
                                     Optional ByRef Installmentvalue As Double, _
                                     Optional ByRef NewAccDepreciation As Double, _
                                     Optional ByRef purchaseprice As Double, _
                                     Optional ByRef FixedAssetName As String, _
                                     Optional ByRef currentvalue As Double, _
                                     Optional ByRef fullcode As String, _
                                     Optional ByRef KhordaPrice As Double, _
                                     Optional ByRef group_id As Integer, _
                                     Optional DepitAccount As String, _
                                     Optional CreditAccount As String, _
                                     Optional ByRef PurchaseDate As Date, _
                                     Optional branch_no As Integer)
    Dim noOfInstallments As Double
    Dim sql              As String
    Dim EXEInstallments  As Double
 
    Dim rs               As ADODB.Recordset
    sql = "Select sum(InstallmentProduct)as EXEInstallments ,sum(InstallmentValue)  as AccDepreciation From FixedAssetInstallmentsDetails where   FixedAssetID =" & FixedassetId    '«·Õ’Ê· ⁄·Ï ⁄œœ «·«ﬁ”«ÿ «·„‰›–…
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    GetAllDataAboutFixedAsset FixedassetId, FixedAssetName, group_id, branch_no, , , currentvalue, , , , , , , noOfInstallments, , , purchaseprice, PurchaseDate, , KhordaPrice, Installmentvalue, , , fullcode, DepitAccount, CreditAccount

    If rs.RecordCount > 0 Then
        EXEInstallments = IIf(IsNull(rs("EXEInstallments").value), 0, rs("EXEInstallments").value)
        RemianInstallments = noOfInstallments - EXEInstallments
        AccDepreciation = IIf(IsNull(rs("AccDepreciation").value), 0, rs("AccDepreciation").value)

        If RemianInstallments > 0 Or purchaseprice > AccDepreciation Then
            RemianInstallments = RemianInstallments
            CurrentInstalmentNo = EXEInstallments
            NewAccDepreciation = AccDepreciation ' + Installmentvalue
        
        Else
            CurrentInstalmentNo = -1 '«‰ Â«¡ «·«ﬁ”«ÿ
        End If
    End If

End Function
      
Public Function getinsttPayedTocontract2(Optional ContNo As Double = 0, _
                                         Optional ByRef RentValuePayed As Double, _
                                         Optional ByRef CommissionsPayed As Double, _
                                         Optional ByRef InsurancePayed As Double, _
                                         Optional ByRef WaterPayed As Double, _
                                         Optional ByRef ElectricPayed As Double, _
                                         Optional ByRef TelandNetPayed As Double, _
                                         Optional ByRef TotalOldValue As Double, _
                                         Optional NoteID As Double, _
                                         Optional Typ As Integer = 0, _
                                         Optional ByRef VATPayed As Double, _
                                         Optional ByVal mData As String = "", _
                                         Optional ByVal DateTypeHij As Boolean = False) As Double
    On Error Resume Next

    Dim total As Single

    Dim Rs3   As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset
    Dim sql        As String
     
    '       sql = " SELECT Sum(IsNull(dbo.TblContractInstallments.RentValue,0) - IsNull(T2.RentValuePayed,0)) RentValueRemain,"
    '       sql = sql & " Sum(IsNull(dbo.TblContractInstallments.Commissions,0) - IsNull(T2.CommissionsPayed,0) )CommissionsRemain,"
    '       'Sql = Sql & "Sum(IsNull(dbo.TblContractInstallments.Insurance,0) - IsNull(t2.InsurancePayed,0) )InsuranceRemain,"
    '       sql = sql & "Sum(IsNull(t2.InsurancePayed,0) )InsuranceRemain,"
    '       sql = sql & "Sum(IsNull(dbo.TblContractInstallments.Water,0) - IsNull(T2.WaterPayed,0) ) WaterRemain,"
    '       sql = sql & "Sum(IsNull(dbo.TblContractInstallments.Electric,0) -IsNull( T2.ElectricPayed,0) )ElectricRemain,"
    '       sql = sql & "Sum(IsNull(dbo.TblContractInstallments.TelandNet,0) - IsNull(T2.TelandNetPayed,0) )TelandNetRemain   ,           "
    '       sql = sql & "Sum(IsNull(dbo.TblContractInstallments.VATValue,0) - IsNull(T2.VATPayed,0) )VatRemain ,"
    '       sql = sql & "        Sum(IsNull(dbo.TblContractInstallments.OldValue,0) - IsNull(T2.OldValuePayed,0))OldValueRemain "
    '
    '
    '
    '       sql = sql & "From "
    '       sql = sql & " dbo.TblContractInstallments"
    '
    '        sql = sql & " LEFT OUTER JOIN"
    '
    '       sql = sql & " ContracttBillInstallmentsDone T2 ON T2.istallid =TblContractInstallments.ID"
    '       sql = sql & " WHERE  "
    '       '("
    '       'sql = sql & "           ("
    '       'sql = sql & "               dbo.TblContractInstallments.Status IS NULL"
    '       'sql = sql & "               OR dbo.TblContractInstallments.Status = 0"
    '       'sql = sql & "           )"
    '
    '       'sql = sql & ")"
    '        If mData <> "" Then
    '            If Not DateTypeHij Then
    '                sql = sql & " (TblContractInstallments.Installdate <= " & SQLDate(CDate(mData), True) & ") and "
    '            Else
    '                sql = sql & " (TblContractInstallments.InstalldateH <= '" & (mData) & "') and "
    '            End If
    '        End If
    '        sql = sql & " contNo=" & ContNo
  
    '--------------------
 
    Dim RentValue2 As Double, Commissions2 As Double, Water2 As Double, Electric2 As Double, TelandNet2 As Double, VATValue2 As Double, OldValue2 As Double, InsuranceRemain2 As Double
    '-----------
    sql = " SELECT Sum( IsNull(T2.RentValuePayed,0)) RentValue2,"
    sql = sql & " Sum(IsNull(T2.CommissionsPayed,0) )Commissions2,"
    'Sql = Sql & "Sum(IsNull(dbo.TblContractInstallments.Insurance,0) - IsNull(t2.InsurancePayed,0) )InsuranceRemain,"
    sql = sql & "Sum(IsNull(t2.InsurancePayed,0) )InsuranceRemain2,"
    sql = sql & "Sum(IsNull(T2.WaterPayed,0) ) Water2,"
    sql = sql & "Sum(IsNull( T2.ElectricPayed,0) )Electric2,"
    sql = sql & "Sum(IsNull(T2.TelandNetPayed,0) )TelandNet2   ,           "
    sql = sql & "Sum(IsNull(T2.VATPayed,0) )VATValue2 ,"
    sql = sql & "        Sum(IsNull(T2.OldValuePayed,0))OldValue2 "
       
    sql = sql & "From ContracttBillInstallmentsDone T2 Where T2.istallid In ( "
    sql = sql & " Select TblContractInstallments.ID from  dbo.TblContractInstallments"
       
    sql = sql & " WHERE  "
    '("
    'sql = sql & "           ("
    'sql = sql & "               dbo.TblContractInstallments.Status IS NULL"
    'sql = sql & "               OR dbo.TblContractInstallments.Status = 0"
    'sql = sql & "           )"
           
    'sql = sql & ")"
    If mData <> "" Then
        If Not DateTypeHij Then
            sql = sql & " (TblContractInstallments.Installdate <= " & SQLDate(CDate(mData), True) & ") and "
        Else
            sql = sql & " (TblContractInstallments.InstalldateH <= '" & (mData) & "') and "
        End If
    End If
    sql = sql & " contNo=" & ContNo & " )"
   
    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not Rs4.EOF Then
            
        RentValue2 = IIf(IsNull(Rs4("RentValue2").value), 0, Rs4("RentValue2").value)
        Commissions2 = IIf(IsNull(Rs4("Commissions2").value), 0, Rs4("Commissions2").value)
        Water2 = IIf(IsNull(Rs4("Water2").value), 0, Rs4("Water2").value)
        Electric2 = IIf(IsNull(Rs4("Electric2").value), 0, Rs4("Electric2").value)
        TelandNet2 = IIf(IsNull(Rs4("TelandNet2").value), 0, Rs4("TelandNet2").value)
        VATValue2 = IIf(IsNull(Rs4("VATValue2").value), 0, Rs4("VATValue2").value)
        OldValue2 = IIf(IsNull(Rs4("OldValue2").value), 0, Rs4("OldValue2").value)
        InsuranceRemain2 = IIf(IsNull(Rs4("InsuranceRemain2").value), 0, Rs4("InsuranceRemain2").value)
    End If
    Rs4.Close
    sql = " SELECT "
    sql = sql & "Sum(IsNull(t2.InsurancePayed,0) )InsuranceRemain2 "
    sql = sql & "From ContracttBillInstallmentsDone T2 Where T2.istallid In ( "
    sql = sql & " Select TblContractInstallments.ID from  dbo.TblContractInstallments"
       
    sql = sql & " WHERE  "
 
    sql = sql & " contNo=" & ContNo & " )"
   
    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not Rs4.EOF Then
        InsuranceRemain2 = IIf(IsNull(Rs4("InsuranceRemain2").value), 0, Rs4("InsuranceRemain2").value)
    End If
    Dim RentValue As Double, Commissions As Double, Water As Double, Electric As Double, TelandNet As Double, VATValue As Double, OldValue As Double
    sql = " SELECT Sum(IsNull(dbo.TblContractInstallments.RentValue,0)) RentValue,"
    sql = sql & " Sum(IsNull(dbo.TblContractInstallments.Commissions,0)) Commissions,"
    'Sql = Sql & "Sum(IsNull(dbo.TblContractInstallments.Insurance,0) - IsNull(t2.InsurancePayed,0) )InsuranceRemain,"
       
    sql = sql & "Sum(IsNull(dbo.TblContractInstallments.Water,0)) Water ,"
    sql = sql & "Sum(IsNull(dbo.TblContractInstallments.Electric,0)) Electric ,"
    sql = sql & "Sum(IsNull(dbo.TblContractInstallments.TelandNet,0)) TelandNet,           "
    sql = sql & "Sum(IsNull(dbo.TblContractInstallments.VATValue,0)) VATValue ,"
    sql = sql & "        Sum(IsNull(dbo.TblContractInstallments.OldValue,0)) OldValue "
       
    sql = sql & "From "
    sql = sql & " dbo.TblContractInstallments"
    
    sql = sql & " WHERE  "
    '("
    'sql = sql & "           ("
    'sql = sql & "               dbo.TblContractInstallments.Status IS NULL"
    'sql = sql & "               OR dbo.TblContractInstallments.Status = 0"
    'sql = sql & "           )"
           
    'sql = sql & ")"
    If mData <> "" Then
        If Not DateTypeHij Then
            sql = sql & " (TblContractInstallments.Installdate <= " & SQLDate(CDate(mData), True) & ") and "
        Else
            sql = sql & " (TblContractInstallments.InstalldateH <= '" & (mData) & "') and "
        End If
    End If
    sql = sql & " contNo=" & ContNo
    
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        total = 0
        RentValuePayed = 0
        CommissionsPayed = 0
        InsurancePayed = 0
        WaterPayed = 0
        ElectricPayed = 0
        TelandNetPayed = 0
        TotalOldValue = 0
        VATPayed = 0
  
    Else

        RentValue = IIf(IsNull(Rs3("RentValue").value), 0, Rs3("RentValue").value)
        Commissions = IIf(IsNull(Rs3("Commissions").value), 0, Rs3("Commissions").value)
        Water = IIf(IsNull(Rs3("Water").value), 0, Rs3("Water").value)
        Electric = IIf(IsNull(Rs3("Electric").value), 0, Rs3("Electric").value)
        TelandNet = IIf(IsNull(Rs3("TelandNet").value), 0, Rs3("TelandNet").value)
        VATValue = IIf(IsNull(Rs3("VATValue").value), 0, Rs3("VATValue").value)
        OldValue = IIf(IsNull(Rs3("OldValue").value), 0, Rs3("OldValue").value)
        
        ' Total = IIf(IsNull(Rs3("total").value), 0, Rs3("total").value)
        VATPayed = VATValue - VATValue2
        RentValuePayed = RentValue - RentValue2
        CommissionsPayed = Commissions - Commissions2
        InsurancePayed = InsuranceRemain2
        WaterPayed = Water - Water2
        ElectricPayed = Electric - Electric2
        TelandNetPayed = TelandNet - TelandNet2
        TotalOldValue = OldValue - OldValue2
 
    End If

    Rs3.Close
    Rs4.Close
    getinsttPayedTocontract2 = total
End Function

Public Function Voucher_codingByBreaks(my_branch As Integer, _
                                       date1 As Date, _
                                       Sanad_No As Integer, _
                                       NoteType As Integer, _
                                       Optional departement_name As Integer = 1, _
                                       Optional Transaction_Type As Integer = 0, _
                                       Optional Prefix As String = "", _
                                       Optional StoreID As Integer = 0, _
                                       Optional BillType As Integer = 0, _
                                       Optional MosemID As Double = 0, _
                                       Optional ByVal mTableName As String = "", _
                                       Optional ByVal mUserId As Long = 0, _
                                       Optional ByRef mSerInv As Long = 0) As String
    
    On Error Resume Next
    If my_branch = 0 Then
        Exit Function
    End If
    Dim start_at       As Double
    Dim end_at         As Double
    Dim auto_sanad_no  As String
    Dim NO             As Double
    Dim numbering_type As Integer
    Dim noOfDigit      As Double
    Dim Zeros          As Double
    Dim StoreCoding    As Double
    Dim IsBreaks       As Boolean
    Dim IsCodeByBranch As Boolean
    Dim IsSerialByUser As Boolean
    Dim Breaks         As String
    
    Dim YearDigit      As Double
    Dim branchpadidng  As Integer
    Dim storepadding   As Integer
    If mUserId = 0 Then mUserId = user_id
    Dim mNoOfUser     As Integer
    Dim mUserIdSerial As String
    Dim mLenUser      As Integer

    Dim mFormatUser   As String

    mFormatUser = ""
    Dim mm As Integer
    mm = 1
    For mm = 1 To SystemOptions.NoOFDigitUserTrans
        If mUserId < 9 And SystemOptions.NoOFDigitUserTrans > 1 And SystemOptions.NoOFDigitUserTrans > mm Then
            mFormatUser = mFormatUser & "0"
        ElseIf mUserId > 9 And mUserId < 100 And SystemOptions.NoOFDigitUserTrans > 1 And SystemOptions.NoOFDigitUserTrans > mm + 1 Then
            mFormatUser = mFormatUser & "0"
        ElseIf mUserId > 99 And SystemOptions.NoOFDigitUserTrans > 1 And SystemOptions.NoOFDigitUserTrans > mm + 2 Then
            mFormatUser = mFormatUser & "0"
        End If

    Next

    'If mUserId > 9 And mUserId < 100 Then
    mUserIdSerial = mFormatUser & mUserId
    'End If
    mUserIdSerial = mUserId

    auto_sanad_no = ""
 
    Dim first_serial As Boolean
    Dim rs           As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim mWhereBranchYear As String
    Dim sql              As String
    Dim i                As Integer
    Dim storecode        As String

    first_serial = False
    sql = sql & "SELECT ISNULL(numbering_id, 0) numbering_id, "
    sql = sql & "       ISNULL(start_at, 0) start_at, "
    sql = sql & "       ISNULL(end_at, 0) end_at, "
    sql = sql & "       ISNULL(no_of_digit, 0) no_of_digit, "
    sql = sql & "       ISNULL(zeros, 0) zeros, "
    sql = sql & "       ISNULL(StoreCoding, 0) StoreCoding, "
    sql = sql & "       ISNULL(YearDigit, 0) YearDigit, "
    sql = sql & "       ISNULL(IsBreaks, 0) IsBreaks, "
    sql = sql & "       ISNULL(IsCodeByBranch, 0) IsCodeByBranch, "
    sql = sql & "       ISNULL(IsSerialByUser, 0) IsSerialByUser, "
    sql = sql & "       ISNULL(Breaks, 0) Breaks "
    sql = sql & "   from sanad_numbering where branch_no=" & my_branch & " and  sanad_no=" & Sanad_No
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly   ', adCmdText
  
    If rs.EOF Then
        numbering_type = 0
                
    Else

        numbering_type = rs!numbering_id 'IIf(IsNull(rs("numbering_id").value), 0, rs("numbering_id").value)
        start_at = rs!start_at 'IIf(IsNull(rs("start_at").value), 0, rs("start_at").value)
        end_at = rs!end_at 'IIf(IsNull(rs("end_at").value), 0, rs("end_at").value)
        noOfDigit = rs!no_of_digit 'IIf(IsNull(rs("no_of_digit").value), 0, rs("no_of_digit").value)
        If noOfDigit = 0 Then noOfDigit = 3
        Zeros = rs!Zeros 'IIf(IsNull(rs("zeros").value), 0, rs("zeros").value)
        StoreCoding = rs!StoreCoding 'IIf(IsNull(rs("StoreCoding").value), 0, rs("StoreCoding").value)
        YearDigit = rs!YearDigit 'IIf(IsNull(rs("YearDigit").value), 4, rs("YearDigit").value)
        IsBreaks = rs!IsBreaks 'IIf(IsNull(rs("IsBreaks").value), 0, rs("IsBreaks").value)
        IsCodeByBranch = rs!IsCodeByBranch 'IIf(IsNull(rs("IsCodeByBranch").value), 0, rs("IsCodeByBranch").value)
        IsSerialByUser = rs!IsSerialByUser 'IIf(IsNull(rs("IsSerialByUser").value), 0, rs("IsSerialByUser").value)
        Breaks = rs!Breaks 'IIf(IsNull(rs("Breaks").value), "", rs("Breaks").value)
        
        storepadding = SystemOptions.StoreDigit - 1
        If YearDigit = 0 Then YearDigit = 4
        
        branchpadidng = SystemOptions.BranchDigit - 1

        If StoreCoding = True Then
            If StoreID <> 0 Then
                storecode = getStoreCoding(StoreID)
            End If
        End If
        
    End If
    
    Dim mWhere4     As String
    Dim mWhereUser  As String
    Dim mWhereUser2 As String
    
    If IsSerialByUser Then
       
        ' mWhere4 = "SUBSTRING(CAST(cast(NoteSerial1 AS BIGINT) AS VARCHAR(100)),2 , " & SystemOptions.NoOFDigitUserTrans & ") = " & mUserID
        'mWhereUser = mWhereUser & " AND " & mUserID & "  IN ("
        'mWhereUser = mWhereUser & " SELECT UserID FROM DOUBLE_ENTREY_VOUCHERS AS dev WHERE dev.Notes_ID = Notes.NoteID)"
        mWhereUser2 = " And UserID = " & mUserId

    Else
        mWhereUser2 = ""
        mUserIdSerial = ""
        ' mWhere4 = mWhere4 & " and UserID = " & mUserId
    End If
    mWhereUser = mWhereUser2
    If IsCodeByBranch Then
        If Transaction_Type <> 0 Then
            mWhere3 = " BranchId =  " & my_branch
        Else
            mWhere3 = " Branch_Id =  " & my_branch
        End If
    Else
        mWhere3 = " "
    End If
   
    '    If my_branch > 9 Then
    '        mWhere3 = " SUBSTRING(CAST(cast(NoteSerial1 AS BIGINT) AS VARCHAR(50))," & SystemOptions.NoOFDigitUserTrans + 2 & ", 2) = " & my_branch
    '    Else
    '        mWhere3 = " SUBSTRING(CAST(cast(NoteSerial1 AS BIGINT) AS VARCHAR(50)), " & SystemOptions.NoOFDigitUserTrans + 2 & ", 1) = " & my_branch
    '    End If
    ' mWhere3 = " 1 = 1 "
    If numbering_type = 1 Then ' «·Ì
        mWhereBranchYear = " Year = " & year(date1) & " "
        
        mWhereBranchYear = " Year(NoteDate) = " & year(date1)

        mWhereUser2 = mWhereUser2 & " And " & mWhereBranchYear
        sql = "select max(Ser) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and   NoteType=" & NoteType ' & " and   numbering_type1=" & numbering_type
        sql = sql & mWhereUser
        
        If Sanad_No = 5 Then
            sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and not(BTCashAccountcode is null )"
            sql = sql & mWhereUser2 'mWhereUser
            
        End If
   
        If Sanad_No = 1 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1)
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
        End If
   
        If Sanad_No = 4 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1)
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            Else
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            End If
            sql = sql & mWhereUser
        End If
   
        ' ﬂÊÌœ ”‰œ «· À”Ìÿ ‰›” ”‰œ «·ﬁ»÷
        If Sanad_No = 25 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1)
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 18) " '  and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
            
        End If
   
        If Sanad_No = 2 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1)
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 18 )" ' And numbering_type1 = " & numbering_type"
            Else
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType  ' & " and   numbering_type1=" & numbering_type
            End If
            sql = sql & mWhereUser
        End If
   
        If Sanad_No = 26 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1)
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
        End If
   
        If Sanad_No = 2 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1)
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
            Else
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type
            End If
            sql = sql & mWhereUser
        End If

        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ «·’—›
   
        If Sanad_No = 1 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1)
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(Ser) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 14) " ' and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
        End If
   
        If Sanad_No = 16 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1)
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 3 or NoteType= 14 ) " ' and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            Else
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType  '& " and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            End If
            sql = sql & mWhereUser
        End If
      
        If Sanad_No = 50 Then
            mWhereBranchYear = " Year(RecordDate) = " & year(date1)
            sql = "select max(Ser) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            sql = sql & mWhereUser2
        End If
        
        If Sanad_No = 58 Then
            mWhereBranchYear = " Year(RecordDate) = " & year(date1)
            sql = "select max(Ser) as last_sand_no from  TblExchange    where  BranchID= " & my_branch
            sql = sql & mWhereUser2
        End If
        
        If Sanad_No = 60 Then
            mWhereBranchYear = " Year(ContDate) = " & year(date1)
            sql = "select max(Ser) as last_sand_no from  TblContract    where  Branch_NO= " & my_branch
            sql = sql & mWhereUser2
        End If
        
        If Sanad_No = 62 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1)
            sql = "select max(Ser) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= " & NoteType & ") "
            sql = sql & mWhereUser
        End If
        
        If Sanad_No = 64 Then
            sql = "select max(CAST (Ser AS BigInt))as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   "
            sql = sql & mWhereUser2
        End If
        
        If Sanad_No = 65 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1 "
                Else
                    sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0 "
                End If
            Else
            
                sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "   "
            End If
        End If
        If Sanad_No = 66 Then
            mWhereBranchYear = " Year(Transaction_Date) = " & year(date1)
            sql = "select max(CAST (Ser AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  "
            sql = sql & "  and ( Transaction_Type=990 or Transaction_Type=18)"
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 67 Then
            mWhereBranchYear = " Year(Transaction_Date) = " & year(date1)
            sql = "select max(CAST (Ser AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  "
            sql = sql & "  and  (Transaction_Type=66 or Transaction_Type=991) "
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 68 Then
            mWhereBranchYear = " Year(recorddate) = " & year(date1)
            sql = "select max(CAST (Ser AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  "
            sql = sql & "  and  (ImportExport=0 ) "
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 69 Then
            mWhereBranchYear = " Year(recorddate) = " & year(date1)
            sql = "select max(CAST (Ser AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  "
            sql = sql & "  and  (ImportExport=1 ) "
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 70 Then
            mWhereBranchYear = " Year(SDate) = " & year(date1)
            sql = "select max (Ser )  as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  "
            sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 71 Then
            mWhereBranchYear = " Year(SDate) = " & year(date1)
            sql = "select max (Ser )  as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  "
            sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
            sql = sql & mWhereUser2
        End If
        
        If Sanad_No = 72 Then
            mWhereBranchYear = " Year(RecordDate) = " & year(date1)
        
            sql = "select max (Ser )  as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  "
            sql = sql & "  and  (SeasonsID=" & MosemID & " ) "
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 74 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1)
            sql = "select max (Ser )  as last_sand_no from  notes_all where  branch_no= " & my_branch & "  and  notetype=370 "
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 75 Then
            '  Sql = "select max(CAST (dbo.TblContractInstallments.Ser AS BigInt))   as last_sand_no from  TblContractInstallments where  branch_no= " & my_branch & "  "
            sql = " SELECT     max(CAST (dbo.TblContractInstallments.Ser AS BigInt)) as last_sand_no "
            sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
            sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
            sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")"
        End If
        If Sanad_No = 76 Then
            sql = "select max (Ser )  as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "  "
            sql = sql & " and " & mWhere3
        End If
        If Transaction_Type <> 0 Then
            mWhereBranchYear = " Year(Transaction_Date) = " & year(date1)
            sql = "select max(CAST (Ser AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
        If Transaction_Type <> 0 Then
            mWhereBranchYear = " Year(Transaction_Date) = " & year(date1)
            sql = "select max(CAST (Ser AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
        
        If Prefix = "" Then
            sql = sql & "  and   Prefix is null"
 
        Else
            sql = sql & "  and   Prefix='" & Prefix & "'"
        End If
        sql = sql & " and " & mWhere3
        sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        sql = sql & " and " & IIf(mWhereBranchYear = "", " 1 = 1 ", mWhereBranchYear)
  
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly   ', adCmdText

        If Not IsNull(Rs3("last_sand_no").value) Then
            If end_at = 0 Then end_at = val(Rs3("last_sand_no").value) + 1
            If Rs3("last_sand_no").value >= end_at Then
                Voucher_codingByBreaks = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 2 Then ' „ ’· ‘Â—Ì
        
        mWhereBranchYear = " Year(Transaction_Date) = " & year(date1) & " and Month(Transaction_Date) =  " & Month(date1)
        mWhereBranchYear = " Year(NoteDate) = " & year(date1) & " and Month(NoteDate) =  " & Month(date1)
        ' mWhereUser2 = mWhereUser2 & " And " & mWhereBranchYear
        
        sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
        sql = sql & " and " & mWhere3
        sql = sql & mWhereUser
        If Sanad_No = 5 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1) & " and Month(NoteDate) =  " & Month(date1)
            sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and not(BTCashAccountcode is null )"
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser
        End If
    
        If Sanad_No = 1 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1) & " and Month(NoteDate) =  " & Month(date1)
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        End If
    
        If Sanad_No = 4 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1) & " and Month(NoteDate) =  " & Month(date1)
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            Else
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        
        End If
    
        If Sanad_No = 25 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1) & " and Month(NoteDate) =  " & Month(date1)
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        End If
    
        If Sanad_No = 2 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1) & " and Month(NoteDate) =  " & Month(date1)
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            Else
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "       and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1) & " and Month(NoteDate) =  " & Month(date1)
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        End If
    
        If Sanad_No = 16 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1) & " and Month(NoteDate) =  " & Month(date1)
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)     and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            Else
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        
        End If
        
        If Sanad_No = 50 Then
            mWhereBranchYear = " Year(recorddate) = " & year(date1) & " and Month(recorddate) =  " & Month(date1)
            '        sql = "select max(Ser) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            sql = "select max(Ser) as last_sand_no from  TblCarBillMentains where  BranchID= " & my_branch & " and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
        End If
        If Sanad_No = 58 Then
            mWhereBranchYear = " Year(recorddate) = " & year(date1) & " and Month(recorddate) =  " & Month(date1)
            '        sql = "select max(Ser) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            sql = "select max(Ser) as last_sand_no from  TblExchange where  BranchID= " & my_branch & " and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
        
        End If
        
        If Sanad_No = 60 Then
            mWhereBranchYear = " Year(ContDate) = " & year(date1) & " and Month(ContDate) =  " & Month(date1)
         
            sql = "select max(Ser) as last_sand_no from  TblContract where  Branch_NO= " & my_branch & " and year(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
            sql = sql & " and " & mWhere3
        
        End If
        
        'TblContract  Branch_NO
        'Dim stockSettelmentsstr As String
        'stockSettelmentsstr = ""
        
        If Sanad_No = 62 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1) & " and Month(NoteDate) =  " & Month(date1)
            sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= " & NoteType & ")   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser
        End If
        ''////
        If Sanad_No = 64 Then
            mWhereBranchYear = " Year(recorddate) = " & year(date1) & " and Month(recorddate) =  " & Month(date1)
            sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 66 Then
            mWhereBranchYear = " Year(Transaction_Date) = " & year(date1) & " and  Month(Transaction_Date) = " & Month(date1)
            sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR  Transaction_Type=18)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 67 Then
            mWhereBranchYear = " Year(Transaction_Date) = " & year(date1) & " and  Month(Transaction_Date) = " & Month(date1)
            sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
        If Sanad_No = 68 Then
            mWhereBranchYear = " Year(recorddate) = " & year(date1) & " and Month(recorddate) =  " & Month(date1)
            sql = "select max  max(CAST (Ser AS BigInt)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=0)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
          
        End If
        If Sanad_No = 69 Then
            mWhereBranchYear = " Year(recorddate) = " & year(date1) & " and Month(recorddate) =  " & Month(date1)
            sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=1)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 70 Then
            sql = "select max (Ser ) as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 71 Then
            sql = "select max (Ser ) as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 72 Then
            sql = "select max (Ser ) as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 74 Then
            mWhereBranchYear = " Year(NoteDate) = " & year(date1) & " and Month(NoteDate) =  " & Month(date1)
            sql = "select max (Ser ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "   and  notetype=370   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
         
        If Sanad_No = 76 Then
            sql = "select max (Ser ) as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "      and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 75 Then
            sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS BigInt))as last_sand_no  "
            sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
            sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
            sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")and year(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                
            'Sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "     and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
                
        If Sanad_No = 65 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                Else
                    sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    'sql = sql & " and " & mWhere3
                End If
            Else
            
                sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                'sql = sql & " and " & mWhere3
            End If
        End If
        sql = sql & " and " & mWhere3
        ''//////
        If Transaction_Type <> 0 Then
            '   If Transaction_Type = 15 Or Transaction_Type = 16 Then
            '    stockSettelmentsstr
            '   End If
        
            If StoreCoding = True Then
                'Or Transaction_Type= 992)
                If Transaction_Type = 10 Then
                    mWhereBranchYear = " Year(Transaction_Date) = " & year(date1) & " and  Month(Transaction_Date) = " & Month(date1)
                    sql = "select max(  (Ser    )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & "   Or Transaction_Type= 992)  and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    sql = sql & " and " & mWhere3
                     
                Else
                    mWhereBranchYear = " Year(Transaction_Date) = " & year(date1) & " and  Month(Transaction_Date) = " & Month(date1)
                    sql = "select max(  (Ser    )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    sql = sql & " and " & mWhere3
                End If
                sql = sql & mWhereUser2
            Else
                If SystemOptions.BranchDigit > 1 Then
                 
                    '   sql = "select max(  (Ser  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                    'edit edit salim here
                    If Transaction_Type = 10 Then
                        mWhereBranchYear = " Year(Transaction_Date) = " & year(date1) & " and  Month(Transaction_Date) = " & Month(date1)
                        sql = "select max(  (Ser  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & "  Or Transaction_Type= 992)   "
                        sql = sql & " and " & mWhere3
                    Else
                        mWhereBranchYear = " Year(Transaction_Date) = " & year(date1) & " and  Month(Transaction_Date) = " & Month(date1)
                        sql = "select max(  (Ser  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
                        sql = sql & " and " & mWhere3
                    End If
                    sql = sql & mWhereUser2
                    'edit edit salim here
            
                Else
                    If Transaction_Type = 10 Then
                        mWhereBranchYear = " Year(Transaction_Date) = " & year(date1) & " and  Month(Transaction_Date) = " & Month(date1)
                        sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and  ( Transaction_Type=" & Transaction_Type & "   Or Transaction_Type= 992)  "
                        sql = sql & " and " & mWhere3
                    Else
                        mWhereBranchYear = " Year(Transaction_Date) = " & year(date1) & " and  Month(Transaction_Date) = " & Month(date1)
                        sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
                    End If
                    sql = sql & mWhereUser2
                End If
       
            End If
            
            '
            If Prefix = "" Then
                sql = sql & "  and   Prefix is null"
 
            Else
                sql = sql & "  and   Prefix='" & Prefix & "'"
            End If
  
            If StoreCoding = True And StoreID <> 0 Then
                sql = sql & "  and   StoreID=" & StoreID
            End If
  
        End If
    
        If Prefix = "" Then
            If Sanad_No <> 58 And Sanad_No <> 50 And Sanad_No <> 60 Then
                sql = sql & "  and   Prefix is null"
            End If
 
        Else
            sql = sql & "  and   Prefix='" & Prefix & "'"
        End If
        
        If mTableName <> "" Then
            sql = "select max (Ser ) as last_sand_no from  " & mTableName & "  where  BranchId= " & my_branch & "      and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3

        End If
        
        sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        sql = sql & " and " & IIf(mWhereBranchYear = "", " 1 = 1 ", mWhereBranchYear)
        Rs3.Open sql, Cn, adOpenForwardOnly, adLockReadOnly  ' , adCmdText

        Dim startrreadding As Integer
        Dim noofreadinchar As Integer
        If Not IsNull(Rs3("last_sand_no").value) Then
            If StoreCoding = True And StoreID <> 0 Then
                startrreadding = SystemOptions.BranchDigit + SystemOptions.StoreDigit + YearDigit + noOfDigit
                noofreadinchar = startrreadding - 1
                '    If YearDigit = 2 Then
                     
                '      no = Mid(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
                '    Else
                     
                '    no = Mid(Rs3("last_sand_no").value, 10, Len(Rs3("last_sand_no").value) - 9)
                '   End If
               
            Else
           
                startrreadding = SystemOptions.BranchDigit + YearDigit + noOfDigit
                If Transaction_Type = 0 And (Sanad_No <> 66 And Sanad_No <> 67) Then
                    startrreadding = 1 + YearDigit + noOfDigit
                End If
           
                noofreadinchar = startrreadding - 1
                   
                '                      If YearDigit = 2 Then
                '             no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                '           Else
                '
                '           no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                '          End If
           
            End If
            NO = mId(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
            NO = right(Rs3("last_sand_no").value, noOfDigit)
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Voucher_codingByBreaks = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ
        mWhereBranchYear = " Year = " & year(date1) & " and Month =  " & Month(date1)
        mWhereUser2 = mWhereUser2 & " And " & mWhereBranchYear
        mWhere3 = mWhere3 & mWhereUser2
        If Sanad_No = 64 Then
            sql = "select max(CAST (Ser AS BigInt))as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        
        If Sanad_No = 65 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
                If BillType = 1 Then
                    sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Else
                    sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                End If
            Else
            
                sql = "select max(CAST (Ser AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & " and " & mWhere3
            End If
        End If
 
        sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        sql = sql & mWhereUser
        sql = sql & " and " & mWhere3
        If Sanad_No = 5 Then
            sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and not(BTCashAccountcode is null )"
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
      
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)     and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & mWhereUser
                sql = sql & " and " & mWhere3
            End If
        
        End If
     
        If Sanad_No = 4 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            Else
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            End If
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
          
        If Sanad_No = 25 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)      and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & mWhereUser
                sql = sql & " and " & mWhere3
            End If
        
        End If
     
        If Sanad_No = 2 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            Else
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            End If
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›”  —ﬁ”„ ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)   and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & mWhereUser
                sql = sql & " and " & mWhere3
            End If
        
        End If
     
        If Sanad_No = 16 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)        and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            Else
                sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            End If
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
    
        If Sanad_No = 50 Then
      
            '        sql = "select max(Ser) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            'sql = "select max(Ser) as last_sand_no from  TblCarBillMentains where  branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
            sql = "select max(Ser) as last_sand_no from  TblCarBillMentains where  BranchID= " & my_branch & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 58 Then
      
            '        sql = "select max(Ser) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
            'sql = "select max(Ser) as last_sand_no from  TblCarBillMentains where  branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
            sql = "select max(Ser) as last_sand_no from  TblExchange where  BranchID= " & my_branch & " and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            sql = sql & " and " & mWhere3
        End If
        
        If Sanad_No = 60 Then
            sql = "select max(Ser) as last_sand_no from  TblContract where  Branch_NO= " & my_branch & " and year(ContDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
   
        End If
        
        If Sanad_No = 62 Then
           
            sql = "select max(Ser) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= " & NoteType & ")    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 66 Then
            sql = "select  max(CAST (Ser AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR Transaction_Type=18)  and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            sql = sql & mWhereUser2
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 67 Then
            sql = "select  max(CAST (Ser AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and  (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            sql = sql & mWhereUser2
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 68 Then
            sql = "select  max(CAST (Ser AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and  (ImportExport=0)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 69 Then
            sql = "select  max(CAST (Ser AS BigInt))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and  (ImportExport=1)   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 70 Then
            sql = "select  max( (Ser)  as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 71 Then
            sql = "select  max( (Ser)  as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(SDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 72 Then
            sql = "select  max( (Ser)  as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(RecordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 74 Then
            sql = "select  max( (Ser)  as last_sand_no from  notes_all where  branch_no= " & my_branch & " and  notetype=370    and year(NoteDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            sql = sql & mWhereUser2
            sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 76 Then
            sql = "select max (Ser ) as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "     and year(recordDate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        If Sanad_No = 75 Then
            sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS BigInt)) as last_sand_no  "
            sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
            sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
            sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")and year(dbo.TblContractInstallments.Installdate)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " "
                
            'Sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "     and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        
        'TblContract  Branch_NO
        If Transaction_Type <> 0 Then
            If StoreCoding = True Then
                sql = "select  max(  (Ser  ))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
                        
            Else
                sql = "select  max(CAST (Ser AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4)
            End If
            
            If StoreCoding = True And StoreID <> 0 Then
                sql = sql & "  and   StoreID=" & StoreID
            End If
            sql = sql & mWhereUser2
            sql = sql & " and " & mWhere3
        End If
        
        If Prefix = "" Then
            If Sanad_No = 58 Or Sanad_No = 60 Then
            Else
                sql = sql & "  and   Prefix is null"
            End If
        Else
            sql = sql & "  and   Prefix='" & Prefix & "'"
        End If
  
        sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If Not IsNull(Rs3("last_sand_no").value) Then
            If StoreCoding = True And StoreID <> 0 Then
                                           
                startrreadding = SystemOptions.BranchDigit + SystemOptions.StoreDigit + YearDigit + 1
                noofreadinchar = startrreadding - 1
                'If YearDigit = 2 Then
                '            no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                '         Else
                '         no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                '         End If
         
            Else
                If val(getNoOfBranches) > 9 Then
         
                    If mId(Rs3("last_sand_no").value, 1, 1) = "0" Then
         
                        startrreadding = SystemOptions.BranchDigit + YearDigit + 1
                    Else
             
                        If val(my_branch) > 9 Then
                            startrreadding = SystemOptions.BranchDigit + YearDigit + 1
                        Else
                            startrreadding = SystemOptions.BranchDigit + YearDigit
                        End If
             
                    End If
             
                Else
                    startrreadding = SystemOptions.BranchDigit + YearDigit
                    'noofreadinchar = startrreadding
                    If Transaction_Type <> 0 Then
                        If SystemOptions.BranchDigit = 1 Then
                            startrreadding = startrreadding + 1
                        End If
                    Else
                        startrreadding = startrreadding + 1
                    End If
                End If
      
                If Transaction_Type = 0 Then
                    'startrreadding = 1 + YearDigit + 1
                End If
                   
                '                               If YearDigit = 2 Then
                '           no = Mid(Rs3("last_sand_no").value, 4, Len(Rs3("last_sand_no").value) - 3)
                '        Else
                '        no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                '        End If
         
            End If
            noofreadinchar = startrreadding - 1
            NO = mId(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Voucher_codingByBreaks = "error"
                Exit Function
            End If
        End If
 
    End If

    Dim Askcount As Double
    'Askcount = 3
    Askcount = noOfDigit

    If Askcount = 0 Then Askcount = 3

    If Rs3.EOF Or IsNull(Rs3("last_sand_no").value) Then
        first_serial = True

        If numbering_type = 0 Then
                 
        ElseIf numbering_type = 1 Then
            auto_sanad_no = start_at
        ElseIf numbering_type = 2 Then
        
            If IsCodeByBranch Then
                mSerInv = 1
                
                If IsSerialByUser Then
                    auto_sanad_no = my_branch & Breaks & mUserId & Breaks & year(date1) & Breaks & Month(date1) & Breaks & mSerInv
                Else
                    auto_sanad_no = my_branch & Breaks & year(date1) & Breaks & Month(date1) & Breaks & mSerInv
                End If
            Else
                If IsSerialByUser Then
                    auto_sanad_no = mUserId & Breaks & year(date1) & Breaks & Month(date1) & Breaks & mSerInv
                    
                Else
                    auto_sanad_no = year(date1) & Breaks & Month(date1) & Breaks & mSerInv
                End If
            End If
            'mId(Format$(date1, "dd/mm/yyyy"), 9, 2) & mId(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
        
        ElseIf numbering_type = 3 Then
        
            mSerInv = 1
            If IsCodeByBranch Then
                If IsSerialByUser Then
                    auto_sanad_no = my_branch & Breaks & mUserId & Breaks & year(date1) & Breaks & mSerInv
                Else
                    auto_sanad_no = my_branch & Breaks & year(date1) & Breaks & mSerInv
                End If
            Else
                If IsSerialByUser Then
                    auto_sanad_no = mUserId & Breaks & year(date1) & Breaks & mSerInv
                        
                Else
                    auto_sanad_no = year(date1) & Breaks & mSerInv
                End If
            End If
        End If

    Else

        If numbering_type = 0 Then
                
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Rs3("last_sand_no").value + 1
        ElseIf numbering_type = 2 Then
        
            mSerInv = val(Rs3("last_sand_no").value + 1)
            If IsCodeByBranch Then
                
                If IsSerialByUser Then
                    auto_sanad_no = my_branch & Breaks & mUserId & Breaks & year(date1) & Breaks & Month(date1) & Breaks & mSerInv
                Else
                    auto_sanad_no = my_branch & Breaks & year(date1) & Breaks & Month(date1) & Breaks & mSerInv
                End If
            Else
                If IsSerialByUser Then
                    auto_sanad_no = mUserId & Breaks & year(date1) & Breaks & Month(date1) & Breaks & mSerInv
                    
                Else
                    auto_sanad_no = year(date1) & Breaks & Month(date1) & Breaks & mSerInv
                End If
            End If
              
        ElseIf numbering_type = 3 Then
             
            mSerInv = val(Rs3("last_sand_no").value + 1)
            If IsCodeByBranch Then
                
                If IsSerialByUser Then
                    auto_sanad_no = my_branch & Breaks & mUserId & Breaks & year(date1) & Breaks & mSerInv
                Else
                    auto_sanad_no = my_branch & Breaks & year(date1) & Breaks & Breaks & mSerInv
                End If
            Else
                If IsSerialByUser Then
                    auto_sanad_no = mUserId & Breaks & year(date1) & Breaks & Breaks & mSerInv
                    
                Else
                    auto_sanad_no = year(date1) & Breaks & Breaks & mSerInv
                End If
            End If
                      
        End If
        
    End If

    Rs3.Close
    'Dim storeADDZero As String
    'storeADDZero = IIf(StoreID < 10, "0", "")
    Dim brancHcode As String

    If numbering_type = 1 Then
        Voucher_codingByBreaks = auto_sanad_no
        Exit Function
    End If
    If first_serial = True Then
        If auto_sanad_no <> "" Then
            
            If StoreCoding = True And StoreID <> 0 Then
                '       Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
                Voucher_codingByBreaks = auto_sanad_no
            Else
                Voucher_codingByBreaks = auto_sanad_no
            End If
        
        Else
            Voucher_codingByBreaks = auto_sanad_no
        End If

    Else
        '     Voucher_coding = my_branch & auto_sanad_no
        If StoreCoding = True And StoreID <> 0 Then
            ' Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
            Voucher_codingByBreaks = auto_sanad_no
        Else
            Voucher_codingByBreaks = auto_sanad_no
        End If
    End If

End Function

