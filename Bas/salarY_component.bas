Attribute VB_Name = "salarY_component"



    
Public Function get_value(operand As String) As Double
    operand = Replace$(operand, "A", "")
    Dim value As Double
    value = val(operand) * 5
    get_value = value
End Function

Public Function change_equation_string_to_value2(src As String) As Double
    Dim new_pos As Integer
    Dim last_pos As Integer
    Dim cuttent_operand As String
    Dim new_str As String
    Dim objScript As Object
    last_pos = 1
    new_str = ""

    For i = 1 To Len(src)

        If mId(src, i, 1) = "+" Or mId(src, i, 1) = "-" Or mId(src, i, 1) = "*" Or mId(src, i, 1) = "/" Or mId(src, i, 1) = "=" Then
            new_pos = i
            cuttent_operand = mId(src, last_pos, new_pos - last_pos)

            If InStr(cuttent_operand, "A") > 0 Then
                cuttent_operand = get_value(cuttent_operand)
            End If

            new_str = new_str & cuttent_operand & mId(src, i, 1)

            If i < Len(src) Then
                last_pos = new_pos + 1
            Else
                GoTo ll
            End If
        End If
 
    Next i

ll:
    new_str = Replace$(new_str, "=", "")

    Set objScript = CreateObject("MSScriptControl.ScriptControl")
    objScript.Language = "VBScript"

    change_equation_string_to_value2 = objScript.Eval(new_str)

End Function

Public Function addSalaryComponentToEmployee(Emp_id As Integer)
    Dim Emp_Salary_sakn      As Double
    Dim Emp_Salary_bus       As Double
    Dim Emp_Salary_food      As Double
    Dim Emp_Salary_others        As Double
    Dim Emp_Salary_mob       As Double
    Dim Emp_Salary_mang      As Double
    Dim Emp_Salary_sakn1         As Double
    Dim Emp_Salary_bus1      As Double
    Dim Emp_Salary_food1         As Double
    Dim Emp_Salary_others1       As Double
    Dim Emp_Salary_mob1      As Double
    Dim Emp_Salary_mang1         As Double
    Dim sql As String
    Dim rs As ADODB.Recordset

    '«·ﬁÌ„ «·‘Â—Ì…
    Set rs = New ADODB.Recordset
    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=1 and mofrad_type=2 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_sakn = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If

    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=1 and mofrad_type=3 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_bus = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If
    
    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=1 and mofrad_type=4 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_food = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If
    
    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=1 and mofrad_type=5 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_mob = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If
    
    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=1 and mofrad_type=6 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_mang = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If
    
    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=1 and mofrad_type=7 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_others = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If

    '«·ﬁÌ„ «·”‰ÊÌ…
    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=0 and mofrad_type=2 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_sakn1 = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If

    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=0 and mofrad_type=3 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_bus1 = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If
    
    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=0 and mofrad_type=4 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_food1 = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If
    
    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=0 and mofrad_type=5 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_mob1 = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If
    
    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=0 and mofrad_type=6 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_mang1 = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If
    
    sql = "Select sum([Value]) as total From EmpSalaryComponent where  Monthly=0 and mofrad_type=7 and emp_ID=" & Emp_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_Salary_others1 = IIf(IsNull(rs("total").value), 0, rs("total").value)
        rs.Close
    End If
 
    sql = "update TblEmployee Set "
    sql = sql + "Emp_Salary_sakn=" & Emp_Salary_sakn
    sql = sql + ",Emp_Salary_bus=" & Emp_Salary_bus
    sql = sql + ",Emp_Salary_food=" & Emp_Salary_food
    sql = sql + ",Emp_Salary_mob=" & Emp_Salary_mob
    sql = sql + ",Emp_Salary_mang=" & Emp_Salary_mang
    sql = sql + ",Emp_Salary_others=" & Emp_Salary_others
    
    sql = sql + ",Emp_Salary_sakn1=" & Emp_Salary_sakn1
    sql = sql + ",Emp_Salary_bus1=" & Emp_Salary_bus1
    sql = sql + ",Emp_Salary_food1=" & Emp_Salary_food1
    sql = sql + ",Emp_Salary_mob1=" & Emp_Salary_mob1
    sql = sql + ",Emp_Salary_mang1=" & Emp_Salary_mang1
    sql = sql + ",Emp_Salary_others1=" & Emp_Salary_others1
    
    sql = sql + " Where emp_ID=" & Emp_id
    Cn.Execute sql
End Function

Public Function saveExpensesDetails(Grid As Integer, _
                                    Optional NoteSerial As String = "", _
                                    Optional NoteSerial1 As String = "", _
                                    Optional order_no As String = "", _
                                    Optional RecordDate As Date, _
                                    Optional NoteID As Long = 0) As Boolean
    Dim ExpensesID As Double

    Dim line_no As Integer
 
    Dim RsExpensesDetails As ADODB.Recordset
    Set RsExpensesDetails = New ADODB.Recordset
    
    'RsExpensesDetails.Open "ExpensesDetails", Cn, adOpenDynamic, adLockPessimistic, adCmdTableDirect
    StrSQL = "SELECT   *  from dbo.ExpensesDetails Where (1 = -1)"
   RsExpensesDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    
    StrSQL = "Delete From ExpensesDetails Where NoteSerial1='" & NoteSerial1 & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords

    If Grid = 0 Then
        If SystemOptions.gldetails_or_gl_general = 0 And FrmExpenses5.dcproject.BoundText <> "" Then

            With FrmExpenses5.VSFlexGrid1

                For i = .FixedRows To .rows - 1
 
                    If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
           
                        RsExpensesDetails.AddNew
                
                        RsExpensesDetails("Noteid").value = NoteID
                        '                RsExpensesDetails("Destribute").value = IIf(.TextMatrix(I, .ColIndex("Destribute")) = "", Null, .TextMatrix(I, .ColIndex("Destribute")))
                        RsExpensesDetails("DEV_ID_Line_No").value = IIf(.TextMatrix(i, .ColIndex("LineNo")) = "", Null, .TextMatrix(i, .ColIndex("LineNo")))
                        RsExpensesDetails("DEV_ID_Line_No1").value = IIf(.TextMatrix(i, .ColIndex("LineNo1")) = "", Null, .TextMatrix(i, .ColIndex("LineNo1")))
                 
                        RsExpensesDetails("AccountCode").value = IIf(.TextMatrix(i, .ColIndex("AccountCode")) = "", Null, .TextMatrix(i, .ColIndex("AccountCode")))
                        RsExpensesDetails("ExpensesID").value = .TextMatrix(i, .ColIndex("AccountCode"))
                        RsExpensesDetails("ExpensesName").value = .TextMatrix(i, .ColIndex("AccountName"))
                
                        RsExpensesDetails("Value").value = .TextMatrix(i, .ColIndex("value"))
                               
                        RsExpensesDetails("NoteSerial").value = NoteSerial '„”·”· «·ﬁÌœ
                        RsExpensesDetails("NoteSerial1").value = NoteSerial1  '„”·”· «–‰ «·’—›
                        RsExpensesDetails("RecordDate").value = RecordDate
                        RsExpensesDetails("opr_fullcode").value = .TextMatrix(i, .ColIndex("opr_fullcode"))
              
                        If order_no <> "" Then
                            RsExpensesDetails("order_no").value = order_no
                        Else
                            '  RsExpensesDetails("order_no").value = IIf(.TextMatrix(I, .ColIndex("Order_No")) = "", Null, .TextMatrix(I, .ColIndex("Order_No")))
                        End If

                        RsExpensesDetails("Des").value = .TextMatrix(i, .ColIndex("des"))

                        RsExpensesDetails.update
        
                    End If

                Next i

            End With

        Else

            With FrmExpenses5.Fg_Journal

                For i = .FixedRows To .rows - 1
 
                    If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
           
                        RsExpensesDetails.AddNew
                        RsExpensesDetails("Noteid").value = NoteID
                        RsExpensesDetails("Destribute").value = IIf(.TextMatrix(i, .ColIndex("Destribute")) = "", Null, .TextMatrix(i, .ColIndex("Destribute")))
                        RsExpensesDetails("DEV_ID_Line_No").value = IIf(.TextMatrix(i, .ColIndex("AccountCode")) = "", Null, .TextMatrix(i, .ColIndex("LineNo")))
                        RsExpensesDetails("DEV_ID_Line_No1").value = IIf(.TextMatrix(i, .ColIndex("AccountCode")) = "", Null, val(.TextMatrix(i, .ColIndex("LineNo1"))))
                        RsExpensesDetails("AccountCode").value = IIf(.TextMatrix(i, .ColIndex("AccountCode")) = "", Null, .TextMatrix(i, .ColIndex("AccountCode")))
                        RsExpensesDetails("ExpensesID").value = .TextMatrix(i, .ColIndex("ExpensesID"))
                        RsExpensesDetails("ExpensesName").value = .TextMatrix(i, .ColIndex("AccountName"))
                
                        RsExpensesDetails("Value").value = .TextMatrix(i, .ColIndex("value"))
                               
                        RsExpensesDetails("NoteSerial").value = NoteSerial '„”·”· «·ﬁÌœ
                        RsExpensesDetails("NoteSerial1").value = NoteSerial1  '„”·”· «–‰ «·’—›
                        RsExpensesDetails("RecordDate").value = RecordDate
                        RsExpensesDetails("opr_fullcode").value = .TextMatrix(i, .ColIndex("opr_fullcode"))
              
                        If order_no <> "" Then
                            RsExpensesDetails("order_no").value = order_no
                        Else
                            RsExpensesDetails("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
                        End If

                        RsExpensesDetails("Des").value = .TextMatrix(i, .ColIndex("des"))

                        RsExpensesDetails.update
        
                    End If

                Next i

            End With

        End If

    Else

        With FrmExpenses3.Fg_Journal

            For i = .FixedRows To .rows - 1
 
                If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
           
                    RsExpensesDetails.AddNew
                    RsExpensesDetails("AccountCode").value = IIf(.TextMatrix(i, .ColIndex("AccountCode")) = "", Null, .TextMatrix(i, .ColIndex("AccountCode")))
                    RsExpensesDetails("ExpensesID").value = .TextMatrix(i, .ColIndex("ExpensesID"))
                    RsExpensesDetails("ExpensesName").value = .TextMatrix(i, .ColIndex("AccountName"))
                
                    RsExpensesDetails("Value").value = val(.TextMatrix(i, .ColIndex("value")))
                               
                    RsExpensesDetails("NoteSerial").value = NoteSerial '„”·”· «·ﬁÌœ
                    RsExpensesDetails("NoteSerial1").value = NoteSerial1  '„”·”· «–‰ «·’—›
                    RsExpensesDetails("RecordDate").value = RecordDate
                    RsExpensesDetails("opr_fullcode").value = .TextMatrix(i, .ColIndex("opr_fullcode"))
              
                    If order_no <> "" Then
                        RsExpensesDetails("order_no").value = order_no
                    Else
                        RsExpensesDetails("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
                    End If

                    RsExpensesDetails("Des").value = .TextMatrix(i, .ColIndex("des"))

                    RsExpensesDetails.update
        
                End If

            Next i

        End With

    End If

End Function

Public Function getProjectAccountwhereString(project_id As Integer) As String
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim str As String
    Set rs = New ADODB.Recordset
    sql = "select * From projects where id=" & project_id
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        str = "ACCOUNTS.Account_Code='" & IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value) & "' or ACCOUNTS.Account_Code='" & IIf(IsNull(rs("REVENUE_account").value), "", rs("REVENUE_account").value)
        str = str & "' or ACCOUNTS.Account_Code='" & IIf(IsNull(rs("Material_account").value), "", rs("Material_account").value)
        str = str & "' or ACCOUNTS.Account_Code='" & IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
        str = str & "' or ACCOUNTS.Account_Code='" & IIf(IsNull(rs("legal").value), "", rs("legal").value) & "'"

        getProjectAccountwhereString = str
    Else
        getProjectAccountwhereString = "ACCOUNTS.Account_Code ='XX'"
    End If

End Function

Public Function GetArrowsBocketData(BocketId As Integer, _
                                    Optional ByRef BocketCode As String, _
                                    Optional ByRef Balance As Double)

    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from Bockets where BocketId=" & BocketId
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        BocketCode = IIf(IsNull(Rs1("BocketCode").value), "", Rs1("BocketCode").value)
        Balance = IIf(Not IsNumeric(Rs1("Balance").value), 0, Rs1("Balance").value)
    End If
  
    Rs1.Close
End Function

Public Function GetArrowsCompanyData(CompanyID As Integer, _
                                     Optional ByRef CompanySymbol As String = "", _
                                     Optional ByRef CompanyName As String, _
                                     Optional ByRef GroupID As Integer, _
                                     Optional ByRef FinMarketId As Integer, _
                                     Optional ByRef currentvalue As Double, _
                                     Optional ByRef CurrencyId As Integer, _
                                     Optional ByRef StatusId As Integer, _
                                     Optional ByRef Payedcapital As Double)

    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from ArrowsCompanies where CompanyId=" & CompanyID

    If CompanySymbol <> "" Then
        sql = "SELECT * from ArrowsCompanies where CompanySymbol='" & CompanySymbol & "'"
    End If

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        CompanyID = IIf(Not IsNumeric(Rs1("CompanyId").value), 0, Rs1("CompanyId").value)
        CompanySymbol = IIf(IsNull(Rs1("CompanySymbol").value), "", Rs1("CompanySymbol").value)
        CompanyName = IIf(IsNull(Rs1("CompanyName").value), "", Rs1("CompanyName").value)
        GroupID = IIf(Not IsNumeric(Rs1("GroupID").value), 0, Rs1("GroupID").value)
        FinMarketId = IIf(Not IsNumeric(Rs1("FinMarketId").value), 0, Rs1("FinMarketId").value)
        currentvalue = IIf(Not IsNumeric(Rs1("CurrentValue").value), 0, Rs1("CurrentValue").value)
        CurrencyId = IIf(Not IsNumeric(Rs1("CurrencyId").value), 0, Rs1("CurrencyId").value)
        StatusId = IIf(Not IsNumeric(Rs1("StatusId").value), 0, Rs1("StatusId").value)
        Payedcapital = IIf(Not IsNumeric(Rs1("Payedcapital").value), 0, Rs1("Payedcapital").value)
    Else
        CompanyID = 0
 
    End If
  
    Rs1.Close
End Function

Public Function GetCurrencyData(ID As Integer, _
                                Optional ByRef code As String, _
                                Optional ByRef Name As String, _
                                Optional ByRef NameE As String, _
                                Optional ByRef Rate As Double)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from currency where id=" & ID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        code = IIf(IsNull(Rs1("code").value), "", Rs1("code").value)
        Name = IIf(IsNull(Rs1("name").value), "", Rs1("name").value)
        NameE = IIf(IsNull(Rs1("nameE").value), "", Rs1("nameE").value)
        Rate = IIf(Not IsNumeric(Rs1("rate").value), 1, Rs1("rate").value)
 
    End If
  
    Rs1.Close
End Function

Public Function GetArrowsQty(CompanyID As Integer, _
                             Optional ByRef TotalQty As Double)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT  sum(qty) as Totalqty from ArrowsTransactions where CompanyId=" & CompanyID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        TotalQty = IIf(IsNull(Rs1("Totalqty").value), 0, Rs1("Totalqty").value)
  
    End If
  
    Rs1.Close
End Function
Public Function CheckAutoCoding(FIELD_no As Integer) As Boolean
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "select * from coding where   auto=1 and FIELD_no=" & FIELD_no
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
       CheckAutoCoding = True
       Else
       CheckAutoCoding = False
  
    End If
  
    Rs1.Close
End Function

Public Function GetArrowsGroupAccount(GroupID As Integer, _
                                      Optional ByRef Account_code As String, _
                                      Optional ByRef Account_code1 As String, _
                                      Optional ByRef Account_code2 As String, _
                                      Optional ByRef Account_code3 As String, _
                                      Optional ByRef Account_code4 As String, _
                                      Optional ByRef GroupName As String)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from ArrowsGroup where GroupID=" & GroupID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        Account_code = IIf(IsNull(Rs1("Account_Code").value), "", Rs1("Account_Code").value)
        Account_code1 = IIf(IsNull(Rs1("Account_Code1").value), "", Rs1("Account_Code1").value)
        Account_code2 = IIf(IsNull(Rs1("Account_Code2").value), "", Rs1("Account_Code2").value)
        Account_code3 = IIf(IsNull(Rs1("Account_Code3").value), "", Rs1("Account_Code3").value)
        Account_code4 = IIf(IsNull(Rs1("Account_Code4").value), "", Rs1("Account_Code4").value)
        GroupName = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
 
    End If
  
    Rs1.Close
End Function

Public Function GetFixedAssetsGroupAccount(GroupID As Integer, _
                                           Optional account_type_code As Integer, _
                                           Optional branch_id As Integer = 0, _
                                           Optional ByRef Account_Codex As String, _
                                           Optional ByRef account_name As String, _
                                           Optional ByRef Percentage1 As Integer = 0, _
                                           Optional ByRef Percentage2 As Integer = 0, _
                                           Optional ByRef DepType As Integer = 0, _
                                           Optional ByRef Account_code As String, _
                                           Optional ByRef Account_code1 As String, _
                                           Optional ByRef Account_code2 As String, _
                                           Optional ByRef Account_code3 As String, _
                                           Optional ByRef Account_code4 As String)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from FixedAssetsGroup where GroupID=" & GroupID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        Percentage1 = IIf(Not IsNumeric(Rs1("Percentage1").value), 0, Rs1("Percentage1").value)
        Percentage2 = IIf(Not IsNumeric(Rs1("Percentage2").value), 0, Rs1("Percentage2").value)
        DepType = IIf(IsNull(Rs1("DepType").value), 0, Rs1("DepType").value)
        Account_code = IIf(IsNull(Rs1("Account_Code").value), "", Rs1("Account_Code").value)
        Account_code1 = IIf(IsNull(Rs1("Account_Code1").value), "", Rs1("Account_Code1").value)
        Account_code2 = IIf(IsNull(Rs1("Account_Code2").value), "", Rs1("Account_Code2").value)
        Account_code3 = IIf(IsNull(Rs1("Account_Code3").value), "", Rs1("Account_Code3").value)
        Account_code4 = IIf(IsNull(Rs1("Account_Code4").value), "", Rs1("Account_Code4").value)
  
    End If
 
    'Set rs = New ADODB.Recordset
    'sql = "SELECT     dbo.ACCOUNTS.Account_Name, dbo.FixedAssetsGroupsAccount.account_code, dbo.FixedAssetsGroupsAccount.branch_id, dbo.FixedAssetsGroupsAccount.group_id, " & _
    '               "       dbo.FixedAssetsGroupsAccount.account_type_code " & _
    '" FROM         dbo.FixedAssetsGroupsAccount INNER JOIN" & _
    ' "                     dbo.ACCOUNTS ON dbo.FixedAssetsGroupsAccount.account_code = dbo.ACCOUNTS.Account_Code"
 
    'sql = sql & " where group_id =" & GroupID & " and account_type_code='" & account_type_code & "' "
    'If branch_id <> 0 Then
    'sql = sql & " and  branch_id=" & branch_id
    'End If
    'rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If rs.RecordCount > 0 Then
    'account_name = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
    ' Account_Code = IIf(IsNull(rs("account_code").value), "", rs("account_code").value)
    ' End If
    'rs.Close
    'Rs1.Close
End Function

Public Function SavePaymentAndReciveDetails(Stype As Integer, _
                                            Optional NoteSerial As String = "", _
                                            Optional NoteSerial1 As String = "", _
                                            Optional order_no As String = "", _
                                            Optional RecordDate As Date) As Boolean
    Dim ExpensesID As Double

    Dim line_no As Integer
    Dim NoteID As String
    Dim RsPaymentRecivw As ADODB.Recordset
    Set RsPaymentRecivw = New ADODB.Recordset
    RsPaymentRecivw.Open "ReciveDetails", Cn, adOpenDynamic, adLockPessimistic, adCmdTableDirect
    StrSQL = "Delete From ReciveDetails Where NoteSerial1='" & NoteSerial1 & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
        
    RsPaymentRecivw.AddNew
    RsPaymentRecivw("Type").value = Stype
    RsPaymentRecivw("NoteSerial").value = NoteSerial '„”·”· «·ﬁÌœ
    RsPaymentRecivw("NoteSerial1").value = NoteSerial1    '«Ê «·œ›⁄'„”·”· «–‰ «·„ﬁ»Ê÷« 
    RsPaymentRecivw("RecordDate").value = RecordDate
          
    If Stype = 0 Then
   
        If FrmPayments.DCboCashType.ListIndex = 1 Or FrmPayments.DCboCashType.ListIndex = 0 Then
            RsPaymentRecivw("CusID").value = IIf(FrmPayments.DBCboClientName.text = "", Null, FrmPayments.DBCboClientName.BoundText)
        End If
   
        If FrmPayments.CboPayMentType.ListIndex = 0 Then
            RsPaymentRecivw("BoxID").value = val(FrmPayments.DcboBox.BoundText)
            RsPaymentRecivw("BankID").value = Null
            RsPaymentRecivw("ChqueNum").value = Null
            RsPaymentRecivw("DueDate").value = Null
            RsPaymentRecivw("NoteCashingType").value = 0
        ElseIf FrmPayments.CboPayMentType.ListIndex = 1 Then
            RsPaymentRecivw("BoxID").value = Null
            RsPaymentRecivw("BankID").value = val(FrmPayments.DcboBankName.BoundText)
            RsPaymentRecivw("ChqueNum").value = Trim$(FrmPayments.TxtChequeNumber.text)
            RsPaymentRecivw("DueDate").value = FrmPayments.DtpChequeDueDate.value
            RsPaymentRecivw("NoteCashingType").value = 1
        End If

        RsPaymentRecivw("Value").value = val(FrmPayments.XPTxtVal.text)
        RsPaymentRecivw("Des").value = FrmPayments.txt_general_des.text
        RsPaymentRecivw("BillNO").value = 11
                
    ElseIf Stype = 1 Then

        If FrmCashing.DCboCashType.ListIndex = 1 Or FrmCashing.DCboCashType.ListIndex = 0 Then
            RsPaymentRecivw("CusID").value = IIf(FrmCashing.DBCboClientName.text = "", Null, FrmCashing.DBCboClientName.BoundText)
        End If
   
        If FrmCashing.CboPayMentType.ListIndex = 0 Then
            RsPaymentRecivw("BoxID").value = val(FrmCashing.DcboBox.BoundText)
            RsPaymentRecivw("BankID").value = Null
            RsPaymentRecivw("ChqueNum").value = Null
            RsPaymentRecivw("DueDate").value = Null
            RsPaymentRecivw("NoteCashingType").value = 0
        ElseIf FrmCashing.CboPayMentType.ListIndex = 1 Then
            RsPaymentRecivw("BoxID").value = Null
            RsPaymentRecivw("BankID").value = val(FrmCashing.DcboBankName.BoundText)
            RsPaymentRecivw("ChqueNum").value = Trim$(FrmCashing.TxtChequeNumber.text)
            RsPaymentRecivw("DueDate").value = FrmCashing.DtpChequeDueDate.value
            RsPaymentRecivw("NoteCashingType").value = 1
        End If
            
        RsPaymentRecivw("Value").value = val(FrmCashing.XPTxtVal.text)
        RsPaymentRecivw("Des").value = FrmCashing.XPMTxtRemarks.text
        RsPaymentRecivw("BillNO").value = 11
    End If

    RsPaymentRecivw.update
End Function
Public Function get_coding(branch_no As Integer, _
                           tablename As String, _
                           FIELD_no As Integer, _
                           prifix As String, Optional ByVal mIsString As Boolean = False) As String
    On Error GoTo ErrTrap

    Dim tempcode As String
    Dim code As String

    Dim coding_auto As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
 
    '       If prifix = "" Then
    '           get_coding = "Manual"
    '           Exit Function
    '       End If
                        
    'sql = "select * from coding where FIELD_no=" & FIELD_no & "  and  branch_no=" & branch_no & " and prifix='" & prifix & "'"
   If mIsString Then
        sql = "select * from coding where FIELD_no=" & FIELD_no & " "
    Else
        sql = "select * from coding where FIELD_no=" & FIELD_no & " and prifix='" & prifix & "'"
    End If

    If FIELD_no = 3 Then
        'sql = "select * from coding where FIELD_no=" & FIELD_no & "  and  branch_no=" & branch_no
        sql = "select * from coding where FIELD_no=" & FIELD_no
    End If

    If FIELD_no = 1 Or FIELD_no = 8 Then
        'sql = "select * from coding where FIELD_no=" & FIELD_no & "  and  branch_no=" & branch_no
        sql = "select * from coding where FIELD_no=" & FIELD_no
    End If

    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        coding_auto = rs("Auto").value

        If coding_auto = False Then
            get_coding = "Manual"
            Exit Function
        End If

        If coding_auto = True Then
            no_of_digit = IIf(Not IsNumeric(rs("no_of_digit").value), 0, rs("no_of_digit").value)
            Zeros = IIf(Not IsNumeric(rs("zeros").value), 0, rs("zeros").value)
        End If
    End If

    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset

    'sql = "select max(code)  as last_code from " & tablename & "  where branch_no=" & branch_no & "  and prifix='" & prifix & "'"
    If prifix = "" Then
        sql = "select   max(  CAST(code AS float)  )    as last_code from " & tablename & "  where   prifix is null"
    Else
        sql = "select   max(  CAST(code AS float)  )    as last_code from " & tablename & "  where   prifix='" & prifix & "'"

    End If

    If FIELD_no = 4 Then  '  ???? ?????
        sql = sql & " And  Type=1"
    ElseIf FIELD_no = 5 Then
        sql = sql & " And  Type=2"
        
        
        ElseIf FIELD_no = 8 Then
        sql = sql & " And  Type=3"
        
          ElseIf FIELD_no = 9 Then ' „”«Â„ '
        sql = sql & " And  Type=20"
          ElseIf FIELD_no = 15 Then '„” √Ã—
        sql = sql & " And  Type=56"
         ElseIf FIELD_no = 16 Then '„«·ﬂ
        sql = sql & " And  Type=57"
    End If
  
    If FIELD_no = 6 Then 'employees
        If prifix = "" Then
            sql = "select   max(  CAST(Emp_Code AS float)  )    as last_code from " & tablename & "  where   prifix is null"
        Else
            sql = "select   max(  CAST(Emp_Code AS float)  )    as last_code from " & tablename & "  where   prifix='" & prifix & "'"

        End If
    End If
  
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If IsNull(Rs1("last_code").value) Then

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
                ' MsgBox "??? ??????? ???? ??? ???????  ???? ????? ????? ??? ???? ??????? ?? ???? ????? ?????? ?? ??????? ?????? ??????"
                get_coding = "miniError"
                '  fullcode = "error"
                Exit Function
              
            Else
                fullcode = tempcode
                code = 1
            End If
        End If

    Else
        tempcode = prifix & val(Rs1("last_code").value) + 1
        If val(Rs1("last_code").value) = 999 And no_of_digit = 4 Then
        End If
        
       ' If Len(tempcode) > no_of_digit And val(Rs1("last_code").value) = 999 Then no_of_digit = no_of_digit + 1
        
        If Len(tempcode) < no_of_digit Then
            
            diffrent = no_of_digit - Len(tempcode)
            tempcode = prifix

            If Zeros = True Then

                For i = 1 To diffrent
                    tempcode = tempcode & "0"
                    code = code & "0"
                Next i
                        
            End If

            fullcode = tempcode & val(Rs1("last_code").value) + 1
            code = code & val(Rs1("last_code").value + 1)
        Else

            If Len(tempcode) > no_of_digit And UCase(tablename) <> "TBLITEMS" Then
                'MsgBox "??? ??????? ???? ??? ??????? ????? ??? ???? ???????"
                ' fullcode = "error"
                get_coding = "miniError"
                Exit Function
              
            Else
                fullcode = tempcode
                code = val(Rs1("last_code").value) + 1
            End If
        End If

    End If

    get_coding = code
    Exit Function
ErrTrap:
    get_coding = "miniError"
End Function
 






'==== Module: modSalaryVoucherFilters.bas ====


' ????? SGN -> ?????/????? ????? ??? yyyy-mm-dd



'==== Module: modSalaryVoucherFilters.bas ====


' ????? SGN -> ?????/????? ????? ??? yyyy-mm-dd
Public Sub SgnToMonthBounds(ByVal sgn2 As String, ByRef MonthStart As String, ByRef monthNext As String)
    Dim Y As Integer, m As Integer
    If Len(sgn2) < 5 Then Err.Raise vbObjectError + 101, , "SGN ??? ????"

    Y = CInt(left$(sgn2, 4))
    If Len(sgn2) = 5 Then
        m = CInt(mId$(sgn2, 5, 1))
    Else
        m = CInt(mId$(sgn2, 5, 2))
    End If
    If m < 1 Or m > 12 Then Err.Raise vbObjectError + 102, , "??? SGN ??? ????"

    Dim dStart As Date, dNext As Date
    dStart = DateSerial(Y, m, 1)
    dNext = DateSerial(Y, m + 1, 1)

    MonthStart = Format$(dStart, "yyyy-mm-dd")
    monthNext = Format$(dNext, "yyyy-mm-dd")
End Sub

' ???? ???? NoteId ??????
Public Function VoucherFilterByNoteId(ByVal NoteID As Long, _
                                      Optional ByVal EmployeeQualifier As String = "dbo.TblEmployee", _
                                      Optional ByVal AccountCodeField As String = "Account_code1") As String
    VoucherFilterByNoteId = _
        " AND EXISTS (" & _
        "SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS v " & _
        "WHERE v.Notes_ID = " & CStr(NoteID) & " " & _
        "AND v.Account_Code = " & EmployeeQualifier & "." & AccountCodeField & ")"
End Function

' ???? ??? SGN ? ????? ??? ??????? NoteType=66 ???? ?????/?????
'  ÿ«»ﬁ ‰ ÌÃ… «·«” ⁄·«„ «·„’ÕÕ  „«„«
' -  Œ «— √ﬁœ„ ﬁÌœ —Ê« » (TOP(1) ORDER BY n.NoteDate ASC)
' -  —»ÿ Account_Code „‰ «·ﬁÌœ »ﬂÊœ Õ”«» «·„ÊŸ›
Public Function VoucherFilterBySgnMonthExact(ByVal pSgn As String, _
                                             Optional ByVal EmployeeQualifier As String = "dbo.TblEmployee", _
                                             Optional ByVal AccountCodeField As String = "Account_code1") As String
    Dim MonthStart As String, monthNext As String
    Call SgnToMonthBounds(pSgn, MonthStart, monthNext)

    VoucherFilterBySgnMonthExact = _
        " AND EXISTS (" & _
        "    SELECT 1" & _
        "    FROM dbo.DOUBLE_ENTREY_VOUCHERS v" & _
        "    WHERE v.Account_Code = " & EmployeeQualifier & "." & AccountCodeField & _
        "      AND v.Notes_ID IN (" & _
        "           SELECT TOP (1) n.NoteID" & _
        "           FROM dbo.Notes n" & _
        "           WHERE n.NoteType = 66" & _
        "             AND n.NoteDate >= '" & MonthStart & "'" & _
        "             AND n.NoteDate <  '" & monthNext & "'" & _
        "           ORDER BY n.NoteDate" & _
        "      )" & _
        ")"
End Function




'==== ›Ì ‰›” «·„ÊœÌÊ· «·⁄«„ ====
Public Function VoucherExistsClause(ByVal Y As Integer, ByVal m As Integer, _
                                    Optional ByVal EmployeeQualifier As String = "dbo.TblEmployee", _
                                    Optional ByVal AccountCodeField As String = "Account_code1") As String
    Dim dStart As Date, dNext As Date
    dStart = DateSerial(Y, m, 1)
    dNext = DateAdd("m", 1, dStart)

    VoucherExistsClause = _
        " EXISTS (" & _
        "  SELECT 1 FROM dbo.DOUBLE_ENTREY_VOUCHERS v" & _
        "  WHERE v.Account_Code = " & EmployeeQualifier & "." & AccountCodeField & _
        "    AND v.Notes_ID IN (" & _
        "        SELECT n.NoteID" & _
        "        FROM dbo.Notes n" & _
        "        WHERE n.NoteType = 66" & _
        "          AND n.NoteDate >= '" & Format$(dStart, "yyyy-mm-dd") & "'" & _
        "          AND n.NoteDate <  '" & Format$(dNext, "yyyy-mm-dd") & "'" & _
        "    )" & _
        ")"
End Function


