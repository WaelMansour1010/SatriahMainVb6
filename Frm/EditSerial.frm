VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 -- „ﬂ«‰ «· ⁄œÌ·
 Õ  ﬂ·„…
NewEdit
ÊÌ„ﬂ‰‰« «·»ÕÀ ⁄‰
If SystemOptions.BranchDigit > 1
À„ ‰÷⁄  «·ﬂÊœ «·–Ï »Ì‰ «·ﬂÊ„‰ 


 If SystemOptions.BranchDigit > 1 Then
            'NewEdit
            If Transaction_Type = 10 Then
                sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & " Or Transaction_Type= 992)   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
            Else
                sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
            End If
            'NewEdit
            
            
 
 
 
 
 Public Function Voucher_coding(my_branch As Integer, _
                               date1 As Date, _
                               Sanad_No As Integer, _
                               NoteType As Integer, _
                               Optional departement_name As Integer = 1, _
                               Optional Transaction_Type As Integer = 0, _
                               Optional Prefix As String = "", Optional StoreId As Integer = 0, Optional BillType As Integer = 0, Optional MosemID As Double = 0) As String
    
    
      On Error Resume Next
If my_branch = 0 Then Exit Function
    Dim start_at As Double
    Dim end_at As Double
    Dim auto_sanad_no As String
    Dim NO As Double
    Dim numbering_type As Integer
    Dim noOfDigit As Double
    Dim Zeros As Double
    Dim StoreCoding As Double
Dim YearDigit As Double
Dim branchpadidng As Integer
Dim storepadding As Integer

     auto_sanad_no = ""
 
    Dim first_serial  As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim Rs3 As ADODB.Recordset
     Set Rs3 = New ADODB.Recordset

    Dim sql As String
    Dim i As Integer
Dim storecode As String


     first_serial = False
     sql = "select * from sanad_numbering where branch_no=" & my_branch & " and  sanad_no=" & Sanad_No
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        numbering_type = 0
   
                
    Else
        numbering_type = IIf(IsNull(rs("numbering_id").Value), 0, rs("numbering_id").Value)
        start_at = IIf(IsNull(rs("start_at").Value), 0, rs("start_at").Value)
        end_at = IIf(IsNull(rs("end_at").Value), 0, rs("end_at").Value)
        noOfDigit = IIf(IsNull(rs("no_of_digit").Value), 0, rs("no_of_digit").Value)
        If noOfDigit = 0 Then noOfDigit = 3
        Zeros = IIf(IsNull(rs("zeros").Value), 0, rs("zeros").Value)
        StoreCoding = IIf(IsNull(rs("StoreCoding").Value), 0, rs("StoreCoding").Value)
        YearDigit = IIf(IsNull(rs("YearDigit").Value), 4, rs("YearDigit").Value)
        
        storepadding = SystemOptions.StoreDigit - 1
        If YearDigit = 0 Then YearDigit = 4
        
branchpadidng = SystemOptions.BranchDigit - 1

 

        If StoreCoding = True Then
                    If StoreId <> 0 Then
                      storecode = getStoreCoding(StoreId)
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
            
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 9 or NoteType= 10)  "
             
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

        If Not IsNull(Rs3("last_sand_no").Value) Then
        
            If end_at = 0 Then end_at = Val(Rs3("last_sand_no").Value) + 1
               
            If Rs3("last_sand_no").Value >= end_at Then
                Voucher_coding = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 2 Then ' „ ’· ‘Â—Ì

        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)

        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and not(BTCashAccountcode is null )"
        End If
    
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        
            End If
        End If
    
        If Sanad_No = 4 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
            End If
        
        End If
    
        If Sanad_No = 25 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        
            End If
        End If
    
        If Sanad_No = 2 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)   and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "       and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
            End If
        
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        
            End If
        End If
    
        If Sanad_No = 16 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)     and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
            End If
        
        End If
        
        
                     If Sanad_No = 50 Then
      
        '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
   sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  BranchID= " & my_branch & " and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
        End If
                  If Sanad_No = 58 Then
      
        '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
   sql = "select max(NoteSerial1) as last_sand_no from  TblExchange where  BranchID= " & my_branch & " and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
        
        End If
        
        
                      If Sanad_No = 60 Then
      
         
   sql = "select max(NoteSerial1) as last_sand_no from  TblContract where  Branch_NO= " & my_branch & " and year(ContDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(ContDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
        End If
        
        
        'TblContract  Branch_NO
    'Dim stockSettelmentsstr As String
    'stockSettelmentsstr = ""
        
        
      If Sanad_No = 62 Then
           
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 9 or NoteType= 10)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        
           
        End If
        ''////
       If Sanad_No = 64 Then
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
              If Sanad_No = 66 Then
        sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR  Transaction_Type=18)   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 67 Then
        sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 68 Then
        sql = "select max  max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=0)   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 69 Then
        sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=1)   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 70 Then
        sql = "select max (NoteSerial1 ) as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 71 Then
        sql = "select max (NoteSerial1 ) as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 72 Then
        sql = "select max (NoteSerial1 ) as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 74 Then
        sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "   and  notetype=370   and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
         
        If Sanad_No = 76 Then
        sql = "select max (NoteSerial1 ) as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "      and year(recordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        If Sanad_No = 75 Then
                   sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS FLOAT))as last_sand_no  "
                sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
                sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
                sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                
        'Sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "     and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        
             
                
            If Sanad_No = 65 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
            If BillType = 1 Then
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                Else
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                End If
            Else
            
         sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        End If
  
        ''//////
        If Transaction_Type <> 0 Then
             '   If Transaction_Type = 15 Or Transaction_Type = 16 Then
             '    stockSettelmentsstr
             '   End If
        
        
        
        If StoreCoding = True Then
               sql = "select max(  (NoteSerial1    )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        Else
        If SystemOptions.BranchDigit > 1 Then
            'NewEdit
            If Transaction_Type = 10 Then
                sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & " Or Transaction_Type= 992)   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
            Else
                sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
            End If
            'NewEdit
        Else
                 sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
       End If
       
        End If
            
'
            If Prefix = "" Then
                sql = sql & "  and   Prefix is null"
 
            Else
                sql = sql & "  and   Prefix='" & Prefix & "'"
            End If
  
  
  If StoreCoding = True And StoreId <> 0 Then
       sql = sql & "  and   StoreID=" & StoreId
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
Dim noofreadinchar  As Integer
        If Not IsNull(Rs3("last_sand_no").Value) Then
          If StoreCoding = True And StoreId <> 0 Then
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
             NO = Mid(Rs3("last_sand_no").Value, startrreadding, Len(Rs3("last_sand_no").Value) - noofreadinchar)
             NO = Right(Rs3("last_sand_no").Value, noOfDigit)
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Voucher_coding = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ
           If Sanad_No = 64 Then
                sql = "select max(CAST (NoteSerial1 AS FLOAT))as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        
            If Sanad_No = 65 Then
            If SystemOptions.AllowProjectBill2Serial = True Then
            If BillType = 1 Then
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Else
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
                End If
            Else
            
         sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        End If
        
 
        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)

        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "   and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and not(BTCashAccountcode is null )"
        End If
      
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)     and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            End If
        
        End If
     
        If Sanad_No = 4 Then
            If SystemOptions.ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)  and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            End If
        
        End If
          
        If Sanad_No = 25 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)      and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            End If
        
        End If
     
        If Sanad_No = 2 Then
            If SystemOptions.InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
            End If
        
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›”  —ﬁ”„ ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)   and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            End If
        
        End If
     
        If Sanad_No = 16 Then
            If SystemOptions.ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)        and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            End If
        
        End If
    
    
                         If Sanad_No = 50 Then
      
        '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
   'sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
   sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  BranchID= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
   
        End If
                          If Sanad_No = 58 Then
      
        '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
   'sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
   sql = "select max(NoteSerial1) as last_sand_no from  TblExchange where  BranchID= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
   
        End If
        
        If Sanad_No = 60 Then
    sql = "select max(NoteSerial1) as last_sand_no from  TblContract where  Branch_NO= " & my_branch & " and year(ContDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
   
        End If
        
        
        If Sanad_No = 62 Then
           
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 9 or NoteType= 10)     and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
          
        
        End If
        If Sanad_No = 66 Then
         sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR Transaction_Type=18)  and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
         End If
        If Sanad_No = 67 Then
         sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and  (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
         End If
         If Sanad_No = 68 Then
         sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and  (ImportExport=0)   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
         End If
         If Sanad_No = 69 Then
         sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and  (ImportExport=1)   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
         End If
         If Sanad_No = 70 Then
         sql = "select  max( (NoteSerial1)  as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(SDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
         End If
          If Sanad_No = 71 Then
         sql = "select  max( (NoteSerial1)  as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(SDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
         End If
         If Sanad_No = 72 Then
         sql = "select  max( (NoteSerial1)  as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  and  (SeasonsID=" & MosemID & ")   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
         End If
         If Sanad_No = 74 Then
         sql = "select  max( (NoteSerial1)  as last_sand_no from  notes_all where  branch_no= " & my_branch & " and  notetype=370    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
         End If
        If Sanad_No = 76 Then
        sql = "select max (NoteSerial1 ) as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "     and year(recordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
         If Sanad_No = 75 Then
         sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS FLOAT)) as last_sand_no  "
                sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
                sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
                sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " "
                
        'Sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "     and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        
        'TblContract  Branch_NO
        If Transaction_Type <> 0 Then
                    If StoreCoding = True Then
                   sql = "select  max(  (NoteSerial1  ))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
                        
                     Else
                     sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
                     End If
            
    If StoreCoding = True And StoreId <> 0 Then
       sql = sql & "  and   StoreID=" & StoreId
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
 
        If Not IsNull(Rs3("last_sand_no").Value) Then
                  If StoreCoding = True And StoreId <> 0 Then
                        
                                           
                  startrreadding = SystemOptions.BranchDigit + SystemOptions.StoreDigit + YearDigit + 1
                   noofreadinchar = startrreadding - 1
                   'If YearDigit = 2 Then
                   '            no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                   '         Else
                   '         no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                   '         End If
         
         Else
         If Val(getNoOfBranches) > 9 Then
         
         If Mid(Rs3("last_sand_no").Value, 1, 1) = "0" Then
         
         startrreadding = SystemOptions.BranchDigit + YearDigit + 1
             Else
             
           If Val(my_branch) > 9 Then
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
           NO = Mid(Rs3("last_sand_no").Value, startrreadding, Len(Rs3("last_sand_no").Value) - noofreadinchar)
            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Voucher_coding = "error"
                Exit Function
            End If
        End If
 
    End If

    Dim Askcount As Double
    'Askcount = 3
    Askcount = noOfDigit

    If Askcount = 0 Then Askcount = 3

    If Rs3.RecordCount = 0 Or IsNull(Rs3("last_sand_no").Value) Then
        first_serial = True

        If numbering_type = 0 Then
                 
        ElseIf numbering_type = 1 Then
            auto_sanad_no = start_at
        ElseIf numbering_type = 2 Then
            
        
                 If YearDigit = 2 Then
            auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 9, 2) & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
           Else
           auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
           End If
        
        ElseIf numbering_type = 3 Then
       
        
                       If YearDigit = 2 Then
            auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format(start_at, String(Askcount, "0"))
          Else
          auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(start_at, String(Askcount, "0"))
          End If
        
       
        End If

    Else

        If numbering_type = 0 Then
                
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Rs3("last_sand_no").Value + 1
        ElseIf numbering_type = 2 Then
             If StoreCoding = True And StoreId <> 0 Then
              
              If YearDigit = 2 Then
'            no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
            '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
            auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 9, 2) & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
             Else
           '  no = Mid(Rs3("last_sand_no").value, 10, Len(Rs3("last_sand_no").value) - 9)
            '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
            auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
             End If
             
             Else
             
                         If YearDigit = 2 Then
           ' no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
            '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
            auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 9, 2) & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
             Else
           '  no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
            '  auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 6) & Format((no + 1), String(Askcount, "0"))
            auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format((NO + 1), String(Askcount, "0"))
             End If
             
             End If
        ElseIf numbering_type = 3 Then
           If StoreCoding = True And StoreId <> 0 Then
            
                              If YearDigit = 2 Then
                            '  no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                              'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                              auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format((NO + 1), String(Askcount, "0"))
                               Else
                               
                                '          no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                              'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                              auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format((NO + 1), String(Askcount, "0"))
                    
                               End If
             
                      
              Else
              
              
                                            If YearDigit = 2 Then
                              'no = Mid(Rs3("last_sand_no").value, 4, Len(Rs3("last_sand_no").value) - 3)
                              'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                              auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 9, 2) & Format((NO + 1), String(Askcount, "0"))
                               Else
                               
                              '            no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                              'auto_sanad_no = Mid(Rs3("last_sand_no").value, 1, 4) & Format((no + 1), String(Askcount, "0"))
                              auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format((NO + 1), String(Askcount, "0"))
                    
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

    If numbering_type = 1 Then Voucher_coding = auto_sanad_no: Exit Function
    If first_serial = True Then
        If auto_sanad_no <> "" Then
            
            If StoreCoding = True And StoreId <> 0 Then
     '       Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
            Voucher_coding = brancHcode & storecode & auto_sanad_no
           Else
             Voucher_coding = brancHcode & auto_sanad_no
           End If
           
        
        Else
            Voucher_coding = auto_sanad_no
        End If

    Else
   '     Voucher_coding = my_branch & auto_sanad_no
          If StoreCoding = True And StoreId <> 0 Then
           ' Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
          Voucher_coding = brancHcode & storecode & auto_sanad_no
           Else
             Voucher_coding = brancHcode & auto_sanad_no
           End If
    End If

End Function





