Attribute VB_Name = "ModMain"


    Public NoOFDigitUserTrans As Integer
    Public StoreDigit As Integer
     Public mPosD As String
            Public mServerD As String
    
    
   Public IsSerialByUserTrans As Boolean
   Public ExpensesCoding As Boolean
    
    Public InstallmntsvchrCoding As Boolean
    Public ExpensesCoding2 As Boolean
    Public AllowProjectBill2Serial As Boolean
    Public NoOFDigitUserVouc As Integer
    Public Ked_digit As Integer
    Public JLCodeBasedOnBranch As Boolean
    Public IsSerialByUserVouch As Boolean
    

Public Cn As New ADODB.Connection
Public POSConnection As New ADODB.Connection
Public ServerDb As String
Public POSDb As String

 Public BranchDigit As Integer

  
Public SysSQLServerType As Integer
Public SysSQLServerName As String
Public SysSQLServerTypeTechnical As String
Public StrAppRegPath As String
Public SysSQLServerDataBaseName As String
Public SysSQLServerUserId As String
Public SysSQLServerUserpassword As String
 
 
Public MainBranch             As String
Public MainBranchID           As Long
Public CountAllBranch         As Long
Public CountAllServer         As Long
'-------------------------------
Public MainServer             As String
Public CurrentServer          As String
Public MainServerID           As Long
Public CurrentServerID        As Long


   
   
  Public Function GetIssueData(Transaction_ID As Double, _
 Optional ByRef NoteId As String, _
  Optional ByRef NoteSerial As String, Optional ByRef NoteSerial1 As String)
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
 sql = " SELECT     NoteId, NoteSerial, NoteSerial1"
 sql = sql & "  From [" & POSDb & "].dbo.Transactions"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"

 
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
 
    If Rs3.RecordCount > 0 Then
      
         NoteId = IIf(Not IsNull(Rs3("NoteId").Value), Rs3("NoteId").Value, 0)
 NoteSerial = IIf(Not IsNull(Rs3("NoteSerial").Value), Rs3("NoteSerial").Value, "")
 
  NoteSerial1 = IIf(Not IsNull(Rs3("NoteSerial1").Value), Rs3("NoteSerial1").Value, "")
 
    Else
 NoteId = 0
        NoteSerial = ""
 NoteSerial1 = ""
      
      End If
 
    Rs3.Close
 End Function
  Public Function get_employee_information(MachinCode As String, Optional Emp_ID As Double, _
 Optional ByRef DepartmentID As Double, _
  Optional ByRef BranchID As Double, Optional ByRef project_id As Double)
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
 sql = " select * from TblEmployee   WHERE MachinCode='" & MachinCode & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    If Rs3.RecordCount > 0 Then
      
         project_id = IIf(Not IsNull(Rs3("project_id").Value), Rs3("project_id").Value, 0)
 DepartmentID = IIf(Not IsNull(Rs3("DepartmentID").Value), Rs3("DepartmentID").Value, 0)
 
  BranchID = IIf(Not IsNull(Rs3("BranchId").Value), Rs3("BranchId").Value, 0)
    Emp_ID = IIf(Not IsNull(Rs3("Emp_ID").Value), Rs3("Emp_ID").Value, 0)
    
 
'GroupID
 
    Else
 BranchID = 0
        Emp_ID = 0
 DepartmentID = 0
     project_id = 0
     
      End If
 
    Rs3.Close
 End Function


 Public Sub Main()
  StrAppRegPath = "bisegypt\SimpleAccounting"
SysSQLServerType = Val(GetSetting(StrAppRegPath, "ServerCon", "ServerType", 0)) '0 loca 1 not 2 rem
SysSQLServerName = GetSetting(StrAppRegPath, "ServerCon", "ServerName", "")
SysSQLServerTypeTechnical = GetSetting(StrAppRegPath, "ServerCon", "SysSQLServerTypeTechnical", "0")

SysSQLServerDataBaseName = GetSetting("Byte_DBS", "Setting", "DBPath", "Byte")

     SysSQLServerUserId = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserId", "salim")
    SysSQLServerUserpassword = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserpassword", "salim")







'frmMain.Show
FRMTRansferData3.Show

 End Sub
 
 
 Public Function SQLDate(ConvertDate As Date, Optional BolPutChar As Boolean = False, _
Optional BolPutSep As Boolean = True) As String

Dim StrTemp As String
Dim StrRes As String
Dim IntMonthB As Integer
Dim IntMonthE As Integer
Dim IntDay As Integer
Dim IntMonth As Integer
Dim IntYear As Integer
Dim StrMonthPrev As String
Dim StrTempqq As String

    IntDay = Day(ConvertDate)
    IntMonth = Month(ConvertDate)
    IntYear = Year(ConvertDate)
    StrMonthPrev = Choose(IntMonth, "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    If BolPutSep = True Then
        StrRes = "" & Format(IntDay, "00") & "-" & StrMonthPrev & "-" & IntYear & ""
    Else
        StrRes = "" & Format(IntDay, "00") & " " & StrMonthPrev & " " & IntYear & ""
    End If
    If BolPutChar = False Then
        SQLDate = StrRes
    ElseIf BolPutChar = True Then
   '     If SysDataBaseType = AccessDataBase Then
   '         StrTemp = "#" & StrRes & "#"
   '     ElseIf SysDataBaseType = SQLServerDataBase Then
            StrTemp = "'" & StrRes & "'"
   '     End If
        SQLDate = StrTemp
    End If
 End Function

Public Function new_id(tablename As String, _
                       FieldName As String, _
                       str_code As String, _
                       Optional serial As Boolean = False, _
                       Optional StrWhere As String = "", Optional ByVal mCn As ADODB.Connection = Nothing) As String
    'This Function to
    'Get the New ID and Serials
    Dim My_SQL As String
    Dim Lngid As Long
    Dim Rs_Temp As New ADODB.Recordset
        If mCn Is Nothing Then
            Set mCn = Cn
        End If
             My_SQL = "select max(cast(isnull(" & FieldName & ",0) as float )) as max_n "
            My_SQL = My_SQL + "From " & tablename & ""

            If StrWhere <> "" Then
                My_SQL = My_SQL + " Where " & StrWhere & " AND isnumeric(" & FieldName & ")=1"
            End If

            Rs_Temp.Open My_SQL, mCn, adOpenStatic, adLockReadOnly, adCmdText

            If IsNull(Rs_Temp("max_n").Value) Then
                new_id = "1"
            Else
                new_id = CStr(Val(Rs_Temp("max_n").Value) + 1)
            End If

            Set Rs_Temp = Nothing
 

End Function



  Public Function Voucher_codingByUser(my_branch As Integer, _
                               date1 As Date, _
                               Sanad_No As Integer, _
                               NoteType As Integer, _
                               Optional departement_name As Integer = 1, _
                               Optional Transaction_Type As Integer = 0, _
                               Optional Prefix As String = "", Optional StoreId As Integer = 0, Optional BillType As Integer = 0, Optional MosemID As Double = 0, Optional ByVal mTableName As String = "", Optional ByVal mUserId As Long = 0) As String
    
    
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

Dim mNoOfUser As Integer
Dim mUserIdSerial As String
Dim mLenUser As Integer

Dim mFormatUser As String

mFormatUser = ""
Dim mm As Integer
mm = 1
For mm = 1 To NoOFDigitUserTrans
    If mUserId < 9 And NoOFDigitUserTrans > 1 And NoOFDigitUserTrans > mm Then
        mFormatUser = mFormatUser & "0"
    ElseIf mUserId > 9 And mUserId < 100 And NoOFDigitUserTrans > 1 And NoOFDigitUserTrans > mm + 1 Then
        mFormatUser = mFormatUser & "0"
    ElseIf mUserId > 99 And NoOFDigitUserTrans > 1 And NoOFDigitUserTrans > mm + 2 Then
        mFormatUser = mFormatUser & "0"
    End If

Next

'If mUserId > 9 And mUserId < 100 Then
    mUserIdSerial = mFormatUser & mUserId
'End If


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
        
        storepadding = StoreDigit - 1
        If YearDigit = 0 Then YearDigit = 4
        
branchpadidng = BranchDigit - 1

 

        If StoreCoding = True Then
                    If StoreId <> 0 Then
                      storecode = getStoreCoding(StoreId)
                    End If
        End If
        
        
    End If
    
  
    Dim mWhere4 As String
    Dim mWhereUser As String
    Dim mWhereUser2 As String
    
    If IsSerialByUserTrans Then
       
            mWhere4 = "SUBSTRING(CAST(cast(NoteSerial1 AS BIGINT) AS VARCHAR(100)),2 , " & NoOFDigitUserTrans & ") = " & mUserId
           mWhereUser = mWhereUser & " AND " & mUserId & "  IN ("
           mWhereUser = mWhereUser & " SELECT UserID FROM DOUBLE_ENTREY_VOUCHERS AS dev WHERE dev.Notes_ID = Notes.NoteID)"
            mWhereUser2 = " And UserID = " & mUserId
           

    Else
        mUserIdSerial = ""
       ' mWhere4 = mWhere4 & " and UserID = " & mUserId
    End If
    
    
   Dim mWhere3 As String
    If my_branch > 9 Then
        mWhere3 = " SUBSTRING(CAST(NoteSerial1 AS VARCHAR(50))," & NoOFDigitUserTrans + 2 & ", 2) = " & my_branch
    Else
        mWhere3 = " SUBSTRING(CAST(NoteSerial1 AS VARCHAR(50)), " & NoOFDigitUserTrans + 2 & ", 1) = " & my_branch
    End If
   ' mWhere3 = " 1 = 1 "
    If numbering_type = 1 Then ' «·Ì
        sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and   NoteType=" & NoteType ' & " and   numbering_type1=" & numbering_type
        sql = sql & mWhereUser
        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and not(BTCashAccountcode is null )"
            sql = sql & mWhereUser
            
            
        End If
   
        If Sanad_No = 1 Then
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
        End If
   
        If Sanad_No = 4 Then
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            End If
            sql = sql & mWhereUser
        End If
   
        ' ﬂÊÌœ ”‰œ «· À”Ìÿ ‰›” ”‰œ «·ﬁ»÷
         If Sanad_No = 25 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 18) " '  and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
            
        End If
   
        If Sanad_No = 2 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 18 )" ' And numbering_type1 = " & numbering_type"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType  ' & " and   numbering_type1=" & numbering_type
            End If
            sql = sql & mWhereUser
        End If
   
        If Sanad_No = 26 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
        End If
   
        If Sanad_No = 2 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type
            End If
            sql = sql & mWhereUser
        End If

        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ «·’—›
   
        If Sanad_No = 1 Then
            If ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 14) " ' and   numbering_type1=" & numbering_type
                sql = sql & mWhereUser
            End If
        End If
   
        If Sanad_No = 16 Then
            If ExpensesCoding2 = True Then
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
            
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 9 or NoteType= 10)  "
                sql = sql & mWhereUser
        End If
        
          If Sanad_No = 64 Then
                sql = "select max(CAST (NoteSerial1 AS BigInt))as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   "
                sql = sql & mWhereUser2
        End If
        
            If Sanad_No = 65 Then
            If AllowProjectBill2Serial = True Then
            If BillType = 1 Then
                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1 "
                Else
                sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0 "
                End If
            Else
            
         sql = "select max(CAST (NoteSerial1 AS BigInt)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "   "
        End If
        End If
        If Sanad_No = 66 Then
                sql = "select max(CAST (NoteSerial1 AS BigInt))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  "
                sql = sql & "  and ( Transaction_Type=990 or Transaction_Type=18)"
                sql = sql & mWhereUser2
        End If
          If Sanad_No = 67 Then
                sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  "
                sql = sql & "  and  (Transaction_Type=66 or Transaction_Type=991) "
                sql = sql & mWhereUser2
        End If
          If Sanad_No = 68 Then
                sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  "
                sql = sql & "  and  (ImportExport=0 ) "
                sql = sql & mWhereUser2
        End If
        If Sanad_No = 69 Then
                sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  "
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
              '  Sql = "select max(CAST (dbo.TblContractInstallments.NoteSerial1 AS FLOAT))   as last_sand_no from  TblContractInstallments where  branch_no= " & my_branch & "  "
                sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS FLOAT)) as last_sand_no "
                sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
                sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
                sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")"
        End If
         If Sanad_No = 76 Then
                sql = "select max (NoteSerial1 )  as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "  "
                  sql = sql & " and " & mWhere3
        End If
        If Transaction_Type <> 0 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser2
        End If
        If Transaction_Type <> 0 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
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
  
          If mTableName <> "" Then
                sql = "select max (NoteSerial1 ) as last_sand_no from  " & mTableName & "  where  BranchId= " & my_branch & "      and year(recordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3

        End If

        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rs3("last_sand_no").Value) Then
        
            If end_at = 0 Then end_at = Val(Rs3("last_sand_no").Value) + 1
               
            If Rs3("last_sand_no").Value >= end_at Then
                Voucher_codingByUser = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 2 Then ' „ ’· ‘Â—Ì

        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
          sql = sql & " and " & mWhere3
          sql = sql & mWhereUser
        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and not(BTCashAccountcode is null )"
              sql = sql & " and " & mWhere3
              sql = sql & mWhereUser
        End If
    
        If Sanad_No = 1 Then
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        End If
    
        If Sanad_No = 4 Then
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
                  sql = sql & " and " & mWhere3
                  sql = sql & mWhereUser
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
                  sql = sql & " and " & mWhere3
                  sql = sql & mWhereUser
            End If
        
        End If
    
        If Sanad_No = 25 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                sql = sql & " and " & mWhere3
                sql = sql & mWhereUser
            End If
        End If
    
        If Sanad_No = 2 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)   and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
                  sql = sql & " and " & mWhere3
                  sql = sql & mWhereUser
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "       and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
                  sql = sql & " and " & mWhere3
                  sql = sql & mWhereUser
            End If
        
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            If ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
            sql = sql & mWhereUser
            End If
        End If
    
        If Sanad_No = 16 Then
            If ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)     and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
                  sql = sql & " and " & mWhere3
                  sql = sql & mWhereUser
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
                  sql = sql & " and " & mWhere3
                  sql = sql & mWhereUser
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
        sql = sql & " and " & mWhere3
        
        End If
        
        
        'TblContract  Branch_NO
    'Dim stockSettelmentsstr As String
    'stockSettelmentsstr = ""
        
        
      If Sanad_No = 62 Then
           
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 9 or NoteType= 10)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3
           sql = sql & mWhereUser
        End If
        ''////
       If Sanad_No = 64 Then
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                  sql = sql & " and " & mWhere3
        End If
              If Sanad_No = 66 Then
        sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR  Transaction_Type=18)   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
          sql = sql & " and " & mWhere3
          sql = sql & mWhereUser2
        End If
        If Sanad_No = 67 Then
        sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
          sql = sql & " and " & mWhere3
          sql = sql & mWhereUser2
        End If
        If Sanad_No = 68 Then
        sql = "select max  max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=0)   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
          sql = sql & " and " & mWhere3
          
        End If
        If Sanad_No = 69 Then
        sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  TblTransacRegistr where  BrnchID= " & my_branch & "  and   (ImportExport=1)   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
          sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 70 Then
        sql = "select max (NoteSerial1 ) as last_sand_no from  tblbookingrequest where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
          sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 71 Then
        sql = "select max (NoteSerial1 ) as last_sand_no from  tblbookingrequest2 where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(SDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(SDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
          sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 72 Then
        sql = "select max (NoteSerial1 ) as last_sand_no from  TblDeported where  BranchID= " & my_branch & "  and   (SeasonsID=" & MosemID & ")   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
          sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 74 Then
        sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "   and  notetype=370   and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
          sql = sql & " and " & mWhere3
          sql = sql & mWhereUser2
        End If
         
        If Sanad_No = 76 Then
        sql = "select max (NoteSerial1 ) as last_sand_no from  TblTravDueK where  BranchId= " & my_branch & "      and year(recordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
          sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 75 Then
                   sql = " SELECT     max(CAST (dbo.TblContractInstallments.NoteSerial1 AS FLOAT))as last_sand_no  "
                sql = sql & "        FROM         dbo.TblContractInstallments INNER JOIN"
                sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo"
                sql = sql & "  Where (dbo.TblContract.branch_no = " & my_branch & ")and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                
        'Sql = "select max (NoteSerial1 ) as last_sand_no from  notes_all where  branch_no= " & my_branch & "     and year(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(dbo.TblContractInstallments.Installdate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If
        
             
                
            If Sanad_No = 65 Then
            If AllowProjectBill2Serial = True Then
            If BillType = 1 Then
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                Else
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                  'sql = sql & " and " & mWhere3
                End If
            Else
            
         sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
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
                   sql = "select max(  (NoteSerial1    )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & "   Or Transaction_Type= 992)  and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                     sql = sql & " and " & mWhere3
                     
            Else
                      sql = "select max(  (NoteSerial1    )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                        sql = sql & " and " & mWhere3
            End If
            sql = sql & mWhereUser2
        Else
                 If BranchDigit > 1 Then
                 
                 
              '   sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & mId(Format$(date1, "dd/mm/yyyy"), 4, 2)
         'edit edit salim here
              If Transaction_Type = 10 Then
                sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   (Transaction_Type=" & Transaction_Type & "  Or Transaction_Type= 992)   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                  sql = sql & " and " & mWhere3
            Else
                sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                  sql = sql & " and " & mWhere3
            End If
            sql = sql & mWhereUser2
            'edit edit salim here
            
                 Else
                 If Transaction_Type = 10 Then
                          sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and  ( Transaction_Type=" & Transaction_Type & "   Or Transaction_Type= 992)  and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
                            sql = sql & " and " & mWhere3
                  Else
                                            sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
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
        
        sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        
                If mTableName <> "" Then
                sql = "select max (NoteSerial1 ) as last_sand_no from  " & mTableName & "  where  BranchId= " & my_branch & "      and year(recordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
            sql = sql & " and " & mWhere3

        End If

        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

Dim startrreadding As Integer
Dim noofreadinchar  As Integer
        If Not IsNull(Rs3("last_sand_no").Value) Then
          If StoreCoding = True And StoreId <> 0 Then
                   startrreadding = BranchDigit + StoreDigit + YearDigit + noOfDigit
                   noofreadinchar = startrreadding - 1
                 '    If YearDigit = 2 Then
                     
                     
                 '      no = Mid(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
                 '    Else
                     
                 '    no = Mid(Rs3("last_sand_no").value, 10, Len(Rs3("last_sand_no").value) - 9)
                 '   End If
               
           Else
           
           startrreadding = BranchDigit + YearDigit + noOfDigit
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
                Voucher_codingByUser = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ
           If Sanad_No = 64 Then
                sql = "select max(CAST (NoteSerial1 AS FLOAT))as last_sand_no from  TblOtheExpensAqar    where  BranchID= " & my_branch & "   and year(RecordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
        
            If Sanad_No = 65 Then
            If AllowProjectBill2Serial = True Then
            If BillType = 1 Then
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=1   and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
                Else
                sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "  and bill_to=0  and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
                End If
            Else
            
         sql = "select max(CAST (NoteSerial1 AS FLOAT)) as last_sand_no from  project_billl    where  Branch_NO= " & my_branch & "    and year(bill_date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
         sql = sql & " and " & mWhere3
        End If
        End If
        
 
        sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        sql = sql & mWhereUser
        sql = sql & " and " & mWhere3
        If Sanad_No = 5 Then
            sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "   and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and not(BTCashAccountcode is null )"
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
      
        If Sanad_No = 1 Then
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)     and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & mWhereUser
                sql = sql & " and " & mWhere3
            End If
        
        End If
     
        If Sanad_No = 4 Then
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)  and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            End If
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
          
        If Sanad_No = 25 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)      and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & mWhereUser
                sql = sql & " and " & mWhere3
            End If
        
        End If
     
        If Sanad_No = 2 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
            End If
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›”  —ﬁ”„ ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            If ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)   and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
                sql = sql & mWhereUser
                sql = sql & " and " & mWhere3
            End If
        
        End If
     
        If Sanad_No = 16 Then
            If ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)        and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            End If
            sql = sql & mWhereUser
            sql = sql & " and " & mWhere3
        End If
    
    
                         If Sanad_No = 50 Then
      
        '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
   'sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
   sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  BranchID= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
   sql = sql & " and " & mWhere3
        End If
                          If Sanad_No = 58 Then
      
        '        sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains    where  BranchID= " & my_branch
   'sql = "select max(NoteSerial1) as last_sand_no from  TblCarBillMentains where  branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)   '& "  )"
   sql = "select max(NoteSerial1) as last_sand_no from  TblExchange where  BranchID= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
   sql = sql & " and " & mWhere3
        End If
        
        If Sanad_No = 60 Then
    sql = "select max(NoteSerial1) as last_sand_no from  TblContract where  Branch_NO= " & my_branch & " and year(ContDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
   
        End If
        
        
        If Sanad_No = 62 Then
           
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 9 or NoteType= 10)     and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
                sql = sql & mWhereUser
                sql = sql & " and " & mWhere3
        End If
        If Sanad_No = 66 Then
         sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   ( Transaction_Type=990 OR Transaction_Type=18)  and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
         sql = sql & mWhereUser2
         sql = sql & " and " & mWhere3
         End If
        If Sanad_No = 67 Then
         sql = "select  max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and  (Transaction_Type=66 or Transaction_Type=991)   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
         sql = sql & mWhereUser2
         sql = sql & " and " & mWhere3
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
         sql = sql & mWhereUser2
         sql = sql & " and " & mWhere3
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
 
        If Not IsNull(Rs3("last_sand_no").Value) Then
                  If StoreCoding = True And StoreId <> 0 Then
                        
                                           
                  startrreadding = BranchDigit + StoreDigit + YearDigit + 1
                   noofreadinchar = startrreadding - 1
                   'If YearDigit = 2 Then
                   '            no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                   '         Else
                   '         no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                   '         End If
         
         Else
         If Val(getNoOfBranches) > 9 Then
         
         If Mid(Rs3("last_sand_no").Value, 1, 1) = "0" Then
         
         startrreadding = BranchDigit + YearDigit + 1
             Else
             
           If Val(my_branch) > 9 Then
       startrreadding = BranchDigit + YearDigit + 1
       Else
       startrreadding = BranchDigit + YearDigit
       End If
       
             
             End If
             
         Else
         startrreadding = BranchDigit + YearDigit
             'noofreadinchar = startrreadding
             If Transaction_Type <> 0 Then
                        If BranchDigit = 1 Then
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
                Voucher_codingByUser = "error"
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
 
brancHcode = zeropadding(CStr(my_branch), Int(BranchDigit))
storecode = zeropadding(storecode, Int(StoreDigit))

    If numbering_type = 1 Then Voucher_codingByUser = auto_sanad_no: Exit Function
    If first_serial = True Then
        If auto_sanad_no <> "" Then
            
            If StoreCoding = True And StoreId <> 0 Then
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
          If StoreCoding = True And StoreId <> 0 Then
           ' Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
          Voucher_codingByUser = "1" & mUserIdSerial & brancHcode & storecode & auto_sanad_no
           Else
             Voucher_codingByUser = "1" & mUserIdSerial & brancHcode & auto_sanad_no
           End If
    End If

End Function



 






 





 Public Function Voucher_coding(my_branch As Integer, _
                               date1 As Date, _
                               Sanad_No As Integer, _
                               NoteType As Integer, _
                               Optional departement_name As Integer = 1, _
                               Optional Transaction_Type As Integer = 0, _
                               Optional Prefix As String = "", Optional StoreId As Integer = 0, Optional BillType As Integer = 0, Optional MosemID As Double = 0, Optional ByVal mTableName As String = "", Optional ByVal mUserId As Long = 0) As String
    
    
      On Error Resume Next
If my_branch = 0 Then Exit Function
    Dim start_at As Integer
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

IsSerialByUserTrans = True
If IsSerialByUserTrans Then
    
    Voucher_coding = Voucher_codingByUser(my_branch, date1, Sanad_No, NoteType, departement_name, Transaction_Type, Prefix, StoreId, BillType, MosemID, mTableName, mUserId)
    Exit Function
End If


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
        Zeros = IIf(IsNull(rs("zeros").Value), 0, rs("zeros").Value)
        StoreCoding = IIf(IsNull(rs("StoreCoding").Value), 0, rs("StoreCoding").Value)
        YearDigit = IIf(IsNull(rs("YearDigit").Value), 4, rs("YearDigit").Value)
        
        storepadding = StoreDigit - 1
        If YearDigit = 0 Then YearDigit = 4
        
branchpadidng = BranchDigit - 1

 

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
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type
            End If
        End If
   
        If Sanad_No = 4 Then
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 3 or NoteType= 5 ) " ' and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type '& " and  (BTCashAccountcode is null )"
            End If
        End If
   
        ' ﬂÊÌœ ”‰œ «· À”Ìÿ ‰›” ”‰œ «·ﬁ»÷
         If Sanad_No = 25 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 18) " '  and   numbering_type1=" & numbering_type
            End If
        End If
   
        If Sanad_No = 2 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 18 )" ' And numbering_type1 = " & numbering_type"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType  ' & " and   numbering_type1=" & numbering_type
            End If
        End If
   
        If Sanad_No = 26 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
            End If
        End If
   
        If Sanad_No = 2 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    (NoteType= 4 or NoteType= 19) " ' and   numbering_type1=" & numbering_type
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and    NoteType=" & NoteType '& " and   numbering_type1=" & numbering_type
            End If
        End If

        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ «·’—›
   
        If Sanad_No = 1 Then
            If ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes    where  branch_no= " & my_branch & "  and     (NoteType= 3 or NoteType= 14) " ' and   numbering_type1=" & numbering_type
            End If
        End If
   
        If Sanad_No = 16 Then
            If ExpensesCoding2 = True Then
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
            If AllowProjectBill2Serial = True Then
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
        
        If Transaction_Type <> 0 Then
            sql = "select max(CAST (NoteSerial1 AS FLOAT))  as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type
  
        End If
   
               If Prefix = "" Then
                sql = sql & "  and   Prefix is null"
 
            Else
                sql = sql & "  and   Prefix='" & Prefix & "'"
            End If

       If mTableName <> "" Then
                sql = "select max (NoteSerial1 ) as last_sand_no from  " & mTableName & "  where  BranchId= " & my_branch & "      and year(recordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(recordDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
          sql = sql & " and " & mWhere3

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
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        
            End If
        End If
    
        If Sanad_No = 4 Then
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) '& " and (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)  '& " and (BTCashAccountcode is null )"
            End If
        
        End If
    
        If Sanad_No = 25 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        
            End If
        End If
    
        If Sanad_No = 2 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)   and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "       and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & " and (BTCashAccountcode is null )"
            End If
        
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›” ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            If ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        
            End If
        End If
    
        If Sanad_No = 16 Then
            If ExpensesCoding2 = True Then
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
        
            If Sanad_No = 65 Then
            If AllowProjectBill2Serial = True Then
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
        If BranchDigit > 1 Then
        sql = "select max(  (NoteSerial1  )) as last_sand_no from  Transactions where  BranchId= " & my_branch & "  and   Transaction_Type=" & Transaction_Type & "   and year(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(Transaction_Date)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
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
       '     Cn.DefaultDatabase = ServerDb
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

Dim startrreadding As Integer
Dim noofreadinchar  As Integer
        If Not IsNull(Rs3("last_sand_no").Value) Then
          If StoreCoding = True And StoreId <> 0 Then
                   startrreadding = BranchDigit + StoreDigit + YearDigit + 3
                   noofreadinchar = startrreadding - 1
                 '    If YearDigit = 2 Then
                     
                     
                 '      no = Mid(Rs3("last_sand_no").value, startrreadding, Len(Rs3("last_sand_no").value) - noofreadinchar)
                 '    Else
                     
                 '    no = Mid(Rs3("last_sand_no").value, 10, Len(Rs3("last_sand_no").value) - 9)
                 '   End If
               
           Else
           
           startrreadding = BranchDigit + YearDigit + 3
           If Transaction_Type = 0 And (Sanad_No <> 66 And Sanad_No <> 67) Then
startrreadding = 1 + YearDigit + 3
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
            If AllowProjectBill2Serial = True Then
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
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)     and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            End If
        
        End If
     
        If Sanad_No = 4 Then
            If ExpensesCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 5)  and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & "  and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & "  "
            End If
        
        End If
          
        If Sanad_No = 25 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)      and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            End If
        
        End If
     
        If Sanad_No = 2 Then
            If InstallmntsvchrCoding = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 4 or NoteType= 18)    and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
            Else
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   NoteType=" & NoteType & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
            End If
        
        End If
    
        '”‰œ«  «· ÕÊÌ· ‰›”  —ﬁ”„ ”‰œ«  «·’—›
        If Sanad_No = 1 Then
            If ExpensesCoding2 = True Then
                sql = "select max(NoteSerial1) as last_sand_no from  Notes where  branch_no= " & my_branch & "  and   (NoteType= 3 or NoteType= 14)   and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
            End If
        
        End If
     
        If Sanad_No = 16 Then
            If ExpensesCoding2 = True Then
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
         
        '
        
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
                        
                                           
                  startrreadding = BranchDigit + StoreDigit + YearDigit + 1
                   noofreadinchar = startrreadding - 1
                   'If YearDigit = 2 Then
                   '            no = Mid(Rs3("last_sand_no").value, 6, Len(Rs3("last_sand_no").value) - 5)
                   '         Else
                   '         no = Mid(Rs3("last_sand_no").value, 8, Len(Rs3("last_sand_no").value) - 7)
                   '         End If
         
         Else
         If Val(getNoOfBranches) > 9 Then
         
         If Mid(Rs3("last_sand_no").Value, 1, 1) = "0" Then
         
         startrreadding = BranchDigit + YearDigit + 1
             Else
             
           If Val(my_branch) > 9 Then
       startrreadding = BranchDigit + YearDigit + 1
       Else
       startrreadding = BranchDigit + YearDigit
       End If
       
             
             End If
             
         Else
         startrreadding = BranchDigit + YearDigit
             'noofreadinchar = startrreadding
             If Transaction_Type <> 0 Then
                        If BranchDigit = 1 Then
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
 
brancHcode = zeropadding(CStr(my_branch), Int(BranchDigit))
storecode = zeropadding(storecode, Int(StoreDigit))

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

 
 Public Function Note_codingNew(my_branch As Integer, _
                               date1 As Date, _
                               Sanad_No As Integer, _
                               NoteType As Integer, _
                               Optional departement_name As Integer = 1, _
                               Optional Transaction_Type As Integer = 0, _
                               Optional Prefix As String = "", Optional StoreId As Integer = 0) As String
    
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    Dim auto_sanad_no As String
    Dim NO As Integer
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
        Zeros = IIf(IsNull(rs("zeros").Value), 0, rs("zeros").Value)
     '   StoreCoding = IIf(IsNull(rs("StoreCoding").value), 0, rs("StoreCoding").value)
        YearDigit = IIf(IsNull(rs("YearDigit").Value), 4, rs("YearDigit").Value)
        If noOfDigit = 0 Then noOfDigit = 3
      '  storepadding = StoreDigit - 1
        If YearDigit = 0 Then YearDigit = 4
        
branchpadidng = BranchDigit - 1
    
        
        
    End If
If Len(BRANCHNO) < 10 Then
branchpadidng = 1
ElseIf Len(BRANCHNO) >= 10 And Len(BRANCHNO) <= 99 Then
branchpadidng = 2
Else
branchpadidng = 3
End If
    If numbering_type = 1 Then ' «·Ì
        sql = "select max(NoteSerial) as last_sand_no from  Notes    where  NoteType<>1 AND   branch_no= " & my_branch
  
'   Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rs3("last_sand_no").Value) Then
        
              startrreadding = branchpadidng + 1
                   noofreadinchar = startrreadding - 1
                   
 
           NO = Mid(Rs3("last_sand_no").Value, startrreadding, Len(Rs3("last_sand_no").Value) - noofreadinchar)
           
            If end_at = 0 Then end_at = Val(Rs3("last_sand_no").Value) + 1: GoTo xl
               
            If Rs3("last_sand_no").Value >= end_at Then
                Note_codingNew = "error"
                Exit Function
            End If
        End If

xl:
             
    ElseIf numbering_type = 2 Then ' „ ’· ‘Â—Ì

        sql = "select max(NoteSerial) as last_sand_no from  Notes where     branch_no= " & my_branch & " and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
   Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
       If Not IsNull(Rs3("last_sand_no").Value) Then
           startrreadding = branchpadidng + YearDigit + noOfDigit
                   noofreadinchar = startrreadding - 1
  
             NO = Mid(Rs3("last_sand_no").Value, startrreadding, Len(Rs3("last_sand_no").Value) - noofreadinchar)
             
                          If end_at = 0 Then end_at = NO + 1
                        If NO >= end_at Then
                            Note_codingNew = "error"
                            Exit Function
                          End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ
 
        sql = "select max(NoteSerial) as last_sand_no from  Notes where  branch_no= " & my_branch & "and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If Not IsNull(Rs3("last_sand_no").Value) Then
     
      startrreadding = branchpadidng + YearDigit + 1
                   noofreadinchar = startrreadding - 1
                   
 
           NO = Mid(Rs3("last_sand_no").Value, startrreadding, Len(Rs3("last_sand_no").Value) - noofreadinchar)
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

    If Rs3.RecordCount = 0 Or IsNull(Rs3("last_sand_no").Value) Then
        first_serial = True

        If numbering_type = 0 Then
                 'ÌœÊÌ
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Format(start_at, String(Askcount, "0"))
   
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
            auto_sanad_no = Format((NO + 1), String(Askcount, "0"))
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
 
brancHcode = zeropadding(CStr(my_branch), Int(BranchDigit))
'storecode = zeropadding(storecode, Int(StoreDigit))

'    If numbering_type = 1 Then Note_codingNew = auto_sanad_no: Exit Function
    
    If first_serial = True Then
        If auto_sanad_no <> "" Then
            
            If StoreCoding = True And StoreId <> 0 Then
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
          If StoreCoding = True And StoreId <> 0 Then
           ' Voucher_coding = my_branch & storeADDZero & StoreID & auto_sanad_no
          Note_codingNew = brancHcode & storecode & auto_sanad_no
           Else
             Note_codingNew = brancHcode & auto_sanad_no
           End If
    End If

End Function
Public Function Notes_coding(my_branch As Integer, _
                             date1 As Date, _
                             Optional departement_name As Integer = 1) As String
    On Error Resume Next
    Dim start_at As Double
    Dim end_at As Single
    Dim auto_sanad_no As String
    Dim NO As Single
    Dim numbering_type As Integer
    auto_sanad_no = ""

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset

    Dim sql As String
    Dim i As Integer
 
If IsSerialByUserVouch Then
    Notes_coding = Notes_codingByUser(my_branch, date1, departement_name)
    Exit Function
End If
 
'    JLCodeBasedOnBranch = False
'**********************
'JLCodeBasedOnBranch = True
'If JLCodeBasedOnBranch = True Then
If JLCodeBasedOnBranch = True Then
Notes_coding = Note_codingNew(my_branch, date1, 0, 200)

Exit Function

End If
'******************


    sql = "select * from sanad_numbering where branch_no=" & my_branch & " and  sanad_no=0"
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = IIf(IsNull(rs("numbering_id").Value), 0, rs("numbering_id").Value)
        start_at = IIf(IsNull(rs("start_at").Value), 0, rs("start_at").Value)
        end_at = IIf(IsNull(rs("end_at").Value), 0, rs("end_at").Value)

    End If

     If numbering_type = 1 Then
        If JLCodeBasedOnBranch = False Then
            sql = "select max(NoteSerial) as last_sand_no from  Notes WHERE NoteType<>1"  'where      numbering_type=" & numbering_type
        Else
            sql = "select max(NoteSerial) as last_sand_no from  Notes where    NoteType<>1 AND  branch_no= " & my_branch '& "  and     numbering_type=" & numbering_type
        End If

        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rs3("last_sand_no").Value) Then
        
            If end_at = 0 Then end_at = Val(Rs3("last_sand_no").Value) + 1: GoTo XL1
               
            If Rs3("last_sand_no").Value >= end_at Then
                Notes_coding = "error"
                Exit Function
            End If
        End If
XL1:
    ElseIf numbering_type = 2 Then '„ ’· ”‰‘Â—ÌÊÌ
 
        If JLCodeBasedOnBranch = False Then
            sql = "select max(NoteSerial) as last_sand_no from  Notes where    year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
            If JLCodeBasedOnBranch = False Then
sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS Bigint) AS varchar(100)), 5, 2) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 1, 2))"
sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS Bigint) AS varchar(100)), 1, 4) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 7, 4))"
End If
        Else
            sql = "select max(NoteSerial) as last_sand_no from  Notes where  branch_no= " & my_branch & " and sanad_year=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and sanad_month=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
        End If

        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rs3("last_sand_no").Value) Then
            NO = Mid(Rs3("last_sand_no").Value, 7, Len(Rs3("last_sand_no").Value) - 6)

            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Notes_coding = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ

        If JLCodeBasedOnBranch = False Then
            sql = "select max(NoteSerial) as last_sand_no from  Notes where     year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
               If JLCodeBasedOnBranch = False Then
            sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS Bigint) AS varchar(100)), 1, 4) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 7, 4))"
            End If
        Else
            sql = "select max(NoteSerial) as last_sand_no from  Notes where    branch_no= " & my_branch & "  and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
        End If
  
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If Not IsNull(Rs3("last_sand_no").Value) Then
            NO = Mid(Rs3("last_sand_no").Value, 5, Len(Rs3("last_sand_no").Value) - 4)

            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Notes_coding = "error"
                Exit Function
            End If
        End If
 
    End If

    Dim Askcount As Integer
    Askcount = Ked_digit ' GetSetting(StrAppRegPath, "Setting", "Count_Ked_digit", 0)
         
    Dim first_serial As Boolean

    If Rs3.RecordCount = 0 Or IsNull(Rs3("last_sand_no").Value) Then
        first_serial = True

        If numbering_type = 0 Then
                 
        ElseIf numbering_type = 1 Then
            auto_sanad_no = start_at
        ElseIf numbering_type = 2 Then
          '  auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            auto_sanad_no = Year(date1) & Format(Month(date1), String(2, "0")) & Format(start_at, String(Askcount, "0"))
            
            ' year(date1) & Format(Month(date1), String(2, "0"))
        ElseIf numbering_type = 3 Then
        '    auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(start_at, String(Askcount, "0"))
        auto_sanad_no = Year(date1) & Format(start_at, String(Askcount, "0"))
        
        End If

    Else

        If numbering_type = 0 Then
                
        ElseIf numbering_type = 1 Then
            auto_sanad_no = Rs3("last_sand_no").Value + 1
        ElseIf numbering_type = 2 Then
              
            NO = Mid(Rs3("last_sand_no").Value, 7, Len(Rs3("last_sand_no").Value) - 6)
            'auto_sanad_no = Mid(rs3("last_sand_no").value, 1, 6) & Format((NO + 1), String(Askcount, "0"))
        auto_sanad_no = Year(date1) & Format(Month(date1), String(2, "0")) & Format((NO + 1), String(Askcount, "0"))
        ElseIf numbering_type = 3 Then
         
            NO = Mid(Rs3("last_sand_no").Value, 5, Len(Rs3("last_sand_no").Value) - 4)
            'auto_sanad_no = Mid(rs3("last_sand_no").value, 1, 4) & Format((NO + 1), String(Askcount, "0"))
            auto_sanad_no = Year(date1) & Format((NO + 1), String(Askcount, "0"))
        End If
 
    End If

    Rs3.Close
    'If first_serial = False Then
    'auto_sanad_no = Mid(auto_sanad_no, 2, Len(auto_sanad_no))
    'End If
    'Notes_coding = my_branch & auto_sanad_no
    Notes_coding = auto_sanad_no
  
End Function

Public Function getStoreCoding(StoreId As Integer) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
 
    sql = " SELECT   * "
    sql = sql & " from   dbo.TblStore"
    sql = sql & " WHERE     (StoreID = " & StoreId & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        getStoreCoding = IIf(IsNull(rs("Code").Value), 0, rs("Code").Value)
    Else
        getStoreCoding = ""
    End If
rs.Close
Set rs = Nothing
End Function

Public Function zeropadding(str As String, _
                        noofchar As Integer) As String
    Dim lenstr As Integer
    Dim DIFF As Integer
    Dim newStr As String
newStr = ""
    lenstr = Len(str)

    If noofchar > lenstr Then
        DIFF = noofchar - lenstr
        newStr = String(DIFF, "0")
     
                    
    End If
   zeropadding = newStr & str
End Function


Public Function LoadMainSystemOptions() As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim IntTemp As Integer
    Dim IntTemp1 As Integer

    Dim StrTemp As String
    Dim StrSQL  As String

      Set rs = New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
       StoreDigit = IIf(IsNull(rs("StoreDigit").Value), 1, (rs("StoreDigit").Value))
    BranchDigit = IIf(IsNull(rs("BranchDigit").Value), 1, (rs("BranchDigit").Value))
    
    LoadMainSystemOptions = True
    Exit Function
hErr:
    Msg = "Â‰«ﬂ Œÿ« ›Ï Load Main System Options"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    LoadMainSystemOptions = False
End Function



Public Function Notes_codingByUser(my_branch As Integer, _
                             date1 As Date, _
                             Optional departement_name As Integer = 1) As String
    On Error Resume Next
    Dim start_at As Double
    Dim end_at As Single
    Dim auto_sanad_no As String
    Dim NO As Single
    Dim numbering_type As Integer
    auto_sanad_no = ""

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset

    Dim sql As String
    Dim i As Integer
    Dim JLCodeBasedOnBranch As Boolean
    JLCodeBasedOnBranch = False
'**********************
'JLCodeBasedOnBranch = True

Dim mNoOfUser As Integer
Dim mUserIdSerial As String
Dim mLenUser As Integer

Dim mFormatUser As String

mFormatUser = ""
Dim mm As Integer
mm = 1
For mm = 1 To NoOFDigitUserVouc
    If mUserId < 9 And NoOFDigitUserVouc > 1 And NoOFDigitUserVouc > mm Then
        mFormatUser = mFormatUser & "0"
    ElseIf mUserId > 9 And mUserId < 100 And NoOFDigitUserVouc > 1 And NoOFDigitUserVouc > mm + 1 Then
        mFormatUser = mFormatUser & "0"
    ElseIf mUserId > 99 And NoOFDigitUserVouc > 1 And NoOFDigitUserVouc > mm + 2 Then
        mFormatUser = mFormatUser & "0"
    End If

Next

'If mUserId > 9 And mUserId < 100 Then
    mUserIdSerial = mFormatUser & mUserId
'End If

If JLCodeBasedOnBranch = True Then
    Notes_codingByUser = Note_codingNew(my_branch, date1, 0, 200)
    Exit Function

End If
 Dim mWhere4 As String
    
    If IsSerialByUserVouch Then
       
            mWhere4 = "SUBSTRING(CAST(cast(NoteSerial AS BIGINT) AS VARCHAR(100)),2 , " & NoOFDigitUserVouc & ") = " & mUserId
            mWhere4 = mWhere4 & " AND " & mUserId & "  IN ("
            mWhere4 = mWhere4 & " SELECT UserID FROM DOUBLE_ENTREY_VOUCHERS AS dev WHERE dev.Notes_ID = Notes.NoteID)"


    Else
        mUserIdSerial = ""
       ' mWhere4 = mWhere4 & " and UserID = " & mUserId
    End If


'******************
    Dim mWhere3 As String
    If my_branch > 9 Then
        mWhere3 = " SUBSTRING(CAST(NoteSerial AS VARCHAR(50)), 1, 2) = " & my_branch
    Else
        mWhere3 = " SUBSTRING(CAST(NoteSerial AS VARCHAR(50)), 1, 1) = " & my_branch
    End If

    sql = "select * from sanad_numbering where branch_no=" & my_branch & " and  sanad_no=0"
        
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = IIf(IsNull(rs("numbering_id").Value), 0, rs("numbering_id").Value)
        start_at = IIf(IsNull(rs("start_at").Value), 0, rs("start_at").Value)
        end_at = IIf(IsNull(rs("end_at").Value), 0, rs("end_at").Value)

    End If

     If numbering_type = 1 Then
        If JLCodeBasedOnBranch = False Then
            sql = "select max(NoteSerial) as last_sand_no from  Notes WHERE NoteType<>1"  'where      numbering_type=" & numbering_type
        Else
            sql = "select max(NoteSerial) as last_sand_no from  Notes where    NoteType<>1 AND  branch_no= " & my_branch '& "  and     numbering_type=" & numbering_type
             sql = sql & " and " & mWhere3
        End If
sql = sql & "  and   NoteType <>1 "
sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rs3("last_sand_no").Value) Then
        
            If end_at = 0 Then end_at = Val(Rs3("last_sand_no").Value) + 1: GoTo XL1
               
            If Rs3("last_sand_no").Value >= end_at Then
                Notes_codingByUser = "error"
                Exit Function
            End If
        End If
XL1:
    ElseIf numbering_type = 2 Then '„ ’· ”‰‘Â—ÌÊÌ
 
        If JLCodeBasedOnBranch = False Then
            sql = "select max(NoteSerial) as last_sand_no from  Notes where    year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and month(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
            If JLCodeBasedOnBranch = False Then
sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS bigint) AS varchar(100)), 5 + " & NoOFDigitUserVouc + 1 & " , 2) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 1, 2))"
sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS bigint) AS varchar(100)), " & NoOFDigitUserVouc + 2 & ", 4) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 7, 4))"
End If
        Else
            sql = "select max(NoteSerial) as last_sand_no from  Notes where  branch_no= " & my_branch & " and sanad_year=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & " and sanad_month=" & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2)
             sql = sql & " and " & mWhere3
        End If
sql = sql & "  and   NoteType <>1 "
sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(Rs3("last_sand_no").Value) Then
            NO = Mid(Rs3("last_sand_no").Value, 7 + NoOFDigitUserVouc + 1, Len(Rs3("last_sand_no").Value) - 6)

            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Notes_codingByUser = "error"
                Exit Function
            End If
        End If

    ElseIf numbering_type = 3 Then '„ ’· ”‰ÊÌ

        If JLCodeBasedOnBranch = False Then
            sql = "select max(NoteSerial) as last_sand_no from  Notes where     year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
               If JLCodeBasedOnBranch = False Then
            sql = sql & " and       (SUBSTRING(CAST(CAST(NoteSerial AS bigint) AS varchar(100)), " & NoOFDigitUserVouc + 1 & ", 4) = SUBSTRING(CONVERT(VARCHAR(10), NoteDate, 110), 7, 4))"
            End If
        Else
            sql = "select max(NoteSerial) as last_sand_no from  Notes where    branch_no= " & my_branch & "  and year(NoteDate)=" & Mid(Format$(date1, "dd/mm/yyyy"), 7, 4)
             sql = sql & " and " & mWhere3
        End If
  
    sql = sql & "  and   NoteType <>1 "
    sql = sql & " and " & IIf(mWhere4 = "", " 1 = 1 ", mWhere4)
        Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If Not IsNull(Rs3("last_sand_no").Value) Then
            NO = Mid(Rs3("last_sand_no").Value, 5 + NoOFDigitUserVouc + 1, Len(Rs3("last_sand_no").Value) - 4)

            If end_at = 0 Then end_at = NO + 1
            If NO >= end_at Then
                Notes_codingByUser = "error"
                Exit Function
            End If
        End If
 
    End If

    Dim Askcount As Integer
    Askcount = Ked_digit ' GetSetting(StrAppRegPath, "Setting", "Count_Ked_digit", 0)
         
    Dim first_serial As Boolean

    If Rs3.RecordCount = 0 Or IsNull(Rs3("last_sand_no").Value) Then
        first_serial = True

        If numbering_type = 0 Then
                 
        ElseIf numbering_type = 1 Then
            auto_sanad_no = "1" & mUserIdSerial & start_at
        ElseIf numbering_type = 2 Then
          '  auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Mid(Format$(date1, "dd/mm/yyyy"), 4, 2) & Format(start_at, String(Askcount, "0"))
            auto_sanad_no = "1" & mUserIdSerial & Year(date1) & Format(Month(date1), String(2, "0")) & Format(start_at, String(Askcount, "0"))
            
            ' year(date1) & Format(Month(date1), String(2, "0"))
        ElseIf numbering_type = 3 Then
        '    auto_sanad_no = Mid(Format$(date1, "dd/mm/yyyy"), 7, 4) & Format(start_at, String(Askcount, "0"))
        auto_sanad_no = "1" & mUserIdSerial & Year(date1) & Format(start_at, String(Askcount, "0"))
        
        End If

    Else

        If numbering_type = 0 Then
                
        ElseIf numbering_type = 1 Then
            auto_sanad_no = "1" & mUserIdSerial & Rs3("last_sand_no").Value + 1
        ElseIf numbering_type = 2 Then
              
            NO = Mid(Rs3("last_sand_no").Value, 7 + NoOFDigitUserVouc + 1, Len(Rs3("last_sand_no").Value) - 6)
            'auto_sanad_no = Mid(rs3("last_sand_no").value, 1, 6) & Format((NO + 1), String(Askcount, "0"))
        auto_sanad_no = "1" & mUserIdSerial & Year(date1) & Format(Month(date1), String(2, "0")) & Format((NO + 1), String(Askcount, "0"))
        ElseIf numbering_type = 3 Then
         
            NO = Mid(Rs3("last_sand_no").Value, 5 + NoOFDigitUserVouc + 1, Len(Rs3("last_sand_no").Value) - 4)
            'auto_sanad_no = Mid(rs3("last_sand_no").value, 1, 4) & Format((NO + 1), String(Askcount, "0"))
            auto_sanad_no = "1" & mUserIdSerial & Year(date1) & Format((NO + 1), String(Askcount, "0"))
        End If
 
    End If

    Rs3.Close
    'If first_serial = False Then
    'auto_sanad_no = Mid(auto_sanad_no, 2, Len(auto_sanad_no))
    'End If
    'Notes_coding = my_branch & auto_sanad_no
    Notes_codingByUser = auto_sanad_no
  
End Function
 




