Attribute VB_Name = "salahnew"

Dim StrFileName As String
Dim StrText As String
Dim Bas64 As ClsSupplierPrice

Dim FileName As String
Dim byteBuffer() As Byte
Dim strBuffer As String
Dim flgFile As Integer
Dim encBuffer As String
Dim decBuffer As String

' [?????] ????? ???????? mBranchID ?? ????? ??????
Public Sub SaveQRCode6(mTable As String, mmIDField As String, mmID As Long, ByVal NoteSerial1 As String, ByVal Transaction_Date As Date, mAmount1 As String, _
    ByRef Picture1 As PictureBox, Optional ByVal mAmountDisc As String = "", Optional ByVal mVat As String = "", Optional ByVal mTotalNet As String = "", _
    Optional ByVal mBranchID As Long = 0)

    '--- (?? ????? ?? ????????? ???????) ---
    Dim mmAmountDisc  As Double
    If val(mAmountDisc) > 0 Then
        mmAmountDisc = Round(mAmountDisc, 3)
    End If
    Dim mmVat As Double
    If val(mVat) > 0 Then
        mmVat = Round(mVat, 3)
    End If
    Dim mmTotalNet As Double
    If val(mTotalNet) > 0 Then
        mmTotalNet = Round(mTotalNet, 3)
    End If
    Dim folderPath As String
    Dim FileName As String
    Dim mmAmount1 As Double
    If val(mAmount1) > 0 Then
        mmAmount1 = Round(mAmount1, 3)
    End If
    
    If Not SystemOptions.IsQrCodePrint Then Exit Sub

    '--- [?? ????? ??? ????? ???????] ---
    ' ?? ????? ???? 'rs' ??? ????? ????? ????? mBranchID
    ' Dim rs As New ADODB.Recordset
    ' s = "Select * from " & mTable & " where " & mmIDField & " = " & mmID
    ' rs.Open s, Cn, adOpenKeyset, adLockOptimistic
    ' If rs.EOF Then Exit Sub
    '-------------------------------------
    
    Dim mQrData As String
    Dim cOptions As New ClsCompanyInfo
    Set cOptions = New ClsCompanyInfo
    
    Dim txtMessage  As String
    Dim STRQRcode As String
    Dim SellerName As String
    Dim Vatregestriationnumber As String
    Dim TimeStamp As String
    Dim invoiceTotalwithVat As String
    Dim Vat As String
    Dim seperator As String
    Dim mPath As String
    Dim s As String ' (???????? ??? rsComp ???)

    '--- [?? ????? ??? ?????] ---
    ' (????? ????? ?? rsDummyBr - ???? ???? ???? ?? ?????? ????? ?????? ??? ??? ?????)
    
    '--- [?????] ??? ?????? ????? ???????? ????????? ?????? ????? ---
    Dim rsComp As New ADODB.Recordset
    Dim oCmdComp As New ADODB.Command
    
    With oCmdComp
        .ActiveConnection = Cn
        .CommandType = adCmdStoredProc
        .CommandText = "sp_GetBranchQRInfo"
        .Parameters.Append .CreateParameter("@BranchID", adInteger, adParamInput, , mBranchID)
    End With
    
    rsComp.Open oCmdComp, , adOpenKeyset, adLockReadOnly

    '--- [?????] ???????? ??? mBranchID ---
    If UCase(mTable) = "TRANSACTIONS" Then
        cOptions.SetBranch = mBranchID
    ElseIf UCase(mTable) = "NOTES" Or UCase(mTable) = "NOTES_ALL" Or UCase(mTable) = "PROJECT_BILLL" Then
        cOptions.SetBranch = mBranchID
    End If

    If Not rsComp.EOF Then
        SellerName = Trim(rsComp!activityName & "")
        Vatregestriationnumber = IIf(Trim(rsComp!VATRegNo & "") = "", "123456789", Trim(rsComp!VATRegNo & ""))
    Else
        ' (??? ???????? ?? ???? ??? ???? ?????)
        SellerName = cOptions.EngCompanyName
        Vatregestriationnumber = IIf(cOptions.VATRegNo = "", "123456789", cOptions.VATRegNo)
    End If
    rsComp.Close
    Set rsComp = Nothing
    Set oCmdComp = Nothing
    
    '--- (???? ??? ????? ??? QR - ?? ????? ?? ???????) ---
    TimeStamp = Format(Transaction_Date, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ssZ")
    
    Dim chkTaxExempt As Boolean
    ' [??????]: ??? ????? ??? ????? ??? 'rs' ???? ???????
    ' ??? ??? ????? 'chkTaxExempt' ??? ?? ????? ????? ?? 'print_report'
    ' If UCase(mTable) = "TRANSACTIONS" Then
    '     If IsNull(rs!chkTaxExempt) Then
    '         chkTaxExempt = False
    '     Else
    '         chkTaxExempt = (rs!chkTaxExempt & "")
    '     End If
    ' End If
    
    ' (?????? chkTaxExempt = False ??????? ???? ??? ???????)
    chkTaxExempt = False
    
    If SystemOptions.PriceWithVAT And UCase(mTable) = "TRANSACTIONS" And Not chkTaxExempt Then
        Vat = Round(val(mmTotalNet) / 1.15 * 0.15, 2)
        cOptions.SetBranch = mBranchID
    Else
        Vat = mmVat
        invoiceTotalwithVat = mmTotalNet
    End If
    
    seperatbor = ""
    txtMessage = createTLVall(SellerName, Vatregestriationnumber, TimeStamp, invoiceTotalwithVat, Vat)
    txtMessage = HexToString(txtMessage)
    txtMessage = Encoden(txtMessage)
    Dim StrText As String ' (?? ??????? ???)
    StrText = txtMessage
    
    ' (?? ??? ????? ?????? ?????? txtMessage)
    
    StrText = Replace(StrText, vbCrLf, vbNullString)
    StrFileName = "Inv" & NoteSerial1 & ".gif"
    StrFileName = RemoveSlashes(StrFileName)
    
    folderPath = App.path
    If right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    folderPath = folderPath & "QRCodeFiles\"

    ' ???? ?????? ?? ??? ?????
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    FileName = folderPath & StrFileName
    n = qrcodeCreateImageInUtf8(FileName, StrText, Options = QRCODE_ESCAPED)

nob:
    If Dir(FileName) <> "" Then
        mFileName = FileName
        Picture1.Picture = LoadPicture(mFileName)
    End If

    '--- [????? + ?????] ---
    ' ???? ????? ?????? ?????? ?????? ?????????
    ' (??? ????? ?????? ??? 'rs' ???? ???? ???? ??? ?????)
    Dim rsUpdate As New ADODB.Recordset
    ' (??? ????? ??? ??? mmID ?? Long ???? ??)
    s = "Select QrCodeData, QrCodeDataPath, QrCodeImage from " & mTable & " where " & mmIDField & " = " & mmID
    rsUpdate.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    If Not rsUpdate.EOF Then
        ' [????? Bug]
        ' mQrData ???? ????? ??????. ????? ??? ??? ???? ??? StrText
        rsUpdate!QrCodeData = StrText  ' (???? mQrData ??? ?????)
        rsUpdate!QrCodeDataPath = CStr(mFileName) ' (???? mPath & mFileName)
        
        If Picture1.Picture <> 0 Then
            SavePictureToDB Picture1, rsUpdate, "QrCodeImage"
        End If
        
        rsUpdate.update
    End If
    rsUpdate.Close
    Set rsUpdate = Nothing
    
End Sub
Public Function GetVal(Optional OrderID As String = 0, Optional NoteID As Double = 0, Optional NoteType As Integer) As Double
    Dim Rs7 As ADODB.Recordset
    Set Rs7 = New ADODB.Recordset
    Dim sql As String
    If OrderID = " " Then GetVal = 0: Exit Function
        sql = " SELECT     SUM(Note_Value) AS Sm"
        sql = sql & " From dbo.Notes"
        sql = sql & " Where  (OrderIDD =" & val(OrderID) & ")"
        If NoteType = 350 Or NoteType = 3 Then ' ’›Ì⁄=Â «·⁄Âœ…
            sql = sql & " and    NoteID  not in (  " & " select NoteID  from notes where notes_all=" & NoteID & "" & ") "
        Else
            sql = sql & " and   NoteID<>" & NoteID & ""
        End If
    
    
    
    
        Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs7.RecordCount > 0 Then
        GetVal = IIf(IsNull(Rs7("Sm").value), 0, Rs7("Sm").value)
    Else
        GetVal = 0
    End If
    

End Function
 
  Public Function GetItemsTotalExpensessByStore(Optional Transaction_ID As Long, Optional StoreID As Integer) As Double
    Dim DblTemp As Double
    Dim RowNum As Long
    Dim Msg  As String
    Dim sql As String
    Dim linetotl As Double
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
     On Local Error GoTo ErrTrap
     
     sql = " SELECT     SUM(LineExpenses*Quantity) AS Price"
     sql = sql & "      From dbo.Transaction_Details"
     sql = sql & "  WHERE     (StoreID2 = " & StoreID & ") AND (Transaction_ID = " & Transaction_ID & ")"
     rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
     If rs2.RecordCount > 0 Then
            linetotl = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
                     If SystemOptions.PursgaseWithoutDecimal = True And GridTrans = PurchaseTransaction Then
                         DblTemp = DblTemp + Int(linetotl)
                     Else
                           DblTemp = DblTemp + Round(linetotl, SystemOptions.SysDefCurrencyForamt)
                    End If
    Else
    DblTemp = 0
    End If

    GetItemsTotalExpensessByStore = DblTemp
    Exit Function
ErrTrap:
    Msg = "?Error "
    GetItemsTotalExpensessByStore = DblTemp
End Function

 Public Function GetAccountByBarnchUser() As String
Dim My_SQL As String
My_SQL = ""
   If SystemOptions.ViewAccountsbyBranch = True Then
            My_SQL = " and (ACCOUNTS.Account_Code in (SELECT    TblAccountBranch.Account_Code"
           My_SQL = My_SQL & " From dbo.TblAccountBranch"
           My_SQL = My_SQL & " WHERE    TblAccountBranch.BranchID  in(" & Current_branchSql & ") "
           
           My_SQL = My_SQL & " and ( ACCOUNTS.Account_Code in (SELECT     TblAccountUser.Account_Code"
           
           My_SQL = My_SQL & " From dbo.TblAccountUser where TblAccountUser.UserID=" & user_id & ")"
           My_SQL = My_SQL & " or  ACCOUNTS.Account_Code not in (SELECT     TblAccountUser.Account_Code"
           My_SQL = My_SQL & " From dbo.TblAccountUser)))"
           
           My_SQL = My_SQL & " or ( ACCOUNTS.Account_Code not in (SELECT     TblAccountBranch.Account_Code"
           My_SQL = My_SQL & " From dbo.TblAccountBranch)"
           My_SQL = My_SQL & " and ACCOUNTS.Account_Code  not in(SELECT     TblAccountUser.Account_Code From dbo.TblAccountUser)  ))"
       
         End If
      GetAccountByBarnchUser = My_SQL
End Function
Public Function GetCustomerIDFromCode(Optional EmpCode As String, _
                                      Optional ByRef Emp_id As Integer, _
                                      Optional Emp_id1 As Integer = 0, _
                                      Optional ByRef EmpCode1 As String, _
                                      Optional ByRef Name1 As String, _
                                      Optional ByRef Name As String, _
                                      Optional ByRef Mobile As String, _
                                       Optional ByRef phone As String, _
                                        Optional ByRef boxmail As String, _
                                       Optional ByRef fax As String, _
                                       Optional ByRef mail As String, _
                                       Optional ByRef adress As String, _
                                        Optional ByRef ZipCode As String, _
                                       Optional ByRef DigCus As String, Optional ByRef jobname As String, _
              Optional ByRef entry As String, Optional ByRef map As String, Optional ByRef ResponsibleContact As String)
            'Dim name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

   'If Emp_id1 <> 0 Then
   '     sql = "select * from TblCustemers where code= " & Emp_id1
   ' Else
 If Name1 <> "" Then
 sql = "select * from TblCustemers where  CusName like '%" & Name1 & "%'"
 Else
        sql = "select * from TblCustemers where  Fullcode  ='" & EmpCode & "'"
        End If
    'End If
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'Dim name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus As String
    If rs.RecordCount > 0 Then
      map = IIf(IsNull(rs("Map").value), "", rs("Map").value)
        jobname = IIf(IsNull(rs("JobName").value), "", rs("JobName").value)
         ResponsibleContact = IIf(IsNull(rs("ResponsibleContact").value), "", rs("ResponsibleContact").value)
    entry = IIf(IsNull(rs("Entry").value), "", rs("Entry").value)
        Emp_id = val(IIf(IsNull(rs("CusID").value), 0, rs("CusID").value))
        EmpCode1 = IIf(IsNull(rs("Fullcode").value), 0, rs("Fullcode").value)
 Name = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
 Mobile = IIf(IsNull(rs("Cus_mobile").value), "", rs("Cus_mobile").value)
 phone = IIf(IsNull(rs("Cus_Phone").value), "", rs("Cus_Phone").value)
 boxmail = IIf(IsNull(rs("BoxMil").value), "", rs("BoxMil").value)
 fax = IIf(IsNull(rs("FaxNumber").value), "", rs("FaxNumber").value)
 mail = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
 adress = IIf(IsNull(rs("Address").value), "", rs("Address").value)
 ZipCode = IIf(IsNull(rs("ZipCode").value), "", rs("ZipCode").value)
 DigCus = IIf(IsNull(rs("TypeCustomer").value), "", rs("TypeCustomer").value)
    Else
        EmpCode1 = 0
    End If

    rs.Close

End Function


Public Function GetCustomerIDFromCode08102018(Optional EmpCode As String, _
                                      Optional ByRef Emp_id As Integer, _
                                      Optional Emp_id1 As Integer = 0, _
                                      Optional ByRef EmpCode1 As String, _
                                      Optional ByRef Name1 As String, _
                                      Optional ByRef Name As String, _
                                      Optional ByRef Mobile As String, _
                                       Optional ByRef phone As String, _
                                        Optional ByRef boxmail As String, _
                                       Optional ByRef fax As String, _
                                       Optional ByRef mail As String, _
                                       Optional ByRef adress As String, _
                                        Optional ByRef ZipCode As String, _
                                       Optional ByRef DigCus As String, Optional ByRef jobname As String, _
              Optional ByRef entry As String, Optional ByRef map As String, Optional ByRef ResponsibleContact As String)
            'Dim name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

   'If Emp_id1 <> 0 Then
   '     sql = "select * from TblCustemers where code= " & Emp_id1
   ' Else
 If Name1 <> "" Then
 sql = "select * from TblCustemers where  CusName like '%" & Name1 & "%'"
 Else
        sql = "select * from TblCustemers where  Fullcode like '%" & EmpCode & "%'"
        End If
    'End If
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'Dim name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus As String
    If rs.RecordCount > 0 Then
      map = IIf(IsNull(rs("Map").value), "", rs("Map").value)
        jobname = IIf(IsNull(rs("JobName").value), "", rs("JobName").value)
         ResponsibleContact = IIf(IsNull(rs("ResponsibleContact").value), "", rs("ResponsibleContact").value)
    entry = IIf(IsNull(rs("Entry").value), "", rs("Entry").value)
        Emp_id = val(IIf(IsNull(rs("CusID").value), 0, rs("CusID").value))
        EmpCode1 = IIf(IsNull(rs("Fullcode").value), 0, rs("Fullcode").value)
 Name = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
 Mobile = IIf(IsNull(rs("Cus_mobile").value), "", rs("Cus_mobile").value)
 phone = IIf(IsNull(rs("Cus_Phone").value), "", rs("Cus_Phone").value)
 boxmail = IIf(IsNull(rs("BoxMil").value), "", rs("BoxMil").value)
 fax = IIf(IsNull(rs("FaxNumber").value), "", rs("FaxNumber").value)
 mail = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
 adress = IIf(IsNull(rs("Address").value), "", rs("Address").value)
 ZipCode = IIf(IsNull(rs("ZipCode").value), "", rs("ZipCode").value)
 DigCus = IIf(IsNull(rs("TypeCustomer").value), "", rs("TypeCustomer").value)
    Else
        EmpCode1 = 0
    End If

    rs.Close

End Function




Public Sub SaveQRCode(mTable As String, mmIDField As String, mmID As Long, ByVal NoteSerial1 As String, ByVal Transaction_Date As Date, mAmount1 As String, _
ByRef Picture1 As PictureBox, Optional ByVal mAmountDisc As String = "", Optional ByVal mVat As String = "", Optional ByVal mTotalNet As String = "")
 
 Dim mmAmountDisc   As Double
 If val(mAmountDisc) > 0 Then
        mmAmountDisc = Round(mAmountDisc, 3)
 End If
Dim mmVat As Double
If val(mVat) > 0 Then
        mmVat = Round(mVat, 3)
 End If
 
Dim mmTotalNet As Double
If val(mTotalNet) > 0 Then
        mmTotalNet = Round(mTotalNet, 3)
 End If
  
 Dim folderPath As String
Dim FileName As String



Dim mmAmount1 As Double
If val(mAmount1) > 0 Then
        mmAmount1 = Round(mAmount1, 3)
 End If
   
    If Not SystemOptions.IsQrCodePrint Then Exit Sub
    
    
    
    
    
    Dim rs As New ADODB.Recordset
    s = "Select * from " & mTable & " where " & mmIDField & " = " & mmID
    rs.Open s, Cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then Exit Sub
    Dim mQrData As String
    Dim cOptions As New ClsCompanyInfo
    Set cOptions = New ClsCompanyInfo
    
    Dim txtMessage  As String
    Dim STRQRcode As String
    Dim SellerName As String
    Dim Vatregestriationnumber As String
    Dim TimeStamp As String
    Dim invoiceTotalwithVat As String
    Dim Vat As String
    Dim seperator As String
        Dim mPath As String
        
    Dim rsDummyBr As New ADODB.Recordset
'    s = "Select IsNull(a790,'100') as a790,IsNull(a791,'') as a791  from branches where IsNull(a790,'') <> '' "
'    rsDummyBr.Open s, Cn, adOpenKeyset, adLockOptimistic
'    If Not rsDummyBr.EOF Then
'                                If val(rsDummyBr!a791 & "") >= val(rsDummyBr!a790 & "") Then
'                                  ' MsgBox("Your QR is not working ")
'                                  '  Exit Sub
'                                End If
'    Else
'        rsDummyBr.Close
'        s = "Select a790,a791 from branches"
'        rsDummyBr.Open s, Cn, adOpenKeyset, adLockOptimistic
'        rsDummyBr!a790 = "500"
'        rsDummyBr.update
'    End If
    
    
    
    Dim mCount As Long
'    mCount = val(rsDummyBr!a791 & "") + 1
'    s = "Update branches set  a791 = '" & mCount & "' where IsNull(a790,'') <> '' "
'    Cn.Execute s

s = " SELECT"
s = s & "       BB.branch_namee,"
s = s & "       ActivityName ="
s = s & "           CASE WHEN NULLIF(LTRIM(RTRIM(BB.branch_namee)), '') IS NOT NULL AND NULLIF(LTRIM(RTRIM(BB.VATRegNo)), '') IS NOT NULL "
s = s & "               THEN BB.branch_namee "
s = s & "               ELSE TA.namee END,"
s = s & "       VATRegNo ="
s = s & "           CASE WHEN NULLIF(LTRIM(RTRIM(TA.VATRegNo)), '') IS NOT NULL"
s = s & "               THEN TA.VATRegNo ELSE BB.VATRegNo END "
s = s & "FROM TblBranchesData AS BB "
s = s & "LEFT JOIN tblActivitesType AS TA "
s = s & " ON TA.id = BB.ActivityTypeId"

    If UCase(mTable) = "TRANSACTIONS" Then
           
            cOptions.SetBranch = val(rs!BranchID & "")
        s = s & " Where BB.branch_id = " & val(rs!BranchID & "")
    ElseIf UCase(mTable) = "NOTES" Or UCase(mTable) = "NOTES_ALL" Or UCase(mTable) = "PROJECT_BILLL" Then
            s = s & " Where BB.branch_id = " & val(rs!branch_no & "")
    
           
            cOptions.SetBranch = val(rs!branch_no & "")
        
    End If
    Dim rsComp As New ADODB.Recordset
    rsComp.Open s, Cn, adOpenKeyset, adLockReadOnly
    If Not rsComp.EOF Then
         SellerName = Trim(rsComp!activityName & "")
        'Vatregestriationnumber = IIf(cOptions.VATRegNo = "", "123456789", cOptions.VATRegNo)
        Vatregestriationnumber = IIf(Trim(rsComp!VATRegNo & "") = "", IIf(cOptions.VATRegNo = "", "123456789", cOptions.VATRegNo), Trim(rsComp!VATRegNo & ""))
    Else
        SellerName = cOptions.EngCompanyName
        Vatregestriationnumber = IIf(cOptions.VATRegNo = "", "123456789", cOptions.VATRegNo)
    End If
    
   
     TimeStamp = Format(Transaction_Date, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ssZ")
    
    
    
        
        

        
        
8

       Dim chkTaxExempt As Boolean
        If UCase(mTable) = "TRANSACTIONS" Then
             If IsNull(rs!chkTaxExempt) Then
                chkTaxExempt = False
            Else
                chkTaxExempt = (rs!chkTaxExempt & "")
            End If
        End If
        
        
        If SystemOptions.PriceWithVAT And UCase(mTable) = "TRANSACTIONS" And Not chkTaxExempt Then
            Vat = Round(val(mmTotalNet) / 1.15 * 0.15, 2)
            cOptions.SetBranch = val(rs!BranchID & "")
          '  invoiceTotalwithVat = val(mmTotalNet / 1.15)
        Else
            Vat = mmVat
            invoiceTotalwithVat = mmTotalNet
        End If
        seperatbor = ""
        txtMessage = createTLVall(SellerName, Vatregestriationnumber, TimeStamp, invoiceTotalwithVat, Vat)
        'HexToString
        
        txtMessage = HexToString(txtMessage)
        txtMessage = Encoden(txtMessage)
        StrText = txtMessage
        StrFileName = "Inv.gif"
        n = qrcodeCreateImageInUtf8(StrFileName, StrText, Options = QRCODE_ESCAPED)
   ' n = qrcodeCreateImageInUtf8(filename + StrFileName, StrText, Options = QRCODE_ESCAPED)
             
             
             If SystemOptions.PriceWithVAT And UCase(mTable) = "TRANSACTIONS" And Not chkTaxExempt Then
            Vat = Round(val(mmTotalNet) / 1.15 * 0.15, 2)
         '   invoiceTotalwithVat = Round(val(mmTotalNet / 1.15), 3)
        Else
            Vat = mmVat
            
        End If
        invoiceTotalwithVat = mmTotalNet
        seperatbor = ""
        txtMessage = ""
        txtMessage = createTLVall(SellerName, Vatregestriationnumber, TimeStamp, invoiceTotalwithVat, Vat)
        'HexToString

        txtMessage = HexToString(txtMessage)
        txtMessage = Encoden(txtMessage)
        StrText = txtMessage
        StrText = Replace(StrText, vbCrLf, vbNullString)
        StrFileName = "Inv" & NoteSerial1 & ".gif"
        'StrFileName = "Inv.gif"
        StrFileName = RemoveSlashes(StrFileName)
        folderPath = App.path
If right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
folderPath = folderPath & "QRCodeFiles\"

' √‰‘∆ «·„Ã·œ ·Ê €Ì— „ÊÃÊœ
If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

FileName = folderPath & StrFileName
n = qrcodeCreateImageInUtf8(FileName, StrText, Options = QRCODE_ESCAPED)

        'n = qrcodeCreateImageInUtf8(StrFileName, StrText, Options = QRCODE_ESCAPED)

nob:
' StrFileName = "Inv" & NoteSerial1 & ".gif"
  '  mFileName = StrFileName
'    mPath = App.path
'    If left(mPath, Len(Trim(mPath))) = "\" Then
'        mPath = left(mPath, Len(Trim(mPath)) - 1)
'    End If
'    If right(mPath, 1) <> "\" Then
'            mPath = mPath & "\"
'    End If
'    If Dir(mPath & mFileName) <> "" Then
'        mFileName = CStr(CStr(mPath) & mFileName)
'
'        Picture1.Picture = LoadPicture(mFileName)
'    End If
    If Dir(FileName) <> "" Then
        mFileName = FileName
        
        Picture1.Picture = LoadPicture(mFileName)
    End If
    

    rs!QrCodeData = mQrData
    rs!QrCodeDataPath = CStr(CStr(mPath) & mFileName)
'    If Picture1.Picture = 0 And CStr(CStr(mPath) & mFileName) <> "" Then
'     Picture1.Picture = LoadPicture(CStr(CStr(mPath) & mFileName))
'     End If
    If Picture1.Picture <> 0 Then
        SavePictureToDB Picture1, rs, "QrCodeImage"
        
    End If
    
    rs.update
        
        

End Sub

Public Function RemoveSlashes(ByVal Txt As String) As String
    RemoveSlashes = Replace(Txt, "/", "")
End Function



Public Sub SaveQRCode2(mTable As String, mmIDField As String, mmID As Long, ByVal NoteSerial1 As String, ByVal Transaction_Date As Date, mAmount1 As String, _
ByRef Picture1 As PictureBox, Optional ByVal mAmountDisc As String = "", Optional ByVal mVat As String = "", Optional ByVal mTotalNet As String = "", Optional QRCode As String, Optional docType As String, Optional fromscreen As Integer = 0)
  
    
    Dim mFileName As String
    Dim rs As New ADODB.Recordset
    s = "Select * from " & mTable & " where " & mmIDField & " = " & mmID
     If mTable <> "transactions" Then
     s = s & " and  isnull(isdeleted,0)=0"
   End If
    
    rs.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF Then Exit Sub
    Dim mQrData As String
    Dim cOptions As New ClsCompanyInfo
    Set cOptions = New ClsCompanyInfo
    Dim txtMessage  As String
    Dim STRQRcode As String
    Dim SellerName As String
    Dim Vatregestriationnumber As String
    Dim TimeStamp As String
    Dim invoiceTotalwithVat As String
    Dim Vat As String
    Dim seperator As String
        Dim mPath As String
      ' „‰ Õ—ﬂ«  «·»Ì⁄ «·„Œ“‰Ì… «·„—œÊœ« '
       If docType = "38801" Then
        mFileName = "SalesInv-" & docType & "- " & NoteSerial1 & ".gif"
      ElseIf docType = "38101" Then
       mFileName = "CreditNoteNote-" & docType & "- " & NoteSerial1 & ".gif"
    ElseIf docType = "38301" Then
       mFileName = "DebitNote-" & docType & "- " & NoteSerial1 & ".gif"
      End If
      
    mFileName = "Invoices\IMAGES\" & mFileName
 
        If SystemOptions.SAveInhomePath = True Then
                              mFileName = "C:" & Environ("HOMEPATH") & "\" & mFileName
                    Else
                                       mPath = App.path
                                 If mId(mPath, Len(Trim(mPath)), 1) = "\" Then
                                     mPath = left(mPath, Len(Trim(mPath)) - 1)
                                 End If
                               '  mPath = mPath & "\"
                      
                     mFileName = mPath & "\" & mFileName
                    End If
                    
                 n = qrcodeCreateImageInUtf8(mFileName, QRCode, Options = QRCODE_ESCAPED)
      
                  
          
   
 
  '0   1 sales 2 retur
    rs!QrCodeData = QRCode
    rs!QrCodeDataPath = mFileName
        If fromscreen = 0 Then
       FrmAnalysItems.Picture1.Picture = LoadPicture(mFileName)
     
    ' SavePictureToDB FrmAnalysItems.Picture1, rs, "QrCodeImage", mFileName
     
                             If SystemOptions.SAveInhomePath = True Then
                    SavePictureToDB FrmAnalysItems.Picture1, rs, "QrCodeImage", "C:" & Environ("HOMEPATH") & "\"
                    Else
                         SavePictureToDB FrmAnalysItems.Picture1, rs, "QrCodeImage", mFileName
                    End If
                    
                    
 End If
 
 
      rs.update
      

'     If fromscreen = 1 Then
'       frmsalebill.Picture1.Picture = LoadPicture(mFileName)
'     FrmAnalysItems
'     SavePictureToDB Picture1, rs, "QrCodeImage", mFileName
' End If
'
'      If fromscreen = 2 Then
'       FrmReturnSalling.Picture1.Picture = LoadPicture(mFileName)
'
'     SavePictureToDB Picture1, rs, "QrCodeImage", mFileName
' End If
 
    

        
        

End Sub




Public Sub SaveQRCode2old(mTable As String, mmIDField As String, mmID As Long, ByVal NoteSerial1 As String, ByVal Transaction_Date As Date, mAmount1 As String, _
ByRef Picture1 As PictureBox, Optional ByVal mAmountDisc As String = "", Optional ByVal mVat As String = "", Optional ByVal mTotalNet As String = "")
 
 Dim mmAmountDisc   As Double
 If val(mAmountDisc) > 0 Then
        mmAmountDisc = Round(mAmountDisc, 3)
 End If
Dim mmVat As Double
If val(mVat) > 0 Then
        mmVat = Round(mVat, 3)
 End If
 
Dim mmTotalNet As Double
If val(mTotalNet) > 0 Then
        mmTotalNet = Round(mTotalNet, 3)
 End If
   
 
Dim mmAmount1 As Double
If val(mAmount1) > 0 Then
        mmAmount1 = Round(mAmount1, 3)
 End If
   
    If Not SystemOptions.IsQrCodePrint Then Exit Sub
    
    
    
    
    
    Dim rs As New ADODB.Recordset
    s = "Select * from " & mTable & " where " & mmIDField & " = " & mmID
    rs.Open s, Cn, adOpenKeyset, adLockOptimistic
    Dim mQrData As String
    Dim cOptions As New ClsCompanyInfo
    Set cOptions = New ClsCompanyInfo
    Dim txtMessage  As String
    Dim STRQRcode As String
    Dim SellerName As String
    Dim Vatregestriationnumber As String
    Dim TimeStamp As String
    Dim invoiceTotalwithVat As String
    Dim Vat As String
    Dim seperator As String
        Dim mPath As String
        
    Dim rsDummyBr As New ADODB.Recordset
    s = "Select IsNull(a790,'100') as a790,IsNull(a791,'') as a791  from branches where IsNull(a790,'') <> '' "
    rsDummyBr.Open s, Cn, adOpenKeyset, adLockOptimistic
    If Not rsDummyBr.EOF Then
                                If val(rsDummyBr!a791 & "") >= val(rsDummyBr!a790 & "") Then
                                    MsgBox ("Your QR is not working ")
                                    Exit Sub
                                End If
    Else
        rsDummyBr.Close
        s = "Select a790,a791 from branches"
        rsDummyBr.Open s, Cn, adOpenKeyset, adLockOptimistic
        rsDummyBr!a790 = "500"
        rsDummyBr.update
    End If
    
    
    
    Dim mCount As Long
    mCount = val(rsDummyBr!a791 & "") + 1
    s = "Update branches set  a791 = '" & mCount & "' where IsNull(a790,'') <> '' "
    Cn.Execute s
    SellerName = cOptions.EngCompanyName
    Vatregestriationnumber = IIf(cOptions.VATRegNo = "", "123456789", cOptions.VATRegNo)
     TimeStamp = Format(Transaction_Date, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ssZ")
    
    
    GoTo 8
    If Not SystemOptions.IsBlue Then
    
       
        mQrData = mQrData & " «”„ «·„Ê—œ:"
        mQrData = mQrData & cOptions.ArabCompanyName & vbNewLine
        mQrData = mQrData & "«·—ﬁ„ «·÷—Ì»Ì:"
        If SystemOptions.VATNoAccordActivity = False Then
            mQrData = mQrData & cOptions.VATRegNo & vbNewLine
        Else
            mQrData = mQrData & GetRegVATNo(CInt(my_branch)) & vbNewLine
        End If
        
        
        mQrData = mQrData & "—ﬁ„ «·›« Ê—…:"
        mQrData = mQrData & NoteSerial1 & vbNewLine
        
        mQrData = mQrData & "«· «—ÌŒ:"
        mQrData = mQrData & Transaction_Date & vbNewLine
        mQrData = mQrData & "«·Êﬁ :"
        mQrData = mQrData & Time & vbNewLine
        mQrData = mQrData & "«·«Ã„«·Ì ﬁ»· «·÷—Ì»…:"
        mQrData = mQrData & mmAmount1 & vbNewLine
        mQrData = mQrData & "«·Œ’Ê„« :"
        mQrData = mQrData & mmAmountDisc & vbNewLine
        mQrData = mQrData & "«·÷—Ì»…:"
        mQrData = mQrData & mmVat & vbNewLine
        mQrData = mQrData & "«·’«›Ì:"
        mQrData = mQrData & mmTotalNet & vbNewLine
        Dim mFileName As String
        mFileName = "Inv" & Trim(NoteSerial1) & ".gif"
        '  n = qrcodeCreateImageInUtf8("china11.gif", ss, Options = QRCODE_ESCAPED)
        Dim n
        '  n = qrcodeCreateImageInUtf8("", mQrData, QRCODE_ESCAPED)
        '  mFileName = "china11.gif"
        StrText = mQrData
        n = qrcodeCreateImageInUtf8(mFileName, mQrData, QRCODE_ESCAPED)
        
GoTo nob:
        
        
    Else
        
        
8
        
        Dim chkTaxExempt As Boolean
        If UCase(mTable) = "TRANSACTIONS" Then
             If IsNull(rs!chkTaxExempt) Then
                chkTaxExempt = False
            Else
                chkTaxExempt = (rs!chkTaxExempt & "")
            End If
        End If
              
        If (SystemOptions.PriceWithVAT And UCase(mTable) = "TRANSACTIONS" And Not chkTaxExempt) Then
            Vat = Round(val(mmTotalNet) / 1.15 * 0.15, 2)
          '  invoiceTotalwithVat = val(mmTotalNet / 1.15)
        Else
            Vat = mmVat
            invoiceTotalwithVat = mmTotalNet
        End If
        If UCase(mTable) = "TBLCARBILLMENTAINS" Then
        End If
        seperatbor = ""
        txtMessage = createTLVall(SellerName, Vatregestriationnumber, TimeStamp, invoiceTotalwithVat, Vat)
        'HexToString
        
        txtMessage = HexToString(txtMessage)
        txtMessage = Encoden(txtMessage)
        StrText = txtMessage
        StrText = "fdsfd"
        StrFileName = "Inv.gif"
        StrFileName = "Inv" & NoteSerial1 & ".gif"
        n = qrcodeCreateImageInUtf8(StrFileName, StrText, Options = QRCODE_ESCAPED)
    End If

             
        If (SystemOptions.PriceWithVAT And UCase(mTable) = "TRANSACTIONS" And Not chkTaxExempt) Then
            Vat = Round(val(mmTotalNet) / 1.15 * 0.15, 2)
         '   invoiceTotalwithVat = Round(val(mmTotalNet / 1.15), 3)
        Else
            Vat = mmVat
            
        End If
        invoiceTotalwithVat = mmTotalNet
        seperatbor = ""
        txtMessage = createTLVall(SellerName, Vatregestriationnumber, TimeStamp, invoiceTotalwithVat, Vat)
        'HexToString
        
        txtMessage = HexToString(txtMessage)
        txtMessage = Encoden(txtMessage)
        StrText = txtMessage
        StrFileName = "Inv" & NoteSerial1 & ".gif"
        'StrFileName = "Inv.gif"
        n = qrcodeCreateImageInUtf8(StrFileName, StrText, Options = QRCODE_ESCAPED)

nob:
 
    mFileName = StrFileName
    mPath = App.path
    If left(mPath, Len(Trim(mPath))) = "\" Then
        mPath = left(mPath, Len(Trim(mPath)) - 1)
    End If
  '  mPath = mPath & "\"
    If Dir(mPath & mFileName) <> "" Then
        mFileName = CStr(CStr(mPath) & mFileName)
        
        Picture1.Picture = LoadPicture(mFileName)
    End If

    rs!QrCodeData = mQrData
    rs!QrCodeDataPath = CStr(mFileName)
    If Picture1.Picture <> 0 Then
        SavePictureToDB Picture1, rs, "QrCodeImage"
    End If
    
    rs.update
        
        

End Sub




Function Encoden(strMessage As String) As String
  Set Bas64 = New ClsSupplierPrice
    If flgFile = 0 Then
        strBuffer = strMessage
    End If
    'Debug.Print "Length of message = " & CStr(Len(strBuffer))
    If flgFile > 0 Then
        Bas64.bBuffer = byteBuffer
    Else
        Bas64.sBuffer = strBuffer
    End If
    Call Bas64.Base64Encode
    encBuffer = Bas64.Base64Buf
    If flgFile > 0 Then
        flgFile = flgFile + 1
     Encoden = strMessage & "Length of Encoded file = " & CStr(Len(encBuffer)) & vbCrLf
    Else
      Encoden = encBuffer
        Debug.Print "Length of encoded message = " & CStr(Len(encBuffer))
        Call DebugPrintString("Encoded String", encBuffer)
    End If
End Function
Function createTLVall(SellerName As String, Vatregestriationnumber As String, TimeStamp As String, invoiceTotalwithVat As String, Vat As String) As String
Dim STRQRcode As String
STRQRcode = createTLV("01", (SellerName)) & seperator
STRQRcode = STRQRcode & createTLV("02", Vatregestriationnumber) & seperator
STRQRcode = STRQRcode & createTLV("03", TimeStamp) & seperator
STRQRcode = STRQRcode & createTLV("04", invoiceTotalwithVat) & seperator
STRQRcode = STRQRcode & createTLV("05", Vat) & seperator
createTLVall = Trim(STRQRcode)

End Function


''Dim StrFileName As String
''Dim StrText As String
''Dim Bas64 As ClsSupplierPrice
''
''Dim FileName As String
''Dim byteBuffer() As Byte
''Dim strBuffer As String
''Dim flgFile As Integer
''Dim encBuffer As String
''Dim decBuffer As String
''
''
''Public Function GetVal(Optional OrderID As String = 0, Optional NoteID As Double = 0, Optional NoteType As Integer) As Double
''    Dim Rs7 As ADODB.Recordset
''    Set Rs7 = New ADODB.Recordset
''    Dim sql As String
''    If OrderID = " " Then GetVal = 0: Exit Function
''        sql = " SELECT     SUM(Note_Value) AS Sm"
''        sql = sql & " From dbo.Notes"
''        sql = sql & " Where  (OrderIDD =" & OrderID & ")"
''        If NoteType = 350 Or NoteType = 3 Then ' ’›Ì⁄=Â «·⁄Âœ…
''            sql = sql & " and    NoteID  not in (  " & " select NoteID  from notes where notes_all=" & NoteID & "" & ") "
''        Else
''            sql = sql & " and   NoteID<>" & NoteID & ""
''        End If
''
''
''
''
''        Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
''    If Rs7.RecordCount > 0 Then
''        GetVal = IIf(IsNull(Rs7("Sm").value), 0, Rs7("Sm").value)
''    Else
''        GetVal = 0
''    End If
''
''
''End Function
''
''  Public Function GetItemsTotalExpensessByStore(Optional Transaction_ID As Long, Optional StoreID As Integer) As Double
''    Dim DblTemp As Double
''    Dim RowNum As Long
''    Dim Msg  As String
''    Dim sql As String
''    Dim linetotl As Double
''    Dim rs2 As ADODB.Recordset
''    Set rs2 = New ADODB.Recordset
''     On Local Error GoTo ErrTrap
''
''     sql = " SELECT     SUM(LineExpenses*Quantity) AS Price"
''     sql = sql & "      From dbo.Transaction_Details"
''     sql = sql & "  WHERE     (StoreID2 = " & StoreID & ") AND (Transaction_ID = " & Transaction_ID & ")"
''     rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
''     If rs2.RecordCount > 0 Then
''            linetotl = IIf(IsNull(rs2("Price").value), 0, rs2("Price").value)
''                     If SystemOptions.PursgaseWithoutDecimal = True And GridTrans = PurchaseTransaction Then
''                         DblTemp = DblTemp + Int(linetotl)
''                     Else
''                           DblTemp = DblTemp + Round(linetotl, SystemOptions.SysDefCurrencyForamt)
''                    End If
''    Else
''    DblTemp = 0
''    End If
''
''    GetItemsTotalExpensessByStore = DblTemp
''    Exit Function
''ErrTrap:
''    Msg = "?Error "
''    GetItemsTotalExpensessByStore = DblTemp
''End Function
''
'' Public Function GetAccountByBarnchUser() As String
''Dim My_SQL As String
''My_SQL = ""
''   If SystemOptions.ViewAccountsbyBranch = True Then
''            My_SQL = " and (ACCOUNTS.Account_Code in (SELECT    TblAccountBranch.Account_Code"
''           My_SQL = My_SQL & " From dbo.TblAccountBranch"
''           My_SQL = My_SQL & " WHERE    TblAccountBranch.BranchID  in(" & Current_branchSql & ") "
''
''           My_SQL = My_SQL & " and ( ACCOUNTS.Account_Code in (SELECT     TblAccountUser.Account_Code"
''
''           My_SQL = My_SQL & " From dbo.TblAccountUser where TblAccountUser.UserID=" & user_id & ")"
''           My_SQL = My_SQL & " or  ACCOUNTS.Account_Code not in (SELECT     TblAccountUser.Account_Code"
''           My_SQL = My_SQL & " From dbo.TblAccountUser)))"
''
''           My_SQL = My_SQL & " or ( ACCOUNTS.Account_Code not in (SELECT     TblAccountBranch.Account_Code"
''           My_SQL = My_SQL & " From dbo.TblAccountBranch)"
''           My_SQL = My_SQL & " and ACCOUNTS.Account_Code  not in(SELECT     TblAccountUser.Account_Code From dbo.TblAccountUser)  ))"
''
''         End If
''      GetAccountByBarnchUser = My_SQL
''End Function
''Public Function GetCustomerIDFromCode(Optional EmpCode As String, _
''                                      Optional ByRef Emp_id As Integer, _
''                                      Optional Emp_id1 As Integer = 0, _
''                                      Optional ByRef EmpCode1 As String, _
''                                      Optional ByRef Name1 As String, _
''                                      Optional ByRef Name As String, _
''                                      Optional ByRef Mobile As String, _
''                                       Optional ByRef phone As String, _
''                                        Optional ByRef boxmail As String, _
''                                       Optional ByRef fax As String, _
''                                       Optional ByRef mail As String, _
''                                       Optional ByRef adress As String, _
''                                        Optional ByRef ZipCode As String, _
''                                       Optional ByRef DigCus As String, Optional ByRef jobname As String, _
''              Optional ByRef entry As String, Optional ByRef map As String, Optional ByRef ResponsibleContact As String)
''            'Dim name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus As String
''    Dim sql As String
''    Dim rs As New ADODB.Recordset
''    Dim Balance As Double
''
''   'If Emp_id1 <> 0 Then
''   '     sql = "select * from TblCustemers where code= " & Emp_id1
''   ' Else
'' If Name1 <> "" Then
'' sql = "select * from TblCustemers where  CusName like '%" & Name1 & "%'"
'' Else
''        sql = "select * from TblCustemers where  Fullcode  ='" & EmpCode & "'"
''        End If
''    'End If
''
''    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'''Dim name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus As String
''    If rs.RecordCount > 0 Then
''      map = IIf(IsNull(rs("Map").value), "", rs("Map").value)
''        jobname = IIf(IsNull(rs("JobName").value), "", rs("JobName").value)
''         ResponsibleContact = IIf(IsNull(rs("ResponsibleContact").value), "", rs("ResponsibleContact").value)
''    entry = IIf(IsNull(rs("Entry").value), "", rs("Entry").value)
''        Emp_id = val(IIf(IsNull(rs("CusID").value), 0, rs("CusID").value))
''        EmpCode1 = IIf(IsNull(rs("Fullcode").value), 0, rs("Fullcode").value)
'' Name = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
'' Mobile = IIf(IsNull(rs("Cus_mobile").value), "", rs("Cus_mobile").value)
'' phone = IIf(IsNull(rs("Cus_Phone").value), "", rs("Cus_Phone").value)
'' boxmail = IIf(IsNull(rs("BoxMil").value), "", rs("BoxMil").value)
'' fax = IIf(IsNull(rs("FaxNumber").value), "", rs("FaxNumber").value)
'' mail = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
'' adress = IIf(IsNull(rs("Address").value), "", rs("Address").value)
'' ZipCode = IIf(IsNull(rs("ZipCode").value), "", rs("ZipCode").value)
'' DigCus = IIf(IsNull(rs("TypeCustomer").value), "", rs("TypeCustomer").value)
''    Else
''        EmpCode1 = 0
''    End If
''
''    rs.Close
''
''End Function
''
''
''Public Function GetCustomerIDFromCode08102018(Optional EmpCode As String, _
''                                      Optional ByRef Emp_id As Integer, _
''                                      Optional Emp_id1 As Integer = 0, _
''                                      Optional ByRef EmpCode1 As String, _
''                                      Optional ByRef Name1 As String, _
''                                      Optional ByRef Name As String, _
''                                      Optional ByRef Mobile As String, _
''                                       Optional ByRef phone As String, _
''                                        Optional ByRef boxmail As String, _
''                                       Optional ByRef fax As String, _
''                                       Optional ByRef mail As String, _
''                                       Optional ByRef adress As String, _
''                                        Optional ByRef ZipCode As String, _
''                                       Optional ByRef DigCus As String, Optional ByRef jobname As String, _
''              Optional ByRef entry As String, Optional ByRef map As String, Optional ByRef ResponsibleContact As String)
''            'Dim name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus As String
''    Dim sql As String
''    Dim rs As New ADODB.Recordset
''    Dim Balance As Double
''
''   'If Emp_id1 <> 0 Then
''   '     sql = "select * from TblCustemers where code= " & Emp_id1
''   ' Else
'' If Name1 <> "" Then
'' sql = "select * from TblCustemers where  CusName like '%" & Name1 & "%'"
'' Else
''        sql = "select * from TblCustemers where  Fullcode like '%" & EmpCode & "%'"
''        End If
''    'End If
''
''    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'''Dim name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus As String
''    If rs.RecordCount > 0 Then
''      map = IIf(IsNull(rs("Map").value), "", rs("Map").value)
''        jobname = IIf(IsNull(rs("JobName").value), "", rs("JobName").value)
''         ResponsibleContact = IIf(IsNull(rs("ResponsibleContact").value), "", rs("ResponsibleContact").value)
''    entry = IIf(IsNull(rs("Entry").value), "", rs("Entry").value)
''        Emp_id = val(IIf(IsNull(rs("CusID").value), 0, rs("CusID").value))
''        EmpCode1 = IIf(IsNull(rs("Fullcode").value), 0, rs("Fullcode").value)
'' Name = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
'' Mobile = IIf(IsNull(rs("Cus_mobile").value), "", rs("Cus_mobile").value)
'' phone = IIf(IsNull(rs("Cus_Phone").value), "", rs("Cus_Phone").value)
'' boxmail = IIf(IsNull(rs("BoxMil").value), "", rs("BoxMil").value)
'' fax = IIf(IsNull(rs("FaxNumber").value), "", rs("FaxNumber").value)
'' mail = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
'' adress = IIf(IsNull(rs("Address").value), "", rs("Address").value)
'' ZipCode = IIf(IsNull(rs("ZipCode").value), "", rs("ZipCode").value)
'' DigCus = IIf(IsNull(rs("TypeCustomer").value), "", rs("TypeCustomer").value)
''    Else
''        EmpCode1 = 0
''    End If
''
''    rs.Close
''
''End Function
''
''
''
''
''Public Sub SaveQRCode(mTable As String, mmIDField As String, mmID As Long, ByVal NoteSerial1 As String, ByVal Transaction_Date As Date, mAmount1 As String, _
''ByRef Picture1 As PictureBox, Optional ByVal mAmountDisc As String = "", Optional ByVal mVat As String = "", Optional ByVal mTotalNet As String = "")
''
'' Dim mmAmountDisc   As Double
'' If val(mAmountDisc) > 0 Then
''        mmAmountDisc = Round(mAmountDisc, 3)
'' End If
''Dim mmVat As Double
''If val(mVat) > 0 Then
''        mmVat = Round(mVat, 3)
'' End If
''
''Dim mmTotalNet As Double
''If val(mTotalNet) > 0 Then
''        mmTotalNet = Round(mTotalNet, 3)
'' End If
''
''
''Dim mmAmount1 As Double
''If val(mAmount1) > 0 Then
''        mmAmount1 = Round(mAmount1, 3)
'' End If
''
''    If Not SystemOptions.IsQrCodePrint Then Exit Sub
''
''
''
''
''
''    Dim rs As New ADODB.Recordset
''    s = "Select * from " & mTable & " where " & mmIDField & " = " & mmID
''    rs.Open s, Cn, adOpenKeyset, adLockOptimistic
''    If rs.EOF Then Exit Sub
''    Dim mQrData As String
''    Dim cOptions As New ClsCompanyInfo
''    Set cOptions = New ClsCompanyInfo
''
''    Dim txtMessage  As String
''    Dim STRQRcode As String
''    Dim SellerName As String
''    Dim Vatregestriationnumber As String
''    Dim TimeStamp As String
''    Dim invoiceTotalwithVat As String
''    Dim Vat As String
''    Dim seperator As String
''        Dim mPath As String
''
''    Dim rsDummyBr As New ADODB.Recordset
''    s = "Select IsNull(a790,'100') as a790,IsNull(a791,'') as a791  from branches where IsNull(a790,'') <> '' "
''    rsDummyBr.Open s, Cn, adOpenKeyset, adLockOptimistic
''    If Not rsDummyBr.EOF Then
''                                If val(rsDummyBr!a791 & "") >= val(rsDummyBr!a790 & "") Then
''                                    MsgBox ("Your QR is not working ")
''                                    Exit Sub
''                                End If
''    Else
''        rsDummyBr.Close
''        s = "Select a790,a791 from branches"
''        rsDummyBr.Open s, Cn, adOpenKeyset, adLockOptimistic
''        rsDummyBr!a790 = "500"
''        rsDummyBr.update
''    End If
''
''
''
''    Dim mCount As Long
''    mCount = val(rsDummyBr!a791 & "") + 1
''    s = "Update branches set  a791 = '" & mCount & "' where IsNull(a790,'') <> '' "
''    Cn.Execute s
''    If UCase(mTable) = "TRANSACTIONS" Then
''
''            cOptions.SetBranch = val(rs!BranchID & "")
''    ElseIf UCase(mTable) = "NOTES" Then
''
''
''            cOptions.SetBranch = val(rs!branch_no & "")
''    End If
''    SellerName = cOptions.EngCompanyName
''    Vatregestriationnumber = IIf(cOptions.VATRegNo = "", "123456789", cOptions.VATRegNo)
''     TimeStamp = Format(Transaction_Date, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ssZ")
''
''
''
''
''
''
''
''
''8
''
''       Dim chkTaxExempt As Boolean
''        If UCase(mTable) = "TRANSACTIONS" Then
''             If IsNull(rs!chkTaxExempt) Then
''                chkTaxExempt = False
''            Else
''                chkTaxExempt = (rs!chkTaxExempt & "")
''            End If
''        End If
''
''
''        If SystemOptions.PriceWithVAT And UCase(mTable) = "TRANSACTIONS" And Not chkTaxExempt Then
''            Vat = Round(val(mmTotalNet) / 1.15 * 0.15, 2)
''            cOptions.SetBranch = val(rs!BranchID & "")
''          '  invoiceTotalwithVat = val(mmTotalNet / 1.15)
''        Else
''            Vat = mmVat
''            invoiceTotalwithVat = mmTotalNet
''        End If
''        seperatbor = ""
''        txtMessage = createTLVall(SellerName, Vatregestriationnumber, TimeStamp, invoiceTotalwithVat, Vat)
''        'HexToString
''
''        txtMessage = HexToString(txtMessage)
''        txtMessage = Encoden(txtMessage)
''        StrText = txtMessage
''        StrFileName = "Inv.gif"
''        n = qrcodeCreateImageInUtf8(StrFileName, StrText, Options = QRCODE_ESCAPED)
''
''
''
''             If SystemOptions.PriceWithVAT And UCase(mTable) = "TRANSACTIONS" And Not chkTaxExempt Then
''            Vat = Round(val(mmTotalNet) / 1.15 * 0.15, 2)
''         '   invoiceTotalwithVat = Round(val(mmTotalNet / 1.15), 3)
''        Else
''            Vat = mmVat
''
''        End If
''        invoiceTotalwithVat = mmTotalNet
''        seperatbor = ""
''        txtMessage = createTLVall(SellerName, Vatregestriationnumber, TimeStamp, invoiceTotalwithVat, Vat)
''        'HexToString
''
''        txtMessage = HexToString(txtMessage)
''        txtMessage = Encoden(txtMessage)
''        StrText = txtMessage
''                 StrText = Replace(StrText, vbCrLf, vbNullString)
''        StrFileName = "Inv" & NoteSerial1 & ".gif"
''        'StrFileName = "Inv.gif"
''        n = qrcodeCreateImageInUtf8(StrFileName, StrText, Options = QRCODE_ESCAPED)
''
''nob:
'' StrFileName = "Inv" & NoteSerial1 & ".gif"
''    mFileName = StrFileName
''    mPath = App.path
''    If left(mPath, Len(Trim(mPath))) = "\" Then
''        mPath = left(mPath, Len(Trim(mPath)) - 1)
''    End If
''    If right(mPath, 1) <> "\" Then
''            mPath = mPath & "\"
''    End If
''    If Dir(mPath & mFileName) <> "" Then
''        mFileName = CStr(CStr(mPath) & mFileName)
''
''        Picture1.Picture = LoadPicture(mFileName)
''    End If
''
''    rs!QrCodeData = mQrData
''    rs!QrCodeDataPath = CStr(CStr(mPath) & mFileName)
''    If Picture1.Picture <> 0 Then
''        SavePictureToDB Picture1, rs, "QrCodeImage"
''    End If
''
''    rs.update
''
''
''
''End Sub
''
''
''
''
''
''Public Sub SaveQRCode2(mTable As String, mmIDField As String, mmID As Long, ByVal NoteSerial1 As String, ByVal Transaction_Date As Date, mAmount1 As String, _
''ByRef Picture1 As PictureBox, Optional ByVal mAmountDisc As String = "", Optional ByVal mVat As String = "", Optional ByVal mTotalNet As String = "")
''
'' Dim mmAmountDisc   As Double
'' If val(mAmountDisc) > 0 Then
''        mmAmountDisc = Round(mAmountDisc, 3)
'' End If
''Dim mmVat As Double
''If val(mVat) > 0 Then
''        mmVat = Round(mVat, 3)
'' End If
''
''Dim mmTotalNet As Double
''If val(mTotalNet) > 0 Then
''        mmTotalNet = Round(mTotalNet, 3)
'' End If
''
''
''Dim mmAmount1 As Double
''If val(mAmount1) > 0 Then
''        mmAmount1 = Round(mAmount1, 3)
'' End If
''
''    If Not SystemOptions.IsQrCodePrint Then Exit Sub
''
''
''
''
''
''    Dim rs As New ADODB.Recordset
''    s = "Select * from " & mTable & " where " & mmIDField & " = " & mmID
''    rs.Open s, Cn, adOpenKeyset, adLockOptimistic
''    Dim mQrData As String
''    Dim cOptions As New ClsCompanyInfo
''    Set cOptions = New ClsCompanyInfo
''    Dim txtMessage  As String
''    Dim STRQRcode As String
''    Dim SellerName As String
''    Dim Vatregestriationnumber As String
''    Dim TimeStamp As String
''    Dim invoiceTotalwithVat As String
''    Dim Vat As String
''    Dim seperator As String
''        Dim mPath As String
''
''    Dim rsDummyBr As New ADODB.Recordset
''    s = "Select IsNull(a790,'100') as a790,IsNull(a791,'') as a791  from branches where IsNull(a790,'') <> '' "
''    rsDummyBr.Open s, Cn, adOpenKeyset, adLockOptimistic
''    If Not rsDummyBr.EOF Then
''                                If val(rsDummyBr!a791 & "") >= val(rsDummyBr!a790 & "") Then
''                                    MsgBox ("Your QR is not working ")
''                                    Exit Sub
''                                End If
''    Else
''        rsDummyBr.Close
''        s = "Select a790,a791 from branches"
''        rsDummyBr.Open s, Cn, adOpenKeyset, adLockOptimistic
''        rsDummyBr!a790 = "500"
''        rsDummyBr.update
''    End If
''
''
''
''    Dim mCount As Long
''    mCount = val(rsDummyBr!a791 & "") + 1
''    s = "Update branches set  a791 = '" & mCount & "' where IsNull(a790,'') <> '' "
''    Cn.Execute s
''    SellerName = cOptions.EngCompanyName
''    Vatregestriationnumber = IIf(cOptions.VATRegNo = "", "123456789", cOptions.VATRegNo)
''     TimeStamp = Format(Transaction_Date, "yyyy-mm-dd") & "T" & Format(Time, "hh:mm:ssZ")
''
''
''    GoTo 8
''    If Not SystemOptions.IsBlue Then
''
''
''        mQrData = mQrData & " «”„ «·„Ê—œ:"
''        mQrData = mQrData & cOptions.ArabCompanyName & vbNewLine
''        mQrData = mQrData & "«·—ﬁ„ «·÷—Ì»Ì:"
''        If SystemOptions.VATNoAccordActivity = False Then
''            mQrData = mQrData & cOptions.VATRegNo & vbNewLine
''        Else
''            mQrData = mQrData & GetRegVATNo(CInt(my_branch)) & vbNewLine
''        End If
''
''
''        mQrData = mQrData & "—ﬁ„ «·›« Ê—…:"
''        mQrData = mQrData & NoteSerial1 & vbNewLine
''
''        mQrData = mQrData & "«· «—ÌŒ:"
''        mQrData = mQrData & Transaction_Date & vbNewLine
''        mQrData = mQrData & "«·Êﬁ :"
''        mQrData = mQrData & Time & vbNewLine
''        mQrData = mQrData & "«·«Ã„«·Ì ﬁ»· «·÷—Ì»…:"
''        mQrData = mQrData & mmAmount1 & vbNewLine
''        mQrData = mQrData & "«·Œ’Ê„« :"
''        mQrData = mQrData & mmAmountDisc & vbNewLine
''        mQrData = mQrData & "«·÷—Ì»…:"
''        mQrData = mQrData & mmVat & vbNewLine
''        mQrData = mQrData & "«·’«›Ì:"
''        mQrData = mQrData & mmTotalNet & vbNewLine
''        Dim mFileName As String
''        mFileName = "Inv" & Trim(NoteSerial1) & ".gif"
''        '  n = qrcodeCreateImageInUtf8("china11.gif", ss, Options = QRCODE_ESCAPED)
''        Dim n
''        '  n = qrcodeCreateImageInUtf8("", mQrData, QRCODE_ESCAPED)
''        '  mFileName = "china11.gif"
''        StrText = mQrData
''        n = qrcodeCreateImageInUtf8(mFileName, mQrData, QRCODE_ESCAPED)
''
''GoTo nob:
''
''
''    Else
''
''
''8
''
''        Dim chkTaxExempt As Boolean
''        If UCase(mTable) = "TRANSACTIONS" Then
''             If IsNull(rs!chkTaxExempt) Then
''                chkTaxExempt = False
''            Else
''                chkTaxExempt = (rs!chkTaxExempt & "")
''            End If
''        End If
''
''        If (SystemOptions.PriceWithVAT And UCase(mTable) = "TRANSACTIONS" And Not chkTaxExempt) Then
''            Vat = Round(val(mmTotalNet) / 1.15 * 0.15, 2)
''          '  invoiceTotalwithVat = val(mmTotalNet / 1.15)
''        Else
''            Vat = mmVat
''            invoiceTotalwithVat = mmTotalNet
''        End If
''        If UCase(mTable) = "TBLCARBILLMENTAINS" Then
''        End If
''        seperatbor = ""
''        txtMessage = createTLVall(SellerName, Vatregestriationnumber, TimeStamp, invoiceTotalwithVat, Vat)
''        'HexToString
''
''        txtMessage = HexToString(txtMessage)
''        txtMessage = Encoden(txtMessage)
''        StrText = txtMessage
''        StrText = "fdsfd"
''        StrFileName = "Inv.gif"
''        StrFileName = "Inv" & NoteSerial1 & ".gif"
''        n = qrcodeCreateImageInUtf8(StrFileName, StrText, Options = QRCODE_ESCAPED)
''    End If
''
''
''        If (SystemOptions.PriceWithVAT And UCase(mTable) = "TRANSACTIONS" And Not chkTaxExempt) Then
''            Vat = Round(val(mmTotalNet) / 1.15 * 0.15, 2)
''         '   invoiceTotalwithVat = Round(val(mmTotalNet / 1.15), 3)
''        Else
''            Vat = mmVat
''
''        End If
''        invoiceTotalwithVat = mmTotalNet
''        seperatbor = ""
''        txtMessage = createTLVall(SellerName, Vatregestriationnumber, TimeStamp, invoiceTotalwithVat, Vat)
''        'HexToString
''
''        txtMessage = HexToString(txtMessage)
''        txtMessage = Encoden(txtMessage)
''        StrText = txtMessage
''        StrFileName = "Inv" & NoteSerial1 & ".gif"
''        'StrFileName = "Inv.gif"
''        n = qrcodeCreateImageInUtf8(StrFileName, StrText, Options = QRCODE_ESCAPED)
''
''nob:
''
''    mFileName = StrFileName
''    mPath = App.path
''    If left(mPath, Len(Trim(mPath))) = "\" Then
''        mPath = left(mPath, Len(Trim(mPath)) - 1)
''    End If
''  '  mPath = mPath & "\"
''    If Dir(mPath & mFileName) <> "" Then
''        mFileName = CStr(CStr(mPath) & mFileName)
''
''        Picture1.Picture = LoadPicture(mFileName)
''    End If
''
''    rs!QrCodeData = mQrData
''    rs!QrCodeDataPath = CStr(mFileName)
''    If Picture1.Picture <> 0 Then
''        SavePictureToDB Picture1, rs, "QrCodeImage"
''    End If
''
''    rs.update
''
''
''
''End Sub
''
''
''
''
''Function Encoden(strMessage As String) As String
''  Set Bas64 = New ClsSupplierPrice
''    If flgFile = 0 Then
''        strBuffer = strMessage
''    End If
''    'Debug.Print "Length of message = " & CStr(Len(strBuffer))
''    If flgFile > 0 Then
''        Bas64.bBuffer = byteBuffer
''    Else
''        Bas64.sBuffer = strBuffer
''    End If
''    Call Bas64.Base64Encode
''    encBuffer = Bas64.Base64Buf
''    If flgFile > 0 Then
''        flgFile = flgFile + 1
''     Encoden = strMessage & "Length of Encoded file = " & CStr(Len(encBuffer)) & vbCrLf
''    Else
''      Encoden = encBuffer
''        Debug.Print "Length of encoded message = " & CStr(Len(encBuffer))
''        Call DebugPrintString("Encoded String", encBuffer)
''    End If
''End Function
''Function createTLVall(SellerName As String, Vatregestriationnumber As String, TimeStamp As String, invoiceTotalwithVat As String, Vat As String) As String
''Dim STRQRcode As String
''STRQRcode = createTLV("01", (SellerName)) & seperator
''STRQRcode = STRQRcode & createTLV("02", Vatregestriationnumber) & seperator
''STRQRcode = STRQRcode & createTLV("03", TimeStamp) & seperator
''STRQRcode = STRQRcode & createTLV("04", invoiceTotalwithVat) & seperator
''STRQRcode = STRQRcode & createTLV("05", Vat) & seperator
''createTLVall = Trim(STRQRcode)
''
''End Function
''




